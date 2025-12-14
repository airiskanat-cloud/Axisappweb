import math
import os
import sys
import shutil
from io import BytesIO
import zipfile
import logging
import json
import ast
import operator as op

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ #1: Импорт openpyxl для создания Excel-файла ---
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage 
# -------------------------------------------------------------------------


# ----------------------------------------

# =========================
# КОНСТАНТЫ / НАСТРОЙКИ 
# =========================

DEBUG = False
logger = logging.getLogger(__name__)
# Конфигурируем базовый логгер для Streamlit
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

# --- УДАЛЕНИЕ ЛОКАЛЬНОЙ ЛОГИКИ ФАЙЛОВ ---
def resource_path(relative_path: str) -> str:
    """Возвращает корректный путь к ресурсу для PyInstaller или локального запуска."""
    try:
        if hasattr(sys, "_MEIPASS"):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(os.path.dirname(__file__))
    except Exception:
        base_path = os.getcwd()
    return os.path.join(base_path, relative_path)

# ID ВАШЕЙ ТАБЛИЦЫ (Убедитесь, что этот ID правильный)
GSPREAD_SHEET_ID = "1RJCkHf9qbjO0z3E2rdHQWAQyrGEHNL-W" 

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# Заголовки (Исключены неразрывные пробелы \xa0)
FORM_HEADER = [
    "Номер заказа", "№ позиции",
    "Тип изделия", "Вид изделия", "Створки",
    "Профильная система",
    "Тип стеклопакета",
    "Режим заполнения",
    "Ширина, мм", "Высота, мм",
    "LEFT, мм", "CENTER, мм", "RIGHT, мм", "TOP, мм",
    "Ширина створки, мм", "Высота створки, мм",
    "Кол-во Nwin",
    "Тонировка", "Сборка", "Монтаж",
    "Тип ручек", "Доводчик"
]

# Брендинг КП
COMPANY_NAME = "ООО «AXIS»"
COMPANY_CITY = "Город Астана"
COMPANY_PHONE = "+7 707 504 4040"
COMPANY_EMAIL = "Axisokna.kz@mail.ru"
COMPANY_SITE = "www.axis.kz"
LOGO_FILENAME = "logo_axis.png"

# =========================
# УТИЛИТЫ 
# =========================

def normalize_key(k):
    """Очищает ключ заголовка: удаляет неразрывные пробелы и нормализует."""
    if k is None:
        return None
    s = str(k)
    s = s.replace("\xa0", " ")
    s = " ".join(s.split())
    return s.strip().lower()

def _clean_cell_val(v):
    """Очищает строковое значение ячейки: удаляет неразрывные пробелы."""
    if v is None:
        return ""
    s = str(v)
    s = s.replace("\xa0", " ").strip()
    return s

def safe_float(value, default=0.0):
    """Безопасное преобразование в float, обрабатывая пробелы и запятые."""
    try:
        if value is None:
            return default
        s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
        if s == "":
            return default
        return float(s)
    except Exception:
        return default

def safe_int(value, default=0):
    """Безопасное преобразование в int."""
    try:
        if value is None:
            return default
        s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
        if s == "":
            return default
        return int(float(s))
    except Exception:
        return default

def get_field(row: dict, needle: str, default=None):
    """Получает значение поля, используя частичное совпадение в нижнем регистре."""
    needle = (needle or "").lower().strip()
    for k, v in row.items():
        if k and needle in str(k).lower():
            return v
    return default

# =========================
# БЕЗОПАСНЫЙ EVAL (ФОРМУЛЫ)
# =========================

_allowed_ops = {
    ast.Add: op.add, ast.Sub: op.sub, ast.Mult: op.mul, ast.Div: op.truediv,
    ast.Pow: op.pow, ast.USub: op.neg, ast.UAdd: op.pos, ast.Mod: op.mod,
    ast.FloorDiv: op.floordiv, ast.Lt: op.lt, ast.Gt: op.gt, ast.LtE: op.le,
    ast.GtE: op.ge, ast.Eq: op.eq, ast.NotEq: op.ne,
    ast.And: lambda a,b: a and b, ast.Or:  lambda a,b: a or b,
}

def _eval_ast(node, names):
    """Рекурсивная безопасная оценка AST-узла."""
    if isinstance(node, ast.Expression):
        return _eval_ast(node.body, names)

    if isinstance(node, (ast.Constant, ast.Num)):
        return node.value if isinstance(node, ast.Constant) else node.n

    if isinstance(node, ast.UnaryOp):
        val = _eval_ast(node.operand, names)
        fn = _allowed_ops.get(type(node.op))
        if fn: return fn(val)

    if isinstance(node, ast.BinOp):
        left = _eval_ast(node.left, names)
        right = _eval_ast(node.right, names)
        fn = _allowed_ops.get(type(node.op))
        if fn: return fn(left, right)

    if isinstance(node, ast.Name):
        if node.id in names:
            return names[node.id]
        raise ValueError(f"Недопустимое имя '{node.id}'")

    if isinstance(node, ast.Call):
        func = node.func
        # Разрешаем math.func()
        if isinstance(func, ast.Attribute) and isinstance(func.value, ast.Name) and func.value.id == "math":
            fname = func.attr
            if hasattr(math, fname):
                args = [_eval_ast(a, names) for a in node.args]
                return getattr(math, fname)(*args)

        # Разрешаем min() и max()
        if isinstance(func, ast.Name) and func.id in ("min", "max"):
            args = [_eval_ast(a, names) for a in node.args]
            return globals()[func.id](*args)

    if isinstance(node, ast.Compare):
        if len(node.ops) != 1:
            raise ValueError("Сложные сравнения запрещены")
        left = _eval_ast(node.left, names)
        right = _eval_ast(node.comparators[0], names)
        fn = _allowed_ops.get(type(node.ops[0]))
        if fn: return fn(left, right)

    raise ValueError(f"Недопустимый элемент формулы: {type(node).__name__}")

def safe_eval_formula(formula: str, context: dict) -> float:
    """Безопасно вычисляет формулу из строки с заданным контекстом переменных."""
    formula = (formula or "").strip()
    if not formula:
        return 0.0

    # Заменяем все неразрывные пробелы в формуле
    formula = formula.replace('\xa0', ' ')

    names = {
        **context,
        "math": math,
        "min": min,
        "max": max,
    }

    try:
        # Убеждаемся, что все переменные в контексте являются float или int
        # Это предотвратит ошибки, если в таблице будет строка (например, '1000.0')
        safe_context = {k: safe_float(v, 0.0) if not isinstance(v, (int, float)) else v for k, v in names.items()}
        
        node = ast.parse(formula, mode="eval")
        result = _eval_ast(node, safe_context)
        return float(result)
    except Exception:
        logger.exception("Ошибка вычисления формулы: %s", formula)
        return 0.0

# =========================
# GOOGLE SHEETS CLIENT 
# =========================

class GoogleSheetsClient:

    def __init__(self, sheet_id: str):
        self.sheet_id = sheet_id
        self._worksheets_cache = {}
        self.load()

    @st.cache_resource
    def _auth(_self):
        """Авторизация gspread через Streamlit/Render Secrets."""
        key_file_path = "/etc/secrets/gcp-key.json"

        if not os.path.exists(key_file_path):
            st.error("❌ Файл gcp-key.json не найден в Secrets. Проверьте конфигурацию хостинга.")
            st.stop()

        try:
            creds = Credentials.from_service_account_file(
                key_file_path,
                scopes=[
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"
                ]
            )
            return gspread.authorize(creds)

        except Exception as e:
            st.error(f"❌ Ошибка аутентификации Google Sheets: {e}")
            st.stop()

    def load(self):
        """Загружает рабочую книгу по ID."""
        try:
            client = self._auth()
            self.wb = client.open_by_key(self.sheet_id)
            logger.info("Успешно подключен к Google Sheets.")
        except Exception as e:
            st.error(f"Критическая ошибка при подключении к Google Sheets. {e}")
            st.stop()
            
    def ws(self, name: str):
        """Получает рабочий лист по имени, используя кеш."""
        if name in self._worksheets_cache:
            return self._worksheets_cache[name]
        try:
            ws = self.wb.worksheet(name)
            self._worksheets_cache[name] = ws
            return ws
        except gspread.WorksheetNotFound:
            if name == SHEET_FORM:
                # Создаем лист ЗАПРОСЫ, если он не найден
                ws = self.wb.add_worksheet(name, rows="100", cols="30")
                self._worksheets_cache[name] = ws
                ws.append_row(FORM_HEADER)
                return ws
            
            st.error(f"Лист '{name}' не найден в Google Sheets. Проверьте название листа в таблице.")
            st.stop()

    @st.cache_data(ttl=3600) # Кешируем данные справочников на 1 час
    def read_records(_self, sheet_name: str):
        """Читает записи из листа, используя заголовок первой строки как ключи."""
        ws = _self.ws(sheet_name)
        rows = ws.get_all_values()
        
        if not rows:
            return []
            
        header_raw = rows[0]
        header = []
        used = {}

        for h in header_raw:
            key = normalize_key(h)
            if not key: # Пропускаем пустые столбцы
                header.append(None)
                continue
            if key in used:
                used[key] += 1
                key = f"{key}_{used[key]}"
            else:
                used[key] = 1
            header.append(key)

        records = []
        for r in rows[1:]:
            # Пропускаем полностью пустые строки
            if all(v is None or v == "" for v in r):
                continue
            row = {}
            for i, k in enumerate(header):
                if k is None:
                    continue
                if i < len(r):
                    row[k] = r[i]
                else:
                    row[k] = None
            records.append(row)
        return records

    def clear_and_write(self, sheet_name: str, header: list, rows: list):
        """Отключаем запись промежуточных расчетов в Google Sheets, как было в оригинале."""
        pass

    def append_form_row(self, row: list):
        """Добавляет строку в лист ЗАПРОСЫ."""
        try:
            ws = self.ws(SHEET_FORM)
            ws.append_row(row, value_input_option='USER_ENTERED')
            logger.info("Строка успешно добавлена в лист ЗАПРОСЫ.")
        except Exception as e:
            logger.error("Ошибка при записи в лист ЗАПРОСЫ: %s", e)
            st.error(f"Ошибка при записи в Google Sheets: {e}")

# =========================
# ПОЛЬЗОВАТЕЛИ (ЛОГИН)
# =========================

def load_users(excel: GoogleSheetsClient):
    """Загружает пользователей для аутентификации."""
    # УДАЛЕНО: excel.load() - т.к. загрузка происходит в __init__
    rows = excel.read_records(SHEET_USERS)
    users = {}

    for r in rows:
        login = _clean_cell_val(get_field(r, "логин", "")).lower()
        pwd = _clean_cell_val(get_field(r, "парол", "")).strip()
        role = _clean_cell_val(get_field(r, "роль", ""))

        if login:
            users[login] = {"password": pwd, "role": role, "_raw_login": login}

    return users

def login_form(excel: GoogleSheetsClient):
    """Форма входа в Streamlit."""
    if "current_user" in st.session_state:
        return st.session_state["current_user"]

    st.sidebar.title("🔐 Вход в систему")
    with st.sidebar.form("login_form"):
        login = st.text_input("Логин")
        password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти")

    users = load_users(excel)

    if submitted:
        entered_login = (login or "").strip().lower()
        entered_pass = (password or "").strip()

        user = users.get(entered_login)

        if user:
            real_pass = (user["password"] or "").strip()
            if entered_pass == real_pass:
                st.session_state["current_user"] = {
                    "login": user["_raw_login"],
                    "role": user["role"],
                }
                st.sidebar.success(f"Привет, {user['_raw_login']}!")
                # Используем st.rerun() вместо deprecated st.experimental_rerun()
                st.rerun() 
                return st.session_state["current_user"]

        st.sidebar.error("Неверный логин или пароль")

    return None

# =========================
# CALCULATORS 
# =========================

class GabaritCalculator:
    HEADER = ["Тип элемента", "Фактическое значение"]

    def __init__(self, excel_client: GoogleSheetsClient):
        self.excel = excel_client

    def _calc_imposts_context(self, width, height, left, center, right, top):
        """Рассчитывает количество импостов и прямоугольников."""
        n_sections_vert = 0
        if left > 0: n_sections_vert += 1
        if center > 0: n_sections_vert += 1
        if right > 0: n_sections_vert += 1
        
        n_imp_vert = max(0, n_sections_vert - 1)
        n_imp_hor = 0
        if top > 0: n_imp_hor += 1

        n_impost = n_imp_vert + n_imp_hor
        n_frame_rect = 1 + n_imp_vert + n_imp_hor 
        n_rect = n_frame_rect
        n_corners = 4 * n_frame_rect

        return {
            "n_imp_vert": n_imp_vert,
            "n_imp_hor": n_imp_hor,
            "n_impost": n_impost,
            "n_frame_rect": n_frame_rect,
            "n_rect": n_rect,
            "n_corners": n_corners,
        }

    def calculate(self, order: dict, sections: list):
        """Вычисляет габариты (длины, штуки) по формулам из СПРАВОЧНИК-3."""
        ref_rows = self.excel.read_records(SHEET_REF3)

        total_area = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
        total_perimeter = sum(s.get("perimeter_m", 0.0) * s.get("Nwin", 1) for s in sections)

        if not ref_rows:
            return [], total_area, total_perimeter

        gabarit_values = []

        for row in ref_rows:
            type_elem = get_field(row, "тип элемент", "")
            formula = get_field(row, "формула_python", "")
            if not type_elem or not formula:
                continue

            total_value = 0.0

            for s in sections:
                
                is_door_section = s.get("kind") == "door"
                is_non_tamur_section = s.get("kind") in ["window", "door"] and order.get("product_type") != "Тамбур"

                # Определяем габариты рамы/изделия
                if is_door_section:
                    # Для дверного блока Тамбура (door) используем frame_width/height
                    width = s.get("frame_width_mm", 0.0)
                    height = s.get("frame_height_mm", 0.0)
                else:
                    # Для окна/панели используем width/height
                    width = s.get("width_mm", 0.0)
                    height = s.get("height_mm", 0.0)
                
                sash_w = 0.0
                sash_h = 0.0

                if s.get("leaves"):
                    # Берем размеры первой створки (для формул)
                    first_leaf = s.get("leaves", [{}])[0]
                    sash_w = first_leaf.get("width_mm", 0.0)
                    sash_h = first_leaf.get("height_mm", 0.0)
                    
                left = s.get("left_mm", 0.0)
                center = s.get("center_mm", 0.0)
                right = s.get("right_mm", 0.0)
                top = s.get("top_mm", 0.0)
                
                # --- ИНЖЕНЕРНАЯ КОРРЕКТИРОВКА ГАБАРИТОВ СТВОРКИ ---
                # Используется, если створки есть (n_leaves > 0), но их размеры не заданы (0.0)
                if is_non_tamur_section and (sash_w <= 0.0 or sash_h <= 0.0) and s.get("n_leaves", 0) > 0:
                    C_DED = 60.0 # Смещение/вычет для фурнитуры/створочного профиля
                    
                    if sash_w <= 0.0:
                        # Упрощенная логика: если есть Left/Center/Right, деление не учитывается в этом блоке
                        # Принимаем всю ширину, если деление не задано явно для створки
                        if left > 0 and center == 0 and right == 0 and s.get("n_leaves", 0) == 1:
                            sash_w = max(0.0, width - left - C_DED) # Если створка в одной секции
                        else:
                            sash_w = width 
                    
                    if sash_h <= 0.0:
                        if top > 0:
                            sash_h = max(0.0, height - top - C_DED)
                        else:
                            sash_h = height
                # ----------------------------------------------------

                area = s.get("area_m2", 0.0)
                perimeter = s.get("perimeter_m", 0.0)
                qty = s.get("Nwin", 1)

                nsash = s.get("n_leaves", len(s.get("leaves", [])) or 0)

                ctx = {
                    "width": width, "height": height, "left": left, "center": center, "right": right, "top": top,
                    "area": area, "perimeter": perimeter, "qty": qty,
                    "sash_width": sash_w, "sash_height": sash_h, "sash_w": sash_w, "sash_h": sash_h,
                    "n_sash": nsash,
                    "n_sash_active": 1 if nsash >= 1 else 0, # Условно, что 1-я створка активная
                    "n_sash_passive": max(nsash - 1, 0),
                    "hinges_per_sash": 3,
                    "is_door": 1 if is_door_section else 0,
                }

                try:
                    geom = self._calc_imposts_context(width, height, left, center, right, top)
                    if isinstance(geom, dict):
                        ctx.update(geom)
                except Exception:
                    pass

                try:
                    # Умножаем результат формулы на количество идентичных изделий
                    calculated_value = safe_eval_formula(str(formula), ctx) * qty 
                    total_value += calculated_value
                except Exception:
                    logger.exception("Error evaluating formula for element %s", type_elem)

            gabarit_values.append([type_elem, total_value])

        self.excel.clear_and_write(SHEET_GABARITS, self.HEADER, gabarit_values)
        return gabarit_values, total_area, total_perimeter

class MaterialCalculator:
    HEADER = [
        "Тип изделия", "Система профиля", "Тип элемента", "Артикул", "Товар",
        "Ед.", "Цена за ед.", "Ед. фактического расхода",
        "Кол-во факт. расхода", "Норма к упаковке", "Ед. к отгрузке",
        "Кол-во к отгрузке", "Сумма"
    ]

    def __init__(self, excel_client: GoogleSheetsClient):
        self.excel = excel_client

    def _calc_imposts_context(self, width, height, left, center, right, top):
        """Рассчитывает количество импостов и прямоугольников."""
        n_sections_vert = 0
        if left > 0: n_sections_vert += 1
        if center > 0: n_sections_vert += 1
        if right > 0: n_sections_vert += 1
        
        n_imp_vert = max(0, n_sections_vert - 1)
        n_imp_hor = 0
        if top > 0: n_imp_hor += 1

        n_impost = n_imp_vert + n_imp_hor
        n_frame_rect = 1 + n_imp_vert + n_imp_hor
        n_rect = n_frame_rect
        n_corners = 4 * n_frame_rect

        return {
            "n_imp_vert": n_imp_vert, "n_imp_hor": n_imp_hor, "n_impost": n_impost,
            "n_frame_rect": n_frame_rect, "n_rect": n_rect, "n_corners": n_corners,
        }

    def calculate(self, order: dict, sections: list, selected_duplicates: dict):
        """Вычисляет расход материалов по формулам из СПРАВОЧНИК-1."""
        ref_rows = self.excel.read_records(SHEET_REF1)
        total_area = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
        if not ref_rows:
            return [], 0.0, total_area

        result_rows = []
        total_sum = 0.0

        for row in ref_rows:
            row_type = get_field(row, "тип издел", "")
            row_profile = get_field(row, "система проф", "")
            type_elem = get_field(row, "тип элемент", "")
            product_name = str(get_field(row, "товар", "") or "")
            
            # Фильтрация по типу изделия и профилю
            if row_type and str(row_type).strip().lower() != order.get("product_type", "").strip().lower():
                continue

            if row_profile and str(row_profile).strip().lower() != order.get("profile_system", "").strip().lower():
                continue

            # Фильтрация по выбору дубликатов
            if type_elem in selected_duplicates and selected_duplicates[type_elem]:
                chosen_names = selected_duplicates[type_elem]
                if product_name not in chosen_names:
                    continue
                
            formula = get_field(row, "формула_python", "")
            if not formula:
                formula = get_field(row, "формула фактического расхода", "") # Дублирующее поле
            if not formula:
                continue

            qty_fact_total = 0.0
            
            # Определение категории элемента для фильтрации Тамбура
            is_panel_frame = "рамный контур" in type_elem.lower() or "импост" in type_elem.lower() or "сухарь усилительный" in type_elem.lower()
            is_door_item = ("рама двери" in type_elem.lower() or "порог дверной" in type_elem.lower() or "створочный профиль" in type_elem.lower() or "петля" in type_elem.lower() or "замок" in type_elem.lower() or "цилиндр" in type_elem.lower() or "ручка" in type_elem.lower() or "фиксатор" in type_elem.lower() or "доводчик" in type_elem.lower())

            for s in sections:
                is_door_section = s.get("kind") == "door"
                is_panel_section = s.get("kind") == "panel" or s.get("kind") == "window"
                is_non_tamur_section = s.get("kind") in ["window", "door"] and order.get("product_type") != "Тамбур"
                
                # --- ЛОГИКА ФИЛЬТРАЦИИ ДЛЯ ТАМБУРА ---
                if order.get("product_type") == "Тамбур":
                    # Если это дверной элемент (петли/ручки/профиль) и текущая секция - глухая панель
                    if is_door_item and s.get("kind") == "panel":
                        continue
                    
                    # Если это рамный профиль/импост и текущая секция - дверной блок.
                    # Предполагаем, что профили тамбура используются для рамы двери и для рам глухих секций.
                    # Эта логика была сложной в оригинале; оставляем только базовую фильтрацию:
                    if s.get("kind") == "panel" and is_door_item and "сухарь усилительный" not in type_elem.lower():
                        continue
                    if is_door_section and is_panel_frame and "рама двери" not in type_elem.lower() and "сухарь усилительный" not in type_elem.lower():
                        pass # Если это дверь Тамбура, профиль/импост используется

                # Определяем габариты рамы/изделия
                if is_door_section:
                    width = s.get("frame_width_mm", 0.0)
                    height = s.get("frame_height_mm", 0.0)
                else:
                    width = s.get("width_mm", 0.0)
                    height = s.get("height_mm", 0.0)

                left = s.get("left_mm", 0.0)
                center = s.get("center_mm", 0.0)
                right = s.get("right_mm", 0.0)
                top = s.get("top_mm", 0.0)
                
                nsash = s.get("n_leaves", len(s.get("leaves", [])) or 0)
                sash_w = 0.0
                sash_h = 0.0

                if nsash > 0 and s.get("leaves"):
                    first_leaf = s.get("leaves", [{}])[0]
                    sash_w = first_leaf.get("width_mm", 0.0)
                    sash_h = first_leaf.get("height_mm", 0.0)
                    
                # --- ИНЖЕНЕРНАЯ КОРРЕКТИРОВКА ГАБАРИТОВ СТВОРКИ (дублируем, т.к. расчет независим) ---
                if is_non_tamur_section and nsash > 0 and (sash_w <= 0.0 or sash_h <= 0.0):
                    C_DED = 60.0 
                    
                    if sash_w <= 0.0:
                        if left > 0 and center == 0 and right == 0 and nsash == 1:
                            sash_w = max(0.0, width - left - C_DED)
                        else:
                            sash_w = width
                    
                    if sash_h <= 0.0:
                        if top > 0:
                            sash_h = max(0.0, height - top - C_DED)
                        else:
                            sash_h = height
                # ----------------------------------------------------
                    
                area = s.get("area_m2", 0.0)
                perimeter = s.get("perimeter_m", 0.0)
                qty = s.get("Nwin", 1)

                geom = self._calc_imposts_context(width, height, left, center, right, top)
                
                ctx = {
                    "width": width, "height": height, "left": left, "center": center, "right": right, "top": top,
                    "sash_width": sash_w, "sash_height": sash_h, "sash_w": sash_w, "sash_h": sash_h,
                    "area": area, "perimeter": perimeter, "qty": qty,
                    "nsash": nsash,
                    "n_sash": nsash, 
                    "n_sash_active": 1 if nsash >= 1 else 0,
                    "n_sash_passive": max(nsash - 1, 0),
                    "hinges_per_sash": 3,
                    "is_door": 1 if is_door_section else 0,
                }
                ctx.update(geom)

                try:
                    # Умножаем результат формулы на количество идентичных изделий
                    calculated_value = safe_eval_formula(str(formula), ctx) * qty
                    qty_fact_total += calculated_value
                except Exception:
                    logger.exception("Error evaluating material formula for %s (Formula: %s)", type_elem, formula)

            unit_price = safe_float(get_field(row, "цена за", 0.0))
            norm_per_pack = safe_float(get_field(row, "кол-во норм", 0.0))
            unit_pack = str(get_field(row, "ед .норма к упаковке", "") or "").strip()
            unit = str(get_field(row, "ед.", "") or "").strip()
            unit_fact = str(get_field(row, "ед. фактического расхода", "") or "").strip()

            if norm_per_pack > 0:
                # Округление до упаковки
                qty_to_ship = math.ceil(qty_fact_total / norm_per_pack)
                effective_qty = qty_to_ship * norm_per_pack
            else:
                qty_to_ship = qty_fact_total
                effective_qty = qty_fact_total

            sum_row = effective_qty * unit_price
            total_sum += sum_row

            result_rows.append([
                row_type if row_type is not None else "",
                row_profile if row_profile is not None else "",
                type_elem,
                get_field(row, "артикул", ""),
                product_name,
                unit,
                unit_price,
                unit_fact,
                qty_fact_total,
                norm_per_pack,
                unit_pack,
                qty_to_ship,
                sum_row
            ])

        self.excel.clear_and_write(SHEET_MATERIAL, self.HEADER, result_rows)
        return result_rows, total_sum, total_area

class FinalCalculator:
    HEADER = ["Наименование услуг", "Стоимость за м²/шт", "Ед", "Итого"]

    def __init__(self, excel_client: GoogleSheetsClient):
        self.excel = excel_client

    def _lookup_ref2_rows(self):
        """Читает справочник-2 (кэшированный)."""
        return self.excel.read_records(SHEET_REF2)

    def _find_price_by_header_match(self, needle_list: list, default=0.0):
        """Ищет стоимость по совпадению в заголовке."""
        ref2 = self._lookup_ref2_rows()
        if not ref2: return default
        
        for r in ref2:
            for k in r.keys():
                if k is None: continue
                hk = str(k).lower()
                for needle in needle_list:
                    if needle in hk and "стоимость" in hk:
                        return safe_float(r[k], default)
        return default

    def _find_price_for_filling(self, filling_value):
        """Ищет стоимость заполнения (для панелей)."""
        ref2 = self._lookup_ref2_rows()
        if not ref2: return 0.0
        fv = str(filling_value or "").strip().lower()
        
        for r in ref2:
            fill_key_found = None
            price_key_found = None
            
            # 1. Находим столбец заполнения/панели
            for k in r.keys():
                if k is None: continue
                if "панел" in str(k).lower() or "заполн" in str(k).lower():
                    if str(r[k] or "").strip().lower() == fv:
                        fill_key_found = r[k]
                        break
            
            if fill_key_found:
                # 2. Находим столбец стоимости в той же строке
                for kk in r.keys():
                    if kk is None: continue
                    if "стоимость" in str(kk).lower():
                        return safe_float(r[kk], 0.0)
                        
        return 0.0

    def _find_price_for_montage(self, montage_type):
        """Ищет стоимость монтажа (цена за м²)."""
        return self._find_price_by_header_match(["монтаж", "стоимость", "за м"], 0.0)

    def _find_price_for_glass_by_type(self, glass_type):
        """Ищет стоимость стеклопакета по типу."""
        ref2 = self._lookup_ref2_rows()
        if not ref2: return 0.0
        gt = str(glass_type or "").strip().lower()
        
        chosen = None
        for r in ref2:
            for k in r.keys():
                if k is None: continue
                # Ищем строку, где тип стеклопакета совпадает
                if "тип стеклопак" in str(k).lower():
                    v = r[k]
                    if v and str(v).strip().lower() == gt:
                        chosen = r
                        break
            if chosen: break
            
        if not chosen:
             # Если не нашли по типу, ищем "дефолтную" цену за м²
            for r in ref2:
                for k in r.keys():
                    if k is None: continue
                    hk = str(k).lower()
                    if "стоимость" in hk and ("стеклопак" in hk or "за м" in hk):
                        return safe_float(r[k], 0.0)
            return 0.0
        
        # Если нашли строку, ищем цену в ней
        for k in chosen.keys():
            if k is None: continue
            hk = str(k).lower()
            if "стоимость" in hk and ("стеклопак" in hk or "за м" in hk or "за м²" in hk):
                return safe_float(chosen[k], 0.0)
        return 0.0

    def _find_price_for_toning(self):
        """Ищет стоимость тонировки (цена за м²)."""
        return self._find_price_by_header_match(["тониров", "стоимость", "за м"], 0.0)

    def _find_price_for_assembly(self):
        """Ищет стоимость сборки (цена за м²)."""
        return self._find_price_by_header_match(["сбор", "стоимость", "за м"], 0.0)

    def _find_price_for_handles(self):
        """Ищет стоимость ручек (цена за шт)."""
        return self._find_price_by_header_match(["ручк", "стоимость", "шт"], 0.0)

    def _find_price_for_closer(self):
        """Ищет стоимость доводчика (цена за шт)."""
        return self._find_price_by_header_match(["доводчик", "стоимость", "шт"], 0.0)


    def calculate(self, order: dict, total_area_all: float, material_total: float, lambr_cost: float = 0.0, handles_qty: int = 0, closer_qty: int = 0):
        """Финальный расчет стоимости услуг и итоговой суммы."""
        
        glass_type = order.get("glass_type", "")
        toning = order.get("toning", "Нет")
        assembly = order.get("assembly", "Нет")
        montage = order.get("montage", "Нет")
        door_closer = order.get("door_closer", "Нет")

        price_glass = self._find_price_for_glass_by_type(glass_type)
        price_toning = self._find_price_for_toning()
        price_assembly = self._find_price_for_assembly()
        price_montage = self._find_price_for_montage(montage)
        price_handles = self._find_price_for_handles()
        price_closer = self._find_price_for_closer()

        rows = []

        # 1. Стеклопакет
        glass_sum = total_area_all * price_glass if total_area_all > 0 else 0.0
        rows.append(["Стеклопакет", price_glass, "за м²", glass_sum])

        # 2. Тонировка
        toning_sum = total_area_all * price_toning if (toning.lower() != "нет" and total_area_all > 0) else 0.0
        rows.append(["Тонировка", price_toning, "за м²", toning_sum])

        # 3. Сборка
        assembly_sum = total_area_all * price_assembly if assembly.lower() != "нет" else 0.0
        rows.append(["Сборка", price_assembly, "за м²", assembly_sum])

        # 4. Монтаж
        montage_sum = total_area_all * price_montage if montage.lower() != "нет" and total_area_all > 0 else 0.0
        rows.append(["Монтаж (" + str(montage) + ")", price_montage, "за м²", montage_sum])

        # 5. Материал (Профиль, фурнитура)
        rows.append(["Материал", "-", "-", material_total])
        
        # 6. Панели (Ламбри/Сэндвич) - рассчитываются отдельно
        if lambr_cost > 0.0:
            rows.append(["Панели (Ламбри/Сэндвич)", "-", "-", lambr_cost])

        # 7. Ручки
        handles_sum = price_handles * handles_qty if handles_qty > 0 else 0.0
        rows.append(["Ручки", price_handles, "шт.", handles_sum])

        # 8. Доводчик
        closer_sum = price_closer * closer_qty if closer_qty > 0 and door_closer.lower() != "нет" else 0.0
        rows.append(["Доводчик", price_closer, "шт.", closer_sum])

        base_sum = (
            glass_sum + toning_sum + assembly_sum + montage_sum + material_total +
            lambr_cost + handles_sum + closer_sum
        )

        # 9. Обеспечение (60%)
        ensure_sum = base_sum * 0.6
        rows.append(["Обеспечение (60%)", "", "", ensure_sum])

        # ИТОГО
        total_sum = base_sum + ensure_sum
        extra_rows = [["ИТОГО", "", "", total_sum]]

        self.excel.clear_and_write(SHEET_FINAL, self.HEADER, rows + extra_rows)
        return rows, total_sum, ensure_sum


# =========================
# EXPORT: коммерческое предложение 
# =========================

def build_smeta_workbook(order: dict,
                         base_positions: list,
                         lambr_positions: list,
                         total_area: float,
                         total_perimeter: float,
                         total_sum: float) -> bytes:
    
    # КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ #1: Workbook и XLImage доступны
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"

    logo_path = resource_path(LOGO_FILENAME)
    current_row = 1

    if os.path.exists(logo_path):
        try:
            # Загрузка изображения для openpyxl
            img = XLImage(logo_path) 
            img.height = 80
            img.width = 80
            ws.add_image(img, "A1")
        except Exception:
            pass

    contact_col = 3
    # Удаление \xa0 в тексте
    ws.cell(row=current_row, column=contact_col, value=COMPANY_NAME); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=COMPANY_CITY); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"Тел.: {COMPANY_PHONE}"); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"E-mail: {COMPANY_EMAIL}"); current_row += 1
    if COMPANY_SITE:
        ws.cell(row=current_row, column=contact_col, value=f"Сайт: {COMPANY_SITE}"); current_row += 1

    current_row += 1
    ws.cell(row=current_row, column=1, value="Коммерческое предложение"); current_row += 2

    # --- Общая информация о заказе ---
    
    # КОРРЕКЦИЯ #2: Обработка пустого filling_mode и заполнение из первой позиции
    filling_mode_val = order.get('filling_mode', '')
    if not filling_mode_val and base_positions:
        filling_mode_val = base_positions[0].get('filling', '')

    
    ws.cell(row=current_row, column=1, value=f"Заказ № {order.get('order_number','')}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип изделия: {order.get('product_type','')}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Профильная система: {order.get('profile_system','')}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип заполнения (панели): {filling_mode_val or '—'}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип стеклопакета: {order.get('glass_type','')}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тонировка: {order.get('toning','')}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Сборка: {order.get('assembly','')}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Монтаж: {order.get('montage','')}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип ручек: {order.get('handle_type','') or '—'}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Доводчик: {order.get('door_closer','')}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 2

    ws.cell(row=current_row, column=1, value="Состав позиции:"); current_row += 1

    # --- Детализация позиций ---
    for idx, p in enumerate(base_positions, start=1):
        # КОРРЕКЦИЯ #3: Корректное определение размеров для Тамбура/Двери
        if p.get('kind') == 'door' and order.get('product_type') == 'Тамбур':
            w = p.get('frame_width_mm', 0)
            h = p.get('frame_height_mm', 0)
        else:
            w = p.get('width_mm', 0)
            h = p.get('height_mm', 0)
            
        fill = p.get('filling', '') or (p.get('leaves', [{}])[0].get('filling', '') if p.get('leaves') else '')
        
        ws.cell(row=current_row, column=1, value=f"Позиция {idx}: {p.get('kind','').capitalize()}, {w} × {h} мм, N = {p.get('Nwin',1)}, заполнение={fill}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' ')
        current_row += 1

    if lambr_positions:
        current_row += 1
        ws.cell(row=current_row, column=1, value="Панели Ламбри / Сэндвич:"); current_row += 1
        for idx, p in enumerate(lambr_positions, start=1):
            w = p.get('width_mm', 0)
            h = p.get('height_mm', 0)
            ws.cell(row=current_row, column=1, value=f"Панель {idx}: {w} × {h} мм, N = {p.get('Nwin',1)}, заполнение={p.get('filling','')}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' ')
            current_row += 1

    current_row += 2
    ws.cell(row=current_row, column=1, value=f"Общая площадь: {total_area:.3f} м²").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Суммарный периметр: {total_perimeter:.3f} м").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' '); current_row += 1
    ws.cell(row=current_row, column=1, value=f"ИТОГО к оплате: {total_sum:.2f}").value = ws.cell(row=current_row, column=1).value.replace('\xa0', ' ')

    try:
        for col in ['A','B','C','D','E','F']:
            ws.column_dimensions[col].width = 25
    except Exception:
        pass

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================
# STREAMLIT UI: main
# =========================

def ensure_session_state():
    """Инициализация session_state."""
    if "tam_door_count" not in st.session_state:
        st.session_state["tam_door_count"] = 0
    if "tam_panel_count" not in st.session_state:
        st.session_state["tam_panel_count"] = 0
    if "sections_inputs" not in st.session_state:
        st.session_state["sections_inputs"] = []
    if 'pos_count' not in st.session_state:
         st.session_state['pos_count'] = 1

def _calculate_lambr_cost(sections: list, fin_calc: FinalCalculator):
    """
    Рассчитывает стоимость Ламбри/Сэндвич панелей.
    
    ПРЕДПОЛОЖЕНИЕ: Стоимость заполнения в СПРАВОЧНИК-2 указана за 1 погонный метр (м/п)
    6-метрового хлыста, и расчет производится по периметру.
    Если цена в СПРАВОЧНИК-2 указана за м², необходимо изменить формулу на area * price.
    """
    lambr_cost = 0.0
    
    for s in sections:
        
        # Заполнение для секции (для Тамбура: глухая панель или заполнение створки двери)
        fills = []
        if s.get("kind") == "door" and s.get("leaves"):
            for leaf in s["leaves"]:
                fills.append((str(leaf.get("filling") or "").strip().lower(), leaf.get("width_mm", 0.0), leaf.get("height_mm", 0.0), s.get("Nwin", 1)))
        elif s.get("kind") in ["panel", "window"]:
             fills.append((str(s.get("filling") or "").strip().lower(), s.get("width_mm", 0.0), s.get("height_mm", 0.0), s.get("Nwin", 1)))
             
        for fill_name, w_mm, h_mm, nwin in fills:
            if fill_name in ["ламбри без термо", "ламбри с термо", "сэндвич"]:
                price_per_meter = fin_calc._find_price_for_filling(fill_name)
                
                if price_per_meter > 0:
                    perimeter_m = 2 * (w_mm + h_mm) / 1000.0
                    
                    # Расчет количества 6-метровых хлыстов
                    count_hlyst = math.ceil(perimeter_m / 6.0) if perimeter_m > 0 else 0
                    price_per_hlyst = price_per_meter * 6.0
                    
                    lambr_cost += count_hlyst * price_per_hlyst * nwin 
                    
    return lambr_cost

def main():
    st.set_page_config(page_title="Axis Pro GF • Калькулятор", layout="wide") 
    
    ensure_session_state()

    # --- Инициализация Google Sheets Client ---
    excel = GoogleSheetsClient(GSPREAD_SHEET_ID)
    # -----------------------------------------

    user = login_form(excel)
    if not user:
        st.stop()

    st.title("📘 Калькулятор алюминиевых изделий (Axis Pro GF)")
    st.info(f"Пользователь: **{user['login']}**")

    # Загружаем справочники
    ref2_records = excel.read_records(SHEET_REF2)
    filling_types_set = set()
    montage_types_set = set()
    handle_types_set = set()
    glass_types_set = set()

    def _clean_for_set(v):
        if v is None:
            return None
        s = str(v).replace("\xa0", " ").strip()
        return s if s else None

    for row in ref2_records:
        # Для заполнения панелей
        f = _clean_for_set(get_field(row, "панел") or get_field(row, "заполн") or get_field(row, "заполнение"))
        if f: filling_types_set.add(f)
        # Для монтажа
        m = _clean_for_set(get_field(row, "монтаж", None))
        if m: montage_types_set.add(m)
        # Для ручек
        h = _clean_for_set(get_field(row, "ручк", None))
        if h: handle_types_set.add(h)
        # Для стеклопакетов
        g = _clean_for_set(get_field(row, "тип стеклопак", None) or get_field(row, "тип стеклопакета", None))
        if g: glass_types_set.add(g)

    filling_options_for_panels = sorted(list(filling_types_set))
    if 'Стеклопакет' not in filling_options_for_panels:
        filling_options_for_panels.append('Стеклопакет')
    # Устанавливаем дефолтное значение для панелей (для Тамбура)
    default_panel_fill_index = filling_options_for_panels.index('Стеклопакет') if 'Стеклопакет' in filling_options_for_panels else 0
    if 'Ламбри без термо' in filling_options_for_panels:
        default_panel_fill_index = filling_options_for_panels.index('Ламбри без термо')

    # Опции монтажа
    montage_options = sorted(list(montage_types_set)) if montage_types_set else ["Есть", "Нет"]
    if "Нет" not in montage_options: montage_options.append("Нет")
    montage_options.insert(0, montage_options.pop(montage_options.index("Нет")))

    # Опции ручек и стеклопакетов
    handle_types = sorted(list(handle_types_set)) if handle_types_set else [""]
    glass_types = sorted(list(glass_types_set)) if glass_types_set else ["двойной"]
    default_glass_index = 0
    if "двойной" in glass_types:
        default_glass_index = glass_types.index("двойной")


    # ---------- Sidebar: общие данные ----------
    with st.sidebar:
        st.header("Общие данные заказа")
        order_number = st.text_input("Номер заказа", value="")
        product_type = st.selectbox("Тип изделия", ["Окно", "Дверь", "Тамбур"])
        profile_system = st.selectbox("Профильная система", ["ALG 2030-45C", "ALG RUIT 63i", "ALG RUIT 73"])
        glass_type = st.selectbox("Тип стеклопакета (цена из СПРАВОЧНИК-2)", glass_types, index=default_glass_index)
        st.markdown("### Прочее")
        toning = st.selectbox("Тонировка", ["Нет", "Есть"])
        assembly = st.selectbox("Сборка", ["Нет", "Есть"])
        montage = st.selectbox("Монтаж (из СПРАВОЧНИК-2)", montage_options, index=0)
        handle_type = st.selectbox("Тип ручек", handle_types, index=0)
        door_closer = st.selectbox("Доводчик", ["Нет", "Есть"])
            
    col_left, col_right = st.columns([2, 1])

    with col_left:
        st.header(f"Позиции ({product_type.lower()})")
        
        base_positions_inputs = [] # Для Окна/Двери (не Тамбур)
        
        if product_type != "Тамбур":
            positions_count = st.number_input("Количество позиций", min_value=1, max_value=10, value=st.session_state.get('pos_count', 1), step=1, key='pos_count')
            
            for i in range(int(positions_count)):
                st.subheader(f"Позиция {i+1}")
                st.markdown("**Габариты рамы/изделия**")
                c1, c2, c_nwin = st.columns(3)
                width_mm = c1.number_input(f"Ширина изделия, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"w_{i}")
                height_mm = c2.number_input(f"Высота изделия, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"h_{i}")
                nwin = c_nwin.number_input(f"Кол-во идентичных рам (N) (поз. {i+1})", min_value=1, value=1, step=1, key=f"nwin_{i}")
                
                st.markdown("**Размеры импостов (для деления)**")
                c_l, c_c, c_r, c_t = st.columns(4)
                left_mm = c_l.number_input(f"LEFT, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"l_{i}", value=0.0)
                center_mm = c_c.number_input(f"CENTER, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"c_{i}", value=0.0)
                right_mm = c_r.number_input(f"RIGHT, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"r_{i}", value=0.0)
                top_mm = c_t.number_input(f"TOP, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"t_{i}", value=0.0)
                
                st.markdown("**Створки и фурнитура**")
                n_leaves = st.number_input(f"Общее количество створок (N_sash) (поз. {i+1})", min_value=0, value=0, step=1, key=f"nleaves_{i}")

                leaves_data = []
                if n_leaves > 0:
                    for L in range(int(n_leaves)):
                        st.markdown(f"**Размеры створки {L+1}**")
                        c_sash_w, c_sash_h = st.columns(2)
                        # Запрашиваем размер створки (если не задан, то 0.0)
                        sash_width_mm = c_sash_w.number_input(f"Ширина створки {L+1}, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"sw_{i}_{L}")
                        sash_height_mm = c_sash_h.number_input(f"Высота створки {L+1}, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"sh_{i}_{L}")
                        
                        leaves_data.append({
                            "width_mm": sash_width_mm, 
                            "height_mm": sash_height_mm,
                            "filling": glass_type 
                        })

                first_sash_w = leaves_data[0]['width_mm'] if leaves_data else 0.0
                first_sash_h = leaves_data[0]['height_mm'] if leaves_data else 0.0
                
                base_positions_inputs.append({
                    "width_mm": width_mm, "height_mm": height_mm,
                    "left_mm": left_mm, "center_mm": center_mm, "right_mm": right_mm, "top_mm": top_mm,
                    "sash_width_mm": first_sash_w, "sash_height_mm": first_sash_h,
                    "Nwin": nwin, "filling": glass_type,
                    "kind": "window" if product_type == "Окно" else "door",
                    "n_leaves": n_leaves, "leaves": leaves_data 
                })
        else:
            # --- Динамический блок для Тамбура (обновление логики) ---
            st.header("Параметры тамбура (дверные блоки и глухие панели)")

            c_add = st.columns([1,1,6])
            if c_add[0].button("Добавить дверной блок"): st.session_state["tam_door_count"] += 1
            if c_add[1].button("Добавить глухую секцию"): st.session_state["tam_panel_count"] += 1
            
            # --- Управление секциями Тамбура (удаление секций) ---
            current_sections = st.session_state.get("sections_inputs", [])
            st.markdown("---")
            st.markdown("**Управление текущими секциями:**")
            sections_to_remove = []
            
            # Предварительная подготовка секций для ввода
            if len(current_sections) != st.session_state.get("tam_door_count", 0) + st.session_state.get("tam_panel_count", 0):
                # Если счетчики не совпадают с количеством секций, нужно их синхронизировать.
                # Это сложно в Streamlit, лучше полагаться на кнопки "Добавить/обновить".
                pass

            # Дверные блоки
            for i in range(st.session_state.get("tam_door_count", 0)):
                # Поиск существующего блока по ключу/индексу
                existing_section = next((s for s in current_sections if s.get("id") == f"door_{i}"), None)
                
                with st.expander(f"🚪 Дверной блок #{i+1}", expanded=False):
                    name = st.text_input(f"Название блока #{i+1}", value=existing_section.get("block_name", f"Дверной блок {i+1}") if existing_section else f"Дверной блок {i+1}", key=f"door_name_{i}")
                    count = st.number_input(f"Кол-во одинаковых блоков #{i+1}", min_value=1, value=existing_section.get("Nwin", 1) if existing_section else 1, key=f"door_count_{i}")
                    dtype = st.selectbox(f"Тип двери #{i+1}", ["Одностворчатая","Двухстворчатая"], index=0, key=f"door_type_{i}")
                    frame_w = st.number_input(f"Ширина рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, value=existing_section.get("frame_width_mm", 0.0) if existing_section else 0.0, key=f"frame_w_{i}")
                    frame_h = st.number_input(f"Высота рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, value=existing_section.get("frame_height_mm", 0.0) if existing_section else 0.0, key=f"frame_h_{i}")
                    
                    st.subheader("Внутренние импосты (для деления рамы)")
                    c_imp1, c_imp2 = st.columns(2)
                    left = c_imp1.number_input(f"LEFT, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, value=existing_section.get("left_mm", 0.0) if existing_section else 0.0, key=f"left_{i}")
                    center = c_imp2.number_input(f"CENTER, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, value=existing_section.get("center_mm", 0.0) if existing_section else 0.0, key=f"center_{i}")
                    c_imp3, c_imp4 = st.columns(2)
                    right = c_imp3.number_input(f"RIGHT, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, value=existing_section.get("right_mm", 0.0) if existing_section else 0.0, key=f"right_{i}")
                    top = c_imp4.number_input(f"TOP, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, value=existing_section.get("top_mm", 0.0) if existing_section else 0.0, key=f"top_{i}")

                    default_leaves = 1 if dtype == "Одностворчатая" else 2
                    n_leaves_val = existing_section.get("n_leaves", default_leaves) if existing_section else default_leaves
                    n_leaves = st.number_input(f"Кол-во створок #{i+1}", min_value=1, value=n_leaves_val, key=f"n_leaves_{i}")

                    leaves = []
                    for L in range(int(n_leaves)):
                        st.markdown(f"**Створка {L+1}**")
                        # Поиск существующей створки
                        existing_leaf = existing_section.get("leaves", [{}])[L] if existing_section and L < len(existing_section.get("leaves", [])) else {}
                        
                        lw = st.number_input(f"Ширина створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, value=existing_leaf.get("width_mm", 0.0), key=f"leaf_w_{i}_{L}")
                        lh = st.number_input(f"Высота створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, value=existing_leaf.get("height_mm", 0.0), key=f"leaf_h_{i}_{L}")
                        
                        default_fill_idx = filling_options_for_panels.index(existing_leaf.get("filling", glass_type)) if existing_leaf.get("filling") in filling_options_for_panels else (filling_options_for_panels.index(glass_type) if glass_type in filling_options_for_panels else 0)
                        fill = st.selectbox(f"Заполнение створки {L+1} — блок {i+1}", options=filling_options_for_panels, index=default_fill_idx, key=f"leaf_fill_{i}_{L}")
                        leaves.append({"width_mm": lw, "height_mm": lh, "filling": fill})
                    
                    c_save, c_del = st.columns(2)
                    if c_save.button(f"✅ Обновить ДБ #{i+1}", key=f"save_door_{i}"):
                        new_section = {
                            "id": f"door_{i}", # Уникальный ID для обновления
                            "kind": "door",
                            "block_name": name,
                            "frame_width_mm": frame_w, "frame_height_mm": frame_h,
                            "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top, 
                            "n_leaves": int(n_leaves), "leaves": leaves,
                            "Nwin": int(count), "filling": glass_type 
                        }
                        # Удаление старой версии и добавление новой
                        st.session_state["sections_inputs"] = [s for s in st.session_state["sections_inputs"] if not (s.get("id") == f"door_{i}")]
                        st.session_state["sections_inputs"].append(new_section)
                        st.success(f"Дверной блок '{name}' добавлен/обновлён.")
                        st.rerun() # Перезапуск для обновления состояния
                    
                    if c_del.button(f"❌ Удалить ДБ #{i+1}", key=f"del_door_{i}"):
                        sections_to_remove.append(f"door_{i}")

            # Глухие секции (панели)
            for i in range(st.session_state.get("tam_panel_count", 0)):
                existing_section = next((s for s in current_sections if s.get("id") == f"panel_{i}"), None)
                
                with st.expander(f"🔲 Глухая секция #{i+1}", expanded=False):
                    name = st.text_input(f"Название панели #{i+1}", value=existing_section.get("block_name", f"Панель {i+1}") if existing_section else f"Панель {i+1}", key=f"panel_name_{i}")
                    count = st.number_input(f"Кол-во одинаковых панелей #{i+1}", min_value=1, value=existing_section.get("Nwin", 1) if existing_section else 1, key=f"panel_count_{i}")
                    p1, p2 = st.columns(2)
                    w = p1.number_input(f"Ширина панели, мм #{i+1}", min_value=0.0, step=10.0, value=existing_section.get("width_mm", 0.0) if existing_section else 0.0, key=f"panel_w_{i}")
                    h = p2.number_input(f"Высота панели, мм #{i+1}", min_value=0.0, step=10.0, value=existing_section.get("height_mm", 0.0) if existing_section else 0.0, key=f"panel_h_{i}")
                    
                    default_fill_idx = filling_options_for_panels.index(existing_section.get("filling", filling_options_for_panels[default_panel_fill_index])) if existing_section and existing_section.get("filling") in filling_options_for_panels else default_panel_fill_index
                    fill = st.selectbox(f"Заполнение панели #{i+1}", options=filling_options_for_panels, index=default_fill_idx, key=f"panel_fill_{i}")
                    
                    st.subheader("Внутренние импосты (для деления рамы)")
                    c_imp5, c_imp6 = st.columns(2)
                    left = c_imp5.number_input(f"LEFT, мм #{i+1} (ГС)", min_value=0.0, step=10.0, value=existing_section.get("left_mm", 0.0) if existing_section else 0.0, key=f"panel_left_{i}")
                    center = c_imp6.number_input(f"CENTER, мм #{i+1} (ГС)", min_value=0.0, step=10.0, value=existing_section.get("center_mm", 0.0) if existing_section else 0.0, key=f"panel_center_{i}")
                    c_imp7, c_imp8 = st.columns(2)
                    right = c_imp7.number_input(f"RIGHT, мм #{i+1} (ГС)", min_value=0.0, step=10.0, value=existing_section.get("right_mm", 0.0) if existing_section else 0.0, key=f"panel_right_{i}")
                    top = c_imp8.number_input(f"TOP, мм #{i+1} (ГС)", min_value=0.0, step=10.0, value=existing_section.get("top_mm", 0.0) if existing_section else 0.0, key=f"panel_top_{i}")

                    c_save, c_del = st.columns(2)
                    if c_save.button(f"✅ Обновить Панель #{i+1}", key=f"save_panel_{i}"):
                        new_section = {
                            "id": f"panel_{i}", # Уникальный ID
                            "kind": "panel", "block_name": name,
                            "width_mm": w, "height_mm": h,
                            "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top, 
                            "filling": fill, "Nwin": int(count)
                        }
                        st.session_state["sections_inputs"] = [s for s in st.session_state["sections_inputs"] if not (s.get("id") == f"panel_{i}")]
                        st.session_state["sections_inputs"].append(new_section)
                        st.success(f"Панель '{name}' добавлена/обновлена.")
                        st.rerun() # Перезапуск для обновления состояния
                        
                    if c_del.button(f"❌ Удалить Панель #{i+1}", key=f"del_panel_{i}"):
                        sections_to_remove.append(f"panel_{i}")
                        
            # Удаление секций после цикла
            if sections_to_remove:
                st.session_state["sections_inputs"] = [s for s in st.session_state["sections_inputs"] if s.get("id") not in sections_to_remove]
                # Корректируем счетчики, чтобы не было "пропусков"
                st.session_state["tam_door_count"] = len([s for s in st.session_state["sections_inputs"] if s.get("kind") == "door"])
                st.session_state["tam_panel_count"] = len([s for s in st.session_state["sections_inputs"] if s.get("kind") == "panel"])
                st.info(f"Удалены {len(sections_to_remove)} секций. Перезагрузка...")
                st.rerun()
            
            st.markdown("**Текущие секции Тамбура:**")
            if st.session_state["sections_inputs"]:
                 for idx, s in enumerate(st.session_state["sections_inputs"], start=1):
                    main_dim = f"{s.get('width_mm', s.get('frame_width_mm'))}x{s.get('height_mm', s.get('frame_height_mm'))}"
                    imposts = f" L{s.get('left_mm',0)} C{s.get('center_mm',0)} R{s.get('right_mm',0)} T{s.get('top_mm',0)}"
                    st.write(f"**{idx}. {s.get('kind').capitalize()}** ({s.get('block_name')}) — {main_dim}, N={s.get('Nwin',1)} | Заполнение: {s.get('filling', glass_type)} | Импосты:{imposts}")
            else:
                 st.info("Нет добавленных секций.")
        
        st.markdown("---")

    with col_right:
        st.header("Информация")
        st.info("Данные справочников кешируются на 1 час для ускорения работы.")
        st.info("Обратите внимание на логику расчета стоимости панелей (Ламбри/Сэндвич): предполагается, что цена указана за м/п 6-метрового хлыста (как в оригинальном коде).")
        
        # ---------- Выбор материалов при дублях ----------
        st.header("🧾 Выбор материалов при дублях")
        selected_duplicates = {}

        ref1 = excel.read_records(SHEET_REF1)
        groups = {}
        for row in ref1:
            row_type = str(get_field(row, "тип издел", "") or "").strip()
            row_profile = str(get_field(row, "система проф", "") or "").strip()

            if row_type and row_type.lower() != product_type.lower():
                continue
            if row_profile and row_profile.lower() != profile_system.lower():
                continue

            type_elem = str(get_field(row, "тип элемент", "") or "").strip()
            product_name = str(get_field(row, "товар", "") or "").strip()
            if not type_elem or not product_name:
                continue

            groups.setdefault(type_elem, set()).add(product_name)

        if not groups:
            st.info("Для выбранного типа изделия и профиля дублей материалов не найдено.")
        else:
            for type_elem, products in sorted(groups.items(), key=lambda kv: kv[0]):
                if len(products) <= 1:
                    continue
                default = sorted(list(products))
                chosen = st.multiselect(
                    f"Тип элемента: {type_elem}",
                    options=sorted(list(products)),
                    default=default,
                    key=f"dup_{type_elem}"
                )
                selected_duplicates[type_elem] = set(chosen)

    # ---------- Кнопка расчёта ----------
    st.markdown("---")
    calc_button = st.button("💾 Сохранить в Excel и выполнить расчёт", type='primary')

    if calc_button:
        if not order_number.strip():
            st.error("Введите номер заказа.")
            st.stop()

        sections = []
        
        if product_type != "Тамбур":
            sections = base_positions_inputs
        else:
            sections = st.session_state["sections_inputs"]
            
        if not sections:
            st.error("Необходимо задать хотя бы одну позицию/секцию.")
            st.stop()
            
        # Валидация габаритов и расчет площади/периметра
        valid_sections = []
        for p in sections:
            # Получаем фактические размеры для расчета
            w_val = p.get("width_mm", 0.0) if p.get("kind") != "door" else p.get("frame_width_mm", 0.0)
            h_val = p.get("height_mm", 0.0) if p.get("kind") != "door" else p.get("frame_height_mm", 0.0)

            if w_val <= 0 or h_val <= 0:
                st.error(f"Секция/позиция '{p.get('block_name', p.get('kind'))}' имеет нулевую ширину или высоту. Исправьте.")
                st.stop()
            
            area_m2 = (w_val * h_val) / 1_000_000.0
            perimeter_m = 2 * (w_val + h_val) / 1000.0
            valid_sections.append({**p, "area_m2": area_m2, "perimeter_m": perimeter_m})
            
        sections = valid_sections
            
        # --- Gabarit Calculation ---
        gab_calc = GabaritCalculator(excel)
        gabarit_rows, total_area_gab, total_perimeter_gab = gab_calc.calculate({"product_type": product_type}, sections)

        # --- Material Calculation ---
        mat_calc = MaterialCalculator(excel)
        material_rows, material_total, total_area_mat = mat_calc.calculate({"product_type": product_type, "profile_system": profile_system}, sections, selected_duplicates)
        
        # --- Intermediate Sums for FinalCalc ---
        total_area_all = total_area_gab # Используем общую площадь из GabaritCalc

        fin_calc = FinalCalculator(excel)
        lambr_cost = _calculate_lambr_cost(sections, fin_calc)

        # --- Handles / Door Closer Counts (1 шт на дверной блок/изделие) ---
        handles_count = 0
        closer_count = 0
        if product_type == "Дверь" or product_type == "Тамбур":
            for s in sections:
                if s.get("kind") == "door":
                    handles_count += s.get("Nwin", 1)
                    
                    if door_closer.lower() == "есть":
                        closer_count += s.get("Nwin", 1) 
                        
        # Для Окна: Ручки считаются по количеству створок, но в исходном коде такой логики не было.
        # Если нужно считать ручки для окон:
        # if product_type == "Окно":
        #    handles_count += sum(s.get("n_leaves", 0) * s.get("Nwin", 1) for s in sections)
        # Оставляем логику по умолчанию (только для дверей), как было в оригинале.
        
        # --- Final Calculation ---
        final_rows, total_sum, ensure_sum = fin_calc.calculate(
            {
                "product_type": product_type, "glass_type": glass_type, "toning": toning,
                "assembly": assembly, "montage": montage, "handle_type": handle_type, "door_closer": door_closer
            },
            total_area_all=total_area_all, material_total=material_total,
            lambr_cost=lambr_cost, handles_qty=handles_count, closer_qty=closer_count
        )
        
        st.success(f"Расчёт выполнен. Итоговая сумма: {total_sum:.2f}")

        # --- Сохраняем в ЗАПРОСЫ ---
        rows_for_form = []
        pos_index = 1
        
        for p in sections:
            sash_w = 0.0
            sash_h = 0.0
            if p.get("leaves"):
                first_leaf = p["leaves"][0]
                sash_w = first_leaf.get("width_mm", 0.0)
                sash_h = first_leaf.get("height_mm", 0.0)
            
            width_val = p.get("width_mm", 0.0) if p.get("kind") != "door" else p.get("frame_width_mm", 0.0)
            height_val = p.get("height_mm", 0.0) if p.get("kind") != "door" else p.get("frame_height_mm", 0.0)

            rows_for_form.append([
                order_number, pos_index, product_type,
                p.get("kind", ""), 
                p.get("n_leaves", 0),
                profile_system, glass_type, p.get("filling",""),
                width_val, height_val,
                p.get("left_mm", 0.0), p.get("center_mm", 0.0), p.get("right_mm", 0.0), p.get("top_mm", 0.0),
                sash_w, sash_h, 
                p.get("Nwin", 1),
                toning, assembly, montage, handle_type, door_closer,
            ])
            pos_index += 1

        for row in rows_for_form:
            excel.append_form_row(row)
        st.info("Данные сохранены в Google Sheets на листе 'ЗАПРОСЫ'.")

        # --- Вывод результатов и экспорт ---
        tab1, tab2, tab3 = st.tabs(["Габариты", "Материалы", "Итоговый расчет"])
        
        with tab1:
            st.subheader("Расчет по габаритам")
            if gabarit_rows:
                gab_disp = [{"Тип элемента": t, "Фактическое значение": v} for t, v in gabarit_rows]
                st.dataframe(gab_disp, use_container_width=True)
            st.write(f"Общая площадь: **{total_area_gab:.3f} м²**")
            st.write(f"Суммарный периметр: **{total_perimeter_gab:.3f} м**")
            
        with tab2:
            st.subheader("Расчёт материалов")
            st.warning("⚠️ Проверьте, что формулы в СПРАВОЧНИК-1 используют корректные Python переменные (width, sash_w, n_sash и т.д.)")
            
            if material_rows:
                mat_disp = []
                for r in material_rows:
                    mat_disp.append({
                        "Тип изделия": r[0], "Система профиля": r[1], "Тип элемента": r[2], "Артикул": r[3],
                        "Товар": r[4], "Ед.": r[5], "Цена за ед.": round(safe_float(r[6]), 2),
                        "Ед. факт. расхода": r[7], "Кол-во факт. расхода": round(safe_float(r[8]), 3),
                        "Норма к упаковке": r[9], "Ед. к отгрузке": r[10],
                        "Кол-во к отгрузке": round(safe_float(r[11]), 3), "Сумма": round(safe_float(r[12]), 2),
                    })
                st.dataframe(mat_disp, use_container_width=True)
            st.write(f"Итого по материалам (Профиль, Фурнитура): **{material_total:.2f}**")
            st.write(f"Панели (ламбри/сэндвич) — Итого: **{lambr_cost:.2f}**")

        with tab3:
            st.subheader("Итоговый расчет с монтажом")
            if final_rows:
                fin_disp = []
                for name, price, unit, total_val in final_rows:
                    fin_disp.append({
                        "Наименование услуг": name,
                        "Стоимость за м²/шт": price if isinstance(price, str) else round(price, 2),
                        "Ед": unit,
                        "Итого": total_val if isinstance(total_val, str) else round(total_val, 2),
                    })
                st.dataframe(fin_disp, use_container_width=True)
            st.write(f"Обеспечение (60%): **{ensure_sum:.2f}**")
            st.write(f"ИТОГО к оплате: **{total_sum:.2f}**")

        # --- Экспорт коммерческого предложения ---
        base_pos = [s for s in sections if s.get("kind") in ["window"] and product_type == "Окно"]
        tam_door_pos = [s for s in sections if s.get("kind") == "door"] # Двери, включая двери Тамбура
        lambr_pos = [s for s in sections if s.get("kind") == "panel"] # Глухие панели Тамбура
        
        smeta_bytes = build_smeta_workbook(
            order={
                "order_number": order_number, "product_type": product_type, "profile_system": profile_system,
                "filling_mode": "", "glass_type": glass_type, "toning": toning, "assembly": assembly, 
                "montage": montage, "handle_type": handle_type, "door_closer": door_closer,
            },
            base_positions=base_pos + tam_door_pos,
            lambr_positions=lambr_pos,
            total_area=total_area_all,
            total_perimeter=total_perimeter_gab,
            total_sum=total_sum,
        )

        default_name = f"Коммерческое_предложение_Заказ_{order_number}.xlsx"
        st.download_button(
            "⬇️ Скачать коммерческое предложение в Excel",
            data=smeta_bytes,
            file_name=default_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        
    # ---------- Кнопка выхода ----------
    if st.sidebar.button("Выйти"):
        st.session_state.pop("current_user", None)
        # КОРРЕКЦИЯ #4: st.rerun()
        st.rerun()


if __name__ == "__main__":
    main()
