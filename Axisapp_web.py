import math
import os
import sys
import zipfile
from io import BytesIO

import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.drawing.image import Image as XLImage

# =========================
# Константы и пути
# =========================

def resource_path(relative_path: str) -> str:
    """Возвращает корректный путь к файлу (поддержка PyInstaller)."""
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


TEMPLATE_EXCEL_NAME = "axis_pro_gf.xlsx"
EXCEL_FILE = resource_path(TEMPLATE_EXCEL_NAME)

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# Шапка для листа ЗАПРОСЫ
FORM_HEADER = [
    "Номер заказа", "№ позиции",
    "Тип изделия", "Вид изделия", "Створки",
    "Профильная система",
    "Тип стеклопакета",
    "Режим заполнения",  # Ламбри / Сэндвич
    "Ширина, мм", "Высота, мм",
    "LEFT, мм", "CENTER, мм", "RIGHT, мм", "TOP, мм",
    "Ширина створки, мм", "Высота створки, мм",
    "Кол-во Nwin",
    "Тонировка", "Сборка", "Монтаж",
    "Тип ручек", "Доводчик"
]

# Брендинг для Excel (логотип+контакты)
COMPANY_NAME = "ООО «AXIS»"
COMPANY_CITY = "Город Астана"
COMPANY_PHONE = "+7 707 504 4040"
COMPANY_EMAIL = "Axisokna.kz@mail.ru"
COMPANY_SITE = "www.axis.kz"  # опционально
LOGO_FILENAME = "logo_axis.png"  # файл логотипа рядом с .py

# =========================
# Утилиты
# =========================

def safe_float(value, default=0.0):
    try:
        if value is None:
            return default
        return float(str(value).replace(",", "."))
    except Exception:
        return default


def safe_int(value, default=0):
    try:
        if value is None:
            return default
        return int(float(str(value).replace(",", ".")))
    except Exception:
        return default


def normalize_key(k):
    if k is None:
        return None
    return str(k).strip()


def get_field(row: dict, needle: str, default=None):
    """Поиск значения в записи по подстроке имени поля (независимо от регистра и пробелов)."""
    if row is None:
        return default
    needle = (needle or "").lower().strip()
    for k in row.keys():
        if k is None:
            continue
        if needle in str(k).lower().strip():
            return row[k]
    return default


def eval_formula(formula: str, context: dict) -> float:
    """Выполняет python-формулу (строго ограниченный контекст) для ОДНОЙ секции."""
    formula = (formula or "").strip()
    if not formula:
        return 0.0

    allowed_names = {
        "width": context.get("width", 0.0),
        "height": context.get("height", 0.0),
        "left": context.get("left", 0.0),
        "center": context.get("center", 0.0),
        "right": context.get("right", 0.0),
        "top": context.get("top", 0.0),
        "sash_width": context.get("sash_width", 0.0),
        "sash_height": context.get("sash_height", 0.0),
        "area": context.get("area", 0.0),
        "perimeter": context.get("perimeter", 0.0),
        "qty": context.get("qty", 1.0),
        "nsash": context.get("nsash", 1),
        "n_sash_active": context.get("n_sash_active", 1),
        "n_sash_passive": context.get("n_sash_passive", 0),
        "hinges_per_sash": context.get("hinges_per_sash", 3),
        "n_rect": context.get("n_rect", 1),
        "n_frame_rect": context.get("n_frame_rect", 1),
        "n_impost": context.get("n_impost", 0),
        "n_imp_vert": context.get("n_imp_vert", 0),
        "n_imp_hor": context.get("n_imp_hor", 0),
        "n_corners": context.get("n_corners", 0),
        "math": math,
        "max": max,
        "min": min,
    }

    try:
        result = eval(formula, {"__builtins__": {}}, allowed_names)
        return float(result)
    except Exception as e:
        print(f"Ошибка в формуле '{formula}': {e}")
        return 0.0

# =========================
# Excel client с проверкой
# =========================

def is_probably_xlsx(path: str) -> bool:
    """
    Надёжная проверка .xlsx: файл должен существовать, быть zip и содержать ключевые файлы.
    """
    try:
        if not os.path.exists(path) or not os.path.isfile(path):
            return False
        with zipfile.ZipFile(path, "r") as z:
            names = z.namelist()
            if '[Content_Types].xml' in names and any(n.startswith('xl/') for n in names):
                return True
            return False
    except Exception:
        return False


class ExcelClient:
    def __init__(self, filename: str):
        self.filename = filename
        # Если файла нет — создаём корректный шаблон
        if not os.path.exists(self.filename):
            self._create_template()
        self.load()

    def _create_template(self):
        try:
            wb = Workbook()
            # удаляем стандартный лист
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]
            wb.create_sheet(SHEET_FORM)
            wb.create_sheet(SHEET_REF1)
            wb.create_sheet(SHEET_REF2)
            wb.create_sheet(SHEET_REF3)
            wb.create_sheet(SHEET_USERS)
            wb.save(self.filename)
            print(f"Создан новый шаблон Excel: {self.filename}")
        except Exception as e:
            print(f"Ошибка при создании шаблона Excel: {e}")

    def load(self):
        try:
            self.wb = load_workbook(self.filename, data_only=True)
        except zipfile.BadZipFile:
            print(f"BadZipFile: {self.filename} is not a valid xlsx.")
            # пытаемся сделать бэкап поврежденного файла и создать новый шаблон
            try:
                bak = self.filename + ".corrupt." + str(int(os.path.getmtime(self.filename)))
                os.rename(self.filename, bak)
                print(f"Старый файл переименован в {bak}. Создаём новый шаблон {self.filename}")
            except Exception as e:
                print(f"Не удалось переименовать повреждённый файл: {e}. Попробуем перезаписать.")
            self._create_template()
            self.wb = load_workbook(self.filename, data_only=True)
        except Exception as e:
            print(f"Ошибка при загрузке Excel: {e}")
            # создаём рабочую книгу в памяти, но не перезаписываем файл
            self.wb = Workbook()

    def save(self):
        try:
            self.wb.save(self.filename)
        except Exception as e:
            print(f"Ошибка при сохранении Excel: {e}")

    def ws(self, name: str):
        if name in self.wb.sheetnames:
            return self.wb[name]
        ws = self.wb.create_sheet(name)
        self.save()
        return ws

    def read_records(self, sheet_name: str):
        ws = self.ws(sheet_name)
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        raw_header = rows[0]
        # нормализуем заголовки: оставляем строки, stripped
        header = []
        for h in raw_header:
            header.append(normalize_key(h) if h is not None else None)
        data_rows = rows[1:]
        records = []
        for r in data_rows:
            if all(v is None for v in r):
                continue
            rec = {}
            for i, key in enumerate(header):
                if key is None:
                    continue
                rec[str(key)] = r[i] if i < len(r) else None
            records.append(rec)
        return records

    def clear_and_write(self, sheet_name: str, header: list, rows: list):
        ws = self.ws(sheet_name)
        # удаляем все строки
        try:
            ws.delete_rows(1, ws.max_row or 1)
        except Exception:
            if sheet_name in self.wb.sheetnames:
                del self.wb[sheet_name]
            ws = self.wb.create_sheet(sheet_name)
        if header:
            ws.append(header)
        for row in rows:
            ws.append(row)
        self.save()

    def append_form_row(self, row: list):
        ws = self.ws(SHEET_FORM)
        # Если лист пустой, добавляем шапку
        if ws.max_row == 1 and all(cell.value is None for cell in ws[1]):
            ws.append(FORM_HEADER)
        ws.append(row)
        self.save()

# =========================
# Пользователи
# =========================

def load_users(excel: ExcelClient):
    excel.load()
    try:
        rows = excel.read_records(SHEET_USERS)
    except Exception:
        return {}

    users = {}
    for row in rows:
        # Попробуем найти столбцы похожие на логин/пароль/роль
        raw_login = get_field(row, "логин", "") or get_field(row, "login", "")
        raw_password = get_field(row, "парол", "") or get_field(row, "password", "")
        raw_role = get_field(row, "роль", "") or get_field(row, "role", "")

        # Нормализация
        login = str(raw_login).strip()
        # логин приведём к нижнему регистру, чтобы ввод был нечувствителен к регистру
        login_norm = login.lower()

        password = str(raw_password or "").strip()
        # убираем "звёздочки" маскировки в конце или начале, если они есть
        if password.startswith("*") or password.endswith("*"):
            password = password.strip("*").strip()
        # можно также убрать случайные пробелы по краям
        password = password.strip()

        role = str(raw_role or "").strip()

        if login_norm:
            users[login_norm] = {"password": password, "role": role, "_raw_login": login}
    return users



def login_form(excel: ExcelClient):
    if "current_user" in st.session_state:
        return st.session_state["current_user"]

    st.sidebar.title("🔐 Вход в систему")
    with st.sidebar.form("login_form", clear_on_submit=False):
        login = st.text_input("Логин")
        password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти")

    users = load_users(excel)

    # Временно показываем debug (при необходимости можно отключить)
    # st.sidebar.write("DEBUG: users:", users)

    if submitted:
    entered_login = (login or "").strip().lower()
    entered_password = (password or "").strip()
    user = users.get(entered_login)
    if user and entered_password == user["password"]:
        st.session_state["current_user"] = {"login": user["_raw_login"], "role": user.get("role", "")}
        st.sidebar.success(f"Привет, {user['_raw_login']}!")
        return st.session_state["current_user"]
    else:
        st.sidebar.error("Неверный логин или пароль")


# =========================
# Gabarit / Material / Final calculators
# =========================

class GabaritCalculator:
    HEADER = ["Тип элемента", "Фактическое значение"]

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def _calc_imposts_context(self, width, height, left, center, right, top):
        n_imp_vert = 0
        if left > 0:
            n_imp_vert += 1
        if center > 0:
            n_imp_vert += 1
        if right > 0:
            n_imp_vert += 1

        n_imp_hor = 0
        if top > 0:
            n_imp_hor += 1

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
        ref_rows = self.excel.read_records(SHEET_REF3)
        # если нет справочника — не ломаемся, но возвращаем пустой список
        if not ref_rows:
            # всё равно подсчитаем суммарную площадь и периметр
            total_area = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
            total_perimeter = sum(s.get("perimeter_m", 0.0) * s.get("Nwin", 1) for s in sections)
            return [], total_area, total_perimeter

        total_area = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
        total_perimeter = sum(s.get("perimeter_m", 0.0) * s.get("Nwin", 1) for s in sections)
        gabarit_values = []

        for row in ref_rows:
            type_elem = get_field(row, "тип элемент", "")
            formula = get_field(row, "формула_python", "")
            if not type_elem or not formula:
                continue

            total_value = 0.0

            for s in sections:
                width = s.get("width_mm", 0.0)
                height = s.get("height_mm", 0.0)
                left = s.get("left_mm", 0.0)
                center = s.get("center_mm", 0.0)
                right = s.get("right_mm", 0.0)
                top = s.get("top_mm", 0.0)
                sash_w = s.get("sash_width_mm", width)
                sash_h = s.get("sash_height_mm", height)
                area = s.get("area_m2", 0.0)
                perimeter = s.get("perimeter_m", 0.0)
                qty = s.get("Nwin", 1)

                geom = self._calc_imposts_context(width, height, left, center, right, top)

                nsash = s.get("nsash", 1)
                if s.get("kind") == "door":
                    if s.get("door_type") == "double":
                        nsash = 2
                    else:
                        nsash = 1

                ctx = {
                    "width": width,
                    "height": height,
                    "left": left,
                    "center": center,
                    "right": right,
                    "top": top,
                    "sash_width": sash_w,
                    "sash_height": sash_h,
                    "area": area,
                    "perimeter": perimeter,
                    "qty": qty,
                    "nsash": nsash,
                    "n_sash_active": 1 if nsash >= 1 else 0,
                    "n_sash_passive": max(nsash - 1, 0),
                    "hinges_per_sash": 3,
                }
                ctx.update(geom)

                total_value += eval_formula(str(formula), ctx)

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

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def _calc_imposts_context(self, width, height, left, center, right, top):
        n_imp_vert = 0
        if left > 0:
            n_imp_vert += 1
        if center > 0:
            n_imp_vert += 1
        if right > 0:
            n_imp_vert += 1

        n_imp_hor = 0
        if top > 0:
            n_imp_hor += 1

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

    def calculate(self, order: dict, sections: list, selected_duplicates: dict):
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

            # фильтр по типу изделия
            if row_type:
                if str(row_type).strip().lower() != order.get("product_type", "").strip().lower():
                    continue

            # фильтр по системе профиля
            if row_profile:
                if str(row_profile).strip().lower() != order.get("profile_system", "").strip().lower():
                    continue

            # фильтр по выбору дублей
            if type_elem in selected_duplicates and selected_duplicates[type_elem]:
                chosen_names = selected_duplicates[type_elem]
                if product_name not in chosen_names:
                    continue

            formula = get_field(row, "формула_python", "")
            if not formula:
                formula = get_field(row, "формула фактического расхода", "")
            if not formula:
                continue

            qty_fact_total = 0.0

            for s in sections:
                width = s.get("width_mm", 0.0)
                height = s.get("height_mm", 0.0)
                left = s.get("left_mm", 0.0)
                center = s.get("center_mm", 0.0)
                right = s.get("right_mm", 0.0)
                top = s.get("top_mm", 0.0)
                sash_w = s.get("sash_width_mm", width)
                sash_h = s.get("sash_height_mm", height)
                area = s.get("area_m2", 0.0)
                perimeter = s.get("perimeter_m", 0.0)
                qty = s.get("Nwin", 1)

                geom = self._calc_imposts_context(width, height, left, center, right, top)

                ctx = {
                    "width": width,
                    "height": height,
                    "left": left,
                    "center": center,
                    "right": right,
                    "top": top,
                    "sash_width": sash_w,
                    "sash_height": sash_h,
                    "area": area,
                    "perimeter": perimeter,
                    "qty": qty,
                    "nsash": s.get("nsash", 1),
                    "n_sash_active": 1 if s.get("nsash", 1) >= 1 else 0,
                    "n_sash_passive": max(s.get("nsash", 1) - 1, 0),
                    "hinges_per_sash": 3,
                }
                ctx.update(geom)

                qty_fact_total += eval_formula(str(formula), ctx)

            unit_price = safe_float(get_field(row, "цена за", 0.0))
            norm_per_pack = safe_float(get_field(row, "кол-во норм", 0.0))
            unit_pack = str(get_field(row, "ед .норма к упаковке", "") or "").strip()
            unit = str(get_field(row, "ед.", "") or "").strip()
            unit_fact = str(get_field(row, "ед. фактического расхода", "") or "").strip()

            if norm_per_pack > 0:
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

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def _lookup_ref2_rows(self):
        return self.excel.read_records(SHEET_REF2)

    def _find_price_for_filling(self, filling_value):
        """
        Ищем в СПРАВОЧНИК-2 строку с заполнением == filling_value и берем подходящую колонку с 'стоимость' и 'заполн'
        """
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        # Поиск строки по полю 'заполнение'
        chosen = None
        for r in ref2:
            fill_name = get_field(r, "заполнение", "")
            if fill_name and str(fill_name).strip().lower() == str(filling_value).strip().lower():
                chosen = r
                break
        if not chosen:
            return 0.0
        # Поиск колонки с ценой
        for k in chosen.keys():
            if k is None:
                continue
            h = str(k).lower()
            if "стоимость" in h and ("заполн" in h or "запол" in h or "за" in h):
                return safe_float(chosen[k], 0.0)
        # fallback: любая колонка с стоимостью
        for k in chosen.keys():
            if k is None:
                continue
            if "стоимость" in str(k).lower():
                return safe_float(chosen[k], 0.0)
        return 0.0

    def _find_price_for_montage(self, montage_type):
        """
        Ищем цену монтажа по типу монтажа в СПРАВОЧНИК-2.
        Если montage_type == 'Нет' или пусто — вернём 0.
        """
        if not montage_type:
            return 0.0
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        chosen = None
        # сначала ищем строку где поле 'монтаж' или 'тип монтаж' соответствует
        for r in ref2:
            # возможно тип монтажа хранится в поле "монтаж" или в колонке "тип монтаж"
            m_val = get_field(r, "монтаж", None)
            if m_val and str(m_val).strip().lower() == str(montage_type).strip().lower():
                chosen = r
                break
            m_val2 = get_field(r, "тип монтаж", None) or get_field(r, "тип монта", None)
            if m_val2 and str(m_val2).strip().lower() == str(montage_type).strip().lower():
                chosen = r
                break
        # Если строка не найдена, можно попробовать взять первую строку содержащую цену монтажа
        if not chosen:
            for r in ref2:
                # check if any montage price exists in this row
                for k in r.keys():
                    if k is None:
                        continue
                    if "монтаж" in str(k).lower() and "стоимость" in str(k).lower():
                        chosen = r
                        break
                if chosen:
                    break
        if not chosen:
            return 0.0
        # выбираем колонку с монтажной ценой
        for k in chosen.keys():
            if k is None:
                continue
            hk = str(k).lower()
            if "монтаж" in hk and "стоимость" in hk:
                return safe_float(chosen[k], 0.0)
        # fallback — любая стоимость
        for k in chosen.keys():
            if k is None:
                continue
            if "стоимость" in str(k).lower():
                return safe_float(chosen[k], 0.0)
        return 0.0

    def _find_price_for_glass_by_type(self, glass_type):
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        # find row with glass type
        chosen = None
        for r in ref2:
            rt = get_field(r, "тип стеклопак", "") or get_field(r, "тип стеклопакета", "")
            if rt and str(rt).strip().lower() == str(glass_type).strip().lower():
                chosen = r
                break
        if not chosen:
            chosen = ref2[0]
        # get cost field
        for k in chosen.keys():
            if k is None:
                continue
            hk = str(k).lower()
            if "стоимость" in hk and ("стеклопак" in hk or "стеклопакет" in hk or "за м" in hk or "за м²" in hk or "за м2" in hk):
                return safe_float(chosen[k], 0.0)
        # fallback
        for k in chosen.keys():
            if k is None:
                continue
            if "стоимость" in str(k).lower():
                return safe_float(chosen[k], 0.0)
        return 0.0

    def _find_price_for_toning(self):
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        # ищем колонку с тонировкой
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                hk = str(k).lower()
                if "тониров" in hk and "стоимость" in hk:
                    return safe_float(r[k], 0.0)
        return 0.0

    def _find_price_for_handles(self):
        # ищем колонку для цены ручек
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                hk = str(k).lower()
                if ("ручк" in hk or "ручки" in hk) and "стоимость" in hk:
                    return safe_float(r[k], 0.0)
        # fallback: try fields names containing 'ручк'
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "ручк" in str(k).lower():
                    return safe_float(r[k], 0.0)
        return 0.0

    def _find_price_for_closer(self):
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                hk = str(k).lower()
                if ("доводчик" in hk or "доводч" in hk) and "стоимость" in hk:
                    return safe_float(r[k], 0.0)
        # fallback
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "довод" in str(k).lower():
                    return safe_float(r[k], 0.0)
        return 0.0

    def calculate(self,
                  order: dict,
                  total_area_all: float,
                  total_area_glass: float,
                  material_total: float,
                  door_blocks: int = 0,
                  lambr_cost: float = 0.0,
                  handles_qty: int = 0,
                  closer_qty: int = 0):
        # Получаем справочник-2
        ref2_rows = self._lookup_ref2_rows()

        glass_type = order.get("glass_type", "")
        toning = order.get("toning", "Нет")
        assembly = order.get("assembly", "Нет")
        montage = order.get("montage", "Нет")
        handle_type = order.get("handle_type", "")
        door_closer = order.get("door_closer", "Нет")

        price_glass = self._find_price_for_glass_by_type(glass_type)
        price_toning = self._find_price_for_toning()
        # price assembly (сборка) — ищем колонку со словом 'сбор' и 'стоимость'
        price_assembly = 0.0
        if ref2_rows:
            for r in ref2_rows:
                for k in r.keys():
                    if k is None:
                        continue
                    hk = str(k).lower()
                    if "сбор" in hk and "стоимость" in hk:
                        price_assembly = safe_float(r[k], 0.0)
                        break
                if price_assembly:
                    break

        price_montage = self._find_price_for_montage(montage)
        price_handles = self._find_price_for_handles()
        price_closer = self._find_price_for_closer()

        rows = []

        # Стеклопакет
        glass_sum = total_area_glass * price_glass if total_area_glass > 0 else 0.0
        rows.append(["Стеклопакет", price_glass, "за м²", glass_sum])

        # Тонировка
        toning_sum = total_area_glass * price_toning if (toning == "Есть" and total_area_glass > 0) else 0.0
        rows.append(["Тонировка", price_toning, "за м²", toning_sum])

        # Сборка
        assembly_sum = total_area_all * price_assembly if assembly == "Есть" else 0.0
        rows.append(["Сборка", price_assembly, "за м²", assembly_sum])

        # Монтаж
        montage_sum = total_area_all * price_montage if montage != "" and montage.lower() != "нет" else 0.0
        rows.append(["Монтаж (" + str(montage) + ")", price_montage, "за м²", montage_sum])

        # Материалы
        rows.append(["Материал", "-", "-", material_total])

        # Ламбри/Сэндвич (lambr_cost уже рассчитан)
        rows.append(["Панели (Ламбри/Сэндвич)", "-", "-", lambr_cost])

        # Ручки
        handles_sum = price_handles * handles_qty if handles_qty > 0 else 0.0
        rows.append(["Ручки", price_handles, "шт.", handles_sum])

        # Доводчик
        closer_sum = price_closer * closer_qty if closer_qty > 0 else 0.0
        rows.append(["Доводчик", price_closer, "шт.", closer_sum])

        base_sum = (
            glass_sum
            + toning_sum
            + assembly_sum
            + montage_sum
            + material_total
            + lambr_cost
            + handles_sum
            + closer_sum
        )

        ensure_sum = base_sum * 0.6
        rows.append(["Обеспечение (60%)", "", "", ensure_sum])

        total_sum = base_sum + ensure_sum
        extra_rows = [["ИТОГО", "", "", total_sum]]

        # Записываем итоговый лист
        self.excel.clear_and_write(SHEET_FINAL, self.HEADER, rows + extra_rows)
        return rows, total_sum, ensure_sum

# =========================
# Экспорт коммерческого предложения
# =========================

def build_smeta_workbook(order: dict,
                         base_positions: list,
                         lambr_positions: list,
                         total_area: float,
                         total_perimeter: float,
                         total_sum: float) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"

    logo_path = resource_path(LOGO_FILENAME)
    current_row = 1

    # Логотип в левом верхнем углу (A1)
    if os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.height = 80
            img.width = 80
            ws.add_image(img, "A1")
        except Exception as e:
            print(f"Не удалось вставить логотип: {e}")

    contact_col = 3  # колонка C
    ws.cell(row=current_row, column=contact_col, value=COMPANY_NAME)
    current_row += 1
    ws.cell(row=current_row, column=contact_col, value=COMPANY_CITY)
    current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"Тел.: {COMPANY_PHONE}")
    current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"E-mail: {COMPANY_EMAIL}")
    current_row += 1
    if COMPANY_SITE:
        ws.cell(row=current_row, column=contact_col, value=f"Сайт: {COMPANY_SITE}")
        current_row += 1

    current_row += 1
    ws.cell(row=current_row, column=1, value="Коммерческое предложение")
    current_row += 2

    # Общие данные заказа
    ws.cell(row=current_row, column=1, value=f"Заказ № {order.get('order_number','')}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип изделия: {order.get('product_type','')}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Профильная система: {order.get('profile_system','')}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип заполнения (панели): {order.get('filling_mode','')}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип стеклопакета: {order.get('glass_type','')}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тонировка: {order.get('toning','')}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Сборка: {order.get('assembly','')}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Монтаж: {order.get('montage','')}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип ручек: {order.get('handle_type','') or '—'}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Доводчик: {order.get('door_closer','')}")
    current_row += 2

    # Состав позиции / секций
    ws.cell(row=current_row, column=1, value="Состав позиции:")
    current_row += 1

    for idx, p in enumerate(base_positions, start=1):
        ws.cell(
            row=current_row,
            column=1,
            value=(
                f"Позиция {idx}: {order.get('product_type','')}, {p.get('width_mm',0)} × {p.get('height_mm',0)} мм, N = {p.get('Nwin',1)}"
            )
        )
        current_row += 1

    if lambr_positions:
        current_row += 1
        ws.cell(row=current_row, column=1, value="Панели Ламбри / Сэндвич:")
        current_row += 1
        for idx, p in enumerate(lambr_positions, start=1):
            ws.cell(
                row=current_row,
                column=1,
                value=(
                    f"Панель {idx}: {p.get('width_mm',0)} × {p.get('height_mm',0)} мм, N = {p.get('Nwin',1)}"
                )
            )
            current_row += 1

    current_row += 2
    ws.cell(row=current_row, column=1, value=f"Общая площадь: {total_area:.3f} м²")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Суммарный периметр: {total_perimeter:.3f} м")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"ИТОГО к оплате: {total_sum:.2f}")

    # простая попытка выставить ширину колонок
    try:
        for col in ['A','B','C','D','E','F']:
            ws.column_dimensions[col].width = 20
    except Exception:
        pass

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# =========================
# Streamlit UI: main
# =========================

def main():
    st.set_page_config(page_title="Axis Pro GF • Калькулятор", layout="wide")

    excel = ExcelClient(EXCEL_FILE)

    # Авторизация
    user = login_form(excel)
    if not user:
        st.stop()

    st.title("📘 Калькулятор алюминиевых изделий (Axis Pro GF)")
    st.info(f"Пользователь: **{user['login']}**")

    # Загружаем Справочник-2 для типов ручек/стекла/заполнений и цен
    ref2_records = excel.read_records(SHEET_REF2)
    # Собираем варианты заполнений (из колонки 'заполнение' в СПРАВОЧНИК-2)
    filling_types_set = set()
    montage_types_set = set()
    handle_types_set = set()
    glass_types_set = set()

    for row in ref2_records:
        f = get_field(row, "заполнение", None)
        if f is not None:
            filling_types_set.add(str(f).strip())
        # монтаж: возможно есть колонка 'монтаж' или значение в колонке
        m = get_field(row, "монтаж", None)
        if m is not None:
            montage_types_set.add(str(m).strip())
        # ручка
        h = get_field(row, "ручк", None)
        if h is not None:
            handle_types_set.add(str(h).strip())
        # стеклопакет
        g = get_field(row, "тип стеклопак", None) or get_field(row, "тип стеклопакета", None)
        if g is not None:
            glass_types_set.add(str(g).strip())

    # добавим явные опции на случай отсутствия
    if not filling_types_set:
        filling_options = ["Ламбри", "Сэндвич", "Стеклопакет", "Нет"]
    else:
        filling_options = sorted(list(filling_types_set))
        # ensure special option "Нет" present for disabling panel calculations
        if "Нет" not in filling_options:
            filling_options.append("Нет")

    if not montage_types_set:
        montage_options = ["Есть", "Нет"]
    else:
        montage_options = sorted(list(montage_types_set))
        if "Нет" not in montage_options:
            montage_options.append("Нет")

    handle_types = sorted(list(handle_types_set)) if handle_types_set else [""]
    glass_types = sorted(list(glass_types_set)) if glass_types_set else ["двойной"]

    # ---------- Сайдбар: общие данные заказа ----------
    with st.sidebar:
        st.header("Общие данные заказа")

        order_number = st.text_input("Номер заказа", value="")
        product_type = st.selectbox("Тип изделия", ["Окно", "Дверь", "Тамбур"])
        profile_system = st.selectbox(
            "Профильная система",
            [
                "ALG 2030-45C",
                "ALG RUIT 63i",
                "ALG RUIT 73",
            ]
        )

        glass_type = st.selectbox(
            "Тип стеклопакета (цена берётся из справочника-2)",
            glass_types
        )

        # Заполнение панелей: варианты берём из СПРАВОЧНИК-2
        st.markdown("### Режим панелей")
        filling_global = st.selectbox("Заполнение панелей (из справочника-2)", filling_options, index=0)

        toning = st.selectbox("Тонировка", ["Нет", "Есть"])
        assembly = st.selectbox("Сборка", ["Нет", "Есть"])

        # Монтаж: варианты берём из СПРАВОЧНИК-2
        montage = st.selectbox("Монтаж (из СПРАВОЧНИК-2)", montage_options, index=0)

        handle_type = st.selectbox(
            "Тип ручек",
            handle_types,
            index=0 if handle_types else 0
        )
        door_closer = st.selectbox("Доводчик", ["Нет", "Есть"])

        # Кнопка применения заполнения к позициям — будет считаться при клике
        apply_filling = st.button("Применить заполнение панелей к позициям (не-тамбур)")

    # ---------- Основная часть ----------
    col_left, col_right = st.columns([2, 1])

    # Справа: информационный блок
    with col_right:
        st.header("Информация")
        st.info("Заполнения панелей берутся из СПРАВОЧНИК-2. Если выбран режим 'Нет' — панели не учитываются в расчёте.")
        # показываем предупреждение если excel явно повреждён (is_probably_xlsx)
        if not is_probably_xlsx(EXCEL_FILE):
            st.warning("Excel-файл может быть повреждён или не стандартного формата — создан или восстановлен шаблон. Проверьте справочники в Excel.")

    # Левая колонка: позиции
    with col_left:
        st.header("Позиции (габариты изделий)")

        positions_count = st.number_input(
            "Количество позиций",
            min_value=1,
            max_value=10,
            value=1,
            step=1
        )

        base_positions_inputs = []
        lambr_positions_inputs = []
        sections_inputs = []

        for i in range(int(positions_count)):
            st.subheader(f"Позиция {i+1}")
            c1, c2, c3, c4 = st.columns(4)
            width_mm = c1.number_input(f"Ширина, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"w_{i}")
            height_mm = c2.number_input(f"Высота, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"h_{i}")
            left_mm = c3.number_input(f"LEFT, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"l_{i}")
            right_mm = c4.number_input(f"RIGHT, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"r_{i}")

            c5, c6, c7, c8 = st.columns(4)
            center_mm = c5.number_input(f"CENTER, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"c_{i}")
            top_mm = c6.number_input(f"TOP, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"t_{i}")
            sash_width_mm = c7.number_input(f"Ширина створки, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"sw_{i}")
            sash_height_mm = c8.number_input(f"Высота створки, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"sh_{i}")

            nwin = st.number_input(f"Кол-во идентичных рам (N) (поз. {i+1})", min_value=1, value=1, step=1, key=f"nwin_{i}")

            if product_type != "Тамбур":
                base_positions_inputs.append({
                    "width_mm": width_mm,
                    "height_mm": height_mm,
                    "left_mm": left_mm,
                    "center_mm": center_mm,
                    "right_mm": right_mm,
                    "top_mm": top_mm,
                    "sash_width_mm": sash_width_mm if sash_width_mm > 0 else width_mm,
                    "sash_height_mm": sash_height_mm if sash_height_mm > 0 else height_mm,
                    "Nwin": nwin,
                    "filling": filling_global
                })
            else:
                with st.expander(f"Параметры Тамбура — Позиция {i+1}", expanded=False):
                    st.markdown("**Двери**")
                    door_count = st.number_input(f"Сколько отдельных дверей добавить в позицию {i+1}?", min_value=0, value=1, step=1, key=f"tdc_{i}")
                    for d in range(int(door_count)):
                        st.markdown(f"--- Дверь {d+1} ---")
                        dt = st.selectbox(f"Тип двери {d+1} (поз.{i+1})", ["one", "double"], key=f"door_type_{i}_{d}")
                        if dt == "one":
                            dw = st.number_input(f"Ширина двери {d+1} (поз.{i+1}), мм", min_value=0.0, step=10.0, key=f"door_w_{i}_{d}")
                            dh = st.number_input(f"Высота двери {d+1} (поз.{i+1}), мм", min_value=0.0, step=10.0, key=f"door_h_{i}_{d}")
                            sections_inputs.append({
                                "kind": "door",
                                "door_type": "one",
                                "width_mm": dw,
                                "height_mm": dh,
                                "left_mm": 0.0,
                                "center_mm": 0.0,
                                "right_mm": 0.0,
                                "top_mm": 0.0,
                                "sash_width_mm": dw,
                                "sash_height_mm": dh,
                                "Nwin": 1,
                                "filling": "Стеклопакет"
                            })
                        else:
                            dw_l = st.number_input(f"Ширина левой створки {d+1} (поз.{i+1}), мм", min_value=0.0, step=10.0, key=f"door_wl_{i}_{d}")
                            dw_r = st.number_input(f"Ширина правой створки {d+1} (поз.{i+1}), мм", min_value=0.0, step=10.0, key=f"door_wr_{i}_{d}")
                            dh = st.number_input(f"Высота двери {d+1} (поз.{i+1}), мм", min_value=0.0, step=10.0, key=f"door_hd_{i}_{d}")
                            sections_inputs.append({
                                "kind": "door",
                                "door_type": "double",
                                "width_mm": dw_l,
                                "height_mm": dh,
                                "left_mm": 0.0,
                                "center_mm": 0.0,
                                "right_mm": 0.0,
                                "top_mm": 0.0,
                                "sash_width_mm": dw_l,
                                "sash_height_mm": dh,
                                "Nwin": 1,
                                "filling": "Стеклопакет"
                            })
                            sections_inputs.append({
                                "kind": "door",
                                "door_type": "double",
                                "width_mm": dw_r,
                                "height_mm": dh,
                                "left_mm": 0.0,
                                "center_mm": 0.0,
                                "right_mm": 0.0,
                                "top_mm": 0.0,
                                "sash_width_mm": dw_r,
                                "sash_height_mm": dh,
                                "Nwin": 1,
                                "filling": "Стеклопакет"
                            })

                    st.markdown("**Глухие секции (панели)**")
                    panel_count = st.number_input(f"Сколько глухих секций добавить в позицию {i+1}?", min_value=0, value=1, step=1, key=f"tp_{i}")
                    for pidx in range(int(panel_count)):
                        st.markdown(f"--- Глухая секция {pidx+1} ---")
                        pw = st.number_input(f"Ширина глухой секции {pidx+1} (поз.{i+1}), мм", min_value=0.0, step=10.0, key=f"panel_w_{i}_{pidx}")
                        ph = st.number_input(f"Высота глухой секции {pidx+1} (поз.{i+1}), мм", min_value=0.0, step=10.0, key=f"panel_h_{i}_{pidx}")
                        # options for filling per-panel: use filling_options
                        pf = st.selectbox(f"Заполнение глухой секции {pidx+1} (поз.{i+1})", options=filling_options, index=0, key=f"panel_fill_{i}_{pidx}")
                        sections_inputs.append({
                            "kind": "panel",
                            "width_mm": pw,
                            "height_mm": ph,
                            "left_mm": 0.0,
                            "center_mm": 0.0,
                            "right_mm": 0.0,
                            "top_mm": 0.0,
                            "sash_width_mm": pw,
                            "sash_height_mm": ph,
                            "Nwin": 1,
                            "filling": pf
                        })

                    base_positions_inputs.append({
                        "width_mm": width_mm,
                        "height_mm": height_mm,
                        "left_mm": left_mm,
                        "center_mm": center_mm,
                        "right_mm": right_mm,
                        "top_mm": top_mm,
                        "sash_width_mm": sash_width_mm if sash_width_mm > 0 else width_mm,
                        "sash_height_mm": sash_height_mm if sash_height_mm > 0 else height_mm,
                        "Nwin": nwin
                    })

        # Non-tambur: дополнительные панели (ламбри/сэндвич)
        if product_type != "Тамбур":
            st.subheader("Панели (Ламбри/Сэндвич) — дополнительные")
            panel_count_ls = st.number_input("Количество дополнительных панелей", min_value=0, value=0, step=1, key="ls_panel_count")
            for i in range(int(panel_count_ls)):
                st.markdown(f"**Панель {i+1}**")
                p1, p2, p3 = st.columns(3)
                w = p1.number_input(f"Ширина панели {i+1}, мм", min_value=0.0, step=10.0, key=f"ls_w_{i}")
                h = p2.number_input(f"Высота панели {i+1}, мм", min_value=0.0, step=10.0, key=f"ls_h_{i}")
                q = p3.number_input(f"N (панель {i+1})", min_value=1, value=1, step=1, key=f"ls_q_{i}")
                lambr_positions_inputs.append({
                    "width_mm": w,
                    "height_mm": h,
                    "Nwin": q,
                    "left_mm": 0.0,
                    "center_mm": 0.0,
                    "right_mm": 0.0,
                    "top_mm": 0.0,
                    "sash_width_mm": w,
                    "sash_height_mm": h,
                    "filling": filling_global
                })

    # ---------- Выбор материалов при дублях ----------
    st.header("🧾 Выбор материалов при дублях (если в справочнике несколько товаров на один элемент)")
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
    calc_button = st.button("💾 Сохранить в Excel и выполнить расчёт")

    if calc_button:
        if not order_number.strip():
            st.error("Введите номер заказа.")
            st.stop()

        # Собираем base_positions и sections
        base_positions = []
        for p in base_positions_inputs:
            if p["width_mm"] <= 0 or p["height_mm"] <= 0:
                st.error("Во всех позициях ширина и высота должны быть больше 0.")
                st.stop()
            area_m2 = (p["width_mm"] * p["height_mm"]) / 1_000_000.0
            perimeter_m = 2 * (p["width_mm"] + p["height_mm"]) / 1000.0
            base_positions.append({
                **p,
                "area_m2": area_m2,
                "perimeter_m": perimeter_m,
            })

        lambr_positions = []
        for p in lambr_positions_inputs:
            if p["width_mm"] > 0 and p["height_mm"] > 0:
                area_m2 = (p["width_mm"] * p["height_mm"]) / 1_000_000.0
                perimeter_m = 2 * (p["width_mm"] + p["height_mm"]) / 1000.0
                lambr_positions.append({
                    **p,
                    "area_m2": area_m2,
                    "perimeter_m": perimeter_m,
                })

        # sections
        sections = []
        if product_type == "Тамбур":
            for s in sections_inputs:
                if s["width_mm"] <= 0 or s["height_mm"] <= 0:
                    st.warning("Одна из секций тамбура имеет 0 ширину или высоту и будет пропущена.")
                    continue
                area_m2 = (s["width_mm"] * s["height_mm"]) / 1_000_000.0
                perimeter_m = 2 * (s["width_mm"] + s["height_mm"]) / 1000.0
                sections.append({
                    **s,
                    "area_m2": area_m2,
                    "perimeter_m": perimeter_m
                })
        else:
            for p in base_positions:
                sections.append({
                    **p,
                    "area_m2": p["area_m2"],
                    "perimeter_m": p["perimeter_m"],
                    "filling": p.get("filling", filling_global)
                })
            for p in lambr_positions:
                sections.append({
                    **p,
                    "area_m2": p["area_m2"],
                    "perimeter_m": p["perimeter_m"],
                    "filling": p.get("filling", filling_global)
                })

        # Apply filling global if button pressed and non-tambur
        if apply_filling and product_type != "Тамбур":
            for s in sections:
                s["filling"] = filling_global

        # Если заполнение панелей выбрано как "Нет" в СПРАВОЧНИК-2 — игнорируем панели в расчетах
        # Проверим, есть ли в СПРАВОЧНИК-2 запись 'Нет' для заполнения
        disable_panels = False
        # Если глобально выбрано "Нет", то панели игнорируются
        if str(filling_global).strip().lower() == "нет":
            disable_panels = True

        # --- Gabarit ---
        gab_calc = GabaritCalculator(excel)
        gabarit_rows, total_area_gab, total_perimeter_gab = gab_calc.calculate(
            {"product_type": product_type},
            sections
        )

        # --- Materials ---
        mat_calc = MaterialCalculator(excel)
        material_rows, material_total, total_area_mat = mat_calc.calculate(
            {"product_type": product_type, "profile_system": profile_system},
            sections,
            selected_duplicates
        )

        # --- Lambr/Sandwich calculation (по хлыстам 6 м) ---
        linear_meters = 0.0
        for s in sections:
            # игнорируем панели если выключено
            if disable_panels and (s.get("filling") in ("Ламбри", "Сэндвич") or str(s.get("filling")).strip().lower() == "ламбри" or str(s.get("filling")).strip().lower() == "сэндвич"):
                continue
            if s.get("filling") in ("Ламбри", "Сэндвич"):
                linear_meters += s.get("perimeter_m", 0.0) * s.get("Nwin", 1)

        count_hlyst = math.ceil(linear_meters / 6.0) if linear_meters > 0 else 0

        # цена за метр заполнения из СПРАВОЧНИК-2
        # используем FinalCalculator helper
        fin_calc = FinalCalculator(excel)
        price_per_meter_fill = 0.0
        # если глобальное заполнение == "Нет", цена не нужна
        if not disable_panels and linear_meters > 0:
            # берем цену по названию filling_global
            price_per_meter_fill = fin_calc._find_price_for_filling(filling_global)

        if price_per_meter_fill <= 0 and linear_meters > 0:
            st.warning("Не найдена цена за заполнение (Ламбри/Сэндвич) в СПРАВОЧНИК-2. Установлена 0.")

        price_per_hlyst = price_per_meter_fill * 6.0
        lambr_cost = count_hlyst * price_per_hlyst

        # --- Areas for glass etc.
        total_area_glass = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections if s.get("filling") == "Стеклопакет")
        total_area_all = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)

        # --- Doors / handles / closer counts
        door_blocks = 0
        double_pairs = 0
        for s in sections:
            if s.get("kind") == "door":
                if s.get("door_type") == "one":
                    door_blocks += 1
                elif s.get("door_type") == "double":
                    double_pairs += 1
        if double_pairs:
            # double_pairs counts sections labeled double (we inserted two sections per double door)
            door_blocks += double_pairs / 2.0
        door_blocks = int(math.ceil(door_blocks))

        # handles_count: number of leaves (sections with kind=='door'), but if double door blocks counted as 1 block with 2 leaves, we approximate:
        handles_count = sum(1 for s in sections if s.get("kind") == "door")
        # closer count: one per block
        closer_count = door_blocks

        # --- Final calculation ---
        final_rows, total_sum, ensure_sum = fin_calc.calculate(
            {
                "product_type": product_type,
                "glass_type": glass_type,
                "toning": toning,
                "assembly": assembly,
                "montage": montage,
                "handle_type": handle_type,
                "door_closer": door_closer
            },
            total_area_all=total_area_all,
            total_area_glass=total_area_glass,
            material_total=material_total,
            door_blocks=door_blocks,
            lambr_cost=lambr_cost,
            handles_qty=handles_count,
            closer_qty=closer_count
        )

        st.success("Расчёт выполнен. Результаты ниже.")

        tab1, tab2, tab3 = st.tabs(["Габариты", "Материалы", "Итоговый расчет"])

        with tab1:
            st.subheader("Расчет по габаритам")
            # Отображаем габаритные строки (если нужны)
            if gabarit_rows:
                gab_disp = [{"Тип элемента": t, "Фактическое значение": v} for t, v in gabarit_rows]
                st.dataframe(gab_disp, use_container_width=True)
            # Внизу — только общая площадь и суммарный периметр изделия (по вашему запросу)
            st.write(f"Общая площадь (служебная): **{total_area_gab:.3f} м²**")
            st.write(f"Суммарный периметр изделия: **{total_perimeter_gab:.3f} м**")

        with tab2:
            st.subheader("Расчёт материалов")
            if material_rows:
                mat_disp = []
                for r in material_rows:
                    mat_disp.append({
                        "Тип изделия": r[0],
                        "Система профиля": r[1],
                        "Тип элемента": r[2],
                        "Артикул": r[3],
                        "Товар": r[4],
                        "Ед.": r[5],
                        "Цена за ед.": round(r[6], 2),
                        "Ед. факт. расхода": r[7],
                        "Кол-во факт. расхода": round(r[8], 3),
                        "Норма к упаковке": r[9],
                        "Ед. к отгрузке": r[10],
                        "Кол-во к отгрузке": round(r[11], 3),
                        "Сумма": round(r[12], 2),
                    })
                st.dataframe(mat_disp, use_container_width=True)
            st.write(f"Итого по материалам: **{material_total:.2f}**")
            st.write(f"Панели (ламбри/сэндвич) — линейная длина: **{linear_meters:.3f} м**, Хлыстов(6м): **{count_hlyst}**, Цена/м: **{price_per_meter_fill:.2f}**, Итого: **{lambr_cost:.2f}**")

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

        # --- Сохраняем в ЗАПРОСЫ ---
        rows_for_form = []
        pos_index = 1
        for p in base_positions:
            rows_for_form.append([
                order_number,
                pos_index,
                product_type,
                "",  # вид изделия
                "",  # створки
                profile_system,
                glass_type,
                filling_global if product_type != "Тамбур" else "Тамбур",
                p["width_mm"],
                p["height_mm"],
                p.get("left_mm", 0.0),
                p.get("center_mm", 0.0),
                p.get("right_mm", 0.0),
                p.get("top_mm", 0.0),
                p.get("sash_width_mm", p["width_mm"]),
                p.get("sash_height_mm", p["height_mm"]),
                p["Nwin"],
                toning,
                assembly,
                montage,
                handle_type,
                door_closer,
            ])
            pos_index += 1

        for row in rows_for_form:
            excel.append_form_row(row)

        # --- Коммерческий Excel ---
        smeta_bytes = build_smeta_workbook(
            order={
                "order_number": order_number,
                "product_type": product_type,
                "profile_system": profile_system,
                "filling_mode": filling_global,
                "glass_type": glass_type,
                "toning": toning,
                "assembly": assembly,
                "montage": montage,
                "handle_type": handle_type,
                "door_closer": door_closer,
                "sections": sections
            },
            base_positions=base_positions,
            lambr_positions=lambr_positions,
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


if __name__ == "__main__":
    main()
