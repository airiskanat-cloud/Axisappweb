import math
import os
import sys
import zipfile
from io import BytesIO

import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.drawing.image import Image as XLImage
import ast
import operator as op
import logging
import json

# =========================
# НАСТРОЙКИ / КОНСТАНТЫ
# =========================

DEBUG = False
logger = logging.getLogger(__name__)

def resource_path(relative_path: str) -> str:
    try:
        if hasattr(sys, "_MEIPASS"):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(os.path.dirname(__file__))
    except Exception:
        base_path = os.getcwd()
    return os.path.join(base_path, relative_path)

# Хранение данных вне каталога проекта (Streamlit не сбрасывает сессию)
DATA_DIR = os.path.join(os.path.expanduser("~"), ".axis_app_data")
os.makedirs(DATA_DIR, exist_ok=True)

import shutil

# Путь к шаблону в репозитории
BUNDLED_TEMPLATE = resource_path("axis_pro_gf.xlsx")  # если в корне

# Если файл не существует в DATA_DIR — копируем шаблон
if os.path.exists(BUNDLED_TEMPLATE) and not os.path.exists(EXCEL_FILE):
    try:
        shutil.copyfile(BUNDLED_TEMPLATE, EXCEL_FILE)
    except Exception as e:
        logger.error(f"Error copying template: {e}")


TEMPLATE_EXCEL_NAME = "axis_pro_gf.xlsx"
EXCEL_FILE = os.path.join(DATA_DIR, TEMPLATE_EXCEL_NAME)
SESSION_FILE = os.path.join(DATA_DIR, "session_user.json")

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# Шапка записи для ЗАПРОСЫ
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

# Брендинг для коммерческого предложения
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
    if k is None:
        return None
    s = str(k)
    s = s.replace("\xa0", " ")
    s = " ".join(s.split())
    return s.strip()

def _clean_cell_val(v):
    if v is None:
        return ""
    s = str(v)
    s = s.replace("\xa0", " ").strip()
    return s

def safe_float(value, default=0.0):
    try:
        if value is None:
            return default
        s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
        if s == "":
            return default
        return float(s)
    except:
        return default

def safe_int(value, default=0):
    try:
        if value is None:
            return default
        s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
        if s == "":
            return default
        return int(float(s))
    except:
        return default

def get_field(row: dict, needle: str, default=None):
    needle = (needle or "").lower().strip()
    for k, v in row.items():
        if k and needle in str(k).lower():
            return v
    return default

# =========================
# БЕЗОПАСНЫЙ EVAL (ФОРМУЛЫ)
# =========================

_allowed_ops = {
    ast.Add: op.add,
    ast.Sub: op.sub,
    ast.Mult: op.mul,
    ast.Div: op.truediv,
    ast.Pow: op.pow,
    ast.USub: op.neg,
    ast.UAdd: op.pos,
    ast.Mod: op.mod,
    ast.FloorDiv: op.floordiv,
    ast.Lt: op.lt,
    ast.Gt: op.gt,
    ast.LtE: op.le,
    ast.GtE: op.ge,
    ast.Eq: op.eq,
    ast.NotEq: op.ne,
    ast.And: lambda a,b: a and b,
    ast.Or:  lambda a,b: a or b,
}

def _eval_ast(node, names):
    if isinstance(node, ast.Expression):
        return _eval_ast(node.body, names)

    if isinstance(node, ast.Constant):
        return node.value

    if isinstance(node, ast.Num):  # старые версии
        return node.n

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
        # math.x
        if isinstance(func, ast.Attribute) and isinstance(func.value, ast.Name) and func.value.id == "math":
            fname = func.attr
            if hasattr(math, fname):
                args = [_eval_ast(a, names) for a in node.args]
                return getattr(math, fname)(*args)

        # max/min
        if isinstance(func, ast.Name) and func.id in ("max", "min"):
            args = [_eval_ast(a, names) for a in node.args]
            return globals()[func.id](*args)

        raise ValueError("Разрешены только math.*, max, min")

    if isinstance(node, ast.Compare):
        if len(node.ops) != 1:
            raise ValueError("Сложные сравнения запрещены")
        left = _eval_ast(node.left, names)
        right = _eval_ast(node.comparators[0], names)
        fn = _allowed_ops.get(type(node.ops[0]))
        if fn: return fn(left, right)

    raise ValueError(f"Недопустимый элемент формулы: {type(node).__name__}")

def safe_eval_formula(formula: str, context: dict) -> float:
    formula = (formula or "").strip()
    if not formula:
        return 0.0

    names = {
        **context,
        "math": math,
        "min": min,
        "max": max,
    }

    try:
        node = ast.parse(formula, mode="eval")
        return float(_eval_ast(node, names))
    except:
        return 0.0


# =========================
# EXCEL CLIENT
# =========================

def is_probably_xlsx(path: str) -> bool:
    try:
        if not os.path.exists(path):
            return False
        if os.path.getsize(path) < 3000:
            return False
        with zipfile.ZipFile(path, "r") as z:
            return (
                "[Content_Types].xml" in z.namelist()
                and "xl/workbook.xml" in z.namelist()
            )
    except:
        return False


class ExcelClient:
    def __init__(self, filename: str):
        self.filename = filename
        if not os.path.exists(self.filename):
            self._create_template()
        self.load()

    def _create_template(self):
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        wb.create_sheet(SHEET_FORM)
        wb.create_sheet(SHEET_REF1)
        wb.create_sheet(SHEET_REF2)
        wb.create_sheet(SHEET_REF3)
        wb.create_sheet(SHEET_USERS)
        wb.save(self.filename)

    def load(self):
        try:
            self.wb = load_workbook(self.filename, data_only=True)
        except:
            # если файл повреждён — пересоздаём
            try:
                os.remove(self.filename)
            except:
                pass
            self._create_template()
            self.wb = load_workbook(self.filename, data_only=True)

    def save(self):
        try:
            self.wb.save(self.filename)
        except Exception as e:
            print("Ошибка сохранения:", e)

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
        header_raw = rows[0]
        header = []
        used = {}

        for h in header_raw:
            key = normalize_key(h)
            if key in used:
                used[key] += 1
                key = f"{key}_{used[key]}"
            else:
                used[key] = 1
            header.append(key)

        records = []
        for r in rows[1:]:
            if all(v is None for v in r):
                continue
            row = {}
            for i, k in enumerate(header):
                row[k] = r[i]
            records.append(row)
        return records

    def clear_and_write(self, sheet_name: str, header: list, rows: list):
        ws = self.ws(sheet_name)
        try:
            ws.delete_rows(1, ws.max_row)
        except:
            pass

        if header:
            ws.append(header)
        for row in rows:
            ws.append(row)
        self.save()

    def append_form_row(self, row: list):
        ws = self.ws(SHEET_FORM)
        if ws.max_row == 1 and not any(ws[1]):
            ws.append(FORM_HEADER)
        ws.append(row)
        self.save()


# =========================
# ПОЛЬЗОВАТЕЛИ (ЛОГИН)
# =========================

def load_users(excel: ExcelClient):
    excel.load()
    rows = excel.read_records(SHEET_USERS)
    users = {}

    for r in rows:
        login = _clean_cell_val(get_field(r, "логин", "")).lower()
        pwd = _clean_cell_val(get_field(r, "парол", "")).replace("*", "").strip()
        role = _clean_cell_val(get_field(r, "роль", ""))

        if login:
            users[login] = {"password": pwd, "role": role, "_raw_login": login}

    return users


def login_form(excel: ExcelClient):
    # если сессия есть — пользователь уже вошёл
    if "current_user" in st.session_state:
        return st.session_state["current_user"]

    # пробуем загрузить из файла (стойкая авторизация)
    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE, "r", encoding="utf-8") as sf:
                st.session_state["current_user"] = json.load(sf)
                return st.session_state["current_user"]
        except:
            pass

    st.sidebar.title("🔐 Вход в систему")
    with st.sidebar.form("login_form"):
        login = st.text_input("Логин")
        password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти")

    users = load_users(excel)

    if submitted:
        entered_login = (login or "").strip().lower()
        entered_pass = (password or "").replace("\xa0", "").strip()

        user = users.get(entered_login)

        if user:
            real_pass = (user["password"] or "").strip().replace("\xa0", "")
            if entered_pass == real_pass:
                st.session_state["current_user"] = {
                    "login": user["_raw_login"],
                    "role": user["role"],
                }
                # сохраняем сессию
                try:
                    with open(SESSION_FILE, "w", encoding="utf-8") as sf:
                        json.dump(st.session_state["current_user"], sf, ensure_ascii=False)
                except:
                    pass

                st.sidebar.success(f"Привет, {user['_raw_login']}!")
                return st.session_state["current_user"]

        st.sidebar.error("Неверный логин или пароль")

    return None
# =========================
# CALCULATORS: GABARIT / MATERIAL / FINAL
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
        total_area = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
        total_perimeter = sum(s.get("perimeter_m", 0.0) * s.get("Nwin", 1) for s in sections)

        if not ref_rows:
            # нет справочника — просто вернуть суммы
            return [], total_area, total_perimeter

        gabarit_values = []
        for row in ref_rows:
            type_elem = get_field(row, "тип элемент", "")
            formula = get_field(row, "формула_python", "")
            if not type_elem or not formula:
                continue
            total_value = 0.0
            for s in sections:
                if s.get("kind") == "door":
                    width = s.get("frame_width_mm", 0.0)
                    height = s.get("frame_height_mm", 0.0)
                    sash_w = s.get("leaves", [{}])[0].get("width_mm", width) if s.get("leaves") else width
                    sash_h = s.get("leaves", [{}])[0].get("height_mm", height) if s.get("leaves") else height
                else:
                    width = s.get("width_mm", 0.0)
                    height = s.get("height_mm", 0.0)
                    sash_w = s.get("sash_width_mm", width)
                    sash_h = s.get("sash_height_mm", height)

                left = s.get("left_mm", 0.0)
                center = s.get("center_mm", 0.0)
                right = s.get("right_mm", 0.0)
                top = s.get("top_mm", 0.0)
                area = s.get("area_m2", 0.0)
                perimeter = s.get("perimeter_m", 0.0)
                qty = s.get("Nwin", 1)

                geom = self._calc_imposts_context(width, height, left, center, right, top)

                nsash = s.get("n_leaves", len(s.get("leaves", [])) or 1)
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
                total_value += safe_eval_formula(str(formula), ctx)
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

            if row_type:
                if str(row_type).strip().lower() != order.get("product_type", "").strip().lower():
                    continue

            if row_profile:
                if str(row_profile).strip().lower() != order.get("profile_system", "").strip().lower():
                    continue

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
                if s.get("kind") == "door":
                    width = s.get("frame_width_mm", 0.0)
                    height = s.get("frame_height_mm", 0.0)
                else:
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
                    "nsash": s.get("n_leaves", len(s.get("leaves", [])) or 1),
                    "n_sash_active": 1 if s.get("n_leaves", len(s.get("leaves", [])) or 1) >= 1 else 0,
                    "n_sash_passive": max(s.get("n_leaves", len(s.get("leaves", [])) or 1) - 1, 0),
                    "hinges_per_sash": 3,
                }
                ctx.update(geom)

                qty_fact_total += safe_eval_formula(str(formula), ctx)

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
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        fv = str(filling_value or "").replace("\xa0", " ").strip().lower()
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "панел" in str(k).lower() or "заполн" in str(k).lower():
                    v = r[k]
                    if v is None:
                        continue
                    if str(v).replace("\xa0", " ").strip().lower() == fv:
                        for kk in r.keys():
                            if kk is None:
                                continue
                            if "стоимость" in str(kk).lower():
                                return safe_float(r[kk], 0.0)
        return 0.0

    def _find_price_for_montage(self, montage_type):
        if not montage_type:
            return 0.0
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        mt = str(montage_type or "").replace("\xa0", " ").strip().lower()
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "монтаж" in str(k).lower() and "стоимость" in str(k).lower():
                    return safe_float(r[k], 0.0)
        return 0.0

    def _find_price_for_glass_by_type(self, glass_type):
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        gt = str(glass_type or "").replace("\xa0", " ").strip().lower()
        chosen = None
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "тип стеклопак" in str(k).lower() or "тип стеклопакета" in str(k).lower():
                    v = r[k]
                    if v and str(v).strip().lower() == gt:
                        chosen = r
                        break
            if chosen:
                break
        if not chosen:
            for r in ref2:
                for k in r.keys():
                    if k is None:
                        continue
                    if "стоимость" in str(k).lower() and ("стеклопак" in str(k).lower() or "за м" in str(k).lower()):
                        return safe_float(r[k], 0.0)
            return 0.0
        for k in chosen.keys():
            if k is None:
                continue
            hk = str(k).lower()
            if "стоимость" in hk and ("стеклопак" in hk or "за м" in hk or "за м²" in hk or "за м2" in hk):
                return safe_float(chosen[k], 0.0)
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
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                hk = str(k).lower()
                if "тониров" in hk and "стоимость" in hk:
                    return safe_float(r[k], 0.0)
        return 0.0

    def _find_price_for_handles(self):
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
        ref2_rows = self._lookup_ref2_rows()

        glass_type = order.get("glass_type", "")
        toning = order.get("toning", "Нет")
        assembly = order.get("assembly", "Нет")
        montage = order.get("montage", "Нет")
        handle_type = order.get("handle_type", "")
        door_closer = order.get("door_closer", "Нет")

        price_glass = self._find_price_for_glass_by_type(glass_type)
        price_toning = self._find_price_for_toning()
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

        glass_sum = total_area_glass * price_glass if total_area_glass > 0 else 0.0
        rows.append(["Стеклопакет", price_glass, "за м²", glass_sum])

        toning_sum = total_area_glass * price_toning if (toning == "Есть" and total_area_glass > 0) else 0.0
        rows.append(["Тонировка", price_toning, "за м²", toning_sum])

        assembly_sum = total_area_all * price_assembly if assembly == "Есть" else 0.0
        rows.append(["Сборка", price_assembly, "за м²", assembly_sum])

        montage_sum = total_area_all * price_montage if montage != "" and montage.lower() != "нет" else 0.0
        rows.append(["Монтаж (" + str(montage) + ")", price_montage, "за м²", montage_sum])

        rows.append(["Материал", "-", "-", material_total])
        rows.append(["Панели (Ламбри/Сэндвич)", "-", "-", lambr_cost])

        handles_sum = price_handles * handles_qty if handles_qty > 0 else 0.0
        rows.append(["Ручки", price_handles, "шт.", handles_sum])

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
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"

    logo_path = resource_path(LOGO_FILENAME)
    current_row = 1

    if os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.height = 80
            img.width = 80
            ws.add_image(img, "A1")
        except:
            pass

    contact_col = 3
    ws.cell(row=current_row, column=contact_col, value=COMPANY_NAME); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=COMPANY_CITY); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"Тел.: {COMPANY_PHONE}"); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"E-mail: {COMPANY_EMAIL}"); current_row += 1
    if COMPANY_SITE:
        ws.cell(row=current_row, column=contact_col, value=f"Сайт: {COMPANY_SITE}"); current_row += 1

    current_row += 1
    ws.cell(row=current_row, column=1, value="Коммерческое предложение"); current_row += 2

    ws.cell(row=current_row, column=1, value=f"Заказ № {order.get('order_number','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип изделия: {order.get('product_type','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Профильная система: {order.get('profile_system','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип заполнения (панели): {order.get('filling_mode','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип стеклопакета: {order.get('glass_type','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тонировка: {order.get('toning','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Сборка: {order.get('assembly','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Монтаж: {order.get('montage','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип ручек: {order.get('handle_type','') or '—'}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Доводчик: {order.get('door_closer','')}"); current_row += 2

    ws.cell(row=current_row, column=1, value="Состав позиции:"); current_row += 1

    for idx, p in enumerate(base_positions, start=1):
        w = p.get('width_mm', p.get('frame_width_mm', 0))
        h = p.get('height_mm', p.get('frame_height_mm', 0))
        fill = p.get('filling', '') or (p.get('leaves', [{}])[0].get('filling', '') if p.get('leaves') else '')
        ws.cell(row=current_row, column=1, value=f"Позиция {idx}: {order.get('product_type','')}, {w} × {h} мм, N = {p.get('Nwin',1)}, filling={fill}")
        current_row += 1

    if lambr_positions:
        current_row += 1
        ws.cell(row=current_row, column=1, value="Панели Ламбри / Сэндвич:"); current_row += 1
        for idx, p in enumerate(lambr_positions, start=1):
            ws.cell(row=current_row, column=1, value=f"Панель {idx}: {p.get('width_mm',0)} × {p.get('height_mm',0)} мм, N = {p.get('Nwin',1)}, filling={p.get('filling','')}")
            current_row += 1

    current_row += 2
    ws.cell(row=current_row, column=1, value=f"Общая площадь: {total_area:.3f} м²"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Суммарный периметр: {total_perimeter:.3f} м"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"ИТОГО к оплате: {total_sum:.2f}")

    try:
        for col in ['A','B','C','D','E','F']:
            ws.column_dimensions[col].width = 20
    except:
        pass

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================
# STREAMLIT UI: main
# =========================

def ensure_session_state():
    if "tam_door_count" not in st.session_state:
        st.session_state["tam_door_count"] = 0
    if "tam_panel_count" not in st.session_state:
        st.session_state["tam_panel_count"] = 0
    if "sections_inputs" not in st.session_state:
        st.session_state["sections_inputs"] = []

def main():
    st.set_page_config(page_title="Axis Pro GF • Калькулятор", layout="wide")
    ensure_session_state()

    excel = ExcelClient(EXCEL_FILE)

    # восстановление сессии, если файл изменился на диске
    if "current_user" not in st.session_state:
        try:
            if os.path.exists(SESSION_FILE):
                with open(SESSION_FILE, "r", encoding="utf-8") as sf:
                    st.session_state["current_user"] = json.load(sf)
        except:
            pass

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
        f = _clean_for_set(get_field(row, "панел") or get_field(row, "заполн") or get_field(row, "заполнение"))
        if f:
            filling_types_set.add(f)
        m = _clean_for_set(get_field(row, "монтаж", None))
        if m:
            montage_types_set.add(m)
        h = _clean_for_set(get_field(row, "ручк", None))
        if h:
            handle_types_set.add(h)
        g = _clean_for_set(get_field(row, "тип стеклопак", None) or get_field(row, "тип стеклопакета", None))
        if g:
            glass_types_set.add(g)

    # Жёстко задаём опции панелей: только три варианта
    filling_options_for_panels = ["Ламбри без термо", "Ламбри с термо", "Стеклопакет"]

    if not montage_types_set:
        montage_options = ["Есть", "Нет"]
    else:
        montage_options = sorted(list(montage_types_set))
        if "Нет" not in montage_options:
            montage_options.append("Нет")

    handle_types = sorted(list(handle_types_set)) if handle_types_set else [""]
    glass_types = sorted(list(glass_types_set)) if glass_types_set else ["двойной"]

    # ---------- Sidebar: общие данные ----------
    with st.sidebar:
        st.header("Общие данные заказа")
        order_number = st.text_input("Номер заказа", value="")
        product_type = st.selectbox("Тип изделия", ["Окно", "Дверь", "Тамбур"])
        profile_system = st.selectbox("Профильная система", ["ALG 2030-45C", "ALG RUIT 63i", "ALG RUIT 73"])
        glass_type = st.selectbox("Тип стеклопакета (цена из СПРАВОЧНИК-2)", glass_types)
        st.markdown("### Прочее")
        toning = st.selectbox("Тонировка", ["Нет", "Есть"])
        assembly = st.selectbox("Сборка", ["Нет", "Есть"])
        montage = st.selectbox("Монтаж (из СПРАВОЧНИК-2)", montage_options, index=0)
        handle_type = st.selectbox("Тип ручек", handle_types, index=0 if handle_types else 0)
        door_closer = st.selectbox("Доводчик", ["Нет", "Есть"])

    col_left, col_right = st.columns([2, 1])

    with col_right:
        st.header("Информация")
        st.info("Тамбур детализируется отдельными секциями: дверные блоки и глухие панели.")
        if not is_probably_xlsx(EXCEL_FILE):
            st.warning("Excel-файл справочников может быть не в порядке — проверь СПРАВОЧНИК-2/1/3.")
        if DEBUG:
            st.write("DEBUG ref2:", ref2_records[:5])
            st.write("DEBUG sections_inputs:", st.session_state.get("sections_inputs", []))

    with col_left:
        st.header("Позиции (окна/двери/тамбур)")
        positions_count = st.number_input("Количество позиций", min_value=1, max_value=10, value=1, step=1)

        base_positions_inputs = []
        lambr_positions_inputs = []

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
                    "filling": "Стеклопакет",
                    "kind": "window"
                })
            else:
                st.markdown("Позиция тамбура: не задаём общий габарит. Добавляй дверные блоки и панели ниже.")

        # дополнительные панели для не-тамбур
        if product_type != "Тамбур":
            st.subheader("Панели (Ламбри/Сэндвич) — дополнительные")
            panel_count_ls = st.number_input("Количество дополнительных панелей", min_value=0, value=0, step=1, key="ls_panel_count")
            for i in range(int(panel_count_ls)):
                st.markdown(f"**Панель {i+1}**")
                p1, p2, p3 = st.columns(3)
                w = p1.number_input(f"Ширина панели {i+1}, мм", min_value=0.0, step=10.0, key=f"ls_w_{i}")
                h = p2.number_input(f"Высота панели {i+1}, мм", min_value=0.0, step=10.0, key=f"ls_h_{i}")
                q = p3.number_input(f"N (панель {i+1})", min_value=1, value=1, step=1, key=f"ls_q_{i}")
                fill_opt = st.selectbox(f"Заполнение панели {i+1}", options=filling_options_for_panels, index=0, key=f"ls_fill_{i}")
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
                    "filling": fill_opt
                })

    # ---------- Параметры тамбура (новая логика) ----------
    if product_type == "Тамбур":
        st.markdown("---")
        st.header("Параметры тамбура (дверные блоки и глухие панели)")

        c_add = st.columns([1,1,6])
        if c_add[0].button("Добавить дверной блок"):
            st.session_state["tam_door_count"] += 1
        if c_add[1].button("Добавить глухую секцию"):
            st.session_state["tam_panel_count"] += 1

        # Дверные блоки
        for i in range(st.session_state.get("tam_door_count", 0)):
            with st.expander(f"Дверной блок #{i+1}", expanded=False):
                name = st.text_input(f"Название блока #{i+1}", value=f"Дверной блок {i+1}", key=f"door_name_{i}")
                count = st.number_input(f"Кол-во одинаковых блоков #{i+1}", min_value=1, value=1, key=f"door_count_{i}")
                dtype = st.selectbox(f"Тип двери #{i+1}", ["Одностворчатая","Двухстворчатая"], key=f"door_type_{i}")
                frame_w = st.number_input(f"Ширина рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, key=f"frame_w_{i}")
                frame_h = st.number_input(f"Высота рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, key=f"frame_h_{i}")
                left = st.number_input(f"LEFT, мм #{i+1}", min_value=0.0, step=10.0, key=f"left_{i}")
                center = st.number_input(f"CENTER, мм #{i+1}", min_value=0.0, step=10.0, key=f"center_{i}")
                right = st.number_input(f"RIGHT, мм #{i+1}", min_value=0.0, step=10.0, key=f"right_{i}")
                top = st.number_input(f"TOP, мм #{i+1}", min_value=0.0, step=10.0, key=f"top_{i}")

                default_leaves = 1 if dtype == "Одностворчатая" else 2
                n_leaves = st.number_input(f"Кол-во створок #{i+1}", min_value=1, value=default_leaves, key=f"n_leaves_{i}")

                leaves = []
                for L in range(int(n_leaves)):
                    lw = st.number_input(f"Ширина створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, key=f"leaf_w_{i}_{L}")
                    lh = st.number_input(f"Высота створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, key=f"leaf_h_{i}_{L}")
                    fill = st.selectbox(f"Заполнение створки {L+1} — блок {i+1}", options=filling_options_for_panels, index=0, key=f"leaf_fill_{i}_{L}")
                    leaves.append({"width_mm": lw, "height_mm": lh, "filling": fill})

                if st.button(f"Добавить/обновить дверной блок #{i+1} в секциях", key=f"save_door_{i}"):
                    # обновляем если есть с таким именем, иначе добавляем
                    found = False
                    for idx, s in enumerate(st.session_state["sections_inputs"]):
                        if s.get("block_name") == name and s.get("kind") == "door":
                            st.session_state["sections_inputs"][idx] = {
                                "kind": "door",
                                "block_name": name,
                                "count": int(count),
                                "frame_width_mm": frame_w,
                                "frame_height_mm": frame_h,
                                "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                                "n_leaves": int(n_leaves),
                                "leaves": leaves,
                                "Nwin": int(count)
                            }
                            found = True
                            break
                    if not found:
                        st.session_state["sections_inputs"].append({
                            "kind": "door",
                            "block_name": name,
                            "count": int(count),
                            "frame_width_mm": frame_w,
                            "frame_height_mm": frame_h,
                            "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                            "n_leaves": int(n_leaves),
                            "leaves": leaves,
                            "Nwin": int(count)
                        })
                    st.success(f"Блок '{name}' добавлен/обновлён.")

        # Глухие секции (панели)
        for i in range(st.session_state.get("tam_panel_count", 0)):
            with st.expander(f"Глухая секция #{i+1}", expanded=False):
                name = st.text_input(f"Название панели #{i+1}", value=f"Панель {i+1}", key=f"panel_name_{i}")
                count = st.number_input(f"Кол-во одинаковых панелей #{i+1}", min_value=1, value=1, key=f"panel_count_{i}")
                w = st.number_input(f"Ширина панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_w_{i}")
                h = st.number_input(f"Высота панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_h_{i}")
                left = st.number_input(f"LEFT панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_left_{i}")
                center = st.number_input(f"CENTER панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_center_{i}")
                right = st.number_input(f"RIGHT панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_right_{i}")
                top = st.number_input(f"TOP панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_top_{i}")
                fill = st.selectbox(f"Заполнение панели #{i+1}", options=filling_options_for_panels, index=0, key=f"panel_fill_{i}")

                if st.button(f"Добавить/обновить панель #{i+1} в секциях", key=f"save_panel_{i}"):
                    found = False
                    for idx, s in enumerate(st.session_state["sections_inputs"]):
                        if s.get("block_name") == name and s.get("kind") == "panel":
                            st.session_state["sections_inputs"][idx] = {
                                "kind": "panel",
                                "block_name": name,
                                "count": int(count),
                                "width_mm": w,
                                "height_mm": h,
                                "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                                "filling": fill,
                                "Nwin": int(count)
                            }
                            found = True
                            break
                    if not found:
                        st.session_state["sections_inputs"].append({
                            "kind": "panel",
                            "block_name": name,
                            "count": int(count),
                            "width_mm": w,
                            "height_mm": h,
                            "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                            "filling": fill,
                            "Nwin": int(count)
                        })
                    st.success(f"Панель '{name}' добавлена/обновлена.")

        st.markdown("**Текущие секции**")
        for idx, s in enumerate(st.session_state["sections_inputs"], start=1):
            st.write(f"{idx}. {s.get('kind')} — {s.get('block_name')} — N={s.get('Nwin',1)}")

        st.markdown("---")

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

        # Собираем base_positions
        base_positions = []
        for p in base_positions_inputs:
            if p["width_mm"] <= 0 or p["height_mm"] <= 0:
                st.error("Во всех позициях ширина и высота должны быть больше 0.")
                st.stop()
            area_m2 = (p["width_mm"] * p["height_mm"]) / 1_000_000.0
            perimeter_m = 2 * (p["width_mm"] + p["height_mm"]) / 1000.0
            base_positions.append({**p, "area_m2": area_m2, "perimeter_m": perimeter_m})

        lambr_positions = []
        for p in lambr_positions_inputs:
            if p["width_mm"] > 0 and p["height_mm"] > 0:
                area_m2 = (p["width_mm"] * p["height_mm"]) / 1_000_000.0
                perimeter_m = 2 * (p["width_mm"] + p["height_mm"]) / 1000.0
                lambr_positions.append({**p, "area_m2": area_m2, "perimeter_m": perimeter_m})

        # Собираем sections
        sections = []
        if product_type == "Тамбур":
            if not st.session_state["sections_inputs"]:
                st.error("Добавьте хотя бы одну секцию (дверной блок или панель).")
                st.stop()
            for s in st.session_state["sections_inputs"]:
                if s.get("kind") == "door":
                    fw = safe_float(s.get("frame_width_mm", 0.0))
                    fh = safe_float(s.get("frame_height_mm", 0.0))
                    area_m2 = (fw * fh) / 1_000_000.0
                    perimeter_m = 2 * (fw + fh) / 1000.0
                    nleaves = int(s.get("n_leaves", len(s.get("leaves", [])) or 1))
                    sections.append({**s, "area_m2": area_m2, "perimeter_m": perimeter_m, "nsash": nleaves})
                else:
                    w = safe_float(s.get("width_mm", 0.0))
                    h = safe_float(s.get("height_mm", 0.0))
                    area_m2 = (w * h) / 1_000_000.0
                    perimeter_m = 2 * (w + h) / 1000.0
                    sections.append({**s, "area_m2": area_m2, "perimeter_m": perimeter_m, "nsash": 0})
        else:
            for p in base_positions:
                sections.append({**p, "area_m2": p["area_m2"], "perimeter_m": p["perimeter_m"], "filling": p.get("filling", "Стеклопакет")})
            for p in lambr_positions:
                sections.append({**p, "area_m2": p["area_m2"], "perimeter_m": p["perimeter_m"], "filling": p.get("filling", "Ламбри без термо")})

        # --- Gabarit ---
        gab_calc = GabaritCalculator(excel)
        gabarit_rows, total_area_gab, total_perimeter_gab = gab_calc.calculate({"product_type": product_type}, sections)

        # --- Materials ---
        mat_calc = MaterialCalculator(excel)
        material_rows, material_total, total_area_mat = mat_calc.calculate({"product_type": product_type, "profile_system": profile_system}, sections, selected_duplicates)

        # --- Lambr/Sandwich calculation (по хлыстам 6 м)
        fill_linear = {}
        for s in sections:
            fill = str(s.get("filling") or "").replace("\xa0", " ").strip().lower()
            if fill in ("ламбри без термо", "ламбри с термо", "сэндвич"):
                fill_linear.setdefault(fill, 0.0)
                fill_linear[fill] += s.get("perimeter_m", 0.0) * s.get("Nwin", 1)

        fin_calc = FinalCalculator(excel)
        lambr_cost = 0.0
        for fill_name, linear_m in fill_linear.items():
            price_per_meter = fin_calc._find_price_for_filling(fill_name)
            if price_per_meter <= 0:
                st.warning(f"Не найдена цена за заполнение '{fill_name}' в СПРАВОЧНИК-2. Установлена 0.")
            count_hlyst = math.ceil(linear_m / 6.0) if linear_m > 0 else 0
            price_per_hlyst = price_per_meter * 6.0
            lambr_cost += count_hlyst * price_per_hlyst

        # --- Areas for glass etc.
        total_area_glass = 0.0
        for s in sections:
            if str(s.get("filling") or "").strip().lower() == "стеклопакет":
                total_area_glass += s.get("area_m2", 0.0) * s.get("Nwin", 1)
            # also check leaves
            for leaf in s.get("leaves", []):
                if str(leaf.get("filling","")).strip().lower() == "стеклопакет":
                    total_area_glass += (leaf.get("width_mm", 0.0) * leaf.get("height_mm", 0.0) / 1_000_000.0) * s.get("Nwin",1)

        total_area_all = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)

        # --- Handles / door blocks counts
        handles_count = 0
        door_blocks = 0
        for s in sections:
            if s.get("kind") == "door":
                nleaves = int(s.get("n_leaves", len(s.get("leaves", [])) or 1))
                handles_count += nleaves * s.get("Nwin", 1)
                # blocks: каждые 2 створки = 1 дверной блок
                door_blocks += int(math.ceil(nleaves / 2.0) * s.get("Nwin", 1))

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
            if gabarit_rows:
                gab_disp = [{"Тип элемента": t, "Фактическое значение": v} for t, v in gabarit_rows]
                st.dataframe(gab_disp, use_container_width=True)
            st.write(f"Общая площадь (служебная): **{total_area_gab:.3f} м²**")
            st.write(f"Суммарный периметр изделия: **{total_perimeter_gab:.3f} м**")
            if DEBUG:
                st.write("DEBUG sections:", sections)

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
            st.write(f"Панели (ламбри/сэндвич) — линейная длина по группам: **{', '.join([f'{k}: {v:.3f}м' for k,v in fill_linear.items()])}**, Итого: **{lambr_cost:.2f}**")

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
        if product_type == "Тамбур":
            for p in sections:
                if p.get("kind") == "door":
                    first_leaf = p.get("leaves", [{}])[0]
                    rows_for_form.append([
                        order_number,
                        pos_index,
                        product_type,
                        "",
                        "",
                        profile_system,
                        glass_type,
                        p.get("filling",""),
                        p.get("frame_width_mm", 0),
                        p.get("frame_height_mm", 0),
                        p.get("left_mm", 0.0),
                        p.get("center_mm", 0.0),
                        p.get("right_mm", 0.0),
                        p.get("top_mm", 0.0),
                        first_leaf.get("width_mm", p.get("frame_width_mm", 0)),
                        first_leaf.get("height_mm", p.get("frame_height_mm", 0)),
                        p.get("Nwin", 1),
                        toning,
                        assembly,
                        montage,
                        handle_type,
                        door_closer,
                    ])
                else:
                    rows_for_form.append([
                        order_number,
                        pos_index,
                        product_type,
                        "",
                        p.get("filling",""),
                        profile_system,
                        glass_type,
                        p.get("filling",""),
                        p.get("width_mm", 0),
                        p.get("height_mm", 0),
                        p.get("left_mm", 0.0),
                        p.get("center_mm", 0.0),
                        p.get("right_mm", 0.0),
                        p.get("top_mm", 0.0),
                        p.get("width_mm", 0),
                        p.get("height_mm", 0),
                        p.get("Nwin", 1),
                        toning,
                        assembly,
                        montage,
                        handle_type,
                        door_closer,
                    ])
                pos_index += 1
        else:
            for p in base_positions:
                rows_for_form.append([
                    order_number,
                    pos_index,
                    product_type,
                    "",
                    "",
                    profile_system,
                    glass_type,
                    p.get("filling",""),
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

        # --- экспорт коммерческого предложения ---
        smeta_bytes = build_smeta_workbook(
            order={
                "order_number": order_number,
                "product_type": product_type,
                "profile_system": profile_system,
                "filling_mode": "",  # убрали глобальный режим
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

    # ---------- Кнопка выхода ----------
    if st.sidebar.button("Выйти"):
        st.session_state.pop("current_user", None)
        try:
            if os.path.exists(SESSION_FILE):
                os.remove(SESSION_FILE)
        except:
            pass
        st.experimental_rerun()

if __name__ == "__main__":
    main()



