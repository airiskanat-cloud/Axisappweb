import math
import os
import sys
import shutil
from io import BytesIO
import logging
import json
import ast
import operator as op

import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.drawing.image import Image as XLImage

# =========================
# КОНСТАНТЫ / НАСТРОЙКИ
# =========================

DEBUG = False
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

def resource_path(relative_path: str) -> str:
    try:
        if hasattr(sys, "_MEIPASS"):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(os.path.dirname(__file__))
    except Exception:
        base_path = os.getcwd()
    return os.path.join(base_path, relative_path)

DATA_DIR = os.getenv("AXIS_DATA_DIR", os.path.join(os.path.expanduser("~"), ".axis_app_data"))
os.makedirs(DATA_DIR, exist_ok=True)

TEMPLATE_EXCEL_NAME = "axis_pro_gf.xlsx"
EXCEL_FILE = os.path.join(DATA_DIR, TEMPLATE_EXCEL_NAME)
SESSION_FILE = os.path.join(DATA_DIR, "session_user.json")

BUNDLED_TEMPLATE = resource_path(TEMPLATE_EXCEL_NAME)
if os.path.exists(BUNDLED_TEMPLATE) and not os.path.exists(EXCEL_FILE):
    try:
        shutil.copyfile(BUNDLED_TEMPLATE, EXCEL_FILE)
        logger.info("Copied bundled template %s -> %s", BUNDLED_TEMPLATE, EXCEL_FILE)
    except Exception:
        logger.exception("Error copying bundled template")

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

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
    except Exception:
        return default

def safe_int(value, default=0):
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

    if isinstance(node, ast.Num):  # compatibility
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
    except Exception:
        return 0.0

# =========================
# EXCEL CLIENT (с бэкапом)
# =========================

class ExcelClient:
    def __init__(self, filename: str):
        self.filename = filename
        if not os.path.exists(self.filename):
            self._create_template()
        self.load()

    def _create_template(self):
        wb = Workbook()
        # ensure named sheets exist
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
        except Exception as e:
            logger.exception("Ошибка при загрузке Excel, делаю бэкап и пересоздаю шаблон: %s", e)
            try:
                if os.path.exists(self.filename):
                    shutil.copyfile(self.filename, self.filename + ".corrupt.bak")
            except Exception:
                pass
            try:
                if os.path.exists(self.filename):
                    os.remove(self.filename)
            except Exception:
                pass
            self._create_template()
            self.wb = load_workbook(self.filename, data_only=True)

    def save(self):
        try:
            self.wb.save(self.filename)
        except Exception as e:
            logger.exception("Ошибка сохранения: %s", e)

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
        except Exception:
            pass

        if header:
            ws.append(header)
        for row in rows:
            ws.append(row)
        self.save()

    def append_form_row(self, row: list):
        ws = self.ws(SHEET_FORM)
        try:
            # if header empty -> write header
            if ws.max_row == 1 and not any(ws[1]):
                ws.append(FORM_HEADER)
        except Exception:
            pass
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
    if "current_user" in st.session_state:
        return st.session_state["current_user"]

    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE, "r", encoding="utf-8") as sf:
                st.session_state["current_user"] = json.load(sf)
                return st.session_state["current_user"]
        except Exception:
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
                try:
                    with open(SESSION_FILE, "w", encoding="utf-8") as sf:
                        json.dump(st.session_state["current_user"], sf, ensure_ascii=False)
                except Exception:
                    pass

                st.sidebar.success(f"Привет, {user['_raw_login']}!")
                return st.session_state["current_user"]

        st.sidebar.error("Неверный логин или пароль")

    return None

# =========================
# CALCULATORS
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
            return [], total_area, total_perimeter

        gabarit_values = []

        for row in ref_rows:
            type_elem = get_field(row, "тип элемент", "")
            formula = get_field(row, "формула_python", "")
            if not type_elem or not formula:
                continue

            total_value = 0.0

            for s in sections:
                # размеры: отдельная логика для дверей
                if s.get("kind") == "door":
                    width = s.get("frame_width_mm", 0.0) or s.get("width_mm", 0.0)
                    height = s.get("frame_height_mm", 0.0) or s.get("height_mm", 0.0)
                    if s.get("leaves"):
                        first_leaf = s.get("leaves", [{}])[0]
                        sash_w = first_leaf.get("width_mm", width)
                        sash_h = first_leaf.get("height_mm", height)
                    else:
                        sash_w = width
                        sash_h = height
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

                nsash = s.get("n_leaves", len(s.get("leaves", [])) or 1)

                ctx = {
                    "width": width,
                    "height": height,
                    "left": left,
                    "center": center,
                    "right": right,
                    "top": top,
                    "area": area,
                    "perimeter": perimeter,
                    "qty": qty,
                    "sash_width": sash_w,
                    "sash_height": sash_h,
                    "sash_w": sash_w,
                    "sash_h": sash_h,
                    "n_sash": nsash,
                    "n_sash_active": 1 if nsash >= 1 else 0,
                    "n_sash_passive": max(nsash - 1, 0),
                    "hinges_per_sash": 3,
                }

                try:
                    geom = self._calc_imposts_context(width, height, left, center, right, top)
                    if isinstance(geom, dict):
                        ctx.update(geom)
                except Exception:
                    pass

                try:
                    total_value += safe_eval_formula(str(formula), ctx)
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
                    width = s.get("frame_width_mm", 0.0) or s.get("width_mm", 0.0)
                    height = s.get("frame_height_mm", 0.0) or s.get("height_mm", 0.0)
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

                try:
                    qty_fact_total += safe_eval_formula(str(formula), ctx)
                except Exception:
                    logger.exception("Error evaluating material formula for %s", type_elem)

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

        # ИСПРАВЛЕНИЕ ОШИБКИ SyntaxError: '(' was never closed (строка 844)
        toning_sum = total_area_glass * price_toning if (toning.lower() != "нет" and total_area_glass > 0) else 0.0
        rows.append(["Тонировка", price_toning, "за м²", toning_sum])

        assembly_sum = total_area_all * price_assembly if assembly.lower() != "нет" else 0.0
        rows.append(["Сборка", price_assembly, "за м²", assembly_sum])

        montage_sum = total_area_all * price_montage if montage.lower() != "нет" and total_area_all > 0 else 0.0
        rows.append(["Монтаж", price_montage, "за м²", montage_sum])

        # Расчет ручек и доводчиков
        handles_sum = handles_qty * price_handles
        rows.append([f"Ручки ({handle_type})", price_handles, "за шт", handles_sum])

        closer_sum = closer_qty * price_closer if door_closer.lower() != "нет" else 0.0
        rows.append([f"Доводчик", price_closer, "за шт", closer_sum])

        # Добавление стоимости материалов
        rows.append(["Материалы (профиль, фурнитура и т.д.)", "-", "-", material_total])

        # Добавление Ламбрина (если есть)
        if lambr_cost > 0.0:
            rows.append(["Ламинация", "-", "-", lambr_cost])

        # Итоговая сумма
        total_sum_final = glass_sum + toning_sum + assembly_sum + montage_sum + handles_sum + closer_sum + material_total + lambr_cost
        rows.append(["ИТОГО (без скидки)", "-", "-", total_sum_final])

        # Запись в Excel
        self.excel.clear_and_write(SHEET_FINAL, self.HEADER, rows)
        return rows, total_sum_final
