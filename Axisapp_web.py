# Axisapp_web.py
# Полностью самодостаточный файл — переработан калькулятор материалов.
# - Встроен безопасный safe_eval (AST)
# - Встроена обработка справочника (CSV/XLSX) и старые листы Excel
# - Генерация контекста заказа с дефолтами (ensure_defaults)
# - Встроенные fallback-функции по группам
# - Поддержка pack_size / norm_per_pack, фурнитуры и профильных групп
# - Возвращает итоговые таблицы: by_item, by_group, summary
# - Логирование нулевых строк
# Запускается как Streamlit-приложение (как и раньше).

import math
import os
import sys
import shutil
from io import BytesIO, StringIO
import zipfile
import logging
import json
import ast
import operator as op
import csv

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
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(ch)

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

    if isinstance(node, ast.Num):
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
        # allow math.sin etc
        if isinstance(func, ast.Attribute) and isinstance(func.value, ast.Name) and func.value.id == "math":
            fname = func.attr
            if hasattr(math, fname):
                args = [_eval_ast(a, names) for a in node.args]
                return getattr(math, fname)(*args)

        if isinstance(func, ast.Name) and func.id in ("max", "min", "round"):
            args = [_eval_ast(a, names) for a in node.args]
            return globals()[func.id](*args)

        raise ValueError("Разрешены только math.*, max, min, round")

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

    # build allowed names: copy context but ensure numbers
    names = {}
    for k, v in (context or {}).items():
        try:
            names[k] = float(v) if isinstance(v, (int, float, str)) and str(v) != "" else v
        except Exception:
            names[k] = v

    names.update({
        "math": math,
        "min": min,
        "max": max,
        "round": round,
    })

    try:
        node = ast.parse(formula, mode="eval")
        val = _eval_ast(node, names)
        # Some formulas may return booleans; cast to float safely
        try:
            return float(val)
        except Exception:
            return 0.0
    except Exception as e:
        logger.debug("safe_eval_formula error: %s; formula=%s; ctx=%s", e, formula, context)
        return 0.0

# =========================
# EXCEL CLIENT (с бэкапом)
# =========================

def is_probably_xlsx(path: str) -> bool:
    try:
        if not os.path.exists(path):
            return False
        if os.path.getsize(path) < 3000:
            return False
        import zipfile 
        with zipfile.ZipFile(path, "r") as z:
            return (
                "[Content_Types].xml" in z.namelist()
                and "xl/workbook.xml" in z.namelist()
            )
    except Exception:
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
# HELPERS: каталог CSV/XLSX -> records
# =========================

def process_catalog_file(path_or_bytes, sheet_name=None):
    """
    Поддерживает:
      - путь к .csv (str)
      - путь к .xlsx/.xls (str)
      - bytes/BytesIO с Excel (BytesIO)
    Возвращает список записей (list of dict), где ключи — нормализованные заголовки.
    """
    # Если передали BytesIO
    try:
        if isinstance(path_or_bytes, (bytes, bytearray)):
            bio = BytesIO(path_or_bytes)
            wb = load_workbook(bio, data_only=True)
            if sheet_name is None:
                sheet = wb[wb.sheetnames[0]]
            else:
                sheet = wb[sheet_name] if sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                return []
            header = [normalize_key(h) for h in rows[0]]
            recs = []
            for r in rows[1:]:
                if all(v is None for v in r):
                    continue
                row = {}
                for i, k in enumerate(header):
                    row[k] = r[i]
                recs.append(row)
            return recs
    except Exception:
        logger.debug("process_catalog_file: not bytes/xlsx or failed to parse as bytes", exc_info=True)

    # Если строка-путь
    if isinstance(path_or_bytes, str):
        path = path_or_bytes
        if not os.path.exists(path):
            logger.warning("Catalog file not found: %s", path)
            return []
        ext = os.path.splitext(path)[1].lower()
        if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
            try:
                wb = load_workbook(path, data_only=True)
                sheet = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
                rows = list(sheet.iter_rows(values_only=True))
                if not rows:
                    return []
                header = [normalize_key(h) for h in rows[0]]
                recs = []
                for r in rows[1:]:
                    if all(v is None for v in r):
                        continue
                    row = {}
                    for i, k in enumerate(header):
                        row[k] = r[i]
                    recs.append(row)
                return recs
            except Exception:
                logger.exception("Ошибка чтения Excel-файла каталога %s", path)
                return []
        elif ext == ".csv":
            try:
                recs = []
                with open(path, "r", encoding="utf-8-sig") as f:
                    reader = csv.reader(f)
                    rows = list(reader)
                if not rows:
                    return []
                header = [normalize_key(h) for h in rows[0]]
                for r in rows[1:]:
                    if all((c is None or str(c).strip() == "") for c in r):
                        continue
                    row = {}
                    for i, k in enumerate(header):
                        row[k] = r[i] if i < len(r) else None
                    recs.append(row)
                return recs
            except Exception:
                logger.exception("Ошибка чтения CSV-файла каталога %s", path)
                return []
    logger.warning("Unsupported catalog source: %s", type(path_or_bytes))
    return []

# =========================
# Контекст заказа с дефолтами
# =========================

def ensure_defaults(order: dict):
    """
    Расширяет order и секции дефолтными полями, чтобы формулы были устойчивы.
    """
    if order is None:
        order = {}
    # top-level defaults
    order.setdefault("order_number", "")
    order.setdefault("product_type", "")
    order.setdefault("profile_system", "")
    order.setdefault("glass_type", "")
    order.setdefault("toning", "Нет")
    order.setdefault("assembly", "Нет")
    order.setdefault("montage", "Нет")
    order.setdefault("handle_type", "")
    order.setdefault("door_closer", "Нет")
    # default numeric tuners
    for k in ["default_hinges_per_sash", "default_hinges_per_leaf"]:
        order.setdefault(k, 3)
    # ensure sections list exists
    order.setdefault("sections", [])
    for s in order["sections"]:
        s.setdefault("width_mm", safe_float(s.get("width_mm", 0.0)))
        s.setdefault("height_mm", safe_float(s.get("height_mm", 0.0)))
        s.setdefault("frame_width_mm", safe_float(s.get("frame_width_mm", s.get("width_mm", 0.0))))
        s.setdefault("frame_height_mm", safe_float(s.get("frame_height_mm", s.get("height_mm", 0.0))))
        s.setdefault("left_mm", safe_float(s.get("left_mm", 0.0)))
        s.setdefault("center_mm", safe_float(s.get("center_mm", 0.0)))
        s.setdefault("right_mm", safe_float(s.get("right_mm", 0.0)))
        s.setdefault("top_mm", safe_float(s.get("top_mm", 0.0)))
        s.setdefault("sash_width_mm", safe_float(s.get("sash_width_mm", s.get("width_mm", 0.0))))
        s.setdefault("sash_height_mm", safe_float(s.get("sash_height_mm", s.get("height_mm", 0.0))))
        s.setdefault("Nwin", int(s.get("Nwin", 1) or 1))
        s.setdefault("n_leaves", int(s.get("n_leaves", len(s.get("leaves", []) or []) or 1)))
        s.setdefault("leaves", s.get("leaves", []))
        # compute area/perimeter if missing
        if "area_m2" not in s or not s.get("area_m2"):
            w = s.get("frame_width_mm", s.get("width_mm", 0.0))
            h = s.get("frame_height_mm", s.get("height_mm", 0.0))
            s["area_m2"] = (safe_float(w) * safe_float(h)) / 1_000_000.0
        if "perimeter_m" not in s or not s.get("perimeter_m"):
            w = s.get("frame_width_mm", s.get("width_mm", 0.0))
            h = s.get("frame_height_mm", s.get("height_mm", 0.0))
            s["perimeter_m"] = 2 * (safe_float(w) + safe_float(h)) / 1000.0
    return order

# =========================
# ФАЛЬБЭК-ФУНКЦИИ для формул по group
# =========================

def fallback_profile_formula(ctx: dict):
    """
    Простейший fallback для профильных элементов:
    - Если есть perimeter и qty -> perimeter * qty
    - Если есть n_corners -> 4 * n_frame_rect
    """
    qty = safe_float(ctx.get("qty", 1))
    perimeter = safe_float(ctx.get("perimeter", 0.0)) or safe_float(ctx.get("perimeter_m", 0.0))
    if perimeter and qty:
        return perimeter * qty
    # если есть число прямоугольников и длина стороны (width/height) — попытаемся
    width = safe_float(ctx.get("width", 0.0))
    height = safe_float(ctx.get("height", 0.0))
    n_rect = int(ctx.get("n_rect", 0) or 0)
    if n_rect and width and height:
        per = 2 * (width + height) / 1000.0
        return per * n_rect * qty
    return 0.0

def fallback_fitting_formula(ctx: dict):
    """
    Фурнитура: стандартные правила
    - Ручки: 1 шт на дверной блок (qty)
    - Петли: hinges_per_sash * n_sash * qty
    - Доводчик: 1 шт на дверной блок (qty) если door
    """
    kind = str(ctx.get("type_elem", "") or "").lower()
    qty = int(ctx.get("qty", 1) or 1)
    n_sash = int(ctx.get("n_sash", 1) or 1)
    hinges_per_sash = int(ctx.get("hinges_per_sash", 3) or 3)
    if "ручк" in kind or "ручка" in kind:
        return qty
    if "петл" in kind or "петля" in kind or "hinge" in kind:
        return hinges_per_sash * n_sash * qty
    if "доводч" in kind or "доводчик" in kind:
        return qty
    # default small usage
    return max(1, qty)

FALLBACK_BY_GROUP = {
    "profile": fallback_profile_formula,
    "fitting": fallback_fitting_formula,
    # можно добавить индивидуальные группы
}

def fallback_formula_eval(formula: str, ctx: dict, group_name: str = ""):
    """
    Попытка вычислить формулу: сначала safe_eval, затем fallback по группе.
    """
    try:
        v = safe_eval_formula(formula, ctx)
        if v and abs(v) > 1e-9:
            return v
    except Exception:
        pass

    # Try group fallback
    if group_name:
        g = group_name.strip().lower()
        for key, fn in FALLBACK_BY_GROUP.items():
            if key in g:
                try:
                    fb = fn(ctx)
                    return float(fb or 0.0)
                except Exception:
                    logger.debug("fallback %s failed for group %s", key, g, exc_info=True)
    # generic fallback: perimeter * qty
    try:
        return float(fallback_profile_formula(ctx))
    except Exception:
        return 0.0

# =========================
# CALCULATORS (обновленные материалы -> by_item/by_group/summary)
# =========================

class GabaritCalculator:
    HEADER = ["Тип элемента", "Фактическое значение"]

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def _calc_imposts_context(self, width, height, left, center, right, top):
        n_sections_vert = 0
        if left > 0:
            n_sections_vert += 1
        if center > 0:
            n_sections_vert += 1
        if right > 0:
            n_sections_vert += 1

        n_imp_vert = max(0, n_sections_vert - 1)
        n_imp_hor = 1 if top > 0 else 0

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
            type_elem = get_field(row, "тип элемент", "") or get_field(row, "тип_элемент", "")
            formula = get_field(row, "формула_python", "") or get_field(row, "формула", "")
            if not type_elem or not formula:
                continue

            total_value = 0.0

            for s in sections:
                # determine dims
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
        n_sections_vert = 0
        if left > 0:
            n_sections_vert += 1
        if center > 0:
            n_sections_vert += 1
        if right > 0:
            n_sections_vert += 1

        n_imp_vert = max(0, n_sections_vert - 1)
        n_imp_hor = 1 if top > 0 else 0

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
        """
        Новая логика:
          - читает СПРАВОЧНИК-1 из Excel (SHEET_REF1)
          - для каждой записи вычисляет фактический расход (формула/формула_python)
          - поддерживает norm_per_pack (кол-во норм/упаковку), pack_size
          - fallback-вычисления по группам (group/type_element)
          - формирует итоговые таблицы: by_item (строки), by_group (агрегация по type_elem/group), summary
          - логирует записи с нулевым расходом (zero_rows)
        """
        ref_rows = self.excel.read_records(SHEET_REF1)
        total_area = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
        if not ref_rows:
            return [], 0.0, total_area

        items = []  # by_item rows as dicts
        zero_rows = []  # keep rows where qty_fact_total == 0
        total_sum = 0.0

        # Normalize selected_duplicates sets to simple lookup
        sel_dup = {k: set(v) if v else set() for k, v in (selected_duplicates or {}).items()}

        for row in ref_rows:
            # Extract fields with normalization
            row_type = str(get_field(row, "тип издел", "") or "").strip()
            row_profile = str(get_field(row, "система проф", "") or "").strip()
            type_elem = str(get_field(row, "тип элемент", "") or "").strip()
            product_name = str(get_field(row, "товар", "") or "").strip()
            group_name = str(get_field(row, "группа", "") or "").strip()
            arтикул = get_field(row, "артикул", "") or get_field(row, "артикул", "")
            formula = get_field(row, "формула_python", "") or get_field(row, "формула фактического расхода", "") or get_field(row, "формула", "")
            unit = str(get_field(row, "ед.", "") or "").strip()
            unit_fact = str(get_field(row, "ед. фактического расхода", "") or "").strip()
            unit_price = safe_float(get_field(row, "цена за", 0.0))
            norm_per_pack = safe_float(get_field(row, "кол-во норм", 0.0)) or safe_float(get_field(row, "norm_per_pack", 0.0))
            unit_pack = str(get_field(row, "ед .норма к упаковке", "") or "").strip() or str(get_field(row, "unit_pack", "") or "")
            pack_size = safe_float(get_field(row, "pack_size", 0.0)) or norm_per_pack

            # Filters by product type and profile_system
            if row_type and row_type.strip().lower() != order.get("product_type", "").strip().lower():
                continue
            if row_profile and row_profile.strip().lower() != order.get("profile_system", "").strip().lower():
                continue

            # Duplicates selection: if present, only include selected product names
            if type_elem in sel_dup and sel_dup[type_elem]:
                if product_name not in sel_dup[type_elem]:
                    continue

            if not type_elem or not formula:
                # If no formula but price exists -> could be pure service/one-time; skip for materials
                continue

            qty_fact_total = 0.0

            # Iterate through sections to compute consumption
            for s in sections:
                # Determine dims for section
                is_door_section = s.get("kind") == "door"
                if is_door_section:
                    width = s.get("frame_width_mm", s.get("width_mm", 0.0))
                    height = s.get("frame_height_mm", s.get("height_mm", 0.0))
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
                    "width": width, "height": height, "left": left, "center": center, "right": right, "top": top,
                    "sash_width": sash_w, "sash_height": sash_h, "sash_w": sash_w, "sash_h": sash_h,
                    "area": area, "perimeter": perimeter, "qty": qty,
                    "nsash": s.get("n_leaves", len(s.get("leaves", [])) or 1),
                    "n_sash": s.get("n_leaves", len(s.get("leaves", [])) or 1),
                    "n_sash_active": 1 if s.get("n_leaves", len(s.get("leaves", [])) or 1) >= 1 else 0,
                    "n_sash_passive": max(s.get("n_leaves", len(s.get("leaves", [])) or 1) - 1, 0),
                    "hinges_per_sash": int(s.get("hinges_per_sash", 3) or 3),
                    "type_elem": type_elem,
                    "group": group_name,
                }
                ctx.update(geom)

                # Evaluate formula with fallback
                try:
                    val = fallback_formula_eval(str(formula), ctx, group_name)
                    # Respect multiplicative factor: many formulas return per 1 item/1m; multiply by qty
                    qty_fact_total += safe_float(val) * safe_float(qty)
                except Exception:
                    logger.exception("Error evaluating material formula for %s (Formula: %s)", type_elem, formula)

            # Pack / norm handling
            if norm_per_pack and norm_per_pack > 0:
                qty_to_ship = math.ceil(qty_fact_total / norm_per_pack)
                effective_qty = qty_to_ship * norm_per_pack
            elif pack_size and pack_size > 0:
                qty_to_ship = math.ceil(qty_fact_total / pack_size)
                effective_qty = qty_to_ship * pack_size
            else:
                qty_to_ship = qty_fact_total
                effective_qty = qty_fact_total

            sum_row = effective_qty * unit_price
            total_sum += sum_row

            item = {
                "Тип изделия": row_type or "",
                "Система профиля": row_profile or "",
                "Тип элемента": type_elem,
                "Артикул": arтикул or "",
                "Товар": product_name or "",
                "Ед.": unit or "",
                "Цена за ед.": round(unit_price, 3),
                "Ед. факт. расхода": unit_fact or "",
                "Кол-во факт. расхода": round(qty_fact_total, 6),
                "Норма к упаковке": norm_per_pack,
                "Ед. к отгрузке": unit_pack or "",
                "Кол-во к отгрузке": round(effective_qty, 6),
                "Сумма": round(sum_row, 2),
                "group": group_name or "",
                "type_elem_raw": type_elem,
            }

            items.append(item)

            if abs(qty_fact_total) < 1e-9:
                # log zero rows
                zero_rows.append({
                    "type_elem": type_elem,
                    "product": product_name,
                    "formula": formula,
                    "row": item
                })
                logger.warning("Zero consumption for item: %s | product=%s | formula=%s", type_elem, product_name, formula)

        # Aggregation by group/type
        by_group = {}
        for it in items:
            g = (it.get("group") or it.get("Тип элемента") or "OTHER").strip()
            key = g
            agg = by_group.setdefault(key, {"Кол-во факт. расхода": 0.0, "Кол-во к отгрузке": 0.0, "Сумма": 0.0, "items": []})
            agg["Кол-во факт. расхода"] += safe_float(it.get("Кол-во факт. расхода", 0.0))
            agg["Кол-во к отгрузке"] += safe_float(it.get("Кол-во к отгрузки", 0.0))
            agg["Сумма"] += safe_float(it.get("Сумма", 0.0))
            agg["items"].append(it)

        by_group_list = []
        for k, v in sorted(by_group.items(), key=lambda kv: kv[0]):
            by_group_list.append({
                "Группа": k,
                "Кол-во факт. расхода": round(v["Кол-во факт. расхода"], 6),
                "Кол-во к отгрузке": round(v["Кол-во к отгрузки"], 6),
                "Сумма": round(v["Сумма"], 2),
                "Кол-элементов": len(v["items"])
            })

        # Summary
        summary = {
            "total_items": len(items),
            "total_groups": len(by_group_list),
            "total_sum": round(total_sum, 2),
            "total_area": round(total_area, 6),
            "zero_rows_count": len(zero_rows),
            "zero_rows": zero_rows[:50],  # sneak peek
        }

        # write to sheet for compatibility (old format)
        rows_for_sheet = []
        for it in items:
            rows_for_sheet.append([
                it.get("Тип изделия", ""),
                it.get("Система профиля", ""),
                it.get("Тип элемента", ""),
                it.get("Артикул", ""),
                it.get("Товар", ""),
                it.get("Ед.", ""),
                it.get("Цена за ед.", 0.0),
                it.get("Ед. факт. расхода", ""),
                it.get("Кол-во факт. расхода", 0.0),
                it.get("Норма к упаковке", 0.0),
                it.get("Ед. к отгрузке", ""),
                it.get("Кол-во к отгрузке", 0.0),
                it.get("Сумма", 0.0),
            ])

        # save to sheet (old behavior)
        try:
            self.excel.clear_and_write(SHEET_MATERIAL, self.HEADER, rows_for_sheet)
        except Exception:
            logger.exception("Failed to write material sheet")

        return items, by_group_list, summary

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
                  material_total: float,
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

        glass_sum = total_area_all * price_glass if total_area_all > 0 else 0.0
        rows.append(["Стеклопакет", price_glass, "за м²", glass_sum])

        toning_sum = total_area_all * price_toning if (toning.lower() != "нет" and total_area_all > 0) else 0.0
        rows.append(["Тонировка", price_toning, "за м²", toning_sum])

        assembly_sum = total_area_all * price_assembly if assembly.lower() != "нет" else 0.0
        rows.append(["Сборка", price_assembly, "за м²", assembly_sum])

        montage_sum = total_area_all * price_montage if montage.lower() != "нет" and total_area_all > 0 else 0.0
        rows.append(["Монтаж (" + str(montage) + ")", price_montage, "за м²", montage_sum])

        rows.append(["Материал", "-", "-", material_total])

        if lambr_cost > 0.0:
            rows.append(["Панели (Ламбри/Сэндвич)", "-", "-", lambr_cost])

        handles_sum = price_handles * handles_qty if handles_qty > 0 else 0.0
        rows.append(["Ручки", price_handles, "шт.", handles_sum])

        closer_sum = price_closer * closer_qty if closer_qty > 0 and door_closer.lower() != "нет" else 0.0
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

        try:
            self.excel.clear_and_write(SHEET_FINAL, self.HEADER, rows + extra_rows)
        except Exception:
            logger.exception("Failed to write final sheet")

        return rows, total_sum, ensure_sum

# =========================
# EXPORT: коммерческое предложение
# (не менял логику)
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
        except Exception:
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
            w = p.get('width_mm', p.get('frame_width_mm', 0))
            h = p.get('height_mm', p.get('frame_height_mm', 0))
            ws.cell(row=current_row, column=1, value=f"Панель {idx}: {w} × {h} мм, N = {p.get('Nwin',1)}, filling={p.get('filling','')}")
            current_row += 1

    current_row += 2
    ws.cell(row=current_row, column=1, value=f"Общая площадь: {total_area:.3f} м²"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Суммарный периметр: {total_perimeter:.3f} м"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"ИТОГО к оплате: {total_sum:.2f}")

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
# STREAMLIT UI: main
# (сохранена общая логика и эндпоинты)
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

    if "current_user" not in st.session_state:
        try:
            if os.path.exists(SESSION_FILE):
                with open(SESSION_FILE, "r", encoding="utf-8") as sf:
                    st.session_state["current_user"] = json.load(sf)
        except Exception:
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

    filling_options_for_panels = sorted(list(filling_types_set))
    if 'Стеклопакет' not in filling_options_for_panels:
         filling_options_for_panels.append('Стеклопакет')
    if 'Ламбри без термо' in filling_options_for_panels:
        default_panel_fill_index = filling_options_for_panels.index('Ламбри без термо')
    else:
        default_panel_fill_index = 0

    if not montage_types_set:
        montage_options = ["Есть", "Нет"]
    else:
        montage_options = sorted(list(montage_types_set))
        if "Нет" not in montage_options:
            montage_options.append("Нет")
    if "Нет" in montage_options:
        montage_options.insert(0, montage_options.pop(montage_options.index("Нет")))

    handle_types = sorted(list(handle_types_set)) if handle_types_set else [""]
    glass_types = sorted(list(glass_types_set)) if glass_types_set else ["двойной"]
    if not handle_types:
        handle_types = [""]
    if not glass_types:
        glass_types = ["двойной"]
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

        if st.button("✨ Новый расчёт / Очистить форму"):
            for k in list(st.session_state.keys()):
                if k.startswith(("w_","h_","l_","r_","c_","t_","sw_","sh_","nwin_","ls_w_","ls_h_","ls_q_","ls_fill_","door_","panel_","leaf_","tam_")):
                    st.session_state.pop(k, None)
            st.session_state["sections_inputs"] = []
            st.session_state["tam_door_count"] = 0
            st.session_state["tam_panel_count"] = 0
            st.experimental_rerun()

    col_left, col_right = st.columns([2, 1])

    with col_left:
        st.header("Позиции (окна/двери)")

        base_positions_inputs = []
        lambr_positions_inputs = []

        if product_type != "Тамбур":
            positions_count = st.number_input("Количество позиций (Окно/Дверь)", min_value=1, max_value=10, value=1, step=1)

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
                    "filling": glass_type,
                    "kind": "window" if product_type == "Окно" else "door"
                })
        else:
            # Tamбур dynamic block unchanged (kept logic)
            st.header("Параметры тамбура (дверные блоки и глухие панели)")

            c_add = st.columns([1,1,6])
            if c_add[0].button("Добавить дверной блок"):
                st.session_state["tam_door_count"] += 1
            if c_add[1].button("Добавить глухую секцию"):
                st.session_state["tam_panel_count"] += 1

            for i in range(st.session_state.get("tam_door_count", 0)):
                with st.expander(f"Дверной блок #{i+1}", expanded=False):
                    name = st.text_input(f"Название блока #{i+1}", value=f"Дверной блок {i+1}", key=f"door_name_{i}")
                    count = st.number_input(f"Кол-во одинаковых блоков #{i+1}", min_value=1, value=1, key=f"door_count_{i}")
                    dtype = st.selectbox(f"Тип двери #{i+1}", ["Одностворчатая","Двухстворчатая"], key=f"door_type_{i}")
                    frame_w = st.number_input(f"Ширина рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, key=f"frame_w_{i}")
                    frame_h = st.number_input(f"Высота рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, key=f"frame_h_{i}")

                    st.subheader("Внутренние импосты (для деления рамы)")
                    c_imp1, c_imp2 = st.columns(2)
                    left = c_imp1.number_input(f"LEFT, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"left_{i}", value=0.0)
                    center = c_imp2.number_input(f"CENTER, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"center_{i}", value=0.0)
                    c_imp3, c_imp4 = st.columns(2)
                    right = c_imp3.number_input(f"RIGHT, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"right_{i}", value=0.0)
                    top = c_imp4.number_input(f"TOP, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"top_{i}", value=0.0)

                    default_leaves = 1 if dtype == "Одностворчатая" else 2
                    n_leaves = st.number_input(f"Кол-во створок #{i+1}", min_value=1, value=default_leaves, key=f"n_leaves_{i}")

                    leaves = []
                    for L in range(int(n_leaves)):
                        st.markdown(f"**Створка {L+1}**")
                        lw = st.number_input(f"Ширина створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, key=f"leaf_w_{i}_{L}")
                        lh = st.number_input(f"Высота створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, key=f"leaf_h_{i}_{L}")
                        fill = st.selectbox(f"Заполнение створки {L+1} — блок {i+1}", options=filling_options_for_panels, index=filling_options_for_panels.index('Стеклопакет') if 'Стеклопакет' in filling_options_for_panels else 0, key=f"leaf_fill_{i}_{L}")
                        leaves.append({"width_mm": lw, "height_mm": lh, "filling": fill})

                    if st.button(f"Добавить/обновить дверной блок #{i+1} в секциях", key=f"save_door_{i}"):
                        new_section = {
                            "kind": "door",
                            "block_name": name,
                            "frame_width_mm": frame_w,
                            "frame_height_mm": frame_h,
                            "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                            "n_leaves": int(n_leaves),
                            "leaves": leaves,
                            "Nwin": int(count),
                            "filling": glass_type
                        }
                        st.session_state["sections_inputs"] = [s for s in st.session_state["sections_inputs"] if not (s.get("block_name") == name and s.get("kind") == "door")]
                        st.session_state["sections_inputs"].append(new_section)
                        st.success(f"Дверной блок '{name}' добавлен/обновлён.")

            for i in range(st.session_state.get("tam_panel_count", 0)):
                with st.expander(f"Глухая секция #{i+1}", expanded=False):
                    name = st.text_input(f"Название панели #{i+1}", value=f"Панель {i+1}", key=f"panel_name_{i}")
                    count = st.number_input(f"Кол-во одинаковых панелей #{i+1}", min_value=1, value=1, key=f"panel_count_{i}")
                    p1, p2 = st.columns(2)
                    w = p1.number_input(f"Ширина панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_w_{i}")
                    h = p2.number_input(f"Высота панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_h_{i}")
                    fill = st.selectbox(f"Заполнение панели #{i+1}", options=filling_options_for_panels, index=default_panel_fill_index, key=f"panel_fill_{i}")

                    st.subheader("Внутренние импосты (для деления рамы)")
                    c_imp5, c_imp6 = st.columns(2)
                    left = c_imp5.number_input(f"LEFT, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_left_{i}", value=0.0)
                    center = c_imp6.number_input(f"CENTER, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_center_{i}", value=0.0)
                    c_imp7, c_imp8 = st.columns(2)
                    right = c_imp7.number_input(f"RIGHT, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_right_{i}", value=0.0)
                    top = c_imp8.number_input(f"TOP, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_top_{i}", value=0.0)

                    if st.button(f"Добавить/обновить панель #{i+1} в секциях", key=f"save_panel_{i}"):
                        new_section = {
                            "kind": "panel",
                            "block_name": name,
                            "width_mm": w,
                            "height_mm": h,
                            "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                            "filling": fill,
                            "Nwin": int(count)
                        }
                        st.session_state["sections_inputs"] = [s for s in st.session_state["sections_inputs"] if not (s.get("block_name") == name and s.get("kind") == "panel")]
                        st.session_state["sections_inputs"].append(new_section)
                        st.success(f"Панель '{name}' добавлена/обновлена.")

            st.markdown("**Текущие секции Тамбура:**")
            if st.session_state["sections_inputs"]:
                 for idx, s in enumerate(st.session_state["sections_inputs"], start=1):
                    main_dim = f"{s.get('width_mm', s.get('frame_width_mm'))}x{s.get('height_mm', s.get('frame_height_mm'))}"
                    imposts = f" L{s.get('left_mm',0)} C{s.get('center_mm',0)} R{s.get('right_mm',0)} T{s.get('top_mm',0)}"
                    st.write(f"**{idx}. {s.get('kind').capitalize()}** ({s.get('block_name')}) — {main_dim}, N={s.get('Nwin',1)} | Импосты:{imposts}")
            else:
                 st.info("Нет добавленных секций.")

        st.markdown("---")

    with col_right:
        st.header("Информация")
        st.info("Тамбур детализируется отдельными секциями: дверные блоки и глухие панели.")
        if not is_probably_xlsx(EXCEL_FILE):
            st.warning("Excel-файл справочников может быть не в порядке — проверь СПРАВОЧНИК-2/1/3.")

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
    calc_button = st.button("💾 Сохранить в Excel и выполнить расчёт")

    if calc_button:
        if not order_number.strip():
            st.error("Введите номер заказа.")
            st.stop()

        # Build sections
        sections = []

        if product_type != "Тамбур":
             for p in base_positions_inputs:
                if p["width_mm"] <= 0 or p["height_mm"] <= 0:
                    st.error("Во всех позициях ширина и высота должны быть больше 0.")
                    st.stop()
                area_m2 = (p["width_mm"] * p["height_mm"]) / 1_000_000.0
                perimeter_m = 2 * (p["width_mm"] + p["height_mm"]) / 1000.0
                sections.append({**p, "area_m2": area_m2, "perimeter_m": perimeter_m})

        else:
             sections = st.session_state["sections_inputs"]
             for s in sections:
                if s.get("kind") == "door":
                    fw = s.get("frame_width_mm", 0.0)
                    fh = s.get("frame_height_mm", 0.0)
                    area_m2 = (fw * fh) / 1_000_000.0
                    perimeter_m = 2 * (fw + fh) / 1000.0
                    s.update({"area_m2": area_m2, "perimeter_m": perimeter_m})
                elif s.get("kind") == "panel":
                    w = s.get("width_mm", 0.0)
                    h = s.get("height_mm", 0.0)
                    area_m2 = (w * h) / 1_000_000.0
                    perimeter_m = 2 * (w + h) / 1000.0
                    s.update({"area_m2": area_m2, "perimeter_m": perimeter_m})

        if not sections:
            st.error("Необходимо задать хотя бы одну позицию с габаритами > 0.")
            st.stop()

        # Prepare order dict and ensure defaults
        order = {
            "order_number": order_number,
            "product_type": product_type,
            "profile_system": profile_system,
            "glass_type": glass_type,
            "toning": toning,
            "assembly": assembly,
            "montage": montage,
            "handle_type": handle_type,
            "door_closer": door_closer,
            "sections": sections
        }
        order = ensure_defaults(order)

        # Gabarit Calculation
        gab_calc = GabaritCalculator(excel)
        gabarit_rows, total_area_gab, total_perimeter_gab = gab_calc.calculate(order, sections)

        # Material Calculation -> returns items, by_group, summary
        mat_calc = MaterialCalculator(excel)
        items, by_group, summary = mat_calc.calculate(order, sections, selected_duplicates)

        # material_total from summary
        material_total = safe_float(summary.get("total_sum", 0.0))

        # compute lambr cost as before
        total_area_all = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
        lambr_cost = 0.0
        fin_calc = FinalCalculator(excel)

        for s in sections:
            fill_name = str(s.get("filling") or "").strip().lower()
            if fill_name in ["ламбри без термо", "ламбри с термо", "сэндвич"]:
                price_per_meter = fin_calc._find_price_for_filling(fill_name)
                if s.get("kind") == "door":
                    for leaf in s.get("leaves", []):
                        leaf_fill = str(leaf.get("filling") or "").strip().lower()
                        if leaf_fill in ["ламбри без термо", "ламбри с термо", "сэндвич"]:
                            leaf_w = leaf.get("width_mm", 0.0)
                            leaf_h = leaf.get("height_mm", 0.0)
                            perimeter_leaf = 2 * (leaf_w + leaf_h) / 1000.0
                            count_hlyst = math.ceil(perimeter_leaf / 6.0) if perimeter_leaf > 0 else 0
                            price_per_hlyst = price_per_meter * 6.0
                            lambr_cost += count_hlyst * price_per_hlyst * s.get("Nwin", 1)
                elif s.get("kind") in ["panel", "window"]:
                    perimeter_s = s.get("perimeter_m", 0.0) * s.get("Nwin", 1)
                    count_hlyst = math.ceil(perimeter_s / 6.0) if perimeter_s > 0 else 0
                    price_per_hlyst = price_per_meter * 6.0
                    lambr_cost += count_hlyst * price_per_hlyst

        # Handles / closers counts
        handles_count = 0
        closer_count = 0
        if product_type in ("Дверь", "Тамбур"):
            for s in sections:
                if s.get("kind") == "door" or (product_type == "Дверь" and s.get("kind") == "door"):
                     handles_count += s.get("Nwin", 1)
                     if door_closer.lower() == "есть":
                         closer_count += s.get("Nwin", 1)

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
            material_total=material_total,
            lambr_cost=lambr_cost,
            handles_qty=handles_count,
            closer_qty=closer_count
        )

        st.success(f"Расчёт выполнен. Итоговая сумма: {total_sum:.2f}")

        # --- Вывод результатов и экспорт ---
        tab1, tab2, tab3, tab4 = st.tabs(["Габариты", "Материалы (по позициям)", "Материалы (по группам)", "Итоговый расчет"])

        with tab1:
            st.subheader("Расчет по габаритам")
            if gabarit_rows:
                gab_disp = [{"Тип элемента": t, "Фактическое значение": v} for t, v in gabarit_rows]
                st.dataframe(gab_disp, use_container_width=True)
            st.write(f"Общая площадь: **{total_area_gab:.3f} м²**")
            st.write(f"Суммарный периметр: **{total_perimeter_gab:.3f} м**")

        with tab2:
            st.subheader("Расчёт материалов — by_item")
            if items:
                # show list of dicts
                st.dataframe(items, use_container_width=True)
            st.write(f"Итого по материалам: **{material_total:.2f}**")
            if summary.get("zero_rows_count", 0) > 0:
                st.warning(f"Найдено {summary['zero_rows_count']} строк(а) со значением расхода 0 — проверь справочник/формулы.")
                if st.checkbox("Показать примеры нулевых строк"):
                    st.json(summary.get("zero_rows", []))

        with tab3:
            st.subheader("Расчёт материалов — by_group")
            if by_group:
                st.dataframe(by_group, use_container_width=True)

        with tab4:
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

        for p in sections:
            rows_for_form.append([
                order_number, pos_index, product_type,
                p.get("kind", ""),
                p.get("n_leaves", 1) if p.get("kind") == "door" else 0,
                profile_system, glass_type, p.get("filling",""),
                p.get("width_mm", 0.0) if not p.get("frame_width_mm") else p.get("frame_width_mm", 0.0),
                p.get("height_mm", 0.0) if not p.get("frame_height_mm") else p.get("frame_height_mm", 0.0),
                p.get("left_mm", 0.0), p.get("center_mm", 0.0), p.get("right_mm", 0.0), p.get("top_mm", 0.0),
                p.get("sash_width_mm", p.get("width_mm", 0.0)),
                p.get("sash_height_mm", p.get("height_mm", 0.0)),
                p.get("Nwin", 1),
                toning, assembly, montage, handle_type, door_closer,
            ])
            pos_index += 1

        for row in rows_for_form:
             try:
                 excel.append_form_row(row)
             except Exception:
                 logger.exception("Failed to append form row")

        # --- Экспорт коммерческого предложения ---
        base_pos = [s for s in sections if s.get("kind") in ["window", "door"] and product_type != "Тамбур"]
        tam_pos = [s for s in sections if s.get("kind") in ["door"] and product_type == "Тамбур"]
        lambr_pos = [s for s in sections if s.get("kind") == "panel" or (product_type == "Тамбур" and s.get("kind") != "door")]

        smeta_bytes = build_smeta_workbook(
            order={
                "order_number": order_number, "product_type": product_type, "profile_system": profile_system,
                "filling_mode": "", "glass_type": glass_type, "toning": toning, "assembly": assembly,
                "montage": montage, "handle_type": handle_type, "door_closer": door_closer,
            },
            base_positions=base_pos + tam_pos,
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
        try:
            if os.path.exists(SESSION_FILE):
                os.remove(SESSION_FILE)
        except Exception:
            pass
        st.experimental_rerun()

if __name__ == "__main__":
    main()
