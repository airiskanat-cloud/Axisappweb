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

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

# =========================
# КОНСТАНТЫ / НАСТРОЙКИ
# =========================

DEBUG = False
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# Списки для интерфейса (согласно запросу)
PRODUCT_TYPES = ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"]
PROFILE_SYSTEMS = [
    "ALG 2030-63C", "ALG 2030-55C", "ALG 2030-73C", 
    "ALG 2030-45C", "ALG 2030-Slim", "Ruit 50F"
]

# Заголовки
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

# =========================
# УТИЛИТЫ
# =========================

def normalize_key(k):
    if k is None: return None
    s = str(k).replace("\xa0", " ")
    s = " ".join(s.split())
    return s.strip().lower()

def _clean_cell_val(v):
    if v is None: return ""
    return str(v).replace("\xa0", " ").strip()

def safe_float(value, default=0.0):
    try:
        if value is None: return default
        s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
        return float(s) if s else default
    except: return default

def get_field(row: dict, needle: str, default=None):
    if not isinstance(row, dict): return default
    needle = (needle or "").lower().strip()
    for k, v in row.items():
        if k and needle in str(k).lower(): return v
    return default

# =========================
# БЕЗОПАСНЫЙ EVAL
# =========================

_allowed_ops = {
    ast.Add: op.add, ast.Sub: op.sub, ast.Mult: op.mul, ast.Div: op.truediv,
    ast.Pow: op.pow, ast.USub: op.neg, ast.UAdd: op.pos, ast.Mod: op.mod,
    ast.FloorDiv: op.floordiv, ast.Lt: op.lt, ast.Gt: op.gt, ast.LtE: op.le,
    ast.GtE: op.ge, ast.Eq: op.eq, ast.NotEq: op.ne,
    ast.And: lambda a,b: a and b, ast.Or: lambda a,b: a or b,
}

def _eval_ast(node, names):
    if isinstance(node, ast.Expression): return _eval_ast(node.body, names)
    if isinstance(node, (ast.Constant, ast.Num)): return node.value if isinstance(node, ast.Constant) else node.n
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
        if node.id in names: return names[node.id]
        raise ValueError(f"Недопустимое имя '{node.id}'")
    if isinstance(node, ast.Call):
        func = node.func
        if isinstance(func, ast.Attribute) and isinstance(func.value, ast.Name) and func.value.id == "math":
            fname = func.attr
            if hasattr(math, fname):
                args = [_eval_ast(a, names) for a in node.args]
                return getattr(math, fname)(*args)
        if isinstance(func, ast.Name) and func.id in ("min", "max"):
            args = [_eval_ast(a, names) for a in node.args]
            return globals()[func.id](*args)
    raise ValueError(f"Недопустимый элемент формулы")

def safe_eval_formula(formula: str, context: dict) -> float:
    formula = (formula or "").strip()
    if not formula: return 0.0
    try:
        names = {**context, "math": math, "min": min, "max": max}
        node = ast.parse(formula, mode="eval")
        return float(_eval_ast(node, names))
    except: return 0.0

# =========================
# GOOGLE SHEETS CLIENT
# =========================

class GoogleSheetsClient:
    def __init__(self, sheet_id: str):
        self.sheet_id = sheet_id
        self._worksheets_cache = {}
        self.load()

    @st.cache_resource
    def _auth_v3(_self):
        import base64
        key_b64 = os.environ.get("GCP_SA_KEYFILE_JSON_BASE64")
        if not key_b64:
            st.error("❌ Ключ GCP_SA_KEYFILE_JSON_BASE64 не найден.")
            st.stop()
        info = json.loads(base64.b64decode(key_b64).decode("utf-8"))
        creds = Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(creds)

    def load(self):
        try:
            client = self._auth_v3()
            self.wb = client.open_by_key(self.sheet_id)
        except Exception as e:
            st.error(f"Ошибка подключения: {e}")
            st.stop()

    def ws(self, name: str):
        if name in self._worksheets_cache: return self._worksheets_cache[name]
        try:
            ws = self.wb.worksheet(name)
            self._worksheets_cache[name] = ws
            return ws
        except:
            if name == SHEET_FORM:
                ws = self.wb.add_worksheet(name, rows="100", cols="30")
                ws.append_row(FORM_HEADER)
                return ws
            st.error(f"Лист {name} не найден.")
            st.stop()

    @st.cache_data(ttl=600)
    def read_records(_self, sheet_name: str):
        rows = _self.ws(sheet_name).get_all_values()
        if not rows: return []
        header = [normalize_key(h) for h in rows[0]]
        records = []
        for r in rows[1:]:
            if any(r): records.append({header[i]: r[i] for i in range(len(header)) if i < len(r)})
        return records

    def append_form_row(self, row: list):
        try: self.ws(SHEET_FORM).append_row(row, value_input_option='USER_ENTERED')
        except Exception as e: st.error(f"Ошибка записи: {e}")

# =========================
# КАЛЬКУЛЯТОРЫ
# =========================

class GabaritCalculator:
    def __init__(self, excel_client: GoogleSheetsClient):
        self.excel = excel_client

    def calculate(self, order: dict, sections: list):
        ref_rows = self.excel.read_records(SHEET_REF3)
        total_area = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
        total_perimeter = sum(s.get("perimeter_m", 0.0) * s.get("Nwin", 1) for s in sections)
        
        g_vals = []
        for row in ref_rows:
            type_elem = get_field(row, "тип элемент", "")
            formula = get_field(row, "формула_python", "")
            if not type_elem or not formula: continue
            
            val_sum = 0.0
            for s in sections:
                ctx = {
                    "width": s.get("width_mm", 0.0), "height": s.get("height_mm", 0.0),
                    "area": s.get("area_m2", 0.0), "qty": s.get("Nwin", 1),
                    "is_facade": 1 if order.get("product_type") == "Фасад" else 0
                }
                val_sum += safe_eval_formula(formula, ctx) * ctx["qty"]
            g_vals.append([type_elem, val_sum])
        return g_vals, total_area, total_perimeter

# =========================
# ОСНОВНОЕ ПРИЛОЖЕНИЕ
# =========================

def main():
    st.set_page_config(page_title="AXIS Калькулятор 15.1", layout="wide")
    client = GoogleSheetsClient(GSPREAD_SHEET_ID)

    st.title("🧮 AXIS: Расчет конструкций (Фасады и Окна)")

    with st.expander("📝 Основные параметры заказа", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            order_num = st.text_input("Номер заказа", "001")
            product_type = st.selectbox("Тип изделия", PRODUCT_TYPES)
        with col2:
            profile_sys = st.selectbox("Серия профиля", PROFILE_SYSTEMS)
            glass_type = st.selectbox("Тип заполнения", ["4мм", "6мм", "Стеклопакет 24мм", "Стеклопакет 32мм"])
        with col3:
            montage = st.radio("Монтаж", ["Да", "Нет"], horizontal=True)
            assembly = st.radio("Сборка", ["Да", "Нет"], horizontal=True)

    # Логика для Фасада (Каркас + Заполнение)
    sections = []
    if product_type == "Фасад":
        st.subheader("🏗️ Параметры Фасада")
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            f_width = st.number_input("Общая ширина фасада (мм)", 0)
            f_height = st.number_input("Общая высота фасада (мм)", 0)
        with f_col2:
            st.info("Добавьте окна и двери, которые встроены в фасад")
            
        # Пример добавления секции
        sections.append({
            "kind": "facade_frame", "width_mm": f_width, "height_mm": f_height, 
            "area_m2": (f_width * f_height) / 1_000_000, "Nwin": 1
        })

    else:
        # Обычный ввод для окон/дверей
        st.subheader("🖼️ Параметры конструкции")
        w = st.number_input("Ширина (мм)", 0)
        h = st.number_input("Высота (мм)", 0)
        qty = st.number_input("Кол-во (шт)", 1)
        sections.append({
            "kind": "standard", "width_mm": w, "height_mm": h, 
            "area_m2": (w * h) / 1_000_000, "Nwin": qty
        })

    if st.button("🚀 Рассчитать и Сохранить"):
        calc = GabaritCalculator(client)
        results, t_area, t_perim = calc.calculate({"product_type": product_type}, sections)
        
        st.success(f"Расчет завершен! Общая площадь: {t_area:.2f} м²")
        st.table(pd.DataFrame(results, columns=["Элемент", "Значение"]))
        
        # Сохранение в Google Sheets
        for s in sections:
            client.append_form_row([
                order_num, "1", product_type, "", "", profile_sys, 
                glass_type, "", s["width_mm"], s["height_mm"], 
                0, 0, 0, 0, 0, 0, s["Nwin"], "Нет", assembly, montage, "", ""
            ])
        st.info("Данные успешно отправлены в таблицу.")

if __name__ == "__main__":
    main()
