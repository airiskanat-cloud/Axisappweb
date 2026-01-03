# =========================================
# Axis Pro GF v18 — Facade Calculator
# =========================================

import math
import ast
import operator as op
import logging
import sys
import os

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# =========================================
# CONFIG
# =========================================

APP_TITLE = "Axis Pro GF — Фасад / Окна / Двери"

GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

ENSURE_COEF = 0.65  # 65%

# =========================================
# LOGGER
# =========================================

logger = logging.getLogger("axis")
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

# =========================================
# UTILS
# =========================================

def normalize_key(v):
    if v is None:
        return ""
    return " ".join(str(v).replace("\xa0", " ").lower().split())

def safe_float(v, default=0.0):
    try:
        if v is None:
            return default
        s = str(v).replace("\xa0", "").replace(" ", "").replace(",", ".")
        if s == "":
            return default
        return float(s)
    except Exception:
        return default

def get_field(row: dict, needle: str, default=None):
    needle = needle.lower()
    for k, v in row.items():
        if k and needle in str(k).lower():
            return v
    return default

# =========================================
# SAFE AST EVAL
# =========================================

_ALLOWED_OPS = {
    ast.Add: op.add,
    ast.Sub: op.sub,
    ast.Mult: op.mul,
    ast.Div: op.truediv,
    ast.Pow: op.pow,
    ast.USub: op.neg,
    ast.UAdd: op.pos,
    ast.Mod: op.mod,
}

def _eval_node(node, names):
    if isinstance(node, ast.Expression):
        return _eval_node(node.body, names)
    if isinstance(node, ast.Constant):
        return node.value
    if isinstance(node, ast.Name):
        if node.id in names:
            return names[node.id]
        raise ValueError(node.id)
    if isinstance(node, ast.BinOp):
        return _ALLOWED_OPS[type(node.op)](
            _eval_node(node.left, names),
            _eval_node(node.right, names),
        )
    if isinstance(node, ast.UnaryOp):
        return _ALLOWED_OPS[type(node.op)](_eval_node(node.operand, names))
    if isinstance(node, ast.Call):
        if isinstance(node.func, ast.Attribute) and node.func.value.id == "math":
            fn = getattr(math, node.func.attr)
            return fn(*[_eval_node(a, names) for a in node.args])
        if isinstance(node.func, ast.Name) and node.func.id in ("min", "max"):
            return globals()[node.func.id](*[_eval_node(a, names) for a in node.args])
    raise ValueError("Unsafe expression")

def safe_eval(formula: str, context: dict) -> float:
    if not formula:
        return 0.0
    try:
        ctx = {k: safe_float(v) for k, v in context.items()}
        ctx["math"] = math
        node = ast.parse(formula, mode="eval")
        return float(_eval_node(node, ctx))
    except Exception as e:
        logger.error("Formula error %s | %s", formula, e)
        return 0.0

# =========================================
# GOOGLE SHEETS
# =========================================

class GoogleSheets:
    @st.cache_resource
    def auth(_self):
        secret_path = "/etc/secrets/gcp_service_account.json"
        creds = Credentials.from_service_account_file(
            secret_path,
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        return gspread.authorize(creds)

    def __init__(self, sheet_id):
        self.client = self.auth()
        self.book = self.client.open_by_key(sheet_id)
        self.cache = {}

    def ws(self, name):
        if name not in self.cache:
            self.cache[name] = self.book.worksheet(name)
        return self.cache[name]

    @st.cache_data(ttl=1800)
    def read(_self, sheet_name):
        return _self.ws(sheet_name).get_all_records()

# =========================================
# LOGIN
# =========================================

def login(gs):
    if "user" in st.session_state:
        return True

    st.sidebar.title("🔐 Вход")
    u = st.sidebar.text_input("Логин")
    p = st.sidebar.text_input("Пароль", type="password")

    if st.sidebar.button("Войти"):
        for row in gs.read(SHEET_USERS):
            if normalize_key(get_field(row, "логин")) == normalize_key(u):
                if str(get_field(row, "пароль")) == p:
                    st.session_state["user"] = u
                    st.rerun()
        st.sidebar.error("Неверный логин или пароль")

    return False

# =========================================
# GEOMETRY CONTEXT
# =========================================

def build_geom_context(section):
    w = safe_float(section["width_mm"])
    h = safe_float(section["height_mm"])
    qty = int(section["qty"])

    area = w * h / 1_000_000
    perimeter = 2 * (w + h) / 1000

    return {
        "width": w,
        "height": h,
        "area": area,
        "perimeter": perimeter,
        "qty": qty,
        "n_sash": int(section.get("n_sash", 0)),
        "is_door": 1 if "Дверь" in section["product_type"] else 0,
        "is_facade": 1 if section.get("kind") == "facade" else 0,
    }

# =========================================
# MATERIAL CALCULATOR
# =========================================

class MaterialCalculator:
    def __init__(self, gs):
        self.ref = gs.read(SHEET_REF1)

    def calculate(self, sections):
        rows = []
        total = 0.0

        for ref_row in self.ref:
            formula = get_field(ref_row, "формула")
            if not formula:
                continue

            qty_total = 0.0
            for s in sections:
                ctx = build_geom_context(s)
                val = safe_eval(str(formula), ctx)
                qty_total += val * ctx["qty"]

            price = safe_float(get_field(ref_row, "цена"), 0)
            if qty_total <= 0:
                continue

            sum_row = qty_total * price
            total += sum_row

            rows.append({
                "Товар": get_field(ref_row, "товар"),
                "Расход": round(qty_total, 3),
                "Цена": price,
                "Сумма": round(sum_row, 2),
            })

        return rows, total

# =========================================
# FINAL CALCULATOR
# =========================================

class FinalCalculator:
    def calculate(self, sections, material_sum):
        total_area = sum(
            build_geom_context(s)["area"] * s["qty"]
            for s in sections
        )
        total_perimeter = sum(
            build_geom_context(s)["perimeter"] * s["qty"]
            for s in sections
        )

        base_sum = material_sum
        ensure = base_sum * ENSURE_COEF
        total = base_sum + ensure

        return {
            "area": total_area,
            "perimeter": total_perimeter,
            "base": base_sum,
            "ensure": ensure,
            "total": total,
        }

# =========================================
# UI HELPERS
# =========================================

def section_form(idx, prefix):
    st.markdown(f"### Позиция {idx+1}")
    c1, c2, c3 = st.columns(3)
    w = c1.number_input("Ширина, мм", 100.0, key=f"{prefix}_w")
    h = c2.number_input("Высота, мм", 100.0, key=f"{prefix}_h")
    q = c3.number_input("Кол-во", 1, step=1, key=f"{prefix}_q")

    n_sash = st.number_input("Кол-во створок", 0, step=1, key=f"{prefix}_s")

    return {
        "product_type": st.session_state["product_type"],
        "profile_system": st.session_state["profile_system"],
        "width_mm": w,
        "height_mm": h,
        "qty": q,
        "n_sash": n_sash,
    }

# =========================================
# MAIN
# =========================================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title("🏗 Axis Pro GF")

    gs = GoogleSheets(GSPREAD_SHEET_ID)
    if not login(gs):
        st.stop()

    with st.sidebar:
        st.session_state["product_type"] = st.selectbox(
            "Тип изделия",
            ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"],
        )
        st.session_state["profile_system"] = st.selectbox(
            "Система профиля",
            ["ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"],
        )

    count = st.number_input("Кол-во позиций", 1, step=1)

    sections = []
    for i in range(count):
        sections.append(section_form(i, f"s{i}"))

    if st.button("🚀 Рассчитать"):
        mat_rows, mat_sum = MaterialCalculator(gs).calculate(sections)
        fin = FinalCalculator().calculate(sections, mat_sum)

        st.success(f"ИТОГО К ОПЛАТЕ: {round(fin['total'],2)}")

        st.subheader("📐 Геометрия")
        st.write(f"Площадь всего: {round(fin['area'],3)} м²")
        st.write(f"Периметр всего: {round(fin['perimeter'],3)} м")

        st.subheader("🧱 Материалы")
        if mat_rows:
            st.dataframe(pd.DataFrame(mat_rows))
        st.write(f"Материалы: {round(mat_sum,2)}")
        st.write(f"Обеспечение 65%: {round(fin['ensure'],2)}")

if __name__ == "__main__":
    main()
