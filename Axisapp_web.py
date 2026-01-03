# =========================================
# Axis Pro GF v17.4 — Facade Calculator
# Google Sheets auth = v16 (SAFE) + os.environ FIX
# =========================================

import math
import ast
import operator as op
import base64
import json
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
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"
SHEET_FORM = "ЗАПРОСЫ"

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
# SAFE AST EVAL (v15 logic)
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
        raise ValueError(f"Unknown var {node.id}")

    if isinstance(node, ast.BinOp):
        return _ALLOWED_OPS[type(node.op)](
            _eval_node(node.left, names),
            _eval_node(node.right, names),
        )

    if isinstance(node, ast.UnaryOp):
        return _ALLOWED_OPS[type(node.op)](
            _eval_node(node.operand, names)
        )

    if isinstance(node, ast.Call):
        if isinstance(node.func, ast.Attribute) and node.func.value.id == "math":
            fn = getattr(math, node.func.attr)
            args = [_eval_node(a, names) for a in node.args]
            return fn(*args)

        if isinstance(node.func, ast.Name) and node.func.id in ("min", "max"):
            args = [_eval_node(a, names) for a in node.args]
            return globals()[node.func.id](*args)

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
        logger.error("Formula error: %s | %s", formula, e)
        return 0.0

# =========================================
# GOOGLE SHEETS CLIENT (FIXED AUTH)
# =========================================

class GoogleSheets:

    @st.cache_resource
    def auth(_self):
        # Приоритет: os.environ (Render), затем st.secrets (Local/Streamlit Cloud)
        key_source = os.environ.get("gcp_service_account") or \
                     os.environ.get("GCP_SA_KEYFILE_JSON") or \
                     st.secrets.get("gcp_service_account")

        if not key_source:
            st.error("❌ Ключ авторизации не найден в переменных окружения.")
            st.stop()

        # Если ключ пришел как объект из secrets, превращаем его в строку для обработки
        if not isinstance(key_source, str):
            try:
                info = dict(key_source)
            except:
                st.error("❌ Формат ключа в st.secrets некорректен.")
                st.stop()
        else:
            # Если это строка, пробуем Base64 или прямой JSON
            key_source = key_source.strip()
            try:
                # 1. Пробуем Base64 (самый надежный способ для Render)
                info = json.loads(base64.b64decode(key_source).decode("utf-8"))
            except Exception:
                try:
                    # 2. Пробуем прямой JSON
                    info = json.loads(key_source)
                except Exception as e:
                    st.error(f"❌ Ошибка в формате JSON ключа: {e}")
                    st.info("Убедитесь, что в Render переменная окружения содержит валидный JSON с двойными кавычками.")
                    st.stop()

        try:
            creds = Credentials.from_service_account_info(
                info,
                scopes=[
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive",
                ],
            )
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"❌ Ошибка Google Auth: {e}")
            st.stop()

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
        ws = _self.ws(sheet_name)
        rows = ws.get_all_records()
        return rows

    def append(self, sheet_name, row):
        self.ws(sheet_name).append_row(row, value_input_option="USER_ENTERED")

# =========================================
# LOGIN
# =========================================

def login(gs: GoogleSheets):
    if "user" in st.session_state:
        return True

    st.sidebar.title("🔐 Вход")
    login_user = st.sidebar.text_input("Логин")
    pwd = st.sidebar.text_input("Пароль", type="password")

    if st.sidebar.button("Войти"):
        users = gs.read(SHEET_USERS)
        for u in users:
            if normalize_key(get_field(u, "логин")) == normalize_key(login_user):
                if str(get_field(u, "пароль")) == pwd:
                    st.session_state["user"] = login_user
                    st.rerun()
        st.sidebar.error("Неверный логин или пароль")

    return False

# =========================================
# DATA MODEL: SECTION / FACADE
# =========================================

def build_geom_context(section: dict):
    width = safe_float(section.get("width_mm", 0))
    height = safe_float(section.get("height_mm", 0))
    qty = int(section.get("qty", 1))

    left = safe_float(section.get("left", 0))
    center = safe_float(section.get("center", 0))
    right = safe_float(section.get("right", 0))
    top = safe_float(section.get("top", 0))

    area = (width * height) / 1_000_000
    perimeter = 2 * (width + height) / 1000

    n_vert = sum(1 for x in (left, center, right) if x > 0)
    n_imp_vert = max(0, n_vert - 1)
    n_imp_hor = 1 if top > 0 else 0

    n_impost = n_imp_vert + n_imp_hor
    n_frame_rect = 1 + n_impost
    n_corners = 4 * n_frame_rect

    n_sash = int(section.get("n_sash", 0))
    sash_w = safe_float(section.get("sash_w", 0))
    sash_h = safe_float(section.get("sash_h", 0))

    kind = section.get("kind")

    ctx = {
        "width": width, "height": height, "area": area, "perimeter": perimeter, "qty": qty,
        "left": left, "center": center, "right": right, "top": top,
        "n_imp_vert": n_imp_vert, "n_imp_hor": n_imp_hor, "n_impost": n_impost,
        "n_frame_rect": n_frame_rect, "n_corners": n_corners,
        "n_sash": n_sash, "n_sash_active": 1 if n_sash > 0 else 0,
        "n_sash_passive": max(n_sash - 1, 0),
        "sash_width": sash_w, "sash_height": sash_h, "sash_w": sash_w, "sash_h": sash_h,
        "is_door": 1 if kind == "door" else 0, "is_facade": 1 if kind == "facade" else 0,
    }
    return ctx

# =========================================
# MATERIAL CALCULATOR
# =========================================

class MaterialCalculator:
    def __init__(self, gs: GoogleSheets):
        self.gs = gs

    def calculate(self, sections: list):
        ref1 = self.gs.read(SHEET_REF1)
        results = []
        total_sum = 0.0

        for row in ref1:
            row_type = str(get_field(row, "тип издел", "") or "").strip()
            row_profile = str(get_field(row, "система проф", "") or "").strip()
            elem_type = str(get_field(row, "тип элемент", "") or "").strip()
            product = str(get_field(row, "товар", "") or "").strip()
            formula = get_field(row, "формула_python")
            
            if not formula: continue

            qty_total = 0.0
            for s in sections:
                if row_type and row_type != s["product_type"]: continue
                if row_profile and row_profile != s["profile_system"]: continue

                ctx = build_geom_context(s)
                val = safe_eval(str(formula), ctx)
                qty_total += val * ctx["qty"]

            price = safe_float(get_field(row, "цена за"))
            norm = safe_float(get_field(row, "кол-во норм"), 1)

            if qty_total <= 0: continue

            if norm > 0:
                ship_qty = math.ceil(qty_total / norm)
                real_qty = ship_qty * norm
            else:
                real_qty = qty_total

            sum_row = real_qty * price
            total_sum += sum_row

            results.append({
                "Тип изделия": row_type,
                "Система профиля": row_profile,
                "Тип элемента": elem_type,
                "Товар": product,
                "Факт. расход": round(qty_total, 3),
                "К отгрузке": real_qty,
                "Цена": price,
                "Сумма": round(sum_row, 2),
            })
        return results, total_sum

# =========================================
# FINAL CALCULATOR
# =========================================

class FinalCalculator:
    def __init__(self, gs: GoogleSheets):
        self.gs = gs
        self.ref2 = self.gs.read(SHEET_REF2)

    def _find_price(self, keywords: list, default=0.0):
        for row in self.ref2:
            for k, v in row.items():
                if not k: continue
                key = normalize_key(k)
                if all(word in key for word in keywords):
                    return safe_float(v, default)
        return default

    def price_glass(self, glass_type: str):
        glass_type_norm = normalize_key(glass_type)
        for row in self.ref2:
            for k, v in row.items():
                if "тип стеклопак" in normalize_key(k):
                    if normalize_key(v) == glass_type_norm:
                        for kk, vv in row.items():
                            if "стоимость" in normalize_key(kk):
                                return safe_float(vv)
        return self._find_price(["стеклопакет", "м"], 0.0)

    def calculate(self, sections: list, material_sum: float, glass_type: str, toning: bool, assembly: bool, montage: bool):
        total_area = sum((safe_float(s["width_mm"]) * safe_float(s["height_mm"]) / 1_000_000) * int(s.get("qty", 1)) for s in sections)
        rows = []

        glass_price = self.price_glass(glass_type)
        glass_sum = total_area * glass_price
        rows.append(("Стеклопакет", glass_price, "м²", glass_sum))

        if toning:
            p = self._find_price(["тониров", "м"])
            rows.append(("Тонировка", p, "м²", total_area * p))
        if assembly:
            p = self._find_price(["сборк", "м"])
            rows.append(("Сборка", p, "м²", total_area * p))
        if montage:
            p = self._find_price(["монтаж", "м"])
            rows.append(("Монтаж", p, "м²", total_area * p))

        rows.append(("Материалы", "-", "-", material_sum))
        
        base_sum = material_sum + glass_sum + sum(r[3] for r in rows if r[0] in ["Тонировка", "Сборка", "Монтаж"])
        ensure = base_sum * 0.65
        rows.append(("Обеспечение 65%", "", "", ensure))
        total = base_sum + ensure
        rows.append(("ИТОГО", "", "", total))

        return rows, total

# =========================================
# UI
# =========================================

def section_form(title, product_type, profile_system, key_prefix=""):
    st.subheader(title)
    c1, c2, c3 = st.columns(3)
    width = c1.number_input("Ширина, мм", min_value=100.0, step=10.0, key=f"{key_prefix}_w")
    height = c2.number_input("Высота, мм", min_value=100.0, step=10.0, key=f"{key_prefix}_h")
    qty = c3.number_input("Кол-во (N)", min_value=1, step=1, value=1, key=f"{key_prefix}_q")

    st.markdown("**Импосты (если есть)**")
    i1, i2, i3, i4 = st.columns(4)
    left = i1.number_input("LEFT", min_value=0.0, step=10.0, key=f"{key_prefix}_l")
    center = i2.number_input("CENTER", min_value=0.0, step=10.0, key=f"{key_prefix}_c")
    right = i3.number_input("RIGHT", min_value=0.0, step=10.0, key=f"{key_prefix}_r")
    top = i4.number_input("TOP", min_value=0.0, step=10.0, key=f"{key_prefix}_t")

    n_sash, sash_w, sash_h = 0, 0.0, 0.0
    if "Окно с откр." in product_type or "Дверь" in product_type:
        n_sash = st.number_input("Кол-во створок", min_value=1, step=1, value=1, key=f"{key_prefix}_ns")
        s1, s2 = st.columns(2)
        sash_w = s1.number_input("Ширина створки, мм", min_value=200.0, step=10.0, key=f"{key_prefix}_sw")
        sash_h = s2.number_input("Высота створки, мм", min_value=200.0, step=10.0, key=f"{key_prefix}_sh")

    return {
        "product_type": product_type, "profile_system": profile_system,
        "kind": "door" if "Дверь" in product_type else "window",
        "width_mm": width, "height_mm": height, "qty": qty,
        "left": left, "center": center, "right": right, "top": top,
        "n_sash": n_sash, "sash_w": sash_w, "sash_h": sash_h,
    }

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title("🏗️ Axis Pro GF — Калькулятор")

    gs = GoogleSheets(GSPREAD_SHEET_ID)
    if not login(gs): st.stop()

    with st.sidebar:
        st.header("Параметры заказа")
        product_main = st.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        profile_system = st.selectbox("Система профиля", ["ALG 2030-63C", "ALG 2030-55C", "ALG 2030-73C", "ALG 2030-45C", "ALG 2030-Slim", "Ruit 50F"])
        glass_type = st.text_input("Тип стеклопакета", value="двойной")
        toning = st.checkbox("Тонировка")
        assembly = st.checkbox("Сборка")
        montage = st.checkbox("Монтаж")

    sections = []
    if product_main != "Фасад":
        sections.append(section_form("Параметры изделия", product_main, profile_system, key_prefix="main"))
    else:
        facade = section_form("Каркас фасада", "Фасад", profile_system, key_prefix="facade_frame")
        facade["kind"] = "facade"
        sections.append(facade)
        st.markdown("---")
        if "f_sections_count" not in st.session_state: st.session_state["f_sections_count"] = 0
        if st.button("➕ Добавить вставку"): st.session_state["f_sections_count"] += 1
        for idx in range(st.session_state["f_sections_count"]):
            ptype = st.selectbox(f"Тип вставки #{idx+1}", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч."], key=f"ptype_{idx}")
            psys = st.selectbox(f"Система профиля #{idx+1}", ["ALG 2030-63C", "ALG 2030-55C", "ALG 2030-73C", "ALG 2030-45C", "ALG 2030-Slim", "Ruit 50F"], key=f"psys_{idx}")
            sections.append(section_form(f"Параметры вставки #{idx+1}", ptype, psys, key_prefix=f"ins_{idx}"))

    if st.button("🚀 Рассчитать", type="primary"):
        mat_rows, mat_sum = MaterialCalculator(gs).calculate(sections)
        fin_rows, total = FinalCalculator(gs).calculate(sections, mat_sum, glass_type, toning, assembly, montage)
        st.success(f"ИТОГО к оплате: {round(total, 2)}")
        tab1, tab2 = st.tabs(["Материалы", "Итог"])
        with tab1:
            if mat_rows: st.dataframe(pd.DataFrame(mat_rows), use_container_width=True)
            st.write(f"**Итого материалы:** {round(mat_sum,2)}")
        with tab2:
            st.dataframe(pd.DataFrame(fin_rows, columns=["Наименование", "Цена", "Ед.", "Сумма"]), use_container_width=True)

if __name__ == "__main__":
    main()
