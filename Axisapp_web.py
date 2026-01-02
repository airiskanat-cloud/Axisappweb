# =========================================
# Axis Pro GF v17.1 — Facade Calculator
# Фикс авторизации: автоматический поиск ключей
# =========================================

import math
import ast
import operator as op
import base64
import json
import logging
import sys

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
# GOOGLE SHEETS CLIENT (V17 MULTI-AUTH)
# =========================================

class GoogleSheets:

    @st.cache_resource
    def auth(_self):
        """
        Универсальный поиск ключа: ищет любое из 3-х имен и понимает как Base64, так и чистый JSON.
        """
        key_source = st.secrets.get("GCP_SA_KEYFILE_JSON_BASE64") or \
                     st.secrets.get("GCP_SA_KEYFILE_JSON") or \
                     st.secrets.get("gcp_service_account")

        if not key_source:
            st.error("❌ Ключ не найден! Добавьте переменную 'gcp_service_account' в Environment Variables на Render.")
            st.stop()

        try:
            # 1. Пробуем декодировать как Base64
            try:
                decoded = base64.b64decode(key_source).decode("utf-8")
                info = json.loads(decoded)
            except Exception:
                # 2. Если не Base64, значит это прямой текст JSON
                info = json.loads(key_source)
                
            creds = Credentials.from_service_account_info(
                info,
                scopes=[
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive",
                ],
            )
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"❌ Ошибка в формате ключа в Render: {e}")
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

# =========================================
# LOGIN
# =========================================

def login(gs: GoogleSheets):
    if "user" in st.session_state:
        return True

    st.sidebar.title("🔐 Вход")
    l_val = st.sidebar.text_input("Логин")
    p_val = st.sidebar.text_input("Пароль", type="password")

    if st.sidebar.button("Войти"):
        users = gs.read(SHEET_USERS)
        for u in users:
            if normalize_key(get_field(u, "логин")) == normalize_key(l_val):
                if str(get_field(u, "пароль")) == p_val:
                    st.session_state["user"] = l_val
                    st.rerun()
        st.sidebar.error("Неверный логин или пароль")
    return False

# =========================================
# GEOM CONTEXT & CALCULATORS
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
    n_sash = int(section.get("n_sash", 0))
    sash_w = safe_float(section.get("sash_w", 0))
    sash_h = safe_float(section.get("sash_h", 0))
    kind = section.get("kind")

    return {
        "width": width, "height": height, "area": area, "perimeter": perimeter, "qty": qty,
        "left": left, "center": center, "right": right, "top": top,
        "n_imp_vert": n_imp_vert, "n_imp_hor": n_imp_hor, "n_impost": n_impost,
        "n_frame_rect": 1 + n_impost, "n_corners": 4 * (1 + n_impost),
        "n_sash": n_sash, "n_sash_active": 1 if n_sash > 0 else 0,
        "n_sash_passive": max(n_sash - 1, 0),
        "sash_w": sash_w, "sash_h": sash_h,
        "is_door": 1 if kind == "door" else 0, "is_facade": 1 if kind == "facade" else 0,
    }

class MaterialCalculator:
    def __init__(self, gs: GoogleSheets):
        self.gs = gs
    def calculate(self, sections: list):
        ref1 = self.gs.read(SHEET_REF1)
        results, total_sum = [], 0.0
        for row in ref1:
            row_type = str(get_field(row, "тип издел", "") or "").strip()
            row_profile = str(get_field(row, "система проф", "") or "").strip()
            formula = get_field(row, "формула_python")
            if not formula: continue
            qty_total = 0.0
            for s in sections:
                if row_type and row_type != s["product_type"]: continue
                if row_profile and row_profile != s["profile_system"]: continue
                ctx = build_geom_context(s)
                qty_total += safe_eval(str(formula), ctx) * ctx["qty"]
            price = safe_float(get_field(row, "цена за"))
            norm = safe_float(get_field(row, "кол-во норм"), 1)
            if qty_total <= 0: continue
            real_qty = math.ceil(qty_total / norm) * norm if norm > 0 else qty_total
            sum_row = real_qty * price
            total_sum += sum_row
            results.append({
                "Тип изделия": row_type, "Система профиля": row_profile,
                "Тип элемента": str(get_field(row, "тип элемент", "")),
                "Товар": str(get_field(row, "товар", "")),
                "Факт. расход": round(qty_total, 3), "К отгрузке": real_qty,
                "Цена": price, "Сумма": round(sum_row, 2),
            })
        return results, total_sum

class FinalCalculator:
    def __init__(self, gs: GoogleSheets):
        self.gs = gs
        self.ref2 = self.gs.read(SHEET_REF2)
    def _find_price(self, keywords: list, default=0.0):
        for row in self.ref2:
            for k, v in row.items():
                if k and all(word in normalize_key(k) for word in keywords):
                    return safe_float(v, default)
        return default
    def calculate(self, sections: list, material_sum: float, glass_type: str, toning: bool, assembly: bool, montage: bool):
        area = sum((safe_float(s["width_mm"])*safe_float(s["height_mm"])/1e6)*int(s.get("qty", 1)) for s in sections)
        rows = []
        # Стеклопакет
        g_price = 0.0
        g_type_norm = normalize_key(glass_type)
        for row in self.ref2:
            if any("тип стеклопак" in normalize_key(k) and normalize_key(v) == g_type_norm for k,v in row.items()):
                g_price = next((safe_float(vv) for kk,vv in row.items() if "стоимость" in normalize_key(kk)), 0.0)
                break
        if g_price == 0: g_price = self._find_price(["стеклопакет", "м"])
        rows.append(("Стеклопакет", g_price, "м²", area * g_price))
        if toning: rows.append(("Тонировка", self._find_price(["тониров"]), "м²", area * self._find_price(["тониров"])))
        if assembly: rows.append(("Сборка", self._find_price(["сборк"]), "м²", area * self._find_price(["сборк"])))
        if montage: rows.append(("Монтаж", self._find_price(["монтаж"]), "м²", area * self._find_price(["монтаж"])))
        rows.append(("Материалы", "-", "-", material_sum))
        base = sum(r[3] for r in rows)
        ensure = base * 0.65
        rows.append(("Обеспечение 65%", "", "", ensure))
        rows.append(("ИТОГО", "", "", base + ensure))
        return rows, base + ensure

# =========================================
# UI
# =========================================

def section_form(title, p_type, p_sys, k_prefix=""):
    st.subheader(title)
    c1, c2, c3 = st.columns(3)
    w = c1.number_input("Ширина, мм", 100.0, step=10.0, key=f"{k_prefix}w")
    h = c2.number_input("Высота, мм", 100.0, step=10.0, key=f"{k_prefix}h")
    q = c3.number_input("Кол-во (N)", 1, step=1, key=f"{k_prefix}q")
    st.markdown("**Импосты**")
    i1, i2, i3, i4 = st.columns(4)
    l = i1.number_input("LEFT", 0.0, step=10.0, key=f"{k_prefix}l")
    c = i2.number_input("CENTER", 0.0, step=10.0, key=f"{k_prefix}c")
    r = i3.number_input("RIGHT", 0.0, step=10.0, key=f"{k_prefix}r")
    t = i4.number_input("TOP", 0.0, step=10.0, key=f"{k_prefix}t")
    ns, sw, sh = 0, 0.0, 0.0
    if "Окно с откр." in p_type or "Дверь" in p_type:
        ns = st.number_input("Кол-во створок", 1, step=1, key=f"{k_prefix}ns")
        s1, s2 = st.columns(2)
        sw = s1.number_input("Ширина створки, мм", 200.0, step=10.0, key=f"{k_prefix}sw")
        sh = s2.number_input("Высота створки, мм", 200.0, step=10.0, key=f"{k_prefix}sh")
    return {"product_type": p_type, "profile_system": p_sys, "kind": "door" if "Дверь" in p_type else "window",
            "width_mm": w, "height_mm": h, "qty": q, "left": l, "center": c, "right": r, "top": t,
            "n_sash": ns, "sash_w": sw, "sash_h": sh}

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title("🏗️ Axis Pro GF — Калькулятор")
    gs = GoogleSheets(GSPREAD_SHEET_ID)
    if not login(gs): st.stop()
    
    with st.sidebar:
        st.header("Заказ")
        p_main = st.selectbox("Тип", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_sys = st.selectbox("Система", ["ALG 2030-63C", "ALG 2030-55C", "ALG 2030-73C", "ALG 2030-45C", "ALG 2030-Slim", "Ruit 50F"])
        g_type = st.text_input("Стеклопакет", "двойной")
        ton, ass, mon = st.checkbox("Тонировка"), st.checkbox("Сборка"), st.checkbox("Монтаж")

    sections = []
    if p_main != "Фасад":
        sections.append(section_form("Параметры", p_main, p_sys, "m"))
    else:
        f = section_form("Каркас фасада", "Фасад", p_sys, "fm"); f["kind"] = "facade"
        sections.append(f)
        if "f_cnt" not in st.session_state: st.session_state.f_cnt = 0
        if st.button("➕ Добавить вставку"): st.session_state.f_cnt += 1
        for i in range(st.session_state.f_cnt):
            st.markdown(f"---")
            it = st.selectbox(f"Тип вставки #{i+1}", ["Окно с откр.", "Окно глух.", "Дверь 1 створч."], key=f"it{i}")
            isys = st.selectbox(f"Система #{i+1}", ["ALG 2030-63C", "ALG 2030-55C"], key=f"is{i}")
            sections.append(section_form(f"Вставка #{i+1}", it, isys, f"i{i}"))

    if st.button("🚀 Рассчитать", type="primary"):
        m_rows, m_sum = MaterialCalculator(gs).calculate(sections)
        f_rows, total = FinalCalculator(gs).calculate(sections, m_sum, g_type, ton, ass, mon)
        st.success(f"ИТОГО: {round(total, 2)}")
        t1, t2 = st.tabs(["Материалы", "Итог"])
        with t1: st.dataframe(pd.DataFrame(m_rows), use_container_width=True)
        with t2: st.dataframe(pd.DataFrame(f_rows, columns=["Название", "Цена", "Ед.", "Сумма"]), use_container_width=True)

if __name__ == "__main__":
    main()
