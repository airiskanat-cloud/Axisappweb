# =========================================
# Axis Pro GF v17.2 — Facade Calculator
# Фикс: Автоматическое определение ключей Google
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
    if v is None: return ""
    return " ".join(str(v).replace("\xa0", " ").lower().split())

def safe_float(v, default=0.0):
    try:
        if v is None: return default
        s = str(v).replace("\xa0", "").replace(" ", "").replace(",", ".")
        return float(s) if s else default
    except Exception:
        return default

def get_field(row: dict, needle: str, default=None):
    needle = needle.lower()
    for k, v in row.items():
        if k and needle in str(k).lower(): return v
    return default

# =========================================
# SAFE AST EVAL
# =========================================

_ALLOWED_OPS = {
    ast.Add: op.add, ast.Sub: op.sub, ast.Mult: op.mul,
    ast.Div: op.truediv, ast.Pow: op.pow, ast.USub: op.neg,
    ast.UAdd: op.pos, ast.Mod: op.mod,
}

def _eval_node(node, names):
    if isinstance(node, ast.Expression): return _eval_node(node.body, names)
    if isinstance(node, ast.Constant): return node.value
    if isinstance(node, ast.Name):
        if node.id in names: return names[node.id]
        raise ValueError(f"Unknown var {node.id}")
    if isinstance(node, ast.BinOp):
        return _ALLOWED_OPS[type(node.op)](_eval_node(node.left, names), _eval_node(node.right, names))
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
    if not formula: return 0.0
    try:
        ctx = {k: safe_float(v) for k, v in context.items()}
        ctx["math"] = math
        return float(_eval_node(ast.parse(formula, mode="eval"), ctx))
    except Exception as e:
        logger.error("Formula error: %s | %s", formula, e)
        return 0.0

# =========================================
# GOOGLE SHEETS CLIENT (Универсальный фикс)
# =========================================

class GoogleSheets:

    @st.cache_resource
    def auth(_self):
        """
        Ищет ключ под любым из возможных имен (Render или локально)
        и автоматически определяет формат (Base64 или прямой JSON).
        """
        # Проверяем все варианты имен переменных
        key_source = st.secrets.get("gcp_service_account") or \
                     st.secrets.get("GCP_SA_KEYFILE_JSON_BASE64") or \
                     st.secrets.get("GCP_SA_KEYFILE_JSON")

        if not key_source:
            st.error("❌ Ключ не найден! Проверьте, что в Render Environment Variables создана переменная 'gcp_service_account'.")
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
        if name not in self.cache: self.cache[name] = self.book.worksheet(name)
        return self.cache[name]

    @st.cache_data(ttl=1800)
    def read(_self, sheet_name):
        return _self.ws(sheet_name).get_all_records()

# =========================================
# БЛОКИ ЛОГИКИ И ИНТЕРФЕЙСА (Без изменений)
# =========================================

def login(gs: GoogleSheets):
    if "user" in st.session_state: return True
    st.sidebar.title("🔐 Вход")
    l_v, p_v = st.sidebar.text_input("Логин"), st.sidebar.text_input("Пароль", type="password")
    if st.sidebar.button("Войти"):
        for u in gs.read(SHEET_USERS):
            if normalize_key(get_field(u, "логин")) == normalize_key(l_v) and str(get_field(u, "пароль")) == p_v:
                st.session_state["user"] = l_v
                st.rerun()
        st.sidebar.error("Неверный логин или пароль")
    return False

def build_geom_context(s: dict):
    w, h, q = safe_float(s.get("width_mm")), safe_float(s.get("height_mm")), int(s.get("qty", 1))
    l, c, r, t = safe_float(s.get("left")), safe_float(s.get("center")), safe_float(s.get("right")), safe_float(s.get("top"))
    area, peri = (w * h) / 1e6, 2 * (w + h) / 1000
    n_v = sum(1 for x in (l, c, r) if x > 0)
    n_imp = max(0, n_v - 1) + (1 if t > 0 else 0)
    ns = int(s.get("n_sash", 0))
    return {
        "width": w, "height": h, "area": area, "perimeter": peri, "qty": q,
        "n_impost": n_imp, "n_frame_rect": 1 + n_imp, "n_corners": 4 * (1 + n_imp),
        "n_sash": ns, "n_sash_active": 1 if ns > 0 else 0, "sash_w": safe_float(s.get("sash_w")), "sash_h": safe_float(s.get("sash_h")),
        "is_door": 1 if s.get("kind") == "door" else 0, "is_facade": 1 if s.get("kind") == "facade" else 0
    }

class MaterialCalculator:
    def __init__(self, gs): self.gs = gs
    def calculate(self, sections):
        ref1 = self.gs.read(SHEET_REF1)
        res, total = [], 0.0
        for row in ref1:
            row_t, row_p, form = str(get_field(row, "тип издел")), str(get_field(row, "система проф")), get_field(row, "формула_python")
            if not form: continue
            q_total = 0.0
            for s in sections:
                if (not row_t or row_t == s["product_type"]) and (not row_p or row_p == s["profile_system"]):
                    ctx = build_geom_context(s)
                    q_total += safe_eval(str(form), ctx) * ctx["qty"]
            if q_total <= 0: continue
            p, n = safe_float(get_field(row, "цена за")), safe_float(get_field(row, "кол-во норм"), 1)
            real_q = math.ceil(q_total / n) * n if n > 0 else q_total
            total += real_q * p
            res.append({"Товар": str(get_field(row, "товар")), "Факт. расход": round(q_total, 3), "К отгрузке": real_q, "Сумма": round(real_q * p, 2)})
        return res, total

class FinalCalculator:
    def __init__(self, gs): self.gs = gs; self.ref2 = gs.read(SHEET_REF2)
    def _get_p(self, kw):
        for row in self.ref2:
            for k, v in row.items():
                if k and all(w in normalize_key(k) for w in kw): return safe_float(v)
        return 0.0
    def calculate(self, sections, mat_sum, g_type, ton, ass, mon):
        area = sum((safe_float(s["width_mm"])*safe_float(s["height_mm"])/1e6)*int(s.get("qty", 1)) for s in sections)
        g_p = 0.0
        for row in self.ref2:
            if any("тип стеклопак" in normalize_key(k) and normalize_key(v) == normalize_key(g_type) for k,v in row.items()):
                g_p = next((safe_float(vv) for kk,vv in row.items() if "стоимость" in normalize_key(kk)), 0.0)
        if g_p == 0: g_p = self._get_p(["стеклопакет", "м"])
        rows = [("Стеклопакет", g_p, "м²", area * g_p)]
        if ton: rows.append(("Тонировка", self._get_p(["тониров"]), "м²", area * self._get_p(["тониров"])))
        if ass: rows.append(("Сборка", self._get_p(["сборк"]), "м²", area * self._get_p(["сборк"])))
        if mon: rows.append(("Монтаж", self._get_p(["монтаж"]), "м²", area * self._get_p(["монтаж"])))
        rows.append(("Материалы", "-", "-", mat_sum))
        base = sum(r[3] for r in rows)
        ensure = base * 0.65
        rows.append(("Обеспечение 65%", "", "", ensure))
        rows.append(("ИТОГО", "", "", base + ensure))
        return rows, base + ensure

def section_form(title, p_t, p_s, kp=""):
    st.subheader(title)
    c1, c2, c3 = st.columns(3)
    w, h, q = c1.number_input("Ширина", 100.0, step=10.0, key=f"{kp}w"), c2.number_input("Высота", 100.0, step=10.0, key=f"{kp}h"), c3.number_input("Кол-во", 1, step=1, key=f"{kp}q")
    i1, i2, i3, i4 = st.columns(4)
    l, c, r, t = i1.number_input("LEFT", 0.0, key=f"{kp}l"), i2.number_input("CENTER", 0.0, key=f"{kp}c"), i3.number_input("RIGHT", 0.0, key=f"{kp}r"), i4.number_input("TOP", 0.0, key=f"{kp}t")
    ns, sw, sh = 0, 0.0, 0.0
    if "Окно с откр." in p_t or "Дверь" in p_t:
        ns = st.number_input("Створки", 1, key=f"{kp}ns")
        sw, sh = st.columns(2)[0].number_input("Ширина ств.", 200.0, key=f"{kp}sw"), st.columns(2)[1].number_input("Высота ств.", 200.0, key=f"{kp}sh")
    return {"product_type": p_t, "profile_system": p_s, "kind": "door" if "Дверь" in p_t else "window", "width_mm": w, "height_mm": h, "qty": q, "left": l, "center": c, "right": r, "top": t, "n_sash": ns, "sash_w": sw, "sash_h": sh}

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title("🏗️ Axis Pro GF — Калькулятор")
    gs = GoogleSheets(GSPREAD_SHEET_ID)
    if not login(gs): st.stop()
    with st.sidebar:
        p_m = st.selectbox("Тип", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_s = st.selectbox("Система", ["ALG 2030-63C", "ALG 2030-55C", "ALG 2030-73C", "ALG 2030-45C", "ALG 2030-Slim", "Ruit 50F"])
        gt, ton, ass, mon = st.text_input("Стеклопакет", "двойной"), st.checkbox("Тонировка"), st.checkbox("Сборка"), st.checkbox("Монтаж")
    sections = []
    if p_m != "Фасад": sections.append(section_form("Параметры", p_m, p_s, "m"))
    else:
        f = section_form("Каркас", "Фасад", p_s, "f"); f["kind"] = "facade"; sections.append(f)
        if "fc" not in st.session_state: st.session_state.fc = 0
        if st.button("➕ Добавить вставку"): st.session_state.fc += 1
        for i in range(st.session_state.fc):
            it = st.selectbox(f"Тип #{i+1}", ["Окно с откр.", "Окно глух.", "Дверь 1 створч."], key=f"it{i}")
            sections.append(section_form(f"Вставка #{i+1}", it, p_s, f"i{i}"))
    if st.button("🚀 Рассчитать", type="primary"):
        m_r, m_s = MaterialCalculator(gs).calculate(sections)
        f_r, total = FinalCalculator(gs).calculate(sections, m_s, gt, ton, ass, mon)
        st.success(f"ИТОГО: {round(total, 2)}")
        t1, t2 = st.tabs(["Материалы", "Итог"])
        with t1: st.dataframe(pd.DataFrame(m_r), use_container_width=True)
        with t2: st.dataframe(pd.DataFrame(f_r, columns=["Название", "Цена", "Ед.", "Сумма"]), use_container_width=True)

if __name__ == "__main__": main()
