# =========================================
# Axis Pro GF — Calculator
# Полная сборка: Части 1-6 + Исправления
# =========================================

import math
import ast
import operator as op
import json
import logging
import sys
import os
from datetime import datetime

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# =========================================
# CONFIG
# =========================================

APP_TITLE = "Axis Pro GF — Фасады / Окна / Двери"
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"
SHEET_REQUESTS = "ЗАПРОСЫ"

ENSURE_PERCENT = 0.65
KP_TEMPLATE_PATH = "template_kp.xlsx" # Убедитесь, что этот файл есть в репозитории

# =========================================
# LOGGER
# =========================================

logger = logging.getLogger("axis_pro_gf")
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

# =========================================
# UTILS
# =========================================

def safe_float(value, default=0.0):
    try:
        if value is None: return default
        s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
        return float(s) if s else default
    except Exception: return default

def safe_int(value, default=0):
    try: return int(float(value))
    except Exception: return default

def normalize_text(value):
    if value is None: return ""
    return " ".join(str(value).strip().split())

def get_field(row, needle, default=None):
    needle = needle.lower()
    for key, value in row.items():
        if key and needle in key.lower(): return value
    return default

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# =========================================
# SAFE AST EVAL
# =========================================

_ALLOWED_OPS = {
    ast.Add: op.add, ast.Sub: op.sub, ast.Mult: op.mul,
    ast.Div: op.truediv, ast.Pow: op.pow, ast.USub: op.neg,
    ast.UAdd: op.pos, ast.Mod: op.mod,
}

def _eval_node(node, context):
    if isinstance(node, ast.Expression): return _eval_node(node.body, context)
    if isinstance(node, ast.Constant): return node.value
    if isinstance(node, ast.Name):
        if node.id in context: return context[node.id]
        raise ValueError(f"Unknown variable: {node.id}")
    if isinstance(node, ast.BinOp):
        return _ALLOWED_OPS[type(node.op)](_eval_node(node.left, context), _eval_node(node.right, context))
    if isinstance(node, ast.UnaryOp):
        return _ALLOWED_OPS[type(node.op)](_eval_node(node.operand, context))
    if isinstance(node, ast.Call):
        if isinstance(node.func, ast.Name) and node.func.id in ("min", "max"):
            args = [_eval_node(a, context) for a in node.args]
            return globals()[node.func.id](*args)
        if isinstance(node.func, ast.Attribute) and node.func.value.id == "math":
            fn = getattr(math, node.func.attr)
            return fn(*[_eval_node(a, context) for a in node.args])
    raise ValueError("Unsafe expression")

def safe_eval(formula, context):
    if not formula: return 0.0
    try:
        prepared = {k: safe_float(v) for k, v in context.items()}
        prepared["math"] = math
        node = ast.parse(str(formula), mode="eval")
        return float(_eval_node(node, prepared))
    except Exception as e:
        logger.error(f"Formula error [{formula}]: {e}")
        return 0.0

# =========================================
# GOOGLE SHEETS CLIENT
# =========================================

class GoogleSheetsClient:
    @st.cache_resource
    def auth(_self):
        # Путь к секрету в Render
        secret_path = "/etc/secrets/gcp.json" 
        
        if not os.path.exists(secret_path):
            # Резервный поиск через переменную окружения
            key_json = os.environ.get("gcp_service_account")
            if key_json:
                info = json.loads(key_json)
                return gspread.authorize(Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]))
            st.error("Service account file not found")
            st.stop()

        credentials = Credentials.from_service_account_file(secret_path, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(credentials)

    def __init__(self, spreadsheet_id):
        self.client = self.auth()
        self.book = self.client.open_by_key(spreadsheet_id)
        self._cache = {}

    def worksheet(self, name):
        if name not in self._cache: self._cache[name] = self.book.worksheet(name)
        return self._cache[name]

    @st.cache_data(ttl=1800)
    def read(_self, sheet_name):
        return _self.worksheet(sheet_name).get_all_records()

    def append_row(self, sheet_name, row):
        self.worksheet(sheet_name).append_row(row, value_input_option="USER_ENTERED")

# =========================================
# LOGIN
# =========================================

def login(gs: GoogleSheetsClient):
    if "user_login" in st.session_state: return True
    st.sidebar.title("Вход")
    login_val = st.sidebar.text_input("Логин")
    pass_val = st.sidebar.text_input("Пароль", type="password")
    if st.sidebar.button("Войти"):
        users = gs.read(SHEET_USERS)
        for user in users:
            if str(get_field(user, "логин", "")).strip() == login_val and str(get_field(user, "пароль", "")).strip() == pass_val:
                st.session_state["user_login"] = login_val
                st.rerun()
        st.sidebar.error("Неверный логин или пароль")
    return False

# =========================================
# GEOMETRY
# =========================================

def build_position_geometry(position):
    w_mm, h_mm, q = safe_float(position.get("width_mm")), safe_float(position.get("height_mm")), safe_int(position.get("qty"), 1)
    w_m, h_m = w_mm / 1000.0, h_mm / 1000.0
    return {"width_mm": w_mm, "height_mm": h_mm, "width_m": w_m, "height_m": h_m, "area_one": w_m * h_m, "perimeter_one": 2*(w_m+h_m), "qty": q, "area_total": w_m * h_m * q, "perimeter_total": 2*(w_m+h_m)*q}

def build_impost_geometry(position):
    l, c, r, t = safe_float(position.get("left_mm")), safe_float(position.get("center_mm")), safe_float(position.get("right_mm")), safe_float(position.get("top_mm"))
    v_cnt = sum(1 for v in (l, c, r) if v > 0)
    return {"impost_vert_count": max(v_cnt - 1, 0), "impost_hor_count": 1 if t > 0 else 0, "impost_vert_length": (l+c+r)/1000.0, "impost_hor_length": t/1000.0, "impost_total_length": (l+c+r+t)/1000.0}

def build_sashes_geometry(position):
    sashes = position.get("sashes", [])
    t_area, t_per, s_cnt = 0.0, 0.0, 0
    for s in sashes:
        w_m, h_m = safe_float(s.get("width_mm"))/1000.0, safe_float(s.get("height_mm"))/1000.0
        if w_m > 0 and h_m > 0:
            t_area += w_m * h_m
            t_per += 2*(w_m + h_m)
            s_cnt += 1
    return {"sash_count": s_cnt, "sash_area_total": t_area, "sash_perimeter_total": t_per}

def build_formula_context(position):
    pos, imp, sash = build_position_geometry(position), build_impost_geometry(position), build_sashes_geometry(position)
    p_type = normalize_text(position.get("product_type"))
    return {"count": pos["qty"], "W": pos["width_m"], "H": pos["height_m"], "area": pos["area_one"], "area_total": pos["area_total"], "perimeter": pos["perimeter_one"], "perimeter_total": pos["perimeter_total"],
            "impost_vert": imp["impost_vert_length"], "impost_hor": imp["impost_hor_length"], "impost_total": imp["impost_total_length"], "impost_vert_count": imp["impost_vert_count"], "impost_hor_count": imp["impost_hor_count"],
            "sash_count": sash["sash_count"], "sash_area_total": sash["sash_area_total"], "sash_perimeter_total": sash["sash_perimeter_total"], "is_window": 1 if "Окно" in p_type else 0, "is_door": 1 if "Дверь" in p_type else 0, "is_facade": 1 if "Фасад" in p_type else 0}

def aggregate_totals(positions):
    t_area, t_per, ts_area, ts_per = 0.0, 0.0, 0.0, 0.0
    for p in positions:
        b, s = build_position_geometry(p), build_sashes_geometry(p)
        t_area += b["area_total"]; t_per += b["perimeter_total"]; ts_area += s["sash_area_total"]; ts_per += s["sash_perimeter_total"]
    return {"total_area": t_area, "total_perimeter": t_per, "total_sash_area": ts_area, "total_sash_perimeter": ts_per}

# =========================================
# CALCULATORS
# =========================================

class MaterialCalculator:
    def __init__(self, gs_client): self.gs = gs_client
    def calculate(self, positions):
        ref_rows = self.gs.read(SHEET_REF1)
        res, t_sum = [], 0.0
        for ref in ref_rows:
            p_ref, pr_ref, formula = normalize_text(get_field(ref, "тип издел")), normalize_text(get_field(ref, "система проф")), get_field(ref, "формула_python")
            price, norm = safe_float(get_field(ref, "цена за")), safe_float(get_field(ref, "кол-во норм"), 1.0)
            if not formula or price <= 0: continue
            t_qty = 0.0
            for pos in positions:
                if (not p_ref or p_ref == normalize_text(pos.get("product_type"))) and (not pr_ref or pr_ref == normalize_text(pos.get("profile_system"))):
                    t_qty += safe_eval(formula, build_formula_context(pos))
            if t_qty <= 0: continue
            ship_qty = math.ceil(t_qty / norm) * norm if norm > 0 else t_qty
            t_sum += ship_qty * price
            res.append({"Тип изделия": p_ref, "Система профиля": pr_ref, "Тип элемента": normalize_text(get_field(ref, "тип элемент")), "Товар": str(get_field(ref, "товар")).strip(), "Факт. расход": round(t_qty, 3), "К отгрузке": ship_qty, "Цена": price, "Сумма": round(ship_qty * price, 2)})
        return res, round(t_sum, 2)

class GlassServiceCatalog:
    def __init__(self, gs_client):
        self.data = gs_client.read(SHEET_REF2)
    def get_glass_types(self):
        return sorted(list(set(str(get_field(r, "тип стеклопак")).strip() for r in self.data if get_field(r, "тип стеклопак"))))
    def get_price_by_type(self, g_type):
        for r in self.data:
            if normalize_text(get_field(r, "тип стеклопак")) == normalize_text(g_type): return safe_float(get_field(r, "стоимость"))
        return 0.0

class GlassServiceCalculator:
    def __init__(self, catalog): self.catalog = catalog
    def calculate(self, positions, g_type):
        area = aggregate_totals(positions)["total_area"]
        price = self.catalog.get_price_by_type(g_type)
        row = {"Наименование": f"Стеклопакет: {g_type}", "Цена": price, "Ед.": "м2", "Кол-во": round(area, 3), "Сумма": round(area * price, 2)}
        return [row], round(area * price, 2)

# =========================================
# UI & MAIN
# =========================================

def save_request(gs_client, **kwargs):
    # Упрощенная заглушка сохранения
    try: gs_client.append_row(SHEET_REQUESTS, [now_str(), kwargs.get("user_login"), kwargs.get("total_sum")])
    except: pass

def position_form(idx):
    st.markdown(f"### Позиция #{idx + 1}")
    c1, c2 = st.columns(2)
    pt = c1.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"], key=f"pt_{idx}")
    ps = c2.selectbox("Система профиля", ["ALG 2030-63C", "ALG 2030-55C", "ALG 2030-73C", "ALG 2030-45C", "ALG 2030-Slim", "Ruit 50F"], key=f"ps_{idx}")
    g1, g2, g3 = st.columns(3)
    w = g1.number_input("Ширина, мм", 100.0, step=10.0, key=f"w_{idx}")
    h = g2.number_input("Высота, мм", 100.0, step=10.0, key=f"h_{idx}")
    q = g3.number_input("Кол-во", 1, step=1, key=f"q_{idx}")
    st.markdown("**Импосты (мм)**")
    i1, i2, i3, i4 = st.columns(4)
    l = i1.number_input("LEFT", 0.0, key=f"l_{idx}"); c = i2.number_input("CENTER", 0.0, key=f"c_{idx}"); r = i3.number_input("RIGHT", 0.0, key=f"r_{idx}"); t = i4.number_input("TOP", 0.0, key=f"t_{idx}")
    sashes = []
    if "Окно" in pt or "Дверь" in pt:
        sc = st.number_input("Кол-во створок", 1, step=1, key=f"sc_{idx}")
        for s in range(sc):
            st.markdown(f"Створка #{s+1}")
            cols = st.columns(2)
            sw = cols[0].number_input("Ширина створки, мм", 200.0, key=f"sw_{idx}_{s}")
            sh = cols[1].number_input("Высота створки, мм", 200.0, key=f"sh_{idx}_{s}")
            sashes.append({"width_mm": sw, "height_mm": sh})
    return {"product_type": pt, "profile_system": ps, "width_mm": w, "height_mm": h, "qty": q, "left_mm": l, "center_mm": c, "right_mm": r, "top_mm": t, "sashes": sashes}

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title("🏗️ Axis Pro GF — Калькулятор")
    gs = GoogleSheetsClient(GSPREAD_SHEET_ID)
    if not login(gs): st.stop()

    if "positions_count" not in st.session_state: st.session_state.positions_count = 1
    positions = [position_form(i) for i in range(st.session_state.positions_count)]
    
    if st.button("➕ Добавить позицию"): st.session_state.positions_count += 1; st.rerun()
    
    st.header("Стеклопакет и услуги")
    catalog = GlassServiceCatalog(gs)
    g_types = catalog.get_glass_types()
    if not g_types: st.error("Стеклопакеты не найдены"); st.stop()
    sel_g = st.selectbox("Тип стеклопакета", g_types)

    if st.button("🚀 Рассчитать", type="primary"):
        m_rows, m_sum = MaterialCalculator(gs).calculate(positions)
        s_rows, s_sum = GlassServiceCalculator(catalog).calculate(positions, sel_g)
        totals = aggregate_totals(positions)
        ensure = (m_sum + s_sum) * ENSURE_PERCENT
        total = m_sum + s_sum + ensure

        st.success(f"ИТОГО: {round(total, 2)}")
        st.subheader("Сводные данные")
        c1, c2, c3 = st.columns(3)
        c1.metric("Площадь, м²", round(totals["total_area"], 3))
        c2.metric("Периметр, м", round(totals["total_perimeter"], 3))
        c3.metric("Обеспечение", round(ensure, 2))
        
        st.subheader("Материалы")
        st.dataframe(pd.DataFrame(m_rows), use_container_width=True)
        st.subheader("Услуги")
        st.dataframe(pd.DataFrame(s_rows), use_container_width=True)
        
        save_request(gs, user_login=st.session_state.user_login, total_sum=total)

if __name__ == "__main__":
    main()
