# =========================================
# Axis Pro GF v17.5 — Facade Calculator
# Полная сборка: Все части + Исправление NameError
# =========================================

import math
import ast
import operator as op
import json
import logging
import sys
import os
import base64
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
        if s == "": return default
        return float(s)
    except Exception:
        return default

def safe_int(value, default=0):
    try:
        return int(float(value))
    except Exception:
        return default

def normalize_text(value):
    if value is None: return ""
    return " ".join(str(value).strip().split())

def get_field(row, needle, default=None):
    needle = needle.lower()
    for key, value in row.items():
        if key and needle in key.lower():
            return value
    return default

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# =========================================
# SAFE AST EVAL (Логика расчета формул)
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
            args = [_eval_node(a, context) for a in node.args]
            return fn(*args)
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
# GOOGLE SHEETS CLIENT (Подключение)
# =========================================

class GoogleSheetsClient:
    @st.cache_resource
    def auth(_self):
        # Используем путь к секретному файлу gcp.json, как на вашем скриншоте Render
        secret_path = "/etc/secrets/gcp.json"
        
        if not os.path.exists(secret_path):
            st.error(f"Файл ключа не найден по пути: {secret_path}. Проверьте раздел Secret Files в Render.")
            st.stop()

        credentials = Credentials.from_service_account_file(
            secret_path,
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        return gspread.authorize(credentials)

    def __init__(self, spreadsheet_id):
        self.client = self.auth()
        self.book = self.client.open_by_key(spreadsheet_id)
        self._cache = {}

    def worksheet(self, name):
        if name not in self._cache:
            self._cache[name] = self.book.worksheet(name)
        return self._cache[name]

    @st.cache_data(ttl=1800)
    def read(_self, sheet_name):
        ws = _self.worksheet(sheet_name)
        return ws.get_all_records()

    def append_row(self, sheet_name, row):
        ws = self.worksheet(sheet_name)
        ws.append_row(row, value_input_option="USER_ENTERED")

# =========================================
# LOGIN (Авторизация)
# =========================================

def login(gs: GoogleSheetsClient):
    if "user_login" in st.session_state:
        return True

    st.sidebar.title("🔐 Вход")
    login_value = st.sidebar.text_input("Логин")
    password_value = st.sidebar.text_input("Пароль", type="password")

    if st.sidebar.button("Войти"):
        users = gs.read(SHEET_USERS)
        for user in users:
            login_cell = str(get_field(user, "логин", "")).strip()
            password_cell = str(get_field(user, "пароль", "")).strip()
            if login_cell == login_value and password_cell == password_value:
                st.session_state["user_login"] = login_cell
                st.rerun()
        st.sidebar.error("Неверный логин или пароль")
    return False

# =========================================
# GEOMETRY (Геометрия изделий)
# =========================================

def build_position_geometry(position):
    width_mm = safe_float(position.get("width_mm"))
    height_mm = safe_float(position.get("height_mm"))
    qty = safe_int(position.get("qty"), 1)
    width_m, height_m = width_mm / 1000.0, height_mm / 1000.0
    return {
        "width_mm": width_mm, "height_mm": height_mm, "width_m": width_m, "height_m": height_m,
        "area_one": width_m * height_m, "perimeter_one": 2.0 * (width_m + height_m),
        "qty": qty, "area_total": (width_m * height_m) * qty, "perimeter_total": 2.0 * (width_m + height_m) * qty,
    }

def build_impost_geometry(position):
    left = safe_float(position.get("left_mm"))
    center = safe_float(position.get("center_mm"))
    right = safe_float(position.get("right_mm"))
    top = safe_float(position.get("top_mm"))
    vertical_count = sum(1 for v in (left, center, right) if v > 0)
    return {
        "impost_vert_count": max(vertical_count - 1, 0), "impost_hor_count": 1 if top > 0 else 0,
        "impost_vert_length": (left + center + right) / 1000.0, "impost_hor_length": top / 1000.0,
        "impost_total_length": (left + center + right + top) / 1000.0,
    }

def build_sashes_geometry(position):
    sashes = position.get("sashes", [])
    total_area, total_perimeter, sash_count = 0.0, 0.0, 0
    for s in sashes:
        w_m, h_m = safe_float(s.get("width_mm")) / 1000.0, safe_float(s.get("height_mm")) / 1000.0
        if w_m > 0 and h_m > 0:
            total_area += (w_m * h_m)
            total_perimeter += 2.0 * (w_m + h_m)
            sash_count += 1
    return {"sash_count": sash_count, "sash_area_total": total_area, "sash_perimeter_total": total_perimeter}

def build_formula_context(position):
    pos = build_position_geometry(position)
    imp = build_impost_geometry(position)
    sash = build_sashes_geometry(position)
    p_type = normalize_text(position.get("product_type"))
    return {
        "count": pos["qty"], "W": pos["width_m"], "H": pos["height_m"],
        "area": pos["area_one"], "area_total": pos["area_total"],
        "perimeter": pos["perimeter_one"], "perimeter_total": pos["perimeter_total"],
        "impost_vert": imp["impost_vert_length"], "impost_hor": imp["impost_hor_length"], "impost_total": imp["impost_total_length"],
        "impost_vert_count": imp["impost_vert_count"], "impost_hor_count": imp["impost_hor_count"],
        "sash_count": sash["sash_count"], "sash_area_total": sash["sash_area_total"], "sash_perimeter_total": sash["sash_perimeter_total"],
        "is_window": 1 if "Окно" in p_type else 0, "is_door": 1 if "Дверь" in p_type else 0, "is_facade": 1 if "Фасад" in p_type else 0,
    }

def aggregate_totals(positions):
    t_area, t_per, ts_area, ts_per = 0.0, 0.0, 0.0, 0.0
    for p in positions:
        base = build_position_geometry(p)
        sash = build_sashes_geometry(p)
        t_area += base["area_total"]
        t_per += base["perimeter_total"]
        ts_area += sash["sash_area_total"]
        ts_per += sash["sash_perimeter_total"]
    return {"total_area": t_area, "total_perimeter": t_per, "total_sash_area": ts_area, "total_sash_perimeter": ts_per}

# =========================================
# CALCULATORS (Классы расчета)
# =========================================

class MaterialCalculator:
    def __init__(self, gs_client):
        self.gs = gs_client

    def calculate(self, positions):
        ref_rows = self.gs.read(SHEET_REF1)
        result_rows, total_sum = [], 0.0
        for ref in ref_rows:
            p_ref = normalize_text(get_field(ref, "тип издел"))
            pr_ref = normalize_text(get_field(ref, "система проф"))
            formula = get_field(ref, "формула_python")
            price = safe_float(get_field(ref, "цена за"))
            norm = safe_float(get_field(ref, "кол-во норм"), 1.0)
            if not formula or price <= 0: continue
            total_qty = 0.0
            for pos in positions:
                if (not p_ref or p_ref == normalize_text(pos.get("product_type"))) and (not pr_ref or pr_ref == normalize_text(pos.get("profile_system"))):
                    total_qty += safe_eval(formula, build_formula_context(pos))
            if total_qty <= 0: continue
            ship_qty = math.ceil(total_qty / norm) * norm if norm > 0 else total_qty
            total_sum += (ship_qty * price)
            result_rows.append({
                "Тип изделия": p_ref, "Система профиля": pr_ref, "Тип элемента": normalize_text(get_field(ref, "тип элемент")),
                "Товар": str(get_field(ref, "товар")).strip(), "Факт. расход": round(total_qty, 3),
                "К отгрузке": ship_qty, "Цена": price, "Сумма": round(ship_qty * price, 2)
            })
        return result_rows, round(total_sum, 2)

class GlassServiceCatalog:
    """Класс для работы со СПРАВОЧНИКОМ-2 (Стеклопакеты и Услуги)"""
    def __init__(self, gs_client):
        self.gs = gs_client
        self.data = self.gs.read(SHEET_REF2)

    def get_glass_types(self):
        types = set()
        for row in self.data:
            val = get_field(row, "тип стеклопак")
            if val: types.add(str(val).strip())
        return sorted(list(types))

    def get_price_by_type(self, glass_type):
        for row in self.data:
            if normalize_text(get_field(row, "тип стеклопак")) == normalize_text(glass_type):
                return safe_float(get_field(row, "стоимость"))
        return 0.0

class GlassServiceCalculator:
    """Класс для расчета итоговой стоимости услуг и стеклопакетов"""
    def __init__(self, catalog: GlassServiceCatalog):
        self.catalog = catalog

    def calculate(self, positions, selected_glass_type):
        totals = aggregate_totals(positions)
        total_area = totals["total_area"]
        glass_price = self.catalog.get_price_by_type(selected_glass_type)
        glass_sum = total_area * glass_price
        result_row = {
            "Наименование": f"Стеклопакет: {selected_glass_type}",
            "Цена": glass_price, "Ед.": "м2", "Кол-во": round(total_area, 3), "Сумма": round(glass_sum, 2)
        }
        return [result_row], round(glass_sum, 2)

# =========================================
# UI COMPONENTS (Формы ввода)
# =========================================

def position_form(idx):
    st.markdown(f"### Позиция #{idx + 1}")
    c1, c2 = st.columns(2)
    p_type = c1.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"], key=f"pt_{idx}")
    p_sys = c2.selectbox("Система профиля", ["ALG 2030-63C", "ALG 2030-55C", "ALG 2030-73C", "ALG 2030-45C", "ALG 2030-Slim", "Ruit 50F"], key=f"ps_{idx}")
    g1, g2, g3 = st.columns(3)
    w = g1.number_input("Ширина, мм", 100.0, step=10.0, key=f"w_{idx}")
    h = g2.number_input("Высота, мм", 100.0, step=10.0, key=f"h_{idx}")
    q = g3.number_input("Кол-во", 1, step=1, key=f"q_{idx}")
    st.markdown("**Импосты (мм)**")
    i1, i2, i3, i4 = st.columns(4)
    l = i1.number_input("LEFT", 0.0, key=f"l_{idx}"); c = i2.number_input("CENTER", 0.0, key=f"c_{idx}")
    r = i3.number_input("RIGHT", 0.0, key=f"r_{idx}"); t = i4.number_input("TOP", 0.0, key=f"t_{idx}")
    sashes = []
    if "Окно" in p_type or "Дверь" in p_type:
        sc = st.number_input("Кол-во створок", 1, step=1, key=f"sc_{idx}")
        for s in range(sc):
            st.markdown(f"Створка #{s+1}")
            cols = st.columns(2)
            sw = cols[0].number_input("Ширина створки, мм", 200.0, key=f"sw_{idx}_{s}")
            sh = cols[1].number_input("Высота створки, мм", 200.0, key=f"sh_{idx}_{s}")
            sashes.append({"width_mm": sw, "height_mm": sh})
    return {"product_type": p_type, "profile_system": p_sys, "width_mm": w, "height_mm": h, "qty": q, "left_mm": l, "center_mm": c, "right_mm": r, "top_mm": t, "sashes": sashes}

# =========================================
# MAIN (Главная функция)
# =========================================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title("🏗️ Axis Pro GF — Калькулятор")
    
    gs = GoogleSheetsClient(GSPREAD_SHEET_ID)
    if not login(gs): st.stop()

    st.header("Параметры изделий")
    if "positions_count" not in st.session_state: st.session_state.positions_count = 1
    positions = []
    for i in range(st.session_state.positions_count):
        positions.append(position_form(i))
        st.divider()

    if st.button("➕ Добавить позицию"):
        st.session_state.positions_count += 1
        st.rerun()

    st.header("Стеклопакет и услуги")
    catalog = GlassServiceCatalog(gs)
    glass_types = catalog.get_glass_types()
    if not glass_types:
        st.error("В СПРАВОЧНИКЕ-2 не найдены типы стеклопакетов.")
        st.stop()
    selected_glass = st.selectbox("Тип стеклопакета", glass_types)

    if st.button("🚀 Рассчитать", type="primary"):
        with st.spinner("Выполняется расчёт..."):
            m_calc = MaterialCalculator(gs)
            m_rows, m_sum = m_calc.calculate(positions)
            s_calc = GlassServiceCalculator(catalog)
            s_rows, s_sum = s_calc.calculate(positions, selected_glass)
            totals = aggregate_totals(positions)
            ensure_sum = (m_sum + s_sum) * ENSURE_PERCENT
            total_pay = m_sum + s_sum + ensure_sum

        st.success(f"ИТОГО к оплате: {round(total_pay, 2)}")
        st.subheader("Сводные данные")
        c1, c2, c3 = st.columns(3)
        c1.metric("Общая площадь, м²", round(totals["total_area"], 3))
        c2.metric("Общий периметр, м", round(totals["total_perimeter"], 3))
        c3.metric("Обеспечение 65%", round(ensure_sum, 2))

        t1, t2 = st.tabs(["Материалы", "Итог"])
        with t1:
            st.dataframe(pd.DataFrame(m_rows), use_container_width=True)
            st.write(f"**Итого материалы:** {round(m_sum, 2)}")
        with t2:
            st.dataframe(pd.DataFrame(s_rows), use_container_width=True)
            st.write(f"**Итого услуги:** {round(s_sum, 2)}")
            st.divider()
            st.write(f"**ОБЩИЙ ИТОГ С НАЦЕНКОЙ:** {round(total_pay, 2)}")

if __name__ == "__main__":
    main()
