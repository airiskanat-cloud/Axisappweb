# =========================================
# Axis Pro GF — Calculator
# Part 1 / 6
# Base, utils, GoogleSheetsClient, login
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

# =========================================
# LOGGER
# =========================================

logger = logging.getLogger("axis_pro_gf")

if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)

logger.setLevel(logging.INFO)

# =========================================
# UTILS
# =========================================

def safe_float(value, default=0.0):
    try:
        if value is None:
            return default
        s = (
            str(value)
            .replace("\xa0", "")
            .replace(" ", "")
            .replace(",", ".")
        )
        if s == "":
            return default
        return float(s)
    except Exception:
        return default


def safe_int(value, default=0):
    try:
        return int(float(value))
    except Exception:
        return default


def normalize_text(value):
    if value is None:
        return ""
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
# SAFE AST EVAL (for formulas from sheets)
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


def _eval_node(node, context):
    if isinstance(node, ast.Expression):
        return _eval_node(node.body, context)

    if isinstance(node, ast.Constant):
        return node.value

    if isinstance(node, ast.Name):
        if node.id in context:
            return context[node.id]
        raise ValueError(f"Unknown variable: {node.id}")

    if isinstance(node, ast.BinOp):
        return _ALLOWED_OPS[type(node.op)](
            _eval_node(node.left, context),
            _eval_node(node.right, context),
        )

    if isinstance(node, ast.UnaryOp):
        return _ALLOWED_OPS[type(node.op)](
            _eval_node(node.operand, context)
        )

    if isinstance(node, ast.Call):
        if isinstance(node.func, ast.Name):
            if node.func.id in ("min", "max"):
                args = [_eval_node(a, context) for a in node.args]
                return globals()[node.func.id](*args)

        if isinstance(node.func, ast.Attribute):
            if node.func.value.id == "math":
                fn = getattr(math, node.func.attr)
                args = [_eval_node(a, context) for a in node.args]
                return fn(*args)

    raise ValueError("Unsafe expression")


def safe_eval(formula, context):
    if not formula:
        return 0.0
    try:
        prepared = {
            k: safe_float(v)
            for k, v in context.items()
        }
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
        secret_path = "/etc/secrets/gcp_service_account.json"

        if not os.path.exists(secret_path):
            st.error("Service account file not found")
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
# LOGIN
# =========================================

def login(gs: GoogleSheetsClient):
    if "user_login" in st.session_state:
        return True

    st.sidebar.title("Вход")

    login_value = st.sidebar.text_input("Логин")
    password_value = st.sidebar.text_input(
        "Пароль",
        type="password",
    )

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
# GEOMETRY & CONTEXT
# Part 2 / 6
# =========================================

def build_position_geometry(position):
    """
    Base geometry for one position (one product line).
    Quantity (qty) applies to the whole position.
    """

    width_mm = safe_float(position.get("width_mm"))
    height_mm = safe_float(position.get("height_mm"))
    qty = safe_int(position.get("qty"), 1)

    width_m = width_mm / 1000.0
    height_m = height_mm / 1000.0

    area_one = width_m * height_m
    perimeter_one = 2.0 * (width_m + height_m)

    return {
        "width_mm": width_mm,
        "height_mm": height_mm,
        "width_m": width_m,
        "height_m": height_m,
        "area_one": area_one,
        "perimeter_one": perimeter_one,
        "qty": qty,
        "area_total": area_one * qty,
        "perimeter_total": perimeter_one * qty,
    }


def build_impost_geometry(position):
    """
    Imposts logic (Excel-compatible):
    LEFT / CENTER / RIGHT / TOP are lengths in mm.
    """

    left = safe_float(position.get("left_mm"))
    center = safe_float(position.get("center_mm"))
    right = safe_float(position.get("right_mm"))
    top = safe_float(position.get("top_mm"))

    vertical_count = sum(1 for v in (left, center, right) if v > 0)
    horizontal_count = 1 if top > 0 else 0

    vertical_length_m = (left + center + right) / 1000.0
    horizontal_length_m = top / 1000.0

    return {
        "impost_vert_count": max(vertical_count - 1, 0),
        "impost_hor_count": horizontal_count,
        "impost_vert_length": vertical_length_m,
        "impost_hor_length": horizontal_length_m,
        "impost_total_length": vertical_length_m + horizontal_length_m,
    }


def build_sashes_geometry(position):
    """
    Multi-sash geometry.
    Supports multiple sashes with DIFFERENT sizes.

    Expected formats:
    - position["sashes"] = list of dicts:
        [{"width_mm": ..., "height_mm": ...}, ...]
    OR (legacy fallback):
    - sash_count + sash_width_mm + sash_height_mm
    """

    sashes = position.get("sashes")

    total_area = 0.0
    total_perimeter = 0.0
    sash_count = 0

    if isinstance(sashes, list) and len(sashes) > 0:
        for sash in sashes:
            w_mm = safe_float(sash.get("width_mm"))
            h_mm = safe_float(sash.get("height_mm"))

            if w_mm <= 0 or h_mm <= 0:
                continue

            w_m = w_mm / 1000.0
            h_m = h_mm / 1000.0

            total_area += w_m * h_m
            total_perimeter += 2.0 * (w_m + h_m)
            sash_count += 1

    else:
        # --- Legacy single-size fallback ---
        sash_count = safe_int(position.get("sash_count"))
        sash_w_mm = safe_float(position.get("sash_width_mm"))
        sash_h_mm = safe_float(position.get("sash_height_mm"))

        if sash_count > 0 and sash_w_mm > 0 and sash_h_mm > 0:
            w_m = sash_w_mm / 1000.0
            h_m = sash_h_mm / 1000.0

            area_one = w_m * h_m
            per_one = 2.0 * (w_m + h_m)

            total_area = area_one * sash_count
            total_perimeter = per_one * sash_count

    return {
        "sash_count": sash_count,
        "sash_area_total": total_area,
        "sash_perimeter_total": total_perimeter,
    }


def build_formula_context(position):
    """
    Formula context for СПРАВОЧНИК-1 (Excel-compatible).
    """

    pos = build_position_geometry(position)
    impost = build_impost_geometry(position)
    sash = build_sashes_geometry(position)

    ctx = {}

    # --- BASIC ---
    ctx["count"] = pos["qty"]
    ctx["W"] = pos["width_m"]
    ctx["H"] = pos["height_m"]

    # --- AREAS ---
    ctx["area"] = pos["area_one"]
    ctx["area_total"] = pos["area_total"]

    # --- PERIMETERS ---
    ctx["perimeter"] = pos["perimeter_one"]
    ctx["perimeter_total"] = pos["perimeter_total"]

    # --- IMPOSTS ---
    ctx["impost_vert"] = impost["impost_vert_length"]
    ctx["impost_hor"] = impost["impost_hor_length"]
    ctx["impost_total"] = impost["impost_total_length"]
    ctx["impost_vert_count"] = impost["impost_vert_count"]
    ctx["impost_hor_count"] = impost["impost_hor_count"]

    # --- SASHES ---
    ctx["sash_count"] = sash["sash_count"]
    ctx["sash_area_total"] = sash["sash_area_total"]
    ctx["sash_perimeter_total"] = sash["sash_perimeter_total"]

    # --- FLAGS ---
    product_type = normalize_text(position.get("product_type"))
    ctx["is_window"] = 1 if "Окно" in product_type else 0
    ctx["is_door"] = 1 if "Дверь" in product_type else 0
    ctx["is_facade"] = 1 if "Фасад" in product_type else 0

    return ctx


def aggregate_totals(positions):
    """
    Aggregates totals across ALL positions.
    Correct for any quantity and any number of sashes.
    """

    total_area = 0.0
    total_perimeter = 0.0
    total_sash_area = 0.0
    total_sash_perimeter = 0.0

    for pos in positions:
        base = build_position_geometry(pos)
        sash = build_sashes_geometry(pos)

        total_area += base["area_total"]
        total_perimeter += base["perimeter_total"]
        total_sash_area += sash["sash_area_total"]
        total_sash_perimeter += sash["sash_perimeter_total"]

    return {
        "total_area": total_area,
        "total_perimeter": total_perimeter,
        "total_sash_area": total_sash_area,
        "total_sash_perimeter": total_sash_perimeter,
    }
# =========================================
# MATERIAL CALCULATOR
# Part 3 / 6
# Based on СПРАВОЧНИК-1
# =========================================

class MaterialCalculator:
    """
    Material calculation strictly follows Excel logic from СПРАВОЧНИК-1.

    Business rules (FIXED):
    - If 'Тип изделия' is filled in reference → must match position product_type
    - If 'Тип изделия' is empty → applies to ALL products
    - If 'Система профиля' is filled in reference → must match position profile_system
    - If 'Система профиля' is empty → applies to ALL profile systems

    Result rows count = actual number of reference rows
    matching selected product type and profile system.
    """

    def __init__(self, gs_client):
        self.gs = gs_client

    def calculate(self, positions):
        ref_rows = self.gs.read(SHEET_REF1)

        result_rows = []
        total_sum = 0.0

        for ref in ref_rows:
            product_type_ref = normalize_text(
                get_field(ref, "тип издел", "")
            )
            profile_ref = normalize_text(
                get_field(ref, "система проф", "")
            )
            element_type = normalize_text(
                get_field(ref, "тип элемент", "")
            )
            product_name = str(
                get_field(ref, "товар", "")
            ).strip()

            formula = get_field(ref, "формула_python")
            price = safe_float(get_field(ref, "цена за"), 0.0)
            norm = safe_float(get_field(ref, "кол-во норм"), 1.0)

            # --- BASIC VALIDATION ---
            if not formula or price <= 0:
                continue

            total_qty = 0.0

            for pos in positions:
                pos_product = normalize_text(pos.get("product_type"))
                pos_profile = normalize_text(pos.get("profile_system"))

                # --- PRODUCT TYPE FILTER ---
                if product_type_ref:
                    if product_type_ref != pos_product:
                        continue

                # --- PROFILE SYSTEM FILTER ---
                if profile_ref:
                    if profile_ref != pos_profile:
                        continue

                # --- FORMULA CONTEXT ---
                ctx = build_formula_context(pos)

                try:
                    qty_value = safe_eval(formula, ctx)
                except Exception:
                    qty_value = 0.0

                total_qty += qty_value

            if total_qty <= 0:
                continue

            # --- NORMALIZATION (bundles, bars, packages) ---
            if norm > 0:
                ship_qty = math.ceil(total_qty / norm) * norm
            else:
                ship_qty = total_qty

            row_sum = ship_qty * price
            total_sum += row_sum

            result_rows.append({
                "Тип изделия": product_type_ref,
                "Система профиля": profile_ref,
                "Тип элемента": element_type,
                "Товар": product_name,
                "Факт. расход": round(total_qty, 3),
                "К отгрузке": ship_qty,
                "Цена": price,
                "Сумма": round(row_sum, 2),
            })

        return result_rows, round(total_sum, 2)
# =========================================
# UI — POSITIONS INPUT
# Part 5 / 6
# =========================================

def position_form(idx, is_facade=False):
    """
    One position form.
    idx — index in positions list
    """

    st.markdown(f"### Позиция #{idx + 1}")

    c1, c2 = st.columns(2)

    with c1:
        product_type = st.selectbox(
            "Тип изделия",
            [
                "Окно с откр.",
                "Окно глух.",
                "Дверь 1 створч.",
                "Дверь 2-х створч.",
                "Фасад",
            ],
            key=f"ptype_{idx}",
        )

    with c2:
        profile_system = st.selectbox(
            "Система профиля",
            [
                "ALG 2030-63C",
                "ALG 2030-55C",
                "ALG 2030-73C",
                "ALG 2030-45C",
                "ALG 2030-Slim",
                "Ruit 50F",
            ],
            key=f"psys_{idx}",
        )

    g1, g2, g3 = st.columns(3)

    with g1:
        width_mm = st.number_input(
            "Ширина, мм",
            min_value=100.0,
            step=10.0,
            key=f"w_{idx}",
        )

    with g2:
        height_mm = st.number_input(
            "Высота, мм",
            min_value=100.0,
            step=10.0,
            key=f"h_{idx}",
        )

    with g3:
        qty = st.number_input(
            "Кол-во (N)",
            min_value=1,
            step=1,
            value=1,
            key=f"q_{idx}",
        )

    # --- IMPOSTS ---
    st.markdown("**Импосты (мм)**")
    i1, i2, i3, i4 = st.columns(4)

    with i1:
        left_mm = st.number_input("LEFT", min_value=0.0, step=10.0, key=f"l_{idx}")
    with i2:
        center_mm = st.number_input("CENTER", min_value=0.0, step=10.0, key=f"c_{idx}")
    with i3:
        right_mm = st.number_input("RIGHT", min_value=0.0, step=10.0, key=f"r_{idx}")
    with i4:
        top_mm = st.number_input("TOP", min_value=0.0, step=10.0, key=f"t_{idx}")

    # --- SASHES ---
    sashes = []
    sash_count = 0

    if "Окно" in product_type or "Дверь" in product_type:
        st.markdown("**Створки**")

        sash_count = st.number_input(
            "Кол-во створок",
            min_value=1,
            step=1,
            value=1,
            key=f"sash_count_{idx}",
        )

        for s in range(sash_count):
            st.markdown(f"Створка #{s + 1}")
            sc1, sc2 = st.columns(2)

            with sc1:
                sw = st.number_input(
                    "Ширина створки, мм",
                    min_value=200.0,
                    step=10.0,
                    key=f"sw_{idx}_{s}",
                )

            with sc2:
                sh = st.number_input(
                    "Высота створки, мм",
                    min_value=200.0,
                    step=10.0,
                    key=f"sh_{idx}_{s}",
                )

            sashes.append({
                "width_mm": sw,
                "height_mm": sh,
            })

    # --- FACADE ---
    stand_step_mm = 0.0
    if product_type == "Фасад":
        stand_step_mm = st.number_input(
            "Шаг стоек фасада, мм",
            min_value=300.0,
            step=50.0,
            value=1000.0,
            key=f"stand_{idx}",
        )

    return {
        "product_type": product_type,
        "profile_system": profile_system,
        "width_mm": width_mm,
        "height_mm": height_mm,
        "qty": qty,
        "left_mm": left_mm,
        "center_mm": center_mm,
        "right_mm": right_mm,
        "top_mm": top_mm,
        "sashes": sashes,
        "sash_count": sash_count,
        "stand_step_mm": stand_step_mm,
    }


def positions_block():
    """
    Block for multiple positions.
    """

    if "positions_count" not in st.session_state:
        st.session_state["positions_count"] = 1

    positions = []

    for i in range(st.session_state["positions_count"]):
        positions.append(position_form(i))
        st.divider()

    c1, c2 = st.columns(2)

    with c1:
        if st.button("➕ Добавить позицию"):
            st.session_state["positions_count"] += 1
            st.rerun()

    with c2:
        if st.session_state["positions_count"] > 1:
            if st.button("➖ Удалить последнюю"):
                st.session_state["positions_count"] -= 1
                st.rerun()

    return positions
# =========================================
# MAIN APPLICATION
# Part 6 / 6
# =========================================

def main():
    st.set_page_config(
        page_title=APP_TITLE,
        layout="wide",
    )

    st.title("🏗️ Axis Pro GF — Калькулятор")

    # --- GOOGLE SHEETS ---
    gs = GoogleSheetsClient(GSPREAD_SHEET_ID)

    # --- LOGIN ---
    if not login(gs):
        st.stop()

    # --- INPUT BLOCK ---
    st.header("Параметры изделий")

    positions = positions_block()

    # --- GLASS TYPE ---
    st.header("Стеклопакет и услуги")

    catalog = GlassServiceCatalog(gs)
    glass_types = catalog.get_glass_types()

    if not glass_types:
        st.error("В СПРАВОЧНИК-2 не найдены типы стеклопакетов")
        st.stop()

    selected_glass_type = st.selectbox(
        "Тип стеклопакета",
        glass_types,
    )

    # --- CALCULATE ---
    if st.button("🚀 Рассчитать", type="primary"):
        with st.spinner("Выполняется расчёт..."):

            # --- MATERIALS ---
            material_calc = MaterialCalculator(gs)
            material_rows, material_sum = material_calc.calculate(positions)

            # --- SERVICES ---
            service_calc = GlassServiceCalculator(catalog)
            service_rows, services_sum = service_calc.calculate(
                positions,
                selected_glass_type,
            )

            # --- TOTALS ---
            totals = aggregate_totals(positions)

            base_sum = material_sum + services_sum
            ensure_sum = base_sum * ENSURE_PERCENT
            total_sum = base_sum + ensure_sum

        # --- RESULT HEADER ---
        st.success(f"ИТОГО к оплате: {round(total_sum, 2)}")

        # --- SUMMARY ---
        st.subheader("Сводные данные")

        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric("Общая площадь, м²", round(totals["total_area"], 3))
        with c2:
            st.metric("Общий периметр, м", round(totals["total_perimeter"], 3))
        with c3:
            st.metric("Обеспечение 65%", round(ensure_sum, 2))

        # --- MATERIALS TABLE ---
        st.subheader("Материалы (СПРАВОЧНИК-1)")

        if material_rows:
            st.dataframe(
                pd.DataFrame(material_rows),
                use_container_width=True,
            )
            st.write(f"**Итого материалы:** {round(material_sum, 2)}")
        else:
            st.info("Материалы не рассчитаны")

        # --- SERVICES TABLE ---
        st.subheader("Стеклопакет и услуги (СПРАВОЧНИК-2)")

        if service_rows:
            st.dataframe(
                pd.DataFrame(service_rows),
                use_container_width=True,
            )
            st.write(f"**Итого услуги:** {round(services_sum, 2)}")
        else:
            st.info("Услуги отсутствуют")

        # --- FINAL TOTAL ---
        st.subheader("Итог")

        итог_df = pd.DataFrame(
            [
                ["Материалы", material_sum],
                ["Услуги", services_sum],
                ["Обеспечение 65%", ensure_sum],
                ["ИТОГО", total_sum],
            ],
            columns=["Позиция", "Сумма"],
        )

        st.dataframe(итог_df, use_container_width=True)

        # --- SAVE REQUEST ---
        save_request(
            gs_client=gs,
            user_login=st.session_state["user_login"],
            positions=positions,
            glass_type=selected_glass_type,
            material_sum=material_sum,
            services_sum=services_sum,
            total_area=totals["total_area"],
            total_perimeter=totals["total_perimeter"],
            ensure_sum=ensure_sum,
            total_sum=total_sum,
            material_rows=material_rows,
        )

        # --- COMMERCIAL PROPOSAL ---
        st.subheader("Коммерческое предложение")

        if os.path.exists(KP_TEMPLATE_PATH):
            kp_path = generate_kp_file(
                user_login=st.session_state["user_login"],
                glass_type=selected_glass_type,
                totals=totals,
                material_rows=material_rows,
                material_sum=material_sum,
                services_sum=services_sum,
                ensure_sum=ensure_sum,
                total_sum=total_sum,
            )

            with open(kp_path, "rb") as f:
                st.download_button(
                    label="📄 Скачать коммерческое предложение (Excel)",
                    data=f,
                    file_name="Коммерческое_предложение.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.warning("Файл шаблона КП не найден")


if __name__ == "__main__":
    main()
