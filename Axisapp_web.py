# =========================================
# Axis Pro GF — Calculator
# Base Part (Part 1 / 7)
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

APP_TITLE = "Axis Pro GF — Facade / Windows / Doors"

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
    def auth(self):
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
    def read(self, sheet_name):
        ws = self.worksheet(sheet_name)
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

    st.sidebar.title("Login")

    login_value = st.sidebar.text_input("Login")
    password_value = st.sidebar.text_input(
        "Password",
        type="password",
    )

    if st.sidebar.button("Sign in"):
        users = gs.read(SHEET_USERS)
        for user in users:
            login_cell = str(get_field(user, "логин", "")).strip()
            password_cell = str(get_field(user, "пароль", "")).strip()

            if login_cell == login_value and password_cell == password_value:
                st.session_state["user_login"] = login_cell
                st.rerun()

        st.sidebar.error("Invalid login or password")

    return False
# =========================================
# GEOMETRY AND CONTEXT (Part 2 / 7)
# =========================================

def build_position_geometry(position):
    """
    One position = one row like in Excel 'ЗАПРОСЫ'
    """

    width_mm = safe_float(position.get("width_mm"))
    height_mm = safe_float(position.get("height_mm"))
    qty = safe_int(position.get("qty"), 1)

    width_m = width_mm / 1000.0
    height_m = height_mm / 1000.0

    area_one = width_m * height_m
    perimeter_one = 2.0 * (width_m + height_m)

    total_area = area_one * qty
    total_perimeter = perimeter_one * qty

    return {
        "width_mm": width_mm,
        "height_mm": height_mm,
        "width_m": width_m,
        "height_m": height_m,
        "area_one": area_one,
        "perimeter_one": perimeter_one,
        "qty": qty,
        "total_area": total_area,
        "total_perimeter": total_perimeter,
    }


def build_impost_geometry(position):
    """
    Imposts exactly as in Excel:
    LEFT / CENTER / RIGHT / TOP are lengths in mm
    """

    left = safe_float(position.get("left_mm"))
    center = safe_float(position.get("center_mm"))
    right = safe_float(position.get("right_mm"))
    top = safe_float(position.get("top_mm"))

    vertical_count = sum(1 for v in (left, center, right) if v > 0)
    horizontal_count = 1 if top > 0 else 0

    vertical_length = (left + center + right) / 1000.0
    horizontal_length = top / 1000.0

    return {
        "impost_vert_count": max(vertical_count - 1, 0),
        "impost_hor_count": horizontal_count,
        "impost_vert_length": vertical_length,
        "impost_hor_length": horizontal_length,
        "impost_total_length": vertical_length + horizontal_length,
    }


def build_sash_geometry(position):
    """
    Sashes logic exactly as Excel:
    each sash has its own width / height
    """

    sash_count = safe_int(position.get("sash_count"))
    sash_width_mm = safe_float(position.get("sash_width_mm"))
    sash_height_mm = safe_float(position.get("sash_height_mm"))

    if sash_count <= 0:
        return {
            "sash_count": 0,
            "sash_area_one": 0.0,
            "sash_area_total": 0.0,
            "sash_perimeter_one": 0.0,
            "sash_perimeter_total": 0.0,
        }

    sash_width_m = sash_width_mm / 1000.0
    sash_height_m = sash_height_mm / 1000.0

    area_one = sash_width_m * sash_height_m
    perimeter_one = 2.0 * (sash_width_m + sash_height_m)

    return {
        "sash_count": sash_count,
        "sash_width_mm": sash_width_mm,
        "sash_height_mm": sash_height_mm,
        "sash_width_m": sash_width_m,
        "sash_height_m": sash_height_m,
        "sash_area_one": area_one,
        "sash_area_total": area_one * sash_count,
        "sash_perimeter_one": perimeter_one,
        "sash_perimeter_total": perimeter_one * sash_count,
    }


def build_formula_context(position):
    """
    CONTEXT FOR СПРАВОЧНИК-1 FORMULAS
    Variables match Excel exactly
    """

    pos = build_position_geometry(position)
    impost = build_impost_geometry(position)
    sash = build_sash_geometry(position)

    context = {}

    # --- BASIC ---
    context["count"] = pos["qty"]
    context["W"] = pos["width_m"]
    context["H"] = pos["height_m"]

    # --- AREAS ---
    context["area"] = pos["area_one"]
    context["area_total"] = pos["total_area"]

    # --- PERIMETERS ---
    context["perimeter"] = pos["perimeter_one"]
    context["perimeter_total"] = pos["total_perimeter"]

    # --- IMPOSTS ---
    context["impost_vert"] = impost["impost_vert_length"]
    context["impost_hor"] = impost["impost_hor_length"]
    context["impost_total"] = impost["impost_total_length"]
    context["impost_vert_count"] = impost["impost_vert_count"]
    context["impost_hor_count"] = impost["impost_hor_count"]

    # --- SASHES ---
    context["sash_count"] = sash["sash_count"]
    context["glass_w"] = sash.get("sash_width_m", 0.0)
    context["glass_h"] = sash.get("sash_height_m", 0.0)
    context["sash_area"] = sash.get("sash_area_one", 0.0)
    context["sash_area_total"] = sash.get("sash_area_total", 0.0)
    context["sash_perimeter"] = sash.get("sash_perimeter_one", 0.0)
    context["sash_perimeter_total"] = sash.get("sash_perimeter_total", 0.0)

    # --- FLAGS ---
    product_type = normalize_text(position.get("product_type"))
    context["is_window"] = 1 if "Окно" in product_type else 0
    context["is_door"] = 1 if "Дверь" in product_type else 0
    context["is_facade"] = 1 if "Фасад" in product_type else 0

    return context


def aggregate_totals(positions):
    """
    Global totals like Excel 'ИТОГО'
    """

    total_area = 0.0
    total_perimeter = 0.0
    total_sash_area = 0.0
    total_sash_perimeter = 0.0

    for pos in positions:
        geo = build_position_geometry(pos)
        sash = build_sash_geometry(pos)

        total_area += geo["total_area"]
        total_perimeter += geo["total_perimeter"]
        total_sash_area += sash.get("sash_area_total", 0.0)
        total_sash_perimeter += sash.get("sash_perimeter_total", 0.0)

    return {
        "total_area": total_area,
        "total_perimeter": total_perimeter,
        "total_sash_area": total_sash_area,
        "total_sash_perimeter": total_sash_perimeter,
    }
# =========================================
# MATERIAL CALCULATOR (Part 3 / 7)
# Based on СПРАВОЧНИК-1
# =========================================

class MaterialCalculator:
    """
    Material calculation strictly follows Excel logic from СПРАВОЧНИК-1.

    Key principles:
    - Exact match of product_type and profile_system (string == string)
    - Formula context = build_formula_context()
    - Supports different element types:
        * per position
        * per sash
        * per square meter
        * per running meter
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

            if not formula or price <= 0:
                continue

            total_qty = 0.0

            for pos in positions:
                # --- STRICT MATCH WITH СПРАВОЧНИК-1 ---
                if product_type_ref and product_type_ref != normalize_text(pos.get("product_type")):
                    continue
                if profile_ref and profile_ref != normalize_text(pos.get("profile_system")):
                    continue

                ctx = build_formula_context(pos)

                try:
                    qty_value = safe_eval(formula, ctx)
                except Exception:
                    qty_value = 0.0

                # --- ELEMENT TYPE LOGIC ---
                # Excel behavior:
                # formula already returns correct quantity unit
                total_qty += qty_value

            if total_qty <= 0:
                continue

            # --- NORMALIZATION (хлысты, упаковки и т.п.) ---
            if norm > 0:
                ship_qty = math.ceil(total_qty / norm) * norm
            else:
                ship_qty = total_qty

            row_sum = ship_qty * price
            total_sum += row_sum

            result_rows.append({
                "Product type": product_type_ref,
                "Profile system": profile_ref,
                "Element type": element_type,
                "Item": product_name,
                "Calculated qty": round(total_qty, 3),
                "Shipping qty": ship_qty,
                "Price": price,
                "Sum": round(row_sum, 2),
            })

        return result_rows, round(total_sum, 2)
# =========================================
# GLASS AND SERVICES (Part 4 / 7)
# Based on СПРАВОЧНИК-2
# =========================================

class GlassServiceCatalog:
    """
    Reads СПРАВОЧНИК-2 and provides:
    - list of glass types for UI
    - prices for glass, panels, toning, assembly, montage
    """

    def __init__(self, gs_client):
        self.gs = gs_client
        self.rows = self.gs.read(SHEET_REF2)

    def get_glass_types(self):
        """
        Returns list of glass types exactly as in СПРАВОЧНИК-2
        """
        types = []
        for row in self.rows:
            glass_type = normalize_text(
                get_field(row, "тип стеклопак", "")
            )
            if glass_type and glass_type not in types:
                types.append(glass_type)
        return types

    def find_row_by_glass(self, glass_type):
        glass_type_norm = normalize_text(glass_type)
        for row in self.rows:
            if normalize_text(get_field(row, "тип стеклопак", "")) == glass_type_norm:
                return row
        return None

    def get_price(self, row, field_name):
        return safe_float(get_field(row, field_name), 0.0)


class GlassServiceCalculator:
    """
    Calculates glass and services strictly from СПРАВОЧНИК-2
    """

    def __init__(self, catalog: GlassServiceCatalog):
        self.catalog = catalog

    def calculate(self, positions, selected_glass_type):
        totals = aggregate_totals(positions)

        total_area = totals["total_area"]

        result_rows = []
        total_sum = 0.0

        ref_row = self.catalog.find_row_by_glass(selected_glass_type)
        if not ref_row:
            return result_rows, 0.0

        # --- GLASS ---
        glass_price = self.catalog.get_price(
            ref_row,
            "стоимость стеклопакета",
        )
        if glass_price > 0:
            glass_sum = total_area * glass_price
            result_rows.append((
                "Glass unit",
                glass_price,
                "m2",
                round(glass_sum, 2),
            ))
            total_sum += glass_sum

        # --- PANELS ---
        panel_price = self.catalog.get_price(
            ref_row,
            "стоимость панел",
        )
        if panel_price > 0:
            panel_sum = total_area * panel_price
            result_rows.append((
                "Panels",
                panel_price,
                "m2",
                round(panel_sum, 2),
            ))
            total_sum += panel_sum

        # --- TONING ---
        toning_price = self.catalog.get_price(
            ref_row,
            "стоимость тониров",
        )
        if toning_price > 0:
            toning_sum = total_area * toning_price
            result_rows.append((
                "Toning",
                toning_price,
                "m2",
                round(toning_sum, 2),
            ))
            total_sum += toning_sum

        # --- ASSEMBLY ---
        assembly_price = self.catalog.get_price(
            ref_row,
            "стоимость сборк",
        )
        if assembly_price > 0:
            assembly_sum = total_area * assembly_price
            result_rows.append((
                "Assembly",
                assembly_price,
                "m2",
                round(assembly_sum, 2),
            ))
            total_sum += assembly_sum

        # --- MONTAGE ---
        montage_price = self.catalog.get_price(
            ref_row,
            "стоимость монтаж",
        )
        if montage_price > 0:
            montage_sum = total_area * montage_price
            result_rows.append((
                "Montage",
                montage_price,
                "m2",
                round(montage_sum, 2),
            ))
            total_sum += montage_sum

        return result_rows, round(total_sum, 2)
# =========================================
# FACADE ENGINEERING (Part 5 / 7)
# Wind load logic based on Excel concept
# =========================================

# Stand profiles table (from Excel)
# Ordered by Jx ascending
FACADE_STAND_TABLE = [
    {"code": "90-5035",  "jx": 79},
    {"code": "100-5009", "jx": 117},
    {"code": "110-5034", "jx": 126},
    {"code": "130-5033", "jx": 190},
    {"code": "150-5032", "jx": 277},
    {"code": "170-5010", "jx": 403},
    {"code": "160-5005", "jx": 422},
    {"code": "200-5006", "jx": 851},
]


def facade_required_jx(height_m):
    """
    Required Jx based on facade height.
    Formula follows Excel concept:
    Jx ~ H^2
    """
    if height_m <= 0:
        return 0.0
    return 55.0 * (height_m ** 2)


def select_facade_stand(height_mm):
    """
    Selects facade stand profile based on height.
    Returns dict with code and jx.
    """
    height_m = height_mm / 1000.0
    required_jx = facade_required_jx(height_m)

    for stand in FACADE_STAND_TABLE:
        if stand["jx"] >= required_jx:
            return stand

    return None


def build_facade_context(position):
    """
    Adds facade-specific parameters to formula context.
    Does NOT override base context.
    """

    context = build_formula_context(position)

    height_mm = safe_float(position.get("height_mm"))
    stand_step_mm = safe_float(position.get("stand_step_mm"))

    stand = select_facade_stand(height_mm)

    context["facade_height_m"] = height_mm / 1000.0
    context["stand_step_m"] = stand_step_mm / 1000.0

    if stand:
        context["facade_stand_jx"] = stand["jx"]
        context["facade_stand_code"] = stand["code"]
    else:
        context["facade_stand_jx"] = 0.0
        context["facade_stand_code"] = ""

    return context


def is_facade_position(position):
    product_type = normalize_text(position.get("product_type"))
    return product_type == "Фасад"
# =========================================
# REQUESTS STORAGE (Part 6 / 7)
# Save calculation history to SHEET_REQUESTS
# =========================================

def serialize_positions(positions):
    """
    Prepare positions data for storage (JSON).
    """
    clean = []
    for pos in positions:
        clean.append({
            "product_type": pos.get("product_type"),
            "profile_system": pos.get("profile_system"),
            "width_mm": safe_float(pos.get("width_mm")),
            "height_mm": safe_float(pos.get("height_mm")),
            "qty": safe_int(pos.get("qty"), 1),
            "left_mm": safe_float(pos.get("left_mm")),
            "center_mm": safe_float(pos.get("center_mm")),
            "right_mm": safe_float(pos.get("right_mm")),
            "top_mm": safe_float(pos.get("top_mm")),
            "sash_count": safe_int(pos.get("sash_count")),
            "sash_width_mm": safe_float(pos.get("sash_width_mm")),
            "sash_height_mm": safe_float(pos.get("sash_height_mm")),
            "stand_step_mm": safe_float(pos.get("stand_step_mm")),
        })
    return json.dumps(clean, ensure_ascii=False)


def serialize_materials(material_rows):
    """
    Prepare materials table for storage (JSON).
    """
    return json.dumps(material_rows, ensure_ascii=False)


def save_request(
    gs_client,
    user_login,
    positions,
    glass_type,
    material_sum,
    services_sum,
    total_area,
    total_perimeter,
    ensure_sum,
    total_sum,
    material_rows,
):
    """
    Append one request row to SHEET_REQUESTS.
    """

    row = [
        now_str(),                 # datetime
        user_login,                # user
        glass_type,                # glass type
        round(total_area, 3),      # total area
        round(total_perimeter, 3), # total perimeter
        round(material_sum, 2),    # materials sum
        round(services_sum, 2),    # glass + services sum
        round(ensure_sum, 2),      # ensure 65%
        round(total_sum, 2),       # total
        serialize_positions(positions),
        serialize_materials(material_rows),
    ]

    gs_client.append_row(SHEET_REQUESTS, row)
# =========================================
# COMMERCIAL PROPOSAL (Part 7 / 7.1)
# Excel template based (v15 logic)
# =========================================

from openpyxl import load_workbook
from tempfile import NamedTemporaryFile


KP_TEMPLATE_PATH = "kp_template.xlsx"


def fill_kp_template(template_path, output_path, data):
    """
    Fill Excel commercial proposal template with calculation data.
    Template structure must match v15.
    """

    wb = load_workbook(template_path)
    ws = wb.active

    # --- HEADER ---
    ws["B2"] = data["date"]
    ws["B3"] = data["user"]
    ws["B4"] = data["glass_type"]

    # --- GEOMETRY ---
    ws["B6"] = data["total_area"]
    ws["B7"] = data["total_perimeter"]

    # --- TOTALS ---
    ws["E10"] = data["materials_sum"]
    ws["E11"] = data["services_sum"]
    ws["E12"] = data["ensure_sum"]
    ws["E13"] = data["total_sum"]

    # --- MATERIALS TABLE ---
    start_row = 16

    for idx, row in enumerate(data["materials_rows"]):
        ws.cell(row=start_row + idx, column=1).value = row.get("Item")
        ws.cell(row=start_row + idx, column=2).value = row.get("Calculated qty")
        ws.cell(row=start_row + idx, column=3).value = row.get("Price")
        ws.cell(row=start_row + idx, column=4).value = row.get("Sum")

    wb.save(output_path)


def generate_kp_file(
    user_login,
    glass_type,
    totals,
    material_rows,
    material_sum,
    services_sum,
    ensure_sum,
    total_sum,
):
    """
    Generate filled KP Excel file and return its path.
    """

    with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        output_path = tmp.name

    data = {
        "date": now_str(),
        "user": user_login,
        "glass_type": glass_type,
        "total_area": round(totals["total_area"], 3),
        "total_perimeter": round(totals["total_perimeter"], 3),
        "materials_sum": round(material_sum, 2),
        "services_sum": round(services_sum, 2),
        "ensure_sum": round(ensure_sum, 2),
        "total_sum": round(total_sum, 2),
        "materials_rows": material_rows,
    }

    fill_kp_template(
        KP_TEMPLATE_PATH,
        output_path,
        data,
    )

    return output_path
# =========================================
# COMMERCIAL PROPOSAL (Part 7 / 7.2)
# Final integration with calculation
# =========================================

def calculate_full_result(gs_client, positions, glass_type):
    """
    Full calculation pipeline.
    """

    # --- MATERIALS ---
    material_calc = MaterialCalculator(gs_client)
    material_rows, material_sum = material_calc.calculate(positions)

    # --- SERVICES ---
    catalog = GlassServiceCatalog(gs_client)
    service_calc = GlassServiceCalculator(catalog)
    service_rows, services_sum = service_calc.calculate(
        positions,
        glass_type,
    )

    # --- TOTALS ---
    totals = aggregate_totals(positions)

    base_sum = material_sum + services_sum
    ensure_sum = base_sum * ENSURE_PERCENT
    total_sum = base_sum + ensure_sum

    return {
        "material_rows": material_rows,
        "material_sum": material_sum,
        "service_rows": service_rows,
        "services_sum": services_sum,
        "totals": totals,
        "ensure_sum": ensure_sum,
        "total_sum": total_sum,
    }


def save_request_and_offer_download(
    gs_client,
    positions,
    glass_type,
    user_login,
    calc_result,
):
    """
    Save request and show KP download button.
    """

    save_request(
        gs_client=gs_client,
        user_login=user_login,
        positions=positions,
        glass_type=glass_type,
        material_sum=calc_result["material_sum"],
        services_sum=calc_result["services_sum"],
        total_area=calc_result["totals"]["total_area"],
        total_perimeter=calc_result["totals"]["total_perimeter"],
        ensure_sum=calc_result["ensure_sum"],
        total_sum=calc_result["total_sum"],
        material_rows=calc_result["material_rows"],
    )

    if not os.path.exists(KP_TEMPLATE_PATH):
        st.error("KP template file not found")
        return

    kp_path = generate_kp_file(
        user_login=user_login,
        glass_type=glass_type,
        totals=calc_result["totals"],
        material_rows=calc_result["material_rows"],
        material_sum=calc_result["material_sum"],
        services_sum=calc_result["services_sum"],
        ensure_sum=calc_result["ensure_sum"],
        total_sum=calc_result["total_sum"],
    )

    with open(kp_path, "rb") as f:
        st.download_button(
            label="Download commercial proposal (Excel)",
            data=f,
            file_name="commercial_proposal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
# =========================================
# MAIN APPLICATION (Final)
# =========================================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    gs = GoogleSheetsClient(GSPREAD_SHEET_ID)

    if not login(gs):
        st.stop()

    st.header("Order parameters")

    # --- PRODUCT PARAMETERS ---
    col1, col2 = st.columns(2)

    with col1:
        product_type = st.selectbox(
            "Product type",
            [
                "Окно с откр.",
                "Окно глух.",
                "Дверь 1 створч.",
                "Дверь 2-х створч.",
                "Фасад",
            ],
        )

    with col2:
        profile_system = st.selectbox(
            "Profile system",
            [
                "ALG 2030-63C",
                "ALG 2030-55C",
                "ALG 2030-73C",
                "ALG 2030-45C",
                "ALG 2030-Slim",
                "Ruit 50F",
            ],
        )

    st.subheader("Geometry")

    g1, g2, g3 = st.columns(3)

    with g1:
        width_mm = st.number_input("Width (mm)", min_value=100.0, step=10.0)

    with g2:
        height_mm = st.number_input("Height (mm)", min_value=100.0, step=10.0)

    with g3:
        qty = st.number_input("Quantity", min_value=1, step=1, value=1)

    st.subheader("Imposts")

    i1, i2, i3, i4 = st.columns(4)

    with i1:
        left_mm = st.number_input("LEFT (mm)", min_value=0.0, step=10.0)

    with i2:
        center_mm = st.number_input("CENTER (mm)", min_value=0.0, step=10.0)

    with i3:
        right_mm = st.number_input("RIGHT (mm)", min_value=0.0, step=10.0)

    with i4:
        top_mm = st.number_input("TOP (mm)", min_value=0.0, step=10.0)

    sash_count = 0
    sash_width_mm = 0.0
    sash_height_mm = 0.0

    if "Окно" in product_type or "Дверь" in product_type:
        st.subheader("Sashes")

        s1, s2, s3 = st.columns(3)

        with s1:
            sash_count = st.number_input("Sash count", min_value=1, step=1, value=1)

        with s2:
            sash_width_mm = st.number_input("Sash width (mm)", min_value=200.0, step=10.0)

        with s3:
            sash_height_mm = st.number_input("Sash height (mm)", min_value=200.0, step=10.0)

    stand_step_mm = 0.0
    if product_type == "Фасад":
        stand_step_mm = st.number_input(
            "Facade stand step (mm)",
            min_value=300.0,
            step=50.0,
            value=1000.0,
        )

    # --- POSITIONS ---
    positions = [
        {
            "product_type": product_type,
            "profile_system": profile_system,
            "width_mm": width_mm,
            "height_mm": height_mm,
            "qty": qty,
            "left_mm": left_mm,
            "center_mm": center_mm,
            "right_mm": right_mm,
            "top_mm": top_mm,
            "sash_count": sash_count,
            "sash_width_mm": sash_width_mm,
            "sash_height_mm": sash_height_mm,
            "stand_step_mm": stand_step_mm,
        }
    ]

    # --- GLASS TYPE ---
    catalog = GlassServiceCatalog(gs)
    glass_types = catalog.get_glass_types()

    if not glass_types:
        st.error("No glass types found in reference sheet")
        st.stop()

    selected_glass_type = st.selectbox(
        "Glass type",
        glass_types,
    )

    # --- CALCULATE ---
    if st.button("Calculate", type="primary"):
        calc_result = calculate_full_result(
            gs_client=gs,
            positions=positions,
            glass_type=selected_glass_type,
        )

        st.success(f"TOTAL: {round(calc_result['total_sum'], 2)}")

        st.subheader("Totals")
        st.write(calc_result["totals"])

        st.subheader("Materials")
        if calc_result["material_rows"]:
            st.dataframe(pd.DataFrame(calc_result["material_rows"]))
        else:
            st.info("No materials")

        st.subheader("Services")
        if calc_result["service_rows"]:
            st.dataframe(
                pd.DataFrame(
                    calc_result["service_rows"],
                    columns=["Name", "Price", "Unit", "Sum"],
                )
            )
        else:
            st.info("No services")

        save_request_and_offer_download(
            gs_client=gs,
            positions=positions,
            glass_type=selected_glass_type,
            user_login=st.session_state["user_login"],
            calc_result=calc_result,
        )


if __name__ == "__main__":
    main()
