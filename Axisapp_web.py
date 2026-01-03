# =========================================
# Axis Pro GF — Calculator
# FULL VERSION / PART 1
# =========================================

import math
import ast
import operator as op
import json
import logging
import sys
import os
import datetime
import copy

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

ENSURE_PERCENT = 0.65

SERVICE_ACCOUNT_PATH = "/etc/secrets/gcp_service_account.json"

# =========================================
# LOGGER
# =========================================

logger = logging.getLogger("axis_pro_gf")
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)

logger.setLevel(logging.INFO)

# =========================================
# UTILS
# =========================================

def normalize_key(value):
    if value is None:
        return ""
    return " ".join(
        str(value)
        .replace("\xa0", " ")
        .replace(",", ".")
        .lower()
        .split()
    )


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


def get_field(row: dict, needle: str, default=None):
    needle = needle.lower()
    for k, v in row.items():
        if k and needle in str(k).lower():
            return v
    return default


def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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
        raise ValueError(f"Unknown variable: {node.id}")

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
        if isinstance(node.func, ast.Attribute):
            if (
                isinstance(node.func.value, ast.Name)
                and node.func.value.id == "math"
            ):
                fn = getattr(math, node.func.attr)
                args = [_eval_node(a, names) for a in node.args]
                return fn(*args)

        if isinstance(node.func, ast.Name):
            if node.func.id in ("min", "max", "round"):
                args = [_eval_node(a, names) for a in node.args]
                return globals()[node.func.id](*args)

    raise ValueError("Unsafe expression")


def safe_eval(formula: str, context: dict) -> float:
    if not formula:
        return 0.0
    try:
        ctx = {k: safe_float(v) for k, v in context.items()}
        ctx["math"] = math
        node = ast.parse(str(formula), mode="eval")
        return float(_eval_node(node, ctx))
    except Exception as e:
        logger.error(f"FORMULA ERROR [{formula}] -> {e}")
        return 0.0

# =========================================
# GOOGLE SHEETS
# =========================================

class GoogleSheets:

    @st.cache_resource
    def auth(_self):
        if not os.path.exists(SERVICE_ACCOUNT_PATH):
            st.error("❌ Не найден файл сервисного аккаунта Google")
            st.stop()

        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_PATH,
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        return gspread.authorize(creds)

    def __init__(self, sheet_id: str):
        self.client = self.auth()
        self.book = self.client.open_by_key(sheet_id)
        self._cache = {}

    def ws(self, name: str):
        if name not in self._cache:
            self._cache[name] = self.book.worksheet(name)
        return self._cache[name]

    @st.cache_data(ttl=1800)
    def read(self, sheet_name: str):
        return self.ws(sheet_name).get_all_records()

    def append(self, sheet_name: str, row: list):
        self.ws(sheet_name).append_row(
            row,
            value_input_option="USER_ENTERED"
        )

# =========================================
# LOGIN
# =========================================

def login(gs: GoogleSheets) -> bool:
    if "user" in st.session_state:
        return True

    st.sidebar.title("🔐 Вход")

    login_value = st.sidebar.text_input("Логин")
    password_value = st.sidebar.text_input("Пароль", type="password")

    if st.sidebar.button("Войти"):
        users = gs.read(SHEET_USERS)
        for u in users:
            if (
                str(get_field(u, "логин", "")).strip() == login_value
                and str(get_field(u, "пароль", "")).strip() == password_value
            ):
                st.session_state["user"] = login_value
                st.session_state["login_time"] = now_str()
                st.rerun()

        st.sidebar.error("Неверный логин или пароль")

    return False
# =========================================
# GEOMETRY / POSITIONS MODEL
# =========================================

def build_geometry_context(position: dict) -> dict:
    """
    Формирует геометрический контекст одной позиции
    """

    width = safe_float(position.get("width_mm"))
    height = safe_float(position.get("height_mm"))
    qty = safe_int(position.get("qty", 1), 1)

    # импосты
    left = safe_float(position.get("left_mm"))
    center = safe_float(position.get("center_mm"))
    right = safe_float(position.get("right_mm"))
    top = safe_float(position.get("top_mm"))

    # базовые величины
    area_one = (width * height) / 1_000_000
    perimeter_one = 2 * (width + height) / 1000

    # импосты (логика как в Excel)
    vert_imposts = sum(1 for x in [left, center, right] if x > 0)
    hor_imposts = 1 if top > 0 else 0

    impost_vert_count = max(0, vert_imposts - 1)
    impost_hor_count = hor_imposts

    impost_total = impost_vert_count + impost_hor_count

    frame_rectangles = 1 + impost_total
    corners = frame_rectangles * 4

    # створки
    sash_count = safe_int(position.get("sash_count", 0))
    sash_width = safe_float(position.get("sash_width_mm"))
    sash_height = safe_float(position.get("sash_height_mm"))

    sash_area_one = 0.0
    sash_perimeter_one = 0.0

    if sash_count > 0 and sash_width > 0 and sash_height > 0:
        sash_area_one = (sash_width * sash_height) / 1_000_000
        sash_perimeter_one = 2 * (sash_width + sash_height) / 1000

    context = {
        # габариты
        "width": width,
        "height": height,
        "qty": qty,

        # площади и периметры
        "area_one": area_one,
        "area_total": area_one * qty,
        "perimeter_one": perimeter_one,
        "perimeter_total": perimeter_one * qty,

        # импосты
        "impost_vert": impost_vert_count,
        "impost_hor": impost_hor_count,
        "impost_total": impost_total,

        # рамы
        "frame_rectangles": frame_rectangles,
        "corners": corners,

        # створки
        "sash_count": sash_count,
        "sash_area_one": sash_area_one,
        "sash_area_total": sash_area_one * sash_count * qty,
        "sash_perimeter_one": sash_perimeter_one,
        "sash_perimeter_total": sash_perimeter_one * sash_count * qty,
    }

    return context


# =========================================
# AGGREGATES (TOTAL AREA / PERIMETER)
# =========================================

def calculate_aggregates(positions: list) -> dict:
    """
    Считает суммарные габариты по ВСЕМ позициям
    """

    total_area = 0.0
    total_perimeter = 0.0
    total_sash_area = 0.0
    total_sash_perimeter = 0.0

    for pos in positions:
        ctx = build_geometry_context(pos)

        total_area += ctx["area_total"]
        total_perimeter += ctx["perimeter_total"]

        total_sash_area += ctx["sash_area_total"]
        total_sash_perimeter += ctx["sash_perimeter_total"]

    return {
        "total_area": round(total_area, 3),
        "total_perimeter": round(total_perimeter, 3),
        "total_sash_area": round(total_sash_area, 3),
        "total_sash_perimeter": round(total_sash_perimeter, 3),
    }


# =========================================
# POSITION FACTORY
# =========================================

def create_position(
    product_type: str,
    profile_system: str,
    width_mm: float,
    height_mm: float,
    qty: int,
    left_mm: float = 0,
    center_mm: float = 0,
    right_mm: float = 0,
    top_mm: float = 0,
    sash_count: int = 0,
    sash_width_mm: float = 0,
    sash_height_mm: float = 0,
    glass_type: str = "",
    toning: bool = False,
    assembly: bool = False,
    montage: bool = False,
):
    """
    Создаёт одну позицию (как строка в Excel ЗАПРОСЫ)
    """

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

        "sash_count": sash_count,
        "sash_width_mm": sash_width_mm,
        "sash_height_mm": sash_height_mm,

        "glass_type": glass_type,
        "toning": toning,
        "assembly": assembly,
        "montage": montage,
    }
# =========================================
# MATERIAL CALCULATOR
# =========================================

class MaterialCalculator:

    def __init__(self, gs: GoogleSheets):
        self.gs = gs
        self.ref1 = self.gs.read(SHEET_REF1)

    def calculate(self, positions: list):
        """
        Основной расчёт материалов
        """
        material_rows = []
        material_sum = 0.0

        for row in self.ref1:
            row_product_type = str(
                get_field(row, "тип изделия", "")
            ).strip()

            row_profile_system = str(
                get_field(row, "система проф", "")
            ).strip()

            element_type = str(
                get_field(row, "тип элемент", "")
            ).strip()

            product_name = str(
                get_field(row, "товар", "")
            ).strip()

            formula = get_field(row, "формула_python")

                        if not formula:
                continue

            total_fact = 0.0

            for pos in positions:
                if row_product_type and row_product_type != pos["product_type"]:
                    continue

                if row_profile_system and row_profile_system != pos["profile_system"]:
                    continue

                geom = build_geometry_context(pos)

                context = {
                    "W": geom["width"],
                    "H": geom["height"],
                    "area": geom["area_one"],
                    "perimeter": geom["perimeter_one"],
                    "qty": geom["qty"],
                    "impost_vert": geom["impost_vert"],
                    "impost_hor": geom["impost_hor"],
                    "impost_total": geom["impost_total"],
                    "sash_count": geom["sash_count"],
                    "sash_area": geom["sash_area_one"],
                    "sash_perimeter": geom["sash_perimeter_one"],
                    "corners": geom["corners"],
                }

                value = safe_eval(str(formula), context)
                total_fact += value * geom["qty"]

            if total_fact расход <= 0:
                continue

            price = safe_float(get_field(row, "цена за"))
            norm = safe_float(get_field(row, "кол-во норм"), 1)

            if norm > 0:
                ship_qty = math.ceil(total_fact расход / norm)
                real_qty = ship_qty * norm
            else:
                real_qty = total_fact расход

            row_sum = real_qty * price
            material_sum += row_sum

            material_rows.append({
                "Тип изделия": row_product_type,
                "Система профиля": row_profile_system,
                "Тип элемента": element_type,
                "Товар": product_name,
                "Фактический расход": round(total_fact расход, 3),
                "К отгрузке": round(real_qty, 3),
                "Цена": price,
                "Сумма": round(row_sum, 2),
            })

        return material_rows, round(material_sum, 2)
# =========================================
# GLASS / SERVICES CALCULATOR
# =========================================

class GlassServiceCalculator:

    def __init__(self, gs: GoogleSheets):
        self.gs = gs
        self.ref2 = self.gs.read(SHEET_REF2)

    # -------------------------------------
    # Поиск строки стеклопакета
    # -------------------------------------
    def _find_glass_row(self, glass_type: str):
        glass_type_norm = normalize_key(glass_type)

        for row in self.ref2:
            row_type = normalize_key(get_field(row, "тип стеклопакета"))
            if row_type == glass_type_norm:
                return row
        return None

    # -------------------------------------
    # Расчёт стекла + услуг
    # -------------------------------------
    def calculate(self, positions: list, aggregates: dict):
        rows = []
        total_sum = 0.0

        total_area = aggregates["total_area"]

        # ---- стеклопакет ----
        glass_type = ""
        for p in positions:
            if p.get("glass_type"):
                glass_type = p["glass_type"]
                break

        glass_row = self._find_glass_row(glass_type)

        glass_price = 0.0
        if glass_row:
            glass_price = safe_float(
                get_field(glass_row, "стоимость стеклопакета")
            )

        glass_sum = total_area * glass_price
        total_sum += glass_sum

        rows.append({
            "Наименование": f"Стеклопакет ({glass_type})",
            "Цена": glass_price,
            "Ед.": "м²",
            "Сумма": round(glass_sum, 2),
        })

        # ---- тонировка ----
        toning_price = 0.0
        toning_enabled = any(p.get("toning") for p in positions)

        if glass_row and toning_enabled:
            toning_flag = normalize_key(
                get_field(glass_row, "тонировка")
            )
            if toning_flag == "есть":
                toning_price = safe_float(
                    get_field(glass_row, "стоимость тонировки")
                )

        if toning_price > 0:
            toning_sum = total_area * toning_price
            total_sum += toning_sum

            rows.append({
                "Наименование": "Тонировка",
                "Цена": toning_price,
                "Ед.": "м²",
                "Сумма": round(toning_sum, 2),
            })

        # ---- сборка ----
        assembly_price = 0.0
        assembly_enabled = any(p.get("assembly") for p in positions)

        if glass_row and assembly_enabled:
            assembly_flag = normalize_key(
                get_field(glass_row, "сборка")
            )
            if assembly_flag == "есть":
                assembly_price = safe_float(
                    get_field(glass_row, "стоимость сборки")
                )

        if assembly_price > 0:
            assembly_sum = total_area * assembly_price
            total_sum += assembly_sum

            rows.append({
                "Наименование": "Сборка",
                "Цена": assembly_price,
                "Ед.": "м²",
                "Сумма": round(assembly_sum, 2),
            })

        # ---- монтаж ----
        montage_price = 0.0
        montage_enabled = any(p.get("montage") for p in positions)

        if glass_row and montage_enabled:
            montage_price = safe_float(
                get_field(glass_row, "стоимость монтаж")
            )

        if montage_price > 0:
            montage_sum = total_area * montage_price
            total_sum += montage_sum

            rows.append({
                "Наименование": "Монтаж",
                "Цена": montage_price,
                "Ед.": "м²",
                "Сумма": round(montage_sum, 2),
            })

        return rows, round(total_sum, 2)
# =========================================
# FINAL TOTAL CALCULATION
# =========================================

def calculate_final_totals(
    material_sum: float,
    glass_service_sum: float,
):
    """
    Финальный расчёт:
    база + 65% + итог
    """

    base_sum = material_sum + glass_service_sum
    ensure_sum = base_sum * ENSURE_PERCENT
    total_sum = base_sum + ensure_sum

    return {
        "base_sum": round(base_sum, 2),
        "ensure_65": round(ensure_sum, 2),
        "total_sum": round(total_sum, 2),
    }


# =========================================
# FINAL TABLE BUILDER
# =========================================

def build_final_table(
    material_rows: list,
    glass_rows: list,
    totals: dict,
):
    """
    Собирает финальную таблицу для UI и КП
    """

    rows = []

    # --- материалы ---
    for r in material_rows:
        rows.append({
            "Наименование": r["Товар"],
            "Цена": r["Цена"],
            "Ед.": "—",
            "Сумма": r["Сумма"],
        })

    # --- стекло и услуги ---
    for r in glass_rows:
        rows.append({
            "Наименование": r["Наименование"],
            "Цена": r["Цена"],
            "Ед.": r["Ед."],
            "Сумма": r["Сумма"],
        })

    # --- итоги ---
    rows.append({
        "Наименование": "ИТОГО (база)",
        "Цена": "",
        "Ед.": "",
        "Сумма": totals["base_sum"],
    })

    rows.append({
        "Наименование": "Обеспечение 65%",
        "Цена": "",
        "Ед.": "",
        "Сумма": totals["ensure_65"],
    })

    rows.append({
        "Наименование": "ИТОГО К ОПЛАТЕ",
        "Цена": "",
        "Ед.": "",
        "Сумма": totals["total_sum"],
    })

    return rows


# =========================================
# DEBUG: TOTALS CHECK
# =========================================

def debug_totals(
    aggregates: dict,
    material_sum: float,
    glass_sum: float,
    totals: dict,
):
    """
    Отладочный вывод итогов
    """

    st.subheader("Проверка итогов")

    st.write("Общая площадь, м²:", aggregates["total_area"])
    st.write("Общий периметр, м:", aggregates["total_perimeter"])
    st.write("Материалы:", material_sum)
    st.write("Стекло + услуги:", glass_sum)
    st.write("База:", totals["base_sum"])
    st.write("65%:", totals["ensure_65"])
    st.write("ИТОГО:", totals["total_sum"])
# =========================================
# REQUESTS / HISTORY
# =========================================

def serialize_positions(positions: list) -> str:
    try:
        return json.dumps(positions, ensure_ascii=False)
    except Exception:
        return ""


def serialize_final_table(final_rows: list) -> str:
    try:
        return json.dumps(final_rows, ensure_ascii=False)
    except Exception:
        return ""


def save_request_to_sheet(
    gs: GoogleSheets,
    user: str,
    positions: list,
    aggregates: dict,
    material_sum: float,
    glass_sum: float,
    totals: dict,
    final_rows: list,
):
    row = [
        now_str(),
        user,
        len(positions),
        round(aggregates["total_area"], 3),
        round(aggregates["total_perimeter"], 3),
        round(material_sum, 2),
        round(glass_sum, 2),
        round(totals["base_sum"], 2),
        round(totals["ensure_65"], 2),
        round(totals["total_sum"], 2),
        serialize_positions(positions),
        serialize_final_table(final_rows),
    ]

    gs.append(SHEET_FORM, row)


def show_history(gs: GoogleSheets):
    st.subheader("📜 История расчётов")

    rows = gs.read(SHEET_FORM)
    if not rows:
        st.info("История пока пуста")
        return

    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True)
# =========================================
# COMMERCIAL PROPOSAL (KP)
# =========================================

from openpyxl import load_workbook
from tempfile import NamedTemporaryFile


def generate_kp_excel(
    template_path: str,
    user: str,
    aggregates: dict,
    final_rows: list,
    totals: dict,
):
    wb = load_workbook(template_path)
    ws = wb.active

    ws["B2"] = "Коммерческое предложение"
    ws["B3"] = f"Менеджер: {user}"
    ws["B4"] = f"Дата: {now_str()}"

    ws["B6"] = aggregates["total_area"]
    ws["B7"] = aggregates["total_perimeter"]

    start_row = 10
    r = start_row

    for row in final_rows:
        ws[f"A{r}"] = row["Наименование"]
        ws[f"B{r}"] = row.get("Ед.", "")
        ws[f"C{r}"] = row.get("Цена", "")
        ws[f"D{r}"] = row.get("Сумма", "")
        r += 1

    ws[f"A{r}"] = "База"
    ws[f"D{r}"] = totals["base_sum"]
    r += 1

    ws[f"A{r}"] = "Обеспечение 65%"
    ws[f"D{r}"] = totals["ensure_65"]
    r += 1

    ws[f"A{r}"] = "ИТОГО К ОПЛАТЕ"
    ws[f"D{r}"] = totals["total_sum"]

    tmp = NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()

    return tmp.name


def kp_download_block(
    user: str,
    aggregates: dict,
    final_rows: list,
    totals: dict,
):
    st.subheader("📄 Коммерческое предложение")

    kp_template = st.file_uploader(
        "Загрузите шаблон КП (xlsx)",
        type=["xlsx"],
    )

    if not kp_template:
        return

    with NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        f.write(kp_template.read())
        template_path = f.name

    kp_path = generate_kp_excel(
        template_path=template_path,
        user=user,
        aggregates=aggregates,
        final_rows=final_rows,
        totals=totals,
    )

    with open(kp_path, "rb") as f:
        st.download_button(
            "⬇️ Скачать коммерческое предложение",
            f,
            file_name="Коммерческое_предложение.xlsx",
        )
# =========================================
# MAIN APPLICATION
# =========================================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title("🏗️ Axis Pro GF — Калькулятор")

    gs = GoogleSheets(GSPREAD_SHEET_ID)
    if not login(gs):
        st.stop()

    st.sidebar.header("Параметры заказа")

    product_type = st.sidebar.selectbox(
        "Тип изделия",
        [
            "Окно с откр.",
            "Окно  глух.",
            "Дверь 2-х створч.",
            "Дверь 1 створч.",
            "Фасад",
        ],
    )

    profile_system = st.sidebar.selectbox(
        "Система профиля",
        [
            "ALG 2030-63C",
            "ALG 2030-55C",
            "ALG 2030-73C",
            "ALG 2030-45C",
            "ALG 2030-Slim",
            "Ruit 50F",
        ],
    )

    glass_type = st.sidebar.selectbox(
        "Тип стеклопакета",
        [row.get("Тип стеклопакета") for row in gs.read(SHEET_REF2)],
    )

    toning = st.sidebar.checkbox("Тонировка")
    assembly = st.sidebar.checkbox("Сборка")
    montage = st.sidebar.checkbox("Монтаж")

    st.header("📦 Позиции")

    positions = []

    pos_count = st.number_input(
        "Количество позиций",
        min_value=1,
        step=1,
        value=1,
    )

    for i in range(int(pos_count)):
        st.subheader(f"Позиция #{i+1}")

        c1, c2, c3 = st.columns(3)
        width = c1.number_input("Ширина, мм", key=f"w_{i}", step=10.0)
        height = c2.number_input("Высота, мм", key=f"h_{i}", step=10.0)
        qty = c3.number_input("Кол-во", key=f"q_{i}", min_value=1, step=1)

        st.markdown("**Импосты**")
        i1, i2, i3, i4 = st.columns(4)
        left = i1.number_input("LEFT", key=f"l_{i}", step=10.0)
        center = i2.number_input("CENTER", key=f"c_{i}", step=10.0)
        right = i3.number_input("RIGHT", key=f"r_{i}", step=10.0)
        top = i4.number_input("TOP", key=f"t_{i}", step=10.0)

        sash_count = 0
        sash_width = 0.0
        sash_height = 0.0

        if "Окно с откр." in product_type or "Дверь" in product_type:
            sash_count = st.number_input(
                "Кол-во створок",
                key=f"sash_n_{i}",
                min_value=1,
                step=1,
            )
            s1, s2 = st.columns(2)
            sash_width = s1.number_input(
                "Ширина створки, мм",
                key=f"sash_w_{i}",
                step=10.0,
            )
            sash_height = s2.number_input(
                "Высота створки, мм",
                key=f"sash_h_{i}",
                step=10.0,
            )

        positions.append(
            create_position(
                product_type=product_type,
                profile_system=profile_system,
                width_mm=width,
                height_mm=height,
                qty=qty,
                left_mm=left,
                center_mm=center,
                right_mm=right,
                top_mm=top,
                sash_count=sash_count,
                sash_width_mm=sash_width,
                sash_height_mm=sash_height,
                glass_type=glass_type,
                toning=toning,
                assembly=assembly,
                montage=montage,
            )
        )

    if st.button("🚀 Рассчитать", type="primary"):
        aggregates = calculate_aggregates(positions)

        material_calc = MaterialCalculator(gs)
        material_rows, material_sum = material_calc.calculate(positions)

        glass_calc = GlassServiceCalculator(gs)
        glass_rows, glass_sum = glass_calc.calculate(
            positions,
            aggregates,
        )

        totals = calculate_final_totals(material_sum, glass_sum)
        final_rows = build_final_table(material_rows, glass_rows, totals)

        save_request_to_sheet(
            gs=gs,
            user=st.session_state.get("user", ""),
            positions=positions,
            aggregates=aggregates,
            material_sum=material_sum,
            glass_sum=glass_sum,
            totals=totals,
            final_rows=final_rows,
        )

        st.success(f"ИТОГО К ОПЛАТЕ: {totals['total_sum']}")

        tab1, tab2, tab3 = st.tabs(
            ["📐 Габариты", "🧱 Материалы", "💰 Итог"]
        )

        with tab1:
            st.metric("Общая площадь, м²", aggregates["total_area"])
            st.metric("Общий периметр, м", aggregates["total_perimeter"])

        with tab2:
            st.dataframe(pd.DataFrame(material_rows), use_container_width=True)
            st.write(f"Итого материалы: {material_sum}")

        with tab3:
            st.dataframe(pd.DataFrame(final_rows), use_container_width=True)
            st.write(f"База: {totals['base_sum']}")
            st.write(f"65%: {totals['ensure_65']}")
            st.write(f"ИТОГО: {totals['total_sum']}")

        kp_download_block(
            user=st.session_state.get("user", ""),
            aggregates=aggregates,
            final_rows=final_rows,
            totals=totals,
        )

    st.markdown("---")
    show_history(gs)


if __name__ == "__main__":
    main()
