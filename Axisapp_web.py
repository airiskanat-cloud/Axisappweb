import math
import os
import sys
from io import BytesIO

import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.drawing.image import Image as XLImage

# ======================================
# КОНСТАНТЫ И ПУТИ
# ======================================

def resource_path(relative_path: str) -> str:
    """
    Возвращает корректный путь к файлу как при обычном запуске,
    так и в упакованном PyInstaller-приложении.
    Для веб-версии по сути просто путь относительно файла.
    """
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


TEMPLATE_EXCEL_NAME = "axis_pro_gf.xlsx"
EXCEL_FILE = resource_path(TEMPLATE_EXCEL_NAME)

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# Шапка для листа ЗАПРОСЫ
FORM_HEADER = [
    "Номер заказа", "№ позиции",
    "Тип изделия", "Вид изделия", "Створки",
    "Профильная система",
    "Тип стеклопакета",
    "Режим заполнения",  # Стеклопакет / Ламбри / Сэндвич
    "Ширина, мм", "Высота, мм",
    "LEFT, мм", "CENTER, мм", "RIGHT, мм", "TOP, мм",
    "Ширина створки, мм", "Высота створки, мм",
    "Кол-во Nwin",
    "Тонировка", "Сборка", "Монтаж",
    "Тип ручек", "Доводчик"
]

# Брендинг для Excel
COMPANY_NAME = "ООО «Ваша Компания»"
COMPANY_CITY = "г. Ваш Город"
COMPANY_PHONE = "+7 (000) 000-00-00"
COMPANY_EMAIL = "info@yourcompany.kz"
COMPANY_SITE = "www.yourcompany.kz"
LOGO_FILENAME = "logo.png"  # логотип рядом с .py


# ======================================
# РАБОТА С EXCEL
# ======================================

class ExcelClient:
    def __init__(self, filename: str):
        self.filename = filename
        if not os.path.exists(self.filename):
            wb = Workbook()
            wb.save(self.filename)
        self.load()

    def load(self):
        self.wb = load_workbook(self.filename, data_only=True)

    def save(self):
        self.wb.save(self.filename)

    def ws(self, name: str):
        if name in self.wb.sheetnames:
            return self.wb[name]
        ws = self.wb.create_sheet(name)
        self.save()
        return ws

    def read_records(self, sheet_name: str):
        ws = self.ws(sheet_name)
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        header = rows[0]
        data_rows = rows[1:]
        records = []
        for r in data_rows:
            if all(v is None for v in r):
                continue
            rec = {}
            for i, key in enumerate(header):
                if key is None:
                    continue
                rec[str(key)] = r[i] if i < len(r) else None
            records.append(rec)
        return records

    def clear_and_write(self, sheet_name: str, header: list, rows: list):
        ws = self.ws(sheet_name)
        ws.delete_rows(1, ws.max_row or 1)
        if header:
            ws.append(header)
        for row in rows:
            ws.append(row)
        self.save()

    def append_form_row(self, row: list):
        ws = self.ws(SHEET_FORM)
        if ws.max_row == 1 and all(c.value is None for c in ws[1]):
            ws.append(FORM_HEADER)
        ws.append(row)
        self.save()


# ======================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ======================================

def safe_float(value, default=0.0):
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return default


def safe_int(value, default=0):
    try:
        return int(float(str(value).replace(",", ".")))
    except Exception:
        return default


def get_field(row: dict, needle: str, default=None):
    needle = needle.lower()
    for k in row.keys():
        if k is None:
            continue
        if needle in str(k).lower():
            return row[k]
    return default


def eval_formula(formula: str, context: dict) -> float:
    """
    Считает формулу на Python для ОДНОЙ позиции.
    Формула из Excel выполняется через eval с ограниченным набором переменных.
    """
    formula = (formula or "").strip()
    if not formula:
        return 0.0

    allowed_names = {
        "width": context.get("width", 0.0),
        "height": context.get("height", 0.0),
        "left": context.get("left", 0.0),
        "center": context.get("center", 0.0),
        "right": context.get("right", 0.0),
        "top": context.get("top", 0.0),
        "sash_width": context.get("sash_width", 0.0),
        "sash_height": context.get("sash_height", 0.0),
        "area": context.get("area", 0.0),
        "perimeter": context.get("perimeter", 0.0),
        "qty": context.get("qty", 0.0),
        "nsash": context.get("nsash", 1),
        "n_sash_active": context.get("n_sash_active", 1),
        "n_sash_passive": context.get("n_sash_passive", 0),
        "hinges_per_sash": context.get("hinges_per_sash", 3),
        "n_rect": context.get("n_rect", 1),
        "n_frame_rect": context.get("n_frame_rect", 1),
        "n_impost": context.get("n_impost", 0),
        "N_impost": context.get("n_impost", 0),
        "n_imp_vert": context.get("n_imp_vert", 0),
        "n_imp_hor": context.get("n_imp_hor", 0),
        "n_corners": context.get("n_corners", 0),
        "n_nodes_12": context.get("n_nodes_12", 0),
        "n_nodes_19": context.get("n_nodes_19", 0),
        "n_nodes_6_5": context.get("n_nodes_6_5", 0),
        "n_nodes_17_2": context.get("n_nodes_17_2", 0),
        "n_nodes_42": context.get("n_nodes_42", 0),
        "Nwin": context.get("qty", 0.0),
        "math": math,
        "max": max,
        "min": min,
    }

    try:
        result = eval(formula, {"__builtins__": {}}, allowed_names)
        return float(result)
    except Exception as e:
        print(f"Ошибка в формуле '{formula}': {e}")
        return 0.0


# ======================================
# ПОЛЬЗОВАТЕЛИ
# ======================================

def load_users(excel: ExcelClient):
    excel.load()
    try:
        rows = excel.read_records(SHEET_USERS)
    except Exception:
        return {}

    users = {}
    for row in rows:
        login = str(get_field(row, "логин", "") or "").strip()
        password = str(get_field(row, "парол", "") or "").strip()
        role = str(get_field(row, "роль", "") or "").strip()
        if login:
            users[login] = {"password": password, "role": role}
    return users


def login_form(excel: ExcelClient):
    if "current_user" in st.session_state:
        return st.session_state["current_user"]

    st.sidebar.title("🔐 Вход в систему")

    login = st.sidebar.text_input("Логин")
    password = st.sidebar.text_input("Пароль", type="password")
    btn = st.sidebar.button("Войти")

    users = load_users(excel)

    if btn:
        user = users.get(login)
        if user and password == user["password"]:
            st.session_state["current_user"] = {
                "login": login,
                "role": user.get("role", ""),
            }
            st.sidebar.success(f"Привет, {login}!")
            return st.session_state["current_user"]
        else:
            st.sidebar.error("Неверный логин или пароль")

    return None


# ======================================
# РАСЧЁТ ПО ГАБАРИТАМ (СПРАВОЧНИК -3)
# ======================================

class GabaritCalculator:
    HEADER = ["Тип элемента", "Фактическое значение"]

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def _calc_imposts_context(self, width, height, left, center, right, top):
        """
        Вспомогательная функция: считает количество импостов/рам/углов
        по габаритам и возвращает словарь с n_imp_vert/n_imp_hor/...
        """
        n_imp_vert = 0
        if left > 0:
            n_imp_vert += 1
        if center > 0:
            n_imp_vert += 1
        if right > 0:
            n_imp_vert += 1

        n_imp_hor = 0
        if top > 0:
            n_imp_hor += 1

        n_impost = n_imp_vert + n_imp_hor
        n_frame_rect = 1 + n_imp_vert + n_imp_hor
        n_rect = n_frame_rect
        n_corners = 4 * n_frame_rect

        return {
            "n_imp_vert": n_imp_vert,
            "n_imp_hor": n_imp_hor,
            "n_impost": n_impost,
            "n_frame_rect": n_frame_rect,
            "n_rect": n_rect,
            "n_corners": n_corners,
        }

    def calculate(self, order: dict, positions: list):
        ref_rows = self.excel.read_records(SHEET_REF3)
        if not ref_rows:
            return [], 0.0

        try:
            nsash = int(order.get("sashes", "1"))
        except ValueError:
            nsash = 1
        n_sash_active = 1 if nsash >= 1 else 0
        n_sash_passive = max(nsash - 1, 0)
        hinges_per_sash = 3

        total_area = sum(p["area_m2"] * p["Nwin"] for p in positions)
        gabarit_values = []

        for row in ref_rows:
            type_elem = get_field(row, "тип элемент", "")
            formula = get_field(row, "формула_python", "")
            if not type_elem or not formula:
                continue

            total_value = 0.0

            for p in positions:
                width = p["width_mm"]
                height = p["height_mm"]
                left = p.get("left_mm", 0.0)
                center = p.get("center_mm", 0.0)
                right = p.get("right_mm", 0.0)
                top = p.get("top_mm", 0.0)
                sash_w = p.get("sash_width_mm", width)
                sash_h = p.get("sash_height_mm", height)
                area = p["area_m2"]
                perimeter = p["perimeter_m"]
                qty = p["Nwin"]

                geom = self._calc_imposts_context(width, height, left, center, right, top)

                ctx = {
                    "width": width,
                    "height": height,
                    "left": left,
                    "center": center,
                    "right": right,
                    "top": top,
                    "sash_width": sash_w,
                    "sash_height": sash_h,
                    "area": area,
                    "perimeter": perimeter,
                    "qty": qty,
                    "nsash": nsash,
                    "n_sash_active": n_sash_active,
                    "n_sash_passive": n_sash_passive,
                    "hinges_per_sash": hinges_per_sash,
                }
                ctx.update(geom)

                total_value += eval_formula(str(formula), ctx)

            gabarit_values.append([type_elem, total_value])

        self.excel.clear_and_write(SHEET_GABARITS, self.HEADER, gabarit_values)
        return gabarit_values, total_area


# ======================================
# РАСЧЁТ МАТЕРИАЛОВ (СПРАВОЧНИК -1)
# ======================================

class MaterialCalculator:
    HEADER = [
        "Тип изделия", "Система профиля", "Тип элемента", "Артикул", "Товар",
        "Ед.", "Цена за ед.", "Ед. фактического расхода",
        "Кол-во фактического расхода (J)",
        "Норма к упаковке", "Ед. к отгрузке",
        "Кол-во к отгрузке", "Сумма"
    ]

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def _calc_imposts_context(self, width, height, left, center, right, top):
        n_imp_vert = 0
        if left > 0:
            n_imp_vert += 1
        if center > 0:
            n_imp_vert += 1
        if right > 0:
            n_imp_vert += 1

        n_imp_hor = 0
        if top > 0:
            n_imp_hor += 1

        n_impost = n_imp_vert + n_imp_hor
        n_frame_rect = 1 + n_imp_vert + n_imp_hor
        n_rect = n_frame_rect
        n_corners = 4 * n_frame_rect

        return {
            "n_imp_vert": n_imp_vert,
            "n_imp_hor": n_imp_hor,
            "n_impost": n_impost,
            "n_frame_rect": n_frame_rect,
            "n_rect": n_rect,
            "n_corners": n_corners,
        }

    def calculate(self, order: dict, positions_for_materials: list, selected_duplicates: dict):
        ref_rows = self.excel.read_records(SHEET_REF1)
        total_area = sum(p["area_m2"] * p["Nwin"] for p in positions_for_materials)
        if not ref_rows:
            return [], 0.0, total_area

        try:
            nsash = int(order.get("sashes", "1"))
        except ValueError:
            nsash = 1
        n_sash_active = 1 if nsash >= 1 else 0
        n_sash_passive = max(nsash - 1, 0)
        hinges_per_sash = 3

        result_rows = []
        total_sum = 0.0

        for row in ref_rows:
            row_type = get_field(row, "тип издел", "")
            row_profile = get_field(row, "система проф", "")
            type_elem = get_field(row, "тип элемент", "")
            product_name = str(get_field(row, "товар", "") or "")

            # фильтр по типу изделия
            if row_type:
                if str(row_type).strip().lower() != order["product_type"].strip().lower():
                    continue

            # фильтр по системе профиля
            if row_profile:
                if str(row_profile).strip().lower() != order["profile_system"].strip().lower():
                    continue

            # фильтр по выбору дублей
            if type_elem in selected_duplicates and selected_duplicates[type_elem]:
                chosen_names = selected_duplicates[type_elem]
                if product_name not in chosen_names:
                    continue

            formula = get_field(row, "формула_python", "")
            if not formula:
                formula = get_field(row, "формула фактического расхода", "")
            if not formula:
                continue

            qty_fact_total = 0.0

            for p in positions_for_materials:
                width = p["width_mm"]
                height = p["height_mm"]
                left = p.get("left_mm", 0.0)
                center = p.get("center_mm", 0.0)
                right = p.get("right_mm", 0.0)
                top = p.get("top_mm", 0.0)
                sash_w = p.get("sash_width_mm", width)
                sash_h = p.get("sash_height_mm", height)
                area = p["area_m2"]
                perimeter = p["perimeter_m"]
                qty = p["Nwin"]

                geom = self._calc_imposts_context(width, height, left, center, right, top)

                ctx = {
                    "width": width,
                    "height": height,
                    "left": left,
                    "center": center,
                    "right": right,
                    "top": top,
                    "sash_width": sash_w,
                    "sash_height": sash_h,
                    "area": area,
                    "perimeter": perimeter,
                    "qty": qty,
                    "nsash": nsash,
                    "n_sash_active": n_sash_active,
                    "n_sash_passive": n_sash_passive,
                    "hinges_per_sash": hinges_per_sash,
                }
                ctx.update(geom)

                qty_fact_total += eval_formula(str(formula), ctx)

            unit_price = safe_float(get_field(row, "цена за", 0.0))
            norm_per_pack = safe_float(get_field(row, "кол-во норм", 0.0))
            unit_pack = str(get_field(row, "ед .норма к упаковке", "") or "").strip()
            unit = str(get_field(row, "ед.", "") or "").strip()
            unit_fact = str(get_field(row, "ед. фактического расхода", "") or "").strip()

            if norm_per_pack > 0:
                qty_to_ship = math.ceil(qty_fact_total / norm_per_pack)
                effective_qty = qty_to_ship * norm_per_pack
            else:
                qty_to_ship = qty_fact_total
                effective_qty = qty_fact_total

            sum_row = effective_qty * unit_price
            total_sum += sum_row

            result_rows.append([
                row_type if row_type is not None else "",
                row_profile if row_profile is not None else "",
                type_elem,
                get_field(row, "артикул", ""),
                product_name,
                unit,
                unit_price,
                unit_fact,
                qty_fact_total,
                norm_per_pack,
                unit_pack,
                qty_to_ship,
                sum_row
            ])

        self.excel.clear_and_write(SHEET_MATERIAL, self.HEADER, result_rows)
        return result_rows, total_sum, total_area


# ======================================
# ИТОГОВЫЙ РАСЧЁТ (СПРАВОЧНИК -2)
# ======================================

class FinalCalculator:
    HEADER = ["Наименование услуг", "Стоимость за м²", "Ед", "Итого"]

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def calculate(self,
                  order: dict,
                  total_area_all: float,
                  total_area_glass: float,
                  material_total: float,
                  tambour_door_count: int = 0):
        ref_rows = self.excel.read_records(SHEET_REF2)

        glass_type = order["glass_type"]
        filling_mode = order["filling_mode"]
        toning = order["toning"]
        assembly = order["assembly"]
        montage = order["montage"]
        handle_type = order["handle_type"]
        door_closer = order["door_closer"]

        selected = None
        for row in ref_rows:
            row_glass = str(get_field(row, "тип стеклопак", "") or "").strip()
            row_fill_mode = str(get_field(row, "заполн", "") or "").strip()
            row_handle_type = str(get_field(row, "ручк", "") or "").strip()

            if row_glass and row_glass != glass_type:
                continue
            if row_fill_mode and row_fill_mode != filling_mode:
                continue
            if handle_type and row_handle_type and row_handle_type != handle_type:
                continue

            selected = row
            break

        if not selected and ref_rows:
            selected = ref_rows[0]
        elif not selected:
            selected = {}

        price_glass = safe_float(get_field(selected, "стоимость стеклопак", 0.0))
        price_toning = safe_float(get_field(selected, "стоимость тониров", 0.0))
        price_assembly = safe_float(get_field(selected, "стоимость сборк", 0.0))
        price_montage = safe_float(get_field(selected, "стоимость монтаж", 0.0))
        price_handles = safe_float(get_field(selected, "стоимость ручек", 0.0))
        price_closer = safe_float(get_field(selected, "стоимость доводчик", 0.0))

        rows = []

        # Стеклопакет
        if filling_mode == "Стеклопакет" and total_area_glass > 0:
            glass_sum = total_area_glass * price_glass
        else:
            glass_sum = 0.0
            price_glass = 0.0
        rows.append(["Стеклопакет", price_glass, "за м²", glass_sum])

        # Тонировка
        if toning == "Есть" and filling_mode == "Стеклопакет" and total_area_glass > 0:
            toning_sum = total_area_glass * price_toning
        else:
            toning_sum = 0.0
            price_toning = 0.0
        rows.append(["Тонировка", price_toning, "за м²", toning_sum])

        # Сборка
        if assembly == "Есть":
            assembly_sum = total_area_all * price_assembly
        else:
            assembly_sum = 0.0
            price_assembly = 0.0
        rows.append(["Сборка", price_assembly, "за м²", assembly_sum])

        # Монтаж
        if montage == "Есть":
            montage_sum = total_area_all * price_montage
        else:
            montage_sum = 0.0
            price_montage = 0.0
        rows.append(["Монтаж", price_montage, "за м²", montage_sum])

        # Материалы
        rows.append(["Материал", "-", "-", material_total])

        # Ручки
        handles_sum = 0.0
        if handle_type:
            if order["product_type"].lower() == "тамбур":
                handles_qty = max(tambour_door_count, 0)
            else:
                handles_qty = 1
            handles_sum = price_handles * handles_qty
        rows.append(["Ручки", price_handles, "шт.", handles_sum])

        # Доводчик
        closer_sum = 0.0
        if door_closer == "Есть":
            if order["product_type"].lower() == "тамбур":
                closer_qty = max(tambour_door_count, 0)
            else:
                closer_qty = 1
            closer_sum = price_closer * closer_qty
        rows.append(["Доводчик", price_closer, "шт.", closer_sum])

        base_sum = (
            glass_sum
            + toning_sum
            + assembly_sum
            + montage_sum
            + material_total
            + handles_sum
            + closer_sum
        )

        ensure_sum = base_sum * 0.6
        rows.append(["Обеспечение", "", "", ensure_sum])

        total_sum = base_sum + ensure_sum
        extra_rows = [["ИТОГО", "", "", total_sum]]

        self.excel.clear_and_write(SHEET_FINAL, self.HEADER, rows + extra_rows)
        return rows, total_sum, ensure_sum


# ======================================
# ЭКСПОРТ КОММЕРЧЕСКОГО ПРЕДЛОЖЕНИЯ
# ======================================

def build_smeta_workbook(order: dict,
                         base_positions: list,
                         lambr_positions: list,
                         total_area: float,
                         total_sum: float) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"

    logo_path = resource_path(LOGO_FILENAME)
    current_row = 1

    # Логотип
    if os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.height = 80
            img.width = 80
            ws.add_image(img, "A1")
        except Exception as e:
            print(f"Не удалось вставить логотип: {e}")

    # Реквизиты
    ws.cell(row=current_row, column=3, value=COMPANY_NAME)
    current_row += 1
    ws.cell(row=current_row, column=3, value=COMPANY_CITY)
    current_row += 1
    ws.cell(row=current_row, column=3, value=f"Тел.: {COMPANY_PHONE}")
    current_row += 1
    ws.cell(row=current_row, column=3, value=f"E-mail: {COMPANY_EMAIL}")
    current_row += 1
    ws.cell(row=current_row, column=3, value=f"Сайт: {COMPANY_SITE}")
    current_row += 2

    ws.cell(row=current_row, column=1, value="Коммерческое предложение")
    current_row += 2

    # Общие данные заказа
    ws.cell(row=current_row, column=1, value=f"Заказ № {order['order_number']}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип изделия: {order['product_type']}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Вид изделия: {order['product_view']}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Профильная система: {order['profile_system']}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип заполнения: {order['filling_mode']}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип стеклопакета: {order['glass_type']}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тонировка: {order['toning']}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Сборка: {order['assembly']}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Монтаж: {order['montage']}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип ручек: {order['handle_type'] or '—'}")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Доводчик: {order['door_closer']}")
    current_row += 2

    # Состав позиции
    ws.cell(row=current_row, column=1, value="Состав позиции:")
    current_row += 1

    if order["product_type"].lower() == "тамбур":
        tambour_sections = order.get("tambour_sections", [])
        ws.cell(row=current_row, column=1, value="Тамбур (единое изделие):")
        current_row += 1

        if base_positions:
            p = base_positions[0]
            ws.cell(
                row=current_row,
                column=1,
                value=f"  Рама: {p['width_mm']} × {p['height_mm']} мм, N = {p['Nwin']}"
            )
            current_row += 1

        door_index = 1
        panel_index = 1
        for sec in tambour_sections:
            kind = sec.get("kind", "section")
            w = sec["width_mm"]
            h = sec["height_mm"]
            q = sec["Nwin"]
            if kind == "door":
                title = f"Дверь {door_index}"
                door_index += 1
            elif kind == "panel":
                title = f"Глухая секция {panel_index}"
                panel_index += 1
            else:
                title = "Секция"

            ws.cell(
                row=current_row,
                column=1,
                value=f"  {title}: {w} × {h} мм, N = {q}"
            )
            current_row += 1
    else:
        for idx, p in enumerate(base_positions, start=1):
            ws.cell(
                row=current_row,
                column=1,
                value=(
                    f"Позиция {idx}: "
                    f"{order['product_type']}, {order['product_view']}, "
                    f"{p['width_mm']} × {p['height_mm']} мм, N = {p['Nwin']}"
                )
            )
            current_row += 1

    # Панели Ламбри / Сэндвич
    if lambr_positions:
        current_row += 1
        ws.cell(row=current_row, column=1, value="Панели Ламбри / Сэндвич:")
        current_row += 1
        for idx, p in enumerate(lambr_positions, start=1):
            ws.cell(
                row=current_row,
                column=1,
                value=(
                    f"Панель {idx}: {p['width_mm']} × {p['height_mm']} мм, "
                    f"N = {p['Nwin']}"
                )
            )
            current_row += 1

    current_row += 2
    ws.cell(row=current_row, column=1, value=f"Общая площадь: {total_area:.3f} м²")
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"ИТОГО к оплате: {total_sum:.2f}")

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ======================================
# WEB-ИНТЕРФЕЙС НА STREAMLIT
# ======================================

def main():
    st.set_page_config(page_title="Axis Pro GF • Калькулятор", layout="wide")

    if not os.path.exists(EXCEL_FILE):
        st.error(f"Не найден Excel-файл справочника: {EXCEL_FILE}")
        st.stop()

    excel = ExcelClient(EXCEL_FILE)

    # Авторизация
    user = login_form(excel)
    if not user:
        st.stop()

    st.title("📘 Калькулятор алюминиевых изделий (Axis Pro GF)")
    st.info(f"Пользователь: **{user['login']}**")

    # Загружаем Справочник-2, чтобы взять типы ручек
    ref2_records = excel.read_records(SHEET_REF2)
    handle_types_set = set()
    for row in ref2_records:
        hname = get_field(row, "ручк", "")
        if hname:
            handle_types_set.add(str(hname).strip())
    handle_types = sorted(list(handle_types_set)) if handle_types_set else [""]

    # ---------- Общие данные заказа (сайдбар) ----------
    with st.sidebar:
        st.header("Общие данные заказа")

        order_number = st.text_input("Номер заказа", value="")
        product_type = st.selectbox("Тип изделия", ["Окно", "Дверь", "Тамбур"])
        product_view = st.selectbox("Вид изделия", ["Стандарт", "С фрамугой"])
        sashes = st.selectbox("Створки", ["1", "2"])

        profile_system = st.selectbox(
            "Профильная система",
            [
                "ALG 2030-45C",
                "ALG RUIT 63i",
                "ALG RUIT 73",
            ]
        )

        glass_type = st.selectbox(
            "Тип стеклопакета",
            [
                "двойной",
                "тройной",
                "энергодвойной",
                "энерготройной",
                "Одинарный 4мм",
                "Одинарный 6мм",
                "Одинарный 4мм закал",
                "Одинарный 6мм закал",
            ]
        )

        toning = st.selectbox("Тонировка", ["Нет", "Есть"])
        assembly = st.selectbox("Сборка", ["Нет", "Есть"])
        montage = st.selectbox("Монтаж", ["Нет", "Есть"])

        handle_type = st.selectbox(
            "Тип ручек",
            handle_types,
            index=0 if handle_types else 0
        )

        door_closer = st.selectbox("Доводчик", ["Нет", "Есть"])

        if product_type == "Тамбур":
            positions_count = 1
            st.caption("Для тамбура считается одна позиция с внутренними секциями.")
        else:
            positions_count = st.number_input(
                "Количество позиций",
                min_value=1,
                max_value=10,
                value=1,
                step=1
            )

    # ---------- Основная часть: две колонки ----------
    col_left, col_right = st.columns([2, 1])

    # Сначала правая колонка: выбор типа заполнения
    with col_right:
        st.header("Заполнение / панели")

        filling_mode = st.radio(
            "Тип заполнения",
            ["Стеклопакет", "Ламбри", "Сэндвич"],
            index=0
        )

        if filling_mode == "Стеклопакет":
            st.caption("Для режима «Стеклопакет» отдельные панели Ламбри/Сэндвич не задаются.")
        else:
            st.caption(
                f"Выбрано заполнение: **{filling_mode}**. "
                f"Габариты панелей задаются под основными габаритами слева."
            )

    # Левая колонка: габариты + панели
    lambr_positions_inputs = []
    positions_inputs = []
    tambour_sections_inputs = []

    with col_left:
        st.header("🧱 Позиции (габариты изделий)")

        for i in range(int(positions_count)):
            st.subheader(f"Позиция {i + 1}")
            c1, c2, c3, c4 = st.columns(4)

            width_mm = c1.number_input(
                f"Ширина, мм (поз. {i+1})",
                min_value=0.0,
                step=10.0,
                key=f"w_{i}"
            )
            height_mm = c2.number_input(
                f"Высота, мм (поз. {i+1})",
                min_value=0.0,
                step=10.0,
                key=f"h_{i}"
            )
            left_mm = c3.number_input(
                f"LEFT, мм (поз. {i+1})",
                min_value=0.0,
                step=10.0,
                key=f"l_{i}"
            )
            right_mm = c4.number_input(
                f"RIGHT, мм (поз. {i+1})",
                min_value=0.0,
                step=10.0,
                key=f"r_{i}"
            )

            c5, c6, c7, c8 = st.columns(4)
            center_mm = c5.number_input(
                f"CENTER, мм (поз. {i+1})",
                min_value=0.0,
                step=10.0,
                key=f"c_{i}"
            )
            top_mm = c6.number_input(
                f"TOP, мм (поз. {i+1})",
                min_value=0.0,
                step=10.0,
                key=f"t_{i}"
            )
            sash_width_mm = c7.number_input(
                f"Ширина створки, мм (поз. {i+1})",
                min_value=0.0,
                step=10.0,
                key=f"sw_{i}"
            )
            sash_height_mm = c8.number_input(
                f"Высота створки, мм (поз. {i+1})",
                min_value=0.0,
                step=10.0,
                key=f"sh_{i}"
            )

            c9, _ = st.columns(2)
            Nwin = c9.number_input(
                f"Кол-во Nwin (поз. {i+1})",
                min_value=1,
                step=1,
                value=1,
                key=f"nwin_{i}"
            )

            position_data = {
                "width_mm": width_mm,
                "height_mm": height_mm,
                "left_mm": left_mm,
                "center_mm": center_mm,
                "right_mm": right_mm,
                "top_mm": top_mm,
                "sash_width_mm": sash_width_mm,
                "sash_height_mm": sash_height_mm,
                "Nwin": Nwin,
            }

            # Расширенная карточка для Тамбура (только первая позиция)
            if product_type == "Тамбур" and i == 0:
                with st.expander("Параметры Тамбура (двери и глухие секции)", expanded=True):
                    door_count = st.number_input(
                        "Количество дверей",
                        min_value=1,
                        value=1,
                        step=1,
                        key="tambour_door_count"
                    )
                    door_inputs = []
                    for d in range(int(door_count)):
                        d1, d2, d3 = st.columns(3)
                        dw = d1.number_input(
                            f"Ширина двери {d+1}, мм",
                            min_value=0.0,
                            step=10.0,
                            key=f"door_w_{d}"
                        )
                        dh = d2.number_input(
                            f"Высота двери {d+1}, мм",
                            min_value=0.0,
                            step=10.0,
                            key=f"door_h_{d}"
                        )
                        dq = d3.number_input(
                            f"N (дверь {d+1})",
                            min_value=1,
                            value=1,
                            step=1,
                            key=f"door_q_{d}"
                        )
                        door_inputs.append({
                            "kind": "door",
                            "width_mm": dw,
                            "height_mm": dh,
                            "Nwin": dq,
                        })

                    panel_count = st.number_input(
                        "Количество глухих секций",
                        min_value=1,
                        value=1,
                        step=1,
                        key="tambour_panel_count"
                    )
                    panel_inputs = []
                    for p_idx in range(int(panel_count)):
                        p1, p2, p3 = st.columns(3)
                        pw = p1.number_input(
                            f"Ширина глухой секции {p_idx+1}, мм",
                            min_value=0.0,
                            step=10.0,
                            key=f"panel_w_{p_idx}"
                        )
                        ph = p2.number_input(
                            f"Высота глухой секции {p_idx+1}, мм",
                            min_value=0.0,
                            step=10.0,
                            key=f"panel_h_{p_idx}"
                        )
                        pq = p3.number_input(
                            f"N (секция {p_idx+1})",
                            min_value=1,
                            value=1,
                            step=1,
                            key=f"panel_q_{p_idx}"
                        )
                        panel_inputs.append({
                            "kind": "panel",
                            "width_mm": pw,
                            "height_mm": ph,
                            "Nwin": pq,
                        })

                    tambour_sections_inputs = door_inputs + panel_inputs
                    position_data["tambour_sections"] = tambour_sections_inputs

            positions_inputs.append(position_data)

        # Панели Ламбри / Сэндвич — теперь здесь
        if filling_mode in ("Ламбри", "Сэндвич"):
            st.subheader(f"Панели {filling_mode}")

            panel_count_ls = st.number_input(
                f"Количество панелей ({filling_mode})",
                min_value=1,
                value=1,
                step=1,
                key="ls_panel_count"
            )

            for i in range(int(panel_count_ls)):
                st.markdown(f"**Панель {i + 1}**")
                p1, p2, p3 = st.columns(3)
                w = p1.number_input(
                    f"Ширина панели {i+1}, мм",
                    min_value=0.0,
                    step=10.0,
                    key=f"ls_w_{i}"
                )
                h = p2.number_input(
                    f"Высота панели {i+1}, мм",
                    min_value=0.0,
                    step=10.0,
                    key=f"ls_h_{i}"
                )
                q = p3.number_input(
                    f"N (панель {i+1})",
                    min_value=1,
                    value=1,
                    step=1,
                    key=f"ls_q_{i}"
                )

                lambr_positions_inputs.append({
                    "width_mm": w,
                    "height_mm": h,
                    "Nwin": q,
                    "left_mm": 0.0,
                    "center_mm": 0.0,
                    "right_mm": 0.0,
                    "top_mm": 0.0,
                    "sash_width_mm": w,
                    "sash_height_mm": h,
                })

    # ---------- Выбор материалов при дублях ----------
    st.header("🧾 Выбор материалов при дублях (если в справочнике несколько товаров на один элемент)")
    selected_duplicates = {}

    ref1 = excel.read_records(SHEET_REF1)
    groups = {}
    for row in ref1:
        row_type = str(get_field(row, "тип издел", "") or "").strip()
        row_profile = str(get_field(row, "система проф", "") or "").strip()

        if row_type.lower() != product_type.lower():
            continue
        if row_profile.lower() != profile_system.lower():
            continue

        type_elem = str(get_field(row, "тип элемент", "") or "").strip()
        product_name = str(get_field(row, "товар", "") or "").strip()
        if not type_elem or not product_name:
            continue

        groups.setdefault(type_elem, set()).add(product_name)

    if not groups:
        st.info("Для выбранного типа изделия и профиля дублей материалов не найдено.")
    else:
        for type_elem, products in sorted(groups.items(), key=lambda kv: kv[0]):
            if len(products) <= 1:
                continue
            default = sorted(list(products))
            chosen = st.multiselect(
                f"Тип элемента: {type_elem}",
                options=sorted(list(products)),
                default=default,
                key=f"dup_{type_elem}"
            )
            selected_duplicates[type_elem] = set(chosen)

    # ---------- Кнопка расчёта ----------
    st.markdown("---")
    calc_button = st.button("💾 Сохранить в Excel и выполнить расчёт")

    if calc_button:
        if not order_number.strip():
            st.error("Введите номер заказа в левой панели.")
            st.stop()

        base_positions = []
        for p in positions_inputs:
            if p["width_mm"] <= 0 or p["height_mm"] <= 0:
                st.error("Во всех позициях ширина и высота должны быть больше 0.")
                st.stop()

            area_m2 = (p["width_mm"] * p["height_mm"]) / 1_000_000.0
            perimeter_m = 2 * (p["width_mm"] + p["height_mm"]) / 1000.0
            base_positions.append({
                **p,
                "area_m2": area_m2,
                "perimeter_m": perimeter_m,
            })

        lambr_positions = []
        for p in lambr_positions_inputs:
            if p["width_mm"] > 0 and p["height_mm"] > 0:
                area_m2 = (p["width_mm"] * p["height_mm"]) / 1_000_000.0
                perimeter_m = 2 * (p["width_mm"] + p["height_mm"]) / 1000.0
                lambr_positions.append({
                    **p,
                    "area_m2": area_m2,
                    "perimeter_m": perimeter_m,
                })

        # Подстановка размеров створок
        try:
            sashes_count = int(sashes)
        except ValueError:
            sashes_count = 1

        if sashes_count >= 1:
            for p in base_positions:
                if p["sash_width_mm"] <= 0:
                    p["sash_width_mm"] = p["width_mm"]
                if p["sash_height_mm"] <= 0:
                    p["sash_height_mm"] = p["height_mm"]

        # Тамбур: внутренние секции
        tambour_sections = []
        tambour_door_count = 0
        if product_type == "Тамбур" and base_positions:
            first_pos = base_positions[0]
            internal = first_pos.get("tambour_sections", [])
            for sec in internal:
                if sec["width_mm"] <= 0 or sec["height_mm"] <= 0:
                    continue
                area_m2 = (sec["width_mm"] * sec["height_mm"]) / 1_000_000.0
                perimeter_m = 2 * (sec["width_mm"] + sec["height_mm"]) / 1000.0
                tambour_sections.append({
                    "kind": sec["kind"],
                    "width_mm": sec["width_mm"],
                    "height_mm": sec["height_mm"],
                    "Nwin": sec["Nwin"],
                    "area_m2": area_m2,
                    "perimeter_m": perimeter_m,
                })
                if sec["kind"] == "door":
                    tambour_door_count += sec["Nwin"]

        order = {
            "order_number": order_number.strip(),
            "product_type": product_type,
            "product_view": product_view,
            "sashes": sashes,
            "profile_system": profile_system,
            "glass_type": glass_type,
            "filling_mode": filling_mode,
            "toning": toning,
            "assembly": assembly,
            "montage": montage,
            "handle_type": handle_type,
            "door_closer": door_closer,
            "tambour_sections": tambour_sections,
        }

        # --- ЗАПРОСЫ ---
        rows_for_form = []

        for idx, p in enumerate(base_positions, start=1):
            rows_for_form.append([
                order["order_number"],
                idx,
                order["product_type"],
                order["product_view"],
                order["sashes"],
                order["profile_system"],
                order["glass_type"],
                order["filling_mode"],
                p["width_mm"],
                p["height_mm"],
                p.get("left_mm", 0.0),
                p.get("center_mm", 0.0),
                p.get("right_mm", 0.0),
                p.get("top_mm", 0.0),
                p.get("sash_width_mm", p["width_mm"]),
                p.get("sash_height_mm", p["height_mm"]),
                p["Nwin"],
                order["toning"],
                order["assembly"],
                order["montage"],
                order["handle_type"],
                order["door_closer"],
            ])

        for idx, p in enumerate(lambr_positions, start=len(rows_for_form) + 1):
            rows_for_form.append([
                order["order_number"],
                idx,
                order["product_type"],
                f"Панель {filling_mode}",
                order["sashes"],
                order["profile_system"],
                order["glass_type"],
                order["filling_mode"],
                p["width_mm"],
                p["height_mm"],
                p.get("left_mm", 0.0),
                p.get("center_mm", 0.0),
                p.get("right_mm", 0.0),
                p.get("top_mm", 0.0),
                p.get("sash_width_mm", p["width_mm"]),
                p.get("sash_height_mm", p["height_mm"]),
                p["Nwin"],
                order["toning"],
                order["assembly"],
                order["montage"],
                order["handle_type"],
                order["door_closer"],
            ])

        for row in rows_for_form:
            excel.append_form_row(row)

        # --- ГАБАРИТЫ ---
        gab_calc = GabaritCalculator(excel)
        if product_type == "Тамбур":
            gabarit_positions = base_positions
        else:
            gabarit_positions = base_positions + lambr_positions

        gabarit_rows, total_area_gab = gab_calc.calculate(order, gabarit_positions)

        # --- МАТЕРИАЛЫ ---
        mat_calc = MaterialCalculator(excel)
        if product_type == "Тамбур":
            positions_for_materials = tambour_sections
        else:
            positions_for_materials = base_positions + lambr_positions

        material_rows, material_total, total_area_mat = mat_calc.calculate(
            order, positions_for_materials, selected_duplicates
        )

        # --- Площади ---
        if filling_mode == "Стеклопакет":
            if product_type == "Тамбур":
                total_area_glass = sum(
                    s["area_m2"] * s["Nwin"] for s in tambour_sections
                )
            else:
                total_area_glass = sum(
                    p["area_m2"] * p["Nwin"] for p in (base_positions + lambr_positions)
                )
        else:
            total_area_glass = 0.0

        if product_type == "Тамбур":
            total_area_all = sum(s["area_m2"] * s["Nwin"] for s in tambour_sections)
        else:
            total_area_all = sum(
                p["area_m2"] * p["Nwin"] for p in (base_positions + lambr_positions)
            )

        total_area = total_area_all

        # --- Финальный расчёт ---
        final_calc = FinalCalculator(excel)
        final_rows, total_sum, ensure_sum = final_calc.calculate(
            order,
            total_area_all=total_area_all,
            total_area_glass=total_area_glass,
            material_total=material_total,
            tambour_door_count=tambour_door_count,
        )

        st.success("Расчёт выполнен. Результаты ниже (служебная информация).")

        tab1, tab2, tab3 = st.tabs(["Габариты", "Материалы", "Итоговый расчет"])

        with tab1:
            st.subheader("Расчет по габаритам")
            if gabarit_rows:
                gab_disp = [
                    {"Тип элемента": t, "Фактическое значение": v}
                    for t, v in gabarit_rows
                ]
                st.dataframe(gab_disp, use_container_width=True)
            st.write(f"Общая площадь (служебная): **{total_area_gab:.3f} м²**")
            st.write(f"Рабочая площадь для расчетов: **{total_area:.3f} м²**")

        with tab2:
            st.subheader("Расчёт материалов")
            if material_rows:
                mat_disp = []
                for r in material_rows:
                    mat_disp.append({
                        "Тип изделия": r[0],
                        "Система профиля": r[1],
                        "Тип элемента": r[2],
                        "Артикул": r[3],
                        "Товар": r[4],
                        "Ед.": r[5],
                        "Цена за ед.": round(r[6], 2),
                        "Ед. факт. расхода": r[7],
                        "Кол-во факт. расхода": round(r[8], 3),
                        "Норма к упаковке": r[9],
                        "Ед. к отгрузке": r[10],
                        "Кол-во к отгрузке": round(r[11], 3),
                        "Сумма": round(r[12], 2),
                    })
                st.dataframe(mat_disp, use_container_width=True)
            st.write(f"Итого по материалам: **{material_total:.2f}**")

        with tab3:
            st.subheader("Итоговый расчет с монтажом (служебно)")
            if final_rows:
                fin_disp = []
                for name, price, unit, total_val in final_rows:
                    fin_disp.append({
                        "Наименование услуг": name,
                        "Стоимость за м²": price if isinstance(price, str) else round(price, 2),
                        "Ед": unit,
                        "Итого": total_val if isinstance(total_val, str) else round(total_val, 2),
                    })
                st.dataframe(fin_disp, use_container_width=True)
            st.write(f"Обеспечение (60%): **{ensure_sum:.2f}**")
            st.write(f"ИТОГО к оплате: **{total_sum:.2f}**")

        # --- Коммерческий Excel ---
        smeta_bytes = build_smeta_workbook(
            order=order,
            base_positions=base_positions,
            lambr_positions=lambr_positions,
            total_area=total_area,
            total_sum=total_sum,
        )
        default_name = f"Коммерческое_предложение_Заказ_{order['order_number']}.xlsx"
        st.download_button(
            "⬇️ Скачать коммерческое предложение в Excel",
            data=smeta_bytes,
            file_name=default_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
