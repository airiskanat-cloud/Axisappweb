# -*- coding: utf-8 -*-
"""
Axis Pro GF — Streamlit приложение (обновлённый полный файл)
Внимание: заменяет логику формы для тамбуров, переносит выбор заполнения в левую часть,
защищает загрузку Excel (BadZipFile) и учитывает стеклопакет по glass_type + по секциям.
"""

import math
import os
import sys
import zipfile
from io import BytesIO

import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.drawing.image import Image as XLImage

# =========================
# Константы и пути
# =========================

def resource_path(relative_path: str) -> str:
    """Возвращает корректный путь к файлу (поддержка PyInstaller)."""
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
    "Режим заполнения",  # Ламбри / Сэндвич
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

# =========================
# Утилиты
# =========================

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

# =========================
# Excel client с проверкой
# =========================

def is_probably_xlsx(path: str) -> bool:
    # базовая проверка: файл существует, не пустой, и можно открыть как zip
    if not os.path.exists(path) or not os.path.isfile(path):
        return False
    try:
        if os.path.getsize(path) < 200:  # слишком маленький — подозрительно
            return False
    except Exception:
        pass
    try:
        with zipfile.ZipFile(path, "r") as z:
            z.namelist()
        return True
    except Exception:
        return False


class ExcelClient:
    def __init__(self, filename: str):
        self.filename = filename
        # если файла нет или он невалидный — пытаемся создать шаблон
        if not is_probably_xlsx(self.filename):
            try:
                # Если файл существует, но невалиден — переименуем его как резервную копию
                if os.path.exists(self.filename):
                    backup = self.filename + ".bad." + str(int(os.path.getmtime(self.filename)))
                    try:
                        os.rename(self.filename, backup)
                        print(f"Renamed invalid excel to backup: {backup}")
                    except Exception:
                        print("Не удалось переименовать повреждённый файл; он будет перезаписан.")
                wb = Workbook()
                # создаём несколько служебных листов для корректной структуры
                if "Sheet" in wb.sheetnames:
                    ws0 = wb["Sheet"]
                    wb.remove(ws0)
                wb.create_sheet(SHEET_FORM)
                wb.create_sheet(SHEET_REF1)
                wb.create_sheet(SHEET_REF2)
                wb.create_sheet(SHEET_REF3)
                wb.create_sheet(SHEET_USERS)
                wb.save(self.filename)
                print(f"Создан новый шаблон Excel: {self.filename}")
            except Exception as e:
                print(f"Ошибка при создании шаблона Excel: {e}")
        self.load()

    def load(self):
        try:
            self.wb = load_workbook(self.filename, data_only=True)
        except zipfile.BadZipFile:
            # Создаём новый рабочий файл, чтобы приложение не падало
            print(f"BadZipFile: {self.filename} is not a valid xlsx. Recreating workbook.")
            wb = Workbook()
            if "Sheet" in wb.sheetnames:
                ws0 = wb["Sheet"]
                wb.remove(ws0)
            wb.create_sheet(SHEET_FORM)
            wb.create_sheet(SHEET_REF1)
            wb.create_sheet(SHEET_REF2)
            wb.create_sheet(SHEET_REF3)
            wb.create_sheet(SHEET_USERS)
            wb.save(self.filename)
            self.wb = load_workbook(self.filename, data_only=True)
        except Exception as e:
            print(f"Ошибка при загрузке Excel: {e}")
            # чтобы self.wb существовал
            self.wb = Workbook()

    def save(self):
        try:
            self.wb.save(self.filename)
        except Exception as e:
            print(f"Ошибка при сохранении Excel: {e}")

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
        # удаляем содержимое
        try:
            ws.delete_rows(1, ws.max_row or 1)
        except Exception:
            # на всякий случай переприсваиваем новый лист
            if sheet_name in self.wb.sheetnames:
                del self.wb[sheet_name]
            ws = self.wb.create_sheet(sheet_name)
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

# =========================
# Пользователи
# =========================

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

# =========================
# Габариты / материалы / итог — классы (используют секции)
# =========================

class GabaritCalculator:
    HEADER = ["Тип элемента", "Фактическое значение"]

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

    def calculate(self, order: dict, sections: list):
        ref_rows = self.excel.read_records(SHEET_REF3)
        if not ref_rows:
            return [], 0.0

        try:
            nsash = int(order.get("sashes", "1"))
        except Exception:
            nsash = 1
        n_sash_active = 1 if nsash >= 1 else 0
        n_sash_passive = max(nsash - 1, 0)
        hinges_per_sash = 3

        total_area = sum(s["area_m2"] * s["Nwin"] for s in sections)
        gabarit_values = []

        for row in ref_rows:
            type_elem = get_field(row, "тип элемент", "")
            formula = get_field(row, "формула_python", "")
            if not type_elem or not formula:
                continue

            total_value = 0.0

            for s in sections:
                width = s.get("width_mm", 0.0)
                height = s.get("height_mm", 0.0)
                left = s.get("left_mm", 0.0)
                center = s.get("center_mm", 0.0)
                right = s.get("right_mm", 0.0)
                top = s.get("top_mm", 0.0)
                sash_w = s.get("sash_width_mm", width)
                sash_h = s.get("sash_height_mm", height)
                area = s["area_m2"]
                perimeter = s["perimeter_m"]
                qty = s["Nwin"]

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

    def calculate(self, order: dict, sections: list, selected_duplicates: dict):
        ref_rows = self.excel.read_records(SHEET_REF1)
        total_area = sum(s["area_m2"] * s["Nwin"] for s in sections)
        if not ref_rows:
            return [], 0.0, total_area

        try:
            nsash = int(order.get("sashes", "1"))
        except Exception:
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

            if row_type:
                if str(row_type).strip().lower() != order["product_type"].strip().lower():
                    continue

            if row_profile:
                if str(row_profile).strip().lower() != order["profile_system"].strip().lower():
                    continue

            # фильтр по дублям
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

            for s in sections:
                width = s.get("width_mm", 0.0)
                height = s.get("height_mm", 0.0)
                left = s.get("left_mm", 0.0)
                center = s.get("center_mm", 0.0)
                right = s.get("right_mm", 0.0)
                top = s.get("top_mm", 0.0)
                sash_w = s.get("sash_width_mm", width)
                sash_h = s.get("sash_height_mm", height)
                area = s["area_m2"]
                perimeter = s["perimeter_m"]
                qty = s["Nwin"]

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


class FinalCalculator:
    HEADER = ["Наименование услуг", "Стоимость за м²", "Ед", "Итого"]

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def calculate(self,
                  order: dict,
                  total_area_all: float,
                  total_area_glass: float,
                  material_total: float,
                  doors_count: int = 0):
        ref_rows = self.excel.read_records(SHEET_REF2)

        glass_type = order["glass_type"]
        toning = order["toning"]
        assembly = order["assembly"]
        montage = order["montage"]
        handle_type = order["handle_type"]
        door_closer = order["door_closer"]

        selected = None
        for row in ref_rows:
            row_glass = str(get_field(row, "тип стеклопак", "") or "").strip()
            row_handle_type = str(get_field(row, "ручк", "") or "").strip()

            if row_glass and row_glass != glass_type:
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

        # Стеклопакет — считаем по фактической площади стекла (передано в total_area_glass)
        if total_area_glass > 0:
            glass_sum = total_area_glass * price_glass
        else:
            glass_sum = 0.0
            price_glass = 0.0
        rows.append(["Стеклопакет", price_glass, "за м²", glass_sum])

        # Тонировка
        if toning == "Есть" and total_area_glass > 0:
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
            handles_qty = max(doors_count, 1) if order["product_type"].lower() == "тамбур" else 1
            handles_sum = price_handles * handles_qty
        rows.append(["Ручки", price_handles, "шт.", handles_sum])

        # Доводчик
        closer_sum = 0.0
        if door_closer == "Есть":
            closer_qty = max(doors_count, 1) if order["product_type"].lower() == "тамбур" else 1
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

# =========================
# Экспорт коммерческого предложения
# =========================

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
    ws.cell(row=current_row, column=1, value=f"Тип заполнения (панели): {order['filling_mode']}")
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
        ws.cell(row=current_row, column=1, value="Тамбур (секции):")
        current_row += 1

        for idx, s in enumerate(order.get("sections", []), start=1):
            kind = s.get("kind", "section")
            w = s["width_mm"]
            h = s["height_mm"]
            q = s["Nwin"]
            filling = s.get("filling", "")
            ws.cell(
                row=current_row,
                column=1,
                value=f"  Секция {idx} ({kind}): {w} × {h} мм, N = {q}, заполнение = {filling}"
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

# =========================
# Streamlit UI: main
# =========================

def main():
    st.set_page_config(page_title="Axis Pro GF • Калькулятор", layout="wide")

    excel_ok = is_probably_xlsx(EXCEL_FILE)
    excel = ExcelClient(EXCEL_FILE)  # ExcelClient сам создаст шаблон, если надо

    # Авторизация
    user = login_form(excel)
    if not user:
        st.stop()

    st.title("📘 Калькулятор алюминиевых изделий (Axis Pro GF)")
    st.info(f"Пользователь: **{user['login']}**")

    # Загружаем Справочник-2, чтобы взять типы ручек и типы стеклопакетов
    ref2_records = excel.read_records(SHEET_REF2)
    handle_types_set = set()
    glass_types_set = set()
    for row in ref2_records:
        hname = get_field(row, "ручк", "")
        if hname:
            handle_types_set.add(str(hname).strip())
        gtype = get_field(row, "тип стеклопак", "")
        if gtype:
            glass_types_set.add(str(gtype).strip())
    handle_types = sorted(list(handle_types_set)) if handle_types_set else [""]
    glass_types = sorted(list(glass_types_set)) if glass_types_set else ["двойной"]

    # ---------- Сайдбар: общие данные ----------
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
            "Тип стеклопакета (цена берётся из справочника-2)",
            glass_types
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

    # ---------- Основная часть: две колонки ----------
    col_left, col_right = st.columns([2, 1])

    # Переносим выбор заполнения (режим панели: Ламбри / Сэндвич) в левую колонку
    with col_left:
        st.header("Настройки заполнения и позиции")
        filling_mode = st.radio(
            "Режим заполнения (панели)",
            ["Ламбри", "Сэндвич"],
            index=0
        )
        st.caption("Стеклопакет убран из общего режима — считается по секциям и по типу стеклопакета.")

    # Справа оставляем вспомогательную информацию
    with col_right:
        st.header("Информация")
        st.info("Заполнения панелей: Ламбри / Сэндвич. "
                "Стеклопакет рассчитывается отдельно на основании выбранного типа стеклопакета и секций, помеченных как стекло.")
        if not excel_ok:
            st.warning("Внимание: исходный Excel-файл либо отсутствует, либо был невалидным. Создан шаблон. Проверьте данные в справочниках.")

    # ---------- Ввод позиций (лево) ----------
    lambr_positions_inputs = []
    base_positions_inputs = []  # для не-тамбура
    sections_inputs = []  # для тамбура — список секций (door/panel)

    with col_left:
        st.subheader("Позиции (габариты изделий)")

        # для всех типов даём возможность задать число позиций (для тамбура также)
        positions_count = st.number_input(
            "Количество позиций",
            min_value=1,
            max_value=10,
            value=1,
            step=1,
            help="Для Тамбура здесь можно задать >1 позиций, каждая позиция — рамная единица"
        )

        # каждая позиция — карточка
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

            # Если не тамбур — обычная позиция (рамная единица)
            if product_type != "Тамбур":
                base_positions_inputs.append({
                    "width_mm": width_mm,
                    "height_mm": height_mm,
                    "left_mm": left_mm,
                    "center_mm": center_mm,
                    "right_mm": right_mm,
                    "top_mm": top_mm,
                    "sash_width_mm": sash_width_mm if sash_width_mm > 0 else width_mm,
                    "sash_height_mm": sash_height_mm if sash_height_mm > 0 else height_mm,
                    "Nwin": Nwin,
                    "filling": filling_mode  # общая панельная логика
                })
            else:
                # Тамбур: у каждой позиции внутри могут быть несколько дверей/панелей
                with st.expander(f"Параметры Тамбура — Позиция {i+1} (двери / глухие секции)", expanded=False):
                    # количество дверей для этой позиции
                    door_count = st.number_input(
                        f"Количество дверей (поз. {i+1})",
                        min_value=0,
                        value=1,
                        step=1,
                        key=f"tamb_dir_count_{i}"
                    )
                    doors_local = []
                    for d in range(int(door_count)):
                        st.markdown(f"**Дверь {d+1} (поз. {i+1})**")
                        d1, d2, d3 = st.columns(3)
                        dw = d1.number_input(
                            f"Ширина двери {d+1} (поз. {i+1}), мм",
                            min_value=0.0,
                            step=10.0,
                            key=f"door_w_{i}_{d}"
                        )
                        dh = d2.number_input(
                            f"Высота двери {d+1} (поз. {i+1}), мм",
                            min_value=0.0,
                            step=10.0,
                            key=f"door_h_{i}_{d}"
                        )
                        dq = d3.number_input(
                            f"N (дверь {d+1} поз. {i+1})",
                            min_value=1,
                            value=1,
                            step=1,
                            key=f"door_q_{i}_{d}"
                        )
                        doors_local.append({
                            "kind": "door",
                            "width_mm": dw,
                            "height_mm": dh,
                            "Nwin": dq,
                            "left_mm": 0.0,
                            "center_mm": 0.0,
                            "right_mm": 0.0,
                            "top_mm": 0.0,
                            "sash_width_mm": dw,
                            "sash_height_mm": dh,
                            "filling": "Стеклопакет"  # обычно двери имеют стекло, но можно изменить ниже
                        })

                    panel_count = st.number_input(
                        f"Количество глухих секций (поз. {i+1})",
                        min_value=0,
                        value=1,
                        step=1,
                        key=f"tamb_panel_count_{i}"
                    )
                    panels_local = []
                    for p_idx in range(int(panel_count)):
                        st.markdown(f"**Глухая секция {p_idx+1} (поз. {i+1})**")
                        p1, p2, p3 = st.columns(3)
                        pw = p1.number_input(
                            f"Ширина глухой секции {p_idx+1} (поз. {i+1}), мм",
                            min_value=0.0,
                            step=10.0,
                            key=f"panel_w_{i}_{p_idx}"
                        )
                        ph = p2.number_input(
                            f"Высота глухой секции {p_idx+1} (поз. {i+1}), мм",
                            min_value=0.0,
                            step=10.0,
                            key=f"panel_h_{i}_{p_idx}"
                        )
                        pq = p3.number_input(
                            f"N (секция {p_idx+1} поз. {i+1})",
                            min_value=1,
                            value=1,
                            step=1,
                            key=f"panel_q_{i}_{p_idx}"
                        )
                        pf = st.selectbox(
                            f"Заполнение глухой секции {p_idx+1} (поз. {i+1})",
                            options=["Стеклопакет", "Ламбри", "Сэндвич"],
                            index=0,
                            key=f"panel_fill_{i}_{p_idx}"
                        )
                        panels_local.append({
                            "kind": "panel",
                            "width_mm": pw,
                            "height_mm": ph,
                            "Nwin": pq,
                            "left_mm": 0.0,
                            "center_mm": 0.0,
                            "right_mm": 0.0,
                            "top_mm": 0.0,
                            "sash_width_mm": pw,
                            "sash_height_mm": ph,
                            "filling": pf
                        })

                    # Сохраняем секции этой позиции
                    # Позиция как рамная единица (может понадобиться для ZAPROS)
                    base_pos_for_this = {
                        "width_mm": width_mm,
                        "height_mm": height_mm,
                        "left_mm": left_mm,
                        "center_mm": center_mm,
                        "right_mm": right_mm,
                        "top_mm": top_mm,
                        "sash_width_mm": sash_width_mm if sash_width_mm > 0 else width_mm,
                        "sash_height_mm": sash_height_mm if sash_height_mm > 0 else height_mm,
                        "Nwin": Nwin
                    }
                    base_positions_inputs.append(base_pos_for_this)

                    # добавляем секции в общий список
                    for d in doors_local:
                        sections_inputs.append(d)
                    for psec in panels_local:
                        sections_inputs.append(psec)

    # ---------- Панели Ламбри / Сэндвич для не-тамбура ----------
    if product_type != "Тамбур":
        if filling_mode in ("Ламбри", "Сэндвич"):
            with col_left:
                st.subheader(f"Панели {filling_mode}")
                panel_count_ls = st.number_input(
                    f"Количество панелей ({filling_mode})",
                    min_value=0,
                    value=0,
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
                        "filling": filling_mode
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
        # валидации
        if not order_number.strip():
            st.error("Введите номер заказа в левой панели.")
            st.stop()

        # Собираем базовые позиции (для не-тамбура они уже в base_positions_inputs)
        base_positions = []
        for p in base_positions_inputs:
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

        # Если Тамбур — sections_inputs уже заполнён (двери и глухие секции)
        # Если не тамбур — хотим считать стекло только если filling == "Стеклопакет"
        sections = []
        if product_type == "Тамбур":
            # sections_inputs уже заполнены в UI. Убедимся в валидности и посчитаем площади
            for s in sections_inputs:
                if s["width_mm"] <= 0 or s["height_mm"] <= 0:
                    st.warning("Одна из секций тамбура имеет 0 ширину или высоту и будет пропущена.")
                    continue
                area_m2 = (s["width_mm"] * s["height_mm"]) / 1_000_000.0
                perimeter_m = 2 * (s["width_mm"] + s["height_mm"]) / 1000.0
                sections.append({
                    **s,
                    "area_m2": area_m2,
                    "perimeter_m": perimeter_m
                })
        else:
            # Для не-тамбура — рассматриваем base_positions + (возможно) ламбри панели.
            # Базовые позиции — это рамные единицы; считаем их как "секции" с filling == filling_mode
            for p in base_positions:
                sections.append({
                    **p,
                    "area_m2": p["area_m2"],
                    "perimeter_m": p["perimeter_m"],
                    "filling": p.get("filling", filling_mode)
                })
            # панели lambr/sandwich — отдельные секции (их filling = filling_mode)
            for p in lambr_positions:
                sections.append({
                    **p,
                    "area_m2": p["area_m2"],
                    "perimeter_m": p["perimeter_m"],
                    "filling": p.get("filling", filling_mode)
                })

        # Подстановка размеров створок: если sashes_count >= 1, заполняем пустые створки
        try:
            sashes_count = int(sashes)
        except Exception:
            sashes_count = 1

        if sashes_count >= 1:
            for s in sections:
                if s.get("sash_width_mm", 0) <= 0:
                    s["sash_width_mm"] = s["width_mm"]
                if s.get("sash_height_mm", 0) <= 0:
                    s["sash_height_mm"] = s["height_mm"]

        # --- Сохранение в ЗАПРОСЫ (служебно) ---
        rows_for_form = []

        # для не-тамбура: записываем позиции и панели
        pos_index = 1
        if product_type != "Тамбур":
            for p in base_positions:
                rows_for_form.append([
                    order_number,
                    pos_index,
                    product_type,
                    product_view,
                    sashes,
                    profile_system,
                    glass_type,
                    filling_mode,
                    p["width_mm"],
                    p["height_mm"],
                    p.get("left_mm", 0.0),
                    p.get("center_mm", 0.0),
                    p.get("right_mm", 0.0),
                    p.get("top_mm", 0.0),
                    p.get("sash_width_mm", p["width_mm"]),
                    p.get("sash_height_mm", p["height_mm"]),
                    p["Nwin"],
                    toning,
                    assembly,
                    montage,
                    handle_type,
                    door_closer,
                ])
                pos_index += 1

            for p in lambr_positions:
                rows_for_form.append([
                    order_number,
                    pos_index,
                    product_type,
                    f"Панель {filling_mode}",
                    sashes,
                    profile_system,
                    glass_type,
                    filling_mode,
                    p["width_mm"],
                    p["height_mm"],
                    p.get("left_mm", 0.0),
                    p.get("center_mm", 0.0),
                    p.get("right_mm", 0.0),
                    p.get("top_mm", 0.0),
                    p.get("sash_width_mm", p["width_mm"]),
                    p.get("sash_height_mm", p["height_mm"]),
                    p["Nwin"],
                    toning,
                    assembly,
                    montage,
                    handle_type,
                    door_closer,
                ])
                pos_index += 1
        else:
            # Тамбур: сохраняем общую позицию(и), но в ком. предложении укажем секции подробно
            for p in base_positions:
                rows_for_form.append([
                    order_number,
                    pos_index,
                    product_type,
                    product_view,
                    sashes,
                    profile_system,
                    glass_type,
                    "Тамбур",
                    p["width_mm"],
                    p["height_mm"],
                    p.get("left_mm", 0.0),
                    p.get("center_mm", 0.0),
                    p.get("right_mm", 0.0),
                    p.get("top_mm", 0.0),
                    p.get("sash_width_mm", p["width_mm"]),
                    p.get("sash_height_mm", p["height_mm"]),
                    p["Nwin"],
                    toning,
                    assembly,
                    montage,
                    handle_type,
                    door_closer,
                ])
                pos_index += 1

        for row in rows_for_form:
            excel.append_form_row(row)

        # --- Расчёты: габариты, материалы, финал ---
        gab_calc = GabaritCalculator(excel)
        gabarit_rows, total_area_gab = gab_calc.calculate(
            {"sashes": sashes, "product_type": product_type},
            sections
        )

        mat_calc = MaterialCalculator(excel)
        material_rows, material_total, total_area_mat = mat_calc.calculate(
            {"product_type": product_type, "profile_system": profile_system, "sashes": sashes},
            sections,
            selected_duplicates
        )

        # --- Площади ---
        # total_area_glass: считаем площадь только тех секций, где filling == "Стеклопакет"
        total_area_glass = sum(s["area_m2"] * s["Nwin"] for s in sections if s.get("filling") == "Стеклопакет")
        total_area_all = sum(s["area_m2"] * s["Nwin"] for s in sections)

        # количество дверей (для ручек/доводчиков)
        doors_count = sum(s["Nwin"] for s in sections if s.get("kind") == "door")

        # Финальный расчёт
        final_calc = FinalCalculator(excel)
        final_rows, total_sum, ensure_sum = final_calc.calculate(
            {
                "product_type": product_type,
                "glass_type": glass_type,
                "handle_type": handle_type,
                "door_closer": door_closer
            },
            total_area_all=total_area_all,
            total_area_glass=total_area_glass,
            material_total=material_total,
            doors_count=doors_count,
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
            st.write(f"Рабочая площадь для расчетов: **{total_area_all:.3f} м²**")
            st.write(f"Площадь стекла (по секциям): **{total_area_glass:.3f} м²**")

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
            order={
                "order_number": order_number,
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
                "sections": sections
            },
            base_positions=base_positions,
            lambr_positions=lambr_positions,
            total_area=total_area_all,
            total_sum=total_sum,
        )
        default_name = f"Коммерческое_предложение_Заказ_{order_number}.xlsx"
        st.download_button(
            "⬇️ Скачать коммерческое предложение в Excel",
            data=smeta_bytes,
            file_name=default_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
