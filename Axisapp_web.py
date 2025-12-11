# axis_app_clean.py
# Единый, упрощённый и укреплённый вариант приложения расчёта материалов.
# - безопасное приведение типов
# - безопасный eval для пользовательских формул
# - минимальная встроенная логика "gabarit" для совместимости формул
# - основной фокус: корректный расчёт материалов и итоговый расчёт
# Как использовать: положи файл рядом с axis_pro_gf.xlsx и запусти через streamlit:
#   streamlit run axis_app_clean.py

import os
import sys
import math
import json
import shutil
import logging
import ast
import operator as op
import re
from io import BytesIO

# Попытка импортировать streamlit; если нет, приложение не сможет показать UI,
# но код можно изучать/править.
try:
    import streamlit as st
except Exception:
    st = None

# Для работы с Excel
try:
    from openpyxl import load_workbook
    from openpyxl.workbook import Workbook
    from openpyxl.drawing.image import Image as XLImage
except Exception:
    load_workbook = None
    Workbook = None
    XLImage = None

# Логгер
logger = logging.getLogger("axis_app_clean")
if not logger.handlers:
    ch = logging.StreamHandler()
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    ch.setFormatter(formatter)
    logger.addHandler(ch)
logger.setLevel(logging.INFO)

# =========================
# Настройки/константы
# =========================
DATA_DIR = os.getenv("AXIS_DATA_DIR", os.path.join(os.path.expanduser("~"), ".axis_app_data"))
os.makedirs(DATA_DIR, exist_ok=True)

TEMPLATE_EXCEL_NAME = "axis_pro_gf.xlsx"
EXCEL_FILE = os.path.join(DATA_DIR, TEMPLATE_EXCEL_NAME)
BUNDLED_TEMPLATE = os.path.join(os.path.dirname(__file__), TEMPLATE_EXCEL_NAME) if "__file__" in globals() else TEMPLATE_EXCEL_NAME
SESSION_FILE = os.path.join(DATA_DIR, "session_user.json")

# Имена листов (как в исходнике)
SHEET_REF1 = "СПРАВОЧНИК -1"   # материалы
SHEET_REF2 = "СПРАВОЧНИК -2"   # цены/параметры
SHEET_REF3 = "СПРАВОЧНИК -3"   # (будет не обязателен)
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"  # можно не использовать
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

FORM_HEADER = [
    "Номер заказа", "№ позиции",
    "Тип изделия", "Вид изделия", "Створки",
    "Профильная система",
    "Тип стеклопакета",
    "Режим заполнения",
    "Ширина, мм", "Высота, мм",
    "LEFT, мм", "CENTER, мм", "RIGHT, мм", "TOP, мм",
    "Ширина створки, мм", "Высота створки, мм",
    "Кол-во Nwin",
    "Тонировка", "Сборка", "Монтаж",
    "Тип ручек", "Доводчик"
]

# =========================
# Утилиты: приведение типов / строки
# =========================

def resource_path(relative_path: str) -> str:
    """Надёжный путь для запуска из pyinstaller или из исходников."""
    try:
        if hasattr(sys, "_MEIPASS"):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(os.path.dirname(__file__))
    except Exception:
        base_path = os.getcwd()
    return os.path.join(base_path, relative_path)

def normalize_key(k):
    if k is None:
        return None
    s = str(k)
    s = s.replace("\xa0", " ")
    s = " ".join(s.split())
    return s.strip()

def _clean_cell_val(v):
    if v is None:
        return ""
    s = str(v)
    s = s.replace("\xa0", " ").strip()
    return s

def safe_float(v, default=0.0):
    """Robust conversion to float: handles None, ints, floats, strings with spaces/NBSP/commas."""
    if v is None:
        return float(default)
    if isinstance(v, (int, float)):
        try:
            return float(v)
        except Exception:
            return float(default)
    try:
        s = str(v).strip()
        s = re.sub(r'[\u00A0\s]', '', s)  # remove spaces and NBSP
        s = s.replace(',', '.')
        if s == '':
            return float(default)
        return float(s)
    except Exception:
        return float(default)

def safe_int(v, default=0):
    try:
        return int(safe_float(v, default))
    except Exception:
        return int(default)

def is_positive_number(v):
    try:
        return safe_float(v) > 0
    except Exception:
        return False

# =========================
# Безопасный eval для формул из Excel
# =========================
_allowed_ops = {
    ast.Add: op.add,
    ast.Sub: op.sub,
    ast.Mult: op.mul,
    ast.Div: op.truediv,
    ast.Pow: op.pow,
    ast.USub: op.neg,
    ast.UAdd: op.pos,
    ast.Mod: op.mod,
    ast.FloorDiv: op.floordiv,
    ast.Lt: op.lt,
    ast.Gt: op.gt,
    ast.LtE: op.le,
    ast.GtE: op.ge,
    ast.Eq: op.eq,
    ast.NotEq: op.ne,
    ast.And: lambda a,b: a and b,
    ast.Or:  lambda a,b: a or b,
}

def _eval_ast(node, names):
    if isinstance(node, ast.Expression):
        return _eval_ast(node.body, names)

    if isinstance(node, ast.Constant):
        return node.value

    if hasattr(ast, "Num") and isinstance(node, ast.Num):
        return node.n

    if isinstance(node, ast.UnaryOp):
        val = _eval_ast(node.operand, names)
        fn = _allowed_ops.get(type(node.op))
        if fn:
            return fn(val)

    if isinstance(node, ast.BinOp):
        left = _eval_ast(node.left, names)
        right = _eval_ast(node.right, names)
        fn = _allowed_ops.get(type(node.op))
        if fn:
            return fn(left, right)

    if isinstance(node, ast.Name):
        if node.id in names:
            return names[node.id]
        raise ValueError(f"Недопустимое имя '{node.id}'")

    if isinstance(node, ast.Call):
        func = node.func
        # math.* calls allowed
        if isinstance(func, ast.Attribute) and isinstance(func.value, ast.Name) and func.value.id == "math":
            fname = func.attr
            if hasattr(math, fname):
                args = [_eval_ast(a, names) for a in node.args]
                return getattr(math, fname)(*args)
        # allow min/max
        if isinstance(func, ast.Name) and func.id in ("max", "min"):
            args = [_eval_ast(a, names) for a in node.args]
            return globals()[func.id](*args)
        raise ValueError("Разрешены только math.*, max, min")

    if isinstance(node, ast.Compare):
        if len(node.ops) != 1:
            raise ValueError("Сложные сравнения запрещены")
        left = _eval_ast(node.left, names)
        right = _eval_ast(node.comparators[0], names)
        fn = _allowed_ops.get(type(node.ops[0]))
        if fn:
            return fn(left, right)

    raise ValueError(f"Недопустимый элемент формулы: {type(node).__name__}")

def safe_eval_formula(formula: str, context: dict) -> float:
    """
    Выполняет ограниченный безопасный eval выражения (только арифметика, math.*, min/max).
    При ошибке логируем DEBUG и возвращаем 0.0.
    """
    formula = (formula or "").strip()
    if not formula:
        return 0.0
    names = {**context, "math": math, "min": min, "max": max}
    try:
        node = ast.parse(formula, mode="eval")
        return float(_eval_ast(node, names))
    except Exception as e:
        logger.debug("safe_eval_formula failed for formula=%r ctx=%s error=%s", formula, context, e, exc_info=True)
        return 0.0

# =========================
# ExcelClient: чтение/запись и минимальный бэкап
# =========================
class ExcelClient:
    def __init__(self, filename: str):
        self.filename = filename
        if not os.path.exists(self.filename):
            self._create_template_if_bundled()
        self.load()

    def _create_template_if_bundled(self):
        # если есть bundled template рядом со скриптом, скопируем
        try:
            if os.path.exists(BUNDLED_TEMPLATE) and not os.path.exists(self.filename):
                shutil.copyfile(BUNDLED_TEMPLATE, self.filename)
                logger.info("Copied bundled template %s -> %s", BUNDLED_TEMPLATE, self.filename)
        except Exception:
            logger.exception("Error copying bundled template")
        # если всё равно нет — создадим минимальный пустой файл
        if not os.path.exists(self.filename) and Workbook is not None:
            wb = Workbook()
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]
            wb.create_sheet(SHEET_FORM)
            wb.create_sheet(SHEET_REF1)
            wb.create_sheet(SHEET_REF2)
            wb.create_sheet(SHEET_REF3)
            wb.create_sheet(SHEET_USERS)
            wb.save(self.filename)

    def load(self):
        if load_workbook is None:
            raise RuntimeError("openpyxl not available in this environment")
        try:
            self.wb = load_workbook(self.filename, data_only=True)
        except Exception as e:
            logger.exception("Ошибка при загрузке Excel, пытаюсь восстановить: %s", e)
            try:
                # бэкап
                if os.path.exists(self.filename):
                    shutil.copyfile(self.filename, self.filename + ".corrupt.bak")
            except Exception:
                pass
            # попытка создать заново
            try:
                if os.path.exists(self.filename):
                    os.remove(self.filename)
            except Exception:
                pass
            self._create_template_if_bundled()
            self.wb = load_workbook(self.filename, data_only=True)

    def save(self):
        try:
            self.wb.save(self.filename)
        except Exception as e:
            logger.exception("Ошибка сохранения Excel: %s", e)

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
        header_raw = rows[0]
        header = []
        used = {}
        for h in header_raw:
            key = normalize_key(h)
            if key in used:
                used[key] += 1
                key = f"{key}_{used[key]}"
            else:
                used[key] = 1
            header.append(key)
        records = []
        for r in rows[1:]:
            if all(v is None for v in r):
                continue
            row = {}
            for i, k in enumerate(header):
                row[k] = r[i]
            records.append(row)
        return records

    def clear_and_write(self, sheet_name: str, header: list, rows: list):
        ws = self.ws(sheet_name)
        try:
            ws.delete_rows(1, ws.max_row)
        except Exception:
            pass
        if header:
            ws.append(header)
        for row in rows:
            ws.append(row)
        self.save()

    def append_form_row(self, row: list):
        ws = self.ws(SHEET_FORM)
        try:
            if ws.max_row == 1 and not any(ws[1]):
                ws.append(FORM_HEADER)
        except Exception:
            pass
        ws.append(row)
        self.save()

# =========================
# Calculators
# =========================

def _calc_imposts_context_minimal(width, height, left, center, right, top):
    """
    Минимальная реализация "gabarit"-переменных, достаточная для большинства формул в СПРАВОЧНИК-1:
    возвращает словарь с N_imp_vert, N_imp_hor, N_impost, N_frame_rect, N_rect, N_corners.
    """
    # приводим к числам
    w = safe_float(width, 0.0)
    h = safe_float(height, 0.0)
    left = safe_float(left, 0.0)
    center = safe_float(center, 0.0)
    right = safe_float(right, 0.0)
    top = safe_float(top, 0.0)

    n_sections_vert = 0
    if left > 0:
        n_sections_vert += 1
    if center > 0:
        n_sections_vert += 1
    if right > 0:
        n_sections_vert += 1

    N_imp_vert = max(0, n_sections_vert - 1)
    N_imp_hor = 1 if top > 0 else 0
    N_impost = N_imp_vert + N_imp_hor
    N_frame_rect = 1 + N_imp_vert + N_imp_hor
    N_rect = N_frame_rect
    N_corners = 4 * N_frame_rect

    return {
        "N_imp_vert": N_imp_vert,
        "N_imp_hor": N_imp_hor,
        "N_impost": N_impost,
        "N_frame_rect": N_frame_rect,
        "N_rect": N_rect,
        "N_corners": N_corners,
    }

class MaterialCalculator:
    """
    Калькулятор расхода материалов.
    Встроена минимальная логика gabarit (через _calc_imposts_context_minimal),
    поэтому отдельный этап Gabarit можно отключить/удалить.
    """
    HEADER = [
        "Тип изделия", "Система профиля", "Тип элемента", "Артикул", "Товар",
        "Ед.", "Цена за ед.", "Ед. фактического расхода",
        "Кол-во факт. расхода", "Норма к упаковке", "Ед. к отгрузке",
        "Кол-во к отгрузке", "Сумма"
    ]

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def calculate(self, order: dict, sections: list, selected_duplicates: dict):
        ref_rows = self.excel.read_records(SHEET_REF1)
        total_area = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)

        if not ref_rows:
            return [], 0.0, total_area

        result_rows = []
        total_sum = 0.0

        for row in ref_rows:
            row_type = get_field(row, "тип издел", "")
            row_profile = get_field(row, "система проф", "")
            type_elem = get_field(row, "тип элемент", "")
            product_name = str(get_field(row, "товар", "") or "")

            # фильтрация по типу изделия/системе
            if row_type:
                if str(row_type).strip().lower() != order.get("product_type", "").strip().lower():
                    continue
            if row_profile:
                if str(row_profile).strip().lower() != order.get("profile_system", "").strip().lower():
                    continue

            # если есть дубли и выбраны конкретные товары
            if type_elem in selected_duplicates and selected_duplicates[type_elem]:
                chosen_names = selected_duplicates[type_elem]
                if product_name not in chosen_names:
                    continue

            # находим формулу
            formula = get_field(row, "формула_python", "")
            if not formula:
                formula = get_field(row, "формула фактического расхода", "")
            if not formula:
                # пропускаем позиции без формулы
                continue

            qty_fact_total = 0.0

            # классификаторы типа элемента
            type_lower = (type_elem or "").lower()
            is_panel_frame = "рамный контур" in type_lower or "импост" in type_lower or "сухарь усилительный" in type_lower
            is_door_item = any(k in type_lower for k in ("рама двери","порог дверной","створочный профиль","петля","замок","цилиндр","ручка","фиксатор","доводчик"))

            # итерация по секциям
            for s in sections:
                is_door_section = s.get("kind") == "door"
                is_panel_section = s.get("kind") in ("panel", "window")

                # логика для Тамбура (копирует прежнюю фильтрацию из оригинала)
                if order.get("product_type") == "Тамбур":
                    if is_door_item and is_panel_section and "сухарь усилительный" not in type_lower:
                        continue
                    if is_panel_frame and is_door_section and "рама двери" not in type_lower and "сухарь усилительный" not in type_lower:
                        continue

                # Сбор/нормализация габаритов для секции
                if is_door_section:
                    width = safe_float(s.get("frame_width_mm", 0.0))
                    height = safe_float(s.get("frame_height_mm", 0.0))
                else:
                    width = safe_float(s.get("width_mm", 0.0))
                    height = safe_float(s.get("height_mm", 0.0))

                left = safe_float(s.get("left_mm", 0.0))
                center = safe_float(s.get("center_mm", 0.0))
                right = safe_float(s.get("right_mm", 0.0))
                top = safe_float(s.get("top_mm", 0.0))

                sash_w = safe_float(s.get("sash_width_mm", width))
                sash_h = safe_float(s.get("sash_height_mm", height))

                area = safe_float(s.get("area_m2", 0.0))
                if not area and width and height:
                    area = (width * height) / 1_000_000.0

                perimeter = safe_float(s.get("perimeter_m", 0.0))
                if not perimeter and width and height:
                    perimeter = 2 * (width + height) / 1000.0

                qty = safe_int(s.get("Nwin", s.get("qty", 1)), default=1)
                nsash = safe_int(s.get("n_sash", s.get("nsash", 1)), default=1)
                n_sash_active = safe_int(s.get("n_sash_active", 1), default=1)
                n_sash_passive = max(nsash - n_sash_active, 0)

                # minimal gabarit context
                geom = _calc_imposts_context_minimal(width, height, left, center, right, top)

                ctx = {
                    "width": width,
                    "height": height,
                    "left": left,
                    "center": center,
                    "right": right,
                    "top": top,
                    "area": area,
                    "perimeter": perimeter,
                    "qty": qty,
                    "sash_width": sash_w,
                    "sash_height": sash_h,
                    "sash_w": sash_w,
                    "sash_h": sash_h,
                    "n_sash": nsash,
                    "n_sash_active": n_sash_active,
                    "n_sash_passive": n_sash_passive,
                    "hinges_per_sash": 3,
                    # include geom computed above (N_imp_vert etc.)
                    **geom,
                }

                # Evaluate formula safely and accumulate
                try:
                    val = safe_eval_formula(str(formula), ctx)
                    try:
                        val = float(val) if val is not None else 0.0
                    except Exception:
                        logger.warning("Non-numeric result from formula %r for %s", formula, type_elem)
                        val = 0.0
                    if val == 0.0:
                        logger.debug("Formula returned zero. element=%s formula=%s ctx_keys=%s", type_elem, formula, list(ctx.keys()))
                    qty_fact_total += val
                except Exception:
                    logger.exception("Error evaluating material formula for %s (Formula: %s) ctx: %s", type_elem, formula, ctx)

            # post-process row: pack, rounding, sum
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

        # записываем результат в Excel
        try:
            self.excel.clear_and_write(SHEET_MATERIAL, self.HEADER, result_rows)
        except Exception:
            logger.exception("Failed to write material sheet")

        return result_rows, total_sum, total_area

class FinalCalculator:
    """
    Финальные расчеты: стеклопакет, тонировка, монтаж, панели, ручки, доводчики, обеспечение.
    (Перенесено из оригинала, с единичной логикой.)
    """
    HEADER = ["Наименование услуг", "Стоимость за м²/шт", "Ед", "Итого"]

    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client

    def _lookup_ref2_rows(self):
        return self.excel.read_records(SHEET_REF2)

    def _find_price_for_filling(self, filling_value):
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        fv = str(filling_value or "").replace("\xa0", " ").strip().lower()
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "панел" in str(k).lower() or "заполн" in str(k).lower():
                    v = r[k]
                    if v is None:
                        continue
                    if str(v).replace("\xa0", " ").strip().lower() == fv:
                        for kk in r.keys():
                            if kk is None:
                                continue
                            if "стоимость" in str(kk).lower():
                                return safe_float(r[kk], 0.0)
        return 0.0

    def _find_price_for_montage(self, montage_type):
        if not montage_type:
            return 0.0
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        mt = str(montage_type or "").replace("\xa0", " ").strip().lower()
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "монтаж" in str(k).lower() and "стоимость" in str(k).lower():
                    return safe_float(r[k], 0.0)
        return 0.0

    def _find_price_for_glass_by_type(self, glass_type):
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        gt = str(glass_type or "").replace("\xa0", " ").strip().lower()
        chosen = None
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "тип стеклопак" in str(k).lower() or "тип стеклопакета" in str(k).lower():
                    v = r[k]
                    if v and str(v).strip().lower() == gt:
                        chosen = r
                        break
            if chosen:
                break
        if not chosen:
            for r in ref2:
                for k in r.keys():
                    if k is None:
                        continue
                    if "стоимость" in str(k).lower() and ("стеклопак" in str(k).lower() or "за м" in str(k).lower()):
                        return safe_float(r[k], 0.0)
            return 0.0
        for k in chosen.keys():
            if k is None:
                continue
            hk = str(k).lower()
            if "стоимость" in hk and ("стеклопак" in hk or "за м" in hk or "за м²" in hk or "за м2" in hk):
                return safe_float(chosen[k], 0.0)
        for k in chosen.keys():
            if k is None:
                continue
            if "стоимость" in str(k).lower():
                return safe_float(chosen[k], 0.0)
        return 0.0

    def _find_price_for_toning(self):
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                hk = str(k).lower()
                if "тониров" in hk and "стоимость" in hk:
                    return safe_float(r[k], 0.0)
        return 0.0

    def _find_price_for_handles(self):
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                hk = str(k).lower()
                if ("ручк" in hk or "ручки" in hk) and "стоимость" in hk:
                    return safe_float(r[k], 0.0)
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "ручк" in str(k).lower():
                    return safe_float(r[k], 0.0)
        return 0.0

    def _find_price_for_closer(self):
        ref2 = self._lookup_ref2_rows()
        if not ref2:
            return 0.0
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                hk = str(k).lower()
                if ("доводчик" in hk or "доводч" in hk) and "стоимость" in hk:
                    return safe_float(r[k], 0.0)
        for r in ref2:
            for k in r.keys():
                if k is None:
                    continue
                if "довод" in str(k).lower():
                    return safe_float(r[k], 0.0)
        return 0.0

    def calculate(self,
                  order: dict,
                  total_area_all: float,
                  material_total: float,
                  lambr_cost: float = 0.0,
                  handles_qty: int = 0,
                  closer_qty: int = 0):
        ref2_rows = self._lookup_ref2_rows()

        glass_type = order.get("glass_type", "")
        toning = order.get("toning", "Нет")
        assembly = order.get("assembly", "Нет")
        montage = order.get("montage", "Нет")
        handle_type = order.get("handle_type", "")
        door_closer = order.get("door_closer", "Нет")

        price_glass = self._find_price_for_glass_by_type(glass_type)
        price_toning = self._find_price_for_toning()
        price_assembly = 0.0
        if ref2_rows:
            for r in ref2_rows:
                for k in r.keys():
                    if k is None:
                        continue
                    hk = str(k).lower()
                    if "сбор" in hk and "стоимость" in hk:
                        price_assembly = safe_float(r[k], 0.0)
                        break
                if price_assembly:
                    break

        price_montage = self._find_price_for_montage(montage)
        price_handles = self._find_price_for_handles()
        price_closer = self._find_price_for_closer()

        rows = []

        glass_sum = total_area_all * price_glass if total_area_all > 0 else 0.0
        rows.append(["Стеклопакет", price_glass, "за м²", glass_sum])

        toning_sum = total_area_all * price_toning if (toning.lower() != "нет" and total_area_all > 0) else 0.0
        rows.append(["Тонировка", price_toning, "за м²", toning_sum])

        assembly_sum = total_area_all * price_assembly if assembly.lower() != "нет" else 0.0
        rows.append(["Сборка", price_assembly, "за м²", assembly_sum])

        montage_sum = total_area_all * price_montage if montage.lower() != "нет" and total_area_all > 0 else 0.0
        rows.append(["Монтаж (" + str(montage) + ")", price_montage, "за м²", montage_sum])

        rows.append(["Материал", "-", "-", material_total])

        if lambr_cost > 0.0:
            rows.append(["Панели (Ламбри/Сэндвич)", "-", "-", lambr_cost])

        handles_sum = price_handles * handles_qty if handles_qty > 0 else 0.0
        rows.append(["Ручки", price_handles, "шт.", handles_sum])

        closer_sum = price_closer * closer_qty if closer_qty > 0 and door_closer.lower() != "нет" else 0.0
        rows.append(["Доводчик", price_closer, "шт.", closer_sum])

        base_sum = (
            glass_sum
            + toning_sum
            + assembly_sum
            + montage_sum
            + material_total
            + lambr_cost
            + handles_sum
            + closer_sum
        )

        ensure_sum = base_sum * 0.6
        rows.append(["Обеспечение (60%)", "", "", ensure_sum])

        total_sum = base_sum + ensure_sum
        extra_rows = [["ИТОГО", "", "", total_sum]]

        try:
            self.excel.clear_and_write(SHEET_FINAL, self.HEADER, rows + extra_rows)
        except Exception:
            logger.exception("Failed to write final sheet")

        return rows, total_sum, ensure_sum

# =========================
# Streamlit UI (упрощённый, но функциональный)
# =========================

def ensure_session_state():
    if st is None:
        return
    if "tam_door_count" not in st.session_state:
        st.session_state["tam_door_count"] = 0
    if "tam_panel_count" not in st.session_state:
        st.session_state["tam_panel_count"] = 0
    if "sections_inputs" not in st.session_state:
        st.session_state["sections_inputs"] = []

def build_smeta_workbook(order: dict,
                         base_positions: list,
                         lambr_positions: list,
                         total_area: float,
                         total_perimeter: float,
                         total_sum: float) -> bytes:
    if Workbook is None:
        return b""
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"
    logo_path = resource_path("logo_axis.png")
    current_row = 1
    if os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.height = 80
            img.width = 80
            ws.add_image(img, "A1")
        except Exception:
            pass
    contact_col = 3
    ws.cell(row=current_row, column=contact_col, value="ООО «AXIS»"); current_row += 1
    ws.cell(row=current_row, column=contact_col, value="Город Астана"); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"Тел.: +7 707 504 4040"); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"E-mail: Axisokna.kz@mail.ru"); current_row += 2
    ws.cell(row=current_row, column=1, value="Коммерческое предложение"); current_row += 2
    ws.cell(row=current_row, column=1, value=f"Заказ № {order.get('order_number','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип изделия: {order.get('product_type','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Профильная система: {order.get('profile_system','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип стеклопакета: {order.get('glass_type','')}"); current_row += 1
    current_row += 1
    ws.cell(row=current_row, column=1, value=f"Общая площадь: {total_area:.3f} м²"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Суммарный периметр: {total_perimeter:.3f} м"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"ИТОГО к оплате: {total_sum:.2f}")
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

def main():
    if st is None:
        print("streamlit is not installed in this environment. Please run this script with streamlit:")
        print("  streamlit run axis_app_clean.py")
        return

    st.set_page_config(page_title="Axis Pro GF • Калькулятор (clean)", layout="wide")
    ensure_session_state()

    excel = ExcelClient(EXCEL_FILE)

    # Авторизация простая (читает лист пользователей)
    user = None
    if "current_user" in st.session_state:
        user = st.session_state["current_user"]
    else:
        # если есть сессия в файле — попробуем загрузить
        try:
            if os.path.exists(SESSION_FILE):
                with open(SESSION_FILE, "r", encoding="utf-8") as sf:
                    st.session_state["current_user"] = json.load(sf)
                    user = st.session_state["current_user"]
        except Exception:
            pass

    if not user:
        # простой логин (показать форму)
        with st.sidebar:
            st.header("🔐 Вход")
            login = st.text_input("Логин")
            password = st.text_input("Пароль", type="password")
            if st.button("Войти"):
                users = {}
                try:
                    rows = excel.read_records(SHEET_USERS)
                    for r in rows:
                        login_k = _clean_cell_val(get_field(r, "логин", "")).lower()
                        pwd = _clean_cell_val(get_field(r, "парол", "")).replace("*", "").strip()
                        role = _clean_cell_val(get_field(r, "роль", ""))
                        if login_k:
                            users[login_k] = {"password": pwd, "role": role, "_raw_login": login_k}
                except Exception:
                    users = {}
                ent = (login or "").strip().lower()
                if ent in users and (password or "").replace("\xa0", "").strip() == users[ent]["password"]:
                    st.session_state["current_user"] = {"login": users[ent]["_raw_login"], "role": users[ent]["role"]}
                    st.experimental_rerun()
                else:
                    st.error("Неверный логин или пароль")
        st.stop()

    st.title("📘 Калькулятор алюминиевых изделий (Axis Pro GF) — clean")
    st.info(f"Пользователь: **{st.session_state['current_user']['login']}**")

    # Загружаем справочники
    ref2_records = excel.read_records(SHEET_REF2)
    filling_types_set = set()
    montage_types_set = set()
    handle_types_set = set()
    glass_types_set = set()

    def _clean_for_set(v):
        if v is None:
            return None
        s = str(v).replace("\xa0", " ").strip()
        return s if s else None

    for row in ref2_records:
        f = _clean_for_set(get_field(row, "панел") or get_field(row, "заполн") or get_field(row, "заполнение"))
        if f:
            filling_types_set.add(f)
        m = _clean_for_set(get_field(row, "монтаж", None))
        if m:
            montage_types_set.add(m)
        h = _clean_for_set(get_field(row, "ручк", None))
        if h:
            handle_types_set.add(h)
        g = _clean_for_set(get_field(row, "тип стеклопак", None) or get_field(row, "тип стеклопакета", None))
        if g:
            glass_types_set.add(g)

    filling_options_for_panels = sorted(list(filling_types_set))
    if 'Стеклопакет' not in filling_options_for_panels:
         filling_options_for_panels.append('Стеклопакет')
    default_panel_fill_index = filling_options_for_panels.index('Ламбри без термо') if 'Ламбри без термо' in filling_options_for_panels else 0

    if not montage_types_set:
        montage_options = ["Есть", "Нет"]
    else:
        montage_options = sorted(list(montage_types_set))
        if "Нет" not in montage_options:
            montage_options.append("Нет")
    if "Нет" in montage_options:
        montage_options.insert(0, montage_options.pop(montage_options.index("Нет")))

    handle_types = sorted(list(handle_types_set)) if handle_types_set else [""]
    glass_types = sorted(list(glass_types_set)) if glass_types_set else ["двойной"]
    default_glass_index = 0
    if "двойной" in glass_types:
        default_glass_index = glass_types.index("двойной")

    # Sidebar общие данные
    with st.sidebar:
        st.header("Общие данные заказа")
        order_number = st.text_input("Номер заказа", value="")
        product_type = st.selectbox("Тип изделия", ["Окно", "Дверь", "Тамбур"])
        profile_system = st.selectbox("Профильная система", ["ALG 2030-45C", "ALG RUIT 63i", "ALG RUIT 73"])
        glass_type = st.selectbox("Тип стеклопакета (справочник)", glass_types, index=default_glass_index)
        st.markdown("### Прочее")
        toning = st.selectbox("Тонировка", ["Нет", "Есть"])
        assembly = st.selectbox("Сборка", ["Нет", "Есть"])
        montage = st.selectbox("Монтаж (из СПРАВОЧНИК-2)", montage_options, index=0)
        handle_type = st.selectbox("Тип ручек", handle_types, index=0)
        door_closer = st.selectbox("Доводчик", ["Нет", "Есть"])

        if st.button("✨ Новый расчёт / Очистить форму"):
            for k in list(st.session_state.keys()):
                if k.startswith(("w_","h_","l_","r_","c_","t_","sw_","sh_","nwin_","ls_w_","ls_h_","ls_q_","ls_fill_","door_","panel_","leaf_","tam_")):
                    st.session_state.pop(k, None)
            st.session_state["sections_inputs"] = []
            st.session_state["tam_door_count"] = 0
            st.session_state["tam_panel_count"] = 0
            st.experimental_rerun()

    col_left, col_right = st.columns([2, 1])

    with col_left:
        st.header("Позиции (окна/двери)")
        base_positions_inputs = []
        lambr_positions_inputs = []
        if product_type != "Тамбур":
            positions_count = st.number_input("Количество позиций (Окно/Дверь)", min_value=1, max_value=10, value=1, step=1)
            for i in range(int(positions_count)):
                st.subheader(f"Позиция {i+1}")
                c1, c2, c3, c4 = st.columns(4)
                width_mm = c1.number_input(f"Ширина, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"w_{i}")
                height_mm = c2.number_input(f"Высота, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"h_{i}")
                left_mm = c3.number_input(f"LEFT, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"l_{i}")
                right_mm = c4.number_input(f"RIGHT, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"r_{i}")

                c5, c6, c7, c8 = st.columns(4)
                center_mm = c5.number_input(f"CENTER, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"c_{i}")
                top_mm = c6.number_input(f"TOP, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"t_{i}")
                sash_width_mm = c7.number_input(f"Ширина створки, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"sw_{i}")
                sash_height_mm = c8.number_input(f"Высота створки, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"sh_{i}")

                nwin = st.number_input(f"Кол-во идентичных рам (N) (поз. {i+1})", min_value=1, value=1, step=1, key=f"nwin_{i}")

                base_positions_inputs.append({
                    "width_mm": width_mm,
                    "height_mm": height_mm,
                    "left_mm": left_mm,
                    "center_mm": center_mm,
                    "right_mm": right_mm,
                    "top_mm": top_mm,
                    "sash_width_mm": sash_width_mm if sash_width_mm > 0 else width_mm,
                    "sash_height_mm": sash_height_mm if sash_height_mm > 0 else height_mm,
                    "Nwin": nwin,
                    "filling": glass_type,
                    "kind": "window" if product_type == "Окно" else "door"
                })
        else:
            # Тамбур: динамические дверные блоки и панели
            st.header("Параметры тамбура (дверные блоки и глухие панели)")
            c_add = st.columns([1,1,6])
            if c_add[0].button("Добавить дверной блок"):
                st.session_state["tam_door_count"] += 1
            if c_add[1].button("Добавить глухую секцию"):
                st.session_state["tam_panel_count"] += 1

            for i in range(st.session_state.get("tam_door_count", 0)):
                with st.expander(f"Дверной блок #{i+1}", expanded=False):
                    name = st.text_input(f"Название блока #{i+1}", value=f"Дверной блок {i+1}", key=f"door_name_{i}")
                    count = st.number_input(f"Кол-во одинаковых блоков #{i+1}", min_value=1, value=1, key=f"door_count_{i}")
                    dtype = st.selectbox(f"Тип двери #{i+1}", ["Одностворчатая","Двухстворчатая"], key=f"door_type_{i}")
                    frame_w = st.number_input(f"Ширина рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, key=f"frame_w_{i}")
                    frame_h = st.number_input(f"Высота рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, key=f"frame_h_{i}")

                    st.subheader("Внутренние импосты (для деления рамы)")
                    c_imp1, c_imp2 = st.columns(2)
                    left = c_imp1.number_input(f"LEFT, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"left_{i}", value=0.0)
                    center = c_imp2.number_input(f"CENTER, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"center_{i}", value=0.0)
                    c_imp3, c_imp4 = st.columns(2)
                    right = c_imp3.number_input(f"RIGHT, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"right_{i}", value=0.0)
                    top = c_imp4.number_input(f"TOP, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"top_{i}", value=0.0)

                    default_leaves = 1 if dtype == "Одностворчатая" else 2
                    n_leaves = st.number_input(f"Кол-во створок #{i+1}", min_value=1, value=default_leaves, key=f"n_leaves_{i}")

                    leaves = []
                    for L in range(int(n_leaves)):
                        st.markdown(f"**Створка {L+1}**")
                        lw = st.number_input(f"Ширина створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, key=f"leaf_w_{i}_{L}")
                        lh = st.number_input(f"Высота створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, key=f"leaf_h_{i}_{L}")
                        fill = st.selectbox(f"Заполнение створки {L+1} — блок {i+1}", options=filling_options_for_panels, index=filling_options_for_panels.index('Стеклопакет') if 'Стеклопакет' in filling_options_for_panels else 0, key=f"leaf_fill_{i}_{L}")
                        leaves.append({"width_mm": lw, "height_mm": lh, "filling": fill})

                    if st.button(f"Добавить/обновить дверной блок #{i+1} в секциях", key=f"save_door_{i}"):
                        new_section = {
                            "kind": "door",
                            "block_name": name,
                            "frame_width_mm": frame_w,
                            "frame_height_mm": frame_h,
                            "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                            "n_leaves": int(n_leaves),
                            "leaves": leaves,
                            "Nwin": int(count),
                            "filling": glass_type
                        }
                        st.session_state["sections_inputs"] = [s for s in st.session_state["sections_inputs"] if not (s.get("block_name") == name and s.get("kind") == "door")]
                        st.session_state["sections_inputs"].append(new_section)
                        st.success(f"Дверной блок '{name}' добавлен/обновлён.")

            for i in range(st.session_state.get("tam_panel_count", 0)):
                with st.expander(f"Глухая секция #{i+1}", expanded=False):
                    name = st.text_input(f"Название панели #{i+1}", value=f"Панель {i+1}", key=f"panel_name_{i}")
                    count = st.number_input(f"Кол-во одинаковых панелей #{i+1}", min_value=1, value=1, key=f"panel_count_{i}")
                    p1, p2 = st.columns(2)
                    w = p1.number_input(f"Ширина панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_w_{i}")
                    h = p2.number_input(f"Высота панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_h_{i}")
                    fill = st.selectbox(f"Заполнение панели #{i+1}", options=filling_options_for_panels, index=default_panel_fill_index, key=f"panel_fill_{i}")

                    st.subheader("Внутренние импосты (для деления рамы)")
                    c_imp5, c_imp6 = st.columns(2)
                    left = c_imp5.number_input(f"LEFT, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_left_{i}", value=0.0)
                    center = c_imp6.number_input(f"CENTER, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_center_{i}", value=0.0)
                    c_imp7, c_imp8 = st.columns(2)
                    right = c_imp7.number_input(f"RIGHT, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_right_{i}", value=0.0)
                    top = c_imp8.number_input(f"TOP, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_top_{i}", value=0.0)

                    if st.button(f"Добавить/обновить панель #{i+1} в секциях", key=f"save_panel_{i}"):
                        new_section = {
                            "kind": "panel",
                            "block_name": name,
                            "width_mm": w,
                            "height_mm": h,
                            "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                            "filling": fill,
                            "Nwin": int(count)
                        }
                        st.session_state["sections_inputs"] = [s for s in st.session_state["sections_inputs"] if not (s.get("block_name") == name and s.get("kind") == "panel")]
                        st.session_state["sections_inputs"].append(new_section)
                        st.success(f"Панель '{name}' добавлена/обновлена.")

            st.markdown("**Текущие секции Тамбура:**")
            if st.session_state["sections_inputs"]:
                 for idx, s in enumerate(st.session_state["sections_inputs"], start=1):
                    main_dim = f"{s.get('width_mm', s.get('frame_width_mm'))}x{s.get('height_mm', s.get('frame_height_mm'))}"
                    imposts = f" L{s.get('left_mm',0)} C{s.get('center_mm',0)} R{s.get('right_mm',0)} T{s.get('top_mm',0)}"
                    st.write(f"**{idx}. {s.get('kind').capitalize()}** ({s.get('block_name')}) — {main_dim}, N={s.get('Nwin',1)} | Импосты:{imposts}")
            else:
                 st.info("Нет добавленных секций.")

        st.markdown("---")

    with col_right:
        st.header("Информация и проверки")
        if not (load_workbook and Workbook):
            st.warning("openpyxl не установлен — некоторые функции Excel будут недоступны.")
        if not os.path.exists(EXCEL_FILE):
            st.warning(f"Excel-файл справочников ({EXCEL_FILE}) не найден. Приложение создаст шаблон при необходимости.")

        # Выбор дублей материалов (при необходимости)
        st.header("🧾 Выбор материалов при дублях")
        selected_duplicates = {}
        ref1 = excel.read_records(SHEET_REF1)
        groups = {}
        for row in ref1:
            row_type = str(get_field(row, "тип издел", "") or "").strip()
            row_profile = str(get_field(row, "система проф", "") or "").strip()

            if row_type and row_type.lower() != product_type.lower():
                continue
            if row_profile and row_profile.lower() != profile_system.lower():
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
        # простая валидация
        if not order_number.strip():
            st.error("Введите номер заказа.")
            st.stop()

        # собираем секции
        sections = []
        if product_type != "Тамбур":
             for p in base_positions_inputs:
                if p["width_mm"] <= 0 or p["height_mm"] <= 0:
                    st.error("Во всех позициях ширина и высота должны быть больше 0.")
                    st.stop()
                area_m2 = (p["width_mm"] * p["height_mm"]) / 1_000_000.0
                perimeter_m = 2 * (p["width_mm"] + p["height_mm"]) / 1000.0
                sections.append({**p, "area_m2": area_m2, "perimeter_m": perimeter_m})

             for p in lambr_positions_inputs:
                if p["width_mm"] > 0 and p["height_mm"] > 0:
                    area_m2 = (p["width_mm"] * p["height_mm"]) / 1_000_000.0
                    perimeter_m = 2 * (p["width_mm"] + p["height_mm"]) / 1000.0
                    sections.append({**p, "area_m2": area_m2, "perimeter_m": perimeter_m, "kind": "panel"})

        else:
             sections = st.session_state["sections_inputs"]
             for s in sections:
                if s.get("kind") == "door":
                    fw = s.get("frame_width_mm", 0.0)
                    fh = s.get("frame_height_mm", 0.0)
                    area_m2 = (fw * fh) / 1_000_000.0
                    perimeter_m = 2 * (fw + fh) / 1000.0
                    s.update({"area_m2": area_m2, "perimeter_m": perimeter_m})
                elif s.get("kind") == "panel":
                    w = s.get("width_mm", 0.0)
                    h = s.get("height_mm", 0.0)
                    area_m2 = (w * h) / 1_000_000.0
                    perimeter_m = 2 * (w + h) / 1000.0
                    s.update({"area_m2": area_m2, "perimeter_m": perimeter_m})

        if not sections:
            st.error("Необходимо задать хотя бы одну позицию с габаритами > 0.")
            st.stop()

        # --- Material Calculation ---
        mat_calc = MaterialCalculator(excel)
        material_rows, material_total, total_area_mat = mat_calc.calculate(
            {"product_type": product_type, "profile_system": profile_system}, sections, selected_duplicates
        )

        # --- Final Calculation ---
        total_area_all = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
        lambr_cost = 0.0

        fin_calc = FinalCalculator(excel)
        # compute lambr cost (if needed): keep original logic (simplified)
        for s in sections:
            fill_name = str(s.get("filling") or "").strip().lower()
            if fill_name in ["ламбри без термо", "ламбри с термо", "сэндвич"]:
                if s.get("kind") == "door":
                    for leaf in s.get("leaves", []):
                        leaf_fill = str(leaf.get("filling") or "").strip().lower()
                        if leaf_fill in ["ламбри без термо", "ламбри с термо", "сэндвич"]:
                            leaf_w = leaf.get("width_mm", 0.0)
                            leaf_h = leaf.get("height_mm", 0.0)
                            perimeter_leaf = 2 * (leaf_w + leaf_h) / 1000.0
                            count_hlyst = math.ceil(perimeter_leaf / 6.0) if perimeter_leaf > 0 else 0
                            price_per_meter = fin_calc._find_price_for_filling(leaf_fill)
                            price_per_hlyst = price_per_meter * 6.0
                            lambr_cost += count_hlyst * price_per_hlyst * s.get("Nwin", 1)
                elif s.get("kind") in ["panel", "window"]:
                    perimeter_s = s.get("perimeter_m", 0.0) * s.get("Nwin", 1)
                    count_hlyst = math.ceil(perimeter_s / 6.0) if perimeter_s > 0 else 0
                    price_per_meter = fin_calc._find_price_for_filling(fill_name)
                    price_per_hlyst = price_per_meter * 6.0
                    lambr_cost += count_hlyst * price_per_hlyst

        # Handles / closers count
        handles_count = 0
        closer_count = 0
        if product_type in ("Дверь", "Тамбур"):
            for s in sections:
                if s.get("kind") == "door":
                    handles_count += s.get("Nwin", 1)
                    if door_closer.lower() == "есть":
                        closer_count += s.get("Nwin", 1)

        final_rows, total_sum, ensure_sum = fin_calc.calculate(
            {
                "product_type": product_type,
                "glass_type": glass_type,
                "toning": toning,
                "assembly": assembly,
                "montage": montage,
                "handle_type": handle_type,
                "door_closer": door_closer
            },
            total_area_all=total_area_all,
            material_total=material_total,
            lambr_cost=lambr_cost,
            handles_qty=handles_count,
            closer_qty=closer_count
        )

        st.success(f"Расчёт выполнен. Итоговая сумма: {total_sum:.2f}")

        # Отобразим результаты
        tab1, tab2, tab3 = st.tabs(["Габариты (по секциям)", "Материалы", "Итоговый расчет"])
        with tab1:
            st.subheader("Секции (Area/Perimeter)")
            st.dataframe([{"kind": s.get("kind"), "area_m2": s.get("area_m2"), "perimeter_m": s.get("perimeter_m"), "Nwin": s.get("Nwin",1)} for s in sections], use_container_width=True)
            st.write(f"Общая площадь: **{total_area_all:.3f} м²**")
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
                        "Цена за ед.": round(safe_float(r[6]), 2),
                        "Ед. факт. расхода": r[7],
                        "Кол-во факт. расхода": round(safe_float(r[8]), 3),
                        "Норма к упаковке": r[9],
                        "Ед. к отгрузке": r[10],
                        "Кол-во к отгрузке": round(safe_float(r[11]), 3),
                        "Сумма": round(safe_float(r[12]), 2),
                    })
                st.dataframe(mat_disp, use_container_width=True)
            st.write(f"Итого по материалам: **{material_total:.2f}**")
            st.write(f"Панели (ламбри/сэндвич) — Итого: **{lambr_cost:.2f}**")
        with tab3:
            st.subheader("Итоговый расчет с монтажом")
            if final_rows:
                fin_disp = []
                for name, price, unit, total_val in final_rows:
                    fin_disp.append({
                        "Наименование услуг": name,
                        "Стоимость за м²/шт": price if isinstance(price, str) else round(price, 2),
                        "Ед": unit,
                        "Итого": total_val if isinstance(total_val, str) else round(total_val, 2),
                    })
                st.dataframe(fin_disp, use_container_width=True)
            st.write(f"Обеспечение (60%): **{ensure_sum:.2f}**")
            st.write(f"ИТОГО к оплате: **{total_sum:.2f}**")

        # Сохраняем в ЗАПРОСЫ
        rows_for_form = []
        pos_index = 1
        for p in sections:
            rows_for_form.append([
                order_number, pos_index, product_type,
                p.get("kind", ""),
                p.get("n_leaves", 1) if p.get("kind") == "door" else 0,
                profile_system, glass_type, p.get("filling",""),
                p.get("width_mm", 0.0) if not p.get("frame_width_mm") else p.get("frame_width_mm", 0.0),
                p.get("height_mm", 0.0) if not p.get("frame_height_mm") else p.get("frame_height_mm", 0.0),
                p.get("left_mm", 0.0), p.get("center_mm", 0.0), p.get("right_mm", 0.0), p.get("top_mm", 0.0),
                p.get("sash_width_mm", p.get("width_mm", 0.0)),
                p.get("sash_height_mm", p.get("height_mm", 0.0)),
                p.get("Nwin", 1),
                toning, assembly, montage, handle_type, door_closer,
            ])
            pos_index += 1
        for row in rows_for_form:
             try:
                 excel.append_form_row(row)
             except Exception:
                 logger.exception("Failed to append form row to Excel")

        # Экспорт коммерческого предложения
        base_pos = [s for s in sections if s.get("kind") in ["window", "door"] and product_type != "Тамбур"]
        tam_pos = [s for s in sections if s.get("kind") in ["door"] and product_type == "Тамбур"]
        lambr_pos = [s for s in sections if s.get("kind") == "panel" or (product_type == "Тамбур" and s.get("kind") != "door")]

        smeta_bytes = build_smeta_workbook(
            order={
                "order_number": order_number, "product_type": product_type, "profile_system": profile_system,
                "filling_mode": "", "glass_type": glass_type, "toning": toning, "assembly": assembly,
                "montage": montage, "handle_type": handle_type, "door_closer": door_closer,
            },
            base_positions=base_pos + tam_pos,
            lambr_positions=lambr_pos,
            total_area=total_area_all,
            total_perimeter=0.0,
            total_sum=total_sum,
        )

        default_name = f"Коммерческое_предложение_Заказ_{order_number}.xlsx"
        try:
            st.download_button(
                "⬇️ Скачать коммерческое предложение в Excel",
                data=smeta_bytes,
                file_name=default_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception:
            st.info("Экспорт недоступен (openpyxl не установлен).")

    # Кнопка выхода
    if st.sidebar.button("Выйти"):
        st.session_state.pop("current_user", None)
        try:
            if os.path.exists(SESSION_FILE):
                os.remove(SESSION_FILE)
        except Exception:
            pass
        st.experimental_rerun()

if __name__ == "__main__":
    main()
