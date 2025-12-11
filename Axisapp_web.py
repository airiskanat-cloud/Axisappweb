import math
import os
import sys
import shutil
from io import BytesIO
import zipfile
import logging
import json
import ast
import operator as op
from typing import Dict, Any, List, Union, Tuple, Set

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.drawing.image import Image as XLImage

# =========================
# ⚙️ КОНСТАНТЫ / НАСТРОЙКИ
# =========================

DEBUG = os.getenv("DEBUG", "False").lower() in ("true", "1", "t")
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG if DEBUG else logging.INFO)

# Конфигурация логирования для Streamlit
if not logger.handlers:
    ch = logging.StreamHandler()
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    logger.addHandler(ch)

def resource_path(relative_path: str) -> str:
    """Определяет корректный путь к ресурсу, учитывая PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

DATA_DIR = os.getenv("AXIS_DATA_DIR", os.path.join(os.path.expanduser("~"), ".axis_app_data"))
os.makedirs(DATA_DIR, exist_ok=True)

TEMPLATE_EXCEL_NAME = "axis_pro_gf.xlsx"
EXCEL_FILE = os.path.join(DATA_DIR, TEMPLATE_EXCEL_NAME)
SESSION_FILE = os.path.join(DATA_DIR, "session_user.json")

# Копирование шаблона при первом запуске
BUNDLED_TEMPLATE = resource_path(TEMPLATE_EXCEL_NAME)
if os.path.exists(BUNDLED_TEMPLATE) and not os.path.exists(EXCEL_FILE):
    try:
        shutil.copyfile(BUNDLED_TEMPLATE, EXCEL_FILE)
        logger.info("Copied bundled template %s -> %s", BUNDLED_TEMPLATE, EXCEL_FILE)
    except Exception:
        logger.exception("Error copying bundled template")

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# Заголовки для ЗАПРОСЫ
FORM_HEADER = [
    "Номер заказа", "№ позиции", "Тип изделия", "Вид изделия", "Створки",
    "Профильная система", "Тип стеклопакета", "Режим заполнения",
    "Ширина, мм", "Высота, мм",
    "LEFT, мм", "CENTER, мм", "RIGHT, мм", "TOP, мм",
    "Ширина створки, мм", "Высота створки, мм", "Кол-во Nwin",
    "Тонировка", "Сборка", "Монтаж", "Тип ручек", "Доводчик"
]

# Брендинг КП
COMPANY_NAME = "ООО «AXIS»"
COMPANY_CITY = "Город Астана"
COMPANY_PHONE = "+7 707 504 4040"
COMPANY_EMAIL = "Axisokna.kz@mail.ru"
COMPANY_SITE = "www.axis.kz"
LOGO_FILENAME = "logo_axis.png"

# =========================
# 🛠️ УТИЛИТЫ
# =========================

def normalize_key(k: Any) -> Union[str, None]:
    """Нормализует ключ (удаляет лишние пробелы, приводит к нижнему регистру)."""
    if k is None:
        return None
    s = str(k).replace("\xa0", " ").strip().lower()
    return " ".join(s.split()) if s else None

def safe_float(value: Any, default: float = 0.0) -> float:
    """Безопасное преобразование к float."""
    try:
        if value is None or (isinstance(value, str) and value.strip() == ""):
            return default
        s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
        return float(s)
    except Exception:
        return default

def safe_int(value: Any, default: int = 0) -> int:
    """Безопасное преобразование к int."""
    try:
        return int(safe_float(value, float(default)))
    except Exception:
        return default

def get_field(row: dict, needle: str, default: Any = None) -> Any:
    """Поиск значения по частичному совпадению ключа (без учета регистра и пробелов)."""
    needle = normalize_key(needle)
    for k, v in row.items():
        if k and needle in normalize_key(k) if normalize_key(k) else False:
            return v
    return default

# =========================
# 🛡️ БЕЗОПАСНЫЙ EVAL (AST)
# =========================

_allowed_ops = {
    ast.Add: op.add, ast.Sub: op.sub, ast.Mult: op.mul,
    ast.Div: op.truediv, ast.Pow: op.pow, ast.USub: op.neg,
    ast.UAdd: op.pos, ast.Mod: op.mod, ast.FloorDiv: op.floordiv,
    ast.Lt: op.lt, ast.Gt: op.gt, ast.LtE: op.le,
    ast.GtE: op.ge, ast.Eq: op.eq, ast.NotEq: op.ne,
    ast.And: lambda a, b: a and b, ast.Or: lambda a, b: a or b,
}

def _eval_ast(node, names: Dict[str, Any]):
    """Рекурсивный обход AST-дерева."""
    if isinstance(node, ast.Expression):
        return _eval_ast(node.body, names)
    if isinstance(node, (ast.Constant, ast.Num)):
        return node.value if isinstance(node, ast.Constant) else node.n
    if isinstance(node, ast.UnaryOp):
        val = _eval_ast(node.operand, names)
        fn = _allowed_ops.get(type(node.op))
        if fn: return fn(val)
    if isinstance(node, ast.BinOp):
        left = _eval_ast(node.left, names)
        right = _eval_ast(node.right, names)
        fn = _allowed_ops.get(type(node.op))
        if fn: return fn(left, right)
    if isinstance(node, ast.Name):
        if node.id in names:
            return names[node.id]
        raise ValueError(f"Недопустимое имя '{node.id}'")
    if isinstance(node, ast.Call):
        func = node.func
        if isinstance(func, ast.Attribute) and isinstance(func.value, ast.Name) and func.value.id == "math":
            fname = func.attr
            if hasattr(math, fname) and not fname.startswith("_"):
                args = [_eval_ast(a, names) for a in node.args]
                return getattr(math, fname)(*args)
        if isinstance(func, ast.Name) and func.id in ("max", "min", "round", "abs"):
            args = [_eval_ast(a, names) for a in node.args]
            return globals()[func.id](*args)
        raise ValueError("Разрешены только math.*, max, min, round, abs")
    if isinstance(node, ast.Compare):
        if len(node.ops) != 1 or len(node.comparators) != 1:
            raise ValueError("Сложные сравнения запрещены")
        left = _eval_ast(node.left, names)
        right = _eval_ast(node.comparators[0], names)
        fn = _allowed_ops.get(type(node.ops[0]))
        if fn: return fn(left, right)
    
    # Добавление для безопасной обработки логических операторов
    if isinstance(node, ast.BoolOp):
        values = [_eval_ast(v, names) for v in node.values]
        op_type = type(node.op)
        if op_type == ast.And:
            return all(values)
        elif op_type == ast.Or:
            return any(values)
        
    raise ValueError(f"Недопустимый элемент формулы: {type(node).__name__}")

def safe_eval_formula(formula: str, context: Dict[str, Any]) -> float:
    """Безопасное вычисление математической формулы с использованием AST."""
    formula = (formula or "").strip()
    if not formula:
        return 0.0

    # Создаем контекст с добавлением безопасных модулей и функций
    names = {
        **context,
        "math": math,
        "min": min, "max": max, "round": round, "abs": abs,
    }

    try:
        # Обработка условных выражений (if ... then ... else ...)
        formula_lower = formula.lower()
        if formula_lower.startswith("if "):
            if " then " not in formula_lower or " else " not in formula_lower:
                 # Если не полный синтаксис, считаем это обычной формулой
                 pass 
            else:
                # Берем оригинальный регистр для формулы, чтобы сохранить имена переменных
                original_parts = formula.split(" else ", 1)
                if len(original_parts) < 2:
                    raise ValueError("Неполный синтаксис if-then-else")

                if_then_part = original_parts[0]
                false_result_str = original_parts[1].strip()

                if " then " not in if_then_part.lower():
                    raise ValueError("Неполный синтаксис if-then-else")
                
                condition_str = if_then_part[3:].split(" then ", 1)[0].strip()
                true_result_str = if_then_part[3:].split(" then ", 1)[1].strip()

                # Вычисляем условие
                condition = bool(_eval_ast(ast.parse(condition_str, mode="eval"), names))
                
                # Вычисляем 'true' и 'false' части
                if condition:
                    return float(_eval_ast(ast.parse(true_result_str, mode="eval"), names))
                else:
                    return float(_eval_ast(ast.parse(false_result_str, mode="eval"), names))

        # Стандартное вычисление (math expression)
        node = ast.parse(formula, mode="eval")
        return float(_eval_ast(node, names))
    except (ValueError, TypeError, ZeroDivisionError, IndexError) as e:
        logger.debug("safe_eval_formula error for formula '%s' with context %s: %s", formula, context, e)
        return 0.0
    except Exception as e:
        logger.error("Critical error in safe_eval for formula '%s': %s", formula, e)
        return 0.0

# =========================
# 🗃️ EXCEL/CATALOG CLIENT
# =========================

class ExcelClient:
    """Клиент для работы с Excel-файлом справочников (с авто-бэкапом)."""
    def __init__(self, filename: str):
        self.filename = filename
        if not os.path.exists(self.filename):
            self._create_template()
        self.load()

    def _create_template(self):
        """Создает пустой шаблон Excel, если файл не найден."""
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        for sheet_name in [SHEET_FORM, SHEET_REF1, SHEET_REF2, SHEET_REF3, SHEET_USERS]:
            wb.create_sheet(sheet_name)
        wb.save(self.filename)
        logger.info("Created new Excel template: %s", self.filename)

    def load(self):
        """Загружает рабочую книгу, выполняя бэкап в случае ошибки."""
        try:
            # data_only=True для чтения значений, а не формул
            self.wb = load_workbook(self.filename, data_only=True)
        except Exception as e:
            logger.exception("Error loading Excel, making backup and recreating template: %s", e)
            try:
                if os.path.exists(self.filename):
                    shutil.copyfile(self.filename, self.filename + ".corrupt.bak")
            except Exception:
                pass
            try:
                if os.path.exists(self.filename):
                    os.remove(self.filename)
            except Exception:
                pass
            self._create_template()
            self.wb = load_workbook(self.filename, data_only=True)

    def ws(self, name: str):
        """Возвращает лист по имени, создавая его при необходимости."""
        if name in self.wb.sheetnames:
            return self.wb[name]
        ws = self.wb.create_sheet(name)
        self.save()
        return ws

    def save(self):
        """Сохраняет рабочую книгу."""
        try:
            self.wb.save(self.filename)
        except Exception as e:
            logger.exception("Save error: %s", e)

    def read_records(self, sheet_name: str) -> List[Dict[str, Any]]:
        """Читает данные из листа, возвращая список словарей (records)."""
        ws = self.ws(sheet_name)
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
            
        header_raw = rows[0]
        header = []
        used_keys: Dict[str, int] = {}

        # Нормализация и обработка дубликатов заголовков
        for h in header_raw:
            key = normalize_key(h)
            if key in used_keys:
                used_keys[key] += 1
                key = f"{key}_{used_keys[key]}"
            elif key:
                used_keys[key] = 1
            header.append(key)

        records = []
        for r in rows[1:]:
            if all(v is None or (isinstance(v, str) and v.strip() == "") for v in r):
                logger.debug("Skipped empty row in sheet: %s", sheet_name)
                continue
            row = {}
            for i, k in enumerate(header):
                if k:
                    # Сохраняем оригинальное значение
                    row[k] = r[i]
            records.append(row)
        return records

    def clear_and_write(self, sheet_name: str, header: List[str], rows: List[List[Any]]):
        """Очищает лист и записывает новые данные."""
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

    def append_form_row(self, row: List[Any]):
        """Добавляет строку в лист ЗАПРОСЫ."""
        ws = self.ws(SHEET_FORM)
        try:
            # Убедиться, что заголовок есть
            if ws.max_row == 1 and not any(ws[1]):
                ws.append(FORM_HEADER)
        except Exception:
            pass
        ws.append(row)
        self.save()

def process_catalog_ref1(ref1_records: List[Dict[str, Any]]) -> Dict[Tuple[str, str, str], Dict[str, Any]]:
    """Обрабатывает СПРАВОЧНИК-1, создавая уникальные ключи для поиска."""
    catalog = {}
    for row in ref1_records:
        product_type = normalize_key(get_field(row, "тип издел", "")) or "universal"
        profile_system = normalize_key(get_field(row, "система проф", "")) or "universal"
        element_type = normalize_key(get_field(row, "тип элемент", ""))
        product_name = normalize_key(get_field(row, "товар", ""))
        
        if not element_type or not product_name:
            continue
            
        # Ключ: (Тип изделия, Система профиля, Тип элемента, Название товара)
        # Добавляем название товара, чтобы различать дубликаты
        key = (product_type, profile_system, element_type, product_name)
        catalog[key] = row
        
    return catalog

# =========================
# 🧠 КОНТЕКСТ И ФОРМУЛЫ
# =========================

def ensure_defaults(order: Dict[str, Any], sections: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Генерирует универсальный контекст заказа с дефолтами для формул."""
    
    # 1. Общие параметры заказа
    ctx: Dict[str, Any] = {
        "product_type": normalize_key(order.get("product_type", "окно")),
        "profile_system": normalize_key(order.get("profile_system", "")),
        "glass_type": normalize_key(order.get("glass_type", "")),
        "toning": normalize_key(order.get("toning", "нет")),
        "assembly": normalize_key(order.get("assembly", "нет")),
        "montage": normalize_key(order.get("montage", "нет")),
        "handle_type": normalize_key(order.get("handle_type", "")),
        "door_closer": normalize_key(order.get("door_closer", "нет")),
    }
    
    total_area_m2 = sum(s.get("area_m2", 0.0) * s.get("Nwin", 1) for s in sections)
    total_perimeter_m = sum(s.get("perimeter_m", 0.0) * s.get("Nwin", 1) for s in sections)
    
    # 2. Суммарные/агрегированные параметры
    ctx.update({
        "total_area": total_area_m2,
        "total_perimeter": total_perimeter_m,
        "n_frames_total": sum(s.get("Nwin", 1) for s in sections),
        "n_doors_total": sum(s.get("Nwin", 1) for s in sections if s.get("kind") == "door"),
        "n_windows_total": sum(s.get("Nwin", 1) for s in sections if s.get("kind") == "window"),
        "n_panels_total": sum(s.get("Nwin", 1) for s in sections if s.get("kind") == "panel"),
    })
    
    return ctx

def fallback_formula_eval(
    formula: str,
    formula_group: str,
    section: Dict[str, Any],
    order_context: Dict[str, Any]
) -> float:
    """
    Вычисляет формулу, используя контекст секции и заказа.
    """
    
    # 1. Формирование контекста для формулы:
    width = safe_float(section.get("width_mm", section.get("frame_width_mm", 0.0)))
    height = safe_float(section.get("height_mm", section.get("frame_height_mm", 0.0)))
    qty = safe_int(section.get("Nwin", 1))

    # Габариты импостов
    left = safe_float(section.get("left_mm", 0.0))
    center = safe_float(section.get("center_mm", 0.0))
    right = safe_float(section.get("right_mm", 0.0))
    top = safe_float(section.get("top_mm", 0.0))
    
    # Логика для подсчета импостов
    n_sections_vert = (1 if left > 0 else 0) + (1 if center > 0 else 0) + (1 if right > 0 else 0)
    n_imp_vert = max(0, n_sections_vert - 1)
    n_imp_hor = 1 if top > 0 else 0
    
    n_impost = n_imp_vert + n_imp_hor
    n_frame_rect = 1 + n_imp_vert + n_imp_hor # Количество прямоугольников в раме
    n_rect = n_frame_rect
    n_corners = 4 * n_frame_rect
    
    # Параметры створки (для створочных профилей и фурнитуры)
    sash_w = safe_float(section.get("sash_width_mm", width))
    sash_h = safe_float(section.get("sash_height_mm", height))
    n_leaves = safe_int(section.get("n_leaves", len(section.get("leaves", [])) or 1))

    # Контекст, который будет доступен в формуле
    context_data = {
        "width": width, "height": height, "w": width, "h": height,
        "sash_width": sash_w, "sash_height": sash_h, "sash_w": sash_w, "sash_h": sash_h,
        "left": left, "center": center, "right": right, "top": top,
        
        "area": safe_float(section.get("area_m2", 0.0)),
        "perimeter": safe_float(section.get("perimeter_m", 0.0)),
        "qty": qty, 
        
        "n_imp_vert": n_imp_vert, "n_imp_hor": n_imp_hor, "n_impost": n_impost,
        "n_frame_rect": n_frame_rect, "n_rect": n_rect, "n_corners": n_corners,
        
        "n_leaves": n_leaves, "n_sash": n_leaves,
        "n_sash_active": 1 if n_leaves >= 1 else 0,
        "n_sash_passive": max(n_leaves - 1, 0),
        "hinges_per_sash": 3,
    }
    
    # 2. Вычисление
    result = safe_eval_formula(formula, {**order_context, **context_data})
    
    return result

# =========================
# 🧮 КАЛЬКУЛЯТОРЫ
# =========================

class OrderProcessor:
    """Главный класс для расчета материалов, использующий унифицированную логику."""
    
    # Группы, которые должны учитывать упаковку (pack_size)
    PROFILE_GROUPS = ["профиль", "усилитель", "сухарь", "импост"]
    
    def __init__(self, excel_client: ExcelClient):
        self.excel = excel_client
        self.ref1_catalog = process_catalog_ref1(self.excel.read_records(SHEET_REF1))
        self.ref3_records = self.excel.read_records(SHEET_REF3)
        self.ref2_records = self.excel.read_records(SHEET_REF2)
        
    def _is_relevant(self, row: Dict[str, Any], order_ctx: Dict[str, Any], section: Dict[str, Any], selected_duplicates: Dict[str, Set[str]]) -> bool:
        """Определяет, применим ли элемент справочника к текущему заказу и секции."""
        
        row_type = normalize_key(get_field(row, "тип издел", "")) or "universal"
        row_profile = normalize_key(get_field(row, "система проф", "")) or "universal"
        type_elem = normalize_key(get_field(row, "тип элемент", ""))
        product_name = normalize_key(get_field(row, "товар", ""))
        
        order_type = order_ctx.get("product_type", "")
        order_profile = order_ctx.get("profile_system", "")
        section_kind = section.get("kind", "") # window, door, panel

        # 1. Фильтрация по Типу изделия (Улучшено: 'universal' и пустые поля считаются совпадением)
        if row_type != "universal" and row_type != order_type:
            return False

        # 2. Фильтрация по Системе профиля (Улучшено: 'universal' и пустые поля считаются совпадением)
        if row_profile != "universal" and row_profile != order_profile:
            return False

        # 3. Фильтрация по дубликатам (если выбраны конкретные товары)
        if type_elem in selected_duplicates and product_name:
            if product_name not in selected_duplicates[type_elem]:
                return False

        # 4. Фильтрация по типу секции (для Тамбура)
        is_door_item = any(k in type_elem for k in ["рама двери", "створочный", "петля", "замок", "цилиндр", "ручка", "доводчик"])
        is_panel_frame = any(k in type_elem for k in ["рамный контур", "импост", "сухарь усилительный", "усилитель", "стеклопакет", "заполнение"])
        
        if order_type == "тамбур":
            if section_kind == "door":
                if not is_door_item and not is_panel_frame:
                    return False
            elif section_kind == "panel":
                # В панели ищем рамные/импостные элементы и заполнение, но исключаем чистую фурнитуру
                if is_door_item and "сухарь усилительный" not in type_elem:
                    return False
                if not is_panel_frame and not is_door_item:
                     return False
                    
        return True

    def calculate_materials(self, order: Dict[str, Any], sections: List[Dict[str, Any]], selected_duplicates: Dict[str, Set[str]]) -> Tuple[pd.DataFrame, float, float]:
        """Расчет материалов из СПРАВОЧНИК-1."""
        order_ctx = ensure_defaults(order, sections)
        material_results: Dict[str, Dict[str, Any]] = {} # Ключ: (Тип элемента, Товар)
        
        total_sum = 0.0
        total_area = order_ctx.get("total_area", 0.0)

        # 1. Сбор итогового фактического расхода
        for row_key, row in self.ref1_catalog.items():
            product_type_row, profile_system_row, element_type, product_name = row_key
            
            formula = str(get_field(row, "формула_python", "") or get_field(row, "формула фактического расхода", "")).strip()
            if not formula:
                continue

            qty_fact_total_for_item = 0.0
            
            for section in sections:
                if not self._is_relevant(row, order_ctx, section, selected_duplicates):
                    continue

                # Вычисляем фактический расход для этой секции
                try:
                    qty_fact_for_section = fallback_formula_eval(formula, element_type, section, order_ctx)
                    
                    qty_fact_total_for_item += qty_fact_for_section * safe_int(section.get("Nwin", 1))
                except Exception as e:
                    logger.error("Error in formula for %s (%s): %s", product_name, formula, e)
                    continue

            # 2. Агрегация и расчет отгрузки (упаковка)
            if qty_fact_total_for_item > 0.0:
                key = (element_type, product_name)
                
                item_data = material_results.setdefault(key, {
                    "Тип изделия": get_field(row, "тип издел", ""),
                    "Система профиля": get_field(row, "система проф", ""),
                    "Тип элемента": get_field(row, "тип элемент", ""),
                    "Артикул": get_field(row, "артикул", ""),
                    "Товар": get_field(row, "товар", ""),
                    "Ед.": get_field(row, "ед.", ""),
                    "Цена за ед.": safe_float(get_field(row, "цена за", 0.0)),
                    "Ед. факт. расхода": get_field(row, "ед. фактического расхода", ""),
                    "Кол-во факт. расхода": 0.0,
                    "Норма к упаковке": safe_float(get_field(row, "кол-во норм", 0.0)), # pack_size
                    "Ед. к отгрузке": str(get_field(row, "ед .норма к упаковке", "") or "").strip(),
                    "Кол-во к отгрузке": 0.0,
                    "Сумма": 0.0,
                })
                
                item_data["Кол-во факт. расхода"] += qty_fact_total_for_item

        # 3. Финальный расчет упаковки и суммы
        final_rows = []
        
        for key, item_data in material_results.items():
            qty_fact_total = item_data["Кол-во факт. расхода"]
            norm_per_pack = item_data["Норма к упаковке"]
            unit_price = item_data["Цена за ед."]
            
            qty_to_ship = qty_fact_total
            effective_qty = qty_fact_total
            
            if norm_per_pack > 0:
                is_profile = any(g in normalize_key(item_data["Тип элемента"]) for g in self.PROFILE_GROUPS)

                if is_profile or "шт" in normalize_key(item_data["Ед. к отгрузке"]):
                    qty_packs = math.ceil(qty_fact_total / norm_per_pack)
                    qty_to_ship = qty_packs # Количество упаковок к отгрузке
                    effective_qty = qty_packs * norm_per_pack # Общее количество товара (суммируется)
                else:
                    qty_to_ship = qty_fact_total
                    effective_qty = qty_fact_total

            sum_row = effective_qty * unit_price
            total_sum += sum_row
            
            final_rows.append([
                item_data["Тип изделия"], item_data["Система профиля"], item_data["Тип элемента"], item_data["Артикул"], 
                item_data["Товар"], item_data["Ед."], item_data["Цена за ед."], item_data["Ед. факт. расхода"],
                qty_fact_total, norm_per_pack, item_data["Ед. к отгрузке"], qty_to_ship, sum_row
            ])

        # Сортировка для чистого вывода (по типу элемента и товару)
        sorted_rows = sorted(final_rows, key=lambda x: (x[2], x[4]))
        
        # Запись в Excel
        self.excel.clear_and_write(SHEET_MATERIAL, MaterialCalculator.HEADER, sorted_rows)
        
        df = pd.DataFrame(sorted_rows, columns=MaterialCalculator.HEADER)
        
        return df, total_sum, total_area

    def calculate_gabarits(self, order: Dict[str, Any], sections: List[Dict[str, Any]]) -> Tuple[pd.DataFrame, float, float]:
        """Расчет габаритов из СПРАВОЧНИК-3."""
        order_ctx = ensure_defaults(order, sections)
        gabarit_values: Dict[str, float] = {}
        
        total_area = order_ctx.get("total_area", 0.0)
        total_perimeter = order_ctx.get("total_perimeter", 0.0)
        
        for row in self.ref3_records:
            type_elem = str(get_field(row, "тип элемент", "") or "").strip()
            formula = str(get_field(row, "формула_python", "") or "").strip()
            
            if not type_elem or not formula:
                continue

            total_value = 0.0

            for section in sections:
                # В СПРАВОЧНИК-3 нет фильтрации по типу/профилю, поэтому считаем для всех секций
                try:
                    total_value += fallback_formula_eval(formula, type_elem, section, order_ctx)
                except Exception as e:
                    logger.error("Error evaluating formula for element %s: %s", type_elem, e)
            
            if total_value > 0.0 or DEBUG:
                gabarit_values[type_elem] = total_value

        # Запись в Excel
        gabarit_list = [[t, v] for t, v in sorted(gabarit_values.items())]
        self.excel.clear_and_write(SHEET_GABARITS, GabaritCalculator.HEADER, gabarit_list)
        
        df = pd.DataFrame(gabarit_list, columns=GabaritCalculator.HEADER)
        
        return df, total_area, total_perimeter
        
    def calculate_final(self, order: Dict[str, Any], material_df: pd.DataFrame, total_area_all: float) -> Tuple[pd.DataFrame, float, float]:
        """Расчет итоговой стоимости из СПРАВОЧНИК-2."""
        
        # 1. Агрегация данных из материального расчета
        material_total = safe_float(material_df["Сумма"].sum())
        
        # Поиск стоимости Ламбри/Сэндвич (по периметру/погонный метр)
        lambr_cost = self._calculate_lambr_cost(order, self.ref2_records)
        
        # Подсчет фурнитуры/штучных элементов (Ручки/Доводчики)
        handles_qty = order.get("n_doors_total", 0) # 1 ручка на дверной блок
        closer_qty = handles_qty if order.get("door_closer", "").lower() == "есть" else 0 # 1 доводчик на дверной блок

        # 2. Поиск цен из СПРАВОЧНИК-2 (Услуги)
        final_calc = FinalCalculator(self.ref2_records)
        
        price_glass = final_calc._find_price_for_glass_by_type(order.get("glass_type", ""))
        price_toning = final_calc._find_price_for_toning()
        price_assembly = final_calc._find_price_for_assembly()
        price_montage = final_calc._find_price_for_montage(order.get("montage", ""))
        price_handles = final_calc._find_price_for_handles(order.get("handle_type", ""))
        price_closer = final_calc._find_price_for_closer(order.get("door_closer", ""))
        
        # 3. Расчет сумм по услугам
        rows = []
        
        # Стеклопакет и Тонировка, Сборка, Монтаж - от общей площади
        glass_sum = total_area_all * price_glass
        rows.append(["Стеклопакет", price_glass, "за м²", glass_sum])

        toning_sum = total_area_all * price_toning if order.get("toning", "").lower() == "есть" else 0.0
        rows.append(["Тонировка", price_toning, "за м²", toning_sum])

        assembly_sum = total_area_all * price_assembly if order.get("assembly", "").lower() == "есть" else 0.0
        rows.append(["Сборка", price_assembly, "за м²", assembly_sum])

        montage_sum = total_area_all * price_montage if order.get("montage", "").lower() != "нет" else 0.0
        rows.append([f"Монтаж ({order.get('montage', 'Нет')})", price_montage, "за м²", montage_sum])

        # Материалы и панели - как есть
        rows.append(["Материал", "-", "-", material_total])
        if lambr_cost > 0.0:
            rows.append(["Панели (Ламбри/Сэндвич)", "-", "-", lambr_cost])

        # Фурнитура - поштучно
        handles_sum = price_handles * handles_qty
        rows.append(["Ручки", price_handles, f"шт. (N={handles_qty})", handles_sum])

        closer_sum = price_closer * closer_qty
        rows.append(["Доводчик", price_closer, f"шт. (N={closer_qty})", closer_sum])

        # 4. Итоги
        base_sum = sum(r[3] for r in rows if isinstance(r[3], (int, float)))
        
        # Обеспечение (60%)
        ensure_sum = base_sum * 0.6
        rows.append(["Обеспечение (60%)", "", "", ensure_sum])

        total_sum = base_sum + ensure_sum
        rows.append(["ИТОГО", "", "", total_sum])
        
        # Запись в Excel
        header = ["Наименование услуг", "Стоимость за м²/шт", "Ед", "Итого"]
        self.excel.clear_and_write(SHEET_FINAL, header, rows)
        
        df = pd.DataFrame(rows, columns=header)
        
        return df, total_sum, ensure_sum
        
    def _calculate_lambr_cost(self, order: Dict[str, Any], ref2_records: List[Dict[str, Any]]) -> float:
        """
        Расчет стоимости Ламбри/Сэндвич (по периметру/пог. метру).
        """
        final_calc = FinalCalculator(ref2_records)
        lambr_cost = 0.0
        sections = order.get("sections_inputs", [])
        
        for section in sections:
            qty_nwin = safe_int(section.get("Nwin", 1))
            
            # 1. Секция - глухая панель или окно
            if section.get("kind") in ["panel", "window"]:
                fill_name = normalize_key(section.get("filling", ""))
                price_per_meter = final_calc._find_price_for_filling(fill_name)
                
                if price_per_meter > 0.0 and ("ламбри" in fill_name or "сэндвич" in fill_name):
                    perimeter_s = safe_float(section.get("perimeter_m", 0.0))
                    
                    # Логика расчета по хлыстам (6м)
                    count_hlyst = math.ceil(perimeter_s / 6.0) if perimeter_s > 0 else 0
                    price_per_hlyst = price_per_meter * 6.0
                    lambr_cost += count_hlyst * price_per_hlyst * qty_nwin

            # 2. Секция - дверной блок с наполнением створок
            elif section.get("kind") == "door":
                for leaf in section.get("leaves", []):
                    leaf_fill = normalize_key(leaf.get("filling", ""))
                    price_per_meter = final_calc._find_price_for_filling(leaf_fill)
                    
                    if price_per_meter > 0.0 and ("ламбри" in leaf_fill or "сэндвич" in leaf_fill):
                        # Рассчитываем периметр створки
                        leaf_w = safe_float(leaf.get("width_mm", 0.0))
                        leaf_h = safe_float(leaf.get("height_mm", 0.0))
                        perimeter_leaf = 2 * (leaf_w + leaf_h) / 1000.0
                        
                        count_hlyst = math.ceil(perimeter_leaf / 6.0) if perimeter_leaf > 0 else 0
                        price_per_hlyst = price_per_meter * 6.0
                        lambr_cost += count_hlyst * price_per_hlyst * qty_nwin

        return lambr_cost

class GabaritCalculator:
    HEADER = ["Тип элемента", "Фактическое значение"]
    # Тело класса интегрировано в OrderProcessor.calculate_gabarits

class MaterialCalculator:
    HEADER = [
        "Тип изделия", "Система профиля", "Тип элемента", "Артикул", "Товар",
        "Ед.", "Цена за ед.", "Ед. факт. расхода",
        "Кол-во факт. расхода", "Норма к упаковке", "Ед. к отгрузке",
        "Кол-во к отгрузке", "Сумма"
    ]
    
class FinalCalculator:
    """Утилиты для поиска цен из СПРАВОЧНИК-2."""
    
    def __init__(self, ref2_records: List[Dict[str, Any]]):
        self.ref2_records = ref2_records
        
    def _find_price(self, search_keys: Union[str, List[str]], filter_key_val: Tuple[str, str] = None) -> float:
        """Общая утилита для поиска цены в СПРАВОЧНИК-2."""
        if isinstance(search_keys, str):
            search_keys = [search_keys]
        
        for r in self.ref2_records:
            is_match = True
            if filter_key_val:
                f_key, f_val = filter_key_val
                if normalize_key(get_field(r, f_key, "")) != normalize_key(f_val):
                    is_match = False
            
            if is_match:
                for k in r.keys():
                    nk = normalize_key(k)
                    if nk and any(sk in nk for sk in search_keys) and "стоимость" in nk:
                        return safe_float(r[k], 0.0)
        return 0.0
        
    def _find_price_for_filling(self, filling_value: str) -> float:
        """Цена за м.п. для Ламбри/Сэндвич."""
        if not filling_value: return 0.0
        f_val = normalize_key(filling_value)
        
        for r in self.ref2_records:
            found_filling = False
            for k in r.keys():
                nk = normalize_key(k)
                if nk and any(n in nk for n in ["панел", "заполн", "заполнение"]):
                    if normalize_key(r[k]) == f_val:
                        found_filling = True
                        break
            
            if found_filling:
                for k in r.keys():
                    nk = normalize_key(k)
                    if nk and "стоимость" in nk:
                        return safe_float(r[k], 0.0)
        return 0.0

    def _find_price_for_montage(self, montage_type: str) -> float:
        """Цена монтажа (берется из колонки 'стоимость монтажа' независимо от типа)."""
        return self._find_price("монтаж", filter_key_val=None)

    def _find_price_for_glass_by_type(self, glass_type: str) -> float:
        """Цена стеклопакета по типу."""
        if not glass_type: return 0.0
        f_val = normalize_key(glass_type)
        
        # Ищем точное совпадение
        for r in self.ref2_records:
            for k in r.keys():
                nk = normalize_key(k)
                if nk and any(n in nk for n in ["тип стеклопак"]):
                    if normalize_key(r[k]) == f_val:
                        return self._find_price("стоимость", filter_key_val=("тип стеклопак", glass_type))
        
        # Fallback: Если тип не найден, ищем просто любую цену стеклопакета
        return self._find_price("стеклопак")

    def _find_price_for_toning(self) -> float:
        """Цена тонировки."""
        return self._find_price("тониров")
        
    def _find_price_for_assembly(self) -> float:
        """Цена сборки."""
        return self._find_price("сбор")
        
    def _find_price_for_handles(self, handle_type: str) -> float:
        """Цена ручки."""
        if not handle_type: return 0.0
        return self._find_price("ручк")

    def _find_price_for_closer(self, closer_type: str) -> float:
        """Цена доводчика."""
        if closer_type.lower() == "нет": return 0.0
        return self._find_price("доводчик")
        
# =========================
# EXPORT: коммерческое предложение
# =========================

def build_smeta_workbook(order: dict,
                         sections: list,
                         total_area: float,
                         total_perimeter: float,
                         total_sum: float) -> bytes:
    """Создает Excel-файл коммерческого предложения."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"

    # 1. Заголовки и лого
    logo_path = resource_path(LOGO_FILENAME)
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
    ws.cell(row=current_row, column=contact_col, value=COMPANY_NAME); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=COMPANY_CITY); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"Тел.: {COMPANY_PHONE}"); current_row += 1
    ws.cell(row=current_row, column=contact_col, value=f"E-mail: {COMPANY_EMAIL}"); current_row += 1
    if COMPANY_SITE:
        ws.cell(row=current_row, column=contact_col, value=f"Сайт: {COMPANY_SITE}"); current_row += 1

    current_row += 1
    ws.cell(row=current_row, column=1, value="Коммерческое предложение"); current_row += 2

    # 2. Общие параметры заказа
    ws.cell(row=current_row, column=1, value=f"Заказ № {order.get('order_number','')}").font = ws.cell(row=current_row, column=1).font.copy(bold=True); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип изделия: {order.get('product_type','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Профильная система: {order.get('profile_system','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип стеклопакета: {order.get('glass_type','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тонировка: {order.get('toning','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Сборка: {order.get('assembly','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Монтаж: {order.get('montage','')}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Тип ручек: {order.get('handle_type','') or '—'}"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Доводчик: {order.get('door_closer','')}"); current_row += 2

    # 3. Детализация позиций (секций)
    ws.cell(row=current_row, column=1, value="Детализация секций:").font = ws.cell(row=current_row, column=1).font.copy(bold=True); current_row += 1
    
    for idx, p in enumerate(sections, start=1):
        is_door = p.get('kind') == 'door'
        w = p.get('frame_width_mm', p.get('width_mm', 0)) if is_door else p.get('width_mm', 0)
        h = p.get('frame_height_mm', p.get('height_mm', 0)) if is_door else p.get('height_mm', 0)
        
        fill_info = f" Заполнение: {p.get('filling', '')}"
        
        if is_door and p.get('leaves'):
            leaves_fills = ", ".join([f"Л{l+1}: {leaf.get('filling', '')}" for l, leaf in enumerate(p['leaves'])])
            fill_info = f" Заполнения створок: {leaves_fills}"

        kind_name = p.get('block_name', f"Позиция {idx}")
        dims = f"{w} × {h} мм"
        qty_info = f" N={p.get('Nwin',1)}"
        
        ws.cell(row=current_row, column=1, value=f"{idx}. {kind_name} ({p.get('kind', '').capitalize()}) — {dims}{qty_info}{fill_info}")
        current_row += 1

    current_row += 2
    
    # 4. Итоговые цифры
    ws.cell(row=current_row, column=1, value=f"Общая площадь: {total_area:.3f} м²"); current_row += 1
    ws.cell(row=current_row, column=1, value=f"Суммарный периметр: {total_perimeter:.3f} м"); current_row += 1
    
    ws.cell(row=current_row, column=1, value=f"ИТОГО к оплате: {total_sum:.2f}").font = ws.cell(row=current_row, column=1).font.copy(bold=True, size=14)

    try:
        for col in ['A','B','C','D','E','F']:
            ws.column_dimensions[col].width = 25
    except Exception:
        pass

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# =========================
# 🌐 STREAMLIT UI: main
# =========================

def ensure_session_state():
    """Инициализация session_state."""
    if "tam_door_count" not in st.session_state:
        st.session_state["tam_door_count"] = 0
    if "tam_panel_count" not in st.session_state:
        st.session_state["tam_panel_count"] = 0
    if "sections_inputs" not in st.session_state:
        st.session_state["sections_inputs"] = []
    if "selected_duplicates" not in st.session_state:
        st.session_state["selected_duplicates"] = {}
    if "last_calculation" not in st.session_state:
        st.session_state["last_calculation"] = None

def load_users(excel: ExcelClient) -> Dict[str, Dict[str, str]]:
    """Загрузка пользователей."""
    excel.load()
    rows = excel.read_records(SHEET_USERS)
    users = {}

    for r in rows:
        login = str(get_field(r, "логин", "") or "").strip().lower()
        pwd = str(get_field(r, "парол", "") or "").replace("*", "").strip()
        role = str(get_field(r, "роль", "") or "").strip()

        if login:
            users[login] = {"password": pwd, "role": role, "_raw_login": login}
    return users

def login_form(excel: ExcelClient) -> Union[Dict[str, str], None]:
    """Форма входа."""
    if "current_user" in st.session_state:
        return st.session_state["current_user"]

    # Попытка восстановить из файла
    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE, "r", encoding="utf-8") as sf:
                st.session_state["current_user"] = json.load(sf)
                return st.session_state["current_user"]
        except Exception:
            pass

    st.sidebar.title("🔐 Вход в систему")
    with st.sidebar.form("login_form"):
        login = st.text_input("Логин")
        password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти")

    users = load_users(excel)

    if submitted:
        entered_login = (login or "").strip().lower()
        entered_pass = (password or "").strip()

        user = users.get(entered_login)

        if user and entered_pass == (user["password"] or "").strip():
            st.session_state["current_user"] = {
                "login": user["_raw_login"],
                "role": user["role"],
            }
            try:
                with open(SESSION_FILE, "w", encoding="utf-8") as sf:
                    json.dump(st.session_state["current_user"], sf, ensure_ascii=False)
            except Exception:
                pass

            st.sidebar.success(f"Привет, {user['_raw_login']}!")
            st.rerun() # ИСПРАВЛЕНО
            return st.session_state["current_user"]

        st.sidebar.error("Неверный логин или пароль")

    return None

def collect_catalog_options(ref2_records: List[Dict[str, Any]]) -> Tuple[List[str], List[str], List[str], List[str]]:
    """Собирает уникальные опции для выпадающих списков из СПРАВОЧНИК-2."""
    filling_types_set = set()
    montage_types_set = set()
    handle_types_set = set()
    glass_types_set = set()

    def _clean_for_set(v):
        s = str(v).replace("\xa0", " ").strip() if v is not None else ""
        return s if s else None

    for row in ref2_records:
        f = _clean_for_set(get_field(row, "панел") or get_field(row, "заполн") or get_field(row, "заполнение"))
        if f: filling_types_set.add(f)
        m = _clean_for_set(get_field(row, "монтаж", None))
        if m: montage_types_set.add(m)
        h = _clean_for_set(get_field(row, "ручк", None))
        if h: handle_types_set.add(h)
        g = _clean_for_set(get_field(row, "тип стеклопак", None) or get_field(row, "тип стеклопакета", None))
        if g: glass_types_set.add(g)

    # Заполнение для панелей
    filling_options_for_panels = sorted(list(filling_types_set))
    if 'Стеклопакет' not in filling_options_for_panels: filling_options_for_panels.append('Стеклопакет')

    # Монтаж
    montage_options = sorted(list(montage_types_set))
    if "Нет" not in montage_options: montage_options.append("Нет")
    if "Нет" in montage_options: montage_options.insert(0, montage_options.pop(montage_options.index("Нет")))

    # Ручки/Стеклопакеты
    handle_types = sorted(list(handle_types_set)) if handle_types_set else [""]
    glass_types = sorted(list(glass_types_set)) if glass_types_set else ["двойной"]

    return filling_options_for_panels, montage_options, handle_types, glass_types

def main():
    st.set_page_config(page_title="Axis Pro GF • Калькулятор", layout="wide")
    ensure_session_state()
    excel = ExcelClient(EXCEL_FILE)

    user = login_form(excel)
    if not user:
        st.stop()

    st.title("📘 Калькулятор алюминиевых изделий (Axis Pro GF)")
    st.info(f"Пользователь: **{user['login']}**")

    # Загружаем опции для selectbox'ов
    ref2_records = excel.read_records(SHEET_REF2)
    filling_options_for_panels, montage_options, handle_types, glass_types = collect_catalog_options(ref2_records)

    default_glass_index = glass_types.index("двойной") if "двойной" in glass_types else 0
    default_handle_index = 0
    if not handle_types: handle_types = [""]

    # ---------- Sidebar: общие данные заказа ----------
    with st.sidebar:
        st.header("Общие данные заказа")
        order_number = st.text_input("Номер заказа", value=st.session_state.get("order_number", ""))
        product_type = st.selectbox("Тип изделия", ["Окно", "Дверь", "Тамбур"], index=["Окно", "Дверь", "Тамбур"].index(st.session_state.get("product_type", "Окно")))
        profile_system = st.selectbox("Профильная система", ["ALG 2030-45C", "ALG RUIT 63i", "ALG RUIT 73"], index=["ALG 2030-45C", "ALG RUIT 63i", "ALG RUIT 73"].index(st.session_state.get("profile_system", "ALG 2030-45C")))
        glass_type = st.selectbox("Тип стеклопакета (цена из СПРАВОЧНИК-2)", glass_types, index=default_glass_index)
        st.markdown("### Прочее")
        toning = st.selectbox("Тонировка", ["Нет", "Есть"], index=["Нет", "Есть"].index(st.session_state.get("toning", "Нет")))
        assembly = st.selectbox("Сборка", ["Нет", "Есть"], index=["Нет", "Есть"].index(st.session_state.get("assembly", "Нет")))
        montage = st.selectbox("Монтаж (из СПРАВОЧНИК-2)", montage_options, index=montage_options.index(st.session_state.get("montage", "Нет")))
        handle_type = st.selectbox("Тип ручек", handle_types, index=default_handle_index)
        door_closer = st.selectbox("Доводчик", ["Нет", "Есть"], index=["Нет", "Есть"].index(st.session_state.get("door_closer", "Нет")))
        
        # Обновление session_state для сохранения значений при переключении
        st.session_state["order_number"] = order_number
        st.session_state["product_type"] = product_type
        st.session_state["profile_system"] = profile_system
        st.session_state["toning"] = toning
        st.session_state["assembly"] = assembly
        st.session_state["montage"] = montage
        st.session_state["door_closer"] = door_closer

        if st.button("✨ Новый расчёт / Очистить форму"):
            for k in list(st.session_state.keys()):
                if k.startswith(("w_","h_","l_","r_","c_","t_","sw_","sh_","nwin_","leaf_","door_","panel_")) or k in ["tam_door_count", "tam_panel_count", "sections_inputs", "selected_duplicates", "last_calculation"]:
                    st.session_state.pop(k, None)
            st.rerun() # ИСПРАВЛЕНО

    # --- Главная колонка: ввод позиций ---
    col_left, col_right = st.columns([2, 1])

    with col_left:
        st.header("Позиции (окна/двери)")
        
        base_positions_inputs: List[Dict[str, Any]] = []

        if product_type != "Тамбур":
            # Логика для Окна/Двери
            positions_count = st.number_input("Количество позиций", min_value=1, max_value=10, value=st.session_state.get("positions_count", 1), step=1, key="positions_count")
            
            for i in range(int(positions_count)):
                st.subheader(f"Позиция {i+1}")
                c1, c2, c3, c4 = st.columns(4)
                # Динамические ключи для сохранения состояния
                w = c1.number_input(f"Ширина, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"w_{i}", value=st.session_state.get(f"w_{i}", 0.0))
                h = c2.number_input(f"Высота, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"h_{i}", value=st.session_state.get(f"h_{i}", 0.0))
                l = c3.number_input(f"LEFT, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"l_{i}", value=st.session_state.get(f"l_{i}", 0.0))
                r = c4.number_input(f"RIGHT, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"r_{i}", value=st.session_state.get(f"r_{i}", 0.0))

                c5, c6, c7, c8 = st.columns(4)
                c = c5.number_input(f"CENTER, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"c_{i}", value=st.session_state.get(f"c_{i}", 0.0))
                t = c6.number_input(f"TOP, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"t_{i}", value=st.session_state.get(f"t_{i}", 0.0))
                sw = c7.number_input(f"Ширина створки, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"sw_{i}", value=st.session_state.get(f"sw_{i}", 0.0))
                sh = c8.number_input(f"Высота створки, мм (поз. {i+1})", min_value=0.0, step=10.0, key=f"sh_{i}", value=st.session_state.get(f"sh_{i}", 0.0))

                nwin = st.number_input(f"Кол-во идентичных рам (N) (поз. {i+1})", min_value=1, value=st.session_state.get(f"nwin_{i}", 1), step=1, key=f"nwin_{i}")
                n_leaves = st.number_input(f"Кол-во створок (для фурнитуры) (поз. {i+1})", min_value=0, value=st.session_state.get(f"n_leaves_{i}", 1 if product_type == "Дверь" else 0), step=1, key=f"n_leaves_{i}")
                
                if w > 0.0 and h > 0.0:
                    area_m2 = (w * h) / 1_000_000.0
                    perimeter_m = 2 * (w + h) / 1000.0
                    base_positions_inputs.append({
                        "width_mm": w, "height_mm": h, "left_mm": l, "center_mm": c, "right_mm": r, "top_mm": t,
                        "sash_width_mm": sw if sw > 0 else w, "sash_height_mm": sh if sh > 0 else h,
                        "Nwin": nwin, "filling": glass_type, "kind": normalize_key(product_type),
                        "area_m2": area_m2, "perimeter_m": perimeter_m, "n_leaves": n_leaves
                    })
            
            # Для не-Тамбура секции берутся из base_positions_inputs
            st.session_state["sections_inputs"] = base_positions_inputs
            st.session_state["tam_door_count"] = 0
            st.session_state["tam_panel_count"] = 0
            
        else:
            # --- Динамический блок для Тамбура ---
            st.header("Параметры тамбура (дверные блоки и глухие панели)")

            c_add = st.columns([1,1,6])
            if c_add[0].button("Добавить дверной блок"):
                st.session_state["tam_door_count"] += 1
            if c_add[1].button("Добавить глухую секцию"):
                st.session_state["tam_panel_count"] += 1
                
            # Дверные блоки
            for i in range(st.session_state.get("tam_door_count", 0)):
                with st.expander(f"Дверной блок #{i+1}", expanded=False):
                    name = st.text_input(f"Название блока #{i+1}", value=st.session_state.get(f"door_name_{i}", f"Дверной блок {i+1}"), key=f"door_name_{i}")
                    count = st.number_input(f"Кол-во одинаковых блоков #{i+1}", min_value=1, value=st.session_state.get(f"door_count_{i}", 1), key=f"door_count_{i}")
                    
                    frame_w = st.number_input(f"Ширина рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, key=f"frame_w_{i}", value=st.session_state.get(f"frame_w_{i}", 0.0))
                    frame_h = st.number_input(f"Высота рамы (изделия), мм #{i+1}", min_value=0.0, step=10.0, key=f"frame_h_{i}", value=st.session_state.get(f"frame_h_{i}", 0.0))
                    
                    st.subheader("Внутренние импосты (для деления рамы)")
                    c_imp1, c_imp2 = st.columns(2)
                    left = c_imp1.number_input(f"LEFT, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"left_{i}", value=st.session_state.get(f"left_{i}", 0.0))
                    center = c_imp2.number_input(f"CENTER, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"center_{i}", value=st.session_state.get(f"center_{i}", 0.0))
                    c_imp3, c_imp4 = st.columns(2)
                    right = c_imp3.number_input(f"RIGHT, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"right_{i}", value=st.session_state.get(f"right_{i}", 0.0))
                    top = c_imp4.number_input(f"TOP, мм #{i+1} (ДБ)", min_value=0.0, step=10.0, key=f"top_{i}", value=st.session_state.get(f"top_{i}", 0.0))

                    n_leaves = st.number_input(f"Кол-во створок #{i+1}", min_value=1, value=st.session_state.get(f"n_leaves_{i}", 1), key=f"n_leaves_{i}")

                    leaves = []
                    for L in range(int(n_leaves)):
                        st.markdown(f"**Створка {L+1}**")
                        lw = st.number_input(f"Ширина створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, key=f"leaf_w_{i}_{L}", value=st.session_state.get(f"leaf_w_{i}_{L}", 0.0))
                        lh = st.number_input(f"Высота створки {L+1} (мм) — блок {i+1}", min_value=0.0, step=10.0, key=f"leaf_h_{i}_{L}", value=st.session_state.get(f"leaf_h_{i}_{L}", 0.0))
                        fill = st.selectbox(f"Заполнение створки {L+1} — блок {i+1}", options=filling_options_for_panels, index=filling_options_for_panels.index('Стеклопакет') if 'Стеклопакет' in filling_options_for_panels else 0, key=f"leaf_fill_{i}_{L}")
                        leaves.append({"width_mm": lw, "height_mm": lh, "filling": fill})

                    if st.button(f"Добавить/обновить дверной блок #{i+1} в секциях", key=f"save_door_{i}"):
                        if frame_w <= 0 or frame_h <= 0:
                            st.error("Ширина и высота рамы дверного блока должны быть > 0.")
                        else:
                            new_section = {
                                "kind": "door", "block_name": name, "frame_width_mm": frame_w, "frame_height_mm": frame_h,
                                "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                                "n_leaves": int(n_leaves), "leaves": leaves, "Nwin": int(count), "filling": glass_type,
                            }
                            new_section.update({"area_m2": (frame_w * frame_h) / 1_000_000.0, "perimeter_m": 2 * (frame_w + frame_h) / 1000.0})
                            
                            # Удаляем старую секцию с тем же именем
                            st.session_state["sections_inputs"] = [s for s in st.session_state["sections_inputs"] if not (s.get("block_name") == name and s.get("kind") == "door")]
                            st.session_state["sections_inputs"].append(new_section)
                            st.success(f"Дверной блок '{name}' добавлен/обновлён.")
                            st.rerun() # Обновляем для показа добавленной секции
                
            # Глухие секции (панели)
            for i in range(st.session_state.get("tam_panel_count", 0)):
                with st.expander(f"Глухая секция #{i+1}", expanded=False):
                    name = st.text_input(f"Название панели #{i+1}", value=st.session_state.get(f"panel_name_{i}", f"Панель {i+1}"), key=f"panel_name_{i}")
                    count = st.number_input(f"Кол-во одинаковых панелей #{i+1}", min_value=1, value=st.session_state.get(f"panel_count_{i}", 1), key=f"panel_count_{i}")
                    p1, p2 = st.columns(2)
                    w = p1.number_input(f"Ширина панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_w_{i}", value=st.session_state.get(f"panel_w_{i}", 0.0))
                    h = p2.number_input(f"Высота панели, мм #{i+1}", min_value=0.0, step=10.0, key=f"panel_h_{i}", value=st.session_state.get(f"panel_h_{i}", 0.0))
                    
                    # Определение дефолтного индекса для заполнения панели
                    default_panel_fill_index = filling_options_for_panels.index('Ламбри без термо') if 'Ламбри без термо' in filling_options_for_panels else 0
                    fill = st.selectbox(f"Заполнение панели #{i+1}", options=filling_options_for_panels, index=default_panel_fill_index, key=f"panel_fill_{i}")
                    
                    st.subheader("Внутренние импосты (для деления рамы)")
                    c_imp5, c_imp6 = st.columns(2)
                    left = c_imp5.number_input(f"LEFT, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_left_{i}", value=st.session_state.get(f"panel_left_{i}", 0.0))
                    center = c_imp6.number_input(f"CENTER, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_center_{i}", value=st.session_state.get(f"panel_center_{i}", 0.0))
                    c_imp7, c_imp8 = st.columns(2)
                    right = c_imp7.number_input(f"RIGHT, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_right_{i}", value=st.session_state.get(f"panel_right_{i}", 0.0))
                    top = c_imp8.number_input(f"TOP, мм #{i+1} (ГС)", min_value=0.0, step=10.0, key=f"panel_top_{i}", value=st.session_state.get(f"panel_top_{i}", 0.0))

                    if st.button(f"Добавить/обновить панель #{i+1} в секциях", key=f"save_panel_{i}"):
                        if w <= 0 or h <= 0:
                            st.error("Ширина и высота панели должны быть > 0.")
                        else:
                            new_section = {
                                "kind": "panel", "block_name": name, "width_mm": w, "height_mm": h,
                                "left_mm": left, "center_mm": center, "right_mm": right, "top_mm": top,
                                "filling": fill, "Nwin": int(count)
                            }
                            new_section.update({"area_m2": (w * h) / 1_000_000.0, "perimeter_m": 2 * (w + h) / 1000.0})
                            
                            # Удаляем старую секцию с тем же именем
                            st.session_state["sections_inputs"] = [s for s in st.session_state["sections_inputs"] if not (s.get("block_name") == name and s.get("kind") == "panel")]
                            st.session_state["sections_inputs"].append(new_section)
                            st.success(f"Панель '{name}' добавлена/обновлена.")
                            st.rerun() # Обновляем для показа добавленной секции
                            
            st.markdown("**Текущие секции Тамбура:**")
            if st.session_state["sections_inputs"]:
                for idx, s in enumerate(st.session_state["sections_inputs"], start=1):
                    is_door = s.get('kind') == 'door'
                    w = s.get('frame_width_mm', s.get('width_mm', 0)) if is_door else s.get('width_mm', 0)
                    h = s.get('frame_height_mm', s.get('height_mm', 0)) if is_door else s.get('height_mm', 0)
                    main_dim = f"{w} × {h}"
                    imposts = f" L{s.get('left_mm',0)} C{s.get('center_mm',0)} R{s.get('right_mm',0)} T{s.get('top_mm',0)}"
                    st.write(f"**{idx}. {s.get('kind').capitalize()}** ({s.get('block_name')}) — {main_dim}, N={s.get('Nwin',1)} | Импосты:{imposts}")
            else:
                st.info("Нет добавленных секций.")

        st.markdown("---")

    # --- Правая колонка: выбор дубликатов ---
    with col_right:
        st.header("Информация")
        if product_type == "Тамбур":
            st.info("Тамбур детализируется отдельными секциями: дверные блоки и глухие панели.")
            
        if not os.path.exists(EXCEL_FILE) or not zipfile.is_zipfile(EXCEL_FILE):
             st.warning("Excel-файл справочников не найден или поврежден. Создан новый шаблон.")
        
        # ---------- Выбор материалов при дублях ----------
        st.header("🧾 Выбор материалов при дублях")
        
        ref1 = excel.read_records(SHEET_REF1)
        groups: Dict[str, Set[str]] = {}
        
        # Собираем дубликаты для текущего типа изделия/профиля
        for row in ref1:
            row_type = normalize_key(get_field(row, "тип издел", "")) or "universal"
            row_profile = normalize_key(get_field(row, "система проф", "")) or "universal"
            type_elem = normalize_key(get_field(row, "тип элемент", ""))
            product_name = normalize_key(get_field(row, "товар", ""))

            if row_type != normalize_key(product_type) and row_type != "universal": continue
            if row_profile != normalize_key(profile_system) and row_profile != "universal": continue
            if not type_elem or not product_name: continue

            groups.setdefault(type_elem, set()).add(product_name)

        current_duplicates = st.session_state.get("selected_duplicates", {})

        if not any(len(products) > 1 for products in groups.values()):
            st.info("Для выбранного типа изделия и профиля дублей материалов не найдено.")
        else:
            for type_elem, products in sorted(groups.items(), key=lambda kv: kv[0]):
                if len(products) <= 1: continue
                
                sorted_products = sorted(list(products))
                
                # Используем сохраненное значение из сессии, или все по дефолту
                default_selection = current_duplicates.get(type_elem, sorted_products)
                
                chosen = st.multiselect(
                    f"Тип элемента: {type_elem.capitalize()}",
                    options=sorted_products,
                    default=default_selection,
                    key=f"dup_{type_elem}"
                )
                current_duplicates[type_elem] = set(normalize_key(c) for c in chosen)
                
        st.session_state["selected_duplicates"] = current_duplicates

    # ---------- Кнопка расчёта ----------
    st.markdown("---")
    calc_button = st.button("💾 Сохранить в Excel и выполнить расчёт", use_container_width=True)

    if calc_button:
        if not order_number.strip():
            st.error("Введите номер заказа.")
            st.stop()
            
        if not st.session_state["sections_inputs"] or all(s.get("area_m2", 0.0) <= 0.0 for s in st.session_state["sections_inputs"]):
            st.error("Необходимо задать хотя бы одну позицию с габаритами > 0.")
            st.stop()
            
        # --- Сборка полного контекста заказа ---
        order_details = {
            "order_number": order_number, "product_type": product_type, "profile_system": profile_system,
            "glass_type": glass_type, "toning": toning, "assembly": assembly,
            "montage": montage, "handle_type": handle_type, "door_closer": door_closer,
            "sections_inputs": st.session_state["sections_inputs"], # Полный список секций
        }
        order_ctx = ensure_defaults(order_details, st.session_state["sections_inputs"])
        
        # --- Расчет ---
        calculator = OrderProcessor(excel)
        
        # 1. Габариты (СПРАВОЧНИК-3)
        gabarit_df, total_area_gab, total_perimeter_gab = calculator.calculate_gabarits(order_ctx, st.session_state["sections_inputs"])

        # 2. Материалы (СПРАВОЧНИК-1)
        material_df, material_total, _ = calculator.calculate_materials(order_ctx, st.session_state["sections_inputs"], st.session_state["selected_duplicates"])
        
        # 3. Итоговый расчет (СПРАВОЧНИК-2)
        final_df, total_sum, ensure_sum = calculator.calculate_final(order_ctx, material_df, total_area_gab)

        # Сохраняем результат в сессию
        st.session_state["last_calculation"] = {
            "gabarit_df": gabarit_df, "material_df": material_df, "final_df": final_df,
            "total_area": total_area_gab, "total_perimeter": total_perimeter_gab, "total_sum": total_sum,
            "lambr_cost": final_df[final_df["Наименование услуг"].str.contains("Панели")]["Итого"].sum()
        }
        
        st.success(f"Расчёт выполнен. Итоговая сумма: {total_sum:.2f}")
        
        # --- Сохраняем в ЗАПРОСЫ ---
        rows_for_form: List[List[Any]] = []
        for pos_index, p in enumerate(st.session_state["sections_inputs"], start=1):
            
            # Определение вида изделия
            kind_item = p.get("kind", "")
            if kind_item == "panel": kind_name = "Глухая секция"
            elif kind_item == "door" and product_type == "Тамбур": kind_name = "Дверной блок"
            elif kind_item == "door": kind_name = "Дверь"
            else: kind_name = "Окно"
                 
            # Ширина/высота
            width_f = p.get("frame_width_mm", p.get("width_mm", 0.0))
            height_f = p.get("frame_height_mm", p.get("height_mm", 0.0))
            
            # Ширина/высота створки
            sash_w_f = p.get("sash_width_mm", 0.0)
            sash_h_f = p.get("sash_height_mm", 0.0)

            # Для Тамбура: сохраняем детализацию заполнения
            filling_mode = p.get("filling", glass_type)
            if kind_item == "door" and p.get("leaves"):
                filling_mode = ", ".join([f"Л{l+1}: {leaf.get('filling')}" for l, leaf in enumerate(p['leaves'])])

            rows_for_form.append([
                order_number, pos_index, product_type,
                kind_name,
                p.get("n_leaves", 1),
                profile_system, glass_type, filling_mode,
                width_f, height_f,
                p.get("left_mm", 0.0), p.get("center_mm", 0.0), p.get("right_mm", 0.0), p.get("top_mm", 0.0),
                sash_w_f, sash_h_f,
                p.get("Nwin", 1),
                toning, assembly, montage, handle_type, door_closer,
            ])

        for row in rows_for_form:
            excel.append_form_row(row)
        
    # --- Вывод результатов, если расчет был ---
    if st.session_state["last_calculation"]:
        calc_data = st.session_state["last_calculation"]
        
        tab1, tab2, tab3, tab4 = st.tabs(["Габариты", "Материалы (по элементам)", "Материалы (по группам)", "Итоговый расчет"])
        
        # 1. Габариты
        with tab1:
            st.subheader("Расчет по габаритам (СПРАВОЧНИК-3)")
            st.dataframe(calc_data["gabarit_df"], use_container_width=True, hide_index=True)
            st.write(f"Общая площадь: **{calc_data['total_area']:.3f} м²**")
            st.write(f"Суммарный периметр: **{calc_data['total_perimeter']:.3f} м**")

        # 2. Материалы (по элементам)
        with tab2:
            st.subheader("Расчёт материалов (СПРАВОЧНИК-1): Детализация")
            
            # Логирование нулевых строк:
            zero_rows = calc_data['material_df'][calc_data['material_df']['Кол-во факт. расхода'] == 0.0]
            if not zero_rows.empty:
                st.warning(f"⚠️ **{len(zero_rows)} строк** в расчете материалов имеют нулевой расход. Проверьте формулы в СПРАВОЧНИК-1.")
                for _, row in zero_rows.iterrows():
                    logger.warning("Zero material consumption: %s - %s", row['Тип элемента'], row['Товар'])
            
            st.dataframe(calc_data["material_df"], use_container_width=True, hide_index=True, column_config={
                "Цена за ед.": st.column_config.NumberColumn(format="%.2f"),
                "Кол-во факт. расхода": st.column_config.NumberColumn(format="%.3f"),
                "Кол-во к отгрузке": st.column_config.NumberColumn(format="%.3f"),
                "Сумма": st.column_config.NumberColumn(format="%.2f"),
            })
            st.write(f"Итого по материалам: **{calc_data['material_df']['Сумма'].sum():.2f}**")

        # 3. Материалы (по группам)
        with tab3:
            st.subheader("Расчёт материалов: Сводка по Типу элемента")
            
            # Агрегация по Типу элемента
            group_summary = calc_data['material_df'].groupby('Тип элемента').agg(
                Товаров_шт=('Товар', 'count'),
                Сумма_группы=('Сумма', 'sum')
            ).reset_index()
            
            # Агрегация по типу профиля
            group_summary_profile = calc_data['material_df'].groupby(['Тип изделия', 'Система профиля']).agg(
                Сумма_профиля=('Сумма', 'sum')
            ).reset_index()

            st.markdown("##### По типу элемента:")
            st.dataframe(group_summary.sort_values(by='Сумма_группы', ascending=False), use_container_width=True, hide_index=True, column_config={
                "Сумма_группы": st.column_config.NumberColumn("Сумма, ИТОГО", format="%.2f"),
            })
            
            st.markdown("##### По системе профиля:")
            st.dataframe(group_summary_profile, use_container_width=True, hide_index=True, column_config={
                "Сумма_профиля": st.column_config.NumberColumn("Сумма, ИТОГО", format="%.2f"),
            })
            
            # Общий итог
            st.markdown("---")
            st.write(f"Общая сумма материалов: **{calc_data['material_df']['Сумма'].sum():.2f}**")


        # 4. Итоговый расчет
        with tab4:
            st.subheader("Итоговый расчет с монтажом (СПРАВОЧНИК-2)")
            final_df_disp = calc_data["final_df"].iloc[:-1] # Убираем итоговую строку для красивого вывода
            final_sum_row = calc_data["final_df"].iloc[-1]
            
            st.dataframe(final_df_disp, use_container_width=True, hide_index=True, column_config={
                "Стоимость за м²/шт": st.column_config.NumberColumn(format="%.2f"),
                "Итого": st.column_config.NumberColumn(format="%.2f"),
            })
            
            st.markdown("---")
            st.write(f"Обеспечение (60%): **{ensure_sum:.2f}**")
            st.markdown(f"**ИТОГО к оплате: {total_sum:.2f}**")

        # --- Экспорт коммерческого предложения ---
        smeta_bytes = build_smeta_workbook(
            order=order_ctx,
            sections=st.session_state["sections_inputs"],
            total_area=calc_data["total_area"],
            total_perimeter=calc_data["total_perimeter"],
            total_sum=calc_data["total_sum"],
        )

        default_name = f"Коммерческое_предложение_Заказ_{order_number}.xlsx"
        st.download_button(
            "⬇️ Скачать коммерческое предложение в Excel",
            data=smeta_bytes,
            file_name=default_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ---------- Кнопка выхода ----------
    if st.sidebar.button("Выйти"):
        st.session_state.pop("current_user", None)
        try:
            if os.path.exists(SESSION_FILE):
                os.remove(SESSION_FILE)
        except Exception:
            pass
        st.rerun() # ИСПРАВЛЕНО

if __name__ == "__main__":
    main()
