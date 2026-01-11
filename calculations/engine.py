import math
import logging
from typing import Dict, List

logger = logging.getLogger(__name__)

# === МАППИНГ СИСТЕМ ПРОФИЛЯ ===
SYSTEM_MAPPING = {
    "ALG RUIT 73i 22MM": "ALG 2030-73C",
    "ALG RUIT 73i": "ALG 2030-73C",
    "ALG RUIT 63i": "ALG 2030-63C",
    "ALG RUIT 55i": "ALG 2030-55C",
    "ALG RUIT 45i": "ALG 2030-45C",
    "ALG 2030-73C": "ALG 2030-73C",
    "ALG 2030-63C": "ALG 2030-63C",
    "ALG 2030-55C": "ALG 2030-55C",
    "ALG 2030-45C": "ALG 2030-45C"
}

# === ОТСТУПЫ ДЛЯ СВЕТОВОГО ПРОЕМА (по системам) ===
SYSTEM_OFFSETS = {
    "ALG 2030-73C": 73,
    "ALG 2030-63C": 63,
    "ALG 2030-55C": 55,
    "ALG 2030-45C": 45
}

def normalize_system(system_id: str) -> str:
    """Нормализация названия системы профиля"""
    return SYSTEM_MAPPING.get(system_id, system_id)

def safe_eval(formula: str, context: dict) -> float:
    """Безопасное вычисление формул Python"""
    try:
        f = str(formula).replace(",", ".").replace(" ", "")
        return float(eval(f, {"__builtins__": None, "math": math}, context))
    except Exception as e:
        logger.warning(f"Ошибка в формуле '{formula}': {e}")
        return 0.0

def calculate_impost_length(W: float, H: float, system: str, impost_type: str = "vertical") -> float:
    """
    Расчет длины импоста
    
    impost_type: "vertical" или "horizontal"
    """
    normalized_sys = normalize_system(system)
    offset = SYSTEM_OFFSETS.get(normalized_sys, 73)
    
    if impost_type == "vertical":
        return H - (offset * 2)
    else:  # horizontal
        return W - (offset * 2)

def calculate_glass_opening(W: float, H: float, system: str) -> tuple:
    """
    Расчет светового проема (w_g, h_g)
    
    Световой проем = габарит окна - толщина профиля × 2
    """
    normalized_sys = normalize_system(system)
    offset = SYSTEM_OFFSETS.get(normalized_sys, 73)
    
    w_g = max(W - (offset * 2), 0)
    h_g = max(H - (offset * 2), 0)
    
    return w_g, h_g

def calculate_window_geometry(position_data: Dict, system: str) -> Dict:
    """
    Расчет геометрии окна/двери
    Возвращает все необходимые переменные для формул
    """
    W = position_data.get("width", 0)
    H = position_data.get("height", 0)
    count = position_data.get("count", 1)
    
    # Импосты
    imposts = position_data.get("imposts", {})
    
    # Автоматический расчет если не указано вручную
    if imposts.get("auto_calculate", True):
        imp_vertical_left = calculate_impost_length(W, H, system, "vertical") if imposts.get("has_left", False) else 0
        imp_vertical_center = calculate_impost_length(W, H, system, "vertical") if imposts.get("has_center", False) else 0
        imp_vertical_right = calculate_impost_length(W, H, system, "vertical") if imposts.get("has_right", False) else 0
        imp_horizontal_top = calculate_impost_length(W, H, system, "horizontal") if imposts.get("has_tor", False) else 0
    else:
        # Ручной ввод
        imp_vertical_left = imposts.get("left", 0)
        imp_vertical_center = imposts.get("center", 0)
        imp_vertical_right = imposts.get("right", 0)
        imp_horizontal_top = imposts.get("tor", 0)
    
    total_imposts = imp_vertical_left + imp_vertical_center + imp_vertical_right + imp_horizontal_top
    
    # Створки
    sashes = position_data.get("sashes", [])
    
    # Если есть створки - берем размеры первой для формул
    if sashes:
        w_s = sashes[0].get("w", 0)
        h_s = sashes[0].get("h", 0)
    else:
        w_s = 0
        h_s = 0
    
    # Световой проем (для штапика) - по системе профиля
    w_g, h_g = calculate_glass_opening(W, H, system)
    
    # Количество точек запирания (по высоте створки)
    if h_s < 1200:
        n_lp = 2
    elif h_s < 2000:
        n_lp = 3
    else:
        n_lp = 4
    
    # Площадь (ИСПРАВЛЕНО: с учетом количества)
    area = (W * H / 1_000_000) * count
    
    # Периметр (ИСПРАВЛЕНО: с учетом количества)
    perimeter = (2 * (W + H) / 1000) * count
    
    return {
        "W": W,
        "H": H,
        "count": count,
        "w_s": w_s,
        "h_s": h_s,
        "w_stvor": w_s,  # алиас для совместимости с формулами
        "h_stvor": h_s,
        "w_g": w_g,
        "h_g": h_g,
        "w_glass": w_g,  # алиас
        "h_glass": h_g,
        "n_lp": n_lp,
        "lock_points": n_lp,  # алиас
        "total_imposts": total_imposts,
        "imp_vertical": imp_vertical_left + imp_vertical_center + imp_vertical_right,
        "imp_horizontal": imp_horizontal_top,
        "area_m2": area,
        "perimeter_m": perimeter
    }

def calculate_window_smeta(order_data: Dict, ref1: List, ref2: Dict, ref3: List) -> Dict:
    """
    Полный расчет сметы для окон
    
    ИСПРАВЛЕНИЯ V2:
    - Маппинг систем профиля
    - Правильный расчет площади и периметра
    - Световой проем по системе
    - Группировка материалов по типу изделия
    """
    
    common = order_data["common"]
    
    result = {
        "metrics": {
            "total_area": 0.0,
            "total_perimeter": 0.0
        },
        "part1_gabarits": [],  # Габаритная ведомость (детальная)
        "part1_summary": {},   # Габаритная ведомость (общая по типам)
        "part2_materials": [],  # Материалы
        "part3_final": {},  # Итоговый расчет
        "total_with_margin": 0.0,
        "debug_info": []  # Отладочная информация
    }
    
    # ===== ОБРАБОТКА ПОЗИЦИЙ =====
    positions = order_data.get("positions", [])
    
    all_contexts = []  # Все контексты для расчета
    materials_by_position = {}  # Материалы по позициям
    
    for pos_idx, position in enumerate(positions):
        pos_data = position.get("data", {})
        pos_count = position.get("count", 1)
        pos_type = position.get("product_type", common.get("main_type", "Окно с откр."))
        pos_system = position.get("system_id", common.get("system_id", "ALG 2030-73C"))
        
        # Нормализуем систему профиля
        normalized_system = normalize_system(pos_system)
        
        result["debug_info"].append({
            "position": pos_idx + 1,
            "original_system": pos_system,
            "normalized_system": normalized_system,
            "product_type": pos_type
        })
        
        # Рассчитываем геометрию позиции
        context = calculate_window_geometry(pos_data, normalized_system)
        context["count"] = pos_count
        all_contexts.append(context)
        
        # Суммируем метрики
        result["metrics"]["total_area"] += context["area_m2"]
        result["metrics"]["total_perimeter"] += context["perimeter_m"]
        
        # ===== ЧАСТЬ 1: ГАБАРИТНАЯ ВЕДОМОСТЬ (Справочник-3) =====
        for row in ref3:
            row_type = str(row.get("Тип изделия", "")).strip()
            
            if row_type == pos_type:
                formula = row.get("Формула_Python", "0")
                val = safe_eval(formula, context)
                
                if val > 0:  # Добавляем только ненулевые
                    element_type = row.get("Тип элемента", "Прочее")
                    element_name = row.get("тип элемент", "Не указано")
                    
                    # Детальная ведомость
                    result["part1_gabarits"].append({
                        "Позиция": f"№{pos_idx + 1}",
                        "Тип изделия": pos_type,
                        "Категория": element_type,
                        "Элемент": element_name,
                        "Значение": round(val, 2)
                    })
                    
                    # Суммарная ведомость по типам
                    summary_key = f"{pos_type}|{element_type}|{element_name}"
                    if summary_key not in result["part1_summary"]:
                        result["part1_summary"][summary_key] = {
                            "Тип изделия": pos_type,
                            "Категория": element_type,
                            "Элемент": element_name,
                            "Значение": 0.0
                        }
                    result["part1_summary"][summary_key]["Значение"] += val
        
        # Сохраняем контекст для расчета материалов
        materials_by_position[pos_idx] = {
            "type": pos_type,
            "system": normalized_system,
            "context": context
        }
    
    # ===== ЧАСТЬ 2: МАТЕРИАЛЫ (Справочник-1) =====
    materials_sum = 0.0
    materials_found_count = 0
    
    # Группируем материалы по типу изделия и системе
    for pos_idx, mat_data in materials_by_position.items():
        target_type = mat_data["type"]
        target_sys = mat_data["system"]
        context = mat_data["context"]
        
        for row in ref1:
            row_type = str(row.get("Тип изделия", "")).strip()
            row_sys_raw = str(row.get("Система профиля", "")).strip()
            row_sys = normalize_system(row_sys_raw)
            
            if row_type == target_type and row_sys == target_sys:
                formula = row.get("Формула_Python", "0")
                qty_fact = safe_eval(formula, context)
                
                # Норма упаковки
                norm_str = str(row.get("кол-во норм к упаковке", "1")).replace(",", ".").replace(" ", "").replace("\xa0", "")
                try:
                    norm = float(norm_str) if norm_str else 1.0
                except:
                    norm = 1.0
                
                # Количество к отгрузке (округление вверх)
                qty_ship = math.ceil(qty_fact / norm) if norm > 0 else math.ceil(qty_fact)
                
                # Цена
                price_str = str(row.get("цена за ед ", "0")).replace(",", ".").replace(" ", "").replace("\xa0", "")
                try:
                    price = float(price_str) if price_str else 0.0
                except:
                    price = 0.0
                
                # Сумма
                row_sum = qty_ship * price
                materials_sum += row_sum
                
                if qty_fact > 0:  # Добавляем только с расходом
                    materials_found_count += 1
                    result["part2_materials"].append({
                        "Позиция": f"№{pos_idx + 1}",
                        "Тип изделия": target_type,
                        "Система": target_sys,
                        "Товар": row.get("Товар", ""),
                        "Артикул": row.get("Артикул", ""),
                        "Тип элемента": row.get("Тип элемента", ""),
                        "Цена": price,
                        "Ед.": row.get("Ед.", "шт"),
                        "Расход факт.": round(qty_fact, 2),
                        "Норма": norm,
                        "К отгрузке": qty_ship,
                        "Сумма": round(row_sum, 0)
                    })
    
    result["debug_info"].append({
        "materials_found": materials_found_count,
        "materials_sum": materials_sum
    })
    
    # ===== ЧАСТЬ 3: ИТОГОВЫЙ РАСЧЕТ =====
    
    # Функция получения цены из Справочника-2
    def get_price_from_ref2(key_word: str, default: float = 0.0) -> float:
        """Ищет цену в Справочнике-2 по ключевому слову"""
        for k, v in ref2.items():
            if key_word.lower() in str(k).lower():
                try:
                    # Берем первое значение из словаря
                    if isinstance(v, dict):
                        val = list(v.values())[0]
                    else:
                        val = v
                    return float(str(val).replace(" ", "").replace(",", ".").replace("\xa0", ""))
                except:
                    return default
        return default
    
    total_area = result["metrics"]["total_area"]
    
    # Стеклопакет / Ламбри
    cost_glass_lambri = 0.0
    
    # Проверяем каждую позицию на тип заполнения
    for position in positions:
        pos_data = position.get("data", {})
        fill_cat = pos_data.get("fill_category", "Стеклопакет")
        glass_type = pos_data.get("glass_type", "Двойной")
        
        W = pos_data.get("width", 0)
        H = pos_data.get("height", 0)
        pos_count = position.get("count", 1)
        pos_area = (W * H / 1_000_000) * pos_count
        
        if fill_cat == "Стеклопакет":
            # Ищем цену стеклопакета по типу
            price_glass = get_price_from_ref2(glass_type, 9000)
            cost_glass_lambri += pos_area * price_glass
        elif "Ламбри" in fill_cat:
            # Ищем цену ламбри
            price_lambri = get_price_from_ref2(fill_cat, 2248)
            cost_glass_lambri += pos_area * price_lambri
    
    # Тонировка
    cost_toning = 0.0
    if common.get("toning_id") == "Есть":
        price_toning = get_price_from_ref2("Тонировка", 2000)
        cost_toning = total_area * price_toning
    
    # Сборка
    cost_assembly = 0.0
    if common.get("assembly_id") == "Есть":
        price_assembly = get_price_from_ref2("Сборка", 10000)
        cost_assembly = total_area * price_assembly
    
    # Монтаж
    cost_installation = 0.0
    if common.get("installation_id") != "Нет":
        price_installation = get_price_from_ref2("Монтаж", 10000)
        cost_installation = total_area * price_installation
    
    # Итого
    result["part3_final"] = {
        "Стеклопакет / Ламбри": round(cost_glass_lambri, 0),
        "Тонировка": round(cost_toning, 0),
        "Сборка": round(cost_assembly, 0),
        "Монтаж": round(cost_installation, 0),
        "Материалы": round(materials_sum, 0)
    }
    
    # Итого без наценки
    subtotal = sum(result["part3_final"].values())
    
    # Наценка 65%
    margin = subtotal * 0.65
    result["part3_final"]["Обеспечение (65%)"] = round(margin, 0)
    
    # Итого к оплате
    result["total_with_margin"] = round(subtotal + margin, 0)
    
    # Округляем метрики
    result["metrics"]["total_area"] = round(result["metrics"]["total_area"], 3)
    result["metrics"]["total_perimeter"] = round(result["metrics"]["total_perimeter"], 3)
    
    # Преобразуем суммарную ведомость в список
    result["part1_summary"] = list(result["part1_summary"].values())
    
    # Округляем значения в суммарной ведомости
    for item in result["part1_summary"]:
        item["Значение"] = round(item["Значение"], 2)
    
    return result
