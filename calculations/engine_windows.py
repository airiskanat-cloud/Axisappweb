import math
import logging
from typing import Dict, List

logger = logging.getLogger(__name__)

def safe_eval(formula: str, context: dict) -> float:
    """Безопасное вычисление формул Python"""
    try:
        f = str(formula).replace(",", ".").replace(" ", "")
        return float(eval(f, {"__builtins__": None, "math": math}, context))
    except Exception as e:
        logger.warning(f"Ошибка в формуле '{formula}': {e}")
        return 0.0

def calculate_window_geometry(position_data: Dict) -> Dict:
    """
    Расчет геометрии окна/двери
    Возвращает все необходимые переменные для формул
    """
    W = position_data.get("width", 0)
    H = position_data.get("height", 0)
    count = position_data.get("count", 1)
    
    # Импосты
    imposts = position_data.get("imposts", [0, 0, 0, 0])
    imp_left = imposts[0] if len(imposts) > 0 else 0
    imp_center = imposts[1] if len(imposts) > 1 else 0
    imp_right = imposts[2] if len(imposts) > 2 else 0
    imp_tor = imposts[3] if len(imposts) > 3 else 0
    total_imposts = imp_left + imp_center + imp_right + imp_tor
    
    # Створки
    sashes = position_data.get("sashes", [])
    
    # Если есть створки - берем размеры первой для формул
    if sashes:
        w_s = sashes[0].get("w", 0)
        h_s = sashes[0].get("h", 0)
    else:
        w_s = 0
        h_s = 0
    
    # Световой проем (для штапика)
    # Упрощенный расчет: минус 100 мм с каждой стороны
    w_g = max(W - 100, 0)
    h_g = max(H - 100, 0)
    
    # Количество точек запирания (по высоте створки)
    if h_s < 1200:
        n_lp = 2
    elif h_s < 2000:
        n_lp = 3
    else:
        n_lp = 4
    
    # Площадь
    area = (W * H / 1_000_000) * count
    
    # Периметр
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
        "area_m2": area,
        "perimeter_m": perimeter
    }

def calculate_window_smeta(order_data: Dict, ref1: List, ref2: Dict, ref3: List) -> Dict:
    """
    Полный расчет сметы для окон
    
    order_data = {
        "common": {
            "order_number": "001",
            "main_type": "Окно с откр.",
            "system_id": "ALG 2030-63C",
            "toning_id": "Нет",
            "assembly_id": "Есть",
            "installation_id": "Монтаж"
        },
        "positions": [
            {
                "count": 2,
                "data": {
                    "width": 1500,
                    "height": 1200,
                    "imposts": [500, 0, 500, 0],
                    "sashes": [{"w": 750, "h": 1100}],
                    "fill_category": "Стеклопакет",
                    "glass_type": "Двойной"
                }
            }
        ]
    }
    """
    
    common = order_data["common"]
    target_type = common["main_type"]
    target_sys = common["system_id"]
    
    result = {
        "metrics": {
            "total_area": 0.0,
            "total_perimeter": 0.0
        },
        "part1_gabarits": [],  # Габаритная ведомость
        "part2_materials": [],  # Материалы
        "part3_final": {},  # Итоговый расчет
        "total_with_margin": 0.0
    }
    
    # ===== ОБРАБОТКА ПОЗИЦИЙ =====
    positions = order_data.get("positions", [])
    
    all_contexts = []  # Все контексты для суммирования
    
    for pos_idx, position in enumerate(positions):
        pos_data = position.get("data", {})
        pos_data["count"] = position.get("count", 1)
        
        # Рассчитываем геометрию позиции
        context = calculate_window_geometry(pos_data)
        all_contexts.append(context)
        
        # Суммируем метрики
        result["metrics"]["total_area"] += context["area_m2"]
        result["metrics"]["total_perimeter"] += context["perimeter_m"]
        
        # ===== ЧАСТЬ 1: ГАБАРИТНАЯ ВЕДОМОСТЬ (Справочник-3) =====
        for row in ref3:
            row_type = str(row.get("Тип изделия", "")).strip()
            
            if row_type == target_type:
                formula = row.get("Формула_Python", "0")
                val = safe_eval(formula, context)
                
                if val > 0:  # Добавляем только ненулевые
                    result["part1_gabarits"].append({
                        "Позиция": f"Позиция №{pos_idx + 1}",
                        "Категория": row.get("Тип элемента", "Прочее"),
                        "Элемент": row.get("тип элемент", "Не указано"),
                        "Значение": round(val, 2)
                    })
    
    # ===== ЧАСТЬ 2: МАТЕРИАЛЫ (Справочник-1) =====
    # Создаем суммарный контекст для расчета материалов
    total_context = {
        "W": sum(c["W"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "H": sum(c["H"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "count": sum(c["count"] for c in all_contexts),
        "w_s": sum(c["w_s"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "h_s": sum(c["h_s"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "w_stvor": sum(c["w_s"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "h_stvor": sum(c["h_s"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "w_g": sum(c["w_g"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "h_g": sum(c["h_g"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "w_glass": sum(c["w_g"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "h_glass": sum(c["h_g"] for c in all_contexts) / len(all_contexts) if all_contexts else 0,
        "n_lp": max(c["n_lp"] for c in all_contexts) if all_contexts else 2,
        "lock_points": max(c["n_lp"] for c in all_contexts) if all_contexts else 2
    }
    
    materials_sum = 0.0
    
    for row in ref1:
        row_type = str(row.get("Тип изделия", "")).strip()
        row_sys = str(row.get("Система профиля", "")).strip()
        
        if row_type == target_type and row_sys == target_sys:
            formula = row.get("Формула_Python", "0")
            qty_fact = safe_eval(formula, total_context)
            
            # Норма упаковки
            norm_str = str(row.get("кол-во норм", "1")).replace(",", ".").replace(" ", "")
            try:
                norm = float(norm_str) if norm_str else 1.0
            except:
                norm = 1.0
            
            # Количество к отгрузке (округление вверх)
            qty_ship = math.ceil(qty_fact / norm) if norm > 0 else qty_fact
            
            # Цена
            price_str = str(row.get("цена за", "0")).replace(",", ".").replace(" ", "")
            try:
                price = float(price_str) if price_str else 0.0
            except:
                price = 0.0
            
            # Сумма
            row_sum = qty_ship * price
            materials_sum += row_sum
            
            if qty_fact > 0:  # Добавляем только с расходом
                result["part2_materials"].append({
                    "Товар": row.get("Товар", ""),
                    "Артикул": row.get("Артикул", ""),
                    "Тип элемента": row.get("Тип элемента", ""),
                    "Цена": price,
                    "Расход факт.": round(qty_fact, 2),
                    "Отгрузка": qty_ship,
                    "Сумма": round(row_sum, 0)
                })
    
    # ===== ЧАСТЬ 3: ИТОГОВЫЙ РАСЧЕТ =====
    
    # Функция получения цены из Справочника-2
    def get_price_from_ref2(key_word: str, default: float = 0.0) -> float:
        """Ищет цену в Справочнике-2 по ключевому слову"""
        for k, v in ref2.items():
            if key_word.lower() in k.lower():
                try:
                    # Берем первое значение из словаря
                    val = list(v.values())[0] if isinstance(v, dict) else v
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
        
        pos_area = (pos_data.get("width", 0) * pos_data.get("height", 0) / 1_000_000) * position.get("count", 1)
        
        if fill_cat == "Стеклопакет":
            # Ищем цену стеклопакета по типу
            price_glass = get_price_from_ref2(glass_type)
            cost_glass_lambri += pos_area * price_glass
        elif "Ламбри" in fill_cat:
            # Ищем цену ламбри
            price_lambri = get_price_from_ref2(fill_cat)
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
    
    return result
