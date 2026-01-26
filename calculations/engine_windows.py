import math
import logging
from typing import Dict, List

# ✅ УНИФИКАЦИЯ V8: Импорт нового расчётного модуля
from calculations.adapter import calculate_window_smeta_unified

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

def safe_float(value, default=0.0):
    """Безопасное преобразование в float"""
    try:
        if value is None:
            return default
        s = str(value).replace(",", ".").replace(" ", "").replace("\xa0", "")
        if s == "":
            return default
        return float(s)
    except:
        return default

def calculate_window_geometry(position_data: Dict, system: str = "ALG 2030-73C") -> Dict:
    """
    Расчет геометрии окна/двери
    
    ИСПРАВЛЕНИЯ V5:
    - Поддержка импостов как словаря {"left": ..., "tor": ...} И как списка
    - Учитываем систему профиля для светового проёма
    - imp_vertical и imp_horizontal для формулы рамы
    - area и perimeter БЕЗ умножения на count (count применяется в формулах!)
    """
    W = position_data.get("width", 0) / 1000  # в метры
    H = position_data.get("height", 0) / 1000
    count = position_data.get("count", 1)
    
    # Нормализуем систему
    normalized_system = normalize_system(system)
    offset_mm = SYSTEM_OFFSETS.get(normalized_system, 73)
    offset_m = offset_mm / 1000  # в метры
    
    # Очищаем тип изделия
    product_raw = position_data.get("product_type", "")
    p_type = " ".join(product_raw.split()).strip()
    is_blind_window = (p_type == "Окно глух.")
    
    # === ИМПОСТЫ: поддержка СЛОВАРЯ, СПИСКА И AUTO_CALCULATE ===
    imposts = position_data.get("imposts", {})
    
    # Проверяем формат импостов
    if isinstance(imposts, dict):
        # НОВЫЙ ФОРМАТ: auto_calculate из app.py
        if imposts.get("auto_calculate"):
            # Автоматический расчёт длин импостов
            imp_left_mm = (H * 1000 - 2 * offset_mm) if imposts.get("has_left") else 0
            imp_center_mm = (H * 1000 - 2 * offset_mm) if imposts.get("has_center") else 0
            imp_right_mm = (H * 1000 - 2 * offset_mm) if imposts.get("has_right") else 0
            imp_tor_mm = (W * 1000 - 2 * offset_mm) if imposts.get("has_tor") else 0
        else:
            # СТАРЫЙ ФОРМАТ: {"left": 500, "center": 0, "right": 0, "tor": 560}
            imp_left_mm = imposts.get("left", 0)
            imp_center_mm = imposts.get("center", 0)
            imp_right_mm = imposts.get("right", 0)
            imp_tor_mm = imposts.get("tor", 0)
    elif isinstance(imposts, (list, tuple)):
        # Формат: [left, center, right, tor]
        imp_left_mm = imposts[0] if len(imposts) > 0 else 0
        imp_center_mm = imposts[1] if len(imposts) > 1 else 0
        imp_right_mm = imposts[2] if len(imposts) > 2 else 0
        imp_tor_mm = imposts[3] if len(imposts) > 3 else 0
    else:
        # Нет импостов
        imp_left_mm = 0
        imp_center_mm = 0
        imp_right_mm = 0
        imp_tor_mm = 0
    
    # Считаем ДЛИНЫ импостов для формулы рамы (в метрах)
    imp_vertical = 0
    if imp_left_mm > 0:
        imp_vertical += (H - 2 * offset_m)
    if imp_center_mm > 0:
        imp_vertical += (H - 2 * offset_m)
    if imp_right_mm > 0:
        imp_vertical += (H - 2 * offset_m)
    
    imp_horizontal = 0
    if imp_tor_mm > 0:
        imp_horizontal += (W - 2 * offset_m)
    
    total_imposts = imp_left_mm + imp_center_mm + imp_right_mm + imp_tor_mm
    
    # Створки
    sashes = position_data.get("sashes", [])
    n_sash = len(sashes) if sashes else 0  # КОЛИЧЕСТВО СТВОРОК
    
    # ИСПРАВЛЕНИЕ: Для глухого окна n_sash минимум 1 (для формул штапика и уплотнителя)
    if n_sash == 0 and W > 0 and H > 0:
        n_sash = 1
    
    if sashes:
        w_s = sashes[0].get("w", 0) / 1000  # в метры (первая створка)
        h_s = sashes[0].get("h", 0) / 1000
    else:
        w_s = 0
        h_s = 0
    
    # ДЛЯ ДВЕРЕЙ: Суммарная ширина всех створок (w_s_total)
    w_s_total = sum(s.get("w", 0) for s in sashes) / 1000  # в метрах
    
    # Алиасы для формул (некоторые формулы используют строчные w, h)
    w = W
    h = H
    
    # Световой проем (по системе профиля!)
    # Для глухого окна световой проём = габарит минус отступы системы
    w_g = max(W - 2 * offset_m, 0)
    h_g = max(H - 2 * offset_m, 0)
    
    # Количество точек запирания
    if h_s * 1000 < 1200:
        n_lp = 2
    elif h_s * 1000 < 2000:
        n_lp = 3
    else:
        n_lp = 4
    
    # ВАЖНО: Площадь и периметр БЕЗ умножения на count!
    # count применяется В ФОРМУЛАХ Справочника-1 и Справочника-3
    area_single = W * H
    perimeter_single = 2 * (W + H)
    
    return {
        "W": W,
        "H": H,
        "count": count,
        "n_sash": n_sash,  # КОЛИЧЕСТВО СТВОРОК для формул штапика/уплотнителей
        "w_s": w_s,
        "h_s": h_s,
        "w_s_total": w_s_total,  # ДЛЯ ДВЕРЕЙ: Суммарная ширина створок
        "w": w,  # Алиас для W (строчная)
        "h": h,  # Алиас для H (строчная)
        "w_stvor": w_s,
        "h_stvor": h_s,
        "w_g": w_g,
        "h_g": h_g,
        "w_glass": w_g,
        "h_glass": h_g,
        "n_lp": n_lp,
        "lock_points": n_lp,
        "total_imposts": total_imposts,
        "imp_vertical": imp_vertical,  # ← ДЛЯ ФОРМУЛЫ РАМЫ!
        "imp_horizontal": imp_horizontal,  # ← ДЛЯ ФОРМУЛЫ РАМЫ!
        "area_m2": area_single,  # БЕЗ count!
        "perimeter_m": perimeter_single,  # БЕЗ count!
        "area": area_single,
        "perimeter": perimeter_single,
        "qty": count,
        "Nwin": count
    }

def calculate_window_smeta(order_data: Dict, ref1: List, ref2: Dict, ref3: List) -> Dict:
    """
    Унифицированный расчёт через новую модель
    Обеспечивает обратную совместимость
    
    ИСПРАВЛЕНИЯ V8 (Унификация):
    - Использует новую архитектуру без хардкодов
    - Геометрия изделия не зависит от контекста
    - Рама всегда имеет 4 стороны
    - Уплотнители по полному периметру
    - Standalone = Embedded (идентичный расчёт)
    
    Для fallback на старую логику (если новая не работает):
    используйте calculate_window_smeta_legacy() ниже
    """
    return calculate_window_smeta_unified(order_data, ref1, ref2, ref3)


def calculate_window_smeta_legacy(order_data: Dict, ref1: List, ref2: Dict, ref3: List) -> Dict:
    """
    ✅ ИСПРАВЛЕННАЯ ВЕРСИЯ V9 - УСТРАНЕНИЕ ДВОЙНОЙ МАРЖИ
    
    Полный расчет сметы для окон и дверей
    
    КЛЮЧЕВОЕ ИЗМЕНЕНИЕ:
    - НЕ начисляет обеспечение (81%) внутри функции
    - Возвращает ТОЛЬКО себестоимость материалов в поле 'materials_cost'
    - Обеспечение начисляется ОДИН РАЗ на уровне всего заказа в app.py
    
    ПОДДЕРЖИВАЕМЫЕ ТИПЫ ИЗДЕЛИЙ:
    - Окно с откр.
    - Окно глух.
    - Дверь 2-х створч.
    - Дверь 1 створч.
    
    Args:
        order_data: {
            "positions": [{
                "data": {
                    "width": int,
                    "height": int,
                    "product_type": str,
                    "imposts": dict,
                    "sashes": list
                },
                "count": int
            }],
            "common": {
                "system": str,
                "fill_category": str,
                "glass_type": str,
                "toning": str,
                "assembly": str,
                "installation": str
            }
        }
        ref1: Справочник-1 (материалы)
        ref2: Справочник-2 (услуги, стеклопакет)
        ref3: Справочник-3 (габаритная ведомость)
        
    Returns:
        {
            "part1_gabarits": [{Элемент, Значение}],
            "part2_materials": [{Артикул, Элемент, Количество, Цена, Стоимость}],
            "part3_final": {
                "Стеклопакет": float,
                "Ламбри": float,
                "Тонировка": float,
                "Сборка": float,
                "Монтаж": float,
                "Дополнительные детали": float,
                "Материалы": float
            },
            "materials_cost": float,  # ← ТОЛЬКО СЕБЕСТОИМОСТЬ (без обеспечения!)
            "total_with_margin": float,  # ← DEPRECATED: для обратной совместимости
            "metrics": {
                "total_area": float,
                "total_perimeter": float
            }
        }
    """
    
    positions = order_data.get("positions", [])
    common = order_data.get("common", {})
    
    if not positions:
        return {"error": "Нет позиций для расчёта"}
    
    system = common.get("system") or common.get("system_id", "ALG 2030-73C")
    
    print("\n" + "="*70)
    print("🔧 ПОЛНЫЙ РАСЧЁТ МАТЕРИАЛОВ (БЕЗ ДВОЙНОЙ МАРЖИ)")
    print("="*70)
    print(f"Тип: window")
    print(f"Система: {system}")
    
    result = {
        "part1_gabarits": [],
        "part1_summary": [],  # Алиас
        "part2_materials": [],
        "part3_final": {},
        "metrics": {"total_area": 0, "total_perimeter": 0}
    }
    
    materials_sum = 0
    all_contexts = []
    
    # ===== ОБРАБОТКА КАЖДОЙ ПОЗИЦИИ =====
    
    for pos_idx, position in enumerate(positions):
        pos_data = position.get("data", {})
        pos_count = position.get("count", 1)
        
        # Геометрия позиции
        context = calculate_window_geometry(pos_data, system)
        context["count"] = pos_count
        all_contexts.append(context)
        
        print(f"\n--- ПОЗИЦИЯ {pos_idx + 1} ---")
        print(f"   Габариты: {context['W']:.2f}м × {context['H']:.2f}м")
        print(f"   Количество: {pos_count}")
        
        # Определяем CODE для поиска материалов
        product_type = pos_data.get("product_type", "")
        
        from calculations.mapping import get_code_for_windows_doors
        code = get_code_for_windows_doors(product_type, system)
        
        print(f"   🔑 CODE: {code}")
        
        # Ищем материалы в Справочнике-1
        print(f"\n🔍 Поиск материалов по CODE={code}...")
        
        materials_for_position = []
        for item in ref1:
            item_code = item.get("CODE", "")
            if item_code == code:
                materials_for_position.append(item)
        
        print(f"Найдено: {len(materials_for_position)} позиций")
        
        if not materials_for_position:
            print(f"⚠️ ВНИМАНИЕ: Материалы для CODE={code} не найдены!")
            continue
        
        # Считаем материалы
        print("\n💰 Округление до упаковок:")
        
        for material in materials_for_position:
            article = material.get("Артикул", "")
            element = material.get("Элемент", "")
            unit = material.get("Единица", "")
            price = safe_float(material.get("Цена за единицу", 0))
            formula_raw = material.get("Формула", "1")
            pack_size = safe_float(material.get("Кратность", 1))
            
            # Вычисляем количество через формулу
            try:
                qty_calc = safe_eval(formula_raw, context)
            except Exception as e:
                print(f"⚠️ Ошибка в формуле '{formula_raw}': {e}")
                qty_calc = 0
            
            if qty_calc <= 0:
                continue
            
            # Округляем до упаковок
            if pack_size > 0:
                qty_packs = math.ceil(qty_calc / pack_size)
                qty_final = qty_packs * pack_size
            else:
                qty_final = qty_calc
            
            cost = qty_final * price
            materials_sum += cost
            
            print(f"   {element}: {qty_calc:.2f}{unit} → {qty_packs} упак × {pack_size}{unit} = {cost:,.0f}₸")
            
            result["part2_materials"].append({
                "Артикул": article,
                "Элемент": element,
                "Количество": round(qty_final, 3),
                "Количество_raw": round(qty_calc, 3),  # ДО округления - для глобальной корзины
                "Единица": unit,
                "Цена": price,
                "Стоимость": round(cost, 0)
            })
    
    print(f"\nИТОГО МАТЕРИАЛЫ: {materials_sum:,.0f}₸")
    
    # ===== ГАБАРИТНАЯ ВЕДОМОСТЬ (Справочник-3) =====
    
    print("\n" + "="*70)
    print("📊 ГАБАРИТНАЯ ВЕДОМОСТЬ (Справочник-3)")
    print("="*70)
    
    for gab_item in ref3:
        elem_name = gab_item.get("Элемент", "")
        formula_raw = gab_item.get("Формула", "0")
        unit = gab_item.get("Единица", "")
        
        # Считаем по всем позициям
        total_value = 0
        for context in all_contexts:
            try:
                val = safe_eval(formula_raw, context)
                total_value += val
            except:
                pass
        
        print(f"{elem_name}: {total_value:.3f} {unit}")
        
        result["part1_gabarits"].append({
            "Элемент": elem_name,
            "Значение": round(total_value, 3),
            "Единица": unit
        })
    
    # Извлекаем метрики из габаритной ведомости
    total_area_calc = 0
    total_perimeter_calc = 0
    
    for item in result["part1_gabarits"]:
        elem = str(item.get("Элемент", "")).lower()
        val = item.get("Значение", 0)
        
        # Ищем площадь
        if "площадь" in elem or "area" in elem:
            total_area_calc += val
        
        # Ищем периметр
        if "периметр" in elem or "perimeter" in elem:
            total_perimeter_calc += val
    
    # Если в габаритной ведомости не нашли, считаем из позиций
    if total_area_calc == 0:
        for context in all_contexts:
            total_area_calc += context["area_m2"] * context["count"]
    
    if total_perimeter_calc == 0:
        for context in all_contexts:
            total_perimeter_calc += context["perimeter_m"] * context["count"]
    
    result["metrics"]["total_area"] = total_area_calc
    result["metrics"]["total_perimeter"] = total_perimeter_calc
    
    # ===== ИТОГОВЫЙ РАСЧЕТ =====
    
    def get_price_from_ref2(key_word: str) -> float:
        """Поиск цены в Справочнике-2"""
        # Нормализуем: приводим к нижнему регистру и убираем пробелы вокруг /
        key_normalized = key_word.lower().replace(" / ", "/")
        price = ref2.get(key_normalized)
        
        # Запасные значения если нет в справочнике
        if price is None:
            defaults = {
                "ламбри без термо": 2248,
                "ламбри с термо": 2800
            }
            price = defaults.get(key_normalized)
        
        print(f"🔎 Ищем: '{key_word}' → normalized: '{key_normalized}' → найдено: {price}")
        
        if price is None:
            print(f"⚠️ WARNING: Цена НЕ НАЙДЕНА!")
            return 0.0
        return float(price)
    
    total_area = result["metrics"]["total_area"]
    
    # ИСПРАВЛЕНО: Разделяем стеклопакет и ламбри
    cost_glass = 0.0
    cost_lambri = 0.0
    
    print(f"\n{'='*60}")
    print(f"🔍 ДИАГНОСТИКА РАСЧЁТА СТЕКЛОПАКЕТА И ЛАМБРИ")
    print(f"{'='*60}")
    print(f"Всего позиций: {len(positions)}")
    
    for pos_idx, position in enumerate(positions):
        pos_data = position.get("data", {})
        # ИСПРАВЛЕНО: Читаем из двух мест (для тамбура данные в position напрямую)
        fill_cat = pos_data.get("fill_category") or position.get("fill_category", "Стеклопакет")
        glass_type = pos_data.get("glass_type") or position.get("glass_type", "Двойной")
        
        # ИСПРАВЛЕНО: Нормализуем чтение размеров - поддержка разных ключей
        W = pos_data.get("width", 0)
        if W == 0:
            W = position.get("width", 0)  # Пробуем читать из корня position
        # Умное определение: если < 100, значит уже в метрах!
        W = W / 1000 if W >= 100 else W
        
        H = pos_data.get("height", 0)
        if H == 0:
            H = position.get("height", 0)  # Пробуем читать из корня position
        # Умное определение: если < 100, значит уже в метрах!
        H = H / 1000 if H >= 100 else H
        
        pos_count = position.get("count", 1)
        pos_area = W * H * pos_count
        
        print(f"\n📦 Позиция {pos_idx + 1}:")
        print(f"   Тип заполнения: {fill_cat}")
        print(f"   Тип стекла: {glass_type}")
        print(f"   Размеры: W={W:.3f}м × H={H:.3f}м")
        print(f"   Количество: {pos_count}")
        print(f"   Площадь: {pos_area:.3f} м²")
        
        if fill_cat == "Стеклопакет":
            # Стеклопакет: площадь × цена за м²
            price_glass = get_price_from_ref2(glass_type)
            cost = pos_area * price_glass
            cost_glass += cost
            print(f"   ✅ Стеклопакет: {pos_area:.3f} м² × {price_glass} тг/м² = {cost:.2f} тг")
        elif "Ламбри" in fill_cat:
            # Ламбри: округляем до кратного 6м (хлысты), потом × цена за 1м
            price_lambri = get_price_from_ref2(fill_cat)
            # Округляем площадь до кратного 6 (завод отпускает хлыстами по 6м)
            qty_hlysti = math.ceil(pos_area / 6) if pos_area > 0 else 0
            total_meters = qty_hlysti * 6
            cost = total_meters * price_lambri
            cost_lambri += cost
            print(f"   ✅ Ламбри: {pos_area:.3f} м² → {qty_hlysti} хлыстов × 6м = {total_meters}м × {price_lambri} тг/м = {cost:.2f} тг")
        else:
            print(f"   ⚠️ Неизвестный тип заполнения: {fill_cat}")
    
    print(f"\n{'='*60}")
    print(f"📊 ИТОГО:")
    print(f"   Стеклопакет: {cost_glass:.2f} тг")
    print(f"   Ламбри: {cost_lambri:.2f} тг")
    print(f"{'='*60}\n")
    
    # Тонировка
    cost_toning = 0.0
    toning = common.get("toning_id") or common.get("toning", "Нет")
    if toning == "Есть":
        price_toning = get_price_from_ref2("Тонировка")
        cost_toning = total_area * price_toning
    
    # Сборка
    cost_assembly = 0.0
    assembly = common.get("assembly_id") or common.get("assembly", "Нет")
    if assembly == "Есть":
        price_assembly = get_price_from_ref2("Сборка")
        cost_assembly = total_area * price_assembly
    
    # Монтаж
    cost_installation = 0.0
    installation = common.get("installation_id") or common.get("installation", "Нет")
    
    print(f"\n🔧 Расчёт монтажа: '{installation}'")
    
    if installation != "Нет":
        # ИСПРАВЛЕНО: ищем цену по конкретному типу монтажа
        # Нормализуем название - убираем лишние пробелы
        installation_clean = " ".join(installation.split())
        price_installation = get_price_from_ref2(installation_clean)
        cost_installation = total_area * price_installation
        print(f"   ✅ {installation_clean}: {total_area:.3f} м² × {price_installation} тг/м² = {cost_installation:.2f} тг")
    else:
        print(f"   ⏭️ Монтаж не требуется")
    
    # ДОБАВЛЕНО: Дополнительные детали
    cost_additional = 0.0
    print(f"\n🔧 Расчёт дополнительных деталей:")
    
    # Берём периметр из метрик
    total_perimeter = result["metrics"]["total_perimeter"]
    
    # Ищем "Нащельник" в ref2
    additional_name = None
    for key in ref2.keys():
        if "нащельник" in key.lower():
            additional_name = key
            break
    
    if additional_name:
        price_additional = ref2.get(additional_name, 0)
        # Формула: ОКРУГЛЕНИЕ ВВЕРХ (периметр / 3) * цена
        cost_additional = math.ceil(total_perimeter / 3) * price_additional
        print(f"   Формула: ⌈периметр / 3⌉ × цена")
        print(f"   Расчёт: ⌈{total_perimeter:.3f} / 3⌉ × {price_additional} = {math.ceil(total_perimeter / 3)} × {price_additional} = {cost_additional:.2f} тг")
    else:
        print(f"   ⚠️ 'Нащельник' не найден в Справочнике-2")
    
    # ===== КРИТИЧНО: УБРАНА ДВОЙНАЯ МАРЖА! =====
    # Раньше было:
    # subtotal = materials_sum + cost_glass + cost_lambri + cost_toning + cost_assembly + cost_installation + cost_additional
    # margin = subtotal * 0.81
    # total_with_margin = subtotal + margin
    #
    # Теперь:
    # Обеспечение начисляется ОДИН РАЗ на уровне всего заказа в app.py
    
    result["part3_final"] = {
        "Стеклопакет": round(cost_glass, 0),
        "Ламбри": round(cost_lambri, 0),
        "Тонировка": round(cost_toning, 0),
        "Сборка": round(cost_assembly, 0),
        "Монтаж": round(cost_installation, 0),
        "Дополнительные детали": round(cost_additional, 0),
        "Материалы": round(materials_sum, 0)
    }
    
    # КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: materials_cost теперь ТОЛЬКО себестоимость без маржи
    materials_cost_only = materials_sum + cost_glass + cost_lambri + cost_toning + cost_assembly + cost_installation + cost_additional
    
    result["materials_cost"] = round(materials_cost_only, 0)  # ← ТОЛЬКО СЕБЕСТОИМОСТЬ!
    
    # Для обратной совместимости оставляем total_with_margin, но помечаем как DEPRECATED
    result["total_with_margin"] = round(materials_cost_only, 0)  # DEPRECATED: используйте materials_cost
    
    result["metrics"]["total_area"] = round(result["metrics"]["total_area"], 3)
    result["metrics"]["total_perimeter"] = round(result["metrics"]["total_perimeter"], 3)
    
    # === АЛИАСЫ ДЛЯ СОВМЕСТИМОСТИ С app.py ===
    result["part1_summary"] = result["part1_gabarits"]
    
    print("\n" + "="*70)
    print("✅ РАСЧЁТ ЗАВЕРШЁН (БЕЗ ДВОЙНОЙ МАРЖИ)")
    print(f"   Себестоимость материалов: {materials_cost_only:,.0f}₸")
    print(f"   Обеспечение (81%) начисляется на уровне заказа в app.py")
    print("="*70)
    
    return result

def calculate_impost_length(width_mm: float, height_mm: float, system: str, direction: str) -> float:
    """
    Вспомогательная функция для расчёта длины одного импоста
    Используется в app.py для отображения информации
    
    Args:
        width_mm: Ширина изделия в мм
        height_mm: Высота изделия в мм
        system: Система профиля
        direction: "vertical" или "horizontal"
    
    Returns:
        Длина импоста в мм
    """
    W = width_mm / 1000  # в метры
    H = height_mm / 1000
    
    normalized_system = normalize_system(system)
    offset_mm = SYSTEM_OFFSETS.get(normalized_system, 73)
    offset_m = offset_mm / 1000
    
    if direction == "vertical":
        # Вертикальный импост
        return (H - 2 * offset_m) * 1000  # возвращаем в мм
    elif direction == "horizontal":
        # Горизонтальный импост (ТОР)
        return (W - 2 * offset_m) * 1000  # возвращаем в мм
    else:
        return 0.0
