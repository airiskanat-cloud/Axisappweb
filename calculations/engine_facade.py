"""
Модуль расчёта фасадных систем (Ruit 50F)
Полный расчёт профилей, заполнения и вставок (двери/окна)
"""

import math
from typing import Dict, List, Any


def parse_price(value):
    """Безопасное преобразование цены в float"""
    if value is None:
        return 0.0
    value = str(value).strip()
    if value == "":
        return 0.0
    try:
        # Убираем все виды пробелов
        for space in ['\xa0', '\u00a0', '\u202f', '\u2009', ' ']:
            value = value.replace(space, '')
        value = value.replace(',', '.')
        return float(value)
    except:
        return 0.0


def calculate_facade_materials(
    W: float,  # Ширина фасада (м)
    H: float,  # Высота фасада (м)
    cols: int,  # Количество столбцов
    rows: int,  # Количество рядов
    count: int,  # Количество сторон фасада
    inserts: List[Dict],  # Вставки (двери/окна)
    facade_profiles_ref: List[Dict],  # Справочник "Фасады - Профили"
    ref1: List[Dict],  # Справочник-1 (для вставок)
    ref2: Dict[str, float],  # Справочник-2 (цены)
    ref3: List[Dict]  # Справочник-3 (формулы)
) -> Dict[str, Any]:
    """
    Полный расчёт материалов для фасадной системы
    
    Возвращает:
    {
        "skeleton": {...},  # Материалы каркаса
        "inserts": {...},   # Материалы вставок
        "total_cost": float,
        "details": [...]
    }
    """
    
    print("\n" + "="*70)
    print("РАСЧЁТ ФАСАДНОЙ СИСТЕМЫ (Ruit 50F)")
    print("="*70)
    
    # Размеры ячейки
    w_cell = W / cols
    h_cell = H / rows
    count_cells = cols * rows * count
    
    print(f"\nИсходные данные:")
    print(f"  Габариты: {W}м × {H}м × {count} сторон")
    print(f"  Сетка: {cols} столбцов × {rows} рядов")
    print(f"  Ячейка: {w_cell:.2f}м × {h_cell:.2f}м")
    print(f"  Всего ячеек: {count_cells}")
    
    result = {
        "skeleton": {},
        "inserts": {},
        "total_cost": 0,
        "details": []
    }
    
    # ============================================================================
    # ЧАСТЬ 1: КАРКАС ФАСАДА (Стойки, Ригели, Соединители)
    # ============================================================================
    
    print("\n" + "="*70)
    print("ЧАСТЬ 1: КАРКАС ФАСАДА")
    print("="*70)
    
    skeleton_cost = 0
    
    # --- СТОЙКИ ---
    # Автоподбор: H≤3м→90мм; 3-4м→110мм; >4м→130мм
    if H <= 3.0:
        mullion_size = 90
        mullion_article = "2-00-5035"
    elif H <= 4.0:
        mullion_size = 110
        mullion_article = "2-00-5034"
    else:
        mullion_size = 130
        mullion_article = "2-00-5033"
    
    count_m = (cols + 1) * count
    L_m = H * count_m
    sticks_m = math.ceil(L_m / 6.0)
    final_m = sticks_m * 6
    
    # Ищем цену в справочнике
    price_m = 0
    mullion_name = f"Стойка {mullion_size} мм"
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if mullion_name in elem:
            price_m = parse_price(item.get('Цена за единицу', 0))
            break
    
    cost_m = final_m * price_m
    skeleton_cost += cost_m
    
    print(f"\n1. СТОЙКИ {mullion_size}мм ({mullion_article}):")
    print(f"   Формула: H × (cols+1) × count")
    print(f"   Расчёт: {H:.2f}м × {count_m} = {L_m:.2f}м")
    print(f"   Округление: ⌈{L_m:.2f}/6⌉ = {sticks_m} хлыстов = {final_m}м")
    print(f"   Цена: {price_m:,}₸/м")
    print(f"   Стоимость: {cost_m:,.0f}₸")
    
    result["skeleton"][f"Стойка {mullion_size}мм"] = {
        "quantity": final_m,
        "unit": "м",
        "price": price_m,
        "cost": cost_m
    }
    
    # --- РИГЕЛИ ---
    # Глубина на 1 шаг меньше стойки
    if mullion_size == 90:
        transom_size = 50
        transom_article = "2-00-5013"
    elif mullion_size == 110:
        transom_size = 70
        transom_article = "2-00-5019"
    else:
        transom_size = 85
        transom_article = "2-00-5014"
    
    count_r = rows * (cols + 1) * count
    L_r = (w_cell - 0.05) * count_r
    sticks_r = math.ceil(L_r / 6.0)
    final_r = sticks_r * 6
    
    price_r = 0
    transom_name = f"Ригель {transom_size} мм"
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if transom_name in elem:
            price_r = parse_price(item.get('Цена за единицу', 0))
            break
    
    cost_r = final_r * price_r
    skeleton_cost += cost_r
    
    print(f"\n2. РИГЕЛИ {transom_size}мм ({transom_article}):")
    print(f"   Формула: (w_cell - 0.05) × rows × (cols+1) × count")
    print(f"   Расчёт: ({w_cell:.2f} - 0.05) × {count_r} = {L_r:.2f}м")
    print(f"   Округление: {sticks_r} хлыстов = {final_r}м")
    print(f"   Цена: {price_r:,}₸/м")
    print(f"   Стоимость: {cost_r:,.0f}₸")
    
    result["skeleton"][f"Ригель {transom_size}мм"] = {
        "quantity": final_r,
        "unit": "м",
        "price": price_r,
        "cost": cost_r
    }
    
    # --- U-СОЕДИНИТЕЛИ РИГЕЛЯ ---
    count_u = 2 * count_r
    price_u = 151  # Из справочника
    cost_u = count_u * price_u
    skeleton_cost += cost_u
    
    print(f"\n3. U-СОЕДИНИТЕЛЬ РИГЕЛЯ:")
    print(f"   Формула: 2 × count_r")
    print(f"   Количество: {count_u} шт")
    print(f"   Стоимость: {cost_u:,.0f}₸")
    
    result["skeleton"]["U-соединитель"] = {
        "quantity": count_u,
        "unit": "шт",
        "price": price_u,
        "cost": cost_u
    }
    
    # --- ДЕРЖАТЕЛИ СТЕКЛОПАКЕТА ---
    # По 2 шт на каждую ячейку
    count_holders = 2 * count_cells
    price_holder = 0
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if 'Держатель СП' in elem:
            price_holder = parse_price(item.get('Цена за единицу', 0))
            break
    
    if price_holder == 0:
        price_holder = 150  # Запасное значение
    
    cost_holders = count_holders * price_holder
    skeleton_cost += cost_holders
    
    print(f"\n4. ДЕРЖАТЕЛИ СТЕКЛОПАКЕТА:")
    print(f"   Формула: 2 × count_cells")
    print(f"   Количество: {count_holders} шт")
    print(f"   Стоимость: {cost_holders:,.0f}₸")
    
    result["skeleton"]["Держатель СП"] = {
        "quantity": count_holders,
        "unit": "шт",
        "price": price_holder,
        "cost": cost_holders
    }
    
    # --- ТЕРМОМОСТ ---
    total_length = L_m + L_r
    L_th = total_length * count * 1.05  # +5% запас
    sticks_th = math.ceil(L_th / 6.0)
    final_th = sticks_th * 6
    
    price_th = 0
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if 'Термомост' in elem or 'термомост' in elem:
            price_th = parse_price(item.get('Цена за единицу', 0))
            break
    
    if price_th == 0:
        price_th = 800  # Запасное значение
    
    cost_th = final_th * price_th
    skeleton_cost += cost_th
    
    print(f"\n5. ТЕРМОМОСТ 18мм:")
    print(f"   Формула: (L_m + L_r) × count × 1.05")
    print(f"   Расчёт: ({L_m:.2f} + {L_r:.2f}) × {count} × 1.05 = {L_th:.2f}м")
    print(f"   Округление: {sticks_th} хлыстов = {final_th}м")
    print(f"   Стоимость: {cost_th:,.0f}₸")
    
    result["skeleton"]["Термомост"] = {
        "quantity": final_th,
        "unit": "м",
        "price": price_th,
        "cost": cost_th
    }
    
    # --- ПРИЖИМНОЙ ПРОФИЛЬ ---
    L_ext = total_length * count
    sticks_ext = math.ceil(L_ext / 6.0)
    final_ext = sticks_ext * 6
    
    price_ext = 0
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if 'Прижимной профиль' in elem or 'прижимной' in elem.lower():
            price_ext = parse_price(item.get('Цена за единицу', 0))
            break
    
    if price_ext == 0:
        price_ext = 1512  # Запасное значение
    
    cost_ext = final_ext * price_ext
    skeleton_cost += cost_ext
    
    print(f"\n6. ПРИЖИМНОЙ ПРОФИЛЬ:")
    print(f"   Длина: {final_ext}м")
    print(f"   Стоимость: {cost_ext:,.0f}₸")
    
    result["skeleton"]["Прижимной профиль"] = {
        "quantity": final_ext,
        "unit": "м",
        "price": price_ext,
        "cost": cost_ext
    }
    
    # --- КРОНШТЕЙНЫ ---
    count_brackets = 2 * (cols + 1) * count
    price_bracket = 0
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if 'Кронштейн' in elem or 'кронштейн' in elem.lower():
            price_bracket = parse_price(item.get('Цена за единицу', 0))
            break
    
    if price_bracket == 0:
        price_bracket = 500  # Запасное значение
    
    cost_brackets = count_brackets * price_bracket
    skeleton_cost += cost_brackets
    
    print(f"\n7. КРОНШТЕЙНЫ:")
    print(f"   Количество: {count_brackets} шт")
    print(f"   Стоимость: {cost_brackets:,.0f}₸")
    
    result["skeleton"]["Кронштейны"] = {
        "quantity": count_brackets,
        "unit": "шт",
        "price": price_bracket,
        "cost": cost_brackets
    }
    
    print(f"\n{'─'*70}")
    print(f"ИТОГО КАРКАС: {skeleton_cost:,.0f}₸")
    
    # ============================================================================
    # ЧАСТЬ 2: ВСТАВКИ (Двери/Окна ALG)
    # ============================================================================
    
    print("\n" + "="*70)
    print("ЧАСТЬ 2: ВСТАВКИ (ДВЕРИ/ОКНА)")
    print("="*70)
    
    inserts_cost = 0
    
    if not inserts or len(inserts) == 0:
        print("\nВставок нет")
    else:
        for i, insert in enumerate(inserts, 1):
            print(f"\nВставка {i}: {insert.get('type', 'Unknown')} {insert.get('system', 'Unknown')}")
            
            # TODO: Здесь нужен полный расчёт вставки через engine_windows
            # Пока упрощённо
            insert_cost = 250000  # Примерная стоимость двери
            inserts_cost += insert_cost
            
            print(f"   Стоимость: {insert_cost:,.0f}₸")
    
    print(f"\n{'─'*70}")
    print(f"ИТОГО ВСТАВКИ: {inserts_cost:,.0f}₸")
    
    # ============================================================================
    # ИТОГО
    # ============================================================================
    
    total = skeleton_cost + inserts_cost
    
    result["total_cost"] = total
    result["skeleton_cost"] = skeleton_cost
    result["inserts_cost"] = inserts_cost
    
    print("\n" + "="*70)
    print(f"ИТОГО МАТЕРИАЛЫ ФАСАДА: {total:,.0f}₸")
    print("="*70)
    
    return result


def calculate_tambour_materials(
    W: float,
    H: float,
    cols: int,
    rows: int,
    count: int,
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict]
) -> Dict[str, Any]:
    """
    Расчёт материалов для оконного тамбура (сцепка рам ALG)
    
    Применяется для тамбуров, собираемых из готовых дверных/оконных блоков.
    """
    
    print("\n" + "="*70)
    print("РАСЧЁТ ОКОННОГО ТАМБУРА (ALG)")
    print("="*70)
    
    w_cell = W / cols
    h_cell = H / rows
    
    print(f"\nИсходные данные:")
    print(f"  Габариты: {W}м × {H}м")
    print(f"  Сетка: {cols} столбцов × {rows} рядов")
    print(f"  Ячейка: {w_cell:.2f}м × {h_cell:.2f}м")
    print(f"  Количество сторон: {count}")
    
    result = {
        "skeleton": {},
        "total_cost": 0,
        "details": []
    }
    
    skeleton_cost = 0
    
    # --- РАМА (Frame) ---
    # Рассчитывается отдельно для каждой из сторон тамбура
    L_f = (W + H) * 2 * count
    price_frame = 3500  # ~3500₸/м для ALG профиля рамы
    cost_frame = L_f * price_frame
    skeleton_cost += cost_frame
    
    print(f"\n1. РАМА (Frame):")
    print(f"   Формула: (W + H) × 2 × count")
    print(f"   Расчёт: ({W:.2f} + {H:.2f}) × 2 × {count} = {L_f:.2f}м")
    print(f"   Цена: {price_frame:,}₸/м")
    print(f"   Стоимость: {cost_frame:,.0f}₸")
    
    result["skeleton"]["Рама"] = {
        "quantity": L_f,
        "unit": "м",
        "price": price_frame,
        "cost": cost_frame
    }
    
    # --- СОЕДИНИТЕЛЬНАЯ ТРУБА 90° (арт. 2-00-2010) ---
    # Два вертикальных угла по всей высоте
    L_pipe = H * 2 * count
    price_pipe = 2500  # ~2500₸/м
    cost_pipe = L_pipe * price_pipe
    skeleton_cost += cost_pipe
    
    print(f"\n2. СОЕДИНИТЕЛЬНАЯ ТРУБА 90° (2-00-2010):")
    print(f"   Формула: H × 2 × count")
    print(f"   Расчёт: {H:.2f} × 2 × {count} = {L_pipe:.2f}м")
    print(f"   Стоимость: {cost_pipe:,.0f}₸")
    
    result["skeleton"]["Труба соединительная"] = {
        "quantity": L_pipe,
        "unit": "м",
        "price": price_pipe,
        "cost": cost_pipe
    }
    
    # --- АДАПТЕР ТРУБЫ ---
    # По 2 "защёлки" на каждый метр трубы для стыковки с рамами
    L_ada = H * 4 * count
    price_adapter = 800  # ~800₸/м
    cost_adapter = L_ada * price_adapter
    skeleton_cost += cost_adapter
    
    print(f"\n3. АДАПТЕР ТРУБЫ:")
    print(f"   Формула: H × 4 × count")
    print(f"   Расчёт: {H:.2f} × 4 × {count} = {L_ada:.2f}м")
    print(f"   Стоимость: {cost_adapter:,.0f}₸")
    
    result["skeleton"]["Адаптер трубы"] = {
        "quantity": L_ada,
        "unit": "м",
        "price": price_adapter,
        "cost": cost_adapter
    }
    
    # --- НАПРАВЛЯЮЩИЙ (арт. 2-00-5581) ---
    # Принудительно добавляется для соединения изделий
    L_guide = (W + H) * count * 1.05  # +5% запас
    price_guide = 1200  # ~1200₸/м
    cost_guide = L_guide * price_guide
    skeleton_cost += cost_guide
    
    print(f"\n4. НАПРАВЛЯЮЩИЙ (2-00-5581):")
    print(f"   Формула: (W + H) × count × 1.05")
    print(f"   Расчёт: ({W:.2f} + {H:.2f}) × {count} × 1.05 = {L_guide:.2f}м")
    print(f"   Стоимость: {cost_guide:,.0f}₸")
    
    result["skeleton"]["Направляющий"] = {
        "quantity": L_guide,
        "unit": "м",
        "price": price_guide,
        "cost": cost_guide
    }
    
    # --- ШТАПИК (Bead) ---
    # Считается для каждой ячейки заполнения
    n_cells = cols * rows * count
    w_g = w_cell - 0.1  # светопроём
    h_g = h_cell - 0.1
    L_b = (w_g + h_g) * 2 * n_cells
    price_bead = 600  # ~600₸/м
    cost_bead = L_b * price_bead
    skeleton_cost += cost_bead
    
    print(f"\n5. ШТАПИК (Bead):")
    print(f"   Формула: (w_g + h_g) × 2 × count_cells")
    print(f"   Расчёт: ({w_g:.2f} + {h_g:.2f}) × 2 × {n_cells} = {L_b:.2f}м")
    print(f"   Стоимость: {cost_bead:,.0f}₸")
    
    result["skeleton"]["Штапик"] = {
        "quantity": L_b,
        "unit": "м",
        "price": price_bead,
        "cost": cost_bead
    }
    
    # --- УПЛОТНИТЕЛЬ ---
    # Два контура (внешний и под штапик) на каждое заполнение
    L_s = L_b * 2 * 1.05  # +5% запас
    price_seal = 300  # ~300₸/м
    cost_seal = L_s * price_seal
    skeleton_cost += cost_seal
    
    print(f"\n6. УПЛОТНИТЕЛЬ:")
    print(f"   Формула: L_b × 2 × 1.05")
    print(f"   Расчёт: {L_b:.2f} × 2 × 1.05 = {L_s:.2f}м")
    print(f"   Стоимость: {cost_seal:,.0f}₸")
    
    result["skeleton"]["Уплотнитель"] = {
        "quantity": L_s,
        "unit": "м",
        "price": price_seal,
        "cost": cost_seal
    }
    
    print(f"\n{'─'*70}")
    print(f"ИТОГО МАТЕРИАЛЫ ТАМБУРА: {skeleton_cost:,.0f}₸")
    print("="*70)
    
    result["total_cost"] = skeleton_cost
    
    return result
