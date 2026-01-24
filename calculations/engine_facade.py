"""
Модуль расчёта фасадных систем (Ruit 50F)
Полный расчёт профилей, заполнения и вставок (двери/окна)
"""

import math
from typing import Dict, List, Any

# Импорты для расчёта вставок
from calculations.engine_windows import calculate_window_smeta
from calculations.mapping import get_code_for_windows_doors


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
    blind_data: Dict,  # ✅ ДОБАВЛЕНО: данные заполнения для blind ячеек
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
        "details": [],
        # ✅ ДОБАВЛЕНО: Метрики (площадь и периметр)
        "metrics": {
            "total_area": W * H * count,  # Общая площадь фасада
            "total_perimeter": 2 * (W + H) * count  # Общий периметр
        }
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
    
    price_u = 151  # Запасное
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if 'U-соединитель' in elem or 'u-соединитель' in elem.lower():
            price_u = parse_price(item.get('Цена за единицу', 0))
            if price_u > 0:
                break
    
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
    
    # --- УПЛОТНИТЕЛЬ (3/5мм) ---
    # Расход по 1м, запас 5%
    L_seal = (L_m + L_r) * 2 * count * 1.05
    price_seal = 0
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if 'Уплотнитель' in elem or 'уплотнитель' in elem.lower():
            price_seal = parse_price(item.get('Цена за единицу', 0))
            break
    
    if price_seal == 0:
        price_seal = 300  # Запасное значение
    
    cost_seal = L_seal * price_seal
    skeleton_cost += cost_seal
    
    print(f"\n8. УПЛОТНИТЕЛЬ:")
    print(f"   Формула: (L_m + L_r) × 2 × count × 1.05")
    print(f"   Расчёт: ({L_m:.2f} + {L_r:.2f}) × 2 × {count} × 1.05 = {L_seal:.2f}м")
    print(f"   Стоимость: {cost_seal:,.0f}₸")
    
    result["skeleton"]["Уплотнитель"] = {
        "quantity": L_seal,
        "unit": "м",
        "price": price_seal,
        "cost": cost_seal
    }
    
    print(f"\n{'─'*70}")
    print(f"ИТОГО КАРКАС: {skeleton_cost:,.0f}₸")
    
    # ============================================================================
    # ЧАСТЬ 1.5: ЗАПОЛНЕНИЕ ЯЧЕЕК (стекло/ламбри для blind ячеек)
    # ============================================================================
    
    # ИСПРАВЛЕНО: Добавлен расчёт заполнения для ячеек без вставок
    blind_cost = 0
    blind_cells = count_cells - len(inserts)  # Количество ячеек без вставок
    
    if blind_cells > 0:
        print("\n" + "="*70)
        print(f"ЧАСТЬ 1.5: ЗАПОЛНЕНИЕ ЯЧЕЕК ({blind_cells} шт)")
        print("="*70)
        
        # Площадь одной ячейки
        cell_area = w_cell * h_cell
        total_blind_area = cell_area * blind_cells
        
        # ИСПРАВЛЕНО: Используем blind_data для определения типа заполнения
        panel_type = blind_data.get("panel_type", "glass")
        
        if panel_type == "glass":
            # Стеклопакет
            glass_type = blind_data.get("glass_type", "двойной")
            price_glass = ref2.get(glass_type, 9500)
            blind_cost = total_blind_area * price_glass
            
            print(f"\n💎 Стеклопакет:")
            print(f"   Тип: {glass_type}")
            print(f"   Ячеек: {blind_cells}")
            print(f"   Площадь ячейки: {cell_area:.3f}м²")
            print(f"   Общая площадь: {total_blind_area:.3f}м²")
            print(f"   Цена: {price_glass:,.0f}₸/м²")
            print(f"   Стоимость: {blind_cost:,.0f}₸")
            
            result["skeleton"]["Заполнение ячеек (стекло)"] = {
                "quantity": total_blind_area,
                "unit": "м²",
                "price": price_glass,
                "cost": blind_cost
            }
        else:
            # Ламбри
            lambri_type = panel_type  # "Ламбри без термо" или "Ламбри с термо"
            price_lambri = ref2.get(lambri_type.lower(), 2248)
            
            # Округляем до хлыстов по 6м
            qty_hlysti = math.ceil(total_blind_area / 6) if total_blind_area > 0 else 0
            total_meters = qty_hlysti * 6
            blind_cost = total_meters * price_lambri
            
            print(f"\n🪵 Ламбри:")
            print(f"   Тип: {lambri_type}")
            print(f"   Ячеек: {blind_cells}")
            print(f"   Площадь: {total_blind_area:.3f}м²")
            print(f"   Хлыстов: {qty_hlysti} × 6м = {total_meters}м")
            print(f"   Цена: {price_lambri:,.0f}₸/м")
            print(f"   Стоимость: {blind_cost:,.0f}₸")
            
            result["skeleton"]["Заполнение ячеек (ламбри)"] = {
                "quantity": total_meters,
                "unit": "м",
                "price": price_lambri,
                "cost": blind_cost
            }
        
        skeleton_cost += blind_cost
        print(f"\nОБНОВЛЁННАЯ СТОИМОСТЬ КАРКАСА (с заполнением): {skeleton_cost:,.0f}₸")
    
    # ============================================================================
    # ЧАСТЬ 2: ВСТАВКИ (Двери/Окна ALG)
    # ============================================================================
    
    print("\n" + "="*70)
    print("ЧАСТЬ 2: ВСТАВКИ (ДВЕРИ/ОКНА)")
    print("="*70)
    
    inserts_cost = 0
    inserts_details = []
    
    if not inserts or len(inserts) == 0:
        print("\nВставок нет")
    else:
        for i, insert in enumerate(inserts, 1):
            insert_type = insert.get('type', 'Unknown')
            insert_system = insert.get('system', 'ALG 2030-63C')
            insert_w = insert.get('width', 0)
            insert_h = insert.get('height', 0)
            product_type = insert.get('product_type', 'Дверь 2-х створч.')  # ДОБАВЛЕНО
            
            print(f"\nВставка {i}: {product_type} {insert_system} ({insert_w}м × {insert_h}м)")
            
            # Формируем данные для расчёта через engine_windows
            # Импорты уже в начале файла
            
            # Генерируем CODE
            code = get_code_for_windows_doors(product_type, insert_system)
            
            # Данные вставки
            # ИСПРАВЛЕНО: Передаём ВСЕ данные из формы window_door_ui (аналогично разделу Окна/Двери)
            insert_order_data = {
                "common": {
                    "order_number": f"INSERT_{i}",
                    "toning": insert.get('data', {}).get('toning', 'Нет'),
                    "assembly": insert.get('data', {}).get('assembly', 'Нет'),
                    "installation": insert.get('data', {}).get('installation', 'Нет')
                },
                "positions": [{
                    "product_type": product_type,
                    "system": insert_system,
                    "code": code,
                    "count": 1,
                    
                    # ✅ ИСПРАВЛЕНО: Передаём ВСЕ данные из формы!
                    "data": {
                        "width": insert_w * 1000,   # в мм
                        "height": insert_h * 1000,  # в мм
                        "count": 1,
                        
                        # ✅ УНИФИКАЦИЯ V8: Флаг для embedded режима
                        "embedded": True,  # Вставка в фасад (унифицированный расчёт)
                        
                        # Заполнение
                        "fill_category": insert.get('data', {}).get('fill_category', 'Стеклопакет'),
                        "glass_type": insert.get('data', {}).get('glass_type', 'двойной'),
                        
                        # Импосты (ВСЕ данные из формы)
                        "imposts": insert.get('data', {}).get('imposts', {
                            "auto_calculate": True,
                            "has_left": False,
                            "has_center": False,
                            "has_right": False,
                            "has_tor": False
                        }),
                        
                        # Створки (ВСЕ данные из формы)
                        "sashes": insert.get('data', {}).get('sashes', [])
                    }
                }]
            }
            
            # Вызываем расчёт
            try:
                insert_result = calculate_window_smeta(insert_order_data, ref1, ref2, ref3)
                
                # 🔧 DEBUG: Выводим ВСЕ материалы вставки
                print(f"\n🔧 DEBUG МАТЕРИАЛЫ ВСТАВКИ:")
                print(f"   part2_materials count: {len(insert_result.get('part2_materials', []))}")
                
                for mat in insert_result.get('part2_materials', []):
                    print(f"   - {mat.get('Товар', 'N/A')}: {mat.get('К отгрузке', 0)} {mat.get('Ед.', 'шт')} × {mat.get('Цена', 0)}₸ = {mat.get('Сумма', 0)}₸")
                
                print(f"\n   part3_final:")
                for key, value in insert_result.get('part3_final', {}).items():
                    print(f"   - {key}: {value:,.0f}₸")
                
                # ИСПРАВЛЕНО: Берём ТОЛЬКО материалы (профили + фурнитура), БЕЗ стекла
                insert_materials = insert_result.get("part3_final", {}).get("Материалы", 0)
                inserts_cost += insert_materials
                
                inserts_details.append({
                    "name": f"{product_type} {insert_system}",
                    "size": f"{insert_w}м × {insert_h}м",
                    "cost": insert_materials
                })
                
                print(f"   Материалы (профили+фурнитура): {insert_materials:,.0f}₸")
            except Exception as e:
                print(f"   ⚠️ Ошибка расчёта вставки: {e}")
                # Запасное значение
                fallback_cost = 250000
                inserts_cost += fallback_cost
                inserts_details.append({
                    "name": f"{product_type} {insert_system}",
                    "size": f"{insert_w}м × {insert_h}м",
                    "cost": fallback_cost
                })
    
    print(f"\n{'─'*70}")
    print(f"ИТОГО ВСТАВКИ: {inserts_cost:,.0f}₸")
    
    result["inserts_details"] = inserts_details
    
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


def calculate_tambour_materials_v2(
    positions: List[Dict],
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict]
) -> Dict[str, Any]:
    """
    Расчёт материалов для оконного тамбура V2
    
    Тамбур = готовые двери/окна + соединительные элементы (направляющий, трубы)
    """
    
            # Импорты уже в начале файла
    
    print("\n" + "="*70)
    print("РАСЧЁТ ОКОННОГО ТАМБУРА V2 (ИЗДЕЛИЯ + НАПРАВЛЯЮЩИЙ)")
    print("="*70)
    
    result = {
        "products": [],  # Изделия (двери/окна)
        "connecting": {},  # Соединительные элементы
        "total_products_cost": 0,
        "total_connecting_cost": 0,
        "total_cost": 0,
        
        # ✅ ДОБАВЛЕНО: Метрики (как в engine_windows)
        "metrics": {
            "total_area": 0.0,
            "total_perimeter": 0.0
        }
    }
    
    # ===== ЧАСТЬ 1: ИЗДЕЛИЯ (ДВЕРИ/ОКНА) =====
    print("\nЧАСТЬ 1: ИЗДЕЛИЯ")
    print("="*70)
    
    products_cost = 0
    total_perimeter = 0
    
    for i, pos in enumerate(positions, 1):
        product_type = pos.get("product_type", "Дверь 2-х створч.")
        system = pos.get("system", "ALG 2030-63C")
        width = pos.get("width", 1800)
        height = pos.get("height", 2200)
        glass_type = pos.get("glass_type", "двойной")
        opening_type = pos.get("opening_type", "Откр.")
        h_imposts = pos.get("horizontal_imposts", 0)
        v_imposts = pos.get("vertical_imposts", 0)
        
        print(f"\nИзделие {i}: {product_type} {system}")
        print(f"  Размер: {width}мм × {height}мм")
        print(f"  Стекло: {glass_type}")
        print(f"  Открывание: {opening_type}")
        print(f"  Импосты: {h_imposts}H × {v_imposts}V")
        
        # ✅ ДОБАВЛЕНО: Считаем метрики
        width_m = width / 1000
        height_m = height / 1000
        area = width_m * height_m
        perimeter = 2 * (width_m + height_m)
        
        result["metrics"]["total_area"] += area
        result["metrics"]["total_perimeter"] += perimeter
        
        # Генерируем CODE
        code = get_code_for_windows_doors(product_type, system)
        
        # Формируем данные для расчёта
        # ✅ ИСПРАВЛЕНО: Полные данные в "data" (как в окнах/дверях)
        order_data = {
            "common": {
                "order_number": f"TAMBOUR_ITEM_{i}",
                "toning": "Нет",
                "assembly": "Нет",
                "installation": "Нет"
            },
            "positions": [{
                "product_type": product_type,
                "system": system,
                "code": code,
                "count": 1,
                
                # ✅ ВСЕ ДАННЫЕ В "data":
                "data": {
                    "width": width,
                    "height": height,
                    "count": 1,
                    
                    # Заполнение
                    "fill_category": pos.get("fill_category", "Стеклопакет"),
                    "glass_type": pos.get("glass_type", glass_type),
                    
                    # Импосты (правильный формат)
                    "imposts": pos.get("imposts", {
                        "auto_calculate": True,
                        "has_left": False,
                        "has_center": False,
                        "has_right": False,
                        "has_tor": False
                    }),
                    
                    # Створки
                    "sashes": pos.get("sashes", [])
                }
            }]
        }
        
        # Вызываем расчёт
        try:
            item_result = calculate_window_smeta(order_data, ref1, ref2, ref3)
            item_cost = item_result.get("materials_cost", 0)
            products_cost += item_cost
            
            # ✅ ДОБАВЛЕНО: Берём метрики из результата calculate_window_smeta
            item_metrics = item_result.get("metrics", {})
            item_area = item_metrics.get("total_area", 0)
            item_perimeter = item_metrics.get("total_perimeter", 0)
            
            # 🔧 DEBUG:
            print(f"\n🔧 DEBUG МЕТРИКИ ИЗДЕЛИЯ {i}:")
            print(f"   item_result keys: {list(item_result.keys())}")
            print(f"   item_metrics: {item_metrics}")
            print(f"   item_area: {item_area}")
            print(f"   item_perimeter: {item_perimeter}")
            
            # ✅ Суммируем метрики в общий result
            result["metrics"]["total_area"] += item_area
            result["metrics"]["total_perimeter"] += item_perimeter
            
            # 🔧 DEBUG:
            print(f"   result['metrics'] ПОСЛЕ: {result['metrics']}")
            
            result["products"].append({
                "name": f"{product_type} {system}",
                "size": f"{width}×{height}мм",
                "cost": item_cost
            })
            
            # Считаем периметр для направляющего
            total_perimeter += 2 * ((width + height) / 1000)  # в метры
            
            print(f"  ✅ Стоимость: {item_cost:,.0f}₸")
        except Exception as e:
            print(f"  ⚠️ Ошибка расчёта: {e}")
    
    result["total_products_cost"] = products_cost
    
    # ===== ЧАСТЬ 2: СОЕДИНИТЕЛЬНЫЕ ЭЛЕМЕНТЫ =====
    print("\n" + "="*70)
    print("ЧАСТЬ 2: СОЕДИНИТЕЛЬНЫЕ ЭЛЕМЕНТЫ")
    print("="*70)
    
    connecting_cost = 0
    
    # --- НАПРАВЛЯЮЩИЙ (2-00-5581) ---
    # Обязательно добавляется для соединения изделий
    L_guide = total_perimeter * 1.05  # +5% запас
    
    price_guide = 1200  # Запасное
    for item in ref1:
        if '2-00-5581' in item.get('Артикул', ''):
            price_guide = item.get('Цена за единицу', 1200)
            break
    
    cost_guide = L_guide * price_guide
    connecting_cost += cost_guide
    
    print(f"\n1. НАПРАВЛЯЮЩИЙ (2-00-5581):")
    print(f"   Формула: Σ(периметры изделий) × 1.05")
    print(f"   Расчёт: {total_perimeter:.2f}м × 1.05 = {L_guide:.2f}м")
    print(f"   Цена: {price_guide:,}₸/м")
    print(f"   Стоимость: {cost_guide:,.0f}₸")
    
    result["connecting"]["Направляющий"] = {
        "quantity": L_guide,
        "unit": "м",
        "price": price_guide,
        "cost": cost_guide
    }
    
    # --- СОЕДИНИТЕЛЬНАЯ ТРУБА (опционально) ---
    # Если изделий > 2, добавляем трубы для угловых соединений
    if len(positions) > 2:
        # Берём максимальную высоту
        max_height = max(pos.get("height", 2200) for pos in positions) / 1000
        L_pipe = max_height * 2  # Две вертикальные стойки
        
        price_pipe = 2500  # Запасное
        for item in ref1:
            if '2-00-2010' in item.get('Артикул', ''):
                price_pipe = item.get('Цена за единицу', 2500)
                break
        
        cost_pipe = L_pipe * price_pipe
        connecting_cost += cost_pipe
        
        print(f"\n2. СОЕДИНИТЕЛЬНАЯ ТРУБА 90° (2-00-2010):")
        print(f"   Длина: {L_pipe:.2f}м")
        print(f"   Стоимость: {cost_pipe:,.0f}₸")
        
        result["connecting"]["Труба соединительная"] = {
            "quantity": L_pipe,
            "unit": "м",
            "price": price_pipe,
            "cost": cost_pipe
        }
    
    result["total_connecting_cost"] = connecting_cost
    result["total_cost"] = products_cost + connecting_cost
    
    print(f"\n{'='*70}")
    print(f"ИТОГО ИЗДЕЛИЯ: {products_cost:,.0f}₸")
    print(f"ИТОГО СОЕДИНЕНИЯ: {connecting_cost:,.0f}₸")
    print(f"ВСЕГО: {result['total_cost']:,.0f}₸")
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
    
    # Ищем цену рамы в ref1 (ALG профиль)
    price_frame = 3500  # Запасное значение
    for item in ref1:
        elem = item.get('Элемент', '')
        system = item.get('Система', '')
        # Ищем раму ALG
        if ('рама' in elem.lower() or 'Рама' in elem) and 'ALG' in system:
            price_frame = item.get('Цена за единицу', 3500)
            break
    
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
    
    price_pipe = 2500  # Запасное
    for item in ref1:
        if '2-00-2010' in item.get('Артикул', ''):
            price_pipe = item.get('Цена за единицу', 2500)
            break
    
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
    
    price_adapter = 800  # Запасное
    for item in ref1:
        if 'адаптер' in item.get('Элемент', '').lower():
            price_adapter = item.get('Цена за единицу', 800)
            break
    
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
    
    price_guide = 1200  # Запасное
    for item in ref1:
        if '2-00-5581' in item.get('Артикул', ''):
            price_guide = item.get('Цена за единицу', 1200)
            break
    
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
    
    price_bead = 600  # Запасное
    for item in ref1:
        elem = item.get('Элемент', '')
        system = item.get('Система', '')
        if 'штапик' in elem.lower() and 'ALG' in system:
            price_bead = item.get('Цена за единицу', 600)
            break
    
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
    
    price_seal = 300  # Запасное
    for item in ref1:
        elem = item.get('Элемент', '')
        system = item.get('Система', '')
        if 'уплотн' in elem.lower() and 'ALG' in system:
            price_seal = item.get('Цена за единицу', 300)
            break
    
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
    
    # --- ЛАМБЕРИ (если требуется) ---
    # Общая площадь ячейки делится на ширину панели (0.1м) + 5% запаса
    S_cell = w_cell * h_cell
    S_total = S_cell * n_cells
    L_lam = (S_total / 0.1) * 1.05
    
    # Цена ламбери из ref2
    price_lambri = ref2.get("ламбри без термо", 2248)  # По умолчанию без термо
    cost_lambri = L_lam * price_lambri
    skeleton_cost += cost_lambri
    
    print(f"\n7. ЛАМБЕРИ:")
    print(f"   Формула: (S_total / 0.1) × 1.05")
    print(f"   Расчёт: ({S_total:.2f} / 0.1) × 1.05 = {L_lam:.2f}м")
    print(f"   Цена: {price_lambri:,}₸/м")
    print(f"   Стоимость: {cost_lambri:,.0f}₸")
    
    result["skeleton"]["Ламбери"] = {
        "quantity": L_lam,
        "unit": "м",
        "price": price_lambri,
        "cost": cost_lambri
    }
    
    print(f"\n{'─'*70}")
    print(f"ИТОГО МАТЕРИАЛЫ ТАМБУРА: {skeleton_cost:,.0f}₸")
    print("="*70)
    
    result["total_cost"] = skeleton_cost
    
    return result
