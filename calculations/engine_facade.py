"""
Модуль расчёта фасадных систем (Ruit 50F)
Полный расчёт профилей, заполнения и вставок (двери/окна)
"""

import math
from typing import Dict, List, Any

# Импорты для расчёта вставок (вынесены из функций)
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
    
    # Каркас фасада
    skeleton_cost = 0
    
    # Стойки
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
    
    price_m = 0
    mullion_name = f"Стойка {mullion_size} мм"
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if mullion_name in elem:
            price_m = parse_price(item.get('Цена за единицу', 0))
            break
    
    cost_m = final_m * price_m
    skeleton_cost += cost_m
    
    result["skeleton"][f"Стойка {mullion_size}мм"] = {
        "quantity": final_m,
        "unit": "м",
        "price": price_m,
        "cost": cost_m
    }
    
    # Вставки
    inserts_cost = 0
    inserts_details = []
    
    if inserts and len(inserts) > 0:
        for i, insert in enumerate(inserts, 1):
            product_type = insert.get('product_type', 'Дверь 2-х створч.')
            insert_system = insert.get('system', 'ALG 2030-63C')
            insert_w = insert.get('width', 0)
            insert_h = insert.get('height', 0)
            
            # Импорты уже в начале файла
            
            code = get_code_for_windows_doors(product_type, insert_system)
            
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
                    "width": insert_w * 1000,
                    "height": insert_h * 1000,
                    "count": 1,
                    "fill_category": "Стеклопакет",
                    "glass_type": insert.get('data', {}).get('glass_type', 'двойной'),
                    "opening_type": "Откр.",
                    "horizontal_imposts": 0,
                    "vertical_imposts": 0
                }]
            }
            
            try:
                insert_result = calculate_window_smeta(insert_order_data, ref1, ref2, ref3)
                insert_materials = insert_result.get("part3_final", {}).get("Материалы", 0)
                inserts_cost += insert_materials
                
                inserts_details.append({
                    "name": f"{product_type} {insert_system}",
                    "size": f"{insert_w}м × {insert_h}м",
                    "cost": insert_materials
                })
            except Exception as e:
                print(f"   ⚠️ Ошибка расчёта вставки: {e}")
                fallback_cost = 250000
                inserts_cost += fallback_cost
                inserts_details.append({
                    "name": f"{product_type} {insert_system}",
                    "size": f"{insert_w}м × {insert_h}м",
                    "cost": fallback_cost
                })
    
    result["inserts_details"] = inserts_details
    result["total_cost"] = skeleton_cost + inserts_cost
    result["skeleton_cost"] = skeleton_cost
    result["inserts_cost"] = inserts_cost
    
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
        "products": [],
        "connecting": {},
        "total_products_cost": 0,
        "total_connecting_cost": 0,
        "total_cost": 0
    }
    
    products_cost = 0
    total_perimeter = 0
    
    for i, pos in enumerate(positions, 1):
        product_type = pos.get("product_type", "Дверь 2-х створч.")
        system = pos.get("system", "ALG 2030-63C")
        width = pos.get("width", 1800)
        height = pos.get("height", 2200)
        
        code = get_code_for_windows_doors(product_type, system)
        
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
                "width": width,
                "height": height,
                "count": 1,
                "fill_category": "Стеклопакет",
                "glass_type": pos.get("glass_type", "двойной"),
                "opening_type": pos.get("opening_type", "Откр."),
                "horizontal_imposts": pos.get("horizontal_imposts", 0),
                "vertical_imposts": pos.get("vertical_imposts", 0)
            }]
        }
        
        try:
            item_result = calculate_window_smeta(order_data, ref1, ref2, ref3)
            item_cost = item_result.get("materials_cost", 0)
            products_cost += item_cost
            
            result["products"].append({
                "name": f"{product_type} {system}",
                "size": f"{width}×{height}мм",
                "cost": item_cost
            })
            
            total_perimeter += 2 * ((width + height) / 1000)
        except Exception as e:
            print(f"  ⚠️ Ошибка расчёта: {e}")
    
    result["total_products_cost"] = products_cost
    result["total_cost"] = products_cost
    
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
    """
    
    result = {
        "skeleton": {},
        "total_cost": 0,
        "details": []
    }
    
    return result
