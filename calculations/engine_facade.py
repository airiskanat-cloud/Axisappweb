"""
Модуль расчёта фасадных систем (Ruit 50F)
ОБНОВЛЕНО: Поддержка трапеции, ручной выбор профилей, исправленные формулы
"""

import math
from typing import Dict, List, Any, Optional

# Импорты для расчёта вставок
try:
    from calculations.engine_windows import calculate_window_smeta
    from calculations.mapping import get_code_for_windows_doors
except ImportError:
    # Fallback если импорты не работают
    calculate_window_smeta = None
    get_code_for_windows_doors = None


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


def calculate_facade_geometry(
    W: float,
    H1: float,
    H2: Optional[float] = None,
    count: int = 1
) -> Dict[str, float]:
    """
    Расчёт геометрии фасада (трапеция или прямоугольник)
    
    Args:
        W: Ширина (м)
        H1: Высота слева (м)
        H2: Высота справа (м), если None или 0 → прямоугольник (H2 = H1)
        count: Количество фасадов
    
    Returns:
        {
            "Havg": средняя высота,
            "area": площадь,
            "perimeter": периметр,
            "Lslope": длина наклонной стороны
        }
    """
    # Если H2 не задано или 0 → прямоугольник
    if H2 is None or H2 == 0:
        H2 = H1
    
    # Средняя высота
    Havg = (H1 + H2) / 2
    
    # Площадь фасада
    area = W * Havg * count
    
    # Наклонная сторона (гипотенуза)
    Lslope = math.sqrt(W**2 + (H1 - H2)**2)
    
    # Периметр фасада
    perimeter = (W + H1 + H2 + Lslope) * count
    
    return {
        "Havg": Havg,
        "area": area,
        "perimeter": perimeter,
        "Lslope": Lslope,
        "is_trapezoid": abs(H1 - H2) > 0.01  # Трапеция если разница > 1см
    }


def round_to_multiple_up(value: float, multiple: float = 6.0) -> float:
    """Округление ВВЕРХ кратно заданному значению"""
    return math.ceil(value / multiple) * multiple


def find_profile_in_ref(
    facade_profiles_ref: List[Dict],
    element_name: str
) -> Dict[str, Any]:
    """
    Поиск профиля в справочнике по названию элемента
    
    Returns:
        {"price": float, "article": str, "found": bool}
    """
    for item in facade_profiles_ref:
        elem = item.get('Элемент', '')
        if element_name.lower() in elem.lower():
            return {
                "price": parse_price(item.get('Цена за единицу', 0)),
                "article": item.get('Артикул', ''),
                "name": elem,
                "found": True
            }
    
    return {"price": 0, "article": "", "name": element_name, "found": False}


def calculate_facade_frame(
    W: float,
    Havg: float,
    cols: int,
    rows: int,
    count: int,
    mullion_size: int,  # Ручной выбор!
    transom_size: int,  # Ручной выбор!
    brackets_per_mullion: int,  # Кронштейнов на 1 стойку
    facade_profiles_ref: List[Dict]
) -> Dict[str, Any]:
    """
    Расчёт каркаса фасада (ИСПРАВЛЕННЫЕ ФОРМУЛЫ по ТЗ)
    
    Args:
        W: Ширина фасада
        Havg: Средняя высота
        cols: Количество столбцов
        rows: Количество рядов
        count: Количество фасадов
        mullion_size: Сечение стойки (мм) - РУЧНОЙ ВЫБОР
        transom_size: Сечение ригеля (мм) - РУЧНОЙ ВЫБОР
        brackets_per_mullion: Кронштейнов на 1 стойку
        facade_profiles_ref: Справочник профилей
    
    Returns:
        {
            "mullions": {...},
            "transoms": {...},
            "press_profile": {...},
            "seals": {...},
            "brackets": {...},
            ...
        }
    """
    
    result = {}
    total_cost = 0
    
    print("\n" + "="*70)
    print("РАСЧЁТ КАРКАСА ФАСАДА (ИСПРАВЛЕННЫЕ ФОРМУЛЫ)")
    print("="*70)
    
    # ============================================================================
    # 1. СТОЙКИ (Mullions)
    # ============================================================================
    
    n_mullions = cols + 1  # Количество стоек
    Lst_raw = n_mullions * Havg * count  # БЕЗ округления
    Lst = round_to_multiple_up(Lst_raw, 6)  # Округление вверх кратно 6м
    
    mullion_info = find_profile_in_ref(facade_profiles_ref, f"Стойка {mullion_size} мм")
    
    cost_mullions = Lst * mullion_info["price"]
    total_cost += cost_mullions
    
    print(f"\n1. СТОЙКИ {mullion_size}мм:")
    print(f"   Формула: (cols + 1) × Havg × count")
    print(f"   Расчёт: {n_mullions} × {Havg:.2f}м × {count} = {Lst_raw:.2f}м")
    print(f"   Округление: ⌈{Lst_raw:.2f}/6⌉ × 6 = {Lst:.0f}м")
    print(f"   Цена: {mullion_info['price']:,}₸/м")
    print(f"   Стоимость: {cost_mullions:,.0f}₸")
    
    result["mullions"] = {
        "quantity": Lst,
        "quantity_raw": Lst_raw,
        "unit": "м",
        "price": mullion_info["price"],
        "cost": cost_mullions,
        "size": mullion_size,
        "article": mullion_info["article"]
    }
    
    # ============================================================================
    # 2. РИГЕЛИ (Transoms)
    # ============================================================================
    
    Lrig_raw = W * rows * count  # БЕЗ округления
    Lrig = round_to_multiple_up(Lrig_raw, 6)  # Округление вверх кратно 6м
    
    transom_info = find_profile_in_ref(facade_profiles_ref, f"Ригель {transom_size} мм")
    
    cost_transoms = Lrig * transom_info["price"]
    total_cost += cost_transoms
    
    print(f"\n2. РИГЕЛИ {transom_size}мм:")
    print(f"   Формула: W × rows × count")
    print(f"   Расчёт: {W:.2f}м × {rows} × {count} = {Lrig_raw:.2f}м")
    print(f"   Округление: ⌈{Lrig_raw:.2f}/6⌉ × 6 = {Lrig:.0f}м")
    print(f"   Цена: {transom_info['price']:,}₸/м")
    print(f"   Стоимость: {cost_transoms:,.0f}₸")
    
    result["transoms"] = {
        "quantity": Lrig,
        "quantity_raw": Lrig_raw,
        "unit": "м",
        "price": transom_info["price"],
        "cost": cost_transoms,
        "size": transom_size,
        "article": transom_info["article"]
    }
    
    # ============================================================================
    # 3. ПРИЖИМНОЙ ПРОФИЛЬ / КРЫШКА (ИСПРАВЛЕНО!)
    # ============================================================================
    # ✅ ПРАВИЛЬНАЯ ФОРМУЛА: Lpr = Lst + Lrig (БЕЗ коэффициентов!)
    
    Lpr = Lst + Lrig  # Просто сумма!
    
    press_info = find_profile_in_ref(facade_profiles_ref, "Прижимной профиль")
    
    cost_press = Lpr * press_info["price"]
    total_cost += cost_press
    
    print(f"\n3. ПРИЖИМНОЙ ПРОФИЛЬ (ИСПРАВЛЕНО!):")
    print(f"   Формула: Lst + Lrig")
    print(f"   Расчёт: {Lst:.0f}м + {Lrig:.0f}м = {Lpr:.0f}м")
    print(f"   Цена: {press_info['price']:,}₸/м")
    print(f"   Стоимость: {cost_press:,.0f}₸")
    
    result["press_profile"] = {
        "quantity": Lpr,
        "unit": "м",
        "price": press_info["price"],
        "cost": cost_press
    }
    
    # Крышка фасадная (аналогично)
    cover_info = find_profile_in_ref(facade_profiles_ref, "Крышка фасадная")
    cost_cover = Lpr * cover_info["price"]
    total_cost += cost_cover
    
    result["cover"] = {
        "quantity": Lpr,
        "unit": "м",
        "price": cover_info["price"],
        "cost": cost_cover
    }
    
    # ============================================================================
    # 4. УПЛОТНИТЕЛЬ (ИСПРАВЛЕНО!)
    # ============================================================================
    # ✅ ПРАВИЛЬНАЯ ФОРМУЛА: Lseal = (Lst + Lrig) × 2 × 1.05
    
    Lseal = (Lst + Lrig) * 2 * 1.05  # × 2 (двусторонний) + 5% запас
    # ВАЖНО: Уплотнитель НЕ округляется кратно 6м!
    
    seal_info = find_profile_in_ref(facade_profiles_ref, "Упл фасада")
    
    cost_seal = Lseal * seal_info["price"]
    total_cost += cost_seal
    
    print(f"\n4. УПЛОТНИТЕЛЬ (ИСПРАВЛЕНО!):")
    print(f"   Формула: (Lst + Lrig) × 2 × 1.05")
    print(f"   Расчёт: ({Lst:.0f} + {Lrig:.0f}) × 2 × 1.05 = {Lseal:.2f}м")
    print(f"   ⚠️ Не округляется кратно 6м!")
    print(f"   Цена: {seal_info['price']:,}₸/м")
    print(f"   Стоимость: {cost_seal:,.0f}₸")
    
    result["seals"] = {
        "quantity": Lseal,
        "unit": "м",
        "price": seal_info["price"],
        "cost": cost_seal
    }
    
    # ============================================================================
    # 5. КРОНШТЕЙНЫ (по новому параметру)
    # ============================================================================
    
    count_brackets = brackets_per_mullion * n_mullions * count
    
    bracket_info = find_profile_in_ref(facade_profiles_ref, "Кронштейн")
    
    cost_brackets = count_brackets * bracket_info["price"]
    total_cost += cost_brackets
    
    print(f"\n5. КРОНШТЕЙНЫ:")
    print(f"   Формула: brackets_per_mullion × n_mullions × count")
    print(f"   Расчёт: {brackets_per_mullion} × {n_mullions} × {count} = {count_brackets} шт")
    print(f"   Стоимость: {cost_brackets:,.0f}₸")
    
    result["brackets"] = {
        "quantity": count_brackets,
        "unit": "шт",
        "price": bracket_info["price"],
        "cost": cost_brackets
    }
    
    # ============================================================================
    # 6. ДОПОЛНИТЕЛЬНЫЕ ЭЛЕМЕНТЫ
    # ============================================================================
    
    # U-соединители ригеля (по 2 на каждый ригель)
    count_u = 2 * rows * (cols + 1) * count
    u_info = find_profile_in_ref(facade_profiles_ref, "U-соединитель")
    cost_u = count_u * u_info["price"]
    total_cost += cost_u
    
    result["u_connectors"] = {
        "quantity": count_u,
        "unit": "шт",
        "price": u_info["price"],
        "cost": cost_u
    }
    
    # Термомост
    L_thermo = (Lst + Lrig) * 1.05  # +5% запас
    thermo_info = find_profile_in_ref(facade_profiles_ref, "Термомост 18мм")
    cost_thermo = L_thermo * thermo_info["price"]
    total_cost += cost_thermo
    
    result["thermobridges"] = {
        "quantity": L_thermo,
        "unit": "м",
        "price": thermo_info["price"],
        "cost": cost_thermo
    }
    
    # Держатели СП (по 2 на ячейку)
    count_cells = cols * rows * count
    count_holders = 2 * count_cells
    holder_info = find_profile_in_ref(facade_profiles_ref, "Держатель")
    cost_holders = count_holders * holder_info["price"]
    total_cost += cost_holders
    
    result["holders"] = {
        "quantity": count_holders,
        "unit": "шт",
        "price": holder_info["price"],
        "cost": cost_holders
    }
    
    print(f"\n" + "="*70)
    print(f"ИТОГО КАРКАС: {total_cost:,.0f}₸")
    print("="*70)
    
    result["total_cost"] = total_cost
    result["summary"] = {
        "Lst": Lst,
        "Lrig": Lrig,
        "Lpr": Lpr,
        "Lseal": Lseal
    }
    
    return result


def calculate_facade_inserts(
    inserts: List[Dict],
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict],
    facade_profiles_ref: List[Dict]
) -> Dict[str, Any]:
    """
    Расчёт материалов вставок (окна/двери)
    
    Args:
        inserts: Список вставок
        ref1, ref2, ref3: Справочники
        facade_profiles_ref: Справочник фасадных профилей
    
    Returns:
        {
            "materials": {...},
            "total_cost": float,
            "adapter_frames": {...}  # Адаптеры рамы
        }
    """
    
    print("\n" + "="*70)
    print("РАСЧЁТ МАТЕРИАЛОВ ВСТАВОК")
    print("="*70)
    
    result = {
        "materials": {},
        "total_cost": 0,
        "adapter_frames": {
            "quantity": 0,
            "cost": 0
        }
    }
    
    if not inserts:
        print("   Вставок нет")
        return result
    
    total_adapter_perimeter = 0
    
    for idx, insert in enumerate(inserts):
        print(f"\nВставка #{idx+1}:")
        print(f"  Тип: {insert.get('type', '?')}")
        print(f"  Размер: {insert.get('width', 0)} × {insert.get('height', 0)} м")
        
        # Периметр вставки для адаптера рамы
        w = insert.get('width', 0)
        h = insert.get('height', 0)
        perimeter = 2 * (w + h)
        total_adapter_perimeter += perimeter
        
        # Расчёт материалов вставки через существующую функцию
        # НО БЕЗ стекла (стекло считается отдельно)
        if calculate_window_smeta:
            insert_result = calculate_window_smeta(
                W=w,
                H=h,
                system_code=insert.get('system_code', ''),
                glass_type=insert.get('glass_type', ''),
                # ... другие параметры
            )
            
            # Добавляем материалы вставки (без стекла)
            for material, data in insert_result.get('materials', {}).items():
                if 'стекло' not in material.lower():
                    if material not in result["materials"]:
                        result["materials"][material] = {
                            "quantity": 0,
                            "cost": 0
                        }
                    result["materials"][material]["quantity"] += data.get("quantity", 0)
                    result["materials"][material]["cost"] += data.get("cost", 0)
    
    # ============================================================================
    # АДАПТЕР РАМЫ (автоматически для всех вставок)
    # ============================================================================
    
    if total_adapter_perimeter > 0:
        adapter_info = find_profile_in_ref(facade_profiles_ref, "Адаптер рамы")
        
        cost_adapter = total_adapter_perimeter * adapter_info["price"]
        
        print(f"\nАДАПТЕР РАМЫ (автоматически):")
        print(f"  Артикул: {adapter_info['article']}")
        print(f"  Количество: {total_adapter_perimeter:.2f} м")
        print(f"  Стоимость: {cost_adapter:,.0f}₸")
        
        result["adapter_frames"] = {
            "quantity": total_adapter_perimeter,
            "unit": "м",
            "price": adapter_info["price"],
            "cost": cost_adapter,
            "article": adapter_info["article"]
        }
        
        result["total_cost"] += cost_adapter
    
    return result


def calculate_facade_materials(
    W: float,
    H1: float,
    H2: Optional[float],
    cols: int,
    rows: int,
    count: int,
    mullion_size: int,  # НОВОЕ: Ручной выбор
    transom_size: int,  # НОВОЕ: Ручной выбор
    brackets_per_mullion: int,  # НОВОЕ: Кронштейнов на стойку
    inserts: List[Dict],
    facade_profiles_ref: List[Dict],
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict]
) -> Dict[str, Any]:
    """
    ГЛАВНАЯ ФУНКЦИЯ: Полный расчёт материалов фасада
    
    ОБНОВЛЕНО ПО ТЗ:
    - Поддержка трапеции (H1, H2)
    - Ручной выбор профилей (mullion_size, transom_size)
    - Исправленные формулы (Lpr, Lseal)
    """
    
    print("\n" + "="*70)
    print("РАСЧЁТ ФАСАДНОЙ СИСТЕМЫ (Ruit 50F)")
    print("="*70)
    
    # ============================================================================
    # 1. ГЕОМЕТРИЯ (трапеция или прямоугольник)
    # ============================================================================
    
    geometry = calculate_facade_geometry(W, H1, H2, count)
    
    print(f"\nГЕОМЕТРИЯ:")
    print(f"  Ширина: {W} м")
    print(f"  Высота слева: {H1} м")
    print(f"  Высота справа: {H2 if H2 else H1} м")
    print(f"  Средняя высота: {geometry['Havg']:.2f} м")
    print(f"  Форма: {'Трапеция' if geometry['is_trapezoid'] else 'Прямоугольник'}")
    print(f"  Площадь: {geometry['area']:.2f} м²")
    print(f"  Периметр: {geometry['perimeter']:.2f} м")
    
    # ============================================================================
    # 2. КАРКАС (с исправленными формулами)
    # ============================================================================
    
    frame = calculate_facade_frame(
        W=W,
        Havg=geometry["Havg"],
        cols=cols,
        rows=rows,
        count=count,
        mullion_size=mullion_size,
        transom_size=transom_size,
        brackets_per_mullion=brackets_per_mullion,
        facade_profiles_ref=facade_profiles_ref
    )
    
    # ============================================================================
    # 3. ВСТАВКИ (окна/двери)
    # ============================================================================
    
    inserts_result = calculate_facade_inserts(
        inserts=inserts,
        ref1=ref1,
        ref2=ref2,
        ref3=ref3,
        facade_profiles_ref=facade_profiles_ref
    )
    
    # ============================================================================
    # 4. ИТОГОВЫЙ РЕЗУЛЬТАТ
    # ============================================================================
    
    result = {
        "geometry": geometry,
        "frame": frame,
        "inserts": inserts_result,
        "total_cost": frame["total_cost"] + inserts_result["total_cost"],
        "metrics": {
            "area": geometry["area"],
            "perimeter": geometry["perimeter"],
            "cost_per_sqm": 0  # Рассчитается позже
        }
    }
    
    # Стоимость за 1 м²
    if geometry["area"] > 0:
        result["metrics"]["cost_per_sqm"] = result["total_cost"] / geometry["area"]
    
    print(f"\n" + "="*70)
    print(f"ИТОГО ФАСАД: {result['total_cost']:,.0f}₸")
    print(f"Стоимость за 1 м²: {result['metrics']['cost_per_sqm']:,.0f}₸/м²")
    print("="*70)
    
    return result
