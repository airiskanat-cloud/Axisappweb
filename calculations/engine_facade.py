"""
Модуль расчёта фасадных систем (Ruit 50F)
ОБНОВЛЕНО V9: Поддержка трапеции, ручной выбор профилей, исправленные формулы.
Этап 1: все цены извлекаются из справочников через get_material_data().
Этап 2: функции возвращают quantity_raw (НЕТТО), округление — только в MaterialAggregator.
"""

import math
from typing import Dict, List, Any, Optional

# Импорт констант (Этап 3) — constants.py кладётся в calculations/
try:
    from calculations.constants import MaterialKeys, TambourArticles, ServiceKeys
except (ImportError, ModuleNotFoundError):
    try:
        from .constants import MaterialKeys, TambourArticles, ServiceKeys
    except (ImportError, ModuleNotFoundError):
        # Финальный fallback: определяем минимально необходимые константы инлайн
        class MaterialKeys:
            ARTICLE = "Артикул"
            ELEMENT = "Элемент"
            PRICE = "Цена за единицу"
            PACKAGE_SIZE = "Кратность"
        class TambourArticles:
            GUIDE = "2-00-5581"
            PIPE = "2-00-2010"
        class ServiceKeys:
            LAMBRI_NO_THERMO = "ламбри без термо"

# Импорты для расчёта вставок
try:
    from calculations.engine_windows import calculate_window_smeta
    from calculations.mapping import get_code_for_windows_doors
except ImportError:
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
        for space in ['\xa0', '\u00a0', '\u202f', '\u2009', ' ']:
            value = value.replace(space, '')
        value = value.replace(',', '.')
        return float(value)
    except:
        return 0.0


def get_material_data(article_id: str, ref_data: List[Dict], search_field: str = "Артикул") -> Dict[str, Any]:
    """
    ✅ ЦЕНТРАЛЬНЫЙ МАППЕР (Этап 1 ТЗ V.9)
    Извлекает данные материала из справочника по артикулу или ключевому слову.
    
    Единственная точка доступа к ценам в движках.
    Запрещено использовать числа как цены — только через эту функцию.
    
    Args:
        article_id: Артикул или ключевое слово для поиска
        ref_data:   Список справочника (ref1 или ref_facade)
        search_field: Поле для поиска ("Артикул" или "Элемент")
    
    Returns:
        {
            "article":  str,   — артикул найденного элемента
            "name":     str,   — название элемента
            "price":    float, — цена за единицу
            "package_size": float, — кратность (размер хлыста)
            "found":    bool   — найден ли элемент
        }
    """
    for item in ref_data:
        val = str(item.get(search_field, ""))
        # Точное совпадение или подстрока (регистронезависимо)
        if val == article_id or article_id.lower() in val.lower():
            result = {
                "article":      item.get(MaterialKeys.ARTICLE, ""),
                "name":         item.get(MaterialKeys.ELEMENT, ""),
                "price":        parse_price(item.get(MaterialKeys.PRICE, 0)),
                "package_size": parse_price(item.get(MaterialKeys.PACKAGE_SIZE, 1)),
                "found":        True
            }
            print(f"🔍 Поиск материала [{article_id}] в Справочнике: Найдено — "
                  f"{result['name']} | {result['article']} | {result['price']:,.0f}₸")
            return result

    print(f"🔍 Поиск материала [{article_id}] в Справочнике: Не найдено")
    return {
        "article":      "",
        "name":         article_id,
        "price":        0.0,
        "package_size": 1.0,
        "found":        False
    }


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
    """Округление ВВЕРХ кратно заданному значению (используется только в тестах и легаси)"""
    return math.ceil(value / multiple) * multiple


def calculate_facade_frame(
    W: float,
    Havg: float,
    cols: int,
    rows: int,
    count: int,
    mullion_size: int,
    transom_size: int,
    brackets_per_mullion: int,
    facade_profiles_ref: List[Dict]
) -> Dict[str, Any]:
    """
    Расчёт каркаса фасада.
    
    ✅ V9 Этап 2: возвращает ТОЛЬКО quantity_raw (НЕТТО).
    Стоимости НЕ считаются здесь — считаются в MaterialAggregator после округления.
    
    ✅ V9 Этап 1: все цены извлекаются через get_material_data().
    """
    
    result = {}
    
    print("\n" + "="*70)
    print("РАСЧЁТ КАРКАСА ФАСАДА (V9 — НЕТТО)")
    print("="*70)
    
    # ============================================================
    # 1. СТОЙКИ
    # ============================================================
    n_mullions = cols + 1
    Lst_raw = n_mullions * Havg * count  # НЕТТО
    
    mullion_info = get_material_data(f"Стойка {mullion_size} мм", facade_profiles_ref, search_field="Элемент")
    
    print(f"\n1. СТОЙКИ {mullion_size}мм:")
    print(f"   ({n_mullions}) × {Havg:.2f}м × {count} = {Lst_raw:.3f}м (НЕТТО)")
    
    result["mullions"] = {
        "quantity_raw": Lst_raw,
        "unit": "м",
        "price": mullion_info["price"],
        "size": mullion_size,
        "article": mullion_info["article"],
        "name": mullion_info["name"]
    }
    
    # ============================================================
    # 2. РИГЕЛИ
    # ============================================================
    Lrig_raw = W * rows * count  # НЕТТО
    
    transom_info = get_material_data(f"Ригель {transom_size} мм", facade_profiles_ref, search_field="Элемент")
    
    print(f"\n2. РИГЕЛИ {transom_size}мм:")
    print(f"   {W:.2f}м × {rows} × {count} = {Lrig_raw:.3f}м (НЕТТО)")
    
    result["transoms"] = {
        "quantity_raw": Lrig_raw,
        "unit": "м",
        "price": transom_info["price"],
        "size": transom_size,
        "article": transom_info["article"],
        "name": transom_info["name"]
    }
    
    # ============================================================
    # 3. ПРИЖИМНОЙ ПРОФИЛЬ (по raw!)
    # ============================================================
    Lpr_raw = Lst_raw + Lrig_raw  # НЕТТО
    
    press_info = get_material_data("Прижимной профиль", facade_profiles_ref, search_field="Элемент")
    
    print(f"\n3. ПРИЖИМНОЙ ПРОФИЛЬ:")
    print(f"   Lst_raw + Lrig_raw = {Lst_raw:.3f} + {Lrig_raw:.3f} = {Lpr_raw:.3f}м (НЕТТО)")
    
    result["press_profile"] = {
        "quantity_raw": Lpr_raw,
        "unit": "м",
        "price": press_info["price"],
        "article": press_info["article"],
        "name": press_info["name"]
    }
    
    # Крышка фасадная
    cover_info = get_material_data("Крышка фасадная", facade_profiles_ref, search_field="Элемент")
    
    result["cover"] = {
        "quantity_raw": Lpr_raw,
        "unit": "м",
        "price": cover_info["price"],
        "article": cover_info["article"],
        "name": cover_info["name"]
    }
    
    # ============================================================
    # 4. УПЛОТНИТЕЛЬ (по raw!)
    # ============================================================
    Lseal_raw = (Lst_raw + Lrig_raw) * 2 * 1.05  # ×2 двусторонний + 5% запас
    
    seal_info = get_material_data("Упл фасада", facade_profiles_ref, search_field="Элемент")
    
    print(f"\n4. УПЛОТНИТЕЛЬ:")
    print(f"   (Lst_raw + Lrig_raw) × 2 × 1.05 = {Lseal_raw:.3f}м (НЕТТО)")
    
    result["seals"] = {
        "quantity_raw": Lseal_raw,
        "unit": "м",
        "price": seal_info["price"],
        "article": seal_info["article"],
        "name": seal_info["name"]
    }
    
    # ============================================================
    # 5. КРОНШТЕЙНЫ (штуки)
    # ============================================================
    count_brackets = brackets_per_mullion * n_mullions * count
    
    bracket_info = get_material_data("Кронштейн", facade_profiles_ref, search_field="Элемент")
    
    print(f"\n5. КРОНШТЕЙНЫ:")
    print(f"   {brackets_per_mullion} × {n_mullions} × {count} = {count_brackets} шт")
    
    result["brackets"] = {
        "quantity_raw": count_brackets,
        "unit": "шт",
        "price": bracket_info["price"],
        "article": bracket_info["article"],
        "name": bracket_info["name"]
    }
    
    # ============================================================
    # 6. ДОПОЛНИТЕЛЬНЫЕ ЭЛЕМЕНТЫ
    # ============================================================
    
    # U-соединители (по 2 на каждый ригель)
    count_u = 2 * rows * (cols + 1) * count
    u_info = get_material_data("U-соединитель", facade_profiles_ref, search_field="Элемент")
    
    result["u_connectors"] = {
        "quantity_raw": count_u,
        "unit": "шт",
        "price": u_info["price"],
        "article": u_info["article"],
        "name": u_info["name"]
    }
    
    # Термомост (+5% запас)
    L_thermo_raw = (Lst_raw + Lrig_raw) * 1.05
    thermo_info = get_material_data("Термомост 18мм", facade_profiles_ref, search_field="Элемент")
    
    result["thermobridges"] = {
        "quantity_raw": L_thermo_raw,
        "unit": "м",
        "price": thermo_info["price"],
        "article": thermo_info["article"],
        "name": thermo_info["name"]
    }
    
    # Держатели СП (по 2 на ячейку)
    count_holders = 2 * cols * rows * count
    holder_info = get_material_data("Держатель", facade_profiles_ref, search_field="Элемент")
    
    result["holders"] = {
        "quantity_raw": count_holders,
        "unit": "шт",
        "price": holder_info["price"],
        "article": holder_info["article"],
        "name": holder_info["name"]
    }
    
    print(f"\n" + "="*70)
    print(f"КАРКАС: все значения НЕТТО. Округление и стоимости — в корзине.")
    print("="*70)
    
    return result

def calculate_facade_inserts(
    inserts: List[Dict],
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict],
    facade_profiles_ref: List[Dict]
) -> Dict[str, Any]:
    """
    Расчёт вставок (окна/двери в фасаде).
    
    ✅ V9 Этап 2: адаптер рамы возвращает quantity_raw (НЕТТО).
    ✅ V9 Этап 1: цена адаптера из справочника через get_material_data().
    ✅ V9 Этап 4: вставки идут в корзину с category='facade_inserts'.
    """
    
    print("\n" + "="*70)
    print("РАСЧЁТ ВСТАВОК (V9 — НЕТТО + МАТРЁШКА)")
    print("="*70)
    
    result = {
        "materials_cost": 0,
        "total_cost": 0,
        "adapter_frames": {
            "quantity_raw": 0,
            "cost": 0
        },
        # ✅ Этап 4: список материалов вставок для корзины
        "insert_materials_raw": []
    }
    
    if not inserts:
        print("   Вставок нет")
        return result
    
    total_adapter_perimeter = 0
    total_inserts_cost = 0
    
    for idx, insert in enumerate(inserts):
        print(f"\nВставка #{idx+1}:")
        print(f"  Тип: {insert.get('type', '?')}")
        
        w = insert.get('width', 0)
        h = insert.get('height', 0)
        
        # Адаптер рамы: 2h + w (без низа)
        adapter_perimeter = h + h + w
        total_adapter_perimeter += adapter_perimeter
        
        print(f"  Размер: {w:.2f}м × {h:.2f}м")
        print(f"  Адаптер рамы: {adapter_perimeter:.3f}м (2h + w) — НЕТТО")
        
        # ✅ Этап 4 (Матрёшка): вызываем calculate_window_smeta для каждой вставки
        if calculate_window_smeta:
            insert_order_data = {
                "positions": [{
                    "data": {
                        "width": w * 1000,
                        "height": h * 1000,
                        "product_type": insert.get("product_type", "Дверь 1 створч."),
                        "imposts": insert.get("imposts", {}),
                        "sashes": insert.get("sashes", [])
                    },
                    "count": 1
                }],
                "common": {
                    "system": insert.get("system", "ALG 2030-45C"),
                    "fill_category": insert.get("fill_category", "Стеклопакет"),
                    "glass_type": insert.get("glass_type", "Двойной"),
                    "toning": "Нет",
                    "assembly": "Нет",
                    "installation": "Нет"
                }
            }
            
            print(f"  🔧 Расчёт материалов вставки (матрёшка)...")
            insert_result = calculate_window_smeta(insert_order_data, ref1, ref2, ref3)
            
            # Себестоимость вставки (без обеспечения)
            insert_cost = insert_result.get("materials_cost", 0)
            total_inserts_cost += insert_cost
            
            # ✅ Этап 4: сохраняем part2_materials для передачи в корзину facade_inserts
            for mat in insert_result.get("part2_materials", []):
                result["insert_materials_raw"].append({
                    "article":      mat.get("Артикул", ""),
                    "name":         mat.get("Товар", mat.get("Элемент", "")),
                    "quantity_raw": mat.get("Расход факт.", mat.get("Количество_raw", mat.get("Количество", 0))),
                    "unit":         mat.get("Ед.", mat.get("Единица", "шт")),
                    "price":        mat.get("Цена", 0)
                })
            
            print(f"  ✅ Себестоимость вставки (НЕТТО): {insert_cost:,.0f}₸")
            print(f"  ✅ Материалов для корзины: {len(insert_result.get('part2_materials', []))} позиций")
        else:
            print(f"  ⚠️ calculate_window_smeta недоступна")
    
    result["materials_cost"] = total_inserts_cost
    
    # АДАПТЕР РАМЫ — через get_material_data, НЕТТО
    if total_adapter_perimeter > 0:
        adapter_info = get_material_data("Адаптер рамы", facade_profiles_ref, search_field="Элемент")
        
        print(f"\nАДАПТЕР РАМЫ:")
        print(f"  Суммарная длина: {total_adapter_perimeter:.3f}м (НЕТТО)")
        print(f"  Округление → в корзине")
        
        result["adapter_frames"] = {
            "quantity_raw": total_adapter_perimeter,  # НЕТТО
            "unit": "м",
            "price": adapter_info["price"],
            "article": adapter_info["article"],
            "name": adapter_info["name"]
        }
        
        result["total_cost"] = total_inserts_cost
    else:
        result["total_cost"] = total_inserts_cost
    
    print(f"\n✅ ИТОГО ВСТАВКИ:")
    print(f"   Себестоимость вставок: {total_inserts_cost:,.0f}₸")
    print(f"   Адаптер рамы: {total_adapter_perimeter:.3f}м (НЕТТО, стоимость в корзине)")
    
    return result

def calculate_facade_materials(
    W: float,
    H1: float,
    H2: Optional[float],
    cols: int,
    rows: int,
    count: int,
    mullion_size: int,
    transom_size: int,
    brackets_per_mullion: int,
    inserts: List[Dict],
    facade_profiles_ref: List[Dict],
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict]
) -> Dict[str, Any]:
    """
    ГЛАВНАЯ ФУНКЦИЯ: Полный расчёт фасада.
    
    ✅ V9 Этап 2: возвращает materials_raw со всеми НЕТТО значениями.
    ✅ V9 Этап 1: все цены из справочников.
    ✅ V9 Этап 4: вставки — через матрёшку, с insert_materials_raw.
    
    Мувиль (нащельник) НЕ считается здесь — считается в app.py по общему периметру проекта.
    """
    
    print("\n" + "="*70)
    print("РАСЧЁТ ФАСАДНОЙ СИСТЕМЫ (Ruit 50F) — V9")
    print("="*70)
    
    # ============================================================
    # 1. ГЕОМЕТРИЯ
    # ============================================================
    geometry = calculate_facade_geometry(W, H1, H2, count)
    
    print(f"\nГЕОМЕТРИЯ:")
    print(f"  Ширина: {W} м")
    print(f"  Высота слева: {H1} м")
    print(f"  Высота справа: {H2 if H2 else H1} м")
    print(f"  Средняя высота: {geometry['Havg']:.2f} м")
    print(f"  Форма: {'Трапеция' if geometry['is_trapezoid'] else 'Прямоугольник'}")
    print(f"  Площадь: {geometry['area']:.2f} м²")
    print(f"  Периметр: {geometry['perimeter']:.2f} м")
    
    # ============================================================
    # 2. КАРКАС (V9 — НЕТТО)
    # ============================================================
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
    
    # ============================================================
    # 3. ВСТАВКИ (матрёшка)
    # ============================================================
    inserts_result = calculate_facade_inserts(
        inserts=inserts,
        ref1=ref1,
        ref2=ref2,
        ref3=ref3,
        facade_profiles_ref=facade_profiles_ref
    )
    
    # ============================================================
    # 4. ФОРМИРОВАНИЕ materials_raw ДЛЯ КОРЗИНЫ
    # ============================================================
    # Каркас — все элементы через одинаковый паттерн
    materials_raw = []
    
    frame_elements = [
        ("mullions",       f"Стойка {mullion_size}мм"),
        ("transoms",       f"Ригель {transom_size}мм"),
        ("press_profile",  "Прижимной профиль"),
        ("cover",          "Крышка фасадная"),
        ("seals",          "Уплотнитель фасадный"),
        ("brackets",       "Кронштейны"),
        ("u_connectors",   "U-соединители"),
        ("thermobridges",  "Термомост"),
        ("holders",        "Держатели СП"),
    ]
    
    for key, default_name in frame_elements:
        if key in frame:
            elem = frame[key]
            materials_raw.append({
                "article":      elem.get("article", ""),
                "name":         elem.get("name", default_name),
                "quantity_raw": elem.get("quantity_raw", 0),
                "unit":         elem.get("unit", "м"),
                "price":        elem.get("price", 0)
            })
    
    # Адаптер рамы (из вставок)
    adapter = inserts_result.get("adapter_frames", {})
    if adapter.get("quantity_raw", 0) > 0:
        materials_raw.append({
            "article":      adapter.get("article", ""),
            "name":         adapter.get("name", "Адаптер рамы"),
            "quantity_raw": adapter.get("quantity_raw", 0),
            "unit":         adapter.get("unit", "м"),
            "price":        adapter.get("price", 0)
        })
    
    # ============================================================
    # 5. РЕЗУЛЬТАТ
    # ============================================================
    result = {
        "geometry":     geometry,
        "frame":        frame,
        "inserts":      inserts_result,
        "materials_raw": materials_raw,          # ← Каркас НЕТТО для facade_frame
        "insert_materials_raw": inserts_result.get("insert_materials_raw", []),  # ← Вставки для facade_inserts
        "total_cost":   inserts_result.get("total_cost", 0),  # Только себестоимость вставок (каркас в корзине)
        "metrics": {
            "total_area":      geometry["area"],
            "total_perimeter": geometry["perimeter"],
            "cost_per_sqm":    0,
            "W":  W,
            "H1": H1,
            "H2": H2 if H2 else 0
        }
    }
    
    print(f"\n" + "="*70)
    print(f"ФАСАД: materials_raw сформирован ({len(materials_raw)} каркас + "
          f"{len(inserts_result.get('insert_materials_raw', []))} вставки)")
    print(f"  ✅ Мувиль НЕ считается здесь — в app.py по общему периметру")
    print("="*70)
    
    return result

def calculate_tambour_materials_v2(
    positions: List[Dict],
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict]
) -> Dict[str, Any]:
    """
    Расчёт оконного тамбура V2.
    ✅ V9 Этап 1: все цены через get_material_data().
    ✅ V9 Этап 3: system через system_id с fallback на system.
    """
    
    print("\n" + "="*70)
    print("РАСЧЁТ ОКОННОГО ТАМБУРА V2 (V9)")
    print("="*70)
    
    result = {
        "products": [],
        "connecting": {},
        "total_products_cost": 0,
        "total_connecting_cost": 0,
        "total_cost": 0,
        "materials_raw": [],  # ← для корзины tambour
        "metrics": {
            "total_area": 0.0,
            "total_perimeter": 0.0
        }
    }
    
    # ===== ЧАСТЬ 1: ИЗДЕЛИЯ =====
    print("\nЧАСТЬ 1: ИЗДЕЛИЯ")
    print("="*70)
    
    products_cost = 0
    total_perimeter = 0
    
    try:
        from calculations.engine_windows import calculate_window_smeta
        from calculations.mapping import get_code_for_windows_doors
    except ImportError:
        print("⚠️ Не удалось импортировать calculate_window_smeta")
        calculate_window_smeta = None
        get_code_for_windows_doors = None
    
    for i, pos in enumerate(positions, 1):
        # ✅ Этап 3: system_id с fallback
        product_type = pos.get("product_type", "Дверь 2-х створч.")
        system = pos.get("system_id", pos.get("system", "ALG 2030-63C"))
        width = pos.get("width", 1800)
        height = pos.get("height", 2200)
        
        print(f"\nИзделие {i}: {product_type} {system}")
        print(f"  Размер: {width}мм × {height}мм")
        
        width_m = width / 1000
        height_m = height / 1000
        area = width_m * height_m
        perimeter = 2 * (width_m + height_m)
        
        result["metrics"]["total_area"] += area
        result["metrics"]["total_perimeter"] += perimeter
        
        if get_code_for_windows_doors:
            code = get_code_for_windows_doors(product_type, system)
        else:
            code = "UNKNOWN"
        
        order_data = {
            "common": {
                "order_number": f"TAMBOUR_ITEM_{i}",
                "toning": "Нет",
                "assembly": "Нет",
                "installation": "Нет"
            },
            "positions": [{
                "product_type": product_type,
                "system_id": system,
                "code": code,
                "count": 1,
                "data": {
                    "width": width,
                    "height": height,
                    "count": 1,
                    "fill_category": pos.get("fill_category", "Стеклопакет"),
                    "glass_type": pos.get("glass_type", "двойной"),
                    "imposts": pos.get("imposts", {
                        "auto_calculate": True,
                        "has_left": False,
                        "has_center": False,
                        "has_right": False,
                        "has_tor": False
                    }),
                    "sashes": pos.get("sashes", [])
                }
            }]
        }
        
        try:
            if calculate_window_smeta:
                item_result = calculate_window_smeta(order_data, ref1, ref2, ref3)
                item_cost = item_result.get("materials_cost", 0)
                products_cost += item_cost
                
                result["products"].append({
                    "name": f"{product_type} {system}",
                    "size": f"{width}×{height}мм",
                    "cost": item_cost
                })
                
                total_perimeter += perimeter
                print(f"  ✅ Стоимость: {item_cost:,.0f}₸")
        except Exception as e:
            print(f"  ⚠️ Ошибка расчёта: {e}")
    
    result["total_products_cost"] = products_cost
    
    # ===== ЧАСТЬ 2: СОЕДИНИТЕЛЬНЫЕ ЭЛЕМЕНТЫ =====
    print("\n" + "="*70)
    print("ЧАСТЬ 2: СОЕДИНИТЕЛЬНЫЕ ЭЛЕМЕНТЫ")
    print("="*70)
    
    connecting_cost = 0
    
    # --- НАПРАВЛЯЮЩИЙ ---
    L_guide = total_perimeter * 1.05  # +5% запас
    guide_info = get_material_data(TambourArticles.GUIDE, ref1)
    
    cost_guide = L_guide * guide_info["price"]
    connecting_cost += cost_guide
    
    print(f"\n1. НАПРАВЛЯЮЩИЙ ({TambourArticles.GUIDE}):")
    print(f"   {total_perimeter:.2f}м × 1.05 = {L_guide:.2f}м | {guide_info['price']:,.0f}₸/м = {cost_guide:,.0f}₸")
    
    result["connecting"]["Направляющий"] = {
        "quantity": L_guide,
        "unit": "м",
        "price": guide_info["price"],
        "cost": cost_guide
    }
    result["materials_raw"].append({
        "article": guide_info["article"],
        "name": "Направляющий",
        "quantity_raw": L_guide,
        "unit": "м",
        "price": guide_info["price"]
    })
    
    # --- ТРУБА (если > 2 изделия) ---
    if len(positions) > 2:
        max_height = max(pos.get("height", 2200) for pos in positions) / 1000
        L_pipe = max_height * 2
        pipe_info = get_material_data(TambourArticles.PIPE, ref1)
        
        cost_pipe = L_pipe * pipe_info["price"]
        connecting_cost += cost_pipe
        
        print(f"\n2. ТРУБА 90° ({TambourArticles.PIPE}):")
        print(f"   {L_pipe:.2f}м | {pipe_info['price']:,.0f}₸/м = {cost_pipe:,.0f}₸")
        
        result["connecting"]["Труба соединительная"] = {
            "quantity": L_pipe,
            "unit": "м",
            "price": pipe_info["price"],
            "cost": cost_pipe
        }
        result["materials_raw"].append({
            "article": pipe_info["article"],
            "name": "Труба соединительная",
            "quantity_raw": L_pipe,
            "unit": "м",
            "price": pipe_info["price"]
        })
    
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
    Расчёт тамбура (сцепка рам ALG).
    ✅ V9 Этап 1: все цены через get_material_data().
    """
    
    print("\n" + "="*70)
    print("РАСЧЁТ ОКОННОГО ТАМБУРА (ALG) — V9")
    print("="*70)
    
    w_cell = W / cols
    h_cell = H / rows
    
    print(f"\n  Габариты: {W}м × {H}м | Сетка: {cols}×{rows} | Ячейка: {w_cell:.2f}×{h_cell:.2f}м")
    
    result = {
        "skeleton": {},
        "total_cost": 0,
        "materials_raw": [],  # для корзины
        "details": []
    }
    
    skeleton_cost = 0
    
    # --- РАМА (Frame) ---
    L_f = (W + H) * 2 * count
    frame_info = get_material_data("Рама", ref1, search_field="Элемент")
    # Если не нашли по слову, пробуем по системе ALG
    if not frame_info["found"]:
        for item in ref1:
            elem = item.get("Элемент", "")
            system = item.get("Система", "")
            if "рама" in elem.lower() and "ALG" in system:
                frame_info = {
                    "article": item.get("Артикул", ""),
                    "name": elem,
                    "price": parse_price(item.get("Цена за единицу", 0)),
                    "found": True
                }
                print(f"🔍 Поиск материала [Рама ALG] в Справочнике: Найдено — {elem} | {frame_info['price']:,.0f}₸")
                break
    
    cost_frame = L_f * frame_info["price"]
    skeleton_cost += cost_frame
    
    print(f"\n1. РАМА: ({W:.2f}+{H:.2f})×2×{count} = {L_f:.2f}м | {frame_info['price']:,.0f}₸/м = {cost_frame:,.0f}₸")
    
    result["skeleton"]["Рама"] = {"quantity": L_f, "unit": "м", "price": frame_info["price"], "cost": cost_frame}
    result["materials_raw"].append({"article": frame_info["article"], "name": "Рама", "quantity_raw": L_f, "unit": "м", "price": frame_info["price"]})
    
    # --- ТРУБА 90° ---
    L_pipe = H * 2 * count
    pipe_info = get_material_data(TambourArticles.PIPE, ref1)
    
    cost_pipe = L_pipe * pipe_info["price"]
    skeleton_cost += cost_pipe
    
    print(f"\n2. ТРУБА 90°: {H:.2f}×2×{count} = {L_pipe:.2f}м | {pipe_info['price']:,.0f}₸/м = {cost_pipe:,.0f}₸")
    
    result["skeleton"]["Труба соединительная"] = {"quantity": L_pipe, "unit": "м", "price": pipe_info["price"], "cost": cost_pipe}
    result["materials_raw"].append({"article": pipe_info["article"], "name": "Труба соединительная", "quantity_raw": L_pipe, "unit": "м", "price": pipe_info["price"]})
    
    # --- АДАПТЕР ТРУБЫ ---
    L_ada = H * 4 * count
    ada_info = get_material_data("Адаптер трубы", ref1, search_field="Элемент")
    if not ada_info["found"]:
        # Fallback: ищем по слову "адаптер"
        for item in ref1:
            if "адаптер" in item.get("Элемент", "").lower():
                ada_info = {
                    "article": item.get("Артикул", ""),
                    "name": item.get("Элемент", ""),
                    "price": parse_price(item.get("Цена за единицу", 0)),
                    "found": True
                }
                print(f"🔍 Поиск материала [Адаптер трубы] в Справочнике: Найдено — {ada_info['name']} | {ada_info['price']:,.0f}₸")
                break
    
    cost_ada = L_ada * ada_info["price"]
    skeleton_cost += cost_ada
    
    print(f"\n3. АДАПТЕР: {H:.2f}×4×{count} = {L_ada:.2f}м | {ada_info['price']:,.0f}₸/м = {cost_ada:,.0f}₸")
    
    result["skeleton"]["Адаптер трубы"] = {"quantity": L_ada, "unit": "м", "price": ada_info["price"], "cost": cost_ada}
    result["materials_raw"].append({"article": ada_info["article"], "name": "Адаптер трубы", "quantity_raw": L_ada, "unit": "м", "price": ada_info["price"]})
    
    # --- НАПРАВЛЯЮЩИЙ ---
    L_guide = (W + H) * count * 1.05
    guide_info = get_material_data(TambourArticles.GUIDE, ref1)
    
    cost_guide = L_guide * guide_info["price"]
    skeleton_cost += cost_guide
    
    print(f"\n4. НАПРАВЛЯЮЩИЙ: ({W:.2f}+{H:.2f})×{count}×1.05 = {L_guide:.2f}м | {guide_info['price']:,.0f}₸/м = {cost_guide:,.0f}₸")
    
    result["skeleton"]["Направляющий"] = {"quantity": L_guide, "unit": "м", "price": guide_info["price"], "cost": cost_guide}
    result["materials_raw"].append({"article": guide_info["article"], "name": "Направляющий", "quantity_raw": L_guide, "unit": "м", "price": guide_info["price"]})
    
    # --- ШТАПИК ---
    n_cells = cols * rows * count
    w_g = w_cell - 0.1
    h_g = h_cell - 0.1
    L_b = (w_g + h_g) * 2 * n_cells
    
    bead_info = get_material_data("Штапик", ref1, search_field="Элемент")
    if not bead_info["found"]:
        for item in ref1:
            elem = item.get("Элемент", "")
            system = item.get("Система", "")
            if "штапик" in elem.lower() and "ALG" in system:
                bead_info = {
                    "article": item.get("Артикул", ""),
                    "name": elem,
                    "price": parse_price(item.get("Цена за единицу", 0)),
                    "found": True
                }
                print(f"🔍 Поиск материала [Штапик ALG] в Справочнике: Найдено — {elem} | {bead_info['price']:,.0f}₸")
                break
    
    cost_bead = L_b * bead_info["price"]
    skeleton_cost += cost_bead
    
    print(f"\n5. ШТАПИК: ({w_g:.2f}+{h_g:.2f})×2×{n_cells} = {L_b:.2f}м | {bead_info['price']:,.0f}₸/м = {cost_bead:,.0f}₸")
    
    result["skeleton"]["Штапик"] = {"quantity": L_b, "unit": "м", "price": bead_info["price"], "cost": cost_bead}
    result["materials_raw"].append({"article": bead_info["article"], "name": "Штапик", "quantity_raw": L_b, "unit": "м", "price": bead_info["price"]})
    
    # --- УПЛОТНИТЕЛЬ ---
    L_s = L_b * 2 * 1.05
    
    seal_info = get_material_data("Уплотнитель", ref1, search_field="Элемент")
    if not seal_info["found"]:
        for item in ref1:
            elem = item.get("Элемент", "")
            system = item.get("Система", "")
            if "уплотн" in elem.lower() and "ALG" in system:
                seal_info = {
                    "article": item.get("Артикул", ""),
                    "name": elem,
                    "price": parse_price(item.get("Цена за единицу", 0)),
                    "found": True
                }
                print(f"🔍 Поиск материала [Уплотнитель ALG] в Справочнике: Найдено — {elem} | {seal_info['price']:,.0f}₸")
                break
    
    cost_seal = L_s * seal_info["price"]
    skeleton_cost += cost_seal
    
    print(f"\n6. УПЛОТНИТЕЛЬ: {L_b:.2f}×2×1.05 = {L_s:.2f}м | {seal_info['price']:,.0f}₸/м = {cost_seal:,.0f}₸")
    
    result["skeleton"]["Уплотнитель"] = {"quantity": L_s, "unit": "м", "price": seal_info["price"], "cost": cost_seal}
    result["materials_raw"].append({"article": seal_info["article"], "name": "Уплотнитель", "quantity_raw": L_s, "unit": "м", "price": seal_info["price"]})
    
    # --- ЛАМБЕРИ ---
    S_cell = w_cell * h_cell
    S_total = S_cell * n_cells
    L_lam = (S_total / 0.1) * 1.05
    
    # ✅ Этап 1: цена ламбри из ref2 через поиск по подстроке
    price_lambri = 0.0
    lambri_key_found = None
    for key in ref2.keys():
        if "ламбри без термо" in key.lower():
            price_lambri = float(ref2[key])
            lambri_key_found = key
            break
    print(f"🔍 Поиск материала [ламбри без термо] в Справочнике-2: {'Найдено — ' + str(price_lambri) + '₸' if lambri_key_found else 'Не найдено'}")
    
    cost_lambri = L_lam * price_lambri
    skeleton_cost += cost_lambri
    
    print(f"\n7. ЛАМБЕРИ: ({S_total:.2f}/0.1)×1.05 = {L_lam:.2f}м | {price_lambri:,.0f}₸/м = {cost_lambri:,.0f}₸")
    
    result["skeleton"]["Ламбери"] = {"quantity": L_lam, "unit": "м", "price": price_lambri, "cost": cost_lambri}
    result["materials_raw"].append({"article": lambri_key_found or "ламбри", "name": "Ламбери", "quantity_raw": L_lam, "unit": "м", "price": price_lambri})
    
    print(f"\n{'─'*70}")
    print(f"ИТОГО МАТЕРИАЛЫ ТАМБУРА: {skeleton_cost:,.0f}₸")
    print("="*70)
    
    result["total_cost"] = skeleton_cost
    
    return result

