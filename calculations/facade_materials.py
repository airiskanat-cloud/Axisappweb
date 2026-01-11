"""
Модуль расчета материалов для фасадов на основе справочника

Этот модуль выполняет:
1. Подбор материалов из справочника по типу профиля
2. Вычисление формул расхода материалов
3. Округление до кратности упаковки
4. Расчет стоимости
5. Формирование спецификации
"""

import math
import logging
from typing import Dict, List, Tuple

logger = logging.getLogger(__name__)


# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========

def safe_eval(formula: str, context: dict) -> float:
    """
    Безопасное вычисление формул Python из справочника
    
    Args:
        formula: Формула в виде строки (например: "w_cell * count_r50")
        context: Словарь переменных для подстановки
    
    Returns:
        Результат вычисления формулы
    """
    try:
        # Очистка формулы
        f = str(formula).replace(",", ".").replace(" ", "")
        
        # Вычисление с доступом только к math
        result = float(eval(f, {"__builtins__": None, "math": math}, context))
        
        return result
    except Exception as e:
        logger.warning(f"Ошибка в формуле '{formula}': {e}")
        return 0.0


def round_to_package(quantity: float, package_multiplicity: float) -> float:
    """
    Округление количества до кратности упаковки (в большую сторону)
    
    Args:
        quantity: Фактическое количество
        package_multiplicity: Кратность упаковки (например, 6 для профилей по 6м)
    
    Returns:
        Округленное количество
    """
    if package_multiplicity <= 0:
        return quantity
    
    return math.ceil(quantity / package_multiplicity) * package_multiplicity


def calculate_material_cost(quantity_fact: float,
                           quantity_rounded: float,
                           price_per_unit: float) -> Dict:
    """
    Расчет стоимости материала
    
    Args:
        quantity_fact: Фактическое количество
        quantity_rounded: Округленное до упаковки
        price_per_unit: Цена за единицу
    
    Returns:
        Dict со стоимостями
    """
    
    cost_fact = quantity_fact * price_per_unit
    cost_rounded = quantity_rounded * price_per_unit
    waste = quantity_rounded - quantity_fact
    waste_percent = (waste / quantity_fact * 100) if quantity_fact > 0 else 0
    
    return {
        'quantity_fact': quantity_fact,
        'quantity_rounded': quantity_rounded,
        'waste': waste,
        'waste_percent': waste_percent,
        'cost_fact': cost_fact,
        'cost_rounded': cost_rounded,
        'price_per_unit': price_per_unit
    }


# ========== ФИЛЬТРАЦИЯ СПРАВОЧНИКА ==========

def filter_facade_reference(facade_ref: List[Dict], 
                            system: str = "Ruit 50F",
                            product_type: str = "Витраж") -> List[Dict]:
    """
    Фильтрация справочника по системе и типу изделия
    
    Args:
        facade_ref: Справочник фасадов (из Google Sheets)
        system: Система профиля
        product_type: Тип изделия
    
    Returns:
        Отфильтрованный список материалов
    """
    
    filtered = []
    
    for item in facade_ref:
        sys = item.get('Система профиля', '')
        prod = item.get('Тип изделия', '')
        
        if sys == system and prod == product_type:
            filtered.append(item)
    
    logger.info(f"Отфильтровано {len(filtered)} материалов для {system} / {product_type}")
    
    return filtered


def get_materials_by_element(facade_ref: List[Dict], 
                             element_type: str) -> List[Dict]:
    """
    Получение материалов по типу элемента
    
    Args:
        facade_ref: Справочник фасадов
        element_type: Тип элемента (например: "Ригель 50 мм", "Стойка 90 мм")
    
    Returns:
        Список материалов
    """
    
    materials = []
    
    for item in facade_ref:
        element = item.get('Элемент', '')
        
        if element == element_type:
            materials.append(item)
    
    return materials


# ========== РАСЧЕТ МАТЕРИАЛОВ ==========

def calculate_facade_materials(geometry: Dict, 
                              facade_ref: List[Dict],
                              system: str = "Ruit 50F") -> Dict:
    """
    Основная функция расчета всех материалов фасада
    
    Args:
        geometry: Результат calculate_vitrazh_geometry()
        facade_ref: Справочник фасадов из Google Sheets
        system: Система профиля
    
    Returns:
        Dict с полной спецификацией материалов
    """
    
    # Фильтруем справочник
    filtered_ref = filter_facade_reference(facade_ref, system, "Витраж")
    
    if not filtered_ref:
        logger.error(f"Справочник пуст для системы {system}")
        return {'error': 'Справочник пуст', 'materials': []}
    
    # Переменные для формул
    formula_vars = geometry['formula_vars']
    profile_types = geometry['profile_types']
    
    # Результаты расчета
    materials_list = []
    total_cost = 0
    
    # === РАСЧЕТ КАЖДОГО МАТЕРИАЛА ИЗ СПРАВОЧНИКА ===
    
    for item in filtered_ref:
        try:
            # Извлекаем данные из справочника
            element = item.get('Элемент', '')
            product_name = item.get('Товар', '')
            article = item.get('Артикул', '')
            formula_str = item.get('Формула_Python', '')
            unit = item.get('Ед. измерения', '')
            price = float(item.get('Цена за единицу', 0))
            package_multiplicity = float(item.get('Кратность к упаковке', 1))
            package_unit = item.get('Ед. упаковки', '')
            
            # Пропускаем если нет формулы
            if not formula_str or formula_str == '':
                logger.debug(f"Пропущен материал без формулы: {product_name}")
                continue
            
            # Вычисляем количество по формуле
            quantity_fact = safe_eval(formula_str, formula_vars)
            
            # Пропускаем если количество = 0
            if quantity_fact <= 0:
                continue
            
            # Округляем до кратности упаковки
            quantity_rounded = round_to_package(quantity_fact, package_multiplicity)
            
            # Расчет стоимости
            cost_calc = calculate_material_cost(
                quantity_fact=quantity_fact,
                quantity_rounded=quantity_rounded,
                price_per_unit=price
            )
            
            # Добавляем в список
            material_item = {
                'element': element,
                'product_name': product_name,
                'article': article,
                'formula': formula_str,
                'quantity_fact': quantity_fact,
                'quantity_rounded': quantity_rounded,
                'unit': unit,
                'price_per_unit': price,
                'package_multiplicity': package_multiplicity,
                'package_unit': package_unit,
                'cost': cost_calc['cost_rounded'],
                'waste': cost_calc['waste'],
                'waste_percent': cost_calc['waste_percent']
            }
            
            materials_list.append(material_item)
            total_cost += cost_calc['cost_rounded']
            
            logger.debug(f"Рассчитан: {product_name} - {quantity_rounded:.2f} {unit} = {cost_calc['cost_rounded']:.0f} тг")
            
        except Exception as e:
            logger.error(f"Ошибка расчета материала {item.get('Товар', 'UNKNOWN')}: {e}")
            continue
    
    # === ГРУППИРОВКА ПО КАТЕГОРИЯМ ===
    
    grouped = group_materials_by_category(materials_list)
    
    # === СТАТИСТИКА ===
    
    total_waste_cost = sum(m['waste'] * m['price_per_unit'] for m in materials_list)
    
    summary = {
        'total_items': len(materials_list),
        'total_cost': total_cost,
        'total_waste_cost': total_waste_cost,
        'waste_percent': (total_waste_cost / total_cost * 100) if total_cost > 0 else 0
    }
    
    return {
        'materials': materials_list,
        'grouped': grouped,
        'summary': summary,
        'geometry': geometry
    }


def group_materials_by_category(materials: List[Dict]) -> Dict:
    """
    Группировка материалов по категориям
    
    Args:
        materials: Список материалов
    
    Returns:
        Dict с группами
    """
    
    # Определяем категории по элементам
    categories = {
        'Профили': [],
        'Комплектующие': [],
        'Уплотнители': [],
        'Крепления': [],
        'Прочее': []
    }
    
    for mat in materials:
        element = mat['element']
        
        if 'Ригель' in element or 'Стойка' in element or 'Крышка' in element or 'Прижимной' in element:
            categories['Профили'].append(mat)
        elif 'соединитель' in element or 'Соединитель' in element or 'Термомост' in element or 'Заглушка' in element:
            categories['Комплектующие'].append(mat)
        elif 'Упл' in element or 'уплотнитель' in element.lower():
            categories['Уплотнители'].append(mat)
        elif 'Кронштейн' in element or 'Монтажная' in element or 'пластина' in element.lower():
            categories['Крепления'].append(mat)
        else:
            categories['Прочее'].append(mat)
    
    # Считаем суммы по категориям
    for cat_name, cat_items in categories.items():
        cat_cost = sum(item['cost'] for item in cat_items)
        categories[cat_name] = {
            'items': cat_items,
            'count': len(cat_items),
            'total_cost': cat_cost
        }
    
    return categories


# ========== РАСЧЕТ ДЛЯ ВСТАВОК (ОКНА/ДВЕРИ) ==========

def calculate_insert_materials(insert_data: Dict,
                               insert_system: str,
                               window_ref1: List[Dict],
                               window_ref2: Dict,
                               window_ref3: List[Dict]) -> Dict:
    """
    Расчет материалов для вставки (окна/двери) в фасаде
    
    Использует существующий движок расчета окон/дверей
    
    Args:
        insert_data: Данные вставки из модального окна
        insert_system: Система профиля вставки (ALG 2030-XX)
        window_ref1-3: Справочники для окон/дверей
    
    Returns:
        Dict с результатами расчета вставки
    """
    
    try:
        # Импортируем функцию расчета окон
        from calculations.engine_windows import calculate_window_smeta
        
        # Формируем данные позиции для движка окон
        position_data = {
            'width': insert_data.get('width', 2000),
            'height': insert_data.get('height', 1560),
            'imposts': insert_data.get('imposts', {}),
            'sashes': insert_data.get('sashes', []),
            'fill_category': insert_data.get('fill_category', 'Стеклопакет'),
            'glass_type': insert_data.get('glass_type', 'Двойной'),
            'count': 1  # Всегда 1 для вставки
        }
        
        # Определяем тип изделия
        product_type = "Окно с откр."  # По умолчанию
        if len(insert_data.get('sashes', [])) == 0:
            product_type = "Окно глух."
        
        # Вызываем расчет
        result = calculate_window_smeta(
            position_data=position_data,
            system_id=insert_system,
            product_type=product_type,
            count=1,
            ref1=window_ref1,
            ref2=window_ref2,
            ref3=window_ref3,
            toning_id="Нет",
            assembly_id="Нет",
            installation_id="Нет"
        )
        
        return {
            'success': True,
            'insert_result': result,
            'insert_system': insert_system
        }
        
    except Exception as e:
        logger.error(f"Ошибка расчета вставки: {e}")
        return {
            'success': False,
            'error': str(e)
        }


# ========== РАСЧЕТ АДАПТЕРА РАМЫ ==========

def calculate_frame_adapter(insert_data: Dict, facade_ref: List[Dict]) -> Dict:
    """
    Расчет адаптера рамы для крепления вставки к фасаду
    
    Args:
        insert_data: Данные вставки
        facade_ref: Справочник фасадов
    
    Returns:
        Dict с расчетом адаптера
    """
    
    # Периметр вставки
    w_s = insert_data.get('width', 0) / 1000  # в метры
    h_s = insert_data.get('height', 0) / 1000
    
    perimeter = (w_s + h_s) * 2
    
    # Ищем адаптер рамы в справочнике
    adapter = None
    for item in facade_ref:
        if 'Адаптер рамы' in item.get('Элемент', ''):
            adapter = item
            break
    
    if not adapter:
        return {'error': 'Адаптер рамы не найден в справочнике'}
    
    # Формула: (w_s + h_s) * 2
    quantity_fact = perimeter
    
    price = float(adapter.get('Цена за единицу', 0))
    package_multiplicity = float(adapter.get('Кратность к упаковке', 6))
    
    quantity_rounded = round_to_package(quantity_fact, package_multiplicity)
    
    cost = quantity_rounded * price
    
    return {
        'product_name': adapter.get('Товар', ''),
        'article': adapter.get('Артикул', ''),
        'quantity_fact': quantity_fact,
        'quantity_rounded': quantity_rounded,
        'unit': 'м',
        'price_per_unit': price,
        'cost': cost
    }


# ========== РАСЧЕТ СТЕКЛОПАКЕТОВ/ПАНЕЛЕЙ ==========

def calculate_glass_materials(geometry: Dict, 
                              panel_type: str = 'glass',
                              glass_type: str = 'Двойной',
                              ref2: Dict = None) -> Dict:
    """
    Расчет стеклопакетов или панелей для глухого остекления
    
    Args:
        geometry: Результат calculate_vitrazh_geometry()
        panel_type: 'glass', 'lambry_no_thermo', 'lambry_with_thermo'
        glass_type: Тип стеклопакета (если panel_type='glass')
        ref2: Справочник 2 с ценами
    
    Returns:
        Dict с расчетом стеклопакетов/панелей
    """
    
    # Площадь остекления
    from facade_geometry import calculate_glass_area
    glass_area_m2 = calculate_glass_area(geometry, panel_type)
    
    if panel_type == 'glass':
        # Стеклопакет - цена за м²
        if ref2 and glass_type in ref2:
            price_m2 = float(ref2[glass_type].get('Цена за кв.м.', 0))
        else:
            # Дефолтные цены если справочника нет
            default_prices = {
                'Двойной': 9000,
                'Тройной': 14000,
                'Энергодвойной': 12000,
                'Энерготройной': 15000,
                'Одинарный 4мм': 4000,
                'Одинарный 6мм': 6000,
                'Нет': 0
            }
            price_m2 = default_prices.get(glass_type, 9000)
        
        cost = glass_area_m2 * price_m2
        
        return {
            'type': 'Стеклопакет',
            'glass_type': glass_type,
            'area_m2': glass_area_m2,
            'price_per_m2': price_m2,
            'cost': cost,
            'unit': 'м²'
        }
    
    else:
        # Ламбри - цена за погонный метр
        # Количество ламбри = периметр всех ячеек
        n_cells = geometry['total_cells']
        perimeter_one_cell = 2 * (geometry['cell_width_m'] + geometry['cell_height_m'])
        total_perimeter = perimeter_one_cell * n_cells
        
        # Цены за метр
        lambry_prices = {
            'lambry_no_thermo': 2248,
            'lambry_with_thermo': 4588
        }
        
        price_per_m = lambry_prices.get(panel_type, 2248)
        cost = total_perimeter * price_per_m
        
        return {
            'type': 'Ламбри',
            'panel_type': panel_type,
            'length_m': total_perimeter,
            'price_per_m': price_per_m,
            'cost': cost,
            'unit': 'м'
        }


# ========== ФОРМИРОВАНИЕ ИТОГОВОЙ СПЕЦИФИКАЦИИ ==========

def create_full_specification(facade_materials: Dict,
                             glass_materials: Dict = None,
                             insert_materials: Dict = None,
                             frame_adapter: Dict = None) -> Dict:
    """
    Создание полной спецификации фасада со всеми материалами
    
    Args:
        facade_materials: Результат calculate_facade_materials()
        glass_materials: Результат calculate_glass_materials()
        insert_materials: Результат calculate_insert_materials()
        frame_adapter: Результат calculate_frame_adapter()
    
    Returns:
        Полная спецификация
    """
    
    specification = {
        'facade_profiles': facade_materials,
        'glass_panels': glass_materials,
        'inserts': insert_materials,
        'frame_adapter': frame_adapter,
        'total_cost': 0
    }
    
    # Считаем общую стоимость
    total = facade_materials['summary']['total_cost']
    
    if glass_materials:
        total += glass_materials.get('cost', 0)
    
    if insert_materials and insert_materials.get('success'):
        insert_result = insert_materials['insert_result']
        total += insert_result.get('Итоговая стоимость', 0)
    
    if frame_adapter and 'cost' in frame_adapter:
        total += frame_adapter['cost']
    
    specification['total_cost'] = total
    
    return specification


# ========== ТЕСТИРОВАНИЕ ==========

if __name__ == "__main__":
    print("=== Тест расчета материалов фасада ===\n")
    
    # Импортируем геометрию
    from facade_geometry import calculate_vitrazh_geometry
    
    # Создаем тестовую геометрию
    geometry = calculate_vitrazh_geometry(W=6.0, H=3.0, n_columns=3, n_rows=2)
    
    # Создаем тестовый справочник (упрощенный)
    test_reference = [
        {
            'Система профиля': 'Ruit 50F',
            'Тип изделия': 'Витраж',
            'Элемент': 'Ригель 50 мм',
            'Товар': 'Ruit 50F Ригель фасадный 50 мм RAL 7024 мат',
            'Артикул': '2-00-5013-60-7024-',
            'Формула_Python': 'w_cell * count_r50',
            'Ед. измерения': 'м',
            'Цена за единицу': 4511,
            'Кратность к упаковке': 6,
            'Ед. упаковки': 'м'
        },
        {
            'Система профиля': 'Ruit 50F',
            'Тип изделия': 'Витраж',
            'Элемент': 'U-соединитель ригеля 50мм',
            'Товар': 'Ruit 50F "U"соединитель ригеля 5013 (40 мм)',
            'Артикул': '2-11-5953-00-0400',
            'Формула_Python': '2 * count_r50',
            'Ед. измерения': 'шт',
            'Цена за единицу': 151,
            'Кратность к упаковке': 1,
            'Ед. упаковки': 'шт'
        },
        {
            'Система профиля': 'Ruit 50F',
            'Тип изделия': 'Витраж',
            'Элемент': 'Стойка 90 мм',
            'Товар': 'Ruit 50F Стойка фасадная 90 мм БКК New RAL 7024 мат',
            'Артикул': '2-00-5035-60-7024-',
            'Формула_Python': 'H_m * count_m',
            'Ед. измерения': 'м',
            'Цена за единицу': 5994,
            'Кратность к упаковке': 6,
            'Ед. упаковки': 'м'
        },
    ]
    
    # Расчет материалов
    result = calculate_facade_materials(geometry, test_reference)
    
    print(f"Всего материалов: {result['summary']['total_items']}")
    print(f"Общая стоимость: {result['summary']['total_cost']:,.0f} тг")
    print(f"Отходы: {result['summary']['waste_percent']:.1f}%\n")
    
    print("Материалы:")
    for mat in result['materials']:
        print(f"  • {mat['product_name']}")
        print(f"    Кол-во: {mat['quantity_fact']:.2f} → {mat['quantity_rounded']:.2f} {mat['unit']}")
        print(f"    Цена: {mat['cost']:,.0f} тг")
        print()
    
    print("\n=== Тест расчета стеклопакетов ===\n")
    
    glass_result = calculate_glass_materials(
        geometry=geometry,
        panel_type='glass',
        glass_type='Двойной',
        ref2={'Двойной': {'Цена за кв.м.': 9000}}
    )
    
    print(f"Тип: {glass_result['type']}")
    print(f"Площадь: {glass_result['area_m2']:.2f} м²")
    print(f"Стоимость: {glass_result['cost']:,.0f} тг")
