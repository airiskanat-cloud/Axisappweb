"""
Модуль расчета геометрии фасадов (Ruit 50F)

Этот модуль выполняет:
1. Расчет размеров ячеек фасада
2. Определение количества и типов профилей (стойки, ригели)
3. Расчет общих длин и площадей
4. Подготовку переменных для формул из справочника
"""

import math
import logging
from typing import Dict, List, Tuple

logger = logging.getLogger(__name__)


# ========== АЛГОРИТМ АВТОПОДБОРА ПРОФИЛЕЙ RUIT 50F ==========

def auto_select_profile_types(height_m: float) -> Dict[str, str]:
    """
    Автоматический подбор типов профилей на основе высоты фасада
    
    Args:
        height_m: Высота фасада в метрах
    
    Returns:
        Dict с типами стоек и ригелей:
        {
            'stoyка_type': '90', '110' или '130',
            'rigel_type': '50', '70', '85', '95', '105', '115', '135',
            'warning': bool (True если нужна статика)
        }
    """
    if height_m <= 3.0:
        return {
            'stoyка_type': '90',  # Стойка 90 мм (2-00-5035)
            'rigel_type': '50',   # Ригель 50 мм (2-00-5013)
            'status': 'Стандарт',
            'description': 'Оптимально для частных домов и витрин',
            'warning': False
        }
    elif height_m <= 4.5:
        return {
            'stoyка_type': '130',  # Стойка 130 мм (2-00-5033)
            'rigel_type': '85',    # Ригель 85 мм (2-00-5014)
            'status': 'Усиленный',
            'description': 'Для высоких первых этажей и ветреных зон',
            'warning': False
        }
    else:
        return {
            'stoyка_type': '130',  # Стойка 130 мм (максимальная)
            'rigel_type': '135',   # Ригель 135 мм (максимальный)
            'status': 'ВНИМАНИЕ',
            'description': 'Требуется расчет статики и стальное армирование!',
            'warning': True
        }


def calculate_vitrazh_geometry(W: float, H: float, n_columns: int, n_rows: int) -> Dict:
    """
    Расчет геометрии витража (фасадная система Ruit 50F)
    
    Args:
        W: Ширина фасада в метрах
        H: Высота фасада в метрах
        n_columns: Количество столбцов (вертикальных делений)
        n_rows: Количество рядов (горизонтальных делений)
    
    Returns:
        Dict с геометрическими параметрами
    """
    
    # === РАЗМЕРЫ ЯЧЕЕК ===
    w_cell = W / n_columns  # Ширина одной ячейки
    h_cell = H / n_rows     # Высота одной ячейки
    
    # === КОЛИЧЕСТВО ПРОФИЛЕЙ ===
    # Стойки (вертикальные): n_columns + 1 (включая крайние)
    count_mullions = n_columns + 1
    
    # Ригели (горизонтальные): n_rows + 1 (включая верхний и нижний)
    count_transoms = n_rows + 1
    
    # === АВТОПОДБОР ТИПОВ ПРОФИЛЕЙ ===
    profile_types = auto_select_profile_types(H)
    
    # === ДЛИНЫ ПРОФИЛЕЙ ===
    # Одна стойка = высота фасада
    length_one_mullion = H
    
    # Один ригель = ширина фасада
    length_one_transom = W
    
    # Общая длина стоек
    total_length_mullions = length_one_mullion * count_mullions
    
    # Общая длина ригелей
    total_length_transoms = length_one_transom * count_transoms
    
    # Общая длина всех профилей
    total_length_all_profiles = total_length_mullions + total_length_transoms
    
    # === ПЛОЩАДИ ===
    area_total_m2 = W * H  # Общая площадь фасада
    area_one_cell_m2 = w_cell * h_cell  # Площадь одной ячейки
    
    # === КОЛИЧЕСТВО УЗЛОВ КРЕПЛЕНИЯ ===
    # Обычно 2 точки крепления на каждую стойку
    count_anchors = count_mullions * 2
    
    # === ПЕРЕМЕННЫЕ ДЛЯ ФОРМУЛ (из справочника) ===
    formula_vars = {
        # Размеры
        'W': W,                    # Ширина фасада (м)
        'H': H,                    # Высота фасада (м)
        'w_cell': w_cell,          # Ширина ячейки (м)
        'h_cell': h_cell,          # Высота ячейки (м)
        'H_m': H,                  # Alias для высоты (используется в формулах)
        
        # Количество элементов
        'n_columns': n_columns,    # Количество столбцов
        'n_rows': n_rows,          # Количество рядов
        'count_m': count_mullions, # Количество стоек (mullions)
        
        # Количество ригелей по типам (зависит от автоподбора)
        'count_r50': count_transoms if profile_types['rigel_type'] == '50' else 0,
        'count_r70': count_transoms if profile_types['rigel_type'] == '70' else 0,
        'count_r85': count_transoms if profile_types['rigel_type'] == '85' else 0,
        'count_r95': count_transoms if profile_types['rigel_type'] == '95' else 0,
        'count_r105': count_transoms if profile_types['rigel_type'] == '105' else 0,
        'count_r115': count_transoms if profile_types['rigel_type'] == '115' else 0,
        'count_r135': count_transoms if profile_types['rigel_type'] == '135' else 0,
        
        # Суммарные значения
        'total_count_rigels': count_transoms,
        'total_length_all_profiles': total_length_all_profiles,
        'count_anchors': count_anchors,
    }
    
    return {
        # Базовые размеры
        'width_m': W,
        'height_m': H,
        'area_total_m2': area_total_m2,
        
        # Ячейки
        'cell_width_m': w_cell,
        'cell_height_m': h_cell,
        'cell_area_m2': area_one_cell_m2,
        'n_columns': n_columns,
        'n_rows': n_rows,
        'total_cells': n_columns * n_rows,
        
        # Профили
        'profile_types': profile_types,
        'count_mullions': count_mullions,
        'count_transoms': count_transoms,
        'length_mullions_m': total_length_mullions,
        'length_transoms_m': total_length_transoms,
        'length_all_profiles_m': total_length_all_profiles,
        
        # Крепления
        'count_anchors': count_anchors,
        
        # Переменные для формул
        'formula_vars': formula_vars
    }


def calculate_tambour_facade_geometry(positions: List[Dict]) -> Dict:
    """
    Расчет геометрии тамбур-фасада (оконный тамбур)
    
    Тамбур состоит из нескольких позиций (обычно 4 стороны)
    
    Args:
        positions: Список позиций фасада, каждая с:
            - width_m: ширина
            - height_m: высота
            - filling_type: 'blind', 'window', 'door'
    
    Returns:
        Dict с суммарными параметрами
    """
    
    total_area = 0
    total_perimeter = 0
    
    for pos in positions:
        W = pos.get('width_m', 0)
        H = pos.get('height_m', 0)
        
        total_area += W * H
        total_perimeter += 2 * (W + H)
    
    return {
        'total_area_m2': total_area,
        'total_perimeter_m': total_perimeter,
        'n_positions': len(positions),
        'positions': positions
    }


def calculate_tambour_window_geometry(W: float, H: float, 
                                      has_insert: bool = False,
                                      insert_data: Dict = None) -> Dict:
    """
    Расчет геометрии одной позиции тамбур-окна/двери
    
    Args:
        W: Ширина позиции (м)
        H: Высота позиции (м)
        has_insert: Есть ли вставка (окно/дверь)
        insert_data: Данные вставки (если есть)
    
    Returns:
        Dict с параметрами позиции
    """
    
    area_m2 = W * H
    perimeter_m = 2 * (W + H)
    
    result = {
        'width_m': W,
        'height_m': H,
        'area_m2': area_m2,
        'perimeter_m': perimeter_m,
        'has_insert': has_insert
    }
    
    # Если есть вставка, добавляем её параметры
    if has_insert and insert_data:
        result['insert'] = {
            'width_m': insert_data.get('width', 0) / 1000,
            'height_m': insert_data.get('height', 0) / 1000,
            'system': insert_data.get('system', 'ALG 2030-63C')
        }
    
    return result


def calculate_glass_area(geometry: Dict, panel_type: str = 'glass') -> float:
    """
    Расчет площади остекления с учетом профилей
    
    Args:
        geometry: Результат calculate_vitrazh_geometry()
        panel_type: 'glass' или 'lambry'
    
    Returns:
        Площадь остекления в м²
    """
    
    # Для фасадной системы учитываем толщину профилей
    # Ширина профиля Ruit 50F ≈ 50 мм = 0.05 м
    profile_width = 0.05
    
    W = geometry['width_m']
    H = geometry['height_m']
    n_columns = geometry['n_columns']
    n_rows = geometry['n_rows']
    
    # Площадь, занимаемая стойками (вертикальные профили)
    # (n_columns + 1) стоек × высота × толщина профиля
    area_mullions = (n_columns + 1) * H * profile_width
    
    # Площадь, занимаемая ригелями (горизонтальные профили)
    # (n_rows + 1) ригелей × ширина × толщина профиля
    area_transoms = (n_rows + 1) * W * profile_width
    
    # Чистая площадь остекления
    glass_area = W * H - area_mullions - area_transoms
    
    return max(glass_area, 0)


def validate_geometry(geometry: Dict) -> Dict:
    """
    Проверка геометрии фасада на соответствие ограничениям
    
    Args:
        geometry: Результат calculate_vitrazh_geometry()
    
    Returns:
        Dict с результатами проверки:
        {
            'is_valid': bool,
            'errors': List[str],
            'warnings': List[str]
        }
    """
    
    errors = []
    warnings = []
    
    W = geometry['width_m']
    H = geometry['height_m']
    w_cell = geometry['cell_width_m']
    h_cell = geometry['cell_height_m']
    
    # Проверка минимальных размеров
    if W < 0.5:
        errors.append("Ширина фасада слишком мала (< 0.5 м)")
    
    if H < 0.5:
        errors.append("Высота фасада слишком мала (< 0.5 м)")
    
    # Проверка максимальных размеров ячейки
    MAX_CELL_WIDTH = 2.5  # максимальная ширина ячейки
    MAX_CELL_HEIGHT = 3.0  # максимальная высота ячейки
    
    if w_cell > MAX_CELL_WIDTH:
        warnings.append(f"Ширина ячейки {w_cell:.2f}м превышает рекомендуемую {MAX_CELL_WIDTH}м")
    
    if h_cell > MAX_CELL_HEIGHT:
        warnings.append(f"Высота ячейки {h_cell:.2f}м превышает рекомендуемую {MAX_CELL_HEIGHT}м")
    
    # Проверка на необходимость статического расчета
    if geometry['profile_types']['warning']:
        warnings.append("⚠️ ВНИМАНИЕ: Высота > 4.5м - требуется расчет статики!")
    
    is_valid = len(errors) == 0
    
    return {
        'is_valid': is_valid,
        'errors': errors,
        'warnings': warnings
    }


# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========

def calculate_cell_dimensions(W: float, H: float, n_columns: int, n_rows: int) -> Tuple[float, float]:
    """
    Простой расчет размеров ячейки
    
    Returns:
        (width_cell, height_cell) в метрах
    """
    return (W / n_columns, H / n_rows)


def get_profile_article(profile_type: str, element: str) -> str:
    """
    Получение артикула профиля по типу и элементу
    
    Args:
        profile_type: '50', '70', '85', '90', '110', '130' и т.д.
        element: 'stoyка', 'rigel', 'u_connector', 'seal'
    
    Returns:
        Артикул профиля
    """
    
    # Справочник артикулов (соответствует таблице Excel)
    articles = {
        # Ригели
        ('50', 'rigel'): '2-00-5013-60-7024-',
        ('70', 'rigel'): '2-00-5019-60-7024-',
        ('85', 'rigel'): '2-00-5014-60-7024-',
        ('95', 'rigel'): '2-00-5018-60-7024-',
        ('105', 'rigel'): '2-00-5015-60-7024-',
        ('115', 'rigel'): '2-00-5038-60-7024-',
        ('135', 'rigel'): '2-00-5037-60-7024-',
        
        # Стойки
        ('90', 'stoyка'): '2-00-5035-60-7024-',
        ('110', 'stoyка'): '2-00-5034-60-7024-',
        ('130', 'stoyка'): '2-00-5033-60-7024-',
        
        # U-соединители ригелей
        ('50', 'u_connector'): '2-11-5953-00-0400',
        ('70', 'u_connector'): '2-11-5953-00-0600',
        ('85', 'u_connector'): '2-11-5953-00-0600',
        ('95', 'u_connector'): '2-11-5953-00-0860',
        ('105', 'u_connector'): '2-11-5953-00-0800',
        ('115', 'u_connector'): '2-11-5953-00-0106',
        ('135', 'u_connector'): '2-11-5953-00-0126',
    }
    
    return articles.get((profile_type, element), 'UNKNOWN')


if __name__ == "__main__":
    # Тестирование модуля
    print("=== Тест расчета геометрии витража ===")
    
    # Пример: фасад 6×3 м, разделенный на 3 столбца и 2 ряда
    result = calculate_vitrazh_geometry(W=6.0, H=3.0, n_columns=3, n_rows=2)
    
    print(f"Размеры фасада: {result['width_m']}×{result['height_m']} м")
    print(f"Размер ячейки: {result['cell_width_m']:.2f}×{result['cell_height_m']:.2f} м")
    print(f"Общая площадь: {result['area_total_m2']:.2f} м²")
    print(f"Количество стоек: {result['count_mullions']}")
    print(f"Количество ригелей: {result['count_transoms']}")
    print(f"Общая длина профилей: {result['length_all_profiles_m']:.2f} м")
    print(f"Тип профилей: Стойка {result['profile_types']['stoyка_type']}мм, Ригель {result['profile_types']['rigel_type']}мм")
    
    # Проверка валидности
    validation = validate_geometry(result)
    print(f"\nВалидность: {validation['is_valid']}")
    if validation['warnings']:
        print("Предупреждения:")
        for w in validation['warnings']:
            print(f"  - {w}")
