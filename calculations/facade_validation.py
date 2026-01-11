"""
Модуль проверки фасадных конструкций на ветровую нагрузку и статику

Этот модуль выполняет:
1. Расчет ветровой нагрузки по СНиП РК
2. Проверку допустимых пролетов для профилей
3. Валидацию конструктивной схемы
4. Рекомендации по усилению
"""

import math
import logging
from typing import Dict, List, Tuple

logger = logging.getLogger(__name__)


# ========== НОРМАТИВНЫЕ ДАННЫЕ ==========

# Ветровые районы Казахстана (СНиП РК)
WIND_ZONES_KZ = {
    'I': {'w0': 0.23, 'description': 'Минимальная нагрузка'},
    'II': {'w0': 0.30, 'description': 'Средняя нагрузка'},
    'III': {'w0': 0.38, 'description': 'Повышенная нагрузка'},
    'IV': {'w0': 0.48, 'description': 'Высокая нагрузка'},
    'V': {'w0': 0.60, 'description': 'Максимальная нагрузка'},
}

# Коэффициенты высоты (тип местности A - открытая)
HEIGHT_COEFFICIENTS = {
    'A': {  # Открытая местность
        5: 0.75,
        10: 1.0,
        20: 1.25,
        40: 1.5,
        60: 1.7,
        80: 1.85,
        100: 2.0,
    },
    'B': {  # Городская застройка
        5: 0.5,
        10: 0.65,
        20: 0.85,
        40: 1.1,
        60: 1.3,
        80: 1.45,
        100: 1.6,
    }
}

# Максимальные пролеты для профилей Ruit 50F (в метрах)
# Значения взяты из технических характеристик производителя
MAX_SPANS_RUIT50F = {
    # Стойки (вертикальные профили)
    'stoyка_90': {'max_height': 3.0, 'max_height_reinforced': 4.0},
    'stoyка_110': {'max_height': 4.0, 'max_height_reinforced': 5.0},
    'stoyка_130': {'max_height': 4.5, 'max_height_reinforced': 6.0},
    
    # Ригели (горизонтальные профили)
    'rigel_50': {'max_span': 2.0, 'max_span_reinforced': 2.5},
    'rigel_70': {'max_span': 2.2, 'max_span_reinforced': 2.8},
    'rigel_85': {'max_span': 2.5, 'max_span_reinforced': 3.0},
    'rigel_95': {'max_span': 2.7, 'max_span_reinforced': 3.2},
    'rigel_105': {'max_span': 3.0, 'max_span_reinforced': 3.5},
    'rigel_115': {'max_span': 3.2, 'max_span_reinforced': 3.8},
    'rigel_135': {'max_span': 3.5, 'max_span_reinforced': 4.0},
}

# Коэффициенты аэродинамики для фасадов
AERO_COEFFICIENTS = {
    'flat_facade': 0.8,      # Плоский фасад
    'corner': 1.0,           # Угловая зона
    'parapet': 1.2,          # Парапетная зона
}


# ========== РАСЧЕТ ВЕТРОВОЙ НАГРУЗКИ ==========

def calculate_wind_load(height_m: float, 
                       wind_zone: str = 'III',
                       terrain_type: str = 'B',
                       aero_coeff: float = 0.8) -> Dict:
    """
    Расчет ветровой нагрузки на фасад по СНиП РК
    
    Args:
        height_m: Высота установки фасада от земли (м)
        wind_zone: Ветровой район ('I', 'II', 'III', 'IV', 'V')
        terrain_type: Тип местности ('A' - открытая, 'B' - городская)
        aero_coeff: Коэффициент аэродинамики (0.8 для плоских фасадов)
    
    Returns:
        Dict с результатами расчета:
        {
            'w0': нормативное давление ветра (кПа),
            'k_height': коэффициент высоты,
            'w_calculated': расчетная нагрузка (кПа),
            'w_force': сила на 1 м² (кг/м²)
        }
    """
    
    # Нормативное давление ветра
    if wind_zone not in WIND_ZONES_KZ:
        logger.warning(f"Неизвестный ветровой район {wind_zone}, используется III")
        wind_zone = 'III'
    
    w0 = WIND_ZONES_KZ[wind_zone]['w0']
    
    # Коэффициент высоты (интерполяция)
    k_height = interpolate_height_coefficient(height_m, terrain_type)
    
    # Коэффициент надежности по ветровой нагрузке
    gamma_f = 1.4
    
    # Расчетная ветровая нагрузка (кПа)
    w_calculated = w0 * k_height * aero_coeff * gamma_f
    
    # Сила на 1 м² (кг/м²) - для наглядности
    # 1 кПа = 1000 Па = 1000 Н/м² ≈ 102 кг/м²
    w_force = w_calculated * 102
    
    return {
        'w0_kPa': w0,
        'wind_zone': wind_zone,
        'wind_zone_description': WIND_ZONES_KZ[wind_zone]['description'],
        'k_height': k_height,
        'aero_coeff': aero_coeff,
        'gamma_f': gamma_f,
        'w_calculated_kPa': w_calculated,
        'w_force_kg_m2': w_force,
        'height_m': height_m,
        'terrain_type': terrain_type
    }


def interpolate_height_coefficient(height_m: float, terrain_type: str) -> float:
    """
    Интерполяция коэффициента высоты
    
    Args:
        height_m: Высота (м)
        terrain_type: Тип местности ('A' или 'B')
    
    Returns:
        Коэффициент высоты
    """
    
    if terrain_type not in HEIGHT_COEFFICIENTS:
        terrain_type = 'B'
    
    coeffs = HEIGHT_COEFFICIENTS[terrain_type]
    heights = sorted(coeffs.keys())
    
    # Если высота меньше минимальной
    if height_m <= heights[0]:
        return coeffs[heights[0]]
    
    # Если высота больше максимальной
    if height_m >= heights[-1]:
        return coeffs[heights[-1]]
    
    # Линейная интерполяция
    for i in range(len(heights) - 1):
        h1, h2 = heights[i], heights[i + 1]
        if h1 <= height_m <= h2:
            k1, k2 = coeffs[h1], coeffs[h2]
            # Линейная интерполяция
            k = k1 + (k2 - k1) * (height_m - h1) / (h2 - h1)
            return k
    
    return 1.0


# ========== ПРОВЕРКА ПРОЛЕТОВ ==========

def validate_spans(geometry: Dict, 
                   wind_load: Dict = None,
                   allow_reinforcement: bool = True) -> Dict:
    """
    Проверка допустимости пролетов для выбранных профилей
    
    Args:
        geometry: Результат calculate_vitrazh_geometry()
        wind_load: Результат calculate_wind_load() (опционально)
        allow_reinforcement: Разрешено ли использование усиления
    
    Returns:
        Dict с результатами проверки:
        {
            'is_valid': bool,
            'stoyка_check': {...},
            'rigel_check': {...},
            'recommendations': [...]
        }
    """
    
    profile_types = geometry['profile_types']
    H = geometry['height_m']
    w_cell = geometry['cell_width_m']
    h_cell = geometry['cell_height_m']
    
    errors = []
    warnings = []
    recommendations = []
    
    # === ПРОВЕРКА СТОЕК ===
    stoyка_type = profile_types['stoyка_type']
    stoyка_key = f'stoyка_{stoyка_type}'
    
    if stoyка_key in MAX_SPANS_RUIT50F:
        max_h = MAX_SPANS_RUIT50F[stoyка_key]['max_height']
        max_h_reinforced = MAX_SPANS_RUIT50F[stoyка_key]['max_height_reinforced']
        
        if H > max_h_reinforced:
            errors.append(
                f"⛔ КРИТИЧНО: Высота {H:.2f}м превышает максимальную "
                f"{max_h_reinforced:.2f}м даже с усилением для стойки {stoyка_type}мм"
            )
        elif H > max_h:
            if allow_reinforcement:
                warnings.append(
                    f"⚠️ Высота {H:.2f}м превышает стандартную {max_h:.2f}м. "
                    f"Требуется стальное усиление стоек."
                )
                recommendations.append("Добавить стальное армирование в стойки")
            else:
                errors.append(
                    f"⛔ Высота {H:.2f}м превышает допустимую {max_h:.2f}м "
                    f"для стойки {stoyка_type}мм без усиления"
                )
        
        stoyка_check = {
            'type': stoyка_type,
            'actual_height_m': H,
            'max_height_m': max_h,
            'max_height_reinforced_m': max_h_reinforced,
            'is_valid': H <= (max_h_reinforced if allow_reinforcement else max_h),
            'needs_reinforcement': H > max_h,
            'safety_factor': max_h / H if H > 0 else 999
        }
    else:
        stoyка_check = {'error': f'Неизвестный тип стойки: {stoyка_type}'}
    
    # === ПРОВЕРКА РИГЕЛЕЙ ===
    rigel_type = profile_types['rigel_type']
    rigel_key = f'rigel_{rigel_type}'
    
    if rigel_key in MAX_SPANS_RUIT50F:
        max_span = MAX_SPANS_RUIT50F[rigel_key]['max_span']
        max_span_reinforced = MAX_SPANS_RUIT50F[rigel_key]['max_span_reinforced']
        
        if w_cell > max_span_reinforced:
            errors.append(
                f"⛔ КРИТИЧНО: Ширина ячейки {w_cell:.2f}м превышает максимальную "
                f"{max_span_reinforced:.2f}м даже с усилением для ригеля {rigel_type}мм"
            )
        elif w_cell > max_span:
            if allow_reinforcement:
                warnings.append(
                    f"⚠️ Ширина ячейки {w_cell:.2f}м превышает стандартную {max_span:.2f}м. "
                    f"Требуется стальное усиление ригелей."
                )
                recommendations.append("Добавить стальное армирование в ригели")
            else:
                errors.append(
                    f"⛔ Ширина ячейки {w_cell:.2f}м превышает допустимую {max_span:.2f}м "
                    f"для ригеля {rigel_type}мм без усиления"
                )
        
        rigel_check = {
            'type': rigel_type,
            'actual_span_m': w_cell,
            'max_span_m': max_span,
            'max_span_reinforced_m': max_span_reinforced,
            'is_valid': w_cell <= (max_span_reinforced if allow_reinforcement else max_span),
            'needs_reinforcement': w_cell > max_span,
            'safety_factor': max_span / w_cell if w_cell > 0 else 999
        }
    else:
        rigel_check = {'error': f'Неизвестный тип ригеля: {rigel_type}'}
    
    # === ПРОВЕРКА ВЕТРОВОЙ НАГРУЗКИ ===
    if wind_load:
        w_force = wind_load['w_force_kg_m2']
        
        # Критичные значения ветровой нагрузки
        if w_force > 150:
            warnings.append(
                f"⚠️ Высокая ветровая нагрузка {w_force:.0f} кг/м². "
                f"Рекомендуется увеличить количество точек крепления."
            )
            recommendations.append("Увеличить количество анкерных креплений на 50%")
        
        if w_force > 200:
            errors.append(
                f"⛔ КРИТИЧНО: Экстремальная ветровая нагрузка {w_force:.0f} кг/м². "
                f"Требуется обязательный расчет статики!"
            )
            recommendations.append("Обязательный расчет статики инженером-конструктором")
    
    # === ОБЩИЕ РЕКОМЕНДАЦИИ ===
    if H > 4.5:
        recommendations.append("Для высоты > 4.5м обязателен расчет статики")
    
    if h_cell > 2.5 or w_cell > 2.5:
        recommendations.append("Увеличить количество делений фасада (уменьшить размер ячеек)")
    
    is_valid = len(errors) == 0
    
    return {
        'is_valid': is_valid,
        'stoyка_check': stoyка_check,
        'rigel_check': rigel_check,
        'errors': errors,
        'warnings': warnings,
        'recommendations': recommendations,
        'wind_load': wind_load
    }


# ========== РАСЧЕТ ПРОГИБОВ ==========

def calculate_deflection(span_m: float, 
                        profile_type: str,
                        load_kg_m2: float = 100) -> Dict:
    """
    Упрощенный расчет прогиба профиля (для оценки)
    
    Args:
        span_m: Пролет (м)
        profile_type: Тип профиля ('rigel_50', 'rigel_85', и т.д.)
        load_kg_m2: Нагрузка (кг/м²)
    
    Returns:
        Dict с расчетными прогибами
    """
    
    # Упрощенные моменты инерции профилей Ruit 50F (см⁴)
    # Реальные значения берутся из технической документации
    inertia_moments = {
        'rigel_50': 15.0,
        'rigel_70': 25.0,
        'rigel_85': 35.0,
        'rigel_95': 45.0,
        'rigel_105': 55.0,
        'rigel_115': 65.0,
        'rigel_135': 85.0,
    }
    
    if profile_type not in inertia_moments:
        return {'error': f'Неизвестный тип профиля: {profile_type}'}
    
    I = inertia_moments[profile_type]  # см⁴
    E = 70000  # МПа - модуль упругости алюминия
    
    # Упрощенная формула прогиба: f = 5 * q * L^4 / (384 * E * I)
    # где q - распределенная нагрузка (Н/м)
    
    L_mm = span_m * 1000  # мм
    q = load_kg_m2 * 9.81 / 1000  # Н/мм (нагрузка на 1 мм длины)
    
    # Прогиб в мм
    f_mm = (5 * q * L_mm**4) / (384 * E * I * 10000)  # 10000 для перевода см⁴ в мм⁴
    
    # Допустимый прогиб: L/200
    f_max_mm = L_mm / 200
    
    return {
        'deflection_mm': f_mm,
        'max_deflection_mm': f_max_mm,
        'is_acceptable': f_mm <= f_max_mm,
        'deflection_ratio': f_mm / f_max_mm if f_max_mm > 0 else 0,
        'span_m': span_m,
        'load_kg_m2': load_kg_m2
    }


# ========== РЕКОМЕНДАЦИИ ПО КРЕПЛЕНИЮ ==========

def calculate_anchor_spacing(geometry: Dict, wind_load: Dict = None) -> Dict:
    """
    Расчет рекомендуемого шага анкерных креплений
    
    Args:
        geometry: Результат calculate_vitrazh_geometry()
        wind_load: Результат calculate_wind_load()
    
    Returns:
        Dict с рекомендациями по креплению
    """
    
    H = geometry['height_m']
    count_mullions = geometry['count_mullions']
    
    # Базовый шаг крепления - каждые 1.5-2м по высоте
    base_spacing = 1.8
    
    # Если есть ветровая нагрузка, корректируем
    if wind_load:
        w_force = wind_load['w_force_kg_m2']
        if w_force > 150:
            base_spacing = 1.5  # Уменьшаем шаг при высокой нагрузке
        if w_force > 200:
            base_spacing = 1.2  # Еще больше уменьшаем при экстремальной нагрузке
    
    # Количество креплений на одну стойку
    anchors_per_mullion = math.ceil(H / base_spacing)
    
    # Общее количество креплений
    total_anchors = count_mullions * anchors_per_mullion
    
    return {
        'spacing_m': base_spacing,
        'anchors_per_mullion': anchors_per_mullion,
        'total_anchors': total_anchors,
        'recommendation': (
            f"Рекомендуется {anchors_per_mullion} точек крепления на стойку "
            f"с шагом ≈{base_spacing:.1f}м. Всего {total_anchors} креплений."
        )
    }


# ========== ПОЛНАЯ ПРОВЕРКА ==========

def full_validation(geometry: Dict, 
                   installation_height_m: float = 3.0,
                   wind_zone: str = 'III',
                   terrain_type: str = 'B',
                   allow_reinforcement: bool = True) -> Dict:
    """
    Полная проверка фасадной конструкции
    
    Args:
        geometry: Результат calculate_vitrazh_geometry()
        installation_height_m: Высота установки от земли (м)
        wind_zone: Ветровой район
        terrain_type: Тип местности
        allow_reinforcement: Разрешено ли усиление
    
    Returns:
        Dict со всеми результатами проверки
    """
    
    # 1. Расчет ветровой нагрузки
    wind_load = calculate_wind_load(
        height_m=installation_height_m,
        wind_zone=wind_zone,
        terrain_type=terrain_type
    )
    
    # 2. Проверка пролетов
    span_validation = validate_spans(
        geometry=geometry,
        wind_load=wind_load,
        allow_reinforcement=allow_reinforcement
    )
    
    # 3. Расчет креплений
    anchor_calc = calculate_anchor_spacing(
        geometry=geometry,
        wind_load=wind_load
    )
    
    # 4. Расчет прогибов
    rigel_type = geometry['profile_types']['rigel_type']
    deflection = calculate_deflection(
        span_m=geometry['cell_width_m'],
        profile_type=f'rigel_{rigel_type}',
        load_kg_m2=wind_load['w_force_kg_m2']
    )
    
    # 5. Общий вердикт
    overall_valid = (
        span_validation['is_valid'] and
        (not deflection.get('is_acceptable') is False)
    )
    
    return {
        'overall_valid': overall_valid,
        'wind_load': wind_load,
        'span_validation': span_validation,
        'anchor_calculation': anchor_calc,
        'deflection_check': deflection,
        'summary': generate_summary(
            overall_valid,
            span_validation,
            wind_load,
            anchor_calc
        )
    }


def generate_summary(overall_valid: bool, 
                    span_validation: Dict,
                    wind_load: Dict,
                    anchor_calc: Dict) -> str:
    """
    Генерация текстовой сводки по результатам проверки
    """
    
    summary = []
    
    if overall_valid:
        summary.append("✅ КОНСТРУКЦИЯ ДОПУСТИМА")
    else:
        summary.append("⛔ КОНСТРУКЦИЯ НЕДОПУСТИМА - ТРЕБУЮТСЯ ИЗМЕНЕНИЯ")
    
    summary.append(f"\n📊 Ветровая нагрузка: {wind_load['w_force_kg_m2']:.0f} кг/м² "
                  f"(зона {wind_load['wind_zone']})")
    
    summary.append(f"🔩 Рекомендуется: {anchor_calc['total_anchors']} анкерных креплений")
    
    if span_validation['warnings']:
        summary.append("\n⚠️ ПРЕДУПРЕЖДЕНИЯ:")
        for w in span_validation['warnings']:
            summary.append(f"  • {w}")
    
    if span_validation['errors']:
        summary.append("\n⛔ КРИТИЧНЫЕ ОШИБКИ:")
        for e in span_validation['errors']:
            summary.append(f"  • {e}")
    
    if span_validation['recommendations']:
        summary.append("\n💡 РЕКОМЕНДАЦИИ:")
        for r in span_validation['recommendations']:
            summary.append(f"  • {r}")
    
    return "\n".join(summary)


if __name__ == "__main__":
    # Тестирование модуля
    print("=== Тест валидации фасада ===\n")
    
    # Создаем тестовую геометрию
    from facade_geometry import calculate_vitrazh_geometry
    
    geometry = calculate_vitrazh_geometry(W=6.0, H=3.0, n_columns=3, n_rows=2)
    
    # Полная проверка
    validation = full_validation(
        geometry=geometry,
        installation_height_m=5.0,  # Высота 1-го этажа
        wind_zone='III',            # Средняя ветровая зона
        terrain_type='B'            # Городская застройка
    )
    
    print(validation['summary'])
    
    print("\n" + "="*50)
    print("Тест с экстремальными параметрами:")
    print("="*50 + "\n")
    
    # Тест с большой высотой
    geometry2 = calculate_vitrazh_geometry(W=8.0, H=5.5, n_columns=3, n_rows=3)
    
    validation2 = full_validation(
        geometry=geometry2,
        installation_height_m=10.0,
        wind_zone='IV',
        terrain_type='A'
    )
    
    print(validation2['summary'])
