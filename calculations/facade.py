"""
Главный движок расчета фасадов (Ruit 50F и Оконный тамбур)

Этот модуль объединяет все подмодули:
- facade_geometry: расчет геометрии
- facade_validation: проверка на ветровую нагрузку
- facade_materials: расчет материалов
- engine_windows: расчет вставок (окна/двери)

Основная функция: calculate_facade_full()
"""

import logging
from typing import Dict, List
from datetime import datetime

# Импорты локальных модулей
try:
    # В продакшн версии (в папке calculations/)
    from calculations.facade_geometry import (
        calculate_vitrazh_geometry,
        calculate_tambour_facade_geometry,
        validate_geometry as validate_geometry_basic
    )
    from calculations.facade_validation import (
        calculate_wind_load,
        validate_spans,
        calculate_anchor_spacing,
        full_validation
    )
    from calculations.facade_materials import (
        calculate_facade_materials,
        calculate_glass_materials,
        calculate_insert_materials,
        calculate_frame_adapter,
        create_full_specification
    )
except ModuleNotFoundError:
    # Для тестирования (в той же папке)
    from facade_geometry import (
        calculate_vitrazh_geometry,
        calculate_tambour_facade_geometry,
        validate_geometry as validate_geometry_basic
    )
    from facade_validation import (
        calculate_wind_load,
        validate_spans,
        calculate_anchor_spacing,
        full_validation
    )
    from facade_materials import (
        calculate_facade_materials,
        calculate_glass_materials,
        calculate_insert_materials,
        calculate_frame_adapter,
        create_full_specification
    )

logger = logging.getLogger(__name__)


# ========== ГЛАВНАЯ ФУНКЦИЯ РАСЧЕТА ==========

def calculate_facade_full(facade_type: str,
                         positions: List[Dict],
                         facade_reference: List[Dict],
                         window_ref1: List[Dict] = None,
                         window_ref2: Dict = None,
                         window_ref3: List[Dict] = None,
                         installation_height_m: float = 3.0,
                         wind_zone: str = 'III',
                         terrain_type: str = 'B',
                         toning_id: str = 'Нет',
                         assembly_id: str = 'Нет',
                         installation_id: str = 'Нет') -> Dict:
    """
    ГЛАВНАЯ ФУНКЦИЯ: Полный расчет фасада
    
    Args:
        facade_type: "Фасадная система (Ruit 50F)" или "Оконный тамбур (ALG)"
        positions: Список позиций фасада [
            {
                'width': 6.0,  # м
                'height': 3.0,  # м
                'columns': 3,
                'rows': 2,
                'filling_type': 'blind' / 'window' / 'door',
                'blind_data': {...},  # если filling_type='blind'
                'insert_data': {...},  # если filling_type='window'/'door'
                'insert_system': 'ALG 2030-63C'
            },
            ...
        ]
        facade_reference: Справочник фасадов (из Google Sheets)
        window_ref1-3: Справочники для окон/дверей
        installation_height_m: Высота установки от земли
        wind_zone: Ветровой район ('I'-'V')
        terrain_type: Тип местности ('A'/'B')
        toning_id, assembly_id, installation_id: Дополнительные услуги
    
    Returns:
        Dict с полными результатами расчета
    """
    
    logger.info(f"🚀 Начало расчета фасада: {facade_type}")
    logger.info(f"   Позиций: {len(positions)}")
    
    try:
        # Определяем тип системы
        is_vitrazh = "Ruit 50F" in facade_type or "Фасадная" in facade_type
        
        if is_vitrazh:
            # === ФАСАДНАЯ СИСТЕМА RUIT 50F (ВИТРАЖ) ===
            result = calculate_vitrazh_facade(
                positions=positions,
                facade_reference=facade_reference,
                window_ref1=window_ref1,
                window_ref2=window_ref2,
                window_ref3=window_ref3,
                installation_height_m=installation_height_m,
                wind_zone=wind_zone,
                terrain_type=terrain_type,
                toning_id=toning_id,
                assembly_id=assembly_id,
                installation_id=installation_id
            )
        else:
            # === ОКОННЫЙ ТАМБУР (ALG) ===
            result = calculate_tambour_facade(
                positions=positions,
                window_ref1=window_ref1,
                window_ref2=window_ref2,
                window_ref3=window_ref3,
                toning_id=toning_id,
                assembly_id=assembly_id,
                installation_id=installation_id
            )
        
        logger.info(f"✅ Расчет завершен успешно")
        
        return result
        
    except Exception as e:
        logger.error(f"❌ Ошибка расчета фасада: {e}")
        import traceback
        logger.error(traceback.format_exc())
        
        return {
            'success': False,
            'error': str(e),
            'traceback': traceback.format_exc()
        }


# ========== РАСЧЕТ ФАСАДНОЙ СИСТЕМЫ (RUIT 50F) ==========

def calculate_vitrazh_facade(positions: List[Dict],
                             facade_reference: List[Dict],
                             window_ref1: List[Dict] = None,
                             window_ref2: Dict = None,
                             window_ref3: List[Dict] = None,
                             installation_height_m: float = 3.0,
                             wind_zone: str = 'III',
                             terrain_type: str = 'B',
                             toning_id: str = 'Нет',
                             assembly_id: str = 'Нет',
                             installation_id: str = 'Нет') -> Dict:
    """
    Расчет фасадной системы Ruit 50F (витраж)
    
    Алгоритм:
    1. Расчет геометрии каждой позиции
    2. Проверка на ветровую нагрузку и статику
    3. Расчет материалов профилей
    4. Расчет стеклопакетов/панелей
    5. Расчет вставок (окна/двери) если есть
    6. Формирование итоговой спецификации
    """
    
    logger.info("📐 Расчет фасадной системы Ruit 50F")
    
    position_results = []
    total_cost = 0
    all_errors = []
    all_warnings = []
    
    # === ОБРАБОТКА КАЖДОЙ ПОЗИЦИИ ===
    
    for idx, pos in enumerate(positions, start=1):
        logger.info(f"Позиция {idx}: {pos.get('width')}×{pos.get('height')}м")
        
        try:
            # 1. ГЕОМЕТРИЯ
            geometry = calculate_vitrazh_geometry(
                W=pos.get('width', 6.0),
                H=pos.get('height', 3.0),
                n_columns=pos.get('columns', 3),
                n_rows=pos.get('rows', 2)
            )
            
            # 2. ВАЛИДАЦИЯ
            validation = full_validation(
                geometry=geometry,
                installation_height_m=installation_height_m,
                wind_zone=wind_zone,
                terrain_type=terrain_type,
                allow_reinforcement=True
            )
            
            # Собираем ошибки и предупреждения
            if validation['span_validation']['errors']:
                all_errors.extend(validation['span_validation']['errors'])
            if validation['span_validation']['warnings']:
                all_warnings.extend(validation['span_validation']['warnings'])
            
            # 3. МАТЕРИАЛЫ ПРОФИЛЕЙ
            materials = calculate_facade_materials(
                geometry=geometry,
                facade_ref=facade_reference,
                system="Ruit 50F"
            )
            
            # 4. ЗАПОЛНЕНИЕ (Стеклопакеты/Панели или Вставки)
            filling_type = pos.get('filling_type', 'blind')
            glass_materials = None
            insert_materials = None
            frame_adapter = None
            
            if filling_type == 'blind':
                # ГЛУХОЕ ОСТЕКЛЕНИЕ
                blind_data = pos.get('blind_data', {})
                panel_type = blind_data.get('panel_type', 'glass')
                glass_type = blind_data.get('glass_type', 'Двойной')
                
                glass_materials = calculate_glass_materials(
                    geometry=geometry,
                    panel_type=panel_type,
                    glass_type=glass_type,
                    ref2=window_ref2
                )
                
                logger.info(f"  Глухое остекление: {glass_type}, {glass_materials.get('area_m2', 0):.2f} м²")
            
            elif filling_type in ['window', 'door']:
                # ВСТАВКА (ОКНО/ДВЕРЬ)
                insert_data = pos.get('insert_data', {})
                insert_system = pos.get('insert_system', 'ALG 2030-63C')
                
                if insert_data and window_ref1:
                    insert_materials = calculate_insert_materials(
                        insert_data=insert_data,
                        insert_system=insert_system,
                        window_ref1=window_ref1,
                        window_ref2=window_ref2,
                        window_ref3=window_ref3
                    )
                    
                    # Адаптер рамы для крепления вставки
                    frame_adapter = calculate_frame_adapter(
                        insert_data=insert_data,
                        facade_ref=facade_reference
                    )
                    
                    logger.info(f"  Вставка: {filling_type}, система {insert_system}")
            
            # 5. ПОЛНАЯ СПЕЦИФИКАЦИЯ ПОЗИЦИИ
            specification = create_full_specification(
                facade_materials=materials,
                glass_materials=glass_materials,
                insert_materials=insert_materials,
                frame_adapter=frame_adapter
            )
            
            position_cost = specification['total_cost']
            total_cost += position_cost
            
            # Сохраняем результат позиции
            position_result = {
                'position_number': idx,
                'geometry': geometry,
                'validation': validation,
                'materials': materials,
                'glass_materials': glass_materials,
                'insert_materials': insert_materials,
                'frame_adapter': frame_adapter,
                'specification': specification,
                'position_cost': position_cost
            }
            
            position_results.append(position_result)
            
            logger.info(f"  ✅ Позиция {idx}: {position_cost:,.0f} тг")
            
        except Exception as e:
            logger.error(f"❌ Ошибка в позиции {idx}: {e}")
            all_errors.append(f"Позиция {idx}: {str(e)}")
            continue
    
    # === ОБЩИЕ РЕЗУЛЬТАТЫ ===
    
    # Применяем дополнительные услуги
    services_cost = calculate_services(
        total_cost=total_cost,
        toning_id=toning_id,
        assembly_id=assembly_id,
        installation_id=installation_id,
        ref2=window_ref2
    )
    
    final_cost = total_cost + services_cost['total_services_cost']
    
    # Формируем итоговый результат
    result = {
        'success': True,
        'facade_type': 'Фасадная система (Ruit 50F)',
        'calculation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        
        # Позиции
        'positions': position_results,
        'total_positions': len(position_results),
        
        # Стоимость
        'materials_cost': total_cost,
        'services': services_cost,
        'total_cost': final_cost,
        
        # Валидация
        'errors': all_errors,
        'warnings': all_warnings,
        'is_valid': len(all_errors) == 0,
        
        # Параметры расчета
        'parameters': {
            'installation_height_m': installation_height_m,
            'wind_zone': wind_zone,
            'terrain_type': terrain_type,
            'toning': toning_id,
            'assembly': assembly_id,
            'installation': installation_id
        }
    }
    
    return result


# ========== РАСЧЕТ ОКОННОГО ТАМБУРА ==========

def calculate_tambour_facade(positions: List[Dict],
                             window_ref1: List[Dict] = None,
                             window_ref2: Dict = None,
                             window_ref3: List[Dict] = None,
                             toning_id: str = 'Нет',
                             assembly_id: str = 'Нет',
                             installation_id: str = 'Нет') -> Dict:
    """
    Расчет оконного тамбура (ALG)
    
    Тамбур состоит из готовых оконных/дверных блоков,
    соединенных трубами и адаптерами
    
    Алгоритм:
    1. Каждая позиция = готовое окно/дверь
    2. Расчет через engine_windows
    3. Добавление труб и адаптеров для соединения
    """
    
    logger.info("🏠 Расчет оконного тамбура")
    
    position_results = []
    total_cost = 0
    
    # === ОБРАБОТКА КАЖДОЙ ПОЗИЦИИ (ГОТОВОЕ ОКНО/ДВЕРЬ) ===
    
    for idx, pos in enumerate(positions, start=1):
        logger.info(f"Позиция {idx}: {pos.get('filling_type', 'window')}")
        
        try:
            # Для тамбура каждая позиция - это готовое окно/дверь
            filling_type = pos.get('filling_type', 'window')
            insert_data = pos.get('insert_data', {})
            insert_system = pos.get('insert_system', 'ALG 2030-63C')
            
            if filling_type in ['window', 'door'] and insert_data and window_ref1:
                # Расчет окна/двери
                insert_result = calculate_insert_materials(
                    insert_data=insert_data,
                    insert_system=insert_system,
                    window_ref1=window_ref1,
                    window_ref2=window_ref2,
                    window_ref3=window_ref3
                )
                
                if insert_result.get('success'):
                    position_cost = insert_result['insert_result'].get('Итоговая стоимость', 0)
                    total_cost += position_cost
                    
                    position_result = {
                        'position_number': idx,
                        'type': filling_type,
                        'system': insert_system,
                        'insert_result': insert_result,
                        'position_cost': position_cost
                    }
                    
                    position_results.append(position_result)
                    
                    logger.info(f"  ✅ Позиция {idx}: {position_cost:,.0f} тг")
            
        except Exception as e:
            logger.error(f"❌ Ошибка в позиции {idx}: {e}")
            continue
    
    # TODO: Добавить расчет труб и адаптеров для соединения позиций
    # Пока это упрощенный расчет
    
    # Применяем дополнительные услуги
    services_cost = calculate_services(
        total_cost=total_cost,
        toning_id=toning_id,
        assembly_id=assembly_id,
        installation_id=installation_id,
        ref2=window_ref2
    )
    
    final_cost = total_cost + services_cost['total_services_cost']
    
    result = {
        'success': True,
        'facade_type': 'Оконный тамбур (ALG)',
        'calculation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        
        'positions': position_results,
        'total_positions': len(position_results),
        
        'materials_cost': total_cost,
        'services': services_cost,
        'total_cost': final_cost,
        
        'parameters': {
            'toning': toning_id,
            'assembly': assembly_id,
            'installation': installation_id
        }
    }
    
    return result


# ========== РАСЧЕТ ДОПОЛНИТЕЛЬНЫХ УСЛУГ ==========

def calculate_services(total_cost: float,
                      toning_id: str,
                      assembly_id: str,
                      installation_id: str,
                      ref2: Dict = None) -> Dict:
    """
    Расчет дополнительных услуг (тонировка, сборка, монтаж)
    
    Args:
        total_cost: Стоимость материалов
        toning_id: "Есть" / "Нет"
        assembly_id: "Есть" / "Нет"
        installation_id: "Монтаж" / "Демонтаж/Монтаж" / "Сложный монтаж" / "Нет"
        ref2: Справочник 2 с ценами услуг
    
    Returns:
        Dict с расчетом услуг
    """
    
    services = {
        'toning': {'id': toning_id, 'cost': 0},
        'assembly': {'id': assembly_id, 'cost': 0},
        'installation': {'id': installation_id, 'cost': 0},
        'total_services_cost': 0
    }
    
    # TODO: Реализовать расчет услуг из ref2
    # Пока возвращаем нулевые значения
    
    return services


# ========== ФОРМАТИРОВАНИЕ РЕЗУЛЬТАТА ДЛЯ ВЫВОДА ==========

def format_result_for_display(result: Dict) -> str:
    """
    Форматирование результата для отображения пользователю
    
    Args:
        result: Результат calculate_facade_full()
    
    Returns:
        Отформатированная строка
    """
    
    if not result.get('success'):
        return f"❌ ОШИБКА РАСЧЕТА:\n{result.get('error', 'Неизвестная ошибка')}"
    
    lines = []
    
    # Заголовок
    lines.append("=" * 60)
    lines.append(f"📊 РАСЧЕТ ФАСАДА: {result['facade_type']}")
    lines.append(f"📅 Дата: {result['calculation_date']}")
    lines.append("=" * 60)
    lines.append("")
    
    # Позиции
    lines.append(f"📦 ПОЗИЦИЙ: {result['total_positions']}")
    lines.append("")
    
    for pos in result.get('positions', []):
        pos_num = pos['position_number']
        pos_cost = pos['position_cost']
        
        lines.append(f"Позиция {pos_num}: {pos_cost:,.0f} тг")
        
        if 'geometry' in pos:
            geom = pos['geometry']
            lines.append(f"  📐 Размер: {geom['width_m']:.1f} × {geom['height_m']:.1f} м")
            lines.append(f"  🔲 Сетка: {geom['n_columns']} × {geom['n_rows']} (ячейка {geom['cell_width_m']:.2f} × {geom['cell_height_m']:.2f} м)")
            lines.append(f"  🔧 Профили: Стойка {geom['profile_types']['stoyка_type']}мм, Ригель {geom['profile_types']['rigel_type']}мм")
        
        lines.append("")
    
    # Стоимость
    lines.append("💰 СТОИМОСТЬ:")
    lines.append(f"  Материалы: {result['materials_cost']:,.0f} тг")
    
    services = result.get('services', {})
    if services.get('total_services_cost', 0) > 0:
        lines.append(f"  Услуги: {services['total_services_cost']:,.0f} тг")
    
    lines.append(f"  {'─' * 40}")
    lines.append(f"  ИТОГО: {result['total_cost']:,.0f} тг")
    lines.append("")
    
    # Предупреждения
    if result.get('warnings'):
        lines.append("⚠️ ПРЕДУПРЕЖДЕНИЯ:")
        for warning in result['warnings']:
            lines.append(f"  • {warning}")
        lines.append("")
    
    # Ошибки
    if result.get('errors'):
        lines.append("❌ ОШИБКИ:")
        for error in result['errors']:
            lines.append(f"  • {error}")
        lines.append("")
    
    lines.append("=" * 60)
    
    return "\n".join(lines)


# ========== ЭКСПОРТ РЕЗУЛЬТАТА ==========

def export_result_to_dict(result: Dict) -> Dict:
    """
    Подготовка результата для экспорта в Excel или JSON
    
    Args:
        result: Результат calculate_facade_full()
    
    Returns:
        Упрощенный Dict для экспорта
    """
    
    export_data = {
        'Дата расчета': result.get('calculation_date'),
        'Тип фасада': result.get('facade_type'),
        'Количество позиций': result.get('total_positions'),
        'Стоимость материалов': result.get('materials_cost'),
        'Стоимость услуг': result.get('services', {}).get('total_services_cost', 0),
        'Итоговая стоимость': result.get('total_cost'),
        'Позиции': []
    }
    
    for pos in result.get('positions', []):
        pos_data = {
            'Номер': pos['position_number'],
            'Стоимость': pos['position_cost']
        }
        
        if 'geometry' in pos:
            geom = pos['geometry']
            pos_data.update({
                'Ширина (м)': geom['width_m'],
                'Высота (м)': geom['height_m'],
                'Столбцов': geom['n_columns'],
                'Рядов': geom['n_rows'],
                'Стойка': f"{geom['profile_types']['stoyка_type']}мм",
                'Ригель': f"{geom['profile_types']['rigel_type']}мм"
            })
        
        export_data['Позиции'].append(pos_data)
    
    return export_data


# ========== ТЕСТИРОВАНИЕ ==========

if __name__ == "__main__":
    print("=" * 60)
    print("🧪 ТЕСТ ГЛАВНОГО ДВИЖКА РАСЧЕТА ФАСАДОВ")
    print("=" * 60)
    print()
    
    # Загружаем тестовые данные
    from facade_geometry import calculate_vitrazh_geometry
    
    # Создаем тестовый справочник
    test_facade_ref = [
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
    
    # Создаем тестовые позиции
    test_positions = [
        {
            'width': 6.0,
            'height': 3.0,
            'columns': 3,
            'rows': 2,
            'filling_type': 'blind',
            'blind_data': {
                'panel_type': 'glass',
                'glass_type': 'Двойной'
            }
        }
    ]
    
    # Тестовый справочник цен
    test_ref2 = {
        'Двойной': {'Цена за кв.м.': 9000}
    }
    
    # ЗАПУСК РАСЧЕТА
    print("🚀 Запуск полного расчета фасада...")
    print()
    
    result = calculate_facade_full(
        facade_type="Фасадная система (Ruit 50F)",
        positions=test_positions,
        facade_reference=test_facade_ref,
        window_ref2=test_ref2,
        installation_height_m=5.0,
        wind_zone='III',
        terrain_type='B'
    )
    
    # ВЫВОД РЕЗУЛЬТАТА
    if result['success']:
        print(format_result_for_display(result))
    else:
        print(f"❌ ОШИБКА: {result.get('error')}")
