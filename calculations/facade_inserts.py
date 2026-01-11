"""
Модуль обработки вставок в фасадах
===================================

Обрабатывает вставки окон, дверей и панелей в конструкции фасада.
Интегрируется с существующей системой расчета окон/дверей.

Автор: Axis Pro GF
Версия: 1.0
"""

import logging
from typing import Dict, List, Optional, Tuple
import copy

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# ==========================================
# КЛАСС МЕНЕДЖЕРА ВСТАВОК
# ==========================================

class FacadeInsertManager:
    """
    Менеджер для управления вставками в фасаде
    """
    
    def __init__(self):
        self.inserts = []
        logger.info("Инициализация менеджера вставок")
    
    def add_insert(
        self,
        insert_type: str,
        position: Tuple[float, float],
        width: float,
        height: float,
        properties: Optional[Dict] = None
    ) -> int:
        """
        Добавить вставку в фасад
        
        Args:
            insert_type: Тип вставки ("window", "door", "panel")
            position: Позиция (x, y) в мм
            width: Ширина вставки (мм)
            height: Высота вставки (мм)
            properties: Дополнительные свойства
        
        Returns:
            ID добавленной вставки
        
        Example:
            >>> manager = FacadeInsertManager()
            >>> insert_id = manager.add_insert("window", (1000, 500), 1200, 1500)
        """
        
        insert_id = len(self.inserts) + 1
        
        insert = {
            'id': insert_id,
            'type': insert_type,
            'position': position,
            'width': width,
            'height': height,
            'area_m2': (width * height) / 1_000_000,
            'properties': properties or {}
        }
        
        self.inserts.append(insert)
        
        logger.info(f"Добавлена вставка #{insert_id}: {insert_type} {width}x{height} мм")
        
        return insert_id
    
    def remove_insert(self, insert_id: int) -> bool:
        """
        Удалить вставку по ID
        
        Args:
            insert_id: ID вставки
        
        Returns:
            True если вставка удалена
        """
        
        initial_count = len(self.inserts)
        self.inserts = [i for i in self.inserts if i['id'] != insert_id]
        
        if len(self.inserts) < initial_count:
            logger.info(f"Вставка #{insert_id} удалена")
            return True
        
        logger.warning(f"Вставка #{insert_id} не найдена")
        return False
    
    def get_insert(self, insert_id: int) -> Optional[Dict]:
        """
        Получить вставку по ID
        
        Args:
            insert_id: ID вставки
        
        Returns:
            Dict с данными вставки или None
        """
        
        for insert in self.inserts:
            if insert['id'] == insert_id:
                return insert
        
        return None
    
    def get_all_inserts(self) -> List[Dict]:
        """
        Получить все вставки
        
        Returns:
            Список всех вставок
        """
        return self.inserts.copy()
    
    def get_inserts_by_type(self, insert_type: str) -> List[Dict]:
        """
        Получить вставки по типу
        
        Args:
            insert_type: Тип вставки ("window", "door", "panel")
        
        Returns:
            Список вставок заданного типа
        """
        return [i for i in self.inserts if i['type'] == insert_type]
    
    def calculate_total_area(self, insert_type: Optional[str] = None) -> float:
        """
        Рассчитать общую площадь вставок
        
        Args:
            insert_type: Тип вставки (опционально, для фильтрации)
        
        Returns:
            Общая площадь в м²
        """
        
        if insert_type:
            inserts = self.get_inserts_by_type(insert_type)
        else:
            inserts = self.inserts
        
        total = sum(i['area_m2'] for i in inserts)
        return round(total, 3)
    
    def clear(self):
        """Очистить все вставки"""
        self.inserts = []
        logger.info("Все вставки удалены")


# ==========================================
# ФУНКЦИЯ: ИНТЕГРАЦИЯ С СИСТЕМОЙ ОКОН
# ==========================================

def calculate_insert_with_window_system(
    insert_width: float,
    insert_height: float,
    profile_system: str,
    glass_type: str,
    reference_data: tuple
) -> Dict:
    """
    Расчет вставки с использованием существующей системы расчета окон
    
    Args:
        insert_width: Ширина вставки (мм)
        insert_height: Высота вставки (мм)
        profile_system: Система профиля
        glass_type: Тип стеклопакета
        reference_data: Кортеж (ref1, ref2, ref3) из Google Sheets
    
    Returns:
        Dict с результатами расчета вставки
    
    Example:
        >>> result = calculate_insert_with_window_system(
        ...     insert_width=1200,
        ...     insert_height=1500,
        ...     profile_system="ALG RUIT 73i 22MM",
        ...     glass_type="Двойной",
        ...     reference_data=(ref1, ref2, ref3)
        ... )
    """
    
    logger.info(f"Расчет вставки через систему окон: {insert_width}x{insert_height} мм")
    
    try:
        # Импорт функции расчета окон
        from calculations.engine_windows import calculate_window
        
        # Формируем данные для расчета как окна
        window_data = {
            'width': insert_width,
            'height': insert_height,
            'product_type': 'Окно глух.',  # Глухое окно как вставка
            'system_id': profile_system,
            'glass_type': glass_type,
            'imposts': {
                'auto_calculate': False,
                'left': 0,
                'center': 0,
                'right': 0,
                'tor': 0
            },
            'sashes': []  # Без створок
        }
        
        # Вызываем расчет окна
        result = calculate_window(
            W=insert_width,
            H=insert_height,
            system_id=profile_system,
            glass_type=glass_type,
            ref1=reference_data[0],
            ref2=reference_data[1],
            ref3=reference_data[2],
            has_imposts=False,
            sashes=[]
        )
        
        logger.info(f"✓ Вставка рассчитана через систему окон")
        
        return {
            'width': insert_width,
            'height': insert_height,
            'area_m2': (insert_width * insert_height) / 1_000_000,
            'calculation_result': result,
            'materials': result.get('materials', []),
            'cost': result.get('total_cost', 0)
        }
        
    except ImportError as e:
        logger.error(f"Ошибка импорта модуля окон: {e}")
        # Возвращаем базовый расчет без интеграции
        return {
            'width': insert_width,
            'height': insert_height,
            'area_m2': (insert_width * insert_height) / 1_000_000,
            'error': 'Модуль расчета окон недоступен',
            'cost': 0
        }
    
    except Exception as e:
        logger.error(f"Ошибка расчета вставки: {e}")
        return {
            'width': insert_width,
            'height': insert_height,
            'area_m2': (insert_width * insert_height) / 1_000_000,
            'error': str(e),
            'cost': 0
        }


# ==========================================
# ФУНКЦИЯ: ПРОВЕРКА ВПИСЫВАЕТСЯ ЛИ ВСТАВКА
# ==========================================

def check_insert_fits(
    facade_width: float,
    facade_height: float,
    insert_position: Tuple[float, float],
    insert_width: float,
    insert_height: float,
    existing_inserts: List[Dict] = None
) -> Dict:
    """
    Проверка, помещается ли вставка в фасаде без пересечений
    
    Args:
        facade_width: Ширина фасада (мм)
        facade_height: Высота фасада (мм)
        insert_position: Позиция вставки (x, y) в мм
        insert_width: Ширина вставки (мм)
        insert_height: Высота вставки (мм)
        existing_inserts: Список существующих вставок
    
    Returns:
        Dict с результатом проверки:
        {
            'fits': bool,
            'errors': List[str],
            'warnings': List[str]
        }
    
    Example:
        >>> result = check_insert_fits(6000, 3000, (1000, 500), 1200, 1500)
        >>> if result['fits']:
        >>>     print("Вставка помещается!")
    """
    
    x, y = insert_position
    errors = []
    warnings = []
    
    # Проверка 1: Выход за границы фасада
    if x < 0 or y < 0:
        errors.append("Вставка выходит за левую/нижнюю границу фасада")
    
    if x + insert_width > facade_width:
        errors.append(f"Вставка выходит за правую границу фасада (x={x}, ширина={insert_width}, макс={facade_width})")
    
    if y + insert_height > facade_height:
        errors.append(f"Вставка выходит за верхнюю границу фасада (y={y}, высота={insert_height}, макс={facade_height})")
    
    # Проверка 2: Минимальные отступы от краев
    MIN_MARGIN = 100  # мм
    
    if x < MIN_MARGIN:
        warnings.append(f"Малый отступ слева ({x} мм), рекомендуется минимум {MIN_MARGIN} мм")
    
    if y < MIN_MARGIN:
        warnings.append(f"Малый отступ снизу ({y} мм), рекомендуется минимум {MIN_MARGIN} мм")
    
    if facade_width - (x + insert_width) < MIN_MARGIN:
        warnings.append(f"Малый отступ справа, рекомендуется минимум {MIN_MARGIN} мм")
    
    if facade_height - (y + insert_height) < MIN_MARGIN:
        warnings.append(f"Малый отступ сверху, рекомендуется минимум {MIN_MARGIN} мм")
    
    # Проверка 3: Пересечение с другими вставками
    if existing_inserts:
        for existing in existing_inserts:
            ex_pos = existing['position']
            ex_width = existing['width']
            ex_height = existing['height']
            
            # Проверка пересечения прямоугольников
            if not (x + insert_width < ex_pos[0] or  # Справа от существующей
                    x > ex_pos[0] + ex_width or       # Слева от существующей
                    y + insert_height < ex_pos[1] or  # Ниже существующей
                    y > ex_pos[1] + ex_height):       # Выше существующей
                errors.append(f"Пересечение с вставкой #{existing['id']}")
    
    fits = len(errors) == 0
    
    result = {
        'fits': fits,
        'errors': errors,
        'warnings': warnings
    }
    
    if fits:
        logger.info(f"✓ Вставка {insert_width}x{insert_height} помещается в позиции ({x}, {y})")
    else:
        logger.warning(f"✗ Вставка не помещается: {', '.join(errors)}")
    
    return result


# ==========================================
# ФУНКЦИЯ: АВТОМАТИЧЕСКОЕ РАЗМЕЩЕНИЕ ВСТАВОК
# ==========================================

def auto_place_inserts(
    facade_width: float,
    facade_height: float,
    inserts_specs: List[Dict],
    spacing: float = 500
) -> List[Dict]:
    """
    Автоматическое размещение вставок в фасаде
    
    Args:
        facade_width: Ширина фасада (мм)
        facade_height: Высота фасада (мм)
        inserts_specs: Список спецификаций вставок [{'width': 1200, 'height': 1500, 'type': 'window'}, ...]
        spacing: Минимальный отступ между вставками (мм)
    
    Returns:
        Список вставок с рассчитанными позициями
    
    Example:
        >>> specs = [
        ...     {'width': 1200, 'height': 1500, 'type': 'window'},
        ...     {'width': 1000, 'height': 2000, 'type': 'door'}
        ... ]
        >>> placed = auto_place_inserts(6000, 3000, specs)
    """
    
    logger.info(f"Автоматическое размещение {len(inserts_specs)} вставок в фасаде {facade_width}x{facade_height} мм")
    
    placed_inserts = []
    current_x = spacing
    current_y = spacing
    max_height_in_row = 0
    
    for idx, spec in enumerate(inserts_specs):
        insert_width = spec['width']
        insert_height = spec['height']
        
        # Проверяем, помещается ли в текущий ряд
        if current_x + insert_width + spacing > facade_width:
            # Переходим на новый ряд
            current_x = spacing
            current_y += max_height_in_row + spacing
            max_height_in_row = 0
        
        # Проверяем, помещается ли по высоте
        if current_y + insert_height + spacing > facade_height:
            logger.warning(f"Вставка #{idx+1} не помещается: превышена высота фасада")
            continue
        
        # Размещаем вставку
        position = (current_x, current_y)
        
        placed_inserts.append({
            'id': idx + 1,
            'type': spec.get('type', 'window'),
            'position': position,
            'width': insert_width,
            'height': insert_height,
            'area_m2': (insert_width * insert_height) / 1_000_000,
            'properties': spec.get('properties', {})
        })
        
        # Обновляем позицию для следующей вставки
        current_x += insert_width + spacing
        max_height_in_row = max(max_height_in_row, insert_height)
        
        logger.info(f"✓ Вставка #{idx+1} размещена в позиции {position}")
    
    logger.info(f"✓ Размещено {len(placed_inserts)} из {len(inserts_specs)} вставок")
    
    return placed_inserts


# ==========================================
# ТЕСТОВЫЙ КОД
# ==========================================

if __name__ == "__main__":
    print("=" * 60)
    print("ТЕСТИРОВАНИЕ МОДУЛЯ facade_inserts.py")
    print("=" * 60)
    
    # Тест 1: Менеджер вставок
    print("\n[ТЕСТ 1] Работа с менеджером вставок")
    try:
        manager = FacadeInsertManager()
        
        # Добавляем вставки
        id1 = manager.add_insert("window", (1000, 500), 1200, 1500)
        id2 = manager.add_insert("door", (3000, 0), 1000, 2400)
        id3 = manager.add_insert("panel", (5000, 1000), 800, 1200)
        
        print(f"✓ Добавлено вставок: {len(manager.get_all_inserts())}")
        print(f"✓ Окон: {len(manager.get_inserts_by_type('window'))}")
        print(f"✓ Дверей: {len(manager.get_inserts_by_type('door'))}")
        print(f"✓ Общая площадь: {manager.calculate_total_area():.2f} м²")
        
        # Удаляем вставку
        manager.remove_insert(id2)
        print(f"✓ После удаления осталось: {len(manager.get_all_inserts())}")
    except Exception as e:
        print(f"✗ Ошибка: {e}")
    
    # Тест 2: Проверка вписывается ли вставка
    print("\n[ТЕСТ 2] Проверка размещения вставки")
    try:
        result = check_insert_fits(
            facade_width=6000,
            facade_height=3000,
            insert_position=(1000, 500),
            insert_width=1200,
            insert_height=1500
        )
        print(f"✓ Вставка помещается: {result['fits']}")
        if result['warnings']:
            print(f"⚠ Предупреждения: {result['warnings']}")
        if result['errors']:
            print(f"✗ Ошибки: {result['errors']}")
    except Exception as e:
        print(f"✗ Ошибка: {e}")
    
    # Тест 3: Проверка выхода за границы
    print("\n[ТЕСТ 3] Проверка выхода за границы")
    try:
        result = check_insert_fits(
            facade_width=6000,
            facade_height=3000,
            insert_position=(5500, 2000),
            insert_width=1200,
            insert_height=1500
        )
        print(f"✓ Вставка помещается: {result['fits']}")
        if result['errors']:
            print(f"✗ Ошибки: {result['errors']}")
    except Exception as e:
        print(f"✗ Ошибка: {e}")
    
    # Тест 4: Автоматическое размещение
    print("\n[ТЕСТ 4] Автоматическое размещение вставок")
    try:
        specs = [
            {'width': 1200, 'height': 1500, 'type': 'window'},
            {'width': 1000, 'height': 2000, 'type': 'door'},
            {'width': 1200, 'height': 1500, 'type': 'window'},
            {'width': 800, 'height': 1200, 'type': 'panel'}
        ]
        
        placed = auto_place_inserts(
            facade_width=8000,
            facade_height=3000,
            inserts_specs=specs,
            spacing=300
        )
        
        print(f"✓ Размещено вставок: {len(placed)}")
        for insert in placed:
            print(f"  - {insert['type']} {insert['width']}x{insert['height']} в позиции {insert['position']}")
    except Exception as e:
        print(f"✗ Ошибка: {e}")
    
    print("\n" + "=" * 60)
    print("ТЕСТИРОВАНИЕ ЗАВЕРШЕНО")
    print("=" * 60)
