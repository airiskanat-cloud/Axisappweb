"""
Модуль глобальной агрегации материалов V.9

КОНЦЕПЦИЯ:
Переход от попозиционного расчёта к проектному.
Все материалы суммируются в единую базу заказа, округление применяется ОДИН РАЗ.

КАТЕГОРИИ:
1. КАРКАС (facade) - стойки, ригели, кронштейны, прижимки
2. ВСТАВКИ (inserts) - профили, фурнитура окон/дверей в фасадах
3. ОКНА/ДВЕРИ (windows_doors) - материалы окон/дверей
4. ТАМБУР (tambour) - направляющие тамбура

АВТОР: Claude V10
ДАТА: 26.01.2026
"""

import math
from typing import Dict, List, Any, Tuple
from collections import defaultdict


def safe_float(value, default=0.0):
    """Безопасное преобразование в float"""
    try:
        if value is None:
            return default
        s = str(value).replace(",", ".").replace(" ", "").replace("\xa0", "")
        if s == "":
            return default
        return float(s)
    except:
        return default


def get_package_size(article: str, ref1: List[Dict]) -> float:
    """
    Получить размер упаковки (кратность) из Справочника-1
    
    Args:
        article: Артикул материала
        ref1: Справочник-1
    
    Returns:
        Размер упаковки в метрах (по умолчанию 6.0 для профилей)
    """
    for item in ref1:
        if item.get('Артикул', '') == article:
            krat = safe_float(item.get('Кратность', 6.0), 6.0)
            if krat > 0:
                return krat
            break
    
    # По умолчанию: профили по 6м, всё остальное по 1
    if any(keyword in str(article).upper() for keyword in ['ПРОФИЛ', 'РАМА', 'СТВОР', 'ИМПОСТ', 'РИГЕЛ', 'СТОЙК']):
        return 6.0
    else:
        return 1.0


class MaterialAggregator:
    """
    Глобальный агрегатор материалов по категориям
    Реализует требования ТЗ V.9
    """
    
    def __init__(self, ref1: List[Dict]):
        """
        Инициализация агрегатора
        
        Args:
            ref1: Справочник-1 для определения кратности
        """
        self.ref1 = ref1
        
        # Категории материалов
        self.categories = {
            'facade_frame': {},      # Каркас фасада
            'facade_inserts': {},    # Вставки фасада
            'windows_doors': {},     # Окна/Двери
            'tambour': {}            # Тамбур
        }
        
        # Метрики
        self.metrics = {
            'total_area': 0.0,
            'total_perimeter': 0.0
        }
        
        # Услуги и стеклопакеты (НЕ материалы!)
        self.services = {
            'glass_total_area': 0.0,
            'glass_cost': 0.0,
            'lambri_cost': 0.0,
            'toning_cost': 0.0,
            'assembly_cost': 0.0,
            'installation_cost': 0.0,
            'additional_details_cost': 0.0
        }
    
    def add_material(self, category: str, article: str, quantity_raw: float, 
                     unit: str, price: float, name: str = ""):
        """
        Добавить материал в категорию БЕЗ округления
        
        Args:
            category: Категория ('facade_frame', 'facade_inserts', 'windows_doors', 'tambour')
            article: Артикул
            quantity_raw: Количество ДО округления (чистое)
            unit: Единица измерения
            price: Цена за единицу
            name: Название элемента
        """
        if category not in self.categories:
            print(f"⚠️ Неизвестная категория: {category}")
            return
        
        qty = safe_float(quantity_raw, 0)
        if qty <= 0:
            return
        
        # Суммируем по артикулу
        if article not in self.categories[category]:
            self.categories[category][article] = {
                'article': article,
                'name': name,
                'quantity_raw': 0.0,
                'quantity_rounded': 0.0,
                'unit': unit,
                'price': safe_float(price, 0),
                'cost': 0.0,
                'package_size': get_package_size(article, self.ref1)
            }
        
        self.categories[category][article]['quantity_raw'] += qty
    
    def add_metrics(self, area: float, perimeter: float):
        """Добавить метрики площади и периметра"""
        self.metrics['total_area'] += safe_float(area, 0)
        self.metrics['total_perimeter'] += safe_float(perimeter, 0)
    
    def add_service(self, service_type: str, value: float):
        """
        Добавить стоимость услуги
        
        Args:
            service_type: Тип услуги (glass_cost, assembly_cost, и т.д.)
            value: Стоимость или площадь
        """
        if service_type in self.services:
            self.services[service_type] += safe_float(value, 0)
    
    def round_all_materials(self):
        """
        Округлить ВСЕ материалы до кратности ОДИН РАЗ
        Это ключевая функция для экономии 80% на профилях
        """
        for category in self.categories:
            for article, data in self.categories[category].items():
                qty_raw = data['quantity_raw']
                package_size = data['package_size']
                
                # Округление вверх до кратности
                packages_needed = math.ceil(qty_raw / package_size)
                qty_rounded = packages_needed * package_size
                
                # Сохраняем
                data['quantity_rounded'] = qty_rounded
                data['cost'] = qty_rounded * data['price']
    
    def get_category_materials(self, category: str) -> List[Dict]:
        """
        Получить список материалов категории (после округления)
        
        Returns:
            Список словарей с материалами
        """
        if category not in self.categories:
            return []
        
        materials = []
        for article, data in self.categories[category].items():
            materials.append({
                'Артикул': data['article'],
                'Элемент': data['name'],
                'Количество_raw': round(data['quantity_raw'], 3),
                'Кратность': data['package_size'],
                'Количество': round(data['quantity_rounded'], 2),
                'Единица': data['unit'],
                'Цена': round(data['price'], 2),
                'Сумма': round(data['cost'], 0)
            })
        
        # Сортируем по названию
        materials.sort(key=lambda x: x['Элемент'])
        return materials
    
    def get_category_total(self, category: str) -> float:
        """Получить общую стоимость категории"""
        if category not in self.categories:
            return 0.0
        
        return sum(data['cost'] for data in self.categories[category].values())
    
    def calculate_final_totals(self, margin_rate: float = 0.81) -> Dict:
        """
        Рассчитать финальные итоги по проектному методу
        
        Args:
            margin_rate: Коэффициент обеспечения (0.81 = 81%)
        
        Returns:
            Словарь с итогами
        """
        # Материалы всех категорий
        materials_total = sum(self.get_category_total(cat) for cat in self.categories)
        
        # Услуги
        services_total = (
            self.services['glass_cost'] +
            self.services['lambri_cost'] +
            self.services['toning_cost'] +
            self.services['assembly_cost'] +
            self.services['installation_cost'] +
            self.services['additional_details_cost']
        )
        
        # Себестоимость
        subtotal = materials_total + services_total
        
        # Обеспечение ОДИН РАЗ
        margin = subtotal * margin_rate
        
        # Итого
        total = subtotal + margin
        
        return {
            'materials_total': round(materials_total, 0),
            'services_total': round(services_total, 0),
            'subtotal': round(subtotal, 0),
            'margin': round(margin, 0),
            'total': round(total, 0),
            'breakdown': {
                'facade_frame': round(self.get_category_total('facade_frame'), 0),
                'facade_inserts': round(self.get_category_total('facade_inserts'), 0),
                'windows_doors': round(self.get_category_total('windows_doors'), 0),
                'tambour': round(self.get_category_total('tambour'), 0),
                'glass': round(self.services['glass_cost'], 0),
                'lambri': round(self.services['lambri_cost'], 0),
                'toning': round(self.services['toning_cost'], 0),
                'assembly': round(self.services['assembly_cost'], 0),
                'installation': round(self.services['installation_cost'], 0),
                'additional_details': round(self.services['additional_details_cost'], 0)
            }
        }
    
    def get_all_materials_for_export(self) -> List[Dict]:
        """
        Получить ВСЕ материалы для экспорта в Excel
        Объединяет все категории
        """
        all_materials = []
        
        # Добавляем заголовки для каждой категории
        categories_names = {
            'facade_frame': 'КАРКАС ФАСАДА',
            'facade_inserts': 'ВСТАВКИ ФАСАДА (окна/двери)',
            'windows_doors': 'ОКНА И ДВЕРИ',
            'tambour': 'ОКОННЫЙ ТАМБУР'
        }
        
        for cat_key, cat_name in categories_names.items():
            materials = self.get_category_materials(cat_key)
            if materials:
                # Добавляем заголовок категории
                all_materials.append({
                    'Артикул': '',
                    'Элемент': f'=== {cat_name} ===',
                    'Количество_raw': '',
                    'Количество': '',
                    'Единица': '',
                    'Цена': '',
                    'Сумма': ''
                })
                # Добавляем материалы
                all_materials.extend(materials)
                # Добавляем итого по категории
                total = self.get_category_total(cat_key)
                all_materials.append({
                    'Артикул': '',
                    'Элемент': f'ИТОГО {cat_name}',
                    'Количество_raw': '',
                    'Количество': '',
                    'Единица': '',
                    'Цена': '',
                    'Сумма': round(total, 0)
                })
                # Пустая строка
                all_materials.append({
                    'Артикул': '', 'Элемент': '', 'Количество_raw': '',
                    'Количество': '', 'Единица': '', 'Цена': '', 'Сумма': ''
                })
        
        return all_materials


def extract_materials_from_facade_result(result: Dict, category: str) -> List[Dict]:
    """
    Извлечь материалы из результата расчёта фасада
    
    Args:
        result: Результат расчёта (из engine_facade)
        category: Категория ('facade_frame' или 'facade_inserts')
    
    Returns:
        Список материалов с полем quantity_raw
    """
    materials = []
    
    if category == 'facade_frame':
        # Каркас: стойки, ригели, прижимки, уплотнители, кронштейны
        frame_materials = result.get('frame_materials', [])
        for mat in frame_materials:
            materials.append({
                'article': mat.get('Артикул', ''),
                'name': mat.get('Элемент', ''),
                'quantity_raw': mat.get('quantity_raw', mat.get('Количество', 0)),
                'unit': mat.get('Единица', 'м'),
                'price': mat.get('Цена', 0)
            })
    
    elif category == 'facade_inserts':
        # Вставки: профили, фурнитура окон/дверей
        insert_materials = result.get('insert_materials', [])
        for mat in insert_materials:
            materials.append({
                'article': mat.get('Артикул', ''),
                'name': mat.get('Элемент', ''),
                'quantity_raw': mat.get('quantity_raw', mat.get('Количество', 0)),
                'unit': mat.get('Единица', 'шт'),
                'price': mat.get('Цена', 0)
            })
    
    return materials


def extract_materials_from_windows_result(result: Dict) -> List[Dict]:
    """
    Извлечь материалы из результата расчёта окон/дверей
    
    Args:
        result: Результат расчёта (из engine_windows)
    
    Returns:
        Список материалов с полем quantity_raw
    """
    materials = []
    
    part2_materials = result.get('part2_materials', [])
    for mat in part2_materials:
        materials.append({
            'article': mat.get('Артикул', ''),
            'name': mat.get('Элемент', ''),
            'quantity_raw': mat.get('Количество_raw', mat.get('Количество', 0)),
            'unit': mat.get('Единица', 'шт'),
            'price': mat.get('Цена', 0)
        })
    
    return materials
