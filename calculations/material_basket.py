"""
Модуль глобальной корзины материалов

ЦЕЛЬ: Устранить перерасход профилей за счёт:
1. Сбора ВСЕХ материалов из всех позиций БЕЗ округления
2. Суммирования по артикулам
3. Округления до упаковок ОДИН РАЗ на уровне всего заказа

ЭКОНОМИЯ: 5× на профилях (вместо 10 дверей × 2хлыста = 2 хлыста на все)
"""

import math
from typing import Dict, List, Tuple


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
        Размер упаковки в метрах (по умолчанию 6.0)
    """
    for item in ref1:
        if item.get('Артикул', '') == article:
            krat = safe_float(item.get('Кратность', 6.0), 6.0)
            if krat > 0:
                return krat
            break
    
    # По умолчанию: профили по 6м, метизы по 1шт
    if any(keyword in article.upper() for keyword in ['ПРОФИЛ', 'РАМА', 'СТВОР', 'ИМПОСТ']):
        return 6.0
    else:
        return 1.0


class MaterialBasket:
    """
    Глобальная корзина материалов для всего заказа
    
    Принцип работы:
    1. Собираем материалы из всех позиций (БЕЗ округления!)
    2. Суммируем по артикулам
    3. Округляем ОДИН РАЗ до упаковок
    4. Считаем стоимость
    """
    
    def __init__(self, ref1: List[Dict]):
        """
        Args:
            ref1: Справочник-1 с кратностями и ценами
        """
        self.ref1 = ref1
        self.materials_raw = {}  # Артикул → чистая длина (нетто)
        self.materials_rounded = {}  # Артикул → округлённая длина
        self.materials_info = {}  # Артикул → {unit, price, name}
    
    def add_material(
        self, 
        article: str, 
        quantity_raw: float,
        unit: str = "м",
        price: float = 0.0,
        name: str = ""
    ):
        """
        Добавить материал в корзину (БЕЗ округления!)
        
        Args:
            article: Артикул
            quantity_raw: Чистое количество (нетто), БЕЗ округления!
            unit: Единица измерения
            price: Цена за единицу
            name: Название материала
        """
        if article not in self.materials_raw:
            self.materials_raw[article] = 0.0
            self.materials_info[article] = {
                "unit": unit,
                "price": price,
                "name": name
            }
        
        self.materials_raw[article] += quantity_raw
    
    def round_all_materials(self):
        """
        Округлить ВСЕ материалы до упаковок ОДИН РАЗ
        
        Это ключевой момент экономии:
        - БЫЛО: 10 дверей × ceil(1.2м / 6м) = 10 × 1 = 10 упаковок = 60м
        - СТАЛО: ceil(10 × 1.2м / 6м) = ceil(12м / 6м) = 2 упаковки = 12м
        
        ЭКОНОМИЯ: 5× на профилях!
        """
        for article, qty_raw in self.materials_raw.items():
            pack_size = get_package_size(article, self.ref1)
            
            if pack_size > 0:
                # Округляем вверх до кратного pack_size
                packages_needed = math.ceil(qty_raw / pack_size)
                qty_rounded = packages_needed * pack_size
            else:
                qty_rounded = qty_raw
            
            self.materials_rounded[article] = qty_rounded
    
    def calculate_costs(self) -> Dict[str, float]:
        """
        Рассчитать стоимость всех материалов
        
        Returns:
            {
                "total_materials_cost": float,  # Общая стоимость материалов
                "total_saved": float  # Сколько сэкономили на округлении
            }
        """
        total_cost = 0.0
        total_raw = 0.0
        total_rounded = 0.0
        
        for article, qty_rounded in self.materials_rounded.items():
            info = self.materials_info[article]
            qty_raw = self.materials_raw[article]
            
            cost = qty_rounded * info["price"]
            total_cost += cost
            
            total_raw += qty_raw
            total_rounded += qty_rounded
        
        # Сколько сэкономили (разница между старым и новым подходом)
        # Старый подход: каждая позиция округляла отдельно
        # Новый подход: округление один раз на весь заказ
        total_saved = (total_rounded - total_raw)
        
        return {
            "total_materials_cost": total_cost,
            "total_raw_quantity": total_raw,
            "total_rounded_quantity": total_rounded,
            "total_saved_quantity": total_saved
        }
    
    def get_materials_list(self) -> List[Dict]:
        """
        Получить список материалов для отображения
        
        Returns:
            [
                {
                    "Артикул": str,
                    "Элемент": str,
                    "Количество нетто": float,  # Чистая длина
                    "Количество брутто": float,  # Округлённая длина
                    "Единица": str,
                    "Цена": float,
                    "Сумма": float
                },
                ...
            ]
        """
        result = []
        
        for article, qty_rounded in self.materials_rounded.items():
            info = self.materials_info[article]
            qty_raw = self.materials_raw[article]
            
            result.append({
                "Артикул": article,
                "Элемент": info["name"],
                "Количество нетто": round(qty_raw, 3),
                "Количество брутто": round(qty_rounded, 3),
                "Единица": info["unit"],
                "Цена": info["price"],
                "Сумма": round(qty_rounded * info["price"], 0)
            })
        
        return result
    
    def clear(self):
        """Очистить корзину"""
        self.materials_raw.clear()
        self.materials_rounded.clear()
        self.materials_info.clear()


def extract_materials_from_result(result: Dict, basket: MaterialBasket):
    """
    Извлечь материалы из результата calculate_window_smeta и добавить в корзину
    
    Args:
        result: Результат от calculate_window_smeta (или calculate_facade_materials)
        basket: Глобальная корзина материалов
    """
    # Извлекаем материалы из part2_materials
    materials = result.get("part2_materials", [])
    
    for material in materials:
        article = material.get("Артикул", "")
        # КРИТИЧНО: берём quantity_raw (ДО округления), а не quantity (ПОСЛЕ округления)
        qty_raw = material.get("Количество_raw", material.get("Количество", 0))
        unit = material.get("Единица", "м")
        price = material.get("Цена", 0)
        name = material.get("Элемент", "")
        
        if article and qty_raw > 0:
            basket.add_material(article, qty_raw, unit, price, name)


def calculate_order_with_global_basket(
    positions: List[Dict],
    common_params: Dict,
    ref1: List[Dict],
    ref2: Dict,
    ref3: List[Dict],
    calculate_function
) -> Tuple[List[Dict], MaterialBasket]:
    """
    Рассчитать заказ с использованием глобальной корзины материалов
    
    Args:
        positions: Список позиций заказа
        common_params: Общие параметры (тонировка, сборка, монтаж)
        ref1, ref2, ref3: Справочники
        calculate_function: Функция расчёта (calculate_window_smeta или calculate_facade_materials)
    
    Returns:
        (results_list, basket) - список результатов по позициям и глобальная корзина
    """
    basket = MaterialBasket(ref1)
    results = []
    
    # Шаг 1: Рассчитываем каждую позицию и собираем материалы БЕЗ округления
    for position in positions:
        order_data = {
            "common": common_params,
            "positions": [position]
        }
        
        result = calculate_function(order_data, ref1, ref2, ref3)
        results.append(result)
        
        # Извлекаем материалы и добавляем в корзину (БЕЗ округления!)
        extract_materials_from_result(result, basket)
    
    # Шаг 2: Округляем ВСЕ материалы ОДИН РАЗ
    basket.round_all_materials()
    
    # Шаг 3: Рассчитываем стоимость
    basket.calculate_costs()
    
    return results, basket
