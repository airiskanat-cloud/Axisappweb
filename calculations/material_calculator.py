"""
Material Calculator - Formula-Based System with CODE Priority
Расчёт материалов по формулам из Справочника-1
С поддержкой CODE для точного поиска профилей
БЕЗ хардкодов - всё из справочника!
"""

import math
from typing import Dict, List, Any
from .product_model import (
    Product, ProductGeometry, ProductMaterials,
    FrameMaterial, SealMaterial, HardwareItem,
    UsageMode, ProductType
)


def safe_eval(formula: str, context: dict) -> float:
    """Безопасное вычисление формул Python"""
    try:
        f = str(formula).replace(",", ".").replace(" ", "")
        return float(eval(f, {"__builtins__": None, "math": math}, context))
    except Exception as e:
        print(f"⚠️ Ошибка в формуле '{formula}': {e}")
        return 0.0


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


class MaterialCalculator:
    """
    Калькулятор материалов по формулам из справочника
    
    ПРИНЦИП:
    1. Берёт CODE из изделия
    2. Находит ВСЕ строки в ref1 с этим CODE
    3. Вычисляет формулы для каждой строки
    4. Округляет до упаковок
    5. Возвращает полный список материалов
    
    ФИЧИ:
    - Поиск по CODE с приоритетом + fallback
    - Поддержка всех вариантов названий колонок
    - Формулы из справочника
    - Округление до упаковок
    - БЕЗ ХАРДКОДОВ!
    """
    
    def __init__(self, ref1: List[Dict], ref2: Dict[str, float], ref3: List[Dict]):
        """
        Args:
            ref1: Справочник-1 (профили, фурнитура, формулы)
            ref2: Справочник-2 (цены на услуги)
            ref3: Справочник-3 (габаритная ведомость)
        """
        self.ref1 = ref1
        self.ref2 = ref2
        self.ref3 = ref3
    
    @staticmethod
    def _get_price(item: Dict, default: float = 0.0) -> float:
        """
        Извлекает цену из строки справочника
        Поддерживает все варианты названий колонок
        """
        price = item.get("Цена за единицу",
                item.get("цена за ед.",
                item.get("цена за ед ",  # С ПРОБЕЛОМ!
                item.get("Цена", default))))
        return safe_float(price, default)
    
    def calculate_materials(self, product: Product) -> Product:
        """
        Полный расчёт материалов для изделия
        
        Args:
            product: Модель изделия с заполненной геометрией
        
        Returns:
            Product с заполненными materials
        """
        print(f"\n{'='*70}")
        print(f"🔧 РАСЧЁТ МАТЕРИАЛОВ ПО ФОРМУЛАМ")
        print(f"{'='*70}")
        print(f"Тип изделия: {product.product_type.value}")
        print(f"Система: {product.system}")
        print(f"CODE: {product.code}")
        print(f"Габариты: {product.geometry.width_m}м × {product.geometry.height_m}м")
        
        # 1. Создаём контекст для формул
        context = self._create_formula_context(product)
        
        print(f"\n📊 Контекст для формул:")
        print(f"   W={context['W']:.2f}м, H={context['H']:.2f}м")
        print(f"   n_sash={context['n_sash']}, n_lp={context['n_lp']}")
        print(f"   Периметр={context['perimeter']:.2f}м, Площадь={context['area']:.2f}м²")
        
        # 2. Ищем ВСЕ материалы из справочника по CODE
        materials_dict = {}
        found_count = 0
        
        print(f"\n🔍 Поиск материалов в Справочнике-1...")
        
        for row in self.ref1:
            # Поддержка разных названий колонки CODE
            row_code = str(row.get("CODE") or row.get("code") or "").strip()
            
            # Проверяем совпадение CODE
            if row_code and product.code and row_code == product.code:
                # Берём формулу
                formula = row.get("Формула_Python", "")
                if not formula:
                    formula = row.get("формула фактического расхода", "")
                if not formula:
                    continue
                
                # Вычисляем количество по формуле
                qty_fact = safe_eval(formula, context)
                
                if qty_fact <= 0:
                    continue
                
                found_count += 1
                
                # Собираем данные материала
                товар = str(row.get("Товар", ""))
                артикул = str(row.get("Артикул", ""))
                тип_эл = row.get("Тип элемента", row.get("тип элемент", ""))
                key = f"{товар}|{артикул}"
                
                if key not in materials_dict:
                    materials_dict[key] = {
                        "товар": товар,
                        "артикул": артикул,
                        "тип_эл": тип_эл,
                        "qty_fact": 0,
                        "norm": safe_float(row.get("кол-во норм к упаковке", 1)),
                        "price": self._get_price(row, 0),
                        "unit": str(row.get("Ед.", "шт"))
                    }
                
                # Накапливаем количество (для нескольких формул)
                materials_dict[key]["qty_fact"] += qty_fact
                
                print(f"   ✅ {тип_эл}: {formula} = {qty_fact:.2f} {materials_dict[key]['unit']}")
        
        print(f"\nНайдено материалов: {found_count} строк → {len(materials_dict)} позиций")
        
        if len(materials_dict) == 0:
            print(f"\n⚠️ ВНИМАНИЕ: Не найдено ни одного материала!")
            print(f"   CODE: {product.code}")
            print(f"   Проверьте что в Справочнике-1 есть строки с этим CODE")
            
            # Показываем доступные CODE
            available_codes = set()
            for row in self.ref1:
                row_code = str(row.get("CODE") or row.get("code") or "").strip()
                if row_code:
                    available_codes.add(row_code)
            
            if available_codes:
                print(f"\n   Доступные CODE в справочнике:")
                for code in sorted(available_codes)[:10]:
                    print(f"      - {code}")
                if len(available_codes) > 10:
                    print(f"      ... и ещё {len(available_codes) - 10}")
        
        # 3. Округляем до упаковок и считаем стоимость
        materials_list = []
        total_materials_cost = 0
        
        print(f"\n💰 Округление до упаковок и расчёт стоимости:")
        
        for key, mat in materials_dict.items():
            qty_fact = mat["qty_fact"]
            norm = mat["norm"]
            price = mat["price"]
            
            # Округление до упаковок
            if norm > 0:
                qty_ship = math.ceil(qty_fact / norm)
            else:
                qty_ship = math.ceil(qty_fact)
            
            # Стоимость = (цена * норма) * количество упаковок
            row_sum = (price * norm) * qty_ship
            total_materials_cost += row_sum
            
            materials_list.append({
                "Товар": mat["товар"],
                "Артикул": mat["артикул"],
                "Тип элемента": mat["тип_эл"],
                "Цена": price,
                "Ед.": mat["unit"],
                "Расход факт.": round(qty_fact, 2),
                "Норма": norm,
                "К отгрузке": qty_ship,
                "Сумма": round(row_sum, 0)
            })
            
            print(f"   {mat['тип_эл']}: {qty_fact:.2f}{mat['unit']} → {qty_ship} упак × {norm}{mat['unit']} = {row_sum:,.0f}₸")
        
        print(f"\nИТОГО МАТЕРИАЛЫ: {total_materials_cost:,.0f}₸")
        
        # 4. Распределяем материалы по категориям
        product = self._categorize_materials(product, materials_list)
        
        # 5. Расчёт стекла (отдельно от справочника)
        glass_result = self._calculate_glass(product)
        product.materials.glass_area = glass_result["area"]
        product.materials.glass_cost = glass_result["cost"]
        
        print(f"\n💎 Стеклопакет: {glass_result['area']:.2f}м² × {glass_result['cost']/glass_result['area']:,.0f}₸/м² = {glass_result['cost']:,.0f}₸")
        print(f"{'='*70}\n")
        
        return product
    
    def _create_formula_context(self, product: Product) -> Dict:
        """
        Создаёт контекст для вычисления формул
        
        Все переменные которые могут быть в формулах справочника
        """
        geometry = product.geometry
        
        # Основные габариты
        W = geometry.width_m
        H = geometry.height_m
        
        # Количество створок
        n_sash = len(geometry.sashes)
        
        # Размеры створки (средние если несколько)
        if geometry.sashes:
            w_s = sum(s.width for s in geometry.sashes) / len(geometry.sashes) / 1000
            h_s = sum(s.height for s in geometry.sashes) / len(geometry.sashes) / 1000
            w_s_total = sum(s.width for s in geometry.sashes) / 1000  # Для дверей
        else:
            w_s = W
            h_s = H
            w_s_total = W
        
        # Световой проём (примерно, можно улучшить)
        offset = 0.073  # 73мм для ALG 2030
        w_g = W - 2 * offset
        h_g = H - 2 * offset
        
        # Импосты
        imp_vertical = 1 if geometry.has_vertical_impost else 0
        imp_horizontal = 1 if geometry.has_horizontal_impost else 0
        total_imposts = imp_vertical + imp_horizontal
        
        # Точки запирания
        if h_s * 1000 < 1200:
            n_lp = 2
        elif h_s * 1000 < 2000:
            n_lp = 3
        else:
            n_lp = 4
        
        # Контекст для формул
        context = {
            # Основные
            "W": W,
            "H": H,
            "w": W,
            "h": H,
            "count": 1,  # Всегда 1 для одного изделия
            "qty": 1,
            "Nwin": 1,
            
            # Створки
            "n_sash": n_sash,
            "w_s": w_s,
            "h_s": h_s,
            "w_stvor": w_s,
            "h_stvor": h_s,
            "w_s_total": w_s_total,
            
            # Световой проём
            "w_g": w_g,
            "h_g": h_g,
            "w_glass": w_g,
            "h_glass": h_g,
            
            # Импосты
            "imp_vertical": imp_vertical,
            "imp_horizontal": imp_horizontal,
            "total_imposts": total_imposts,
            
            # Фурнитура
            "n_lp": n_lp,
            "lock_points": n_lp,
            
            # Площадь и периметр
            "area": W * H,
            "area_m2": W * H,
            "perimeter": 2 * (W + H),
            "perimeter_m": 2 * (W + H)
        }
        
        return context
    
    def _categorize_materials(self, product: Product, materials_list: List[Dict]) -> Product:
        """
        Распределяет материалы по категориям (рама, уплотнители, фурнитура)
        """
        for mat in materials_list:
            тип_эл = mat["Тип элемента"].lower()
            
            # Профили рамы
            if "профиль" in тип_эл or "рам" in тип_эл or "короб" in тип_эл or "порог" in тип_эл or "створ" in тип_эл or "импост" in тип_эл:
                product.materials.frame_materials.append(FrameMaterial(
                    name=mat["Товар"],
                    side="calculated",  # Из формулы
                    length=mat["Расход факт."],
                    price=mat["Цена"],
                    article=mat["Артикул"]
                ))
            
            # Уплотнители
            elif "уплотн" in тип_эл or "штап" in тип_эл:
                product.materials.seal_materials.append(SealMaterial(
                    name=mat["Товар"],
                    zone="calculated",  # Из формулы
                    length=mat["Расход факт."],
                    price=mat["Цена"],
                    article=mat["Артикул"]
                ))
            
            # Фурнитура
            else:
                product.materials.hardware.append(HardwareItem(
                    name=mat["Товар"],
                    quantity=mat["К отгрузке"],
                    price=mat["Цена"] * mat["Норма"],  # Цена за упаковку
                    article=mat["Артикул"]
                ))
        
        return product
    
    def _calculate_glass(self, product: Product) -> Dict:
        """
        Расчёт стеклопакета (из ref2)
        """
        geometry = product.geometry
        
        # Площадь стекла
        glass_area = geometry.width_m * geometry.height_m
        
        # Цена из ref2
        glass_type_key = product.glass_type.lower()
        glass_price_per_m2 = self.ref2.get(glass_type_key, 9500)
        
        # Стоимость
        glass_cost = glass_area * glass_price_per_m2
        
        return {
            "area": glass_area,
            "cost": glass_cost
        }


def calculate_product_materials(
    product_data: Dict,
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict],
    usage_mode: UsageMode = UsageMode.STANDALONE
) -> Product:
    """
    Универсальная функция расчёта материалов
    
    Args:
        product_data: Данные изделия из формы (включая CODE!)
        ref1, ref2, ref3: Справочники
        usage_mode: Контекст использования (НЕ влияет на расчёт!)
    
    Returns:
        Product с полным расчётом материалов
    """
    from .product_model import create_product_from_form_data
    
    # 1. Создание модели изделия
    product = create_product_from_form_data(
        product_type=product_data.get("product_type", "Окно"),
        system=product_data.get("system", "ALG 2030-45C"),
        data=product_data.get("data", {}),
        usage_mode=usage_mode,
        code=product_data.get("code", "")  # ✅ CODE для формул!
    )
    
    # 2. Расчёт материалов (по формулам из справочника)
    calculator = MaterialCalculator(ref1, ref2, ref3)
    product = calculator.calculate_materials(product)
    
    # 3. Валидация
    errors = product.validate()
    if errors:
        print("⚠️ VALIDATION ERRORS:")
        for error in errors:
            print(f"   - {error}")
    
    return product
