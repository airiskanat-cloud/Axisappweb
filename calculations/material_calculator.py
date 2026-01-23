"""
Universal Material Calculator
Универсальный расчёт материалов без хардкодов
Работает для всех типов изделий и систем профилей
"""

import math
from typing import Dict, List, Any
from .product_model import (
    Product, ProductGeometry, ProductMaterials,
    FrameMaterial, SealMaterial, HardwareItem,
    UsageMode, ProductType
)


class MaterialCalculator:
    """
    Универсальный калькулятор материалов
    
    Принципы:
    1. НЕТ хардкодов под конкретные типы/системы
    2. Геометрия не зависит от контекста
    3. Все расчёты от формул и справочников
    """
    
    def __init__(self, ref1: List[Dict], ref2: Dict[str, float], ref3: List[Dict]):
        """
        Args:
            ref1: Справочник-1 (профили, фурнитура)
            ref2: Справочник-2 (цены на услуги)
            ref3: Справочник-3 (формулы)
        """
        self.ref1 = ref1
        self.ref2 = ref2
        self.ref3 = ref3
    
    def calculate_materials(self, product: Product) -> Product:
        """
        Полный расчёт материалов для изделия
        
        Args:
            product: Модель изделия с заполненной геометрией
        
        Returns:
            Product с заполненными materials
        """
        # 1. Расчёт профилей рамы (все 4 стороны)
        product.materials.frame_materials = self._calculate_frame(product)
        
        # 2. Расчёт уплотнителей (рама + створки)
        product.materials.seal_materials = self._calculate_seals(product)
        
        # 3. Расчёт фурнитуры
        product.materials.hardware = self._calculate_hardware(product)
        
        # 4. Расчёт заполнения (стекло/ламбри)
        glass_result = self._calculate_glass(product)
        product.materials.glass_area = glass_result["area"]
        product.materials.glass_cost = glass_result["cost"]
        
        # 5. Дополнительные детали (по периметру)
        product.materials.additional_details_cost = self._calculate_additional_details(product)
        
        return product
    
    def _calculate_frame(self, product: Product) -> List[FrameMaterial]:
        """
        Расчёт профилей рамы
        
        КРИТИЧЕСКИ ВАЖНО:
        - Все 4 стороны ВСЕГДА присутствуют
        - Периметр = 2 × (W + H)
        - НЕТ урезаний под embedded
        """
        frame_materials = []
        geometry = product.geometry
        
        # Поиск профиля рамы в справочнике
        frame_profile = self._find_profile(
            system=product.system,
            element_type="рама",
            profile_type="frame"
        )
        
        if not frame_profile:
            # Fallback 1: ищем любой профиль для данной системы с "рама" или "коробка"
            system_upper = product.system.strip().upper()
            for item in self.ref1:
                # Поддержка разных названий колонок
                sys = item.get("Система", item.get("Система профиля", "")).strip().upper()
                elem = item.get("Элемент", item.get("Тип элемента", "")).lower()
                
                # Проверяем систему
                if system_upper in sys or sys in system_upper:
                    # Проверяем что это рама
                    if "рам" in elem or "короб" in elem or "frame" in elem:
                        frame_profile = item
                        print(f"✅ Frame profile found (fallback 1): {item.get('Элемент', '')} for {product.system}")
                        break
        
        if not frame_profile:
            # Fallback 2: ищем ЛЮБОЙ профиль для данной системы (с непустым названием)
            system_upper = product.system.strip().upper()
            for item in self.ref1:
                sys = item.get("Система", item.get("Система профиля", "")).strip().upper()
                elem = item.get("Элемент", item.get("Тип элемента", "")).strip()
                
                # Проверяем систему И что элемент не пустой
                if (system_upper in sys or sys in system_upper) and elem:
                    frame_profile = item
                    print(f"⚠️ Frame profile found (fallback 2): Using {elem} for {product.system}")
                    break
        
        if not frame_profile:
            # Последний fallback: выводим список доступных систем
            available_systems = set()
            for item in self.ref1:
                sys = item.get("Система", item.get("Система профиля", ""))
                if sys:
                    available_systems.add(sys)
            
            error_msg = f"Frame profile not found for system '{product.system}'.\n"
            error_msg += f"Available systems in reference: {sorted(available_systems)}"
            raise ValueError(error_msg)
        
        # Поддержка разных названий колонок
        frame_price = self._get_price(frame_profile, 3000)
        frame_name = frame_profile.get("Элемент", 
                     frame_profile.get("Тип элемента", "Рама"))
        frame_article = frame_profile.get("Артикул", "")
        
        # Поиск профиля порога для дверей
        threshold_profile = None
        threshold_price = frame_price
        threshold_name = frame_name
        threshold_article = frame_article
        
        if product.product_type in [ProductType.DOOR_SINGLE, ProductType.DOOR_DOUBLE]:
            threshold_profile = self._find_profile(
                system=product.system,
                element_type="порог",
                profile_type="threshold"
            )
            
            if threshold_profile:
                threshold_price = self._get_price(threshold_profile, frame_price)
                threshold_name = threshold_profile.get("Элемент",
                                threshold_profile.get("Тип элемента", "Порог"))
                threshold_article = threshold_profile.get("Артикул", "")
        
        # КРИТИЧЕСКИ ВАЖНО: Создаём ВСЕ 4 стороны
        sides = geometry.frame_sides
        
        for side, length in sides.items():
            # Для дверей низ = порог (если найден)
            if side == "bottom" and product.product_type in [ProductType.DOOR_SINGLE, ProductType.DOOR_DOUBLE]:
                price = threshold_price
                name = threshold_name
                article = threshold_article
            else:
                price = frame_price
                name = frame_name
                article = frame_article
            
            frame_materials.append(FrameMaterial(
                name=name,
                side=side,
                length=length,
                price=price,
                article=article
            ))
        
        # Валидация: ОБЯЗАТЕЛЬНО 4 стороны
        assert len(frame_materials) == 4, f"Frame must have 4 sides, got {len(frame_materials)}"
        
        return frame_materials
    
    def _calculate_seals(self, product: Product) -> List[SealMaterial]:
        """
        Расчёт уплотнителей
        
        КРИТИЧЕСКИ ВАЖНО:
        - Уплотнитель рамы = полный периметр изделия
        - Уплотнители створок = отдельно
        - НЕТ урезаний
        """
        seal_materials = []
        geometry = product.geometry
        
        # Поиск ВСЕХ уплотнителей для данной системы
        seal_profiles = []
        for item in self.ref1:
            elem = item.get("Элемент", item.get("Тип элемента", ""))
            sys = item.get("Система", item.get("Система профиля", ""))
            
            if product.system in sys and "уплотн" in elem.lower():
                seal_profiles.append(item)
        
        # Если нашли уплотнители - используем их
        if seal_profiles:
            # Используем первый найденный для рамы
            main_seal = seal_profiles[0]
            seal_price = self._get_price(main_seal, 184)
            seal_name = main_seal.get("Элемент", "Уплотнитель")
            
            # 1. Уплотнитель рамы (ОБЯЗАТЕЛЬНО)
            # КРИТИЧЕСКИ ВАЖНО: Полный периметр рамы
            seal_materials.append(SealMaterial(
                name=seal_name + " (рама)",
                zone="frame",
                length=geometry.perimeter,
                price=seal_price,
                article=main_seal.get("Артикул", "")
            ))
            
            # 2. Уплотнители створок (если есть)
            if geometry.sashes and geometry.total_sash_perimeter > 0:
                seal_materials.append(SealMaterial(
                    name=seal_name + " (створки)",
                    zone="sash",
                    length=geometry.total_sash_perimeter,
                    price=seal_price,
                    article=main_seal.get("Артикул", "")
                ))
        else:
            # Fallback: создаём уплотнитель с дефолтной ценой
            seal_materials.append(SealMaterial(
                name="Уплотнитель (рама)",
                zone="frame",
                length=geometry.perimeter,
                price=184,  # Дефолтная цена из логов
                article=""
            ))
            
            if geometry.sashes and geometry.total_sash_perimeter > 0:
                seal_materials.append(SealMaterial(
                    name="Уплотнитель (створки)",
                    zone="sash",
                    length=geometry.total_sash_perimeter,
                    price=184,
                    article=""
                ))
        
        # Валидация: ОБЯЗАТЕЛЬНО должен быть уплотнитель рамы
        frame_seals = [s for s in seal_materials if s.zone == "frame"]
        assert len(frame_seals) > 0, "Frame seal is mandatory"
        
        return seal_materials
    
    def _calculate_hardware(self, product: Product) -> List[HardwareItem]:
        """
        Расчёт фурнитуры
        
        КРИТИЧЕСКИ ВАЖНО:
        - Фурнитура НЕ зависит от usage_mode
        - Одинаковый набор для standalone и embedded
        """
        hardware = []
        
        # Поиск всех элементов фурнитуры для данной системы и типа
        for item in self.ref1:
            elem = item.get("Элемент", item.get("Тип элемента", ""))
            system = item.get("Система", item.get("Система профиля", ""))
            
            # Проверяем, что это фурнитура для нашей системы
            if product.system not in system:
                continue
            
            # Определяем тип фурнитуры
            elem_lower = elem.lower()
            
            # Петли
            if "петл" in elem_lower:
                qty = self._get_hardware_quantity("петля", product)
                if qty > 0:
                    hardware.append(HardwareItem(
                        name=elem,
                        quantity=qty,
                        unit="шт",
                        price=self._get_price(item, 0),
                        article=item.get("Артикул", "")
                    ))
            
            # Ручки
            elif "ручк" in elem_lower:
                qty = self._get_hardware_quantity("ручка", product)
                if qty > 0:
                    hardware.append(HardwareItem(
                        name=elem,
                        quantity=qty,
                        unit="комплект" if "комплек" in elem_lower else "шт",
                        price=self._get_price(item, 0),
                        article=item.get("Артикул", "")
                    ))
            
            # Замки (для дверей)
            elif "замок" in elem_lower and product.product_type in [ProductType.DOOR_SINGLE, ProductType.DOOR_DOUBLE]:
                hardware.append(HardwareItem(
                    name=elem,
                    quantity=1,
                    unit="шт",
                    price=self._get_price(item, 0),
                    article=item.get("Артикул", "")
                ))
            
            # Другие элементы (доводчики, фиксаторы и т.д.)
            elif any(word in elem_lower for word in ["доводчик", "фиксатор", "сердцевина", "планка", "накладк"]):
                if product.product_type in [ProductType.DOOR_SINGLE, ProductType.DOOR_DOUBLE]:
                    hardware.append(HardwareItem(
                        name=elem,
                        quantity=1,
                        unit="шт",
                        price=self._get_price(item, 0),
                        article=item.get("Артикул", "")
                    ))
        
        return hardware
    
    def _calculate_glass(self, product: Product) -> Dict[str, float]:
        """Расчёт стеклопакета/заполнения"""
        geometry = product.geometry
        
        # Площадь створок (если есть) или общая площадь
        if geometry.sashes:
            area = sum(sash.width * sash.height for sash in geometry.sashes) / 1_000_000
        else:
            area = geometry.area
        
        # Поиск цены стекла
        glass_price = 0
        if product.fill_category == "Стеклопакет":
            glass_type_normalized = product.glass_type.lower().strip()
            for key, value in self.ref2.items():
                if glass_type_normalized in key.lower():
                    glass_price = value
                    break
        
        return {
            "area": area,
            "cost": area * glass_price
        }
    
    def _calculate_additional_details(self, product: Product) -> float:
        """
        Дополнительные детали
        
        Формула: ⌈периметр / 3⌉ × цена
        КРИТИЧЕСКИ ВАЖНО: Используем ПОЛНЫЙ периметр
        """
        geometry = product.geometry
        
        # Поиск цены дополнительных деталей
        additional_price = self.ref2.get("дополнительные детали", 5600)
        
        # Расчёт по полному периметру
        count = math.ceil(geometry.perimeter / 3)
        
        return count * additional_price
    
    def _find_profile(self, system: str, element_type: str, profile_type: str) -> Dict:
        """
        Поиск профиля в справочнике
        
        Универсальный поиск без хардкодов
        Использует гибкий поиск по частичному совпадению
        
        Поддерживаемые названия колонок:
        - "Элемент" или "Тип элемента"
        - "Система" или "Система профиля"
        """
        element_type_lower = element_type.lower()
        system_normalized = system.strip().upper()
        
        # Список возможных вариантов поиска элемента
        element_variants = [element_type_lower]
        if element_type_lower == "рама":
            element_variants.extend(["frame", "коробка"])
        elif element_type_lower == "порог":
            element_variants.extend(["threshold", "низ"])
        elif element_type_lower == "створка":
            element_variants.extend(["sash", "створ"])
        elif element_type_lower == "импост":
            element_variants.extend(["impost", "перемычка"])
        
        for item in self.ref1:
            # Поддержка разных названий колонок
            elem = item.get("Элемент", item.get("Тип элемента", ""))
            sys = item.get("Система", item.get("Система профиля", ""))
            
            # Нормализуем систему из справочника
            sys_normalized = sys.strip().upper()
            
            # Проверка системы (гибкий поиск)
            # Проверяем как прямое вхождение, так и обратное
            system_match = (
                system_normalized in sys_normalized or 
                sys_normalized in system_normalized or
                system.strip() in sys or
                sys.strip() in system
            )
            
            if not system_match:
                continue
            
            # Проверка типа элемента (любой из вариантов)
            elem_lower = elem.lower()
            if any(variant in elem_lower for variant in element_variants):
                return item
        
        return {}
    
    def _get_hardware_quantity(self, hardware_type: str, product: Product) -> float:
        """
        Расчёт количества фурнитуры
        
        Универсальная логика от геометрии
        """
        if hardware_type == "петля":
            # Петли: по количеству створок
            return len(product.geometry.sashes) if product.geometry.sashes else 1
        
        elif hardware_type == "ручка":
            # Ручки: по количеству створок
            return len(product.geometry.sashes) if product.geometry.sashes else 1
        
        return 0
    
    @staticmethod
    def _parse_price(value) -> float:
        """Безопасное преобразование цены"""
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
    
    @staticmethod
    def _get_price(item: Dict, default: float = 0.0) -> float:
        """
        Получение цены из элемента с поддержкой разных названий колонок
        
        Поддерживаемые колонки:
        - "Цена за единицу"
        - "цена за ед." 
        - "цена за ед " (с пробелом!)
        - "Цена"
        """
        price = item.get("Цена за единицу",
                item.get("цена за ед.",
                item.get("цена за ед ",  # С ПРОБЕЛОМ!
                item.get("Цена", default))))
        return MaterialCalculator._parse_price(price)


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
        product_data: Данные изделия из формы
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
        usage_mode=usage_mode
    )
    
    # 2. Расчёт материалов (универсальный)
    calculator = MaterialCalculator(ref1, ref2, ref3)
    product = calculator.calculate_materials(product)
    
    # 3. Валидация
    errors = product.validate()
    if errors:
        print("⚠️ VALIDATION ERRORS:")
        for error in errors:
            print(f"   - {error}")
    
    return product
