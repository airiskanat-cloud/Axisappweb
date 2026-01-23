"""
Unified Product Model - единая модель изделия для окон и дверей
Геометрия изделия не зависит от контекста использования (standalone/embedded)
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Literal
from enum import Enum


class UsageMode(Enum):
    """Контекст использования изделия"""
    STANDALONE = "standalone"  # Отдельное изделие
    EMBEDDED = "embedded"      # Вставка в фасад


class ProductType(Enum):
    """Тип изделия"""
    WINDOW = "window"
    DOOR_SINGLE = "door_single"
    DOOR_DOUBLE = "door_double"


@dataclass
class Sash:
    """Створка"""
    width: float   # мм
    height: float  # мм
    opening_type: str = "Откр."
    
    @property
    def perimeter(self) -> float:
        """Периметр створки в метрах"""
        return 2 * (self.width + self.height) / 1000


@dataclass
class ProductGeometry:
    """
    Геометрия изделия - ВСЕГДА ПОЛНАЯ
    Не зависит от контекста использования
    """
    width: float   # мм
    height: float  # мм
    sashes: List[Sash] = field(default_factory=list)
    
    # Импосты (если есть)
    has_horizontal_impost: bool = False
    has_vertical_impost: bool = False
    horizontal_impost_count: int = 0
    vertical_impost_count: int = 0
    
    @property
    def width_m(self) -> float:
        """Ширина в метрах"""
        return self.width / 1000
    
    @property
    def height_m(self) -> float:
        """Высота в метрах"""
        return self.height / 1000
    
    @property
    def perimeter(self) -> float:
        """
        Полный периметр изделия в метрах
        ВСЕГДА: 2 × (W + H)
        """
        return 2 * (self.width_m + self.height_m)
    
    @property
    def area(self) -> float:
        """Площадь изделия в м²"""
        return self.width_m * self.height_m
    
    @property
    def frame_sides(self) -> Dict[str, float]:
        """
        Все 4 стороны рамы (в метрах)
        ВСЕГДА присутствуют, независимо от контекста
        """
        return {
            "left": self.height_m,
            "right": self.height_m,
            "top": self.width_m,
            "bottom": self.width_m
        }
    
    @property
    def total_sash_perimeter(self) -> float:
        """Суммарный периметр всех створок в метрах"""
        return sum(sash.perimeter for sash in self.sashes)
    
    def get_impost_length(self) -> float:
        """Длина импостов в метрах"""
        h_length = self.horizontal_impost_count * self.width_m
        v_length = self.vertical_impost_count * self.height_m
        return h_length + v_length


@dataclass
class FrameMaterial:
    """Материал рамы"""
    name: str
    side: str  # left, right, top, bottom
    length: float  # метры
    price: float
    article: str = ""
    
    @property
    def cost(self) -> float:
        return self.length * self.price


@dataclass
class SealMaterial:
    """Уплотнитель"""
    name: str
    zone: str  # frame, sash, glazing_bead
    length: float  # метры
    price: float
    article: str = ""
    
    @property
    def cost(self) -> float:
        return self.length * self.price


@dataclass
class HardwareItem:
    """Элемент фурнитуры"""
    name: str
    quantity: float
    unit: str
    price: float
    article: str = ""
    
    @property
    def cost(self) -> float:
        return self.quantity * self.price


@dataclass
class ProductMaterials:
    """
    Материалы изделия
    Полный набор, независимо от контекста
    """
    # Профили рамы (все 4 стороны)
    frame_materials: List[FrameMaterial] = field(default_factory=list)
    
    # Уплотнители (рама + створки)
    seal_materials: List[SealMaterial] = field(default_factory=list)
    
    # Фурнитура
    hardware: List[HardwareItem] = field(default_factory=list)
    
    # Заполнение (стекло/ламбри)
    glass_area: float = 0.0
    glass_cost: float = 0.0
    
    # Дополнительные детали
    additional_details_cost: float = 0.0
    
    @property
    def total_frame_cost(self) -> float:
        """Стоимость профилей рамы"""
        return sum(m.cost for m in self.frame_materials)
    
    @property
    def total_seal_cost(self) -> float:
        """Стоимость уплотнителей"""
        return sum(m.cost for m in self.seal_materials)
    
    @property
    def total_hardware_cost(self) -> float:
        """Стоимость фурнитуры"""
        return sum(h.cost for h in self.hardware)
    
    @property
    def total_cost(self) -> float:
        """Полная стоимость материалов"""
        return (
            self.total_frame_cost +
            self.total_seal_cost +
            self.total_hardware_cost +
            self.glass_cost +
            self.additional_details_cost
        )


@dataclass
class Product:
    """
    Единая модель изделия (окно/дверь)
    
    Принципы:
    1. Геометрия ВСЕГДА полная
    2. Материалы ВСЕГДА полные
    3. Контекст влияет только на распределение (owner)
    """
    product_type: ProductType
    system: str
    geometry: ProductGeometry
    materials: ProductMaterials
    
    # Контекст использования (НЕ влияет на расчёт)
    usage_mode: UsageMode = UsageMode.STANDALONE
    
    # Данные заполнения
    fill_category: str = "Стеклопакет"
    glass_type: str = "Двойной"
    
    # Дополнительные опции
    toning: str = "Нет"
    assembly: str = "Нет"
    installation: str = "Нет"
    
    def validate(self) -> List[str]:
        """
        Валидация модели изделия
        Проверяет, что все обязательные элементы присутствуют
        """
        errors = []
        
        # Проверка геометрии
        if self.geometry.width <= 0 or self.geometry.height <= 0:
            errors.append("Invalid dimensions: width and height must be positive")
        
        # Проверка рамы (должны быть все 4 стороны)
        if len(self.materials.frame_materials) == 0:
            errors.append("Frame materials missing: all 4 sides required")
        
        frame_sides = {m.side for m in self.materials.frame_materials}
        required_sides = {"left", "right", "top", "bottom"}
        missing_sides = required_sides - frame_sides
        if missing_sides:
            errors.append(f"Missing frame sides: {missing_sides}")
        
        # Проверка уплотнителей (должны быть рама + створки)
        if len(self.materials.seal_materials) == 0:
            errors.append("Seal materials missing")
        
        seal_zones = {m.zone for m in self.materials.seal_materials}
        if "frame" not in seal_zones:
            errors.append("Frame seal missing")
        
        return errors
    
    def get_summary(self) -> Dict:
        """Краткая сводка по изделию"""
        return {
            "type": self.product_type.value,
            "system": self.system,
            "dimensions": f"{self.geometry.width}×{self.geometry.height}мм",
            "perimeter": f"{self.geometry.perimeter:.2f}м",
            "area": f"{self.geometry.area:.2f}м²",
            "usage_mode": self.usage_mode.value,
            "total_cost": f"{self.materials.total_cost:,.0f}₸",
            "frame_cost": f"{self.materials.total_frame_cost:,.0f}₸",
            "seal_cost": f"{self.materials.total_seal_cost:,.0f}₸",
            "hardware_cost": f"{self.materials.total_hardware_cost:,.0f}₸"
        }


def create_product_from_form_data(
    product_type: str,
    system: str,
    data: Dict,
    usage_mode: UsageMode = UsageMode.STANDALONE
) -> Product:
    """
    Создание модели изделия из данных формы
    
    Args:
        product_type: "Окно", "Дверь 1-створч.", "Дверь 2-х створч."
        system: "ALG 2030-45C", etc.
        data: Данные из формы (width, height, sashes, etc.)
        usage_mode: Контекст использования
    """
    # Маппинг типов
    type_mapping = {
        "Окно": ProductType.WINDOW,
        "Дверь 1-створч.": ProductType.DOOR_SINGLE,
        "Дверь 2-х створч.": ProductType.DOOR_DOUBLE,
    }
    
    # Создание створок
    sashes = []
    for sash_data in data.get("sashes", []):
        sashes.append(Sash(
            width=sash_data.get("w", 0),
            height=sash_data.get("h", 0),
            opening_type=sash_data.get("opening_type", "Откр.")
        ))
    
    # Создание геометрии
    imposts = data.get("imposts", {})
    geometry = ProductGeometry(
        width=data.get("width", 0),
        height=data.get("height", 0),
        sashes=sashes,
        has_horizontal_impost=imposts.get("has_tor", False),
        has_vertical_impost=imposts.get("has_center", False),
        horizontal_impost_count=1 if imposts.get("has_tor") else 0,
        vertical_impost_count=1 if imposts.get("has_center") else 0
    )
    
    # Создание изделия (материалы будут заполнены позже)
    product = Product(
        product_type=type_mapping.get(product_type, ProductType.WINDOW),
        system=system,
        geometry=geometry,
        materials=ProductMaterials(),
        usage_mode=usage_mode,
        fill_category=data.get("fill_category", "Стеклопакет"),
        glass_type=data.get("glass_type", "Двойной"),
        toning=data.get("toning", "Нет"),
        assembly=data.get("assembly", "Нет"),
        installation=data.get("installation", "Нет")
    )
    
    return product
