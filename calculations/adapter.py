"""
Adapter для интеграции нового кода с существующим engine_windows
Обеспечивает обратную совместимость
"""

from typing import Dict, List, Any
from .product_model import Product, UsageMode
from .material_calculator import calculate_product_materials


class ProductToLegacyAdapter:
    """
    Адаптер для преобразования Product в legacy формат
    Используется для сохранения совместимости с существующим кодом
    """
    
    @staticmethod
    def to_legacy_materials(product: Product) -> List[Dict]:
        """
        Преобразование materials в legacy формат (part2_materials)
        
        Returns:
            List[Dict] в формате для таблицы материалов
        """
        materials = []
        
        # 1. Профили рамы
        for frame_mat in product.materials.frame_materials:
            materials.append({
                "Товар": frame_mat.name,
                "Ед.": "м",
                "К отгрузке": round(frame_mat.length, 2),
                "Цена": frame_mat.price,
                "Сумма": round(frame_mat.cost, 2),
                "Артикул": frame_mat.article,
                "category": "profile"
            })
        
        # 2. Уплотнители
        for seal_mat in product.materials.seal_materials:
            materials.append({
                "Товар": seal_mat.name,
                "Ед.": "м",
                "К отгрузке": round(seal_mat.length, 2),
                "Цена": seal_mat.price,
                "Сумма": round(seal_mat.cost, 2),
                "Артикул": seal_mat.article,
                "category": "seal"
            })
        
        # 3. Фурнитура
        for hardware in product.materials.hardware:
            materials.append({
                "Товар": hardware.name,
                "Ед.": hardware.unit,
                "К отгрузке": hardware.quantity,
                "Цена": hardware.price,
                "Сумма": round(hardware.cost, 2),
                "Артикул": hardware.article,
                "category": "hardware"
            })
        
        return materials
    
    @staticmethod
    def to_legacy_summary(product: Product) -> Dict:
        """
        Преобразование в legacy формат (part3_final)
        
        Returns:
            Dict с итогами по категориям
        """
        return {
            "Стеклопакет": product.materials.glass_cost,
            "Ламбри": 0.0,  # TODO: добавить поддержку ламбри
            "Тонировка": 0.0,  # Считается отдельно
            "Сборка": 0.0,    # Считается отдельно
            "Монтаж": 0.0,    # Считается отдельно
            "Дополнительные детали": product.materials.additional_details_cost,
            "Материалы": (
                product.materials.total_frame_cost +
                product.materials.total_seal_cost +
                product.materials.total_hardware_cost
            ),
            "Обеспечение": 0.0  # TODO: добавить расчёт обеспечения
        }


def calculate_window_smeta_unified(
    order_data: Dict,
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict]
) -> Dict:
    """
    Унифицированный расчёт с использованием новой модели
    Заменяет старый calculate_window_smeta
    
    Args:
        order_data: Данные заказа (common + positions)
        ref1, ref2, ref3: Справочники
    
    Returns:
        Dict в legacy формате для совместимости
    """
    result = {
        "part2_materials": [],
        "part3_final": {},
        "metrics": {
            "total_area": 0.0,
            "total_perimeter": 0.0
        }
    }
    
    # Обработка каждой позиции
    for position in order_data.get("positions", []):
        # Определяем usage_mode
        # Если в данных есть флаг embedded - используем его
        data = position.get("data", {})
        usage_mode = UsageMode.EMBEDDED if data.get("embedded", False) else UsageMode.STANDALONE
        
        # Расчёт материалов через новую модель
        product = calculate_product_materials(
            product_data={
                "product_type": position.get("product_type", "Окно"),
                "system": position.get("system", "ALG 2030-45C"),
                "data": data
            },
            ref1=ref1,
            ref2=ref2,
            ref3=ref3,
            usage_mode=usage_mode
        )
        
        # Валидация
        errors = product.validate()
        if errors:
            print(f"\n⚠️ VALIDATION ERRORS for {product.product_type.value}:")
            for error in errors:
                print(f"   - {error}")
        
        # Преобразование в legacy формат
        adapter = ProductToLegacyAdapter()
        
        # Добавляем материалы
        result["part2_materials"].extend(adapter.to_legacy_materials(product))
        
        # Добавляем итоги
        summary = adapter.to_legacy_summary(product)
        if not result["part3_final"]:
            result["part3_final"] = summary
        else:
            # Суммируем с предыдущими позициями
            for key in summary:
                result["part3_final"][key] = result["part3_final"].get(key, 0) + summary[key]
        
        # Метрики
        result["metrics"]["total_area"] += product.geometry.area
        result["metrics"]["total_perimeter"] += product.geometry.perimeter
        
        # Диагностическая информация
        print(f"\n✅ Product calculated: {product.product_type.value}")
        print(f"   System: {product.system}")
        print(f"   Dimensions: {product.geometry.width}×{product.geometry.height}мм")
        print(f"   Perimeter: {product.geometry.perimeter:.2f}м")
        print(f"   Usage mode: {product.usage_mode.value}")
        print(f"   Frame materials: {len(product.materials.frame_materials)} items")
        print(f"   Seal materials: {len(product.materials.seal_materials)} items")
        print(f"   Hardware: {len(product.materials.hardware)} items")
        print(f"   Total cost: {product.materials.total_cost:,.0f}₸")
    
    # Добавляем общие услуги (тонировка, сборка, монтаж)
    common = order_data.get("common", {})
    
    # Тонировка
    if common.get("toning", "Нет") != "Нет":
        toning_price = ref2.get("тонировка", 2000)
        result["part3_final"]["Тонировка"] = result["metrics"]["total_area"] * toning_price
    
    # Сборка
    if common.get("assembly", "Нет") != "Нет":
        assembly_price = ref2.get("сборка", 10000)
        result["part3_final"]["Сборка"] = result["metrics"]["total_area"] * assembly_price
    
    # Монтаж
    if common.get("installation", "Нет") != "Нет":
        installation_price = ref2.get("монтаж", 15000)
        result["part3_final"]["Монтаж"] = result["metrics"]["total_area"] * installation_price
    
    return result


def validate_calculation_consistency(
    standalone_result: Dict,
    embedded_result: Dict,
    tolerance: float = 0.20
) -> Dict[str, Any]:
    """
    Проверка консистентности расчётов standalone vs embedded
    
    Args:
        standalone_result: Результат расчёта для отдельного изделия
        embedded_result: Результат расчёта для вставки
        tolerance: Допустимое отклонение (по умолчанию 20%)
    
    Returns:
        Dict с результатами валидации
    """
    validation = {
        "passed": True,
        "errors": [],
        "warnings": [],
        "comparison": {}
    }
    
    # Сравнение материалов
    standalone_materials = standalone_result.get("part3_final", {}).get("Материалы", 0)
    embedded_materials = embedded_result.get("part3_final", {}).get("Материалы", 0)
    
    if standalone_materials > 0:
        diff_percent = abs(standalone_materials - embedded_materials) / standalone_materials
        
        validation["comparison"]["materials"] = {
            "standalone": standalone_materials,
            "embedded": embedded_materials,
            "diff_percent": diff_percent * 100,
            "diff_abs": abs(standalone_materials - embedded_materials)
        }
        
        if diff_percent > tolerance:
            validation["passed"] = False
            validation["errors"].append(
                f"Materials cost difference > {tolerance*100}%: "
                f"{diff_percent*100:.1f}% ({standalone_materials:,.0f} vs {embedded_materials:,.0f})"
            )
    
    # Сравнение периметров
    standalone_perimeter = standalone_result.get("metrics", {}).get("total_perimeter", 0)
    embedded_perimeter = embedded_result.get("metrics", {}).get("total_perimeter", 0)
    
    if standalone_perimeter > 0 and abs(standalone_perimeter - embedded_perimeter) > 0.1:
        validation["warnings"].append(
            f"Perimeter mismatch: {standalone_perimeter:.2f}м vs {embedded_perimeter:.2f}м"
        )
    
    # Проверка наличия всех сторон рамы
    for result, mode in [(standalone_result, "standalone"), (embedded_result, "embedded")]:
        frame_sides = set()
        for mat in result.get("part2_materials", []):
            if mat.get("category") == "profile" and "рама" in mat.get("Товар", "").lower():
                # Проверяем длину профиля
                length = mat.get("К отгрузке", 0)
                if length > 0:
                    frame_sides.add(mat.get("Товар"))
        
        if len(frame_sides) < 1:  # Должен быть хотя бы один профиль рамы
            validation["warnings"].append(
                f"Frame profiles missing in {mode} calculation"
            )
    
    return validation
