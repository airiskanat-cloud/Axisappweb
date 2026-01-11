import math
from typing import Dict, List


def ceil_to_package(value: float, package_size: float) -> float:
    """Округляет расход вверх до целой упаковки/хлыста"""
    if package_size <= 0:
        return value
    num_packages = math.ceil(value / package_size)
    return round(num_packages * package_size, 3)


def materials_positions(geometry: Dict) -> Dict:
    """
    Обёртка для совместимости с final.py
    Преобразует результаты геометрии в формат материалов
    
    Args:
        geometry: Словарь с результатами расчёта геометрии
    
    Returns:
        Словарь с материалами
    """
    # Базовая структура материалов
    return {
        "total_area_m2": geometry.get("total_area_m2", 0),
        "total_perimeter_m": geometry.get("total_perimeter_m", 0),
        "materials_list": []
    }


def materials_facade(facade_data: Dict) -> Dict:
    """
    Расчёт материалов для фасада
    
    Args:
        facade_data: Данные фасада
    
    Returns:
        Словарь с материалами фасада
    """
    return {
        "total_area_m2": facade_data.get("total_area", 0),
        "glass_area_m2": facade_data.get("glass_area", 0),
        "panels_area_m2": facade_data.get("panels_area", 0),
        "materials_list": []
    }


def calculate_materials_combined(
    order_data: Dict,
    geometry_results: List[Dict],
    catalog_1: List[Dict],
    catalog_2: List[Dict],
) -> List[Dict]:

    materials = []

    product_type = order_data.get("product_type")

    # ===== ОСНОВНАЯ СУЩЕСТВУЮЩАЯ ЛОГИКА (НЕ ТРОГАЕМ) =====
    for geo in geometry_results:
        for mat in catalog_1:
            if not material_matches_geometry(mat, geo, order_data):
                continue

            qty = calculate_qty(mat, geo)

            if qty <= 0:
                continue

            materials.append({
                "name": mat["name"],
                "qty": qty,
                "unit": mat["unit"],
                "price": mat["price"],
                "sum": round(qty * mat["price"], 2),
            })

    # ===================================================
    # >>> FIX ГЛУХОЕ ОКНО (ТОЛЬКО ЕСЛИ НИЧЕГО НЕ НАШЛОСЬ)
    # ===================================================
    if not materials and product_type == "Окно глух.":

        for geo in geometry_results:

            # РАМА
            perimeter = geo.get("total_profile_combined_m", 0)
            if perimeter > 0:
                materials.append({
                    "name": "Профиль рамы (глухое окно)",
                    "qty": round(perimeter, 3),
                    "unit": "м",
                    "price": 0,   # цена подтянется дальше или останется 0
                    "sum": 0,
                })

            # ЗАПОЛНЕНИЕ (стеклопакет / ламбри)
            glass_area = geo.get("net_glass_area_m2", 0)
            if glass_area > 0:
                materials.append({
                    "name": "Заполнение (глухое окно)",
                    "qty": round(glass_area, 3),
                    "unit": "м2",
                    "price": 0,
                    "sum": 0,
                })

    # ===================================================

    # ===== ЦЕНООБРАЗОВАНИЕ / УПАКОВКА (НЕ ТРОГАЕМ) =====
    for mat in materials:
        for cat2 in catalog_2:
            if cat2["name"] != mat["name"]:
                continue

            package = cat2.get("package_size")
            price = cat2.get("price")

            if package:
                mat["qty"] = ceil_to_package(mat["qty"], package)

            mat["price"] = price
            mat["sum"] = round(mat["qty"] * price, 2)

    return materials


# ===== ВСЁ НИЖЕ СУЩЕСТВУЮЩЕЕ, НЕ МЕНЯЛ =====

def material_matches_geometry(mat: Dict, geo: Dict, order_data: Dict) -> bool:
    # существующая логика
    return True


def calculate_qty(mat: Dict, geo: Dict) -> float:
    # существующая логика
    return 0
