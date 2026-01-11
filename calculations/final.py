from calculations.engine_windows import calculate_windows_geometry
from calculations.materials import materials_positions
from calculations.pricing import price_windows_doors, price_options_windows

from calculations.facade import facade_glass_and_panels
from calculations.materials import materials_facade
from calculations.pricing import price_facade


def run_calculation(order_data: dict) -> dict:
    """
    Главная функция расчёта заказа
    
    Args:
        order_data: Словарь с данными заказа:
            - meta: метаданные (номер заказа и т.д.)
            - common: общие параметры (тип изделия, система профиля, стекло и т.д.)
            - positions: список позиций окон/дверей
            - facade: данные фасада (если есть)
    
    Returns:
        Словарь с результатами расчёта
    """
    common = order_data.get("common", {})
    positions = order_data.get("positions", [])
    facade_data = order_data.get("facade")
    
    # Определяем тип изделия
    product_type = common.get("product_type", "")
    
    # Получаем параметры
    glass_type = common.get("glass_type", "Двойной")
    profile_system = common.get("profile_system", "ALG 2030-73C")
    toning = "Есть" if common.get("toning") else "Нет"
    assembly = "Есть" if common.get("assembly") else "Нет"
    installation = common.get("installation", "Нет")
    
    # Справочник цен (здесь нужно подключить реальный справочник)
    ref2 = {
        "Двойной": 9000,
        "Тройной": 12000,
        "Тонировка": 1500,
        "Сборка": 2000,
        "Монтаж": 3000
    }
    
    # Если это фасад
    if facade_data:
        result = calc_facade_final(
            facade_data=facade_data,
            ref2=ref2,
            glass_type=glass_type,
            toning=toning,
            assembly=assembly,
            installation=installation
        )
    # Если это окна/двери
    else:
        result = calc_windows_doors_final(
            positions=positions,
            ref2=ref2,
            glass_type=glass_type,
            profile_system=profile_system,
            toning=toning,
            assembly=assembly,
            installation=installation
        )
    
    # Добавляем метаданные
    result["meta"] = order_data.get("meta", {})
    result["common"] = common
    
    return result


def calc_windows_doors_final(
    positions: list,
    ref2: dict,
    glass_type: str,
    profile_system: str,
    toning: str,
    assembly: str,
    installation: str
) -> dict:
    """
    Финальный расчёт окон / дверей
    """

    # 1. Геометрия
    geometry = calculate_windows_geometry(positions)

    # 2. Материалы
    materials = materials_positions(geometry)

    # 3. Базовая цена
    base_price = price_windows_doors(
        geometry=geometry,
        materials=materials,
        ref2=ref2,
        glass_type=glass_type,
        profile_system=profile_system
    )


    # 4. Опции
    options_price = price_options_windows(
        geometry=geometry,
        ref2=ref2,
        glass_type=glass_type,
        toning=toning,
        assembly=assembly,
        installation=installation
    )

    total = base_price["total"] + options_price["total"]

    return {
        "geometry": geometry,
        "materials": materials,
        "price": {
            "base": base_price,
            "options": options_price,
            "total": round(total, 2)
        }
    }


def calc_facade_final(
    facade_data: dict,
    ref2: dict,
    glass_type: str,
    toning: str,
    assembly: str,
    installation: str
) -> dict:
    """
    Финальный расчёт фасада
    """

    geometry = facade_glass_and_panels(facade_data)
    materials = materials_facade(facade_data)

    base_price = price_facade(
        facade_areas=geometry,
        facade_materials=materials
    )

    area = geometry["total_facade_area_m2"]
    options = {}
    total_options = 0.0

    if toning == "Есть":
        cost = area * ref2.get("Тонировка", 0)
        options["toning"] = cost
        total_options += cost

    if assembly == "Есть":
        cost = area * ref2.get("Сборка", 0)
        options["assembly"] = cost
        total_options += cost

    if installation != "Нет":
        cost = area * ref2.get(installation, 0)
        options["installation"] = cost
        total_options += cost

    options["total"] = round(total_options, 2)

    return {
        "geometry": geometry,
        "materials": materials,
        "price": {
            "base": base_price,
            "options": options,
            "total": round(base_price["total"] + options["total"], 2)
        }
    }
