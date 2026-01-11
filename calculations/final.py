from calculations.geometry import geometry_positions
from calculations.geometry import geometry_positions_extended
from calculations.materials import materials_positions
from calculations.pricing import price_windows_doors, price_options_windows


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
    geometry = geometry_positions(positions)
    geometry_ext = geometry_positions_extended(positions)

    # 2. Материалы
    materials = materials_positions(geometry_ext)

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
from calculations.facade import facade_glass_and_panels
from calculations.materials import materials_facade
from calculations.pricing import price_facade


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

    # 1. Геометрия фасада
    geometry = facade_glass_and_panels(facade_data)

    # 2. Материалы фасада
    materials = materials_facade(facade_data)

    # 3. Базовая цена фасада
    base_price = price_facade(
        facade_areas=geometry,
        facade_materials=materials
    )

    # 4. Опции фасада (пока считаем так же, как окна)
    area = geometry["total_facade_area_m2"]
    ref = ref2.get(glass_type, {})

    options = {}
    total_options = 0.0

    if toning == "Есть":
        price = ref.get("Стоимость тонировки за квадратный метр", 0)
        cost = area * float(price)
        options["toning"] = cost
        total_options += cost

    if assembly == "Есть":
        price = ref.get("Стоимость сборки за квадратный метр", 0)
        cost = area * float(price)
        options["assembly"] = cost
        total_options += cost

    if installation.strip() != "Нет":
        price = ref.get("Стоимость монтаж  за квадратный метр", 0)
        cost = area * float(price)
        options["installation"] = cost
        total_options += cost

    options["total"] = round(total_options, 2)

    total = base_price["total"] + options["total"]

    return {
        "geometry": geometry,
        "materials": materials,
        "price": {
            "base": base_price,
            "options": options,
            "total": round(total, 2)
        }
    }

