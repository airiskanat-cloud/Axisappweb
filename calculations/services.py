# calculations/services.py

def calc_services(order_data: dict, geometry: dict, services_ref: dict) -> dict:
    """
    Расчет услуг (сборка, монтаж) от общей площади
    services_ref = {
        "assembly": price_per_m2,
        "installation": price_per_m2
    }
    """

    area = geometry.get("total_area", 0)

    result = {
        "assembly": {
            "price_per_m2": 0,
            "area": area,
            "sum": 0
        },
        "installation": {
            "price_per_m2": 0,
            "area": area,
            "sum": 0
        }
    }

    # --- СБОРКА ---
    if order_data["common"].get("assembly"):
        price = services_ref.get("assembly", 0)
        result["assembly"]["price_per_m2"] = price
        result["assembly"]["sum"] = round(area * price, 2)

    # --- МОНТАЖ ---
    if order_data["common"].get("installation"):
        price = services_ref.get("installation", 0)
        result["installation"]["price_per_m2"] = price
        result["installation"]["sum"] = round(area * price, 2)

    return result
