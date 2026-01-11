# calculations/pricing.py
from typing import Dict, List


def price_windows_doors(geometry: Dict, materials: Dict, ref2: Dict, 
                        glass_type: str, profile_system: str) -> Dict:
    """
    Расчёт базовой цены окон/дверей
    
    Args:
        geometry: Результаты геометрии
        materials: Материалы
        ref2: Справочник цен
        glass_type: Тип стеклопакета
        profile_system: Система профиля
    
    Returns:
        Dict с базовой ценой
    """
    total_area = geometry.get("total_area_m2", 0)
    
    # Цена стеклопакета за м²
    glass_price = ref2.get(glass_type, 0)
    if isinstance(glass_price, dict):
        glass_price = glass_price.get("Цена за кв.м.", 0)
    
    glass_cost = total_area * float(glass_price)
    
    # Цена профиля (примерная, можно доработать)
    profile_cost = total_area * 5000  # базовая цена профиля
    
    total = glass_cost + profile_cost
    
    return {
        "glass": round(glass_cost, 2),
        "profile": round(profile_cost, 2),
        "total": round(total, 2)
    }


def price_options_windows(geometry: Dict, ref2: Dict, glass_type: str,
                          toning: str, assembly: str, installation: str) -> Dict:
    """
    Расчёт цены опций для окон/дверей
    
    Args:
        geometry: Результаты геометрии
        ref2: Справочник цен
        glass_type: Тип стеклопакета
        toning: Тонировка (Есть/Нет)
        assembly: Сборка (Есть/Нет)
        installation: Монтаж
    
    Returns:
        Dict с ценами опций
    """
    total_area = geometry.get("total_area_m2", 0)
    
    options = {}
    total_options = 0.0
    
    # Тонировка
    if toning == "Есть":
        toning_price = ref2.get("Тонировка", 0)
        cost = total_area * float(toning_price)
        options["toning"] = cost
        total_options += cost
    
    # Сборка
    if assembly == "Есть":
        assembly_price = ref2.get("Сборка", 0)
        cost = total_area * float(assembly_price)
        options["assembly"] = cost
        total_options += cost
    
    # Монтаж
    if installation != "Нет":
        installation_price = ref2.get("Монтаж", 0)
        cost = total_area * float(installation_price)
        options["installation"] = cost
        total_options += cost
    
    options["total"] = round(total_options, 2)
    
    return options


def price_facade(facade_areas: Dict, facade_materials: Dict) -> Dict:
    """
    Расчёт цены фасада
    
    Args:
        facade_areas: Площади фасада
        facade_materials: Материалы фасада
    
    Returns:
        Dict с ценой фасада
    """
    total_area = facade_areas.get("total_facade_area_m2", 0)
    
    # Базовая цена фасада (примерная)
    base_price_per_m2 = 15000  # тенге за м²
    
    total = total_area * base_price_per_m2
    
    return {
        "base": round(total, 2),
        "total": round(total, 2)
    }


def calculate_final_pricing(
    materials_results: Dict, 
    geometry_results: List[Dict], 
    ref2: Dict, 
    common_data: Dict
) -> Dict:
    """
    Рассчитывает финальную стоимость заказа с наценкой 65%
    """
    # 1. Получаем базовые цены из Справочника-2
    # Ищем по ключу системы профиля или типа стекла
    glass_type = common_data.get("glass_id", "Двойной")
    prices = ref2.get(glass_type, {})
    
    # Извлекаем цены (с очисткой от пробелов и символов)
    def clean_price(key, default=0):
        val = prices.get(key, default)
        try:
            return float(str(val).replace("\xa0", "").replace(" ", "").replace(",", "."))
        except:
            return default

    price_glass_m2 = clean_price("Стоимость стеклопакета за квадратный метр")
    price_toning_m2 = clean_price("Стоимость тонировки за квадратный метр")
    price_assembly_m2 = clean_price("Стоимость сборки за квадратный метр")
    price_install_m2 = clean_price("Стоимость монтаж  за квадратный метр")
    
    # Цены на профиль (можно вынести в настройки или брать из справочника)
    price_profile_m = 1500  # Примерная цена за метр (нужно уточнить в справочнике 1)

    # 2. Считаем себестоимость (Cost)
    # Суммируем площади всех позиций для стекла и услуг
    total_area = sum(g.get("total_area_m2", g.get("area_m2", 0)) for g in geometry_results)
    
    cost_materials = (
        materials_results["shipping_amounts"]["facade_profile_ship"] + 
        materials_results["shipping_amounts"]["window_profile_ship"]
    ) * price_profile_m
    
    cost_glass = total_area * price_glass_m2
    cost_services = total_area * (price_assembly_m2 + price_install_m2)
    
    if common_data.get("toning_id") != "Нет":
        cost_services += total_area * price_toning_m2

    subtotal = cost_materials + cost_glass + cost_services

    # 3. Применяем наценку 65% (коэффициент 1.65)
    margin_coefficient = 1.65
    final_total = subtotal * margin_coefficient

    return {
        "cost_details": {
            "materials": round(cost_materials, 2),
            "glass": round(cost_glass, 2),
            "services": round(cost_services, 2),
            "subtotal_no_margin": round(subtotal, 2)
        },
        "margin_percent": "65%",
        "total_with_margin": round(final_total, 2)
    }
