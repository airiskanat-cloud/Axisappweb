# calculations/pricing.py
from typing import Dict, List

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
