import math
from typing import Dict, List


def calculate_facade_smeta(order_data: Dict, ref2: Dict) -> Dict:
    """
    Финальный расчет сметы для ФАСАДОВ.
    Основан 1-в-1 на итоговом блоке calculate_window_smeta.
    """

    common = order_data.get("common", {})
    positions = order_data.get("positions", [])

    result = {
        "metrics": {
            "total_area": 0.0
        },
        "part3_final": {},
        "total_with_margin": 0.0
    }

    # ==================================================
    # ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ЦЕН
    # ==================================================
    def get_price_from_ref2(key_word: str) -> float:
        price = ref2.get(key_word)
        if price is None:
            print(f"⚠️ WARNING: Цена для '{key_word}' не найдена в Справочнике-2!")
            return 0.0
        return float(price)

    # ==================================================
    # ПЛОЩАДЬ ФАСАДА
    # ==================================================
    total_area = 0.0

    for position in positions:
        data = position.get("data", {})
        W = data.get("width", 0) / 1000
        H = data.get("height", 0) / 1000
        count = position.get("count", 1)

        total_area += W * H * count

    result["metrics"]["total_area"] = round(total_area, 3)

    # ==================================================
    # СТЕКЛОПАКЕТ И ЛАМБРИ (РАЗДЕЛЬНО)
    # ==================================================
    cost_glass = 0.0
    cost_lambri = 0.0

    for position in positions:
        data = position.get("data", {})
        fill_cat = data.get("fill_category", "Стеклопакет")
        glass_type = data.get("glass_type", "Двойной")

        W = data.get("width", 0) / 1000
        H = data.get("height", 0) / 1000
        count = position.get("count", 1)
        area = W * H * count

        # --- Стеклопакет ---
        if fill_cat == "Стеклопакет":
            price_glass = get_price_from_ref2(glass_type)
            cost_glass += area * price_glass

        # --- Ламбри ---
        elif "Ламбри" in fill_cat:
            price_lambri = get_price_from_ref2(fill_cat)

            # отпуск хлыстами по 6 м
            qty_hlysti = math.ceil(area / 6)
            total_meters = qty_hlysti * 6

            cost_lambri += total_meters * price_lambri

    # ==================================================
    # ДОПОЛНИТЕЛЬНЫЕ РАБОТЫ
    # ==================================================
    # Тонировка
    cost_toning = 0.0
    toning = common.get("toning_id") or common.get("toning", "Нет")
    if toning == "Есть":
        price_toning = get_price_from_ref2("Тонировка")
        cost_toning = total_area * price_toning

    # Сборка
    cost_assembly = 0.0
    assembly = common.get("assembly_id") or common.get("assembly", "Нет")
    if assembly == "Есть":
        price_assembly = get_price_from_ref2("Сборка")
        cost_assembly = total_area * price_assembly

    # Монтаж (любой тип)
    cost_installation = 0.0
    installation = common.get("installation_id") or common.get("installation", "Нет")
    if installation != "Нет":
        price_installation = get_price_from_ref2(installation)
        cost_installation = total_area * price_installation

    # ==================================================
    # ИТОГ
    # ==================================================
    result["part3_final"] = {
        "Стеклопакет": round(cost_glass, 0),
        "Ламбри": round(cost_lambri, 0),
        "Тонировка": round(cost_toning, 0),
        "Сборка": round(cost_assembly, 0),
        "Монтаж": round(cost_installation, 0)
    }

    subtotal = sum(result["part3_final"].values())
    margin = subtotal * 0.65

    result["part3_final"]["Обеспечение (65%)"] = round(margin, 0)
    result["total_with_margin"] = round(subtotal + margin, 0)

    return result
