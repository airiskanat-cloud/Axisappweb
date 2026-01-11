# calculations/geometry.py
from typing import List, Dict

MM_TO_M = 0.001
MM2_TO_M2 = 0.000001


def calc_area_m2(width_mm: float, height_mm: float) -> float:
    return round(width_mm * height_mm * MM2_TO_M2, 3)


def calc_perimeter_m(width_mm: float, height_mm: float) -> float:
    return round(2 * (width_mm + height_mm) * MM_TO_M, 3)


def process_window_geometry(pos: Dict) -> Dict:
    """Расчет геометрии окна/двери (включая импосты и створки)"""

    # === ✅ ДОБАВЛЕННАЯ ЛОГИКА ДЛЯ ГЛУХОГО ОКНА ===
    if pos.get("product_type") == "Окно глух.":

        w = pos.get("width", 0)
        h = pos.get("height", 0)

        area = calc_area_m2(w, h)
        perimeter = calc_perimeter_m(w, h)

        return {
            "total_area_m2": area,
            "net_glass_area_m2": area,
            "panels_area_m2": area,
            "facade_profile_m": 0.0,
            "inserts_profile_m": 0.0,
            "total_profile_combined_m": perimeter,
            "inserts_details": []
        }

    # === ⬇️ НИЖЕ ИДЁТ ТВОЙ СУЩЕСТВУЮЩИЙ КОД (НЕ ТРОГАЛ) ===

    w, h = pos.get("width", 0), pos.get("height", 0)
    area = calc_area_m2(w, h)
    ...
    # Стойки (вертикаль): (кол-во столбцов + 1) * высота
    vertical_profile = (cols + 1) * total_h * MM_TO_M
    # Ригели (горизонталь): (кол-во строк + 1) * ширина
    horizontal_profile = (rows + 1) * total_w * MM_TO_M

    return {
        "total_area_m2": total_area,
        "net_glass_area_m2": round(net_glass_area, 3),
        "panels_area_m2": round(panels_area, 3),
        "facade_profile_m": round(vertical_profile + horizontal_profile, 3),
        "inserts_profile_m": round(sum_inserts_profile, 3),
        "total_profile_combined_m": round(
            vertical_profile + horizontal_profile + sum_inserts_profile, 3
        ),
        "inserts_details": inserts_geometry
    }
