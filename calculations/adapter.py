"""
Unified Adapter - SIMPLE VERSION
Просто вызывает calculate_product_materials который возвращает правильный формат
"""

from typing import Dict, List
from .material_calculator import calculate_product_materials
from .product_model import UsageMode


def calculate_window_smeta_unified(
    order_data: Dict,
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict]
) -> Dict:
    """
    Унифицированный расчёт с использованием ПОЛНОЙ версии калькулятора
    
    Теперь calculate_product_materials возвращает результат
    в формате engine_windows.py, поэтому просто передаём его дальше!
    
    Args:
        order_data: Данные заказа (common + positions)
        ref1, ref2, ref3: Справочники
    
    Returns:
        Dict в формате engine_windows.py:
        {
            "part1_gabarits": [...],
            "part2_materials": [...],  # С qty_fact, norm, qty_ship, row_sum
            "part3_final": {...},      # Материалы, Стеклопакет, Сборка и т.д.
            "metrics": {...},
            "total_with_margin": float
        }
    """
    # Берём первую позицию
    positions = order_data.get("positions", [])
    if not positions:
        return {
            "part1_gabarits": [],
            "part2_materials": [],
            "part3_final": {},
            "metrics": {"total_area": 0.0, "total_perimeter": 0.0},
            "total_with_margin": 0.0,
            "materials_cost": 0.0
        }
    
    position = positions[0]
    
    # Определяем usage_mode
    data = position.get("data", {})
    usage_mode = UsageMode.EMBEDDED if data.get("embedded", False) else UsageMode.STANDALONE
    
    # ПРОСТО ВЫЗЫВАЕМ calculate_product_materials
    # Он сам вернёт результат в правильном формате!
    result = calculate_product_materials(
        product_data={
            "product_type": position.get("product_type", "Окно"),
            "system": position.get("system", "ALG 2030-45C"),
            "code": position.get("code", ""),
            "data": data
        },
        ref1=ref1,
        ref2=ref2,
        ref3=ref3,
        usage_mode=usage_mode
    )
    
    # Добавляем алиасы для совместимости
    result["part1_summary"] = result.get("part1_gabarits", [])
    
    return result
