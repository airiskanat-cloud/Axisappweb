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
    # ИСПРАВЛЕНО: Обрабатываем ВСЕ позиции
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
    
    # Накопительные результаты
    all_materials = []
    total_area = 0.0
    total_perimeter = 0.0
    all_costs = {}
    
    # Обрабатываем каждую позицию
    for idx, position in enumerate(positions, 1):
        data = position.get("data", {})
        usage_mode = UsageMode.EMBEDDED if data.get("embedded", False) else UsageMode.STANDALONE
        
        # Расчёт для одной позиции
        pos_result = calculate_product_materials(
            product_data={
                "product_type": position.get("product_type", "Окно"),
                "system": position.get("system_id", position.get("system", "ALG 2030-45C")),
                "code": position.get("code", ""),
                "data": data,
                "common": order_data.get("common", {})
            },
            ref1=ref1,
            ref2=ref2,
            ref3=ref3,
            usage_mode=usage_mode
        )
        
        # Добавляем материалы
        all_materials.extend(pos_result.get("part2_materials", []))
        
        # Суммируем метрики
        metrics = pos_result.get("metrics", {})
        total_area += metrics.get("total_area", 0.0)
        total_perimeter += metrics.get("total_perimeter", 0.0)
        
        # Суммируем part3_final
        for key, value in pos_result.get("part3_final", {}).items():
            all_costs[key] = all_costs.get(key, 0) + value
    
    # Пересчитываем обеспечение на общую сумму
    if "Обеспечение" in all_costs:
        del all_costs["Обеспечение"]
    
    subtotal = sum(all_costs.values())
    margin = subtotal * 0.81
    all_costs["Обеспечение"] = round(margin, 0)
    
    result = {
        "part1_gabarits": [],
        "part2_materials": all_materials,
        "part3_final": all_costs,
        "metrics": {
            "total_area": round(total_area, 3),
            "total_perimeter": round(total_perimeter, 3)
        },
        "total_with_margin": round(subtotal + margin, 0),
        "materials_cost": round(subtotal, 0),
        "part1_summary": []  # Алиас
    }
    
    return result
