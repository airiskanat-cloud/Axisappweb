"""
Модуль для маппинга типов изделий и систем в CODE
"""

def get_code_for_windows_doors(product_type: str, system_id: str) -> str:
    """
    Преобразует тип изделия и систему в CODE для поиска в справочнике
    
    Args:
        product_type: Тип изделия ("Окно с откр.", "Дверь 2-х створч." и т.д.)
        system_id: ID системы ("ALG 2030-45C", "ALG 2030-63C" и т.д.)
    
    Returns:
        CODE для поиска в справочнике (например: "window_opening_ALG_2030_45C")
    """
    
    # Маппинг типов изделий
    product_mapping = {
        "Окно с откр.": "window_opening",
        "Окно глух.": "window_blind",
        "Дверь 2-х створч.": "door_double",
        "Дверь 1 створч.": "door_single"
    }
    
    # Нормализация системы: пробелы → подчёркивания
    system_normalized = system_id.replace(" ", "_").replace("-", "_")
    
    # Получаем префикс типа
    product_prefix = product_mapping.get(product_type, "unknown")
    
    # Формируем CODE
    code = f"{product_prefix}_{system_normalized}"
    
    return code


def get_code_for_facade(facade_type: str) -> str:
    """
    Преобразует тип фасада в CODE
    
    Args:
        facade_type: Тип фасада ("Фасадная система (Ruit 50F)" и т.д.)
    
    Returns:
        CODE для фасада
    """
    
    facade_mapping = {
        "Фасадная система (Ruit 50F)": "FACADE_RUIT_50F",
        "Оконный тамбур": "TAMBOUR"
    }
    
    return facade_mapping.get(facade_type, "FACADE_UNKNOWN")
