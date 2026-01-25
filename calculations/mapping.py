"""
Модуль для маппинга типов изделий и систем в CODE
"""

def get_code_for_windows_doors(product_type: str, system_id: str) -> str:
    """
    Преобразует тип изделия и систему в CODE для поиска в справочнике
    
    КРИТИЧНО: Формат CODE в Справочнике-1:
    - DOOR_DOUBLE_2030_45C (БЕЗ "ALG_", заглавные буквы)
    - WINDOW_OPEN_2030_63C
    - WINDOW_FIXED_2030_55C
    
    Args:
        product_type: Тип изделия ("Окно с откр.", "Дверь 2-х створч." и т.д.)
        system_id: ID системы ("ALG 2030-45C", "ALG 2030-63C" и т.д.)
    
    Returns:
        CODE для поиска в справочнике (например: "DOOR_DOUBLE_2030_45C")
    """
    
    # Маппинг типов изделий
    product_mapping = {
        "Окно с откр.": "WINDOW_OPEN",
        "Окно глух.": "WINDOW_FIXED",
        "Дверь 2-х створч.": "DOOR_DOUBLE",
        "Дверь 1 створч.": "DOOR_SINGLE"
    }
    
    # Извлекаем только цифровую часть системы (без "ALG ")
    # "ALG 2030-45C" → "2030_45C"
    system_clean = system_id.replace("ALG ", "").replace(" ", "").replace("-", "_")
    
    # Получаем префикс типа
    product_prefix = product_mapping.get(product_type, "UNKNOWN")
    
    # Формируем CODE (ЗАГЛАВНЫМИ БУКВАМИ, БЕЗ ALG)
    code = f"{product_prefix}_{system_clean}".upper()
    
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
