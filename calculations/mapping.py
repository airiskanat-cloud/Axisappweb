"""
Маппинг типов изделий и систем на CODE для поиска в Справочнике1
"""

def get_code_for_windows_doors(product_type: str, system: str) -> str:
    """
    Генерирует CODE для окон/дверей по типу изделия и системе
    
    Args:
        product_type: "Окно с откр.", "Окно глухое", "Дверь 1 створч.", "Дверь 2-х створч."
        system: "ALG 2030-73C", "ALG 2030-63C", "ALG 2030-55C", "ALG 2030-45C", "ALG 2030-Slim"
    
    Returns:
        CODE для поиска в Справочнике1
    
    Raises:
        ValueError: Если комбинация не найдена
    """
    
    # Нормализация системы
    system_normalized = system.upper().strip()
    
    # Извлекаем толщину системы
    system_map = {
        "ALG 2030-73C": "73C",
        "ALG 2030-63C": "63C",
        "ALG 2030-55C": "55C",
        "ALG 2030-45C": "45C",
        "ALG 2030-SLIM": "SLIM",
        # RUIT системы
        "ALG RUIT 73I 22MM": "73C",
        "ALG RUIT 63I": "63C",
        "ALG RUIT 55I": "55C",
        "ALG RUIT 45I": "45C"
    }
    
    system_key = None
    for key, value in system_map.items():
        if key in system_normalized:
            system_key = value
            break
    
    if not system_key:
        raise ValueError(f"[MAPPING ERROR] Неизвестная система: {system}")
    
    # Маппинг по типу изделия
    product_type_lower = product_type.lower()
    
    # ОКНА
    if "окно" in product_type_lower:
        if "глух" in product_type_lower:
            # Окно глухое = FIXED
            code = f"WINDOW_FIXED_2030_{system_key}"
        else:
            # Окно с откр.
            code = f"WINDOW_OPEN_2030_{system_key}"
    
    # ДВЕРИ
    elif "дверь" in product_type_lower:
        if "1" in product_type or "одн" in product_type_lower:
            # Дверь 1 створч.
            code = f"DOOR_SINGLE_2030_{system_key}"
        elif "2" in product_type or "двух" in product_type_lower:
            # Дверь 2-х створч.
            code = f"DOOR_DOUBLE_2030_{system_key}"
        else:
            raise ValueError(f"[MAPPING ERROR] Неизвестный тип двери: {product_type}")
    
    else:
        raise ValueError(f"[MAPPING ERROR] Неизвестный тип изделия: {product_type}")
    
    return code


def get_code_for_facade(facade_type: str) -> str:
    """
    Генерирует CODE для фасадных систем
    
    Args:
        facade_type: "Фасадная система (Ruit 50F)" или "Оконный тамбур (ALG)"
    
    Returns:
        CODE для фасада
    """
    
    if "Ruit 50F" in facade_type or "ruit" in facade_type.lower():
        return "FACADE_RUIT_50F"
    elif "ALG" in facade_type or "тамбур" in facade_type.lower():
        return "FACADE_TAMBOUR_ALG"
    else:
        return "FACADE_RUIT_50F"  # По умолчанию


# Для обратной совместимости
def get_code_for_position(product_type: str, system: str) -> str:
    """Алиас для get_code_for_windows_doors"""
    return get_code_for_windows_doors(product_type, system)
