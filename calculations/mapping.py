"""
Модуль для маппинга типов изделий и систем в CODE
ИСПРАВЛЕНО: Добавлен fallback для неизвестных систем
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
        "Дверь 1 створч.": "DOOR_SINGLE",
        # Добавляем варианты написания
        "окно с откр.": "WINDOW_OPEN",
        "окно глух.": "WINDOW_FIXED",
        "дверь 2-х створч.": "DOOR_DOUBLE",
        "дверь 1 створч.": "DOOR_SINGLE",
        "Окно": "WINDOW_OPEN",
        "Дверь": "DOOR_DOUBLE"
    }
    
    # Извлекаем только цифровую часть системы (без "ALG ")
    # "ALG 2030-45C" → "2030_45C"
    # "ALG RUIT 73i 22MM" → "RUIT_73I_22MM"
    system_clean = system_id.replace("ALG ", "").replace(" ", "_").replace("-", "_")
    
    # Получаем префикс типа
    product_prefix = product_mapping.get(product_type, "")
    
    # FALLBACK 1: Если тип не распознан
    if not product_prefix:
        print(f"⚠️ Неизвестный тип изделия: '{product_type}'")
        print(f"   Доступные типы: {list(product_mapping.keys())}")
        # Пробуем угадать
        if "окн" in product_type.lower():
            product_prefix = "WINDOW_OPEN"
            print(f"   → Использую WINDOW_OPEN")
        elif "двер" in product_type.lower():
            product_prefix = "DOOR_DOUBLE"
            print(f"   → Использую DOOR_DOUBLE")
        else:
            product_prefix = "UNKNOWN"
    
    # Формируем CODE (ЗАГЛАВНЫМИ БУКВАМИ, БЕЗ ALG)
    code = f"{product_prefix}_{system_clean}".upper()
    
    # FALLBACK 2: Если CODE содержит UNKNOWN
    if "UNKNOWN" in code:
        print(f"⚠️ Не удалось определить CODE для '{product_type}' / '{system_id}'")
        print(f"   Сгенерированный CODE: {code}")
        print(f"   Попробую использовать только систему: {system_clean}")
        # Возвращаем хотя бы систему
        return system_clean.upper()
    
    # FALLBACK 3: Валидация CODE
    if not code or code == "_":
        print(f"⚠️ Пустой CODE для '{product_type}' / '{system_id}'")
        return system_clean.upper() if system_clean else "UNKNOWN"
    
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
        "Оконный тамбур": "TAMBOUR",
        # Добавляем варианты
        "Ruit 50F": "FACADE_RUIT_50F",
        "Ruit50F": "FACADE_RUIT_50F",
        "Фахверк": "FACADE_FAKHVERK",
        "Фасад": "FACADE_RUIT_50F"
    }
    
    code = facade_mapping.get(facade_type, "")
    
    # FALLBACK
    if not code:
        print(f"⚠️ Неизвестный тип фасада: '{facade_type}'")
        # Пробуем извлечь из строки
        if "ruit" in facade_type.lower() or "50f" in facade_type.lower():
            code = "FACADE_RUIT_50F"
            print(f"   → Использую FACADE_RUIT_50F")
        elif "тамбур" in facade_type.lower():
            code = "TAMBOUR"
            print(f"   → Использую TAMBOUR")
        elif "фахверк" in facade_type.lower():
            code = "FACADE_FAKHVERK"
            print(f"   → Использую FACADE_FAKHVERK")
        else:
            code = "FACADE_UNKNOWN"
            print(f"   → Использую FACADE_UNKNOWN")
    
    return code


def diagnose_code_issue(product_type: str, system_id: str, ref1: list) -> None:
    """
    Диагностика проблем с CODE
    Помогает понять почему не находятся материалы
    
    Args:
        product_type: Тип изделия
        system_id: Система
        ref1: Справочник-1
    """
    print("\n" + "="*70)
    print("🔍 ДИАГНОСТИКА CODE")
    print("="*70)
    
    code = get_code_for_windows_doors(product_type, system_id)
    print(f"\nИсходные данные:")
    print(f"  Тип изделия: {product_type}")
    print(f"  Система: {system_id}")
    print(f"  Сгенерированный CODE: {code}")
    
    # Ищем CODE в справочнике
    found = False
    for item in ref1[:100]:  # Проверяем первые 100
        item_code = str(item.get("CODE", "")).strip()
        if code and item_code == code:
            found = True
            print(f"\n✅ CODE найден в Справочнике-1!")
            print(f"   Артикул: {item.get('Артикул', 'N/A')}")
            print(f"   Элемент: {item.get('Элемент', 'N/A')}")
            break
    
    if not found:
        print(f"\n❌ CODE не найден в Справочнике-1!")
        
        # Показываем похожие CODE
        print(f"\n📋 Похожие CODE в справочнике:")
        similar = []
        for item in ref1[:200]:
            item_code = str(item.get("CODE", "")).strip()
            if item_code and system_id.replace("ALG ", "").replace(" ", "").replace("-", "") in item_code:
                similar.append(item_code)
        
        if similar:
            for sc in list(set(similar))[:5]:
                print(f"   - {sc}")
        else:
            print("   Похожих не найдено")
        
        # Показываем доступные системы
        print(f"\n📋 Доступные системы в справочнике (первые 10):")
        systems = set()
        for item in ref1[:500]:
            sys = item.get("Система", "")
            if sys:
                systems.add(sys)
            if len(systems) >= 10:
                break
        
        for sys in sorted(systems):
            print(f"   - {sys}")
    
    print("="*70 + "\n")
