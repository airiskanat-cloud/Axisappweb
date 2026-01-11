# order_data.py
# Этот файл определяет стандарт того, как выглядят данные в Axis pro GF.

def create_empty_position(pos_type="Окно/Дверь"):
    """Создает пустую структуру для новой позиции в списке"""
    if pos_type == "Фасад":
        return {
            "type": "Фасад",
            "width": 0,
            "height": 0,
            "grid": {"cols": 1, "rows": 1},
            "panels": [],    # Индексы ячеек, где стоят панели (ламбри)
            "inserts": [],   # Список объектов (окон/дверей), вставленных в ячейки
            "is_active": True
        }
    else:
        return {
            "type": "Окно/Дверь",
            "width": 0,
            "height": 0,
            "imposts": {"left": 0, "right": 0, "center": 0, "tor": 0},
            "sashes": [],    # Список словарей {"width": 0, "height": 0}
            "is_active": True
        }

def create_order_structure():
    """Главная структура заказа для передачи в расчетные модули"""
    return {
        "meta": {
            "order_number": "",
            "date": ""
        },
        "common": {
            "system_id": "",        # Справочник-1
            "glass_id": "",         # Справочник-2
            "toning_id": "Нет",     # Справочник-2
            "assembly_id": "Нет",   # Справочник-2
            "installation_id": "Нет" # Справочник-2
        },
        "positions": []             # Сюда добавляются объекты из create_empty_position
    }
