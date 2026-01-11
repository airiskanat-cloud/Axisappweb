# history/save_history.py

import json
from datetime import datetime
from references.sheets_reader import get_gc  # ИСПРАВЛЕНО: было get_client


def save_history(
    credentials_path: str,
    spreadsheet_id: str,
    user: str,
    order_data: dict,
    result_data: dict
):
    """
    Сохраняет историю расчёта в Google Sheets (лист ИСТОРИЯ)
    """
    try:
        client = get_gc(credentials_path)  # ИСПРАВЛЕНО: было get_client
        sheet = client.open_by_key(spreadsheet_id).worksheet("ИСТОРИЯ")

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # Дата и время
            user,  # Пользователь
            order_data.get("common", {}).get("order_number", ""),  # Номер заказа
            len(order_data.get("positions", [])),  # Количество позиций
            f"{result_data.get('metrics', {}).get('total_area', 0):.2f}",  # Площадь
            f"{result_data.get('total_with_margin', 0):,.0f}",  # Итого
            json.dumps(order_data, ensure_ascii=False),  # Полные данные заказа
            json.dumps(result_data, ensure_ascii=False)  # Полный результат
        ]

        sheet.append_row(row)
        return True
    except Exception as e:
        print(f"❌ Ошибка сохранения истории: {e}")
        return False
