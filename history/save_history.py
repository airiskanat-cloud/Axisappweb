"""
Модуль для сохранения истории заказов в Google Sheets
Каждый расчёт сохраняется как отдельная версия
"""

import datetime
import gspread
from google.oauth2.service_account import Credentials


def save_history(
    credentials_path: str,
    spreadsheet_id: str,
    user_login: str,
    order_data: dict,
    result: dict
):
    """
    Сохраняет историю расчёта в Google Sheets (лист "ИСТОРИЯ")
    
    Args:
        credentials_path: Путь к файлу credentials.json
        spreadsheet_id: ID таблицы Google Sheets
        user_login: Логин пользователя
        order_data: Данные заказа (common + positions)
        result: Результаты расчёта
    """
    
    try:
        # Авторизация
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
        gc = gspread.authorize(creds)
        
        # Открываем таблицу и лист ИСТОРИЯ
        sh = gc.open_by_key(spreadsheet_id)
        ws = sh.worksheet("ИСТОРИЯ")
        
        # Данные для записи
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Извлекаем данные
        common = order_data.get("common", {})
        positions = order_data.get("positions", [])
        
        order_number = common.get("order_number", "")
        n_positions = len(positions)
        
        # Метрики
        metrics = result.get("metrics", {})
        total_area = metrics.get("total_area", 0)
        
        # Итоговая стоимость
        total_cost = result.get("total_with_margin", 0)
        if total_cost == 0:
            total_cost = result.get("total_cost", 0)
        
        # Форматируем стоимость
        cost_formatted = f"{total_cost:,.0f}".replace(",", " ")
        
        # Строка для записи
        row = [
            timestamp,
            user_login,
            order_number,
            n_positions,
            f"{total_area:.2f}",
            cost_formatted
        ]
        
        # Добавляем строку в конец
        ws.append_row(row)
        
        print(f"✅ История сохранена: {order_number} ({timestamp})")
        
    except Exception as e:
        print(f"❌ Ошибка сохранения истории: {e}")
        # Не прерываем работу приложения при ошибке сохранения истории
        pass
