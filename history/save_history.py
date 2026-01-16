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
        # ИСПРАВЛЕНО: Формат времени ДД.ММ.ГГГГ ЧЧ:ММ
        now = datetime.datetime.now()
        date_str = now.strftime("%d.%m.%Y")  # 16.01.2026
        time_str = now.strftime("%H:%M")      # 19:06
        timestamp = f"{date_str} {time_str}"
        
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
        
        # ИСПРАВЛЕНО: Определяем тип изделия ОБЯЗАТЕЛЬНО для корректной истории
        calc_type = "Окно/Дверь"  # По умолчанию
        facade_type = result.get("facade_type", "")
        if facade_type:
            if "тамбур" in facade_type.lower():
                calc_type = "Оконный тамбур"
            elif "фасад" in facade_type.lower():
                calc_type = "Фасад"
        
        # DEBUG
        print(f"\n=== СОХРАНЕНИЕ ИСТОРИИ ===")
        print(f"Тип расчёта: {calc_type}")
        print(f"Позиций: {len(positions)}")
        if positions:
            print(f"Первая позиция: {positions[0]}")
        
        # Габариты позиций + вставки для фасада
        gabarits = []
        for idx, pos in enumerate(positions):
            w = pos.get("width", 0)
            h = pos.get("height", 0)
            
            print(f"Позиция {idx+1}: w={w}, h={h}, тип={type(w)}")
            
            # Конвертируем мм в метры для окон/дверей
            if w > 100 or h > 100:  # Если больше 100, значит в мм
                w = w / 1000
                h = h / 1000
            
            if w > 0 and h > 0:
                # Тип изделия
                product_type = pos.get("product_type", "")
                if product_type:
                    pos_str = f"П{idx+1} ({product_type}): {w:.2f}м×{h:.2f}м"
                else:
                    pos_str = f"П{idx+1}: {w:.2f}м×{h:.2f}м"
                
                # Если фасад - добавляем сетку
                cols = pos.get("columns", 0)
                rows = pos.get("rows", 0)
                if cols > 0 and rows > 0:
                    pos_str += f" ({cols}×{rows})"
                
                # Если есть вставки
                inserts_data = pos.get("insert_data", {})
                if inserts_data:
                    insert_w = inserts_data.get("width", 0)
                    insert_h = inserts_data.get("height", 0)
                    # Конвертируем мм в метры
                    if insert_w > 100:
                        insert_w = insert_w / 1000
                    if insert_h > 100:
                        insert_h = insert_h / 1000
                    if insert_w > 0 and insert_h > 0:
                        pos_str += f" | Вставка: {insert_w:.2f}×{insert_h:.2f}"
                
                gabarits.append(pos_str)
        
        gabarits_str = "; ".join(gabarits) if gabarits else "-"
        
        # Строка для записи (ОБНОВЛЁННАЯ структура с типом изделия)
        row = [
            timestamp,           # A: Дата и время
            user_login,          # B: Пользователь
            calc_type,           # C: Тип изделия (ОБЯЗАТЕЛЬНО!)
            order_number,        # D: Номер заказа
            gabarits_str,        # E: Габариты (полные)
            n_positions,         # F: Позиций
            f"{total_area:.2f}", # G: Площадь
            cost_formatted       # H: Стоимость
        ]
        
        # Добавляем строку в конец
        ws.append_row(row)
        
        print(f"✅ История сохранена: {order_number} ({timestamp})")
        
    except Exception as e:
        print(f"❌ Ошибка сохранения истории: {e}")
        # Не прерываем работу приложения при ошибке сохранения истории
        pass
