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
        
        # РАСШИРЕННАЯ детализация позиций с ПОЛНЫМИ данными
        gabarits = []
        full_details = []  # Полная информация для отдельной колонки
        
        for idx, pos in enumerate(positions):
            # Базовые размеры
            pos_data = pos.get("data", {})
            w = pos_data.get("width", 0) if pos_data else pos.get("width", 0)
            h = pos_data.get("height", 0) if pos_data else pos.get("height", 0)
            
            print(f"Позиция {idx+1}: w={w}, h={h}, тип={type(w)}")
            
            # Конвертируем мм в метры для окон/дверей
            if w > 100 or h > 100:  # Если больше 100, значит в мм
                w_m = w / 1000
                h_m = h / 1000
            else:
                w_m = w
                h_m = h
            
            if w_m > 0 and h_m > 0:
                # === БАЗОВАЯ ИНФОРМАЦИЯ ===
                product_type = pos.get("product_type", "")
                system_id = pos.get("system_id", "")
                
                # Краткая версия (для габаритов)
                if product_type:
                    pos_str = f"П{idx+1} ({product_type}): {w_m:.2f}м×{h_m:.2f}м"
                else:
                    pos_str = f"П{idx+1}: {w_m:.2f}м×{h_m:.2f}м"
                
                # ПОЛНАЯ версия (для детальной колонки)
                detail_parts = []
                detail_parts.append(f"▸ П{idx+1}: {product_type or 'Не указано'}")
                detail_parts.append(f"  Размер: {w_m:.2f}м × {h_m:.2f}м ({w:.0f}×{h:.0f}мм)")
                if system_id:
                    detail_parts.append(f"  Система: {system_id}")
                
                # === ЗАПОЛНЕНИЕ ===
                fill_category = pos_data.get("fill_category", "")
                if fill_category:
                    detail_parts.append(f"  Заполнение: {fill_category}")
                    
                    # Тип стеклопакета
                    if fill_category == "Стеклопакет":
                        glass_type = pos_data.get("glass_type", "")
                        if glass_type:
                            detail_parts.append(f"  └─ Тип: {glass_type}")
                    
                    # Тип ламбри
                    elif "Ламбри" in fill_category:
                        lambri_type = pos_data.get("lambri_type", "")
                        if lambri_type:
                            detail_parts.append(f"  └─ Тип: {lambri_type}")
                
                # === ДОПОЛНИТЕЛЬНЫЕ УСЛУГИ ===
                services = []
                
                # Тонировка
                toning = pos_data.get("toning", "")
                if toning and toning != "Нет":
                    services.append(f"Тонировка: {toning}")
                
                # Сборка
                assembly = pos_data.get("assembly", "")
                if assembly and assembly != "Нет":
                    services.append(f"Сборка: {assembly}")
                
                # Монтаж
                installation = pos_data.get("installation", "")
                if installation and installation != "Нет":
                    services.append(f"Монтаж: {installation}")
                
                # Доп.детали
                additional = pos_data.get("additional", "")
                if additional and additional != "Нет":
                    services.append(f"Доп.детали: {additional}")
                
                if services:
                    detail_parts.append(f"  Услуги: {', '.join(services)}")
                
                # === ФАСАД: Сетка ===
                cols = pos.get("columns", 0)
                rows = pos.get("rows", 0)
                if cols > 0 and rows > 0:
                    pos_str += f" ({cols}×{rows})"
                    detail_parts.append(f"  Сетка: {cols} колонок × {rows} рядов")
                
                # === ФАСАД: Вставки ===
                inserts_data = pos.get("insert_data", {})
                if inserts_data:
                    insert_w = inserts_data.get("width", 0)
                    insert_h = inserts_data.get("height", 0)
                    insert_type = inserts_data.get("insert_product_type", "")
                    insert_system = inserts_data.get("insert_system", "")
                    
                    # Конвертируем мм в метры
                    if insert_w > 100:
                        insert_w_m = insert_w / 1000
                    else:
                        insert_w_m = insert_w
                    if insert_h > 100:
                        insert_h_m = insert_h / 1000
                    else:
                        insert_h_m = insert_h
                    
                    if insert_w_m > 0 and insert_h_m > 0:
                        pos_str += f" | Вставка: {insert_w_m:.2f}×{insert_h_m:.2f}"
                        detail_parts.append(f"  Вставка: {insert_type or 'Дверь/Окно'}")
                        detail_parts.append(f"  └─ Размер: {insert_w_m:.2f}м × {insert_h_m:.2f}м")
                        if insert_system:
                            detail_parts.append(f"  └─ Система: {insert_system}")
                
                # Добавляем в массивы
                gabarits.append(pos_str)
                full_details.append("\n".join(detail_parts))
        
        # Форматируем для записи
        gabarits_str = "; ".join(gabarits) if gabarits else "-"
        details_str = "\n\n".join(full_details) if full_details else "-"
        
        # Строка для записи (РАСШИРЕННАЯ структура)
        row = [
            timestamp,           # A: Дата и время
            user_login,          # B: Пользователь
            calc_type,           # C: Тип изделия
            order_number,        # D: Номер заказа
            gabarits_str,        # E: Габариты (краткие)
            details_str,         # F: ДЕТАЛИ (полные данные)
            n_positions,         # G: Позиций
            f"{total_area:.2f}", # H: Площадь
            cost_formatted       # I: Стоимость
        ]
        
        # Добавляем строку в конец
        ws.append_row(row)
        
        print(f"✅ История сохранена: {order_number} ({timestamp})")
        
    except Exception as e:
        print(f"❌ Ошибка сохранения истории: {e}")
        # Не прерываем работу приложения при ошибке сохранения истории
        pass
