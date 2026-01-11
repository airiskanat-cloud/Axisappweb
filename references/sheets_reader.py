import pandas as pd
import gspread
import streamlit as st
import logging
from google.oauth2.service_account import Credentials

logger = logging.getLogger(__name__)

def parse_price(value):
    if value is None:
        return 0.0

    value = str(value).strip()
    if value == "":
        return 0.0

    return float(
        value
        .replace('\xa0', '')  # неразрывный пробел
        .replace(' ', '')    # обычный пробел
        .replace(',', '.')   # запятая → точка
    )

# Названия листов из твоих справочников
SHEET_REF1 = "Справочник1"
SHEET_REF2 = "Справочник2"
SHEET_REF3 = "Справочник3"
SHEET_FACADE = "Фасады - Профили"

def get_gc(credentials_path):
    """Авторизация в Google Sheets."""
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    return gspread.authorize(creds)

def load_reference_1(spreadsheet_id, credentials_path):
    """Загрузка СПРАВОЧНИКА -1 (Материалы, артикулы, формулы)."""
    gc = get_gc(credentials_path)
    sh = gc.open_by_key(spreadsheet_id)
    worksheet = sh.worksheet(SHEET_REF1)
    data = worksheet.get_all_records()
    return data

def load_reference_2(spreadsheet_id, credentials_path):
    gc = get_gc(credentials_path)
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(SHEET_REF2)

    rows = ws.get_all_values()
    headers = rows[0]
    data = rows[1:]

    ref = {}

    def col(needle):
        needle = needle.lower().strip()
        for i, h in enumerate(headers):
            if needle in h.lower():
                return i
        raise ValueError(f"Колонка не найдена: {needle}")

    # стеклопакеты
    for row in data:
        name = row[col("тип стеклопакет")]
        price = row[col("стоимость стеклопакет")]
        if name and price:
            ref[name.strip()] = parse_price(price)

    # тонировка
    for row in data:
        if row[col("тониров")] == "Есть":
            ref["Тонировка"] = parse_price(row[col("стоимость тониров")])
            break

    # сборка
    for row in data:
        if row[col("сборк")] == "Есть":
            ref["Сборка"] = parse_price(row[col("стоимость сборк")])
            break

    # монтаж
    for row in data:
        name = row[col("монтаж")]
        price = row[col("стоимость монтаж")]
        if name and price:
            ref[name.strip()] = parse_price(price)

    print("DEBUG ref2:", ref)
    return ref

def load_reference_3(spreadsheet_id, credentials_path):
    """Загрузка СПРАВОЧНИКА -3 (Габаритные формулы и нарезка)."""
    gc = get_gc(credentials_path)
    sh = gc.open_by_key(spreadsheet_id)
    worksheet = sh.worksheet(SHEET_REF3)
    data = worksheet.get_all_records()
    return data

@st.cache_data(ttl=3600)
def load_facade_reference(spreadsheet_id: str, credentials_path: str) -> list:
    """
    Загрузка справочника профилей для фасадов из Google Sheets
    Лист: "Фасады - Профили"
    """
    try:
        gc = get_gc(credentials_path)
        sh = gc.open_by_key(spreadsheet_id)
        
        # Загружаем лист "Фасады - Профили"
        ws = sh.worksheet(SHEET_FACADE)
        data = ws.get_all_records()
        
        logger.info(f"✅ Загружено {len(data)} записей из справочника фасадов")
        return data
    except gspread.exceptions.WorksheetNotFound:
        logger.error(f"❌ Лист '{SHEET_FACADE}' не найден в Google Sheets!")
        st.error(f"Лист '{SHEET_FACADE}' не найден в Google Sheets!")
        return []
    except Exception as e:
        logger.error(f"❌ Ошибка загрузки справочника фасадов: {e}")
        st.error(f"Ошибка загрузки справочника фасадов: {e}")
        return []

def get_facade_data():
    """Обертка для удобного вызова из app.py"""
    from config.settings import SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH
    return load_facade_reference(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)

def save_request_to_sheet(spreadsheet_id, credentials_path, row_data):
    """Функция для записи данных в лист ЗАПРОСЫ (для будущего использования)."""
    gc = get_gc(credentials_path)
    sh = gc.open_by_key(spreadsheet_id)
    worksheet = sh.worksheet("ЗАПРОСЫ")
    worksheet.append_row(row_data)
