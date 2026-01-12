import gspread
import streamlit as st
from google.oauth2.service_account import Credentials

# Мы меняем load_users на загрузку через общую функцию из sheets_reader
def load_users_from_sheet(credentials_path, spreadsheet_id):
    """Вспомогательная функция для получения списка пользователей из листа 'ПОЛЬЗОВАТЕЛИ'"""
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(spreadsheet_id)
        # Ищем лист с названием ПОЛЬЗОВАТЕЛИ
        worksheet = sh.worksheet("ПОЛЬЗОВАТЕЛИ")
        return worksheet.get_all_records()
    except Exception as e:
        print(f"Ошибка при загрузке пользователей: {e}")
        return []

def authenticate(login, password, credentials_path, spreadsheet_id):
    # 1. Загружаем данные
    users = load_users_from_sheet(credentials_path, spreadsheet_id)
    
    # Очищаем то, что ввел пользователь
    in_login = str(login).strip().lower()
    in_pass = str(password).strip()

    print("\n" + "="*40)
    print("🔎 ДИАГНОСТИКА ВХОДА")
    print(f"Вы вводите: логин ['{in_login}'], пароль ['{in_pass}']")
    
    if not users:
        print("❌ ОШИБКА: Лист 'ПОЛЬЗОВАТЕЛИ' пуст или не найден!")
        print("="*40 + "\n")
        return None

    print(f"Найдено строк в таблице: {len(users)}")
    
    for i, user in enumerate(users):
        # Берем значения напрямую
        values = list(user.values())
        if len(values) < 2:
            print(f"Строка {i+1}: Ошибка (мало колонок): {user}")
            continue
            
        db_login = str(values[0]).strip().lower()
        db_pass = str(values[1]).strip()
        
        print(f"Строка {i+1} в базе: логин ['{db_login}'], пароль ['{db_pass}']")

        if db_login == in_login and db_pass == in_pass:
            st.session_state["authenticated"] = True
            print("✅ УСПЕХ: Совпадение найдено!")
            print("="*40 + "\n")
            return {"login": db_login}
            
    print("❌ ИТОГ: Совпадений не найдено.")
    print("="*40 + "\n")
    return None
