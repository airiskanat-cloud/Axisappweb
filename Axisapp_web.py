import math
import os
import sys
import shutil
from io import BytesIO
import zipfile
import logging
import json
import ast
import operator as op

import streamlit as st
# openpyxl теперь нужен только для экспорта КП, но не для чтения справочников
from openpyxl.workbook import Workbook
from openpyxl.drawing.image import Image as XLImage

# --- НОВЫЕ ИМПОРТЫ ДЛЯ GOOGLE SHEETS ---
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
# ----------------------------------------

# =========================
# КОНСТАНТЫ / НАСТРОЙКИ (ОБНОВЛЕНО)
# =========================

DEBUG = False
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# --- УДАЛЕНИЕ ЛОКАЛЬНОЙ ЛОГИКИ ФАЙЛОВ ---
# resource_path теперь не используется для Excel/Session
def resource_path(relative_path: str) -> str:
    # Оставляем только для загрузки логотипа
    try:
        if hasattr(sys, "_MEIPASS"):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(os.path.dirname(__file__))
    except Exception:
        base_path = os.getcwd()
    return os.path.join(base_path, relative_path)

# УДАЛЕНО: DATA_DIR, EXCEL_FILE, SESSION_FILE, BUNDLED_TEMPLATE

# --- НОВЫЕ КОНСТАНТЫ GOOGLE SHEETS ---
# ВАЖНО: Мы используем ID вашей таблицы
GSPREAD_SHEET_ID = "1RJCkHf9qbjO0z3E2rdHQWAQyrGEHNL-W" 
# -------------------------------------

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# ... (Остальные константы FORM_HEADER, COMPANY_NAME и т.д. остаются без изменений) ...


# =========================
# УТИЛИТЫ (без изменений)
# =========================
# ... (normalize_key, _clean_cell_val, safe_float, safe_int, get_field остаются без изменений) ...


# =========================
# БЕЗОПАСНЫЙ EVAL (ФОРМУЛЫ) (без изменений)
# =========================
# ... (_allowed_ops, _eval_ast, safe_eval_formula остаются без изменений) ...


# =========================
# GOOGLE SHEETS CLIENT (ЗАМЕНА ExcelClient)
# =========================

class GoogleSheetsClient:
    def __init__(self, sheet_id: str):
        self.sheet_id = sheet_id
        self._worksheets_cache = {} 
        self.load()

    def _auth(self):
        # Аутентификация через JSON-ключ из переменной окружения/секрета Render
        gcp_keyfile_content = os.getenv("GCP_SA_KEYFILE")
        if not gcp_keyfile_content:
            st.error("Ошибка: Ключ сервисного аккаунта GCP_SA_KEYFILE не найден в секретах Render. Расчет невозможен.")
            st.stop()
            
        try:
            creds_data = json.loads(gcp_keyfile_content)
            creds = ServiceAccountCredentials.from_json_keyfile_dict(
                creds_data,
                scope=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            )
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"Ошибка аутентификации Google Sheets. Проверьте формат GCP_SA_KEYFILE. {e}")
            st.stop()

    def load(self):
        try:
            client = self._auth()
            self.wb = client.open_by_key(self.sheet_id)
            logger.info("Успешно подключен к Google Sheets.")
        except Exception as e:
            st.error(f"Критическая ошибка при подключении к Google Sheets. Проверьте ID и права доступа. {e}")
            st.stop()

    def ws(self, name: str):
        if name in self._worksheets_cache:
            return self._worksheets_cache[name]
        try:
            ws = self.wb.worksheet(name)
            self._worksheets_cache[name] = ws
            return ws
        except gspread.WorksheetNotFound:
            # Если это лист для записи (ЗАПРОСЫ), создаем его
            if name == SHEET_FORM:
                ws = self.wb.add_worksheet(name, rows="100", cols="30")
                self._worksheets_cache[name] = ws
                ws.append_row(FORM_HEADER)
                return ws
            
            # Для справочников - это ошибка
            st.error(f"Лист '{name}' не найден в Google Sheets. Проверьте название листа в таблице.")
            st.stop()

    def read_records(self, sheet_name: str):
        ws = self.ws(sheet_name)
        # Получаем все значения
        rows = ws.get_all_values()
        
        if not rows:
            return []
            
        header_raw = rows[0]
        header = []
        used = {}

        for h in header_raw:
            key = normalize_key(h)
            if key in used:
                used[key] += 1
                key = f"{key}_{used[key]}"
            else:
                used[key] = 1
            header.append(key)

        records = []
        for r in rows[1:]:
            # Пропускаем пустые строки
            if all(v is None or v == "" for v in r):
                continue
            row = {}
            for i, k in enumerate(header):
                if i < len(r):
                    # Важно: gspread возвращает все как строки, преобразование в float/int будет в safe_float
                    row[k] = r[i]
                else:
                    row[k] = None
            records.append(row)
        return records

    def clear_and_write(self, sheet_name: str, header: list, rows: list):
        # В облачном калькуляторе мы не записываем промежуточные расчеты
        # в Sheets, чтобы не замедлять работу и не перегружать API.
        # Они отображаются только пользователю в Streamlit.
        logger.warning(f"Расчеты для листа {sheet_name} (Габариты/Материалы/Итог) не сохраняются в Google Sheets.")
        pass

    def append_form_row(self, row: list):
        ws = self.ws(SHEET_FORM)
        # Добавление новой строки в лист ЗАПРОСЫ
        # value_input_option='USER_ENTERED' позволяет Sheets автоматически преобразовывать
        # числа и формулы, если это необходимо (хотя мы передаем только данные)
        ws.append_row(row, value_input_option='USER_ENTERED')
        logger.info("Строка успешно добавлена в лист ЗАПРОСЫ.")

# =========================
# ПОЛЬЗОВАТЕЛИ (ЛОГИН) (ОБНОВЛЕНО)
# =========================

def load_users(excel: GoogleSheetsClient):
    excel.load()
    # ... (логика load_users остается прежней)
    rows = excel.read_records(SHEET_USERS)
    # ...

    users = {}
    for r in rows:
        login = _clean_cell_val(get_field(r, "логин", "")).lower()
        # В Google Sheets пароли могут быть строками, поэтому убираем .replace("*", "")
        pwd = _clean_cell_val(get_field(r, "парол", "")).strip() 
        role = _clean_cell_val(get_field(r, "роль", ""))

        if login:
            users[login] = {"password": pwd, "role": role, "_raw_login": login}

    return users

def login_form(excel: GoogleSheetsClient):
    # УДАЛЕНО: Проверка и чтение/запись SESSION_FILE

    if "current_user" in st.session_state:
        return st.session_state["current_user"]

    st.sidebar.title("🔐 Вход в систему")
    with st.sidebar.form("login_form"):
        login = st.text_input("Логин")
        password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти")

    users = load_users(excel)

    if submitted:
        entered_login = (login or "").strip().lower()
        entered_pass = (password or "").replace("\xa0", "").strip()

        user = users.get(entered_login)

        if user:
            real_pass = (user["password"] or "").strip().replace("\xa0", "")
            if entered_pass == real_pass:
                st.session_state["current_user"] = {
                    "login": user["_raw_login"],
                    "role": user["role"],
                }
                # УДАЛЕНО: Сохранение в SESSION_FILE

                st.sidebar.success(f"Привет, {user['_raw_login']}!")
                return st.session_state["current_user"]

        st.sidebar.error("Неверный логин или пароль")

    return None

# =========================
# CALCULATORS (ИЗМЕНЕНИЯ МИНИМАЛЬНЫЕ)
# =========================

# Все классы (GabaritCalculator, MaterialCalculator, FinalCalculator)
# автоматически используют новый GoogleSheetsClient, потому что они ожидают
# методы read_records() и clear_and_write(), которые мы переопределили.

# ... (Код GabaritCalculator остается без изменений) ...
# ... (Код MaterialCalculator остается без изменений) ...
# ... (Код FinalCalculator остается без изменений) ...


# =========================
# STREAMLIT UI: main (ОБНОВЛЕНО)
# =========================

def main():
    st.set_page_config(page_title="Axis Pro GF • Калькулятор", layout="wide") 
    
    ensure_session_state()

    # --- Инициализация Google Sheets Client ---
    excel = GoogleSheetsClient(GSPREAD_SHEET_ID)
    # -----------------------------------------

    # --- УДАЛЕНА ЛОГИКА ЗАГРУЗКИ SESSION_FILE ---
    # if "current_user" not in st.session_state:
    #     try:
    #         if os.path.exists(SESSION_FILE):
    #             with open(SESSION_FILE, "r", encoding="utf-8") as sf:
    #                 st.session_state["current_user"] = json.load(sf)
    #         except Exception:
    #             pass
    # ---------------------------------------------
    
    user = login_form(excel)
    if not user:
        st.stop()

    st.title("📘 Калькулятор алюминиевых изделий (Axis Pro GF)")
    st.info(f"Пользователь: **{user['login']}**")

    # ... (Остальной код main() остается без изменений) ...
    
    # ---------- Кнопка выхода (ОБНОВЛЕНО) ----------
    if st.sidebar.button("Выйти"):
        st.session_state.pop("current_user", None)
        # УДАЛЕНО: Удаление SESSION_FILE
        st.experimental_rerun()


if __name__ == "__main__":
    main()
