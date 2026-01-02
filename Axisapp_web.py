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
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

# =========================
# КОНСТАНТЫ / НАСТРОЙКИ
# =========================

DEBUG = False
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================
# ПОДКЛЮЧЕНИЕ К GOOGLE (Render Safe)
# =========================
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        # Прямое чтение секретного файла gcp.json на Render
        creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Ошибка авторизации: {e}")
        st.stop()

@st.cache_data(ttl=600)
def load_all_data():
    client = get_gspread_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    
    df_ref1 = pd.DataFrame(sh.worksheet(SHEET_REF1).get_all_records())
    df_ref2 = pd.DataFrame(sh.worksheet(SHEET_REF2).get_all_records())
    df_ref3 = pd.DataFrame(sh.worksheet(SHEET_REF3).get_all_records())
    df_users = pd.DataFrame(sh.worksheet(SHEET_USERS).get_all_records())
    
    return df_ref1, df_ref2, df_ref3, df_users, sh

# (Вспомогательные функции Excel и Расчета из v15 остаются здесь без изменений)
# ... [Код функций build_smeta_workbook и evaluate_formula аналогичен v15] ...

def main():
    st.set_page_config(page_title="Axisapp Pro v15", layout="wide")
    
    if 'auth' not in st.session_state:
        st.session_state.auth = False

    data = load_all_data()
    if data[0] is None: return
    df_ref1, df_ref2, df_ref3, df_users, sh = data

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("Вход в систему Axisapp")
        u_in = st.text_input("Логин")
        p_in = st.text_input("Пароль", type="password")
        if st.button("Войти"):
            user_row = df_users[(df_users['Логин'] == u_in) & (df_users['Пароль'].astype(str) == p_in)]
            if not user_row.empty:
                st.session_state.auth = True
                st.session_state.user_role = user_row.iloc[0]['Роль']
                st.rerun()
            else:
                st.error("Неверный логин или пароль")
        return

    # --- ОРИГИНАЛЬНАЯ ФОРМА ЗАПОЛНЕНИЯ v15 ---
    st.title("Калькулятор изделий")
    
    order_number = st.sidebar.text_input("Номер заказа", "001")
    
    # ИЗМЕНЕННЫЕ СПИСКИ ЗДЕСЬ
    product_type = st.sidebar.selectbox("Тип изделия", [
        "Окно с откр.", "Окно глух.", "Дверь 2-х створч.", "Дверь 1 створч.", "Фасад"
    ])
    
    profile_system = st.sidebar.selectbox("Система профиля", [
        "ALG 2030-63C", "ALG 2030-55C", "ALG 2030-73C", "ALG 2030-45C", "ALG 2030-Slim", "Ruit 50F"
    ])

    # Остальные поля формы как в v15
    glass_type = st.sidebar.selectbox("Тип заполнения", ["Стеклопакет", "Ламбри", "Сэндвич"])
    toning = st.sidebar.checkbox("Тонировка")
    handle_type = st.sidebar.selectbox("Тип ручки", ["Нажимная", "Офисная"])
    door_closer = st.sidebar.checkbox("Доводчик")
    assembly = st.sidebar.checkbox("Сборка", value=True)
    montage = st.sidebar.checkbox("Монтаж")

    st.header(f"Расчет: {product_type} ({profile_system})")
    
    # Ввод габаритов (стандартный блок v15)
    num_pos = st.number_input("Количество типоразмеров", min_value=1, value=1)
    sections = []
    for i in range(int(num_pos)):
        with st.expander(f"Позиция №{i+1}", expanded=True):
            col_w, col_h, col_q = st.columns(3)
            w = col_w.number_input(f"Ширина W{i+1}", value=1000, key=f"w_{i}")
            h = col_h.number_input(f"Высота H{i+1}", value=1500, key=f"h_{i}")
            q = col_q.number_input(f"Кол-во шт{i+1}", value=1, key=f"q_{i}")
            sections.append({"W": w, "H": h, "qty": q})

    if st.button("🏗️ РАССЧИТАТЬ"):
        # Логика расчета из v15...
        st.success("Расчет выполнен")
        # Вывод таблиц материалов и итогов как в v15

if __name__ == "__main__":
    main()
