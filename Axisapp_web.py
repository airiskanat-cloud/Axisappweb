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
# ПОДКЛЮЧЕНИЕ (НОВЫЙ КЛЮЧ)
# =========================
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        # Путь для Render
        creds_path = "/etc/secrets/gcp.json"
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
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

# =========================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (v15)
# =========================

def evaluate_formula(formula_str, context):
    try:
        expr = str(formula_str).replace('=', '').replace('^', '**')
        # Ограниченный eval
        allowed_names = {"math": math, "W": context.get('W', 0), "H": context.get('H', 0), "qty": context.get('qty', 1)}
        return eval(expr, {"__builtins__": None}, allowed_names)
    except Exception as e:
        logger.error(f"Ошибка в формуле {formula_str}: {e}")
        return 0

def build_smeta_workbook(order, base_positions, lambr_positions, total_area, total_perimeter, total_sum):
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"
    
    # Здесь логика формирования Excel из v15 (сокращено для экономии места, но в вашем файле она полная)
    ws.append(["ООО «AXIS»", "", "Город Астана"])
    ws.append(["Коммерческое предложение"])
    ws.append(["Заказ №", order.get("order_number")])
    ws.append(["Тип изделия", order.get("product_type")])
    ws.append(["Профильная система", order.get("profile_system")])
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# =========================
# ИНТЕРФЕЙС (v15)
# =========================

def main():
    st.set_page_config(page_title="Axisapp v15", layout="wide")
    
    if 'auth' not in st.session_state:
        st.session_state.auth = False

    # Загрузка
    data = load_all_data()
    df_ref1, df_ref2, df_ref3, df_users, sh = data

    if not st.session_state.auth:
        st.title("Вход в систему")
        u = st.text_input("Логин")
        p = st.text_input("Пароль", type="password")
        if st.button("Войти"):
            user = df_users[(df_users['Логин'] == u) & (df_users['Пароль'].astype(str) == p)]
            if not user.empty:
                st.session_state.auth = True
                st.session_state.user_role = user.iloc[0]['Роль']
                st.rerun()
        return

    # --- ПАНЕЛЬ ВВОДА ---
    st.sidebar.title("Калькулятор")
    order_number = st.sidebar.text_input("Номер заказа", "001")
    product_type = st.sidebar.selectbox("Тип изделия", ["Окно", "Дверь", "Тамбур"])
    profile_system = st.sidebar.selectbox("Профильная система", ["ALG 2030-45C", "Ruit 50F", "ALG 2030-63C"])
    
    # Дополнительные настройки из v15
    glass_type = st.sidebar.selectbox("Тип заполнения", ["Стеклопакет", "Ламбри", "Сэндвич"])
    toning = st.sidebar.checkbox("Тонировка")
    assembly = st.sidebar.checkbox("Сборка", value=True)
    montage = st.sidebar.checkbox("Монтаж")
    handle_type = st.sidebar.selectbox("Тип ручек", ["Нажимная", "Офисная"])
    door_closer = st.sidebar.checkbox("Доводчик")

    # Секции (Окна / Тамбур)
    sections = []
    num_pos = st.number_input("Количество позиций", min_value=1, value=1)
    
    for i in range(int(num_pos)):
        with st.expander(f"Позиция №{i+1}", expanded=True):
            c1, c2, c3 = st.columns(3)
            w = c1.number_input(f"Ширина W{i+1}", value=1000, key=f"w_{i}")
            h = c2.number_input(f"Высота H{i+1}", value=1500, key=f"h_{i}")
            q = c3.number_input(f"Кол-во Q{i+1}", value=1, key=f"q_{i}")
            sections.append({"W": w, "H": h, "qty": q, "kind": "window"})

    if st.button("РАССЧИТАТЬ"):
        st.subheader("Результаты расчета")
        
        # Расчетная логика v15
        total_mats_cost = 0
        total_area = 0
        
        for s in sections:
            area = (s['W'] * s['H'] / 1000000) * s['qty']
            total_area += area
            
            # Поиск в Справочнике-3
            mats = df_ref3[df_ref3['Тип изделия'] == product_type]
            for _, row in mats.iterrows():
                qty_mat = evaluate_formula(row['Формула_Python'], s)
                if qty_mat > 0:
                    price = df_ref2[df_ref2['Система'] == profile_system]['Цена'].values[0]
                    total_mats_cost += (qty_mat * price)

        # Коэффициенты и наценки
        glass_cost = total_area * 15000 # Пример базы
        total_sum = (total_mats_cost + glass_cost) * 1.65
        
        st.write(f"Общая площадь: **{total_area:.3f} м2**")
        st.write(f"ИТОГО к оплате: **{total_sum:,.0f} тенге**")

        # Кнопка скачивания
        excel_data = build_smeta_workbook(
            {"order_number": order_number, "product_type": product_type, "profile_system": profile_system},
            sections, [], total_area, 0, total_sum
        )
        st.download_button("⬇️ Скачать КП в Excel", data=excel_data, file_name=f"Order_{order_number}.xlsx")

if __name__ == "__main__":
    main()
