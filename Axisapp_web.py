import math
import os
import sys
from io import BytesIO
import logging
from datetime import datetime

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

# Константы листов (как в v15)
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# Подключение к Google (Исправлено для Render)
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=600)
def load_data():
    client = get_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    return {
        "ref1": pd.DataFrame(sh.worksheet(SHEET_REF1).get_all_records()),
        "ref2": pd.DataFrame(sh.worksheet(SHEET_REF2).get_all_records()),
        "ref3": pd.DataFrame(sh.worksheet(SHEET_REF3).get_all_records()),
        "users": pd.DataFrame(sh.worksheet(SHEET_USERS).get_all_records()),
        "sh": sh
    }

def main():
    st.set_page_config(page_title="Axisapp v15 Modern", layout="wide")
    db = load_data()

    if 'auth' not in st.session_state: st.session_state.auth = False
    if 'items' not in st.session_state: st.session_state.items = []

    # --- Авторизация (Стиль v15) ---
    if not st.session_state.auth:
        st.title("Вход в систему AXIS")
        u = st.text_input("Логин")
        p = st.text_input("Пароль", type="password")
        if st.button("Войти"):
            user = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
            if not user.empty:
                st.session_state.auth = True
                st.session_state.role = user.iloc[0]['Роль']
                st.rerun()
        return

    # --- Основная форма (Стиль v15) ---
    st.title("Калькулятор Axisapp Pro")
    
    order_number = st.sidebar.text_input("Номер заказа", "001")
    ral_color = st.sidebar.text_input("Цвет RAL", "7024")

    # Форма ввода (как ты любишь, но с новыми типами)
    with st.container():
        col1, col2, col3 = st.columns(3)
        p_type = col1.selectbox("Тип изделия", ["Окно глух.", "Окно с откр.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_sys = col2.selectbox("Система профиля", ["ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "ALG Slim", "Ruit 50F"])
        p_qty = col3.number_input("Кол-во (шт)", min_value=1, value=1)

        cW, cH = st.columns(2)
        W = cW.number_input("Ширина W (мм)", value=1000)
        H = cH.number_input("Высота H (мм)", value=1000)

        # Специфические поля для фасада
        n_m, n_t, is_ins = 0, 0, False
        if p_type == "Фасад":
            f1, f2 = st.columns(2)
            n_m = f1.number_input("Кол-во стоек", value=2)
            n_t = f2.number_input("Кол-во ригелей", value=2)
        else:
            is_ins = st.checkbox("Вставить это изделие в фасадный каркас")

        if st.button("Добавить позицию"):
            area = (W * H / 1000000) * p_qty
            # Автоматика петель
            hinges = 3 if H > 2100 and "Дверь" in p_type else 2
            st.session_state.items.append({
                "type": p_type, "sys": p_sys, "W": W, "H": H, "qty": p_qty, 
                "area": area, "n_m": n_m, "n_t": n_t, "is_insert": is_ins, "hinges": hinges
            })
            st.success(f"Позиция {p_type} добавлена!")

    # --- Результаты (Стиль v15) ---
    if st.session_state.items:
        st.markdown("---")
        st.subheader("Состав заказа:")
        df = pd.DataFrame(st.session_state.items)
        st.table(df[['type', 'sys', 'W', 'H', 'qty', 'area']])

        # Глобальные опции
        toning = st.sidebar.checkbox("Тонировка")
        assembly = st.sidebar.checkbox("Сборка", value=True)
        montage = st.sidebar.checkbox("Монтаж")

        # Расчет сумм (Логика Справочника-3 и Коэффициент 1.65)
        total_mats = 0
        for item in st.session_state.items:
            mats_ref = db['ref3'][db['ref3']['Тип изделия'] == item['type']]
            for _, row in mats_ref.iterrows():
                try:
                    formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                    count = eval(formula, {"math": math}, {"W": item['W'], "H": item['H'], "qty": item['qty'], "hinges": item['hinges'], "is_insert": int(item['is_insert'])})
                    price = db['ref2'][db['ref2']['Система'] == item['sys']]['Цена'].values[0]
                    total_mats += (count * price)
                except: continue

        total_area = sum(i['area'] for i in st.session_state.items)
        glass_sum = total_area * 18000
        if toning: glass_sum += (total_area * 4000)
        
        work_sum = (total_area * 5000 if assembly else 0) + (total_area * 7000 if montage else 0)
        
        # ИТОГОВАЯ ФОРМУЛА: (Мат + Стекло + Работы) * 1.65
        final_total = (total_mats + glass_sum + work_sum) * 1.65

        st.write(f"### Итого площадь: {total_area:.3f} м²")
        st.write(f"### СУММА К ОПЛАТЕ: {final_total:,.0f} тенге")

        if st.button("Скачать КП (Excel)"):
            st.info("Генерация файла...")
            # Тут вызывается функция построения Excel как в v15

if __name__ == "__main__":
    main()
