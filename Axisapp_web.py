import math
import os
import sys
from io import BytesIO
import logging
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from datetime import datetime

# Настройки остаются из v15
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=600)
def load_data():
    client = get_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    return {
        "ref1": pd.DataFrame(sh.worksheet("СПРАВОЧНИК -1").get_all_records()),
        "ref2": pd.DataFrame(sh.worksheet("СПРАВОЧНИК -2").get_all_records()),
        "ref3": pd.DataFrame(sh.worksheet("СПРАВОЧНИК -3").get_all_records()),
        "users": pd.DataFrame(sh.worksheet("ПОЛЬЗОВАТЕЛИ").get_all_records()),
        "sh": sh
    }

def main():
    st.set_page_config(page_title="Axisapp v16 Pro", layout="wide")
    db = load_data()

    if 'auth' not in st.session_state: st.session_state.auth = False
    if 'items' not in st.session_state: st.session_state.items = []

    # --- Авторизация v15 ---
    if not st.session_state.auth:
        st.title("Вход в Axisapp")
        u, p = st.text_input("Логин"), st.text_input("Пароль", type="password")
        if st.button("Войти"):
            user = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
            if not user.empty:
                st.session_state.auth, st.session_state.role = True, user.iloc[0]['Роль']
                st.rerun()
        return

    st.title("Инженерный калькулятор AXIS")
    
    # 1. ОБНОВЛЕННЫЕ СПИСКИ (как ты просила)
    product_type = st.sidebar.selectbox("Тип изделия", 
        ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
    
    profile_system = st.sidebar.selectbox("Система профиля", 
        ["ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "ALG 2030-Slim", "Ruit 50F"])

    # 2. ПАРАМЕТРЫ ФОРМЫ (v15)
    order_number = st.sidebar.text_input("Заказ №", "001")
    toning = st.sidebar.checkbox("Тонировка")
    assembly = st.sidebar.checkbox("Сборка", value=True)
    montage = st.sidebar.checkbox("Монтаж")

    # 3. ВВОД ГАБАРИТОВ (С восстановлением полей импоста/стоек)
    st.subheader(f"Ввод параметров для: {product_type}")
    colW, colH, colQ = st.columns(3)
    W = colW.number_input("Ширина (мм)", value=1000)
    H = colH.number_input("Высота (мм)", value=1500)
    qty = colQ.number_input("Кол-во (шт)", value=1)

    # Динамические поля в зависимости от типа
    n_imp = 0
    if product_type == "Фасад":
        c1, c2 = st.columns(2)
        n_m = c1.number_input("Кол-во стоек (вертикальных)", value=2)
        n_t = c2.number_input("Кол-во ригелей (горизонтальных)", value=1)
        n_imp = n_m + n_t # Для расчёта Т-соединителей
    else:
        n_imp = st.number_input("Количество импостов (перегородок)", value=0)

    if st.button("🏗️ РАССЧИТАТЬ"):
        # Логика подбора данных из справочников
        mats_spec = []
        mats_total_cost = 0

        # Фильтруем Справочник-3 по типу изделия
        ref3_filtered = db['ref3'][db['ref3']['Тип изделия'] == product_type]
        
        # Контекст для формул (чтобы eval не ломался)
        context = {"W": W, "H": H, "qty": qty, "n_imp": n_imp, "math": math}

        for _, row in ref3_filtered.iterrows():
            try:
                # Считаем количество материала по формуле
                formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                q_mat = eval(formula, {"__builtins__": None}, context)
                
                if q_mat > 0:
                    # Ищем цену в Справочнике-2
                    price_row = db['ref2'][db['ref2']['Система'] == profile_system]
                    price = price_row['Цена'].values[0] if not price_row.empty else 0
                    
                    cost = q_mat * price
                    mats_total_cost += cost
                    mats_spec.append({"Название": row['Наименование'], "Кол-во": q_mat, "Ед": row['Ед'], "Сумма": cost})
            except: continue

        # Экономика: (Мат + Стекло + Работа) * 1.65
        area = (W * H / 1000000) * qty
        glass_sum = area * 18000 + (area * 5000 if toning else 0)
        labor_sum = area * (5000 if assembly else 0) + area * (7000 if montage else 0)
        
        final_total = (mats_total_cost + glass_sum + labor_sum) * 1.65

        # ВЫВОД РЕЗУЛЬТАТОВ (v15)
        st.markdown("---")
        st.success(f"Расчет позиции №{order_number} завершен!")
        
        c1, c2 = st.columns(2)
        c1.metric("Площадь изделия", f"{area:.3f} м2")
        c2.metric("ИТОГО К ОПЛАТЕ", f"{final_total:,.0f} ₸")

        with st.expander("Посмотреть ведомость материалов"):
            st.table(pd.DataFrame(mats_spec))

if __name__ == "__main__":
    main()
