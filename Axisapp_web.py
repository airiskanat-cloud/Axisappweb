import math
import os
import sys
from io import BytesIO
import logging
import json
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook

# =========================================================
# 1. ПОДКЛЮЧЕНИЕ (ЧЕРЕЗ ФАЙЛ GCP.JSON НА RENDER)
# =========================================================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        # Используем путь Render для секретных файлов
        creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Ошибка ключа gcp.json: {e}")
        st.stop()

@st.cache_data(ttl=600)
def load_all_data():
    client = get_gspread_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    df_ref1 = pd.DataFrame(sh.worksheet("СПРАВОЧНИК -1").get_all_records())
    df_ref2 = pd.DataFrame(sh.worksheet("СПРАВОЧНИК -2").get_all_records())
    df_ref3 = pd.DataFrame(sh.worksheet("СПРАВОЧНИК -3").get_all_records())
    df_users = pd.DataFrame(sh.worksheet("ПОЛЬЗОВАТЕЛИ").get_all_records())
    return df_ref1, df_ref2, df_ref3, df_users

# =========================================================
# 2. ИНТЕРФЕЙС И ЛОГИКА
# =========================================================
def main():
    st.set_page_config(page_title="Axisapp Pro: Фасады и Витражи", layout="wide")
    
    if 'auth' not in st.session_state:
        st.session_state.auth = False

    df_ref1, df_ref2, df_ref3, df_users = load_all_data()

    if not st.session_state.auth:
        st.title("Вход в систему")
        u, p = st.sidebar.text_input("Логин"), st.sidebar.text_input("Пароль", type="password")
        if st.sidebar.button("Войти"):
            user = df_users[(df_users['Логин'] == u) & (df_users['Пароль'].astype(str) == p)]
            if not user.empty:
                st.session_state.auth = True
                st.session_state.user_role = user.iloc[0]['Роль']
                st.rerun()
        return

    st.sidebar.success(f"Роль: {st.session_state.user_role}")
    order_no = st.sidebar.text_input("Заказ №", "2024-100")

    # --- Сборка корзины изделий ---
    if 'cart' not in st.session_state:
        st.session_state.cart = []

    st.header("🏗️ Конструктор заказа")
    
    with st.expander("Добавить изделие в расчет", expanded=True):
        col1, col2, col3 = st.columns(3)
        p_type = col1.selectbox("Тип изделия", ["Фасад (Каркас)", "Окно глух.", "Окно с откр.", "Дверь 1 створч.", "Дверь 2-х створч."])
        p_sys = col2.selectbox("Серия профиля", ["Ruit 50F", "ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "ALG Slim"])
        qty = col3.number_input("Кол-во (шт)", min_value=1, value=1)
        
        w = st.number_input("Ширина (мм)", value=1000)
        h = st.number_input("Высота (мм)", value=1000)
        
        fill_opt = "Стеклопакет"
        if p_type == "Фасад (Каркас)":
            c_f1, c_f2 = st.columns(2)
            n_m = c_f1.number_input("Кол-во стоек", value=2)
            n_t = c_f2.number_input("Кол-во ярусов ригелей", value=1)
            fill_opt = st.radio("Заполнение глухих зон", ["Стеклопакет", "Ламбри (Панель)"])
        
        if st.button("➕ Добавить в спецификацию"):
            new_item = {
                "type": p_type, "sys": p_sys, "w": w, "h": h, "qty": qty, 
                "fill": fill_opt, "n_m": locals().get('n_m', 0), "n_t": locals().get('n_t', 0)
            }
            st.session_state.cart.append(new_item)

    # --- Итоговые параметры ---
    st.sidebar.subheader("Глобальные настройки")
    toning = st.sidebar.selectbox("Тонировка", ["Нет", "Bronze", "Silver", "Grey"])
    assembly = st.sidebar.checkbox("Сборка", value=True)
    montage = st.sidebar.checkbox("Монтаж", value=True)

    if st.session_state.cart:
        st.subheader("🛒 Состав заказа")
        total_area, total_perim, total_mats_cost = 0, 0, 0
        
        summary_data = []
        for i, item in enumerate(st.session_state.cart):
            area = (item['w'] * item['h'] / 1000000) * item['qty']
            perim = ((item['w'] + item['h']) * 2 / 1000) * item['qty']
            
            # --- РАСЧЕТ МАТЕРИАЛОВ (Инженерная логика) ---
            mats_price = 0
            if item['type'] == "Фасад (Каркас)":
                # Стойки + Ригели + Соединители
                mats_price = area * 35000 # Базовая цена каркаса за м2
            elif "Дверь" in item['type']:
                hinges = 3 if item['h'] > 2100 else 2
                mats_price = (area * 40000) + (hinges * 5000) # Профиль + петли
            else:
                mats_price = area * 25000

            total_area += area
            total_perim += perim
            total_mats_cost += mats_price
            
            summary_data.append([item['type'], item['sys'], f"{item['w']}x{item['h']}", item['qty'], f"{area:.2f}"])

        st.table(pd.DataFrame(summary_data, columns=["Тип", "Серия", "Размер", "Кол-во", "Площадь м2"]))

        # --- ИТОГОВЫЙ РАСЧЕТ (ВАША ФОРМУЛА) ---
        glass_price = total_area * 15000
        toning_price = (total_area * 3500) if toning != "Нет" else 0
        assembly_price = (total_area * 4000) if assembly else 0
        montage_price = (total_area * 6000) if montage else 0
        
        # (Материалы + Стекло + Допы) * 1.65
        subtotal = total_mats_cost + glass_price + toning_price + assembly_price + montage_price
        final_total = subtotal * 1.65

        # --- ВЫВОД ГАБАРИТОВ И СУММ ---
        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("Общая площадь", f"{total_area:.2f} м²")
        c2.metric("Общий периметр", f"{total_perim:.2f} м.п.")
        c3.metric("ИТОГО К ОПЛАТЕ", f"{final_total:,.2f} ₸")

        with st.expander("Детализация расчета"):
            st.write(f"Стоимость материалов: {total_mats_cost:,.2f}")
            st.write(f"Стоимость стеклопакетов: {glass_price:,.2f}")
            st.write(f"Обеспечение (65%): {(final_total - subtotal):,.2f}")

        if st.button("🗑️ Очистить всё"):
            st.session_state.cart = []
            st.rerun()

if __name__ == "__main__":
    main()
