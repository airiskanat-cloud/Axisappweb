import math
import os
import sys
import logging
from io import BytesIO
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook

# =========================================================
# 1. СЕРВИСНЫЕ НАСТРОЙКИ (БЕЗ SECRETS.TOML)
# =========================================================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        # Прямое чтение gcp.json на Render
        creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Критическая ошибка доступа: {e}")
        st.stop()

@st.cache_data(ttl=600)
def load_all_data():
    client = get_gspread_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    return {
        "ref1": pd.DataFrame(sh.worksheet("СПРАВОЧНИК -1").get_all_records()),
        "ref2": pd.DataFrame(sh.worksheet("СПРАВОЧНИК -2").get_all_records()),
        "ref3": pd.DataFrame(sh.worksheet("СПРАВОЧНИК -3").get_all_records()),
        "users": pd.DataFrame(sh.worksheet("ПОЛЬЗОВАТЕЛИ").get_all_records())
    }

# =========================================================
# 2. МАТЕМАТИЧЕСКАЯ ЛОГИКА (ИЗ ВЕРСИИ 15)
# =========================================================
def safe_eval(expr, context):
    try:
        # Очистка формулы из Справочника-3 для Python
        clean_expr = expr.replace('^', '**').replace('=', '')
        return eval(clean_expr, {"__builtins__": None, "math": math}, context)
    except:
        return 0

# =========================================================
# 3. ИНТЕРФЕЙС
# =========================================================
def main():
    st.set_page_config(page_title="Axisapp | Калькулятор систем", layout="wide")
    
    # Custom CSS для эстетики
    st.markdown("""
        <style>
        .main { background-color: #f8f9fa; }
        .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
        .metric-card { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        </style>
    """, unsafe_allow_html=True)

    data = load_all_data()
    
    if 'auth' not in st.session_state: st.session_state.auth = False
    if 'cart' not in st.session_state: st.session_state.cart = []

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("🔐 Вход в Axisapp")
        with st.container():
            u = st.text_input("Логин")
            p = st.text_input("Пароль", type="password")
            if st.button("Войти"):
                user = data['users'][(data['users']['Логин'] == u) & (data['users']['Пароль'].astype(str) == p)]
                if not user.empty:
                    st.session_state.auth = True
                    st.session_state.user_role = user.iloc[0]['Роль']
                    st.rerun()
        return

    # --- РАБОЧАЯ ПАНЕЛЬ ---
    st.sidebar.title("💎 Управление")
    order_number = st.sidebar.text_input("Заказ №", "NEW-001")
    
    tab1, tab2, tab3 = st.tabs(["🏗️ Добавить изделие", "📋 Спецификация", "💰 Итоговая смета"])

    with tab1:
        st.subheader("Настройка параметров")
        c1, c2, c3 = st.columns([2, 2, 1])
        p_type = c1.selectbox("Тип изделия", ["Окно глух.", "Окно с откр.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад (Каркас)"])
        p_sys = c2.selectbox("Система профиля", ["ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "ALG Slim", "Ruit 50F"])
        p_qty = c3.number_input("Кол-во", min_value=1, value=1)

        col_w, col_h = st.columns(2)
        W = col_w.number_input("Ширина (мм)", min_value=100, value=1000)
        H = col_h.number_input("Высота (мм)", min_value=100, value=1500)

        # Доп параметры для дверей/фасадов
        if "Дверь" in p_type:
            hinges = 3 if H > 2100 else 2
            st.info(f"Автоматически рассчитано: {hinges} петли (т.к. H={H}мм)")
        
        if p_type == "Фасад (Каркас)":
            f1, f2 = st.columns(2)
            n_m = f1.number_input("Стоек", value=2)
            n_t = f2.number_input("Ярусов ригеля", value=2)
            fill_m = st.radio("Заполнение", ["Стеклопакет", "Ламбри"])

        if st.button("✨ Добавить в заказ"):
            item = {
                "type": p_type, "sys": p_sys, "qty": p_qty, "W": W, "H": H,
                "area": (W * H / 1000000) * p_qty,
                "perim": ((W + H) * 2 / 1000) * p_qty,
                "n_m": locals().get('n_m', 0), "n_t": locals().get('n_t', 0),
                "fill": locals().get('fill_m', "Стеклопакет")
            }
            st.session_state.cart.append(item)
            st.toast(f"Добавлено: {p_type}")

    with tab2:
        if not st.session_state.cart:
            st.info("Заказ пуст")
        else:
            df_cart = pd.DataFrame(st.session_state.cart)
            st.table(df_cart[['type', 'sys', 'W', 'H', 'qty', 'area']])
            if st.button("🗑️ Очистить корзину"):
                st.session_state.cart = []
                st.rerun()

    with tab3:
        if st.session_state.cart:
            total_area = sum(i['area'] for i in st.session_state.cart)
            total_perim = sum(i['perim'] for i in st.session_state.cart)
            
            # РАСЧЕТ МАТЕРИАЛОВ (ПО ЛОГИКЕ СПРАВОЧНИКА-3)
            # Здесь происходит магия: мы объединяем цены из ref2 и формулы из ref3
            total_mats_cost = 0
            for item in st.session_state.cart:
                # Фильтруем справочник материалов для конкретного типа
                item_mats = data['ref3'][data['ref3']['Тип изделия'] == item['type']]
                for _, row in item_mats.iterrows():
                    formula = str(row['Формула_Python'])
                    count = safe_eval(formula, {"W": item['W'], "H": item['H'], "qty": item['qty'], "n_m": item['n_m'], "n_t": item['n_t']})
                    # Ищем цену в REF2 для этой системы
                    price_row = data['ref2'][data['ref2']['Система'] == item['sys']]
                    price = price_row['Цена'].values[0] if not price_row.empty else 1000
                    total_mats_cost += count * price

            # Глобальные опции
            toning = st.sidebar.checkbox("Тонировка")
            assembly = st.sidebar.checkbox("Сборка", value=True)
            montage = st.sidebar.checkbox("Монтаж", value=True)

            # --- ЭКОНОМИЧЕСКАЯ ПАНЕЛЬ ---
            st.subheader("Финансовый результат")
            
            glass_sum = total_area * 16000 # Пример цены стеклопакета
            work_sum = (total_area * 5000 if assembly else 0) + (total_area * 7000 if montage else 0)
            
            # ФОРМУЛА: (МАТЕРИАЛЫ + СТЕКЛО + РАБОТЫ) * 1.65
            subtotal = total_mats_cost + glass_sum + work_sum
            final_price = subtotal * 1.65

            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f"<div class='metric-card'><b>Общая площадь</b><br><h2>{total_area:.2f} м²</h2></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='metric-card'><b>Общий периметр</b><br><h2>{total_perim:.2f} м.п.</h2></div>", unsafe_allow_html=True)
            with c3:
                st.markdown(f"<div class='metric-card' style='background:#e3f2fd'><b>ИТОГО С НДС (1.65)</b><br><h2>{final_price:,.0f} ₸</h2></div>", unsafe_allow_html=True)

            # Кнопка экспорта
            if st.button("📥 Сформировать Excel Смету"):
                st.write("Смета генерируется...")

    # --- SIDEBAR ВЫХОД ---
    if st.sidebar.button("🚪 Выйти"):
        st.session_state.auth = False
        st.rerun()

if __name__ == "__main__":
    main()
