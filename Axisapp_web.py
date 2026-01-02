# =========================================================
# 1. СИСТЕМНЫЙ БЛОК (ВЫПОЛНЯЕТСЯ ПЕРВЫМ)
# =========================================================
import os
import shutil

# Пути для Render
SOURCE = "/etc/secrets/secrets.toml"
TARGET_DIR = "/opt/render/project/src/.streamlit"
TARGET = f"{TARGET_DIR}/secrets.toml"

if os.path.exists(SOURCE):
    os.makedirs(TARGET_DIR, exist_ok=True)
    shutil.copyfile(SOURCE, TARGET)
    # Файл скопирован до инициализации Streamlit

# =========================================================
# 2. ИМПОРТЫ
# =========================================================
import math
import sys
import logging
import json
from io import BytesIO

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook

# =========================================================
# 3. КОНСТАНТЫ И НАСТРОЙКИ
# =========================================================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Листы Google Таблицы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================================================
# 4. ИСПРАВЛЕННАЯ ФУНКЦИЯ ПОДКЛЮЧЕНИЯ
# =========================================================
def get_gspread_client():
    # Проверка наличия ключа в secrets
    if "gcp_service_account" not in st.secrets:
        st.error("Критическая ошибка: Ключ 'gcp_service_account' не найден в secrets.toml!")
        st.stop()
        
    # st.secrets["gcp_service_account"] в Streamlit уже является словарем (AttrDict)
    creds_info = st.secrets["gcp_service_account"]

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    # Используем словарь напрямую без json.loads
    creds = Credentials.from_service_account_info(
        creds_info,
        scopes=scope
    )

    return gspread.authorize(creds)

@st.cache_data(ttl=600)
def load_all_data():
    try:
        client = get_gspread_client()
        sh = client.open_by_key(GSPREAD_SHEET_ID)
        df_ref1 = pd.DataFrame(sh.worksheet(SHEET_REF1).get_all_records())
        df_ref2 = pd.DataFrame(sh.worksheet(SHEET_REF2).get_all_records())
        df_ref3 = pd.DataFrame(sh.worksheet(SHEET_REF3).get_all_records())
        df_users = pd.DataFrame(sh.worksheet(SHEET_USERS).get_all_records())
        return df_ref1, df_ref2, df_ref3, df_users
    except Exception as e:
        st.error(f"Ошибка загрузки данных из Google Sheets: {e}")
        return None, None, None, None

# =========================================================
# 5. ФУНКЦИИ ЭКСПОРТА И БИЗНЕС-ЛОГИКА
# =========================================================
def build_excel_kp(order_info, results_df, total_sum):
    wb = Workbook()
    ws = wb.active
    ws.title = "КП"
    ws.append(["КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", "", f"Заказ: {order_info['order_no']}"])
    ws.append(["Тип изделия:", order_info['p_type']])
    ws.append(["Система:", order_info['p_system']])
    ws.append([])
    ws.append(["Наименование", "Расход", "Ед. изм.", "Цена (ед)", "Сумма"])
    for row in results_df.itertuples():
        ws.append([row.Наименование, row.Расход, row.Ед, row.Цена, row.Итого])
    ws.append([])
    ws.append(["ИТОГО К ОПЛАТЕ:", "", "", "", total_sum])
    
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 6. ОСНОВНОЕ ПРИЛОЖЕНИЕ
# =========================================================
def main():
    st.set_page_config(page_title="Axisapp Pro", layout="wide")
    
    # ТЕСТОВАЯ ПРОВЕРКА (УДАЛИТЬ ПОСЛЕ УСПЕШНОГО ЗАПУСКА)
    if os.path.exists(TARGET):
        st.sidebar.success("✅ Файл secrets.toml обнаружен")
    
    if 'auth' not in st.session_state:
        st.session_state.auth = False

    # Загрузка данных
    df_ref1, df_ref2, df_ref3, df_users = load_all_data()
    if df_ref1 is None: return

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("Вход в систему Axisapp")
        l_col, r_col = st.columns(2)
        u_in = l_col.text_input("Логин")
        p_in = l_col.text_input("Пароль", type="password")
        if l_col.button("Войти"):
            user = df_users[(df_users['Логин'] == u_in) & (df_users['Пароль'].astype(str) == p_in)]
            if not user.empty:
                st.session_state.auth = True
                st.session_state.user_role = user.iloc[0]['Роль']
                st.rerun()
            else:
                st.error("Неверный логин или пароль")
        return

    # --- ИНТЕРФЕЙС УПРАВЛЕНИЯ ---
    st.sidebar.header(f"Роль: {st.session_state.user_role}")
    if st.sidebar.button("Выход"):
        st.session_state.auth = False
        st.rerun()

    order_no = st.sidebar.text_input("Заказ №", "2024-001")
    p_type = st.sidebar.selectbox("Тип изделия", ["Окно глух.", "Окно с откр.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
    p_system = st.sidebar.selectbox("Система профиля", ["ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "ALG 2030-Slim", "Ruit 50F"])
    
    color = st.sidebar.selectbox("Цвет RAL", ["7024", "9016", "9005"])
    toning = st.sidebar.selectbox("Тонировка", ["Нет", "Bronze", "Silver", "Grey"])
    assembly = st.sidebar.checkbox("Сборка", value=True)
    montage = st.sidebar.checkbox("Монтаж", value=False)

    # --- ВВОД ГАБАРИТОВ ---
    st.header(f"Расчет: {p_type}")
    positions = []

    if p_type == "Фасад":
        c1, c2, c3 = st.columns(3)
        W_f = c1.number_input("Ширина фасада (мм)", min_value=100, value=1500)
        H_f = c2.number_input("Высота фасада (мм)", min_value=100, value=3000)
        n_m = c1.number_input("Кол-во стоек", min_value=2, value=2)
        n_t = c2.number_input("Кол-во уровней ригелей", min_value=1, value=2)
        f_fill = c3.selectbox("Заполнение", ["Стеклопакет", "Ламбри (Панель)"])
        positions.append({"W": W_f, "H": H_f, "n_m": n_m, "n_t": n_t, "fill": f_fill, "qty": 1})
    else:
        num_items = st.number_input("Кол-во типоразмеров", min_value=1, value=1)
        for i in range(int(num_items)):
            c1, c2, c3 = st.columns(3)
            w = c1.number_input(f"Ширина W{i+1}", min_value=100, value=1000, key=f"w{i}")
            h = c2.number_input(f"Высота H{i+1}", min_value=100, value=1000, key=f"h{i}")
            q = c3.number_input(f"Кол-во шт{i+1}", min_value=1, value=1, key=f"q{i}")
            positions.append({"W": w, "H": h, "qty": q})

    # --- РАСЧЕТ И ВЫВОД ---
    if st.button("📊 РАССЧИТАТЬ"):
        final_mats = []
        total_price = 0
        
        for pos in positions:
            area = (pos['W'] * pos['H'] / 1000000) * pos['qty']
            
            if p_type == "Фасад":
                m_len = (pos['H'] * pos['n_m'] / 1000)
                t_count = (pos['n_m'] - 1) * pos['n_t']
                t_len_total = ((pos['W'] - (pos['n_m'] * 50)) / 1000) * pos['n_t']
                final_mats.append({"Наименование": "Профиль стойки", "Расход": m_len, "Ед": "м.п.", "Цена": 5000, "Итого": m_len*5000})
                final_mats.append({"Наименование": "Профиль ригеля", "Расход": t_len_total, "Ед": "м.п.", "Цена": 4500, "Итого": t_len_total*4500})
            elif "Дверь" in p_type:
                hinges = 3 if pos['H'] > 2100 else 2
                final_mats.append({"Наименование": "Петля дверная", "Расход": hinges * pos['qty'], "Ед": "шт", "Цена": 2500, "Итого": hinges*pos['qty']*2500})
            
            # Пример базовой стоимости для КП
            total_price += area * 65000 

        res_df = pd.DataFrame(final_mats)
        st.table(res_df)
        st.metric("ИТОГО К ОПЛАТЕ", f"{total_price:,.2f} тенге")

        # ЭКСПОРТ
        order_info = {'order_no': order_no, 'p_type': p_type, 'p_system': p_system}
        kp_file = build_excel_kp(order_info, res_df, total_price)
        st.download_button("📥 Скачать КП", data=kp_file, file_name=f"KP_{order_no}.xlsx")

if __name__ == "__main__":
    main()
