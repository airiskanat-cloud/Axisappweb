import math
import os
import sys
import shutil
from io import BytesIO
import logging
import json
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook

# =========================
# КОНСТАНТЫ / НАСТРОЙКИ
# =========================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Листы Google Таблицы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================
# ПОДКЛЮЧЕНИЕ К GOOGLE
# =========================
def get_gspread_client():
    creds_json = st.secrets.get("gcp_service_account")
    if not creds_json:
        st.error("Ключ gcp_service_account не найден в secrets!")
        st.stop()
    info = json.loads(creds_json)
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(info, scopes=scope)
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
        st.error(f"Ошибка загрузки данных: {e}")
        return None, None, None, None

# =========================
# ФУНКЦИЯ ЭКСПОРТА EXCEL
# =========================
def build_excel_kp(order_info, results_df, total_sum):
    wb = Workbook()
    ws = wb.active
    ws.title = "КП"
    ws.append(["КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", "", f"Заказ: {order_info['order_no']}"])
    ws.append(["Тип изделия:", order_info['p_type']])
    ws.append(["Система:", order_info['p_system']])
    ws.append([])
    # Заголовки таблицы
    ws.append(["Наименование", "Расход", "Ед. изм.", "Цена (ед)", "Сумма"])
    for row in results_df.itertuples():
        ws.append([row.Наименование, row.Расход, row.Ед, row.Цена, row.Итого])
    ws.append([])
    ws.append(["ИТОГО К ОПЛАТЕ:", "", "", "", total_sum])
    
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================
# ОСНОВНОЕ ПРИЛОЖЕНИЕ
# =========================
def main():
    st.set_page_config(page_title="Axisapp Pro", layout="wide")
    
    if 'auth' not in st.session_state:
        st.session_state.auth = False

    df_ref1, df_ref2, df_ref3, df_users = load_all_data()
    if df_ref1 is None: return

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("Вход в Axisapp")
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
                st.error("Неверные учетные данные")
        return

    # --- SIDEBAR ---
    st.sidebar.header(f"Пользователь: {st.session_state.user_role}")
    if st.sidebar.button("Выйти"):
        st.session_state.auth = False
        st.rerun()

    order_no = st.sidebar.text_input("Заказ №", "2024-001")
    p_type = st.sidebar.selectbox("Тип изделия", ["Окно глух.", "Окно с откр.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
    p_system = st.sidebar.selectbox("Система профиля", ["ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "ALG 2030-Slim", "Ruit 50F"])
    
    st.sidebar.markdown("---")
    color = st.sidebar.selectbox("Цвет RAL", ["7024", "9016", "9005"])
    glass_th = st.sidebar.selectbox("Толщина пакета (мм)", [4, 10, 24, 32, 40])
    toning = st.sidebar.selectbox("Тонировка", ["Нет", "Bronze", "Silver", "Grey"])
    
    assembly = st.sidebar.checkbox("Сборка", value=True)
    montage = st.sidebar.checkbox("Монтаж", value=False)

    # --- ФОРМА ВВОДА ---
    st.header(f"Параметры: {p_type} / {p_system}")
    positions = []

    if p_type == "Фасад":
        col1, col2, col3 = st.columns(3)
        with col1:
            W_f = st.number_input("Ширина фасада (мм)", min_value=100, value=1500)
            H_f = st.number_input("Высота фасада (мм)", min_value=100, value=3000)
        with col2:
            n_m = st.number_input("Кол-во стоек", min_value=2, value=2)
            n_t = st.number_input("Кол-во уровней ригелей", min_value=1, value=2)
        with col3:
            f_fill = st.selectbox("Заполнение глухих частей", ["Стеклопакет", "Ламбри (Панель)"])
        positions.append({"W": W_f, "H": H_f, "n_m": n_m, "n_t": n_t, "fill": f_fill, "qty": 1})
    else:
        num_items = st.number_input("Кол-во типоразмеров", min_value=1, value=1)
        for i in range(int(num_items)):
            c1, c2, c3 = st.columns(3)
            w = c1.number_input(f"Ширина W{i+1}", min_value=100, value=1000, key=f"w{i}")
            h = c2.number_input(f"Высота H{i+1}", min_value=100, value=1000, key=f"h{i}")
            q = c3.number_input(f"Кол-во шт{i+1}", min_value=1, value=1, key=f"q{i}")
            positions.append({"W": w, "H": h, "qty": q})

    # --- РАСЧЕТ МАТЕРИАЛОВ ---
    if st.button("📊 РАССЧИТАТЬ"):
        final_mats = []
        total_price = 0
        total_area = 0

        for pos in positions:
            area = (pos['W'] * pos['H'] / 1000000) * pos['qty']
            total_area += area
            
            # 1. Фильтрация справочника по Серии и Типу
            mask = (df_ref1['Тип изделия'] == p_type) & (df_ref1['Система профиля'] == p_system)
            current_ref = df_ref1[mask]

            # 2. Логика для Фасада
            if p_type == "Фасад":
                m_len = (pos['H'] * pos['n_m'] / 1000)
                t_count = (pos['n_m'] - 1) * pos['n_t']
                t_len_total = ((pos['W'] - (pos['n_m'] * 50)) / 1000) * pos['n_t']
                
                final_mats.append({"Наименование": "Профиль стойки", "Расход": m_len, "Ед": "м.п."})
                final_mats.append({"Наименование": "Профиль ригеля", "Расход": t_len_total, "Ед": "м.п."})
                final_mats.append({"Наименование": "U-соединитель", "Расход": t_count * 2, "Ед": "шт"})
                final_mats.append({"Наименование": "Упл. торцевой", "Расход": t_count * 2, "Ед": "шт"})
                
                if pos['fill'] == "Ламбри (Панель)":
                    final_mats.append({"Наименование": "Панель Ламбри", "Расход": area, "Ед": "м2"})
            
            # 3. Логика для Дверей (Авто-петли)
            elif "Дверь" in p_type:
                perim = ((pos['W'] + pos['H']) * 2 / 1000) * pos['qty']
                hinges = 3 if pos['H'] > 2100 else 2
                final_mats.append({"Наименование": "Профиль рамы/створки", "Расход": perim, "Ед": "м.п."})
                final_mats.append({"Наименование": "Петля дверная", "Расход": hinges * pos['qty'], "Ед": "шт"})
                final_mats.append({"Наименование": "Замок KALE", "Расход": 1 * pos['qty'], "Ед": "шт"})
            
            # Подтягиваем цены (из REF2/REF3) - упрощенная логика для примера
            # В реальном коде здесь поиск цены по Артикулу в REF3
            base_m2_cost = 55000 
            item_sum = area * base_m2_cost
            
            if toning != "Нет": item_sum *= 1.15
            if assembly: item_sum += (area * 3500)
            if montage: item_sum += (area * 6000)
            
            total_price += item_sum

        # Отрисовка результатов
        st.subheader("Итоговая спецификация")
        res_df = pd.DataFrame(final_mats)
        # Добавляем фиктивные колонки цены для наглядности
        res_df['Цена'] = 1500
        res_df['Итого'] = res_df['Расход'] * res_df['Цена']
        st.table(res_df)
        
        st.metric("ИТОГО К ОПЛАТЕ (тенге)", f"{total_price:,.2f}")
        st.write(f"Общая площадь остекления: {total_area:.2f} м²")

        # --- EXCEL ---
        order_info = {'order_no': order_no, 'p_type': p_type, 'p_system': p_system}
        kp_file = build_excel_kp(order_info, res_df, total_price)
        st.download_button("📥 Скачать КП (Excel)", data=kp_file, file_name=f"KP_{order_no}.xlsx")

if __name__ == "__main__":
    main()
