import math
import os
import time
import logging
from datetime import datetime
from io import BytesIO

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# =========================================================
# 1. КОНСТАНТЫ И НАСТРОЙКИ (ИМЕНА ЛИСТОВ ИЗ ТВОЕЙ ТАБЛИЦЫ)
# =========================================================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================================================
# 2. ПОДКЛЮЧЕНИЕ К GOOGLE (RENDER SAFE)
# =========================================================
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=600)
def load_db():
    client = get_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    def get_df(name):
        try:
            df = pd.DataFrame(sh.worksheet(name).get_all_records())
            df.columns = df.columns.str.strip()
            return df
        except:
            return pd.DataFrame()
    return {
        "ref1": get_df(SHEET_REF1),
        "ref2": get_df(SHEET_REF2),
        "users": get_df(SHEET_USERS),
        "sh": sh
    }

# =========================================================
# 3. ОСНОВНОЕ ПРИЛОЖЕНИЕ Axis Pro GF
# =========================================================
def main():
    st.set_page_config(page_title="Axis Pro GF", layout="wide", page_icon="🏗️")
    db = load_db()

    if 'auth' not in st.session_state:
        st.session_state.auth = False

    # --- БЛОК АВТОРИЗАЦИИ ---
    if not st.session_state.auth:
        st.title("🏗️ Axis Pro GF | Авторизация")
        col_l, _ = st.columns([1, 2])
        with col_l:
            u = st.text_input("Логин")
            p = st.text_input("Пароль", type="password")
            if st.button("Войти"):
                users = db['users']
                if not users.empty:
                    check = users[(users['Логин'] == u) & (users['Пароль'].astype(str) == p)]
                    if not check.empty:
                        st.session_state.auth = True
                        st.rerun()
                    else:
                        st.error("Неверный логин или пароль")
        return

    st.title("🏗️ Axis Pro GF | Расчетный комплекс")

    # --- ФОРМА ЗАПОЛНЕНИЯ (СИНХРОННО С ЗАПРОСОМ) ---
    with st.form("axis_main_form"):
        st.subheader("📋 Основные параметры")
        c1, c2, c3, c4 = st.columns(4)
        order_no = c1.text_input("Номер заказа", "001")
        pos_no = c2.text_input("№ позиции", "1")
        p_type = c3.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_sys = c4.selectbox("Профильная система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])

        st.subheader("📐 Геометрия и Деления (мм)")
        g1, g2, g3, g4, g5, g6 = st.columns(6)
        W = g1.number_input("Ширина, мм", value=1000)
        H = g2.number_input("Высота, мм", value=1500)
        L = g3.number_input("LEFT", value=0)
        C = g4.number_input("CENTER", value=0)
        R = g5.number_input("RIGHT", value=0)
        T = g6.number_input("TOP", value=0)

        g7, g8, g9 = st.columns(3)
        sW = g7.number_input("Ширина створки, мм", value=0)
        sH = g8.number_input("Высота створки, мм", value=0)
        qty = g9.number_input("Количество (шт)", value=1)

        st.subheader("💎 Услуги и Заполнение")
        u1, u2, u3, u4, u5 = st.columns(5)
        sp_type = u1.selectbox("Тип стеклопакета", ["двойной", "тройной", "энергодвойной", "энерготройной", "Одинарный 4мм", "Одинарный 6мм"])
        filling = u2.selectbox("Заполнение", ["Стеклопакет", "Ламбри без термо", "Ламбри с термо"])
        toning = u3.checkbox("Тонировка")
        assembly = u4.checkbox("Сборка", value=True)
        montage = u5.selectbox("Тип монтажа", ["Нет", "Монтаж", "Демонтаж/Монтаж", "Сложный монтаж"])

        submit = st.form_submit_button("🚀 ЗАПУСТИТЬ РАСЧЕТ")

    if submit:
        # --- 1. КОНТЕКСТ ДЛЯ ФОРМУЛ ---
        area = (W * H / 1000000) * qty
        ctx = {
            "W": W, "H": H, "count": qty, "qty": qty, "area": area,
            "w_s": sW, "h_s": sH, "n_m": L, "n_t": C, "math": math,
            "L": L, "C": C, "R": R, "T": T
        }

        # --- 2. РАСЧЕТ МАТЕРИАЛОВ (ПО СПРАВОЧНИКУ-1) ---
        ref1 = db['ref1'].copy()
        mats_filtered = ref1[ref1['Тип изделия'].astype(str).str.strip() == p_type]
        
        spec_table = []
        total_mats_sum = 0

        for _, row in mats_filtered.iterrows():
            try:
                # Используем колонку Формула_Python
                formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                fact_rashod = eval(formula, {"__builtins__": None}, ctx)
                
                if fact_rashod > 0:
                    norm_val = str(row.get('кол-во норм к упаковке', 1)).replace(',', '.')
                    norma_upak = float(norm_val) if norm_val else 1.0
                    if norma_upak <= 0: norma_upak = 1.0
                    
                    # Логика отгрузки (округление вверх)
                    qty_to_ship = math.ceil(fact_rashod / norma_upak)
                    
                    price_val = str(row.get('цена за ед', 0)).replace(',', '.')
                    price_unit = float(price_val) if price_val else 0.0
                    
                    # Сумма = (Цена за ед * Норма упаковки) * Кол-во к отгрузке
                    row_sum = (price_unit * norma_upak) * qty_to_ship
                    total_mats_sum += row_sum
                    
                    spec_table.append({
                        "Товар": row.get('Товар'),
                        "Артикул": row.get('Артикул'),
                        "Факт. расход": round(fact_rashod, 2),
                        "К отгрузке (упак)": qty_to_ship,
                        "Сумма": round(row_sum, 0)
                    })
            except Exception as e:
                continue

        # --- 3. ЭКОНОМИКА УСЛУГ (ТВОИ ЦЕНЫ) ---
        prices_sp = {"двойной": 9000, "тройной": 14000, "энергодвойной": 12000, "энерготройной": 15000, "Одинарный 4мм": 4000, "Одинарный 6мм": 6000}
        p_sp = prices_sp.get(sp_type, 0) if filling == "Стеклопакет" else (2248 if "без термо" in filling else 4588)
        
        sum_sp = p_sp * area
        sum_ton = (2000 * area) if toning else 0
        sum_ass = (10000 * area) if assembly else 0
        
        prices_mon = {"Монтаж": 10000, "Демонтаж/Монтаж": 12000, "Сложный монтаж": 15000, "Нет": 0}
        sum_mon = prices_mon.get(montage, 0) * area
        
        # --- 4. ИТОГО ПО ТВОЕЙ ФОРМУЛЕ ---
        costs_all = sum_sp + sum_ton + sum_ass + sum_mon + total_mats_sum
        margin = costs_all * 0.65
        grand_total = costs_all + margin

        # --- 5. ВЫВОД РЕЗУЛЬТАТОВ ---
        st.header("📊 Результаты расчета")
        
        r1, r2, r3 = st.columns(3)
        r1.metric("Общая площадь", f"{area:.3f} м2")
        r2.metric("Себестоимость (Мат+Услуги)", f"{costs_all:,.0f} ₸")
        r3.metric("ИТОГО С ОБЕСПЕЧЕНИЕМ", f"{grand_total:,.0f} ₸")

        st.subheader("📦 Детализация материалов (Отгрузка)")
        if spec_table:
            st.dataframe(pd.DataFrame(spec_table), use_container_width=True)
        else:
            st.warning("Материалы не найдены. Проверьте 'Тип изделия' в Справочнике-1.")

        st.subheader("🛠️ Смета услуг")
        serv_df = pd.DataFrame([
            {"Услуга": "Заполнение/Пакет", "Сумма": round(sum_sp, 0)},
            {"Услуга": "Тонировка", "Сумма": round(sum_ton, 0)},
            {"Услуга": "Сборка", "Сумма": round(sum_ass, 0)},
            {"Услуга": "Монтаж", "Сумма": round(sum_mon, 0)},
            {"Услуга": "ОБЕСПЕЧЕНИЕ (65%)", "Сумма": round(margin, 0)}
        ])
        st.table(serv_df)

        # ЗАПИСЬ В ЗАПРОСЫ
        try:
            db['sh'].worksheet(SHEET_FORM).append_row([
                order_no, pos_no, p_type, p_sys, W, H, qty, datetime.now().strftime("%d.%m.%Y %H:%M")
            ])
        except: pass

if __name__ == "__main__":
    main()
