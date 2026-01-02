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

# =========================
# КОНСТАНТЫ
# =========================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================
# ПОДКЛЮЧЕНИЕ
# =========================
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=600)
def load_db():
    client = get_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    def get_df(name):
        df = pd.DataFrame(sh.worksheet(name).get_all_records())
        df.columns = df.columns.str.strip()
        return df
    return {
        "ref1": get_df(SHEET_REF1),
        "ref2": get_df(SHEET_REF2),
        "ref3": get_df(SHEET_REF3),
        "users": get_df(SHEET_USERS),
        "sh": sh
    }

# =========================
# ОСНОВНОЙ КОД Axis Pro GF
# =========================
def main():
    st.set_page_config(page_title="Axis Pro GF", layout="wide")
    db = load_db()

    if 'auth' not in st.session_state:
        st.session_state.auth = False

    # --- БЛОК ЛОГИНА (ВОЗВРАЩЕН) ---
    if not st.session_state.auth:
        st.title("🏗️ Axis Pro GF | Вход")
        col_l, col_r = st.columns([1, 1])
        with col_l:
            u = st.text_input("Логин")
            p = st.text_input("Пароль", type="password")
            if st.button("Войти"):
                user_check = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
                if not user_check.empty:
                    st.session_state.auth = True
                    st.session_state.user_role = user_check.iloc[0]['Роль']
                    st.rerun()
                else:
                    st.error("Неверные данные")
        return

    st.title("🏗️ Axis Pro GF | Расчетный комплекс")

    # --- ФОРМА ЗАПОЛНЕНИЯ (ТВОИ 22 ПОЛЯ) ---
    with st.form("axis_form"):
        st.subheader("1. Заказ и Профиль")
        c1, c2, c3, c4 = st.columns(4)
        order_no = c1.text_input("Номер заказа", "001")
        pos_no = c2.text_input("№ позиции", "1")
        p_type = c3.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_sys = c4.selectbox("Профильная система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])

        st.subheader("2. Геометрия (мм)")
        g1, g2, g3, g4, g5, g6 = st.columns(6)
        W = g1.number_input("Ширина, мм", value=1000)
        H = g2.number_input("Высота, мм", value=1500)
        L = g3.number_input("LEFT, мм", value=0)
        C = g4.number_input("CENTER, мм", value=0)
        R = g5.number_input("RIGHT, мм", value=0)
        T = g6.number_input("TOP, мм", value=0)

        g7, g8, g9, g10 = st.columns(4)
        sW = g7.number_input("Ширина створки, мм", value=0)
        sH = g8.number_input("Высота створки, мм", value=0)
        qty = g9.number_input("Количество (шт)", value=1)
        nwin = g10.number_input("Nwin", value=1)

        st.subheader("3. Заполнение и Услуги")
        u1, u2, u3, u4, u5 = st.columns(5)
        sp_type = u1.selectbox("Тип стеклопакета", ["двойной", "тройной", "энергодвойной", "энерготройной", "Одинарный 4мм", "Одинарный 6мм"])
        filling = u2.selectbox("Заполнение", ["Стеклопакет", "Ламбри без термо", "Ламбри с термо"])
        toning = u3.checkbox("Тонировка")
        assembly = u4.checkbox("Сборка", value=True)
        montage = u5.selectbox("Тип монтажа", ["Нет", "Монтаж", "Демонтаж/Монтаж", "Сложный монтаж"])

        submit = st.form_submit_button("🚀 РАССЧИТАТЬ МАТЕРИАЛЫ И СТОИМОСТЬ")

    if submit:
        # --- ПОДГОТОВКА КОНТЕКСТА ---
        ctx = {
            "W": W, "H": H, "count": qty, "qty": qty,
            "w_s": sW, "h_s": sH, "n_m": L, "n_t": C,
            "L": L, "C": C, "R": R, "T": T,
            "math": math
        }

        # --- РАСЧЕТ МАТЕРИАЛОВ (ПО СПРАВОЧНИКУ-1 И СПРАВОЧНИКУ-3) ---
        # Мы ищем в Справочнике-1 те строки, которые подходят под тип и систему
        ref1 = db['ref1'].copy()
        ref1['Тип изделия'] = ref1['Тип изделия'].astype(str).str.strip()
        ref1['Система профиля'] = ref1['Система профиля'].astype(str).str.strip()
        
        # Фильтр по типу и системе
        mats_filtered = ref1[(ref1['Тип изделия'] == p_type) & (ref1['Система профиля'] == p_sys)]
        
        spec_res = []
        total_mats_cost = 0

        for _, row in mats_filtered.iterrows():
            try:
                # Используем формулу из колонки Формула_Python
                formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                res_qty = eval(formula, {"__builtins__": None}, ctx)
                
                if res_qty > 0:
                    price = float(row.get('цена за ед', 0))
                    row_cost = res_qty * price
                    total_mats_cost += row_cost
                    spec_res.append({
                        "Тип элемента": row.get('Тип элемента', ''),
                        "Товар": row.get('Товар', ''),
                        "Артикул": row.get('Артикул', ''),
                        "Расход": round(res_qty, 2),
                        "Ед": row.get('Ед. фактического расхода', 'м.п.'),
                        "Сумма": round(row_cost, 0)
                    })
            except Exception as e:
                continue

        # --- ЭКОНОМИКА (ПО ЦЕНАМ ИЗ ТВОЕГО СПИСКА) ---
        area = (W * H / 1000000) * qty
        prices_sp = {"двойной": 9000, "тройной": 14000, "энергодвойной": 12000, "энерготройной": 15000, "Одинарный 4мм": 4000, "Одинарный 6мм": 6000}
        p_sp = prices_sp.get(sp_type, 0) if filling == "Стеклопакет" else (2248 if "без термо" in filling else 4588)
        
        p_ton = 2000 if toning else 0
        p_ass = 10000 if assembly else 0
        p_mon = 10000 if montage == "Монтаж" else 12000 if montage == "Демонтаж/Монтаж" else 15000 if montage == "Сложный монтаж" else 0
        
        # Сетка услуг
        serv_data = [
            {"Услуга": "Стеклопакет/Заполнение", "Ед": "м2", "Итого": p_sp * area},
            {"Услуга": "Нарезка", "Ед": "м2", "Итого": 4000 * area},
            {"Услуга": "Тонировка", "Ед": "м2", "Итого": p_ton * area},
            {"Услуга": "Сборка", "Ед": "м2", "Итого": p_ass * area},
            {"Услуга": "Монтаж", "Ед": "м2", "Итого": p_mon * area}
        ]
        total_serv = sum(s['Итого'] for s in serv_data)
        
        base_sum = total_mats_cost + total_serv
        margin = base_sum * 0.65
        final_total = base_sum + margin

        # --- ВЫВОД РЕЗУЛЬТАТОВ ---
        st.header("📊 Результаты расчета")
        
        r1, r2, r3 = st.columns(3)
        r1.metric("Площадь", f"{area:.3f} м2")
        r2.metric("Мат. себестоимость", f"{total_mats_cost:,.0f} ₸")
        r3.metric("ИТОГО К ОПЛАТЕ", f"{final_total:,.0f} ₸")

        st.subheader("📋 Детальная смета услуг")
        st.table(pd.DataFrame(serv_data))

        with st.expander("🔍 Посмотреть расход материалов (из Справочника-1)"):
            if spec_res:
                st.dataframe(pd.DataFrame(spec_res), use_container_width=True)
            else:
                st.warning("Материалы не найдены. Проверьте Тип изделия и Систему профиля в Справочнике-1.")

        # ЗАПИСЬ В ЗАПРОСЫ
        try:
            db['sh'].worksheet(SHEET_FORM).append_row([order_no, pos_no, p_type, p_sys, W, H, qty, datetime.now().strftime("%d.%m.%Y")])
        except: pass

if __name__ == "__main__":
    main()
