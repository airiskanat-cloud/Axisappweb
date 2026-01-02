import math
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
from datetime import datetime

# Настройки листов
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"

def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
    return gspread.authorize(creds)

def main():
    st.set_page_config(page_title="Axis Pro GF", layout="wide")
    st.title("🏗️ Axis Pro GF | Инженерный комплекс")

    db = load_data() # Функция загрузки данных (ref2, ref3)
    
    # --- ФОРМА ЗАПОЛНЕНИЯ (СИНХРОННАЯ) ---
    with st.form("main_form"):
        st.subheader("📋 Основные параметры")
        c1, c2, c3, c4 = st.columns(4)
        order_no = c1.text_input("Номер заказа", "001")
        pos_no = c2.text_input("№ позиции", "1")
        p_type = c3.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_sys = c4.selectbox("Профильная система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])

        st.subheader("📐 Геометрия и Конструкция")
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
        qty = g9.number_input("Кол-во (Nwin)", value=1)
        s_count = g10.number_input("Створки (шт)", value=1 if "откр" in p_type else 0)

        st.subheader("💎 Заполнение и Услуги")
        u1, u2, u3, u4, u5 = st.columns(5)
        # Выбор из справочника 2
        sp_type = u1.selectbox("Тип стеклопакета", ["двойной", "тройной", "энергодвойной", "энерготройной", "Одинарный 4мм", "Одинарный 6мм"])
        panel_type = u2.selectbox("Панели", ["Нет", "Ламбри без термо", "Ламбри с термо"])
        toning = u3.checkbox("Тонировка")
        assembly = u4.checkbox("Сборка", value=True)
        montage = u5.selectbox("Тип монтажа", ["Нет", "Монтаж", "Демонтаж/Монтаж", "Сложный монтаж"])

        submit = st.form_submit_button("🚀 РАССЧИТАТЬ МАТЕРИАЛЫ И СТОИМОСТЬ")

    if submit:
        # --- 1. ПОДГОТОВКА ИНЖЕНЕРНЫХ ПЕРЕМЕННЫХ (ДЛЯ FORMULA_PYTHON) ---
        area = (W * H / 1000000) * qty
        # Периметры и вычеты
        context = {
            "W": W, "H": H, "count": qty, 
            "w_s": sW, "h_s": sH, 
            "w_g": W - 100, "h_g": H - 100, # Пример вычета
            "math": math
        }

        # --- 2. РАСЧЕТ МАТЕРИАЛОВ (ПО ТИПУ ИЗДЕЛИЯ) ---
        mats_ref = db['ref3'][db['ref3']['Тип изделия'] == p_type]
        total_mats_price = 0
        spec_list = []

        for _, row in mats_ref.iterrows():
            try:
                formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                q_res = eval(formula, {"__builtins__": None}, context)
                if q_res > 0:
                    # Ищем цену за единицу в Справочнике-2 для системы
                    price_unit = db['ref2'][db['ref2']['Система'] == p_sys]['Цена'].values[0]
                    total_mats_price += (q_res * price_unit)
                    spec_list.append({"Тип": row['Тип элемента'], "Название": row.get('Комплектующие', 'Профиль'), "Расход": q_res})
            except: continue

        # --- 3. РАСЧЕТ УСЛУГ (ИЗ СПРАВОЧНИКА 2) ---
        # Цены из твоего списка
        p_sp = 9000 if sp_type == "двойной" else 14000 # и т.д. по списку
        p_ton = 2000 if toning else 0
        p_ass = 10000 if assembly else 0
        p_mon = 10000 if montage == "Монтаж" else 15000 if montage == "Сложный монтаж" else 0
        
        services_data = [
            {"Услуга": "Стеклопакет", "Цена м2": p_sp, "Итого": p_sp * area},
            {"Услуга": "Тонировка", "Цена м2": p_ton, "Итого": p_ton * area},
            {"Услуга": "Сборка", "Цена м2": p_ass, "Итого": p_ass * area},
            {"Услуга": "Монтаж", "Цена м2": p_mon, "Итого": p_mon * area},
        ]
        total_services = sum(item['Итого'] for item in services_data)

        # --- 4. ИТОГИ (ОБЕСПЕЧЕНИЕ) ---
        subtotal = total_mats_price + total_services
        margin = subtotal * 0.65 # Твоё обеспечение
        grand_total = subtotal + margin

        # --- ВЫВОД ---
        st.header("📊 Результаты расчета Axis Pro GF")
        
        col_res1, col_res2 = st.columns(2)
        with col_res1:
            st.subheader("Смета услуг")
            st.table(pd.DataFrame(services_data))
        with col_res2:
            st.subheader("Итоговые показатели")
            st.metric("Площадь изделия", f"{area:.3f} м2")
            st.metric("Материалы (себест.)", f"{total_mats_price:,.0f} ₸")
            st.write(f"**Обеспечение (65%):** {margin:,.2f} ₸")
            st.title(f"ИТОГО: {grand_total:,.2f} ₸")

        with st.expander("🔍 Детальный расход материалов (Справочник-3)"):
            st.dataframe(pd.DataFrame(spec_list))

# Вспомогательная загрузка
def load_data():
    client = get_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    return {
        "ref2": pd.DataFrame(sh.worksheet(SHEET_REF2).get_all_records()),
        "ref3": pd.DataFrame(sh.worksheet(SHEET_REF3).get_all_records()),
        "sh": sh
    }

if __name__ == "__main__":
    main()
