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
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# =========================================================
# 1. НАСТРОЙКИ И КОНСТАНТЫ
# =========================================================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_FINAL = "Итоговый расчет с монтажом"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# =========================================================
# 2. ПОДКЛЮЧЕНИЕ
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
        df = pd.DataFrame(sh.worksheet(name).get_all_records())
        df.columns = df.columns.str.strip()
        return df
    return {"ref2": get_df(SHEET_REF2), "ref3": get_df(SHEET_REF3), "sh": sh}

# =========================================================
# 3. ОСНОВНОЕ ПРИЛОЖЕНИЕ
# =========================================================
def main():
    st.set_page_config(page_title="Axis Pro GF", layout="wide")
    st.title("🏗️ Axis Pro GF | Инженерный комплекс")

    db = load_db()

    # --- ФОРМА ЗАПОЛНЕНИЯ (СИНХРОННО С ЗАПРОСОМ) ---
    with st.form("axis_form"):
        st.subheader("📋 Данные заказа и профиля")
        c1, c2, c3, c4 = st.columns(4)
        order_no = c1.text_input("Номер заказа", "001")
        pos_no = c2.text_input("№ позиции", "1")
        p_type = c3.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_sys = c4.selectbox("Профильная система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])

        st.subheader("📐 Геометрия (мм)")
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
        qty = g9.number_input("Кол-во шт (Nwin)", value=1)
        s_count = g10.number_input("Створки (шт)", value=0)

        st.subheader("⚙️ Услуги и Заполнение")
        u1, u2, u3, u4, u5 = st.columns(5)
        sp_type = u1.selectbox("Стеклопакет", ["двойной", "тройной", "энергодвойной", "энерготройной", "Одинарный 4мм", "Одинарный 6мм"])
        filling = u2.selectbox("Заполнение", ["Стеклопакет", "Ламбри без термо", "Ламбри с термо"])
        toning = u3.checkbox("Тонировка")
        assembly = u4.checkbox("Сборка", value=True)
        montage = u5.selectbox("Монтаж", ["Нет", "Монтаж", "Демонтаж/Монтаж", "Сложный монтаж"])

        submit = st.form_submit_button("🚀 РАССЧИТАТЬ")

    if submit:
        # --- ПОДГОТОВКА КОНТЕКСТА ДЛЯ ФОРМУЛ ---
        # Здесь мы прописываем все варианты имен переменных из Справочника-3
        ctx = {
            "W": W, "H": H, "count": qty, "qty": qty,
            "w_s": sW, "h_s": sH, "w_stvor": sW, "h_stvor": sH,
            "n_lp": 4, "lock_points": 4, "math": math,
            "n_m": L if p_type == "Фасад" else 0, # Стойки фасада
            "n_t": C if p_type == "Фасад" else 0, # Ригели фасада
            "total_area": (W * H / 1000000) * qty
        }

        # --- РАСЧЕТ МАТЕРИАЛОВ ---
        ref3 = db['ref3'].copy()
        ref3['Тип изделия'] = ref3['Тип изделия'].astype(str).str.strip()
        mats_filtered = ref3[ref3['Тип изделия'] == p_type]
        
        spec_res = []
        total_mats_cost = 0

        for _, row in mats_filtered.iterrows():
            try:
                formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                res_qty = eval(formula, {"__builtins__": None}, ctx)
                if res_qty > 0:
                    # Поиск цены в Справочнике-2
                    prices = db['ref2'][db['ref2']['Система'].str.strip() == p_sys]
                    u_price = prices['Цена'].values[0] if not prices.empty else 0
                    
                    row_cost = res_qty * u_price
                    total_mats_cost += row_cost
                    spec_res.append({
                        "Тип": row.get('Тип элемента', 'Профиль'),
                        "Название": row.get('Комплектующие', 'Профиль'),
                        "Расход": round(res_qty, 2),
                        "Сумма": round(row_cost, 0)
                    })
            except: continue

        # --- ЭКОНОМИКА (ПО ТВОЕМУ СПИСКУ) ---
        area = ctx["total_area"]
        # Цены из твоего сообщения
        prices_sp = {"двойной": 9000, "тройной": 14000, "энергодвойной": 12000, "энерготройной": 15000, "Одинарный 4мм": 4000, "Одинарный 6мм": 6000}
        p_sp = prices_sp.get(sp_type, 0) if filling == "Стеклопакет" else (2248 if "без термо" in filling else 4588)
        
        p_ton = 2000 if toning else 0
        p_ass = 10000 if assembly else 0
        p_mon = 10000 if montage == "Монтаж" else 12000 if montage == "Демонтаж/Монтаж" else 15000 if montage == "Сложный монтаж" else 0
        
        # Сетка услуг
        services = [
            {"Услуга": "Стеклопакет/Панель", "Итого": p_sp * area},
            {"Услуга": "Нарезка", "Итого": 4000 * area},
            {"Услуга": "Тонировка", "Итого": p_ton * area},
            {"Услуга": "Сборка", "Итого": p_ass * area},
            {"Услуга": "Монтаж", "Итого": p_mon * area}
        ]
        total_serv = sum(s['Итого'] for s in services)
        
        # ИТОГИ
        base_sum = total_mats_cost + total_serv
        margin = base_sum * 0.65
        final_total = base_sum + margin

        # --- ВЫВОД ---
        st.header("📊 Итоги расчета Axis Pro GF")
        c_res1, c_res2 = st.columns(2)
        
        with c_res1:
            st.subheader("Смета услуг")
            st.table(pd.DataFrame(services))
        
        with c_res2:
            st.subheader("Финансовый результат")
            st.metric("Общая площадь", f"{area:.3f} м2")
            st.metric("Себестоимость материалов", f"{total_mats_cost:,.0f} ₸")
            st.write(f"**Обеспечение (65%):** {margin:,.0f} ₸")
            st.title(f"ИТОГО: {final_total:,.0f} ₸")

        with st.expander("🔍 Детальный расход материалов"):
            if spec_res:
                st.dataframe(pd.DataFrame(spec_res), use_container_width=True)
            else:
                st.warning("Материалы не найдены. Проверьте 'Тип изделия' в Справочнике-3.")

        # ЗАПИСЬ В ИСТОРИЮ (ЗАПРОСЫ)
        try:
            db['sh'].worksheet(SHEET_FORM).append_row([
                order_no, pos_no, p_type, p_sys, W, H, qty, datetime.now().strftime("%d.%m.%Y %H:%M")
            ])
            st.toast("Запрос сохранен в историю")
        except: pass

if __name__ == "__main__":
    main()
