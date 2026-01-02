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
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# =========================================================
# 1. СИСТЕМНЫЕ НАСТРОЙКИ (Axis Pro GF)
# =========================================================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Названия листов из твоей оригинальной базы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================================================
# 2. ПОДКЛЮЧЕНИЕ К ДАННЫМ (С ЗАЩИТОЙ ОТ ПРОБЕЛОВ)
# =========================================================
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"❌ Ошибка авторизации Google API: {e}")
        st.stop()

@st.cache_data(ttl=600)
def load_all_data():
    client = get_gspread_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    
    # Загрузка с немедленной очисткой названий колонок
    def get_clean_df(sheet_name):
        df = pd.DataFrame(sh.worksheet(sheet_name).get_all_records())
        df.columns = df.columns.str.strip() # Удаляет пробелы из названий колонок
        return df

    return {
        "ref1": get_clean_df(SHEET_REF1),
        "ref2": get_clean_df(SHEET_REF2),
        "ref3": get_clean_df(SHEET_REF3),
        "users": get_clean_df(SHEET_USERS),
        "sh": sh
    }

# =========================================================
# 3. ИНЖЕНЕРНЫЙ РАСЧЕТ (v15 Logic)
# =========================================================
def evaluate_formula(formula_str, context):
    try:
        expr = str(formula_str).replace('=', '').replace('^', '**')
        # В v15 важны W, H, qty
        allowed_names = {
            "math": math, 
            "W": context.get('W', 0), 
            "H": context.get('H', 0), 
            "qty": context.get('qty', 1),
            "n_imp": context.get('n_imp', 0)
        }
        return eval(expr, {"__builtins__": None}, allowed_names)
    except Exception as e:
        return 0

# =========================================================
# 4. ЭСТЕТИЧНЫЙ ИНТЕРФЕЙС
# =========================================================
def main():
    st.set_page_config(page_title="Axis Pro GF", page_icon="🏗️", layout="wide")
    
    # Современный стиль через CSS
    st.markdown("""
        <style>
        .main { background-color: #f5f7f9; }
        .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #1e3d59; color: white; }
        .stMetric { background-color: white; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        </style>
    """, unsafe_allow_html=True)

    db = load_all_data()
    if 'auth' not in st.session_state: st.session_state.auth = False

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.image("https://static.tildacdn.com/tild3133-3131-4131-b331-313131313131/logo_axis.png", width=200)
            st.title("Axis Pro GF | Авторизация")
            u = st.text_input("Логин оператора")
            p = st.text_input("Пароль", type="password")
            if st.button("Войти в систему"):
                user = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
                if not user.empty:
                    st.session_state.auth = True
                    st.session_state.role = user.iloc[0]['Роль']
                    st.rerun()
        return

    # --- ПАНЕЛЬ УПРАВЛЕНИЯ (Sidebar) ---
    st.sidebar.markdown(f"### 👤 {st.session_state.role}")
    st.sidebar.title("Axis Pro GF")
    
    order_number = st.sidebar.text_input("Заказ №", "2025-GF-01")
    
    # Выпадающие списки
    p_type = st.sidebar.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
    p_sys = st.sidebar.selectbox("Профильная система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F", "ALG Slim"])
    
    st.sidebar.markdown("---")
    toning = st.sidebar.checkbox("Тонировка стекла")
    assembly = st.sidebar.checkbox("Сборка изделия", value=True)
    montage = st.sidebar.checkbox("Монтаж на объекте")

    # --- ОСНОВНОЙ КОНСТРУКТОР ---
    st.header(f"🏗️ Конструктор изделия: {p_type}")
    
    with st.container():
        c1, c2, c3, c4 = st.columns(4)
        W = c1.number_input("Ширина W (мм)", value=1000)
        H = c2.number_input("Высота H (мм)", value=1500)
        qty = c3.number_input("Количество (шт)", min_value=1, value=1)
        
        # Интеллектуальное поле импоста
        label_imp = "Кол-во стоек/ригелей" if p_type == "Фасад" else "Кол-во импостов"
        n_imp = c4.number_input(label_imp, value=0)

    st.markdown("---")

    if st.button("🚀 ЗАПУСТИТЬ ИНЖЕНЕРНЫЙ РАСЧЕТ"):
        with st.spinner('Обработка формул Справочника-3...'):
            # Проверка колонки
            if 'Тип изделия' not in db['ref3'].columns:
                st.error("Критическая ошибка: В Справочнике-3 не найдена колонка 'Тип изделия'. Проверьте таблицу.")
                st.stop()

            # Расчет материалов
            mats_spec = []
            total_mats_cost = 0
            
            # Фильтрация по типу
            ref3_filtered = db['ref3'][db['ref3']['Тип изделия'] == p_type]
            context = {"W": W, "H": H, "qty": qty, "n_imp": n_imp}

            for _, row in ref3_filtered.iterrows():
                q_mat = evaluate_formula(row['Формула_Python'], context)
                if q_mat > 0:
                    # Цена из Справочника-2
                    price_row = db['ref2'][db['ref2']['Система'] == p_sys]
                    price = price_row['Цена'].values[0] if not price_row.empty else 0
                    
                    cost = q_mat * price
                    total_mats_cost += cost
                    mats_spec.append({
                        "Наименование": row['Наименование'],
                        "Количество": round(q_mat, 2),
                        "Ед.": row['Ед'],
                        "Сумма": round(cost, 0)
                    })

            # Экономика Axis (1.65)
            area = (W * H / 1000000) * qty
            glass_sum = area * 18500 + (area * 4500 if toning else 0)
            labor_sum = area * (5000 if assembly else 0) + area * (8000 if montage else 0)
            
            grand_total = (total_mats_cost + glass_sum + labor_sum) * 1.65

            # ВЫВОД РЕЗУЛЬТАТОВ
            res1, res2, res3 = st.columns(3)
            res1.metric("Общая площадь", f"{area:.3f} м²")
            res2.metric("Себестоимость мат.", f"{total_mats_cost:,.0f} ₸")
            res3.metric("ИТОГО К ОПЛАТЕ", f"{grand_total:,.0f} ₸", delta="Обеспечение 65%")

            st.markdown("### 📋 Ведомость материалов (Спецификация)")
            st.table(pd.DataFrame(mats_spec))

            # ЗАПИСЬ В ОБЛАКО
            try:
                db['sh'].worksheet(SHEET_FINAL).append_row([
                    order_number, p_type, area, grand_total, datetime.now().strftime("%d.%m.%Y")
                ])
                st.toast("Данные сохранены в Google Sheets", icon="✅")
            except:
                st.warning("Запись в облако не удалась, проверьте доступ к таблице.")

    if st.sidebar.button("🚪 Выход"):
        st.session_state.auth = False
        st.rerun()

if __name__ == "__main__":
    main()
