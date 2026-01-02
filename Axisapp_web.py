import math
import os
import sys
import shutil
from io import BytesIO
import zipfile
import logging
import json
import ast
import operator as op
from datetime import datetime

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# =========================================================
# 1. КОНСТАНТЫ И НАСТРОЙКИ (ВСЕ НАЗВАНИЯ ИЗ v15)
# =========================================================
DEBUG = False
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Листы Google Таблиц (строго как в v15)
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================================================
# 2. ПОДКЛЮЧЕНИЕ К ДАННЫМ (АДАПТАЦИЯ ПОД RENDER)
# =========================================================
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        # Читаем секретный файл gcp.json на Render
        creds_path = "/etc/secrets/gcp.json"
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Ошибка авторизации Google API: {e}")
        st.stop()

@st.cache_data(ttl=600)
def load_all_data():
    client = get_gspread_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    return {
        "ref1": pd.DataFrame(sh.worksheet(SHEET_REF1).get_all_records()),
        "ref2": pd.DataFrame(sh.worksheet(SHEET_REF2).get_all_records()),
        "ref3": pd.DataFrame(sh.worksheet(SHEET_REF3).get_all_records()),
        "users": pd.DataFrame(sh.worksheet(SHEET_USERS).get_all_records()),
        "sh": sh
    }

# =========================================================
# 3. ИНЖЕНЕРНЫЕ ФУНКЦИИ (УСИЛЕННЫЕ v15)
# =========================================================
def evaluate_formula(formula_str, context):
    """
    Выполняет расчет материалов. 
    Добавлена поддержка автоматики петель и сложных вычетов.
    """
    try:
        expr = str(formula_str).replace('=', '').replace('^', '**')
        # Автоматика: если высота > 2100, добавляем 1 петлю к расчету (context['hinges'])
        hinge_extra = 1 if context.get('H', 0) > 2100 else 0
        
        allowed_names = {
            "math": math, 
            "W": context.get('W', 0), 
            "H": context.get('H', 0), 
            "qty": context.get('qty', 1),
            "n_imp": context.get('n_imp', 0),
            "hinges": 2 + hinge_extra # Базово 2 петли + 1 если высокая
        }
        return eval(expr, {"__builtins__": None}, allowed_names)
    except Exception as e:
        logger.error(f"Ошибка в формуле {formula_str}: {e}")
        return 0

# =========================================================
# 4. ИНТЕРФЕЙС (ПРИВЫЧНЫЙ v15)
# =========================================================
def main():
    st.set_page_config(page_title="Axisapp v15 Pro", layout="wide")
    db = load_all_data()

    if 'auth' not in st.session_state: st.session_state.auth = False

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("Вход в систему AXIS")
        u = st.text_input("Логин")
        p = st.text_input("Пароль", type="password")
        if st.button("Войти"):
            user = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
            if not user.empty:
                st.session_state.auth = True
                st.session_state.role = user.iloc[0]['Роль']
                st.rerun()
        return

    # --- ПРИВЫЧНАЯ БОКОВАЯ ПАНЕЛЬ ---
    st.sidebar.title("Параметры заказа")
    order_number = st.sidebar.text_input("Номер заказа", "001")
    
    # ТИПЫ ИЗДЕЛИЙ (Обновленные, как ты просила)
    product_type = st.sidebar.selectbox("Тип изделия", 
        ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
    
    # СИСТЕМЫ ПРОФИЛЯ (Обновленные)
    profile_system = st.sidebar.selectbox("Профильная система", 
        ["ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "ALG 2030-Slim", "Ruit 50F"])

    # Доп. опции v15
    toning = st.sidebar.checkbox("Тонировка")
    assembly = st.sidebar.checkbox("Сборка", value=True)
    montage = st.sidebar.checkbox("Монтаж")
    handle_type = st.sidebar.selectbox("Тип ручек", ["Нажимная", "Офисная"])

    # --- ВВОД ПОЗИЦИЙ ---
    st.header(f"Расчет объекта: {product_type}")
    num_pos = st.number_input("Количество типоразмеров", min_value=1, value=1)
    
    sections = []
    for i in range(int(num_pos)):
        with st.expander(f"Позиция №{i+1}", expanded=True):
            colW, colH, colQ, colI = st.columns(4)
            w = colW.number_input(f"Ширина W{i+1}", value=1000, key=f"w_{i}")
            h = colH.number_input(f"Высота H{i+1}", value=1500, key=f"h_{i}")
            q = colQ.number_input(f"Кол-во шт{i+1}", value=1, key=f"q_{i}")
            # Поле для импостов (или стоек для фасада)
            imp = colI.number_input(f"Деления (импосты) {i+1}", value=0, key=f"i_{i}")
            sections.append({"W": w, "H": h, "qty": q, "n_imp": imp})

    if st.button("🏗️ РАССЧИТАТЬ"):
        st.subheader("Результаты инженерного расчета")
        
        # Основной цикл расчета по Справочнику-3
        all_materials = []
        total_mats_cost = 0
        total_area = 0

        for s in sections:
            area = (s['W'] * s['H'] / 1000000) * s['qty']
            total_area += area
            
            # Фильтруем материалы
            mats = db['ref3'][db['ref3']['Тип изделия'] == product_type]
            for _, row in mats.iterrows():
                q_mat = evaluate_formula(row['Формула_Python'], s)
                if q_mat > 0:
                    # Цена из Справочника-2
                    price_row = db['ref2'][db['ref2']['Система'] == profile_system]
                    price = price_row['Цена'].values[0] if not price_row.empty else 0
                    
                    cost = q_mat * price
                    total_mats_cost += cost
                    all_materials.append({"Материал": row['Наименование'], "Кол-во": q_mat, "Сумма": cost})

        # Экономика v15: (Мат + Стекло + Работа) * 1.65
        glass_sum = total_area * 18000 + (total_area * 4500 if toning else 0)
        labor_sum = total_area * (5000 if assembly else 0) + total_area * (7000 if montage else 0)
        final_total = (total_mats_cost + glass_sum + labor_sum) * 1.65

        # ВЫВОД (v15 Style)
        st.info(f"Общая площадь заказа: {total_area:.3f} м2")
        st.success(f"ИТОГО К ОПЛАТЕ: {final_total:,.0f} тенге")
        
        with st.expander("Детальная спецификация материалов"):
            st.table(pd.DataFrame(all_materials).groupby("Материал").sum())

        # ЗАПИСЬ В ОБЛАКО (v15 Style)
        try:
            db['sh'].worksheet(SHEET_FINAL).append_row([
                order_number, product_type, total_area, final_sum, datetime.now().strftime("%d.%m.%Y")
            ])
        except: pass

if __name__ == "__main__":
    main()
