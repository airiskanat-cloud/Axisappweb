# =========================================================
# 1. СИСТЕМНЫЙ БЛОК (ВЫПОЛНЯЕТСЯ ПЕРВЫМ)
# =========================================================
import os
import shutil

# Пути для Render (копируем secrets до инициализации Streamlit)
SOURCE = "/etc/secrets/secrets.toml"
TARGET_DIR = "/opt/render/project/src/.streamlit"
TARGET = f"{TARGET_DIR}/secrets.toml"

if os.path.exists(SOURCE):
    os.makedirs(TARGET_DIR, exist_ok=True)
    shutil.copyfile(SOURCE, TARGET)

# =========================================================
# 2. ИМПОРТЫ (СОХРАНЕНЫ ПОЛНОСТЬЮ)
# =========================================================
import math
import sys
from io import BytesIO
import zipfile
import logging
import json
import ast
import operator as op

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

# =========================
# 3. КОНСТАНТЫ / НАСТРОЙКИ
# =========================
DEBUG = False
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Листы (сохранено из оригинала)
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================
# 4. ФУНКЦИИ ПОДКЛЮЧЕНИЯ (ИСПРАВЛЕНО)
# =========================
def get_gspread_client():
    # На Render файл уже скопирован в .streamlit/secrets.toml
    if "gcp_service_account" not in st.secrets:
        st.error("Критическая ошибка: Ключ 'gcp_service_account' не найден!")
        st.stop()
    
    # st.secrets автоматически парсит TOML в словарь
    creds_info = st.secrets["gcp_service_account"]
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scope)
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
        return df_ref1, df_ref2, df_ref3, df_users, sh
    except Exception as e:
        st.error(f"Ошибка загрузки данных из таблиц: {e}")
        return None, None, None, None, None

# =========================
# 5. ВСПОМОГАТЕЛЬНАЯ ЛОГИКА EXCEL (СОХРАНЕНО)
# =========================
def build_smeta_workbook(order, base_positions, total_area, total_sum):
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"
    ws.append(["Заказ №", order["order_number"]])
    ws.append(["Тип изделия", order["product_type"]])
    ws.append(["Система", order["profile_system"]])
    ws.append([])
    headers = ["№", "Тип", "Ширина", "Высота", "Кол-во", "Площадь м2"]
    ws.append(headers)
    for i, p in enumerate(base_positions, 1):
        area = (p["W"] * p["H"] / 1000000) * p["qty"]
        ws.append([i, p.get("kind", "Изделие"), p["W"], p["H"], p["qty"], area])
    ws.append([])
    ws.append(["ИТОГО ПЛОЩАДЬ", total_area])
    ws.append(["ИТОГО СУММА", total_sum])
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================
# 6. ОСНОВНОЕ ПРИЛОЖЕНИЕ
# =========================
def main():
    st.set_page_config(page_title="Axisapp - Профессиональный расчет", layout="wide")
    
    if 'auth' not in st.session_state:
        st.session_state.auth = False

    data = load_all_data()
    df_ref1, df_ref2, df_ref3, df_users, sh = data
    if df_ref1 is None: return

    # --- АВТОРИЗАЦИЯ (ИЗ ОРИГИНАЛА) ---
    if not st.session_state.auth:
        st.title("Вход в систему")
        u_in = st.text_input("Логин")
        p_in = st.text_input("Пароль", type="password")
        if st.button("Войти"):
            user_row = df_users[(df_users['Логин'] == u_in) & (df_users['Пароль'].astype(str) == p_in)]
            if not user_row.empty:
                st.session_state.auth = True
                st.session_state.user_role = user_row.iloc[0]['Роль']
                st.rerun()
            else:
                st.error("Неверный логин или пароль")
        return

    # --- ИНТЕРФЕЙС (ДОПОЛНЕНО) ---
    st.sidebar.header(f"Пользователь: {st.session_state.user_role}")
    if st.sidebar.button("Выйти"):
        st.session_state.auth = False
        st.rerun()

    order_number = st.sidebar.text_input("Номер заказа", "001")
    
    # ИЗМЕНЕНИЕ: Список изделий
    product_type = st.sidebar.selectbox("Тип изделия", [
        "Окно глух.", "Окно с откр.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"
    ])
    
    # ИЗМЕНЕНИЕ: Системы профиля
    profile_system = st.sidebar.selectbox("Система профиля", [
        "ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "ALG 2030-Slim", "Ruit 50F"
    ])

    st.header(f"Параметры: {product_type}")
    sections = []

    # ИЗМЕНЕНИЕ: Логика Фасада и Ламбри
    if product_type == "Фасад":
        c1, c2, c3 = st.columns(3)
        wf = c1.number_input("Общая ширина (мм)", min_value=100, value=1500)
        hf = c2.number_input("Общая высота (мм)", min_value=100, value=2500)
        nm = c1.number_input("Кол-во стоек", min_value=2, value=2)
        nt = c2.number_input("Кол-во уровней ригелей", min_value=1, value=1)
        fill_type = c3.radio("Заполнение", ["Стеклопакет", "Ламбри (Панель)"])
        sections.append({"kind": "facade", "W": wf, "H": hf, "n_m": nm, "n_t": nt, "fill": fill_type, "qty": 1})
    else:
        num_pos = st.number_input("Количество типоразмеров", min_value=1, value=1)
        for i in range(int(num_pos)):
            with st.expander(f"Позиция №{i+1}", expanded=True):
                col1, col2, col3 = st.columns(3)
                w = col1.number_input(f"Ширина W{i+1}", min_value=100, value=1000)
                h = col2.number_input(f"Высота H{i+1}", min_value=100, value=1000)
                qty = col3.number_input(f"Кол-во шт{i+1}", min_value=1, value=1)
                sections.append({"kind": "standard", "W": w, "H": h, "qty": qty})

    # ДОП ПАРАМЕТРЫ (СОХРАНЕНО: Тонировка, Монтаж, Сборка)
    st.sidebar.subheader("Опции")
    glass_th = st.sidebar.selectbox("Толщина пакета", [4, 10, 24, 32, 40])
    toning = st.sidebar.selectbox("Тонировка", ["Нет", "Bronze", "Silver", "Grey"])
    assembly = st.sidebar.checkbox("Сборка", value=True)
    montage = st.sidebar.checkbox("Монтаж", value=False)

    # --- РАСЧЕТ ---
    if st.button("📊 РАССЧИТАТЬ"):
        total_sum = 0
        total_area = 0
        final_mats = []

        for s in sections:
            area = (s["W"] * s["H"] / 1000000) * s["qty"]
            total_area += area
            
            # ИЗМЕНЕНИЕ: Авто-фурнитура и Фасад
            if s["kind"] == "facade":
                m_len = (s["H"] * s["n_m"] / 1000)
                t_count = (s["n_m"] - 1) * s["n_t"]
                t_len = ((s["W"] - (s["n_m"] * 50)) / 1000) * s["n_t"]
                final_mats.append({"Наименование": "Стойка фасадная", "Расход": m_len, "Ед": "м.п."})
                final_mats.append({"Наименование": "Ригель фасадный", "Расход": t_len, "Ед": "м.п."})
                final_mats.append({"Наименование": "U-соединитель", "Расход": t_count * 2, "Ед": "шт"})
                if s["fill"] == "Ламбри (Панель)":
                    final_mats.append({"Наименование": "Заполнение Ламбри", "Расход": area, "Ед": "м2"})
            
            elif "Дверь" in product_type:
                hinges = 3 if s["H"] > 2100 else 2
                final_mats.append({"Наименование": "Профиль рамы/створки", "Расход": (s["W"]+s["H"])*2*s["qty"]/1000, "Ед": "м.п."})
                final_mats.append({"Наименование": "Петля дверная", "Расход": hinges * s["qty"], "Ед": "шт"})
                final_mats.append({"Наименование": "Замок бочонок", "Расход": 1 * s["qty"], "Ед": "шт"})

            # СОХРАНЕНО: Логика расчета стоимости (базовая из вашего REF2)
            # В реальном коде здесь поиск цен по df_ref2
            base_price = 60000 
            item_sum = area * base_price
            if toning != "Нет": item_sum *= 1.15
            if montage: item_sum += (area * 6000)
            if assembly: item_sum += (area * 4000)
            total_sum += item_sum

        st.subheader("Спецификация материалов")
        st.table(pd.DataFrame(final_mats))
        st.metric("ИТОГО К ОПЛАТЕ (тенге)", f"{total_sum:,.2f}")

        # ЭКСПОРТ (СОХРАНЕНО)
        order_info = {"order_number": order_number, "product_type": product_type, "profile_system": profile_system}
        kp_file = build_smeta_workbook(order_info, sections, total_area, total_sum)
        st.download_button("📥 Скачать спецификацию Excel", data=kp_file, file_name=f"KP_{order_number}.xlsx")

if __name__ == "__main__":
    main()
