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

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

# =========================================================
# 1. КОНСТАНТЫ / НАСТРОЙКИ
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

# Листы Google Таблицы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================================================
# 2. ПОДКЛЮЧЕНИЕ К GOOGLE (ИСПРАВЛЕННЫЙ МЕТОД)
# =========================================================
def get_gspread_client():
    """
    Подключение через секретный JSON файл Render.
    Путь /etc/secrets/gcp.json настраивается в панели Render.
    """
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    # Прямое чтение файла, без использования st.secrets
    try:
        creds = Credentials.from_service_account_file(
            "/etc/secrets/gcp.json",
            scopes=scopes
        )
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Ошибка доступа к файлу ключей /etc/secrets/gcp.json: {e}")
        st.stop()

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
        st.error(f"Ошибка при загрузке данных из Google Таблиц: {e}")
        return None, None, None, None, None

# =========================================================
# 3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (EXCEL И РАСЧЕТЫ)
# =========================================================
def build_smeta_workbook(order, base_positions, total_area, total_sum):
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"
    
    ws.append(["Заказ №", order.get("order_number")])
    ws.append(["Тип изделия", order.get("product_type")])
    ws.append(["Система", order.get("profile_system")])
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

# =========================================================
# 4. ОСНОВНОЕ ПРИЛОЖЕНИЕ (UI И ЛОГИКА)
# =========================================================
def main():
    st.set_page_config(page_title="Axisapp - Профессиональный расчет", layout="wide")
    
    if 'auth' not in st.session_state:
        st.session_state.auth = False

    # Загрузка данных
    data = load_all_data()
    if data[0] is None:
        return
    df_ref1, df_ref2, df_ref3, df_users, sh = data

    # --- БЛОК АВТОРИЗАЦИИ ---
    if not st.session_state.auth:
        st.title("Вход в систему Axisapp")
        col_login, _ = st.columns([1, 2])
        with col_login:
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

    # --- ИНТЕРФЕЙС ПОСЛЕ ВХОДА ---
    st.sidebar.header(f"Пользователь: {st.session_state.user_role}")
    if st.sidebar.button("Выйти"):
        st.session_state.auth = False
        st.rerun()

    order_number = st.sidebar.text_input("Номер заказа", "001")
    
    # Списки выбора (согласно ТЗ)
    product_type = st.sidebar.selectbox("Тип изделия", [
        "Окно глух.", "Окно с откр.", 
        "Дверь 1 створч.", "Дверь 2-х створч.", 
        "Фасад"
    ])
    
    profile_system = st.sidebar.selectbox("Система профиля", [
        "ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", 
        "ALG 2030-73C", "ALG 2030-Slim", "Ruit 50F"
    ])

    st.header(f"Конфигурация: {product_type} ({profile_system})")
    
    sections = []
    
    # Динамическая форма ввода
    if product_type == "Фасад":
        c1, c2, c3 = st.columns(3)
        with c1:
            wf = st.number_input("Общая ширина фасада (мм)", min_value=100, value=1500)
            hf = st.number_input("Общая высота фасада (мм)", min_value=100, value=3000)
        with c2:
            nm = st.number_input("Количество вертикальных стоек", min_value=2, value=2)
            nt = st.number_input("Количество уровней ригелей", min_value=1, value=2)
        with c3:
            fill_type = st.radio("Тип заполнения", ["Стеклопакет", "Ламбри (Панель)"])
        
        sections.append({"kind": "facade", "W": wf, "H": hf, "n_m": nm, "n_t": nt, "fill": fill_type, "qty": 1})
    else:
        num_pos = st.number_input("Количество типоразмеров", min_value=1, value=1)
        for i in range(int(num_pos)):
            with st.expander(f"Позиция №{i+1}", expanded=True):
                col_w, col_h, col_q = st.columns(3)
                w = col_w.number_input(f"Ширина W{i+1} (мм)", min_value=100, value=1000, key=f"w_{i}")
                h = col_h.number_input(f"Высота H{i+1} (мм)", min_value=100, value=1500, key=f"h_{i}")
                q = col_q.number_input(f"Количество шт{i+1}", min_value=1, value=1, key=f"q_{i}")
                sections.append({"kind": "standard", "W": w, "H": h, "qty": q})

    # Дополнительные параметры заказа
    st.sidebar.subheader("Дополнительно")
    toning = st.sidebar.selectbox("Тонировка", ["Нет", "Bronze", "Silver", "Grey"])
    assembly = st.sidebar.checkbox("Включить сборку", value=True)
    montage = st.sidebar.checkbox("Включить монтаж", value=False)

    # --- ГЛАВНЫЙ РАСЧЕТ ---
    if st.button("🏗️ ВЫПОЛНИТЬ РАСЧЕТ"):
        total_sum = 0
        total_area = 0
        mat_summary = []

        for s in sections:
            pos_area = (s["W"] * s["H"] / 1000000) * s["qty"]
            total_area += pos_area
            
            # Логика материалов для Фасада
            if s["kind"] == "facade":
                m_len = (s["H"] * s["n_m"] / 1000)
                t_count = (s["n_m"] - 1) * s["n_t"]
                # Ригель "в свету" между стойками
                single_t_len = (s["W"] - (s["n_m"] * 50)) / (s["n_m"] - 1)
                total_t_len = (single_t_len * t_count / 1000)
                
                mat_summary.append({"Наименование": "Стойка фасадная Ruit 50F", "Расход": m_len, "Ед": "м.п."})
                mat_summary.append({"Наименование": "Ригель фасадный Ruit 50F", "Расход": total_t_len, "Ед": "м.п."})
                mat_summary.append({"Наименование": "U-соединитель ригеля", "Расход": t_count * 2, "Ед": "шт"})
                mat_summary.append({"Наименование": "Уплотнитель торцевой ригеля", "Расход": t_count * 2, "Ед": "шт"})
                
                if s["fill"] == "Ламбри (Панель)":
                    mat_summary.append({"Наименование": "Заполнение: Панель Ламбри", "Расход": pos_area, "Ед": "м2"})
            
            # Логика материалов для Дверей (Авто-петли)
            elif "Дверь" in product_type:
                hinges = 3 if s["H"] > 2100 else 2
                mat_summary.append({"Наименование": "Профиль (рама + створка)", "Расход": (s["W"] + s["H"]) * 2 * s["qty"] / 1000, "Ед": "м.п."})
                mat_summary.append({"Наименование": "Петля дверная (усиленная)", "Расход": hinges * s["qty"], "Ед": "шт"})
                mat_summary.append({"Наименование": "Замок дверной в сборе", "Расход": 1 * s["qty"], "Ед": "шт"})

            # Расчет стоимости (упрощенная модель на основе площади)
            base_price = 58000
            if "73C" in profile_system: base_price = 82000
            if "Slim" in profile_system: base_price = 48000
            
            item_price = pos_area * base_price
            if toning != "Нет": item_price *= 1.12
            if montage: item_price += (pos_area * 6500)
            if assembly: item_price += (pos_area * 4000)
            
            total_sum += item_price

        # Вывод результатов
        st.subheader("📊 Результаты расчета")
        st.table(pd.DataFrame(mat_summary))
        
        c_res1, c_res2 = st.columns(2)
        c_res1.metric("Общая площадь остекления", f"{total_area:.2f} м²")
        c_res2.metric("ИТОГО К ОПЛАТЕ", f"{total_sum:,.2f} тенге")

        # Формирование Excel
        order_info = {
            "order_number": order_number, 
            "product_type": product_type, 
            "profile_system": profile_system
        }
        excel_data = build_smeta_workbook(order_info, sections, total_area, total_sum)
        
        st.download_button(
            label="💾 Скачать смету в Excel",
            data=excel_data,
            file_name=f"Axis_Order_{order_number}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
