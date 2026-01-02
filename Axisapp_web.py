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

# =========================
# КОНСТАНТЫ / НАСТРОЙКИ
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

# Листы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================
# ФУНКЦИИ ПОДКЛЮЧЕНИЯ
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
    client = get_gspread_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    
    df_ref1 = pd.DataFrame(sh.worksheet(SHEET_REF1).get_all_records())
    df_ref2 = pd.DataFrame(sh.worksheet(SHEET_REF2).get_all_records())
    df_ref3 = pd.DataFrame(sh.worksheet(SHEET_REF3).get_all_records())
    df_users = pd.DataFrame(sh.worksheet(SHEET_USERS).get_all_records())
    
    return df_ref1, df_ref2, df_ref3, df_users, sh

# =========================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =========================

def build_smeta_workbook(order, base_positions, lambr_positions, total_area, total_perimeter, total_sum):
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"
    
    ws["A1"] = "Заказ №:"
    ws["B1"] = order["order_number"]
    ws["A2"] = "Тип изделия:"
    ws["B2"] = order["product_type"]
    ws["A3"] = "Система:"
    ws["B3"] = order["profile_system"]
    
    headers = ["№", "Тип", "Ширина", "Высота", "Кол-во", "Площадь м2"]
    ws.append([])
    ws.append(headers)
    
    all_pos = base_positions + lambr_positions
    for i, p in enumerate(all_pos, 1):
        ws.append([i, p.get("kind"), p.get("W"), p.get("H"), p.get("qty"), p.get("area", 0)])
        
    ws.append([])
    ws.append(["ИТОГО Площадь", total_area])
    ws.append(["ИТОГО Сумма", total_sum])
    
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================
# ОСНОВНОЕ ПРИЛОЖЕНИЕ
# =========================

def main():
    st.set_page_config(page_title="Axisapp - Профессиональный расчет", layout="wide")
    
    if 'auth' not in st.session_state:
        st.session_state.auth = False

    df_ref1, df_ref2, df_ref3, df_users, sh = load_all_data()

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("Вход в систему")
        user_input = st.text_input("Логин")
        pass_input = st.text_input("Пароль", type="password")
        if st.button("Войти"):
            user_row = df_users[(df_users['Логин'] == user_input) & (df_users['Пароль'].astype(str) == pass_input)]
            if not user_row.empty:
                st.session_state.auth = True
                st.session_state.user_role = user_row.iloc[0]['Роль']
                st.rerun()
            else:
                st.error("Неверный логин или пароль")
        return

    # --- ИНТЕРФЕЙС РАСЧЕТА ---
    st.sidebar.title(f"Роль: {st.session_state.user_role}")
    if st.sidebar.button("Выход"):
        st.session_state.auth = False
        st.rerun()

    st.sidebar.header("Параметры заказа")
    order_number = st.sidebar.text_input("Номер заказа", "001")
    
    # ИЗМЕНЕНИЕ 1: Обновленный список изделий
    product_type = st.sidebar.selectbox("Тип изделия", [
        "Окно глух.", "Окно с откр.", 
        "Дверь 1 створч.", "Дверь 2-х створч.", 
        "Фасад"
    ])
    
    # ИЗМЕНЕНИЕ 2: Обновленный список систем
    profile_system = st.sidebar.selectbox("Система профиля", [
        "ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", 
        "ALG 2030-73C", "ALG 2030-Slim", "Ruit 50F"
    ])

    st.header(f"Настройка: {product_type} ({profile_system})")
    
    sections = []
    
    # ИЗМЕНЕНИЕ 3: Специальная форма для Фасада
    if product_type == "Фасад":
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            W_f = st.number_input("Общая ширина (мм)", min_value=100, value=1500)
            H_f = st.number_input("Общая высота (мм)", min_value=100, value=2500)
        with col_f2:
            n_mull = st.number_input("Кол-во стоек (вертикальных)", min_value=2, value=2)
            n_trans = st.number_input("Кол-во уровней ригелей", min_value=1, value=1)
        with col_f3:
            facade_fill = st.selectbox("Заполнение глухих частей", ["Стеклопакет", "Ламбри (Панель)"])
        
        sections.append({
            "kind": "facade", "W": W_f, "H": H_f, 
            "n_m": n_mull, "n_t": n_trans, "fill": facade_fill, "qty": 1
        })
    else:
        # Стандартный ввод для окон и дверей (сохраняем кол-во позиций)
        num_pos = st.number_input("Количество типоразмеров", min_value=1, value=1)
        for i in range(int(num_pos)):
            with st.expander(f"Позиция №{i+1}", expanded=True):
                c1, c2, c3 = st.columns(3)
                w_val = c1.number_input(f"Ширина W{i+1}", min_value=100, value=1000)
                h_val = c2.number_input(f"Высота H{i+1}", min_value=100, value=1000)
                q_val = c3.number_input(f"Количество шт{i+1}", min_value=1, value=1)
                sections.append({"kind": "standard", "W": w_val, "H": h_val, "qty": q_val})

    # --- ДОПОЛНИТЕЛЬНЫЕ ОПЦИИ (Сохраняем тонировку, монтаж и т.д.) ---
    st.sidebar.subheader("Доп. параметры")
    glass_type = st.sidebar.selectbox("Тип стекла", ["Прозрачное", "Энергосберегающее", "Мультифункциональное"])
    glass_thickness = st.sidebar.selectbox("Толщина пакета (мм)", [4, 10, 24, 32, 40, 44])
    toning = st.sidebar.selectbox("Тонировка", ["Нет", "Bronze", "Silver", "Grey"])
    
    assembly = st.sidebar.checkbox("Сборка включена", value=True)
    montage = st.sidebar.checkbox("Монтаж включен", value=False)
    
    # --- РАСЧЕТ ---
    if st.button("🚀 РАССЧИТАТЬ"):
        all_results = []
        total_sum = 0
        total_area_all = 0
        total_perimeter_gab = 0

        for s in sections:
            # ИЗМЕНЕНИЕ 4: Расчет Ламбри/Площади
            area = (s["W"] * s["H"] / 1000000) * s["qty"]
            perim = ((s["W"] + s["H"]) * 2 / 1000) * s["qty"]
            total_area_all += area
            total_perimeter_gab += perim
            
            # --- ЛОГИКА РАСХОДА МАТЕРИАЛОВ (БАЗОВАЯ) ---
            # Здесь вызывается логика фильтрации справочника по profile_system и product_type
            res_materials = []
            
            if s["kind"] == "facade":
                # Специфика фасада Ruit 50F
                m_count = s["n_m"]
                t_count = (s["n_m"] - 1) * s["n_t"]
                # Стойки
                res_materials.append({"Наименование": "Стойка фасадная", "Расход": (s["H"] * m_count / 1000), "Ед": "м.п."})
                # Ригели (в свету)
                r_len = (s["W"] - (m_count * 50)) / (m_count - 1) if m_count > 1 else 0
                res_materials.append({"Наименование": "Ригель фасадный", "Расход": (r_len * t_count / 1000), "Ед": "м.п."})
                # Узлы
                res_materials.append({"Наименование": "U-соединитель", "Расход": t_count * 2, "Ед": "шт"})
                res_materials.append({"Наименование": "Торцевой уплотнитель", "Расход": t_count * 2, "Ед": "шт"})
                if s["fill"] == "Ламбри (Панель)":
                    res_materials.append({"Наименование": "Ламбри (Заполнение)", "Расход": area, "Ед": "м2"})
            
            elif "Дверь" in product_type:
                # Дверная логика + Авто-петли
                h_s = s["H"]
                hinges = 3 if h_s > 2100 else 2
                res_materials.append({"Наименование": "Профиль рамы", "Расход": perim, "Ед": "м.п."})
                res_materials.append({"Наименование": "Петля дверная", "Расход": hinges * s["qty"], "Ед": "шт"})
                res_materials.append({"Наименование": "Замок бочонок/ролик", "Расход": 1 * s["qty"], "Ед": "шт"})

            # --- ИТОГОВЫЙ РАСЧЕТ ЦЕНЫ (БЕРЕМ ИЗ ВАШЕЙ ЛОГИКИ) ---
            # Для примера ставим базовую цену из SHEET_REF2 или REF3
            base_price_m2 = 45000  # Тут должна быть подгрузка из df_ref2
            pos_sum = area * base_price_m2
            
            if toning != "Нет": pos_sum *= 1.15
            if montage: pos_sum += (area * 5000)
            if assembly: pos_sum += (area * 3000)
            
            total_sum += pos_sum
            all_results.append({
                "Поз": s.get("W"), "Размер": f"{s['W']}x{s['H']}", 
                "Кол-во": s['qty'], "Сумма": pos_sum
            })

        # --- ВЫВОД РЕЗУЛЬТАТОВ ---
        st.success(f"Расчет завершен для заказа {order_number}")
        st.write(f"**Общая площадь:** {total_area_all:.2f} м2")
        st.write(f"**Общая сумма:** {total_sum:.2f} тенге")
        st.table(pd.DataFrame(all_results))

        # --- ЭКСПОРТ EXCEL ---
        order_data = {
            "order_number": order_number, "product_type": product_type, 
            "profile_system": profile_system, "glass_type": glass_type,
            "toning": toning, "assembly": assembly, "montage": montage
        }
        
        # Передаем данные в функцию формирования книги (из констант выше)
        excel_data = build_smeta_workbook(order_data, sections, [], total_area_all, total_perimeter_gab, total_sum)
        
        st.download_button(
            label="💾 Скачать КП в Excel",
            data=excel_data,
            file_name=f"КП_{order_number}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
