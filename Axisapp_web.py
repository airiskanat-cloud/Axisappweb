import math
import os
import sys
import shutil
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
# 1. СИСТЕМНЫЕ НАСТРОЙКИ
# =========================================================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Листы Google Таблиц
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================================================
# 2. ПОДКЛЮЧЕНИЕ К ДАННЫМ
# =========================================================
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        # Прямое использование ключа на Render (файл gcp.json)
        creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"❌ Ошибка авторизации: {e}")
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
# 3. МАТЕМАТИЧЕСКАЯ ЛОГИКА
# =========================================================
def calculate_materials(pos, ref3, ref2):
    """Расчет материалов на основе формул из Справочника-3"""
    spec = []
    cost = 0
    mats_ref = ref3[ref3['Тип изделия'] == pos['type']]
    
    context = {
        "W": pos['W'], "H": pos['H'], "qty": pos['qty'],
        "n_m": pos.get('n_m', 0), "n_t": pos.get('n_t', 0),
        "hinges": pos.get('hinges', 2), "is_insert": int(pos.get('is_insert', False)),
        "math": math
    }

    for _, row in mats_ref.iterrows():
        try:
            formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
            qty_mat = eval(formula, {"__builtins__": None}, context)
            if qty_mat > 0:
                price_row = ref2[ref2['Система'] == pos['sys']]
                price = price_row['Цена'].values[0] if not price_row.empty else 0
                sum_mat = qty_mat * price
                cost += sum_mat
                spec.append({
                    "Наименование": row['Наименование'],
                    "Количество": round(qty_mat, 2),
                    "Ед": row['Ед'],
                    "Сумма": round(sum_mat, 0)
                })
        except:
            continue
    return spec, cost

# =========================================================
# 4. ГЕНЕРАЦИЯ КП (EXCEL)
# =========================================================
def build_excel(order_meta, items, grand_total, total_area):
    wb = Workbook()
    ws = wb.active
    ws.title = "КП Axisapp"
    
    # Стили
    bold = Font(bold=True)
    ws['C1'] = "ООО «AXIS»"
    ws['C1'].font = Font(bold=True, size=14)
    ws.append(["Заказ №", order_meta['no']])
    ws.append(["Цвет RAL", order_meta['ral']])
    ws.append([])
    
    headers = ["№", "Тип", "Габариты", "Система", "Кол-во", "Площадь"]
    ws.append(headers)
    for i, p in enumerate(items, 1):
        ws.append([i, p['type'], f"{p['W']}x{p['H']}", p['sys'], p['qty'], round(p['area'], 3)])
    
    ws.append([])
    ws.append(["ИТОГО ПЛОЩАДЬ:", f"{total_area:.3f} м2"])
    ws.append(["ИТОГО К ОПЛАТЕ:", f"{grand_total:,.0f} тенге"])
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# =========================================================
# 5. ИНТЕРФЕЙС STREAMLIT
# =========================================================
def main():
    st.set_page_config(page_title="Axisapp Pro v16", layout="wide")
    db = load_all_data()

    if 'auth' not in st.session_state: st.session_state.auth = False
    if 'items' not in st.session_state: st.session_state.items = []

    # --- ЛОГИН ---
    if not st.session_state.auth:
        st.title("🧱 Вход в систему AXIS")
        u = st.text_input("Логин")
        p = st.text_input("Пароль", type="password")
        if st.button("Войти"):
            user = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
            if not user.empty:
                st.session_state.auth = True
                st.session_state.role = user.iloc[0]['Роль']
                st.rerun()
        return

    # --- ГЛАВНЫЙ ЭКРАН ---
    st.sidebar.title(f"👤 {st.session_state.role}")
    order_no = st.sidebar.text_input("Номер заказа", "2025-001")
    ral = st.sidebar.text_input("RAL", "7024")

    tab1, tab2, tab3 = st.tabs(["📐 Конструктор", "📋 Список позиций", "💰 Расчет"])

    with tab1:
        st.subheader("Добавление изделия")
        col1, col2, col3 = st.columns([2, 2, 1])
        
        # 1) Обновленные типы
        p_type = col1.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 2-х створч.", "Дверь 1 створч.", "Фасад"])
        # 2) Обновленные серии
        p_sys = col2.selectbox("Серия профиля", ["ALG 2030-63C", "ALG 2030-55C", "ALG 2030-73C", "ALG 2030-45C", "ALG 2030-Slim", "Ruit 50F"])
        p_qty = col3.number_input("Кол-во", min_value=1, value=1)

        cW, cH = st.columns(2)
        W = cW.number_input("Ширина (мм)", value=1000)
        H = cH.number_input("Высота (мм)", value=1500)

        # 4) Логика Фасада и каркаса
        n_m, n_t, is_ins = 0, 0, False
        if p_type == "Фасад":
            st.info("📏 Введите габариты каркаса")
            f1, f2 = st.columns(2)
            n_m = f1.number_input("Кол-во стоек", value=2)
            n_t = f2.number_input("Кол-во ригелей", value=1)
        else:
            is_ins = st.checkbox("Вставка в фасад (каркас)", help="Добавляет адаптер 5081")

        if st.button("Добавить в расчет"):
            # Автоматика фурнитуры (3 петли если H > 2100)
            h_count = 3 if H > 2100 and "Дверь" in p_type else 2
            
            st.session_state.items.append({
                "type": p_type, "sys": p_sys, "W": W, "H": H, "qty": p_qty,
                "area": (W * H / 1000000) * p_qty,
                "n_m": n_m, "n_t": n_t, "is_insert": is_ins, "hinges": h_count
            })
            st.success("Позиция добавлена!")

    with tab2:
        if st.session_state.items:
            st.table(pd.DataFrame(st.session_state.items)[['type', 'sys', 'W', 'H', 'qty', 'area']])
            if st.button("Очистить проект"):
                st.session_state.items = []
                st.rerun()

    with tab3:
        if st.session_state.items:
            toning = st.sidebar.checkbox("Тонировка")
            assembly = st.sidebar.checkbox("Сборка", value=True)
            montage = st.sidebar.checkbox("Монтаж")

            total_mats_cost = 0
            full_spec = []

            for item in st.session_state.items:
                spec, cost = calculate_materials(item, db['ref3'], db['ref2'])
                full_spec.extend(spec)
                total_mats_cost += cost

            total_area = sum(i['area'] for i in st.session_state.items)
            glass_cost = total_area * 18000
            if toning: glass_cost += (total_area * 5000)
            
            labor = (total_area * 5000 if assembly else 0) + (total_area * 8000 if montage else 0)
            
            # Коэффициент обеспечения 1.65
            final_sum = (total_mats_cost + glass_cost + labor) * 1.65

            st.metric("ИТОГО К ОПЛАТЕ", f"{final_sum:,.0f} ₸")
            st.write(f"Общая площадь: {total_area:.3f} м2")

            with st.expander("Детальные материалы"):
                st.table(pd.DataFrame(full_spec).groupby("Наименование").sum())

            excel_data = build_excel({"no": order_no, "ral": ral}, st.session_state.items, final_sum, total_area)
            st.download_button("💾 Скачать КП в Excel", data=excel_data, file_name=f"Axis_Offer_{order_no}.xlsx")

if __name__ == "__main__":
    main()
