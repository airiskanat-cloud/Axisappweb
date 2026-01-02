import math
import os
import time
from datetime import datetime
from io import BytesIO

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# =========================================================
# 1. НАСТРОЙКИ
# =========================================================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

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
    return {"ref1": get_df(SHEET_REF1), "users": get_df(SHEET_USERS), "sh": sh}

# =========================================================
# 2. ФУНКЦИЯ ЭКСПОРТА (v15 STYLE)
# =========================================================
def create_excel(order_data, items, services, grand_total, total_area, total_perim):
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"
    
    # Шапка
    ws["C1"] = "ООО «AXIS»"
    ws["C2"] = "Город Астана"
    ws["A7"] = f"Заказ № {order_data['no']}"
    ws["A8"] = f"Тип изделия: {order_data['type']}"
    ws["A9"] = f"Профильная система: {order_data['sys']}"
    
    # Заголовки таблицы
    ws.append(["", "Наименование", "Размеры", "Кол-во", "Сумма"])
    
    for item in items:
        ws.append(["", item['name'], f"{item['W']}x{item['H']}", item['qty'], ""])

    ws.append([])
    ws.append(["", "Общая площадь:", f"{total_area:.3f} м2"])
    ws.append(["", "ИТОГО К ОПЛАТЕ:", f"{grand_total:,.0f} ₸"])
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# =========================================================
# 3. ОСНОВНОЕ ПРИЛОЖЕНИЕ
# =========================================================
def main():
    st.set_page_config(page_title="Axis Pro GF v19", layout="wide")
    db = load_db()

    if 'auth' not in st.session_state: st.session_state.auth = False

    if not st.session_state.auth:
        st.title("🏗️ Axis Pro GF | Вход")
        u, p = st.text_input("Логин"), st.text_input("Пароль", type="password")
        if st.button("Войти"):
            check = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
            if not check.empty:
                st.session_state.auth = True
                st.rerun()
        return

    st.title("🏗️ Axis Pro GF | Расчетный комплекс")

    with st.sidebar:
        st.header("Настройки")
        order_no = st.text_input("Заказ №", "Шевченка-01")
        p_type = st.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_sys = st.selectbox("Система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])
        
        st.markdown("---")
        sp_price = st.number_input("Цена стеклопакета (м2)", value=9000)
        toning = st.checkbox("Тонировка (2000 ₸/м2)")
        assembly = st.checkbox("Сборка (10000 ₸/м2)", value=True)
        montage_val = st.number_input("Монтаж (₸/м2)", value=10000)

    # ДИНАМИЧЕСКИЕ ИЗДЕЛИЯ
    num_items = st.number_input("Количество разных изделий в заказе", min_value=1, value=1)
    all_items_data = []

    for i in range(int(num_items)):
        st.markdown(f"### Позиция №{i+1}")
        col1, col2, col3 = st.columns(3)
        W = col1.number_input(f"Ширина изделия {i+1}", value=1000, key=f"W_{i}")
        H = col2.number_input(f"Высота изделия {i+1}", value=1500, key=f"H_{i}")
        qty = col3.number_input(f"Кол-во штук {i+1}", value=1, key=f"Q_{i}")
        
        # СТВОРКИ ВНУТРИ ИЗДЕЛИЯ
        num_s = st.number_input(f"Кол-во створок в изделии {i+1}", min_value=0, value=0, key=f"NS_{i}")
        stvor_list = []
        if num_s > 0:
            for s in range(int(num_s)):
                sc1, sc2 = st.columns(2)
                sw = sc1.number_input(f"Ширина створки {s+1} (изд {i+1})", value=600, key=f"sw_{i}_{s}")
                sh = sc2.number_input(f"Высота створки {s+1} (изд {i+1})", value=1200, key=f"sh_{i}_{s}")
                stvor_list.append({"sw": sw, "sh": sh})
        
        all_items_data.append({"W": W, "H": H, "qty": qty, "stvor": stvor_list})

    if st.button("🚀 ПОЛНЫЙ РАСЧЕТ"):
        total_area = 0
        total_perim = 0
        total_mats_cost = 0
        spec_final = []

        for item in all_items_data:
            item_area = (item['W'] * item['H'] / 1000000) * item['qty']
            item_perim = ((item['W'] + item['H']) * 2 / 1000) * item['qty']
            total_area += item_area
            total_perim += item_perim

            # РАСЧЕТ МАТЕРИАЛОВ (ФИЛЬТР ТИП + СИСТЕМА)
            mats = db['ref1'][(db['ref1']['Тип изделия'] == p_type) & (db['ref1']['Система профиля'] == p_sys)]
            
            for _, row in mats.iterrows():
                try:
                    ctx = {"W": item['W'], "H": item['H'], "qty": item['qty'], "count": item['qty'], "math": math}
                    # Если есть створки, берем первую для формул w_s/h_s
                    if item['stvor']:
                        ctx["w_s"] = item['stvor'][0]['sw']
                        ctx["h_s"] = item['stvor'][0]['sh']
                    
                    formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                    fact = eval(formula, {"__builtins__": None}, ctx)
                    
                    if fact > 0:
                        norma = float(str(row.get('кол-во норм к упаковке', 1)).replace(',', '.'))
                        qty_ship = math.ceil(fact / norma)
                        price = float(str(row.get('цена за ед', 0)).replace(',', '.'))
                        sum_row = (price * norma) * qty_ship
                        total_mats_cost += sum_row
                        spec_final.append({"Товар": row['Товар'], "Кол-во": qty_ship, "Сумма": sum_row})
                except: continue

        # ЭКОНОМИКА
        sum_sp = total_area * sp_price
        sum_ton = (total_area * 2000) if toning else 0
        sum_ass = (total_area * 10000) if assembly else 0
        sum_mon = total_area * montage_val
        
        all_costs = sum_sp + sum_ton + sum_ass + sum_mon + total_mats_cost
        margin = all_costs * 0.65
        grand_total = all_costs + margin

        # ВЫВОД
        st.header("📊 Результат")
        c1, c2, c3 = st.columns(3)
        c1.metric("Общая площадь", f"{total_area:.3f} м2")
        c2.metric("Общий периметр", f"{total_perim:.1f} м.п.")
        c3.metric("ИТОГО К ОПЛАТЕ", f"{grand_total:,.0f} ₸")

        st.subheader("🛠️ Смета услуг")
        serv_df = pd.DataFrame([
            {"Услуга": "Итого материалов", "Сумма": total_mats_cost},
            {"Услуга": "Стеклопакет", "Сумма": sum_sp},
            {"Услуга": "Тонировка", "Сумма": sum_ton},
            {"Услуга": "Сборка", "Сумма": sum_ass},
            {"Услуга": "Монтаж", "Сумма": sum_mon},
            {"Услуга": "Обеспечение (65%)", "Сумма": margin}
        ])
        st.table(serv_df)

        # КНОПКА EXCEL
        excel_file = create_excel({"no": order_no, "type": p_type, "sys": p_sys}, all_items_data, serv_df, grand_total, total_area, total_perim)
        st.download_button("📥 Скачать Коммерческое предложение", data=excel_file, file_name=f"Axis_Offer_{order_no}.xlsx")

if __name__ == "__main__":
    main()
