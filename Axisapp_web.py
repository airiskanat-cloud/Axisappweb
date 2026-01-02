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
# 1. КОНСТАНТЫ И ПОДКЛЮЧЕНИЕ
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
# 2. ЭКСПОРТ В EXCEL (ПО ОБРАЗЦУ ШЕВЧЕНКА)
# =========================================================
def create_excel_report(order_info, items_list, services_df, total_res):
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"
    
    # Стили
    bold = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Шапка компании
    ws.merge_cells('C1:E1')
    ws['C1'] = "ООО «AXIS»"
    ws['C1'].font = Font(bold=True, size=14)
    ws['C2'] = "Город Астана. Тел.: +7 707 504 4040"
    
    ws.append([])
    ws.append(["Заказ №", order_info['no']])
    ws.append(["Тип изделия", order_info['type']])
    ws.append(["Профильная система", order_info['sys']])
    ws.append(["Дата", datetime.now().strftime("%d.%m.%Y")])
    ws.append([])

    # Таблица позиций
    ws.append(["№", "Наименование позиции", "Ширина", "Высота", "Кол-во", "Площадь"])
    for i, item in enumerate(items_list, 1):
        ws.append([i, f"Позиция {i}", item['W'], item['H'], item['qty'], item['area']])

    ws.append([])
    ws.append(["", "", "", "", "ИТОГО К ОПЛАТЕ:", f"{total_res:,.0f} ₸"])
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# =========================================================
# 3. ОСНОВНОЙ ИНТЕРФЕЙС v15 + ДОПОЛНЕНИЯ
# =========================================================
def main():
    st.set_page_config(page_title="Axis Pro GF v20", layout="wide")
    db = load_db()

    if 'auth' not in st.session_state: st.session_state.auth = False

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("🛡️ Вход в Axis Pro GF")
        u = st.text_input("Логин")
        p = st.text_input("Пароль", type="password")
        if st.button("Войти"):
            user = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
            if not user.empty:
                st.session_state.auth = True
                st.rerun()
        return

    st.title("🏗️ Axis Pro GF | Профессиональный калькулятор")

    # --- ВЕРХНЯЯ ПАНЕЛЬ (ПОЛНЫЙ ИНТЕРФЕЙС v15) ---
    with st.sidebar:
        st.header("🏢 Основные настройки")
        order_no = st.text_input("Номер заказа", "Шевченка_001")
        p_type = st.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад", "Тамбур"])
        p_sys = st.selectbox("Профильная система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])
        
        st.markdown("---")
        sp_type = st.selectbox("Тип стеклопакета", ["двойной", "тройной", "энергодвойной", "энерготройной", "Одинарный 4мм"])
        toning = st.checkbox("Тонировка")
        assembly = st.checkbox("Сборка", value=True)
        montage = st.selectbox("Монтаж", ["Нет", "Монтаж", "Демонтаж/Монтаж", "Сложный монтаж"])
        
        st.header("🛒 Дополнительно")
        handle_type = st.selectbox("Тип ручек", ["Нажимная T-MZS70", "Офисная 1000мм", "Офисная 1500мм", "Нет"])
        closer = st.checkbox("Доводчик")

    # --- ДИНАМИЧЕСКИЙ ВВОД ИЗДЕЛИЙ (ДОПОЛНЕНИЕ) ---
    st.subheader("📐 Габариты изделий")
    num_items = st.number_input("Сколько разных конструкций в заказе?", min_value=1, value=1)
    
    all_final_items = []
    
    for i in range(int(num_items)):
        with st.expander(f"Конструкция №{i+1}", expanded=True):
            c1, c2, c3, c4 = st.columns(4)
            W = c1.number_input(f"Ширина W{i+1}, мм", value=1000, key=f"W_{i}")
            H = c2.number_input(f"Высота H{i+1}, мм", value=1500, key=f"H_{i}")
            qty = c3.number_input(f"Кол-во штук{i+1}", value=1, key=f"Q_{i}")
            n_stvor = c4.number_input(f"Кол-во створок{i+1}", value=0, key=f"S_{i}")

            # Детальные габариты (v15)
            g1, g2, g3, g4 = st.columns(4)
            L = g1.number_input(f"LEFT{i+1}", value=0, key=f"L_{i}")
            C = g2.number_input(f"CENTER{i+1}", value=0, key=f"C_{i}")
            R = g3.number_input(f"RIGHT{i+1}", value=0, key=f"R_{i}")
            T = g4.number_input(f"TOP{i+1}", value=0, key=f"T_{i}")

            stvor_data = []
            if n_stvor > 0:
                for s in range(int(n_stvor)):
                    sc1, sc2 = st.columns(2)
                    sw = sc1.number_input(f"Ширина створки {s+1} изд {i+1}", value=600, key=f"sw_{i}_{s}")
                    sh = sc2.number_input(f"Высота створки {s+1} изд {i+1}", value=1200, key=f"sh_{i}_{s}")
                    stvor_data.append({"sw": sw, "sh": sh})

            all_final_items.append({
                "W": W, "H": H, "qty": qty, "L": L, "C": C, "R": R, "T": T, 
                "stvor": stvor_data, "area": (W*H/1000000)*qty, "perim": ((W+H)*2/1000)*qty
            })

    # --- РАСЧЕТ ---
    if st.button("🚀 РАССЧИТАТЬ ПО АЛГОРИТМУ v15"):
        total_mats_cost = 0
        total_area = sum(it['area'] for it in all_final_items)
        total_perim = sum(it['perim'] for it in all_final_items)
        spec_list = []

        # Фильтр материалов (Тип + Система)
        ref1 = db['ref1']
        mats_for_type = ref1[(ref1['Тип изделия'] == p_type) & (ref1['Система профиля'] == p_sys)]

        for item in all_final_items:
            for _, row in mats_for_type.iterrows():
                try:
                    # Контекст для формул (v15)
                    ctx = {
                        "W": item['W'], "H": item['H'], "qty": item['qty'], "count": item['qty'],
                        "L": item['L'], "C": item['C'], "R": item['R'], "T": item['T'],
                        "w_s": item['stvor'][0]['sw'] if item['stvor'] else 0,
                        "h_s": item['stvor'][0]['sh'] if item['stvor'] else 0,
                        "math": math
                    }
                    formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                    fact_res = eval(formula, {"__builtins__": None}, ctx)
                    
                    if fact_res > 0:
                        norma = float(str(row.get('кол-во норм к упаковке', 1)).replace(',', '.'))
                        qty_ship = math.ceil(fact_res / (norma if norma > 0 else 1))
                        price = float(str(row.get('цена за ед', 0)).replace(',', '.'))
                        
                        row_sum = (price * (norma if norma > 0 else 1)) * qty_ship
                        total_mats_cost += row_sum
                        spec_list.append({"Товар": row['Товар'], "Артикул": row['Артикул'], "Кол-во": qty_ship, "Сумма": row_sum})
                except: continue

        # Экономика услуг
        prices_sp = {"двойной": 9000, "тройной": 14000, "энергодвойной": 12000, "энерготройной": 15000, "Одинарный 4мм": 4000}
        sum_sp = prices_sp.get(sp_type, 9000) * total_area
        sum_ton = (2000 * total_area) if toning else 0
        sum_ass = (10000 * total_area) if assembly else 0
        
        mon_p = {"Монтаж": 10000, "Демонтаж/Монтаж": 12000, "Сложный монтаж": 15000, "Нет": 0}
        sum_mon = mon_p.get(montage, 0) * total_area

        # ИТОГО (Твоя формула)
        all_expenses = sum_sp + sum_ton + sum_ass + sum_mon + total_mats_cost
        margin = all_expenses * 0.65
        grand_total = all_expenses + margin

        # Вывод результатов
        st.markdown("---")
        res1, res2, res3 = st.columns(3)
        res1.metric("Общая площадь", f"{total_area:.3f} м2")
        res2.metric("Суммарный периметр", f"{total_perim:.1f} м.п.")
        res3.metric("ИТОГО К ОПЛАТЕ", f"{grand_total:,.0f} ₸")

        st.subheader("🛠️ Смета расходов и услуг")
        serv_df = pd.DataFrame([
            {"Наименование": "Итого материалов", "Сумма": round(total_mats_cost, 0)},
            {"Наименование": "Заполнение (Стеклопакет)", "Сумма": round(sum_sp, 0)},
            {"Наименование": "Тонировка", "Сумма": round(sum_ton, 0)},
            {"Наименование": "Сборка", "Сумма": round(sum_ass, 0)},
            {"Наименование": "Монтаж", "Сумма": round(sum_mon, 0)},
            {"Наименование": "Обеспечение (наценка 65%)", "Сумма": round(margin, 0)}
        ])
        st.table(serv_df)

        # Скачивание
        excel_data = create_excel_report({"no": order_no, "type": p_type, "sys": p_sys}, all_final_items, serv_df, grand_total)
        st.download_button("📥 Скачать Коммерческое предложение (Excel)", data=excel_data, file_name=f"Axis_Offer_{order_no}.xlsx")

if __name__ == "__main__":
    main()
