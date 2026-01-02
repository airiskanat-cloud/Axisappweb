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
from openpyxl.styles import Font, Alignment, Border, Side

# =========================
# КОНСТАНТЫ
# =========================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"
SHEET_REF1 = "СПРАВОЧНИК -1"
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

# =========================
# ИНТЕРФЕЙС И ЛОГИКА
# =========================
def main():
    st.set_page_config(page_title="Axis Pro GF v22", layout="wide")
    db = load_db()

    if 'auth' not in st.session_state: st.session_state.auth = False
    if 'order_items' not in st.session_state: st.session_state.order_items = [] # Список всех позиций заказа

    # --- АВТОРИЗАЦИЯ v15 ---
    if not st.session_state.auth:
        st.title("🏗️ Axis Pro GF | Вход")
        u, p = st.text_input("Логин"), st.text_input("Пароль", type="password")
        if st.button("Войти"):
            check = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
            if not check.empty:
                st.session_state.auth, st.session_state.role = True, check.iloc[0]['Роль']
                st.rerun()
        return

    st.title(f"🏗️ Axis Pro GF | Заказ: {st.session_state.get('order_no', 'Новый')}")

    # --- СИНХРОННЫЕ НАСТРОЙКИ (SIDEBAR) ---
    with st.sidebar:
        st.header("🏢 Параметры заказа")
        order_no = st.text_input("Номер заказа", value="001")
        st.session_state.order_no = order_no
        p_type = st.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад", "Тамбур"])
        p_sys = st.selectbox("Система профиля", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])
        
        st.markdown("---")
        sp_price = st.number_input("Стеклопакет (₸/м2)", value=9000)
        toning = st.checkbox("Тонировка (2000 ₸/м2)")
        assembly = st.checkbox("Сборка (10000 ₸/м2)", value=True)
        montage = st.number_input("Монтаж (₸/м2)", value=10000)

    # --- ДОБАВЛЕНИЕ ПОЗИЦИИ (ЛОГИКА v15) ---
    with st.expander("➕ Добавить новую позицию (окно/дверь)", expanded=True):
        c1, c2, c3 = st.columns(3)
        W = c1.number_input("Ширина W, мм", value=1000)
        H = c2.number_input("Высота H, мм", value=1500)
        qty = c3.number_input("Кол-во изделий, шт", value=1, min_value=1)

        g1, g2, g3, g4 = st.columns(4)
        L, C, R, T = g1.number_input("LEFT"), g2.number_input("CENTER"), g3.number_input("RIGHT"), g4.number_input("TOP")

        n_stvor = st.number_input("Количество створок в этом изделии", min_value=0, value=0)
        sashes = []
        for s in range(int(n_stvor)):
            sc1, sc2 = st.columns(2)
            sw = sc1.number_input(f"Ширина створки {s+1}", value=600, key=f"sw_{s}")
            sh = sc2.number_input(f"Высота створки {s+1}", value=1200, key=f"sh_{s}")
            sashes.append({"sw": sw, "sh": sh})

        if st.button("✅ Добавить в список расчета"):
            item_area = (W * H / 1000000) * qty
            item_perim = ((W + H) * 2 / 1000) * qty
            st.session_state.order_items.append({
                "type": p_type, "sys": p_sys, "W": W, "H": H, "qty": qty,
                "L": L, "C": C, "R": R, "T": T, "sashes": sashes,
                "area": item_area, "perim": item_perim
            })
            st.success(f"Позиция добавлена! Текущая площадь заказа: {sum(it['area'] for it in st.session_state.order_items):.3f} м2")

    # --- ТАБЛИЦА ТЕКУЩИХ ПОЗИЦИЙ ---
    if st.session_state.order_items:
        st.subheader("📋 Состав заказа")
        df_order = pd.DataFrame(st.session_state.order_items)
        st.table(df_order[["type", "sys", "W", "H", "qty", "area"]])
        if st.button("🗑️ Очистить заказ"):
            st.session_state.order_items = []
            st.rerun()

        # --- ИТОГОВЫЙ РАСЧЕТ ---
        if st.button("🚀 РАССЧИТАТЬ ВЕСЬ ЗАКАЗ"):
            total_mats_cost = 0
            total_area = sum(it['area'] for it in st.session_state.order_items)
            total_perim = sum(it['perim'] for it in st.session_state.order_items)
            
            # Фильтр Справочника-1 (Жесткий: Тип + Система)
            ref1 = db['ref1']
            
            detailed_results = []
            for it in st.session_state.order_items:
                # Фильтруем материалы именно под эту позицию
                mats = ref1[(ref1['Тип изделия'] == it['type']) & (ref1['Система профиля'] == it['sys'])]
                
                # Контекст для формул
                ctx = {
                    "W": it['W'], "H": it['H'], "count": it['qty'], "qty": it['qty'],
                    "L": it['L'], "C": it['C'], "R": it['R'], "T": it['T'],
                    "w_s": it['sashes'][0]['sw'] if it['sashes'] else 0,
                    "h_s": it['sashes'][0]['sh'] if it['sashes'] else 0,
                    "math": math
                }

                for _, row in mats.iterrows():
                    try:
                        formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                        fact_res = eval(formula, {"__builtins__": None}, ctx)
                        if fact_res > 0:
                            norma = float(str(row.get('кол-во норм к упаковке', 1)).replace(',', '.'))
                            qty_ship = math.ceil(fact_res / (norma if norma > 0 else 1))
                            price = float(str(row.get('цена за ед', 0)).replace(',', '.'))
                            
                            row_sum = (price * (norma if norma > 0 else 1)) * qty_ship
                            total_mats_cost += row_sum
                            detailed_results.append({"Товар": row['Товар'], "Кол-во": qty_ship, "Сумма": row_sum})
                    except: continue

            # Экономика услуг
            sum_sp = total_area * sp_price
            sum_ton = (total_area * 2000) if toning else 0
            sum_ass = (total_area * 10000) if assembly else 0
            sum_mon = total_area * montage
            
            # Итоговая формула (твоя: расходы + 65%)
            base_costs = sum_sp + sum_ton + sum_ass + sum_mon + total_mats_cost
            margin = base_costs * 0.65
            grand_total = base_costs + margin

            # Вывод
            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            c1.metric("ОБЩАЯ ПЛОЩАДЬ", f"{total_area:.3f} м2")
            c2.metric("ОБЩИЙ ПЕРИМЕТР", f"{total_perim:.1f} м.п.")
            c3.metric("ИТОГО К ОПЛАТЕ", f"{grand_total:,.0f} ₸")

            st.subheader("🛠️ Смета услуг и материалов")
            st.table(pd.DataFrame([
                {"Наименование": "Материалы (Справочник-1)", "Сумма": f"{total_mats_cost:,.0f}"},
                {"Наименование": "Стеклопакеты", "Сумма": f"{sum_sp:,.0f}"},
                {"Наименование": "Тонировка / Сборка / Монтаж", "Сумма": f"{sum_ton+sum_ass+sum_mon:,.0f}"},
                {"Наименование": "ОБЕСПЕЧЕНИЕ (65%)", "Сумма": f"{margin:,.0f}"}
            ]))



if __name__ == "__main__":
    main()
