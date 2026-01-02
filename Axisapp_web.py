import math
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook

# Настройки остаются прежними
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=300)
def load_db():
    client = get_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    def get_df(name):
        df = pd.DataFrame(sh.worksheet(name).get_all_records())
        df.columns = df.columns.str.strip() # Чистим заголовки
        return df
    return {"ref1": get_df("СПРАВОЧНИК -1"), "users": get_df("ПОЛЬЗОВАТЕЛИ"), "sh": sh}

def main():
    st.set_page_config(page_title="Axis Pro GF v21", layout="wide")
    db = load_db()

    # --- АВТОРИЗАЦИЯ (v15) ---
    if 'auth' not in st.session_state: st.session_state.auth = False
    if not st.session_state.auth:
        st.title("🛡️ Вход в Axis Pro GF")
        u, p = st.text_input("Логин"), st.text_input("Пароль", type="password")
        if st.button("Войти"):
            check = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
            if not check.empty:
                st.session_state.auth = True
                st.rerun()
        return

    # --- ИНТЕРФЕЙС (v15 ПОЛНЫЙ) ---
    with st.sidebar:
        st.header("📋 Заказ")
        order_no = st.text_input("Номер заказа", "001")
        p_type = st.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад", "Тамбур"])
        p_sys = st.selectbox("Система профиля", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])
        
        st.markdown("---")
        # Цены из твоего сообщения
        p_glass = st.number_input("Цена СП (м2)", value=9000)
        p_assembly = st.number_input("Цена Сборки (м2)", value=10000)
        p_montage = st.number_input("Цена Монтажа (м2)", value=10000)
        toning = st.checkbox("Тонировка (2000 ₸)")

    # ВВОД ПОЗИЦИЙ
    num_pos = st.number_input("Кол-во типоразмеров (позиций)", min_value=1, value=1)
    all_positions = []

    for i in range(int(num_pos)):
        with st.expander(f"Позиция №{i+1}", expanded=True):
            col1, col2, col3 = st.columns(3)
            W = col1.number_input(f"Ширина W{i+1}", value=1000, key=f"w{i}")
            H = col2.number_input(f"Высота H{i+1}", value=1500, key=f"h{i}")
            qty = col3.number_input(f"Кол-во изделий N{i+1}", value=1, key=f"q{i}")
            
            # Доп габариты (L, C, R, T)
            c1, c2, c3, c4 = st.columns(4)
            L = c1.number_input(f"LEFT {i+1}", value=0, key=f"l{i}")
            C = c2.number_input(f"CENTER {i+1}", value=0, key=f"c{i}")
            R = c3.number_input(f"RIGHT {i+1}", value=0, key=f"r{i}")
            T = c4.number_input(f"TOP {i+1}", value=0, key=f"t{i}")

            all_positions.append({"W":W, "H":H, "qty":qty, "L":L, "C":C, "R":R, "T":T})

    if st.button("🚀 РАССЧИТАТЬ МАТЕРИАЛЫ"):
        total_area = 0
        total_mats_cost = 0
        spec_output = []

        # 1. ПОДГОТОВКА СПРАВОЧНИКА (ОЧИСТКА ДЛЯ ФИЛЬТРА)
        ref1 = db['ref1'].copy()
        ref1['Тип изделия'] = ref1['Тип изделия'].astype(str).str.strip().str.lower()
        ref1['Система профиля'] = ref1['Система профиля'].astype(str).str.strip().str.lower()
        
        # 2. ФИЛЬТРАЦИЯ (МЯГКАЯ)
        search_type = p_type.lower().strip()
        search_sys = p_sys.lower().strip()
        
        mats_filtered = ref1[
            (ref1['Тип изделия'] == search_type) & 
            ((ref1['Система профиля'] == search_sys) | (ref1['Система профиля'] == ""))
        ]

        # 3. ЦИКЛ ПО ПОЗИЦИЯМ
        for pos in all_positions:
            pos_area = (pos['W'] * pos['H'] / 1000000) * pos['qty']
            total_area += pos_area
            
            # Контекст переменных для Справочника-1
            ctx = {
                "W": pos['W'], "H": pos['H'], "count": pos['qty'], "qty": pos['qty'],
                "L": pos['L'], "C": pos['C'], "R": pos['R'], "T": pos['T'],
                "math": math
            }

            for _, row in mats_filtered.iterrows():
                try:
                    formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                    fact_rashod = eval(formula, {"__builtins__": None}, ctx)
                    
                    if fact_rashod > 0:
                        norma = float(str(row.get('кол-во норм к упаковке', 1)).replace(',', '.'))
                        if norma <= 0: norma = 1
                        
                        qty_ship = math.ceil(fact_rashod / norma)
                        price = float(str(row.get('цена за ед', 0)).replace(',', '.'))
                        
                        row_sum = (price * norma) * qty_ship
                        total_mats_cost += row_sum
                        spec_output.append({
                            "Товар": row['Товар'], "Артикул": row['Артикул'], 
                            "Расход": round(fact_rashod, 2), "Упак": qty_ship, "Сумма": row_sum
                        })
                except: continue

        # 4. ИТОГИ (ТВОЯ ФОРМУЛА)
        sum_services = (p_glass + p_assembly + p_montage + (2000 if toning else 0)) * total_area
        all_costs = sum_services + total_mats_cost
        margin = all_costs * 0.65
        grand_total = all_costs + margin

        # 5. ВЫВОД
        st.header("📊 Результаты (Axis Pro GF)")
        c1, c2, c3 = st.columns(3)
        c1.metric("Общая площадь (м2)", f"{total_area:.3f}")
        c2.metric("Себестоимость", f"{all_costs:,.0f} ₸")
        c3.metric("ИТОГО К ОПЛАТЕ", f"{grand_total:,.0f} ₸")

        st.subheader("📋 Спецификация материалов (Все позиции)")
        if spec_output:
            # Группируем, чтобы одинаковые товары не дублировались
            final_df = pd.DataFrame(spec_output).groupby(["Товар", "Артикул"]).sum().reset_index()
            st.dataframe(final_df)
        else:
            st.warning("Материалы не найдены. Проверьте соответствие 'Тип изделия' и 'Система профиля' в Справочнике.")

if __name__ == "__main__":
    main()
