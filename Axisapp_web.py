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
    try:
        client = get_client()
        sh = client.open_by_key(GSPREAD_SHEET_ID)
        def get_df(name):
            data = sh.worksheet(name).get_all_records()
            df = pd.DataFrame(data)
            df.columns = df.columns.str.strip()
            return df
        return {"ref1": get_df(SHEET_REF1), "users": get_df(SHEET_USERS), "sh": sh}
    except Exception as e:
        st.error(f"Ошибка загрузки базы: {e}")
        return None

# =========================================================
# 2. ФУНКЦИЯ ЭКСПОРТА (ОБРАЗЕЦ ШЕВЧЕНКА)
# =========================================================
def create_excel_axis(order_info, positions, services_df, grand_total, total_area, total_perim):
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"
    
    # Стили
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center')
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Шапка
    ws.merge_cells('C1:E1')
    ws['C1'] = "ООО «AXIS»"
    ws['C1'].font = Font(bold=True, size=14)
    ws['C1'].alignment = center_align
    ws['C2'] = "Город Астана"
    ws['C2'].alignment = center_align
    ws['C3'] = "Тел.: +7 707 504 4040"
    ws['C3'].alignment = center_align
    
    ws.append([])
    ws.append(["Коммерческое предложение"])
    ws[ws.max_row][0].font = Font(bold=True, size=12)
    
    ws.append(["Заказ №:", order_info['no']])
    ws.append(["Тип изделия:", order_info['type']])
    ws.append(["Профильная система:", order_info['sys']])
    ws.append(["Дата:", datetime.now().strftime("%d.%m.%Y")])
    ws.append([])

    # Таблица позиций
    headers = ["№", "Наименование позиции", "Ширина (мм)", "Высота (мм)", "Кол-во (шт)", "Площадь (м2)"]
    ws.append(headers)
    for cell in ws[ws.max_row]:
        cell.font = bold_font
        cell.border = border

    for i, pos in enumerate(positions, 1):
        ws.append([i, f"Позиция {i}", pos['W'], pos['H'], pos['qty'], pos['area']])
        for cell in ws[ws.max_row]:
            cell.border = border

    ws.append([])
    ws.append(["", "", "", "", "Общая площадь:", f"{total_area:.3f} м2"])
    ws.append(["", "", "", "", "Общий периметр:", f"{total_perim:.1f} м.п."])
    ws.append(["", "", "", "", "ИТОГО К ОПЛАТЕ:", f"{grand_total:,.0f} ₸"])
    ws[ws.max_row][-1].font = bold_font

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# =========================================================
# 3. ОСНОВНОЙ ИНТЕРФЕЙС
# =========================================================
def main():
    st.set_page_config(page_title="Axis Pro GF v20", layout="wide")
    db = load_db()
    if not db: return

    if 'auth' not in st.session_state: st.session_state.auth = False

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("🛡️ Axis Pro GF | Вход")
        u = st.text_input("Логин")
        p = st.text_input("Пароль", type="password")
        if st.button("Войти"):
            users = db['users']
            if not users.empty:
                check = users[(users['Логин'] == u) & (users['Пароль'].astype(str) == p)]
                if not check.empty:
                    st.session_state.auth = True
                    st.rerun()
            st.error("Ошибка входа")
        return

    st.title("🏗️ Axis Pro GF | Профессиональный расчет")

    # --- ПАНЕЛЬ НАСТРОЕК ---
    with st.sidebar:
        st.header("🏢 Настройки заказа")
        order_no = st.text_input("Номер заказа", "001")
        p_type = st.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад", "Тамбур"])
        p_sys = st.selectbox("Профильная система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])
        
        st.markdown("---")
        st.subheader("Цены за м² (₸)")
        p_glass = st.number_input("Стеклопакет", value=9000)
        p_toning = st.number_input("Тонировка", value=2000)
        p_assembly = st.number_input("Сборка", value=10000)
        p_montage = st.number_input("Монтаж", value=10000)
        
        st.markdown("---")
        toning_enabled = st.checkbox("Применить тонировку")
        assembly_enabled = st.checkbox("Применить сборку", value=True)
        montage_enabled = st.checkbox("Применить монтаж", value=True)

    # --- ВВОД ПОЗИЦИЙ ---
    num_positions = st.number_input("Количество позиций в заказе", min_value=1, value=1)
    all_positions = []

    for i in range(int(num_positions)):
        with st.expander(f"📦 Позиция №{i+1}", expanded=True):
            col1, col2, col3 = st.columns(3)
            W = col1.number_input(f"Ширина изделия {i+1}, мм", value=1000, key=f"W_{i}")
            H = col2.number_input(f"Высота изделия {i+1}, мм", value=1500, key=f"H_{i}")
            qty = col3.number_input(f"Количество шт {i+1}", value=1, key=f"Q_{i}")

            col4, col5, col6, col7 = st.columns(4)
            L = col4.number_input(f"LEFT {i+1}", value=0, key=f"L_{i}")
            C = col5.number_input(f"CENTER {i+1}", value=0, key=f"C_{i}")
            R = col6.number_input(f"RIGHT {i+1}", value=0, key=f"R_{i}")
            T = col7.number_input(f"TOP {i+1}", value=0, key=f"T_{i}")

            n_stvor = st.number_input(f"Кол-во створок {i+1}", min_value=0, value=1 if "откр" in p_type else 0, key=f"NS_{i}")
            sashes = []
            if n_stvor > 0:
                st.markdown(f"**Габариты створок для поз. {i+1}:**")
                for s in range(int(n_stvor)):
                    sc1, sc2 = st.columns(2)
                    sw = sc1.number_input(f"Ширина створки {s+1}, мм", value=600, key=f"sw_{i}_{s}")
                    sh = sc2.number_input(f"Высота створки {s+1}, мм", value=1200, key=f"sh_{i}_{s}")
                    sashes.append({"sw": sw, "sh": sh})

            all_positions.append({
                "W": W, "H": H, "qty": qty, "L": L, "C": C, "R": R, "T": T,
                "n_stvor": n_stvor, "sashes": sashes,
                "area": (W * H / 1000000) * qty,
                "perim": ((W + H) * 2 / 1000) * qty
            })

    # --- РАСЧЕТ ---
    if st.button("🚀 ВЫПОЛНИТЬ РАСЧЕТ"):
        total_mats_sum = 0
        total_area = sum(p['area'] for p in all_positions)
        total_perim = sum(p['perim'] for p in all_positions)
        
        # 1. Фильтр материалов (Тип + Система)
        ref1 = db['ref1']
        mats_filtered = ref1[
            (ref1['Тип изделия'].astype(str).str.strip() == p_type) & 
            ((ref1['Система профиля'].astype(str).str.strip() == p_sys) | (ref1['Система профиля'].astype(str).str.strip() == ""))
        ]

        if mats_filtered.empty:
            st.warning("⚠️ Материалы для данного типа и системы не найдены в Справочнике-1.")

        detailed_mats = []
        for pos in all_positions:
            # Для формул переводим размеры в МЕТРЫ
            ctx = {
                "W": pos['W'] / 1000, "H": pos['H'] / 1000,
                "qty": pos['qty'], "count": pos['qty'],
                "L": pos['L'] / 1000, "C": pos['C'] / 1000, "R": pos['R'] / 1000, "T": pos['T'] / 1000,
                "math": math
            }
            # Створки (суммируем габариты для формул периметра)
            ctx["w_s"] = sum(s['sw'] for s in pos['sashes']) / 1000 if pos['sashes'] else 0
            ctx["h_s"] = sum(s['sh'] for s in pos['sashes']) / 1000 if pos['sashes'] else 0

            for _, row in mats_filtered.iterrows():
                try:
                    formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                    fact_res = eval(formula, {"__builtins__": None}, ctx)
                    
                    if fact_res > 0:
                        norma_str = str(row.get('кол-во норм к упаковке', 1)).replace(',', '.')
                        norma = float(norma_str) if norma_str else 1.0
                        if norma <= 0: norma = 1.0
                        
                        # Кол-во к отгрузке (упаковки)
                        qty_ship = math.ceil(fact_res / norma)
                        
                        price_str = str(row.get('цена за ед', 0)).replace(',', '.')
                        price = float(price_str) if price_str else 0.0
                        
                        # Сумма = (Цена за ед * Норма упаковки) * Кол-во к отгрузке
                        row_sum = (price * norma) * qty_ship
                        total_mats_sum += row_sum
                        detailed_mats.append({
                            "Товар": row['Товар'], "Расход": round(fact_res, 2), "Упак.": qty_ship, "Сумма": round(row_sum, 0)
                        })
                except Exception as e:
                    continue

        # 2. Услуги
        sum_glass = total_area * p_glass
        sum_toning = (total_area * p_toning) if toning_enabled else 0
        sum_assembly = (total_area * p_assembly) if assembly_enabled else 0
        sum_montage = (total_area * p_montage) if montage_enabled else 0
        
        # 3. ИТОГО (Формула: Затраты * 0.65 + Затраты)
        base_expenses = sum_glass + sum_toning + sum_assembly + sum_montage + total_mats_sum
        margin = base_expenses * 0.65
        grand_total = base_expenses + margin

        # 4. ОТОБРАЖЕНИЕ
        st.markdown("---")
        st.header("📊 Результаты расчета")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Общая площадь", f"{total_area:.3f} м²")
        c2.metric("Общий периметр", f"{total_perim:.1f} м.п.")
        c3.metric("ИТОГО К ОПЛАТЕ", f"{grand_total:,.0f} ₸")

        # Смета услуг
        st.subheader("🛠️ Смета расходов и услуг")
        serv_data = [
            {"Наименование": "Итого материалов (себест.)", "Сумма": f"{total_mats_sum:,.0f} ₸"},
            {"Наименование": "Стеклопакеты", "Сумма": f"{sum_glass:,.0f} ₸"},
            {"Наименование": "Тонировка", "Сумма": f"{sum_toning:,.0f} ₸"},
            {"Наименование": "Сборка изделий", "Сумма": f"{sum_assembly:,.0f} ₸"},
            {"Наименование": "Монтажные работы", "Сумма": f"{sum_montage:,.0f} ₸"},
            {"Наименование": "ОБЕСПЕЧЕНИЕ (наценка 65%)", "Сумма": f"{margin:,.0f} ₸"}
        ]
        st.table(pd.DataFrame(serv_data))

        # Детализация материалов
        with st.expander("🔍 Посмотреть детальный расход материалов"):
            if detailed_mats:
                st.dataframe(pd.DataFrame(detailed_mats).groupby("Товар").sum(), use_container_width=True)
            else:
                st.info("Материалы не найдены.")

        # Кнопка Excel
        excel_out = create_excel_axis({"no": order_no, "type": p_type, "sys": p_sys}, all_positions, None, grand_total, total_area, total_perim)
        st.download_button("📥 Скачать Коммерческое предложение", data=excel_out, file_name=f"Axis_Offer_{order_no}.xlsx")

if __name__ == "__main__":
    main()
