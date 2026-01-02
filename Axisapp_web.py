import math
import os
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# Названия листов из твоей базы
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=600)
def load_all_data():
    client = get_gspread_client()
    sh = client.open_by_key(GSPREAD_SHEET_ID)
    
    def get_clean_df(name):
        df = pd.DataFrame(sh.worksheet(name).get_all_records())
        df.columns = df.columns.str.strip() # Чистим заголовки
        return df

    return {
        "ref1": get_clean_df(SHEET_REF1),
        "ref2": get_clean_df(SHEET_REF2),
        "ref3": get_clean_df(SHEET_REF3),
        "sh": sh
    }

def main():
    st.set_page_config(page_title="Axis Pro GF", layout="wide")
    
    # Стилизация под алюминиевый бизнес
    st.markdown("""<style> .stMetric { background-color: #f0f2f6; padding: 20px; border-radius: 10px; } </style>""", unsafe_allow_html=True)
    
    db = load_all_data()

    st.title("🏗️ Axis Pro GF")
    
    # Панель управления
    with st.sidebar:
        st.header("Настройки заказа")
        order_no = st.text_input("Номер заказа", "001")
        # ВАЖНО: Эти названия должны быть в Справочнике-3!
        p_type = st.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_sys = st.selectbox("Система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])
        
        st.markdown("---")
        toning = st.checkbox("Тонировка")
        assembly = st.checkbox("Сборка", value=True)
        montage = st.checkbox("Монтаж")

    # Ввод данных
    col1, col2, col3, col4 = st.columns(4)
    W = col1.number_input("Ширина W (мм)", value=1000)
    H = col2.number_input("Высота H (мм)", value=1500)
    qty = col3.number_input("Кол-во (шт)", value=1)
    n_imp = col4.number_input("Деления/Импосты", value=0)

    if st.button("🚀 РАССЧИТАТЬ МАТЕРИАЛЫ И ИТОГИ"):
        # 1. Фильтруем справочник материалов
        # Очищаем данные в таблице для сравнения
        ref3 = db['ref3'].copy()
        ref3['Тип изделия'] = ref3['Тип изделия'].astype(str).str.strip()
        
        mats_to_calc = ref3[ref3['Тип изделия'] == p_type]
        
        if mats_to_calc.empty:
            st.error(f"❌ В Справочнике-3 нет материалов для типа '{p_type}'. Проверь написание в таблице!")
            return

        spec = []
        mats_cost = 0
        
        # Контекст для формул
        ctx = {"W": W, "H": H, "qty": qty, "n_imp": n_imp, "math": math}

        for _, row in mats_to_calc.iterrows():
            try:
                formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
                count = eval(formula, {"__builtins__": None}, ctx)
                
                if count > 0:
                    # Ищем цену в Справочнике-2
                    prices = db['ref2'][db['ref2']['Система'].str.strip() == p_sys.strip()]
                    price = prices['Цена'].values[0] if not prices.empty else 0
                    
                    sum_row = count * price
                    mats_cost += sum_row
                    spec.append({
                        "Наименование": row['Наименование'],
                        "Расход": round(count, 2),
                        "Ед.": row['Ед'],
                        "Сумма": round(sum_row, 0)
                    })
            except Exception as e:
                logger.error(f"Ошибка в формуле: {e}")

        # 2. Итоговый расчет (Формула 1.65)
        area = (W * H / 1000000) * qty
        glass = area * 18000 + (area * 4000 if toning else 0)
        works = area * (5000 if assembly else 0) + area * (8000 if montage else 0)
        
        final_total = (mats_cost + glass + works) * 1.65

        # 3. Вывод результатов
        if not spec:
            st.warning("⚠️ Материалы не найдены. Проверь формулы в Справочнике-3.")
        else:
            st.markdown("### 📊 Итоговые показатели")
            m1, m2, m3 = st.columns(3)
            m1.metric("Площадь", f"{area:.3f} м²")
            m2.metric("Мат. себест.", f"{mats_cost:,.0f} ₸")
            m3.metric("ИТОГО К ОПЛАТЕ", f"{final_total:,.0f} ₸")

            st.markdown("### 📋 Спецификация материалов")
            st.table(pd.DataFrame(spec))

if __name__ == "__main__":
    main()
