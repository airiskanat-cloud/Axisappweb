import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
from datetime import datetime

# =========================
# КОНСТАНТЫ
# =========================
GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"

def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
    return gspread.authorize(creds)

def main():
    st.set_page_config(page_title="Axis Pro GF", layout="wide", page_icon="🏗️")
    
    # Стилизация
    st.markdown("""
        <style>
        .stButton>button { background-color: #1e3d59; color: white; border-radius: 5px; height: 3em; width: 100%; }
        .block-container { padding-top: 2rem; }
        .stMetric { background-color: #ffffff; border: 1px solid #e6e9ef; padding: 15px; border-radius: 10px; }
        </style>
    """, unsafe_allow_html=True)

    st.title("🏗️ Axis Pro GF | Профессиональный расчет")

    try:
        client = get_client()
        sh = client.open_by_key(GSPREAD_SHEET_ID)
    except Exception as e:
        st.error(f"Ошибка доступа к Google Sheets: {e}")
        return

    # --- ФОРМА ЗАПОЛНЕНИЯ (СОГЛАСНО ТВОЕМУ СПИСКУ) ---
    
    with st.form("main_form"):
        st.subheader("1. Основные данные и Профиль")
        c1, c2, c3, c4 = st.columns(4)
        order_no = c1.text_input("Номер заказа", "001")
        pos_no = c2.text_input("№ позиции", "1")
        p_type = c3.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        v_type = c4.selectbox("Вид изделия", ["Стандарт", "Витраж", "Входная группа"])
        
        c5, c6, c7, c8 = st.columns(4)
        s_count = c5.number_input("Створки (кол-во)", value=0)
        p_sys = c6.selectbox("Профильная система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F", "ALG Slim"])
        glass_thick = c7.selectbox("Толщина стеклопакета", ["24 мм", "32 мм", "40 мм"])
        glass_type = c8.selectbox("Тип стеклопакета", ["Однокамерный", "Двухкамерный", "Энергосберегающий"])

        st.subheader("2. Габариты и Деления (мм)")
        g1, g2, g3, g4, g5 = st.columns(5)
        W = g1.number_input("Ширина, мм", value=1000)
        H = g2.number_input("Высота, мм", value=1500)
        L = g3.number_input("LEFT", value=0)
        C = g4.number_input("CENTER", value=0)
        R = g5.number_input("RIGHT", value=0)

        g6, g7, g8, g9, g10 = st.columns(5)
        T = g6.number_input("TOP", value=0)
        sW = g7.number_input("Ширина створки", value=0)
        sH = g8.number_input("Высота створки", value=0)
        qty = g9.number_input("Кол-во шт", value=1)
        nwin = g10.number_input("Nwin", value=1)

        st.subheader("3. Заполнение и Услуги")
        s1, s2, s3, s4, s5, s6 = st.columns(6)
        filling = s1.selectbox("Заполнение", ["Стекло", "Сэндвич", "Ламбри"])
        cutting = s2.selectbox("Нарезка", ["Заводская", "Цех"])
        toning = s3.checkbox("Тонировка")
        assembly = s4.checkbox("Сборка", value=True)
        montage = s5.checkbox("Монтаж")

        submit = st.form_submit_button("🚀 ЗАПУСТИТЬ СИНХРОННЫЙ РАСЧЕТ")

    if submit:
        with st.spinner('Синхронизация данных с таблицей...'):
            # Сбор данных в одну строку (строго по твоему списку из запроса)
            row_to_send = [
                order_no, pos_no, p_type, v_type, s_count, p_sys, 
                glass_thick, glass_type, filling, W, H, L, C, R, T, 
                sW, sH, qty, nwin, cutting, 
                "Да" if toning else "Нет", 
                "Да" if assembly else "Нет", 
                "Да" if montage else "Нет",
                datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            ]

            # 1. Запись в "ЗАПРОСЫ"
            ws_form = sh.worksheet(SHEET_FORM)
            ws_form.append_row(row_to_send)
            
            # 2. Пауза для облачного пересчета
            time.sleep(3) 
            
            # 3. Получение итогов
            try:
                # Читаем результаты материалов
                df_mats = pd.DataFrame(sh.worksheet(SHEET_MATERIAL).get_all_records())
                df_final = pd.DataFrame(sh.worksheet(SHEET_FINAL).get_all_records())
                
                st.success(f"Заказ {order_no} успешно рассчитан!")
                
                # Показываем итоги
                if not df_final.empty:
                    last_res = df_final.iloc[-1]
                    r1, r2, r3 = st.columns(3)
                    r1.metric("Площадь", f"{last_res.get('Площадь', 0)} м2")
                    r2.metric("Мат. расход", f"{last_res.get('Сумма материалов', 0):,.0f} ₸")
                    r3.metric("ИТОГО", f"{last_res.get('Итоговая сумма', 0):,.0f} ₸")

                with st.expander("🔍 Посмотреть детализацию материалов"):
                    st.dataframe(df_mats.tail(20))
            except Exception as e:
                st.warning(f"Данные записаны, но не удалось считать результат: {e}")

if __name__ == "__main__":
    main()
