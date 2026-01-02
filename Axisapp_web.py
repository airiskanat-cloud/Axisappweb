import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
from datetime import datetime

# Настройки листов (строго по твоей ссылке)
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
    st.set_page_config(page_title="Axis Pro GF", layout="wide")
    st.title("🏗️ Axis Pro GF (Engine v15)")

    try:
        client = get_client()
        sh = client.open_by_key(GSPREAD_SHEET_ID)
    except Exception as e:
        st.error(f"Ошибка подключения к Google Sheets: {e}")
        return

    # --- ИНТЕРФЕЙС ---
    with st.sidebar:
        st.header("Ввод данных")
        order_no = st.text_input("Номер заказа", "001")
        p_type = st.selectbox("Тип изделия", ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"])
        p_sys = st.selectbox("Система", ["ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "Ruit 50F"])
        
    col1, col2, col3, col4 = st.columns(4)
    W = col1.number_input("Ширина (мм)", value=1000)
    H = col2.number_input("Высота (мм)", value=1500)
    qty = col3.number_input("Кол-во (шт)", value=1)
    n_imp = col4.number_input("Деления", value=0)

    if st.button("🚀 ВЫПОЛНИТЬ ПОЛНЫЙ РАСЧЕТ"):
        with st.spinner('Синхронизация с Google Sheets...'):
            # 1. Записываем в лист ЗАПРОСЫ (как в 15 версии)
            worksheet_form = sh.worksheet(SHEET_FORM)
            timestamp = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            
            # Очищаем старые данные или добавляем новые (в v15 обычно очищали для нового расчета)
            # Для истории лучше добавлять. Давай добавим строку:
            new_row = [order_no, p_type, p_sys, W, H, qty, n_imp, timestamp]
            worksheet_form.append_row(new_row)
            
            # 2. Ждем обновления формул в облаке
            time.sleep(2) 
            
            # 3. Читаем результаты из листа "Расчетом расходов материалов"
            res_mats = pd.DataFrame(sh.worksheet(SHEET_MATERIAL).get_all_records())
            res_mats.columns = res_mats.columns.str.strip()
            
            # 4. Читаем итог из "Итоговый расчет с монтажом"
            res_final = pd.DataFrame(sh.worksheet(SHEET_FINAL).get_all_records())
            res_final.columns = res_final.columns.str.strip()

            # --- ВЫВОД РЕЗУЛЬТАТОВ ---
            st.success("Данные в таблице обновлены!")
            
            # Показываем последние данные из расчета материалов
            if not res_mats.empty:
                st.subheader("📋 Спецификация материалов (из Справочника)")
                # Фильтруем только те строки, где есть количество (чтобы не спамить пустыми)
                display_mats = res_mats[res_mats['Количество'].astype(float) > 0] if 'Количество' in res_mats.columns else res_mats
                st.table(display_mats.tail(10)) # Показываем последние 10 записей

            # Показываем финальную сумму
            if not res_final.empty:
                st.markdown("---")
                last_row = res_final.iloc[-1]
                m1, m2 = st.columns(2)
                # Предполагаем названия колонок из твоего файла
                m1.metric("Общая площадь", f"{last_row.get('Площадь', 0)} м2")
                m2.metric("ИТОГО К ОПЛАТЕ", f"{last_row.get('Сумма', 0):,.0f} ₸")

            st.info("История сохранена в Google Sheets на листах ЗАПРОСЫ и Итоговый расчет.")

if __name__ == "__main__":
    main()
