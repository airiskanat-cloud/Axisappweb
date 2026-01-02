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
# 1. СИСТЕМНЫЕ НАСТРОЙКИ И ЛОГГИРОВАНИЕ
# =========================================================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Имена листов в Google Таблице
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================================================
# 2. ПОДКЛЮЧЕНИЕ (RENDER SAFE)
# =========================================================
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        # Используем секретный файл gcp.json, который вы загрузили в Render
        creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"❌ Ошибка доступа к ключу /etc/secrets/gcp.json: {e}")
        st.stop()

@st.cache_data(ttl=600)
def load_all_data():
    try:
        client = get_gspread_client()
        sh = client.open_by_key(GSPREAD_SHEET_ID)
        return {
            "ref1": pd.DataFrame(sh.worksheet(SHEET_REF1).get_all_records()),
            "ref2": pd.DataFrame(sh.worksheet(SHEET_REF2).get_all_records()),
            "ref3": pd.DataFrame(sh.worksheet(SHEET_REF3).get_all_records()),
            "users": pd.DataFrame(sh.worksheet(SHEET_USERS).get_all_records()),
            "sh": sh
        }
    except Exception as e:
        st.error(f"❌ Ошибка загрузки данных из Google Sheets: {e}")
        return None

# =========================================================
# 3. ИНЖЕНЕРНОЕ ЯДРО: РАСЧЕТ МАТЕРИАЛОВ
# =========================================================
def calculate_materials_for_pos(pos, ref3_data, ref2_prices):
    """
    Выполняет расчет всех комплектующих для одной позиции 
    на основе формул из Справочника-3.
    """
    spec = []
    total_cost = 0
    
    # Фильтруем справочник материалов под конкретный тип изделия
    mats_for_type = ref3_data[ref3_data['Тип изделия'] == pos['type']]
    
    # Контекст для формул Python (W, H, кол-во и доп. параметры)
    context = {
        "W": pos['W'], "H": pos['H'], "qty": pos['qty'],
        "n_m": pos.get('n_m', 0), "n_t": pos.get('n_t', 0),
        "hinges": pos.get('hinges', 2), "is_insert": int(pos.get('is_insert', False)),
        "math": math
    }

    for _, row in mats_for_type.iterrows():
        formula = str(row['Формула_Python']).replace('=', '').replace('^', '**')
        try:
            count = eval(formula, {"__builtins__": None}, context)
            if count > 0:
                # Ищем цену в Справочнике-2 для выбранной системы
                price_row = ref2_prices[ref2_prices['Система'] == pos['sys']]
                # Если в справочнике 2 есть привязка к артикулу — берем её, иначе — базовую цену системы
                price_unit = price_row['Цена'].values[0] if not price_row.empty else 0
                
                sum_mat = count * price_unit
                total_cost += sum_mat
                spec.append({
                    "Наименование": row['Наименование'],
                    "Артикул": row.get('Артикул', '-'),
                    "Количество": round(count, 2),
                    "Ед": row['Ед'],
                    "Сумма": round(sum_mat, 0)
                })
        except Exception as e:
            logger.error(f"Ошибка в формуле для {row['Наименование']}: {e}")
            
    return spec, total_cost

# =========================================================
# 4. ГЕНЕРАЦИЯ КП (EXCEL ПО ОБРАЗЦУ ШЕВЧЕНКО)
# =========================================================
def build_excel_offer(order_meta, positions, grand_total, total_area):
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"

    # Оформление
    blue_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    bold_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Шапка Axis
    ws['C1'] = "ООО «AXIS»"
    ws['C1'].font = Font(bold=True, size=14)
    ws['C2'] = "Город Астана. Тел: +7 707 504 4040"
    ws.append([])

    ws.append(["ЗАКАЗ №", order_meta['no']])
    ws.append(["Цвет RAL:", order_meta['ral']])
    ws.append([])

    # Заголовки таблицы
    headers = ["№", "Тип изделия", "Размеры (ШхВ)", "Система", "Кол-во", "Площадь (м2)"]
    ws.append(headers)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=ws.max_row, column=col_num)
        cell.font = bold_font
        cell.fill = blue_fill
        cell.border = border

    # Данные позиций
    for i, p in enumerate(positions, 1):
        row = [i, p['type'], f"{p['W']}x{p['H']}", p['sys'], p['qty'], round(p['area'], 3)]
        ws.append(row)
        for col_num in range(1, 7):
            ws.cell(row=ws.max_row, column=col_num).border = border

    ws.append([])
    ws.append(["ИТОГО ПЛОЩАДЬ:", f"{total_area:.3f} м2"])
    ws.append(["ИТОГО К ОПЛАТЕ:", f"{grand_total:,.0f} тенге"])
    ws.cell(row=ws.max_row, column=1).font = bold_font

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# =========================================================
# 5. ИНТЕРФЕЙС STREAMLIT
# =========================================================
def main():
    st.set_page_config(page_title="Axisapp Pro v16", layout="wide")
    
    # Загрузка данных
    db = load_all_data()
    if not db: return

    if 'auth' not in st.session_state: st.session_state.auth = False
    if 'cart' not in st.session_state: st.session_state.cart = []

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("🧱 Вход в инженерную систему AXIS")
        with st.container():
            u = st.text_input("Логин")
            p = st.text_input("Пароль", type="password")
            if st.button("Войти"):
                user = db['users'][(db['users']['Логин'] == u) & (db['users']['Пароль'].astype(str) == p)]
                if not user.empty:
                    st.session_state.auth = True
                    st.session_state.user_role = user.iloc[0]['Роль']
                    st.rerun()
                else:
                    st.error("Ошибка входа")
        return

    # --- ПАНЕЛЬ УПРАВЛЕНИЯ ---
    st.sidebar.title(f"👤 {st.session_state.user_role}")
    order_id = st.sidebar.text_input("Номер заказа", "2025-AX-001")
    ral_color = st.sidebar.text_input("Цвет конструкции (RAL)", "7024")
    
    tabs = st.tabs(["📐 Конструктор", "📋 Состав заказа", "📊 Итоги и Выгрузка"])

    # ВКЛАДКА 1: ДОБАВЛЕНИЕ ПОЗИЦИЙ
    with tabs[0]:
        st.subheader("Настройка параметров изделия")
        col1, col2, col3 = st.columns([2, 2, 1])
        
        type_choice = col1.selectbox("Тип изделия", [
            "Окно глух.", "Окно с откр.", 
            "Дверь 1 створч.", "Дверь 2-х створч.", 
            "Фасад"
        ])
        sys_choice = col2.selectbox("Система профиля", [
            "ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", 
            "ALG 2030-73C", "ALG 2030-Slim", "Ruit 50F"
        ])
        qty_input = col3.number_input("Количество (шт)", min_value=1, value=1)

        cW, cH = st.columns(2)
        W_val = cW.number_input("Ширина W (мм)", min_value=100, value=1000)
        H_val = cH.number_input("Высота H (мм)", min_value=100, value=1500)

        # Специфика фасада
        n_m, n_t, is_ins = 0, 0, False
        if type_choice == "Фасад":
            f1, f2 = st.columns(2)
            n_m = f1.number_input("Количество стоек", value=2)
            n_t = f2.number_input("Количество ригелей", value=1)
        else:
            is_ins = st.checkbox("Вставка в фасадный каркас", help="Добавляет адаптерный профиль 5081")

        if st.button("🏗️ Добавить позицию в проект"):
            # Расчет петель по чертежам (H > 2100 = 3 петли)
            hng = 3 if H_val > 2100 and "Дверь" in type_choice else 2
            
            pos_entry = {
                "type": type_choice, "sys": sys_choice, "qty": qty_input,
                "W": W_val, "H": H_val, "area": (W_val * H_val / 1000000) * qty_input,
                "perim": ((W_val + H_val) * 2 / 1000) * qty_input,
                "n_m": n_m, "n_t": n_t, "is_insert": is_ins, "hinges": hng
            }
            st.session_state.cart.append(pos_entry)
            st.toast(f"Добавлено: {type_choice}")

    # ВКЛАДКА 2: КОРЗИНА (РЕВИЗИЯ)
    with tabs[1]:
        if st.session_state.cart:
            st.write("### Объекты в расчете")
            df_cart = pd.DataFrame(st.session_state.cart)
            st.table(df_cart[['type', 'sys', 'W', 'H', 'qty', 'area']])
            if st.button("🗑️ Очистить всё"):
                st.session_state.cart = []
                st.rerun()
        else:
            st.info("Ваш проект пока пуст.")

    # ВКЛАДКА 3: РАСЧЕТ И ИТОГИ
    with tabs[2]:
        if st.session_state.cart:
            st.sidebar.markdown("---")
            toning = st.sidebar.checkbox("Тонировка стекла")
            assembly = st.sidebar.checkbox("Сборка", value=True)
            montage = st.sidebar.checkbox("Монтаж")

            # ГЛАВНЫЙ ЦИКЛ РАСЧЕТА
            full_spec = []
            mats_cost_total = 0
            
            for p in st.session_state.cart:
                item_spec, item_cost = calculate_materials_for_pos(p, db['ref3'], db['ref2'])
                full_spec.extend(item_spec)
                mats_cost_total += item_cost

            # Экономика (v15 + Ваша формула)
            total_area_all = sum(i['area'] for i in st.session_state.cart)
            glass_cost = total_area_all * 18500 # База
            if toning: glass_cost += (total_area_all * 4000)
            
            labor_cost = 0
            if assembly: labor_cost += (total_area_all * 5000)
            if montage: labor_cost += (total_area_all * 8000)

            # (Мат + Стекло + Работы) * 1.65
            subtotal = mats_cost_total + glass_cost + labor_cost
            grand_total = subtotal * 1.65

            # ВИТРИНА
            st.subheader("Результаты расчета проекта")
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Общая площадь", f"{total_area_all:.3f} м²")
            m2.metric("Себестоимость мат.", f"{mats_cost_total:,.0f} ₸")
            m3.metric("ИТОГО (с обеспечением 1.65)", f"{grand_total:,.0f} ₸")

            with st.expander("Детальная ведомость материалов (Спецификация)"):
                st.table(pd.DataFrame(full_spec).groupby("Наименование").sum())

            # ЭКСПОРТ И СОХРАНЕНИЕ
            meta = {"no": order_id, "sys": st.session_state.cart[0]['sys'], "ral": ral_color}
            excel_data = build_excel_offer(meta, st.session_state.cart, grand_total, total_area_all)
            
            st.download_button(
                "💾 Скачать КП в Excel (формат Шевченко)", 
                data=excel_data, 
                file_name=f"Offer_{order_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if st.button("☁️ Сохранить в Google Таблицу"):
                try:
                    db['sh'].worksheet(SHEET_FINAL).append_row([
                        order_id, ral_color, total_area_all, grand_total, datetime.now().strftime("%d.%m.%Y %H:%M")
                    ])
                    st.success("Данные успешно сохранены в лист 'Итоговый расчет'!")
                except Exception as e:
                    st.error(f"Ошибка сохранения: {e}")

    if st.sidebar.button("🚪 Выйти"):
        st.session_state.auth = False
        st.rerun()

if __name__ == "__main__":
    main()
