import math
import os
import sys
import shutil
from io import BytesIO
import zipfile
import logging
import json
import ast
import operator as op
from datetime import datetime

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# =========================================================
# 1. СИСТЕМНЫЕ НАСТРОЙКИ И ЛОГГИРОВАНИЕ (ИЗ v15)
# =========================================================
DEBUG = False
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

GSPREAD_SHEET_ID = "13kxXxhYNkMBhnltEZT6v2cdRu6aTF4_7wm7glqq45O8"

# Листы Google Таблиц
SHEET_REF1 = "СПРАВОЧНИК -1"
SHEET_REF2 = "СПРАВОЧНИК -2"
SHEET_REF3 = "СПРАВОЧНИК -3"
SHEET_FORM = "ЗАПРОСЫ"
SHEET_GABARITS = "Расчет по габаритам"
SHEET_MATERIAL = "Расчетом расходов материалов"
SHEET_FINAL = "Итоговый расчет с монтажом"
SHEET_USERS = "ПОЛЬЗОВАТЕЛИ"

# =========================================================
# 2. ПОДКЛЮЧЕНИЕ К ДАННЫМ (БЕЗ st.secrets)
# =========================================================
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        # Используем секретный файл Render напрямую
        creds = Credentials.from_service_account_file("/etc/secrets/gcp.json", scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"❌ Критическая ошибка доступа к gcp.json: {e}")
        st.stop()

@st.cache_data(ttl=600)
def load_all_data():
    try:
        client = get_gspread_client()
        sh = client.open_by_key(GSPREAD_SHEET_ID)
        
        data = {
            "ref1": pd.DataFrame(sh.worksheet(SHEET_REF1).get_all_records()),
            "ref2": pd.DataFrame(sh.worksheet(SHEET_REF2).get_all_records()),
            "ref3": pd.DataFrame(sh.worksheet(SHEET_REF3).get_all_records()),
            "users": pd.DataFrame(sh.worksheet(SHEET_USERS).get_all_records()),
            "sh": sh
        }
        return data
    except Exception as e:
        st.error(f"❌ Ошибка загрузки базы данных: {e}")
        return None

# =========================================================
# 3. МАТЕМАТИЧЕСКОЕ ЯДРО (ВЫЧЕТЫ И ФОРМУЛЫ)
# =========================================================
def evaluate_formula(formula_str, context):
    """Безопасное выполнение формул из Справочника-3"""
    try:
        # Заменяем Excel-символы на Python-совместимые
        expr = str(formula_str).replace('=', '').replace('^', '**')
        # Разрешаем только базовую математику
        allowed_names = {"math": math, "W": context.get('W', 0), "H": context.get('H', 0), 
                         "qty": context.get('qty', 0), "n_m": context.get('n_m', 0), 
                         "n_t": context.get('n_t', 0), "hinges": context.get('hinges', 0)}
        return eval(expr, {"__builtins__": None}, allowed_names)
    except Exception as e:
        logger.error(f"Ошибка в формуле {formula_str}: {e}")
        return 0

# =========================================================
# 4. ГЕНЕРАЦИЯ КП В EXCEL (ПОЛНАЯ ВЕРСИЯ ШЕВЧЕНКО)
# =========================================================
def build_pro_excel(order_meta, positions, total_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Коммерческое предложение"

    # Стили
    bold_font = Font(bold=True, size=11)
    center_align = Alignment(horizontal='center', vertical='center')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # Шапка компании
    ws.merge_cells('C1:E1')
    ws['C1'] = "ООО «AXIS»"
    ws['C1'].font = Font(bold=True, size=14)
    ws['C2'] = "Город Астана, Тел.: +7 707 504 4040"
    
    ws.append([])
    ws.append(["КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", "", f"Заказ №: {order_meta['no']}"])
    ws.append(["Дата:", datetime.now().strftime("%d.%m.%Y")])
    ws.append(["Система:", order_meta['sys']])
    ws.append(["Цвет RAL:", order_meta['ral']])
    ws.append([])

    # Таблица позиций
    headers = ["№", "Наименование", "Размеры (мм)", "Кол-во", "Площадь (м2)", "Заполнение"]
    ws.append(headers)
    
    for i, p in enumerate(positions, 1):
        ws.append([i, p['type'], f"{p['W']} x {p['H']}", p['qty'], p['area'], p['fill']])

    ws.append([])
    ws.append(["ИТОГОВЫЕ ПОКАЗАТЕЛИ"])
    ws.append(["Общая площадь:", f"{total_data['area']:.3f} м2"])
    ws.append(["Общий периметр:", f"{total_data['perim']:.2f} м.п."])
    ws.append(["СУММА К ОПЛАТЕ:", f"{total_data['sum']:,.0f} тенге"])

    # Настройка ширины колонок
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 15

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# =========================================================
# 5. ОСНОВНОЙ ИНТЕРФЕЙС STREAMLIT
# =========================================================
def main():
    st.set_page_config(page_title="Axisapp Pro v16", layout="wide")
    
    data = load_all_data()
    if not data: return
    
    if 'auth' not in st.session_state: st.session_state.auth = False
    if 'cart' not in st.session_state: st.session_state.cart = []

    # --- АВТОРИЗАЦИЯ ---
    if not st.session_state.auth:
        st.title("🔑 Axisapp: Система инженерных расчетов")
        col_a, _ = st.columns([1, 2])
        with col_a:
            u = st.text_input("Логин")
            p = st.text_input("Пароль", type="password")
            if st.button("Войти"):
                user_match = data['users'][(data['users']['Логин'] == u) & (data['users']['Пароль'].astype(str) == p)]
                if not user_match.empty:
                    st.session_state.auth = True
                    st.session_state.user_role = user_match.iloc[0]['Роль']
                    st.rerun()
                else:
                    st.error("Доступ отклонен")
        return

    # --- ПАНЕЛЬ УПРАВЛЕНИЯ ---
    st.sidebar.image("https://static.tildacdn.com/tild3133-3131-4131-b331-313131313131/logo_axis.png", width=150) # Пример лого
    st.sidebar.title(f"👤 {st.session_state.user_role}")
    
    order_number = st.sidebar.text_input("Заказ №", "2025-001")
    
    tabs = st.tabs(["🏗️ Конструктор изделий", "🛒 Корзина заказа", "📊 Расчет и Смета"])

    # --- ВКЛАДКА 1: КОНСТРУКТОР ---
    with tabs[0]:
        st.subheader("Добавление новой позиции")
        c1, c2, c3 = st.columns([2, 2, 1])
        p_type = c1.selectbox("Тип изделия", ["Окно глух.", "Окно с откр.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад (Каркас)"])
        p_system = c2.selectbox("Система", ["Ruit 50F", "ALG 2030-45C", "ALG 2030-63C", "ALG 2030-73C", "ALG Slim"])
        p_qty = c3.number_input("Кол-во шт", min_value=1, value=1)

        col_w, col_h = st.columns(2)
        W = col_w.number_input("Ширина W (мм)", min_value=100, value=1000)
        H = col_h.number_input("Высота H (мм)", min_value=100, value=1000)

        # Инженерные дополнения (из чертежей)
        st.markdown("---")
        exp_col1, exp_col2 = st.columns(2)
        
        with exp_col1:
            is_facade_insert = st.checkbox("Вставка в фасадный каркас (адаптер)")
            filling = st.radio("Тип заполнения", ["Стеклопакет", "Ламбри (Панель)", "Сэндвич"])
            
        with exp_col2:
            if "Дверь" in p_type:
                handle_type = st.selectbox("Тип ручки", ["Нажимной гарнитур", "Офисная ручка (скоба)", "Офисная 1000мм"])
                has_closer = st.checkbox("Доводчик")
            else:
                handle_type, has_closer = "Стандарт", False

        if p_type == "Фасад (Каркас)":
            f_col1, f_col2 = st.columns(2)
            n_m = f_col1.number_input("Кол-во стоек (мульонов)", value=2)
            n_t = f_col2.number_input("Кол-во ригелей", value=1)
        else:
            n_m, n_t = 0, 0

        if st.button("🚀 Добавить изделие в расчет"):
            # Авто-петли по логике высоты
            hinges = 3 if H > 2100 and "Дверь" in p_type else 2
            
            new_pos = {
                "type": p_type, "sys": p_system, "W": W, "H": H, "qty": p_qty,
                "area": (W * H / 1000000) * p_qty,
                "perim": ((W + H) * 2 / 1000) * p_qty,
                "fill": filling, "hinges": hinges, "is_insert": is_facade_insert,
                "handle": handle_type, "closer": has_closer, "n_m": n_m, "n_t": n_t
            }
            st.session_state.cart.append(new_pos)
            st.toast("Позиция добавлена успешно!")

    # --- ВКЛАДКА 2: КОРЗИНА ---
    with tabs[1]:
        if st.session_state.cart:
            st.write("### Состав текущего заказа")
            cart_df = pd.DataFrame(st.session_state.cart)
            st.dataframe(cart_df[['type', 'sys', 'W', 'H', 'qty', 'area', 'fill']])
            
            if st.button("🗑️ Полностью очистить заказ"):
                st.session_state.cart = []
                st.rerun()
        else:
            st.info("Ваша корзина пуста. Перейдите в Конструктор.")

    # --- ВКЛАДКА 3: РАСЧЕТ МАТЕРИАЛОВ И ЭКОНОМИКА ---
    with tabs[2]:
        if st.session_state.cart:
            st.sidebar.subheader("Дополнительные опции")
            toning = st.sidebar.checkbox("Тонировка стекла")
            assembly = st.sidebar.checkbox("Сборка на производстве", value=True)
            montage = st.sidebar.checkbox("Монтажные работы")
            ral_color = st.sidebar.text_input("Цвет RAL", "7024")

            # 1. Сбор данных из Справочника-3 по всем позициям
            full_mats_list = []
            total_mats_cost = 0

            for item in st.session_state.cart:
                # Фильтруем материалы для конкретного типа изделия
                ref3_filtered = data['ref3'][data['ref3']['Тип изделия'] == item['type']]
                
                context = {
                    "W": item['W'], "H": item['H'], "qty": item['qty'], 
                    "n_m": item['n_m'], "n_t": item['n_t'], 
                    "hinges": item['hinges'], "is_insert": int(item['is_insert'])
                }

                for _, row in ref3_filtered.iterrows():
                    qty_mat = evaluate_formula(row['Формула_Python'], context)
                    if qty_mat > 0:
                        # Получаем цену из Справочника-2
                        price_row = data['ref2'][data['ref2']['Система'] == item['sys']]
                        price_unit = price_row['Цена'].values[0] if not price_row.empty else 0
                        
                        total_mats_cost += (qty_mat * price_unit)
                        full_mats_list.append({
                            "Артикул": row.get('Артикул', '-'),
                            "Наименование": row['Наименование'],
                            "Расход": f"{qty_mat:.2f}",
                            "Ед.": row['Ед'],
                            "Сумма": f"{(qty_mat * price_unit):,.0f}"
                        })

            # 2. ИТОГОВАЯ ЭКОНОМИКА (v15 Формула)
            total_area = sum(i['area'] for i in st.session_state.cart)
            glass_base = total_area * 17500 # Базовая цена за м2
            if toning: glass_base += (total_area * 4000)
            
            work_cost = 0
            if assembly: work_cost += (total_area * 4500)
            if montage: work_cost += (total_area * 6500)

            # (Мат + Стекло + Работа) * 1.65
            subtotal = total_mats_cost + glass_base + work_cost
            final_sum = subtotal * 1.65

            # ВЫВОД МЕТРИК
            st.subheader("Итоговые результаты проекта")
            m1, m2, m3 = st.columns(3)
            m1.metric("Общая площадь", f"{total_area:.3f} м2")
            m2.metric("Чистый металл + фурн.", f"{total_mats_cost:,.0f} ₸")
            m3.metric("СУММА К ОПЛАТЕ (с обесп. 65%)", f"{final_sum:,.0f} ₸")

            with st.expander("Детальная спецификация материалов"):
                st.table(pd.DataFrame(full_mats_list))

            # ЭКСПОРТ
            meta = {"no": order_number, "sys": st.session_state.cart[0]['sys'], "ral": ral_color}
            total_data = {"area": total_area, "perim": sum(i['perim'] for i in st.session_state.cart), "sum": final_sum}
            
            excel_bytes = build_pro_excel(meta, st.session_state.cart, total_data)
            st.download_button("📥 Скачать Коммерческое Предложение (Excel)", 
                               data=excel_bytes, 
                               file_name=f"Axis_KP_{order_number}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # ЗАПИСЬ В GOOGLE ТАБЛИЦУ (v15 функция)
            if st.button("💾 Сохранить расчет в облако"):
                try:
                    sheet_final = data['sh'].worksheet(SHEET_FINAL)
                    sheet_final.append_row([order_number, total_area, final_sum, datetime.now().strftime("%Y-%m-%d %H:%M")])
                    st.success("Данные успешно сохранены в Google Sheets")
                except:
                    st.error("Ошибка записи в облако")

    if st.sidebar.button("🚪 Выход из системы"):
        st.session_state.auth = False
        st.rerun()

if __name__ == "__main__":
    main()
