import streamlit as st
import sys
import os
import pandas as pd
from pathlib import Path
import datetime
import tempfile
import math

# --- ФИКСАЦИЯ ПУТЕЙ (Стандарт Axis Pro GF) ---
current_file = Path(__file__).resolve()
root_dir = current_file.parents[1] 
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

# Импорты внутренних модулей
from auth.auth import authenticate
from config.settings import SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH
from references.sheets_reader import load_reference_1, load_reference_2, load_reference_3, load_facade_reference
from calculations.engine_windows import calculate_window_smeta, calculate_impost_length, SYSTEM_MAPPING
from calculations.engine_facade import calculate_facade_materials, calculate_tambour_materials, calculate_tambour_materials_v2
from calculations.mapping import get_code_for_windows_doors, get_code_for_facade
from export.export_kp import export_to_excel, export_facade_to_excel
from history.save_history import save_history

# --- КОНСТАНТЫ ---
PRODUCT_TYPES = ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч."]
FACADE_SYSTEMS = ["Ruit 50F"]

# Настройка страницы
st.set_page_config(page_title="Axis Pro - Калькулятор систем", layout="wide", page_icon="🏗️")

# --- ИСПРАВЛЕННЫЙ БЛОК АВТОРИЗАЦИИ ---
# Передаем параметры, которые ожидает ваша функция authenticate()
if not authenticate(
    st.secrets.get("LOGIN", "admin"), 
    st.secrets.get("PASSWORD", "admin"), 
    GOOGLE_CREDENTIALS_PATH, 
    SPREADSHEET_ID
):
    st.stop()

# Загрузка данных (Кэширование на 10 минут)
@st.cache_data(ttl=600)
def get_references():
    r1 = load_reference_1()
    r2 = load_reference_2()
    r3 = load_reference_3()
    rf = load_facade_reference()
    return r1, r2, r3, rf

ref1, ref2, ref3, ref_facade = get_references()

# ========================================
# ФУНКЦИИ ОТРИСОВКИ (КАК В "APP 2.PY")
# ========================================

def render_history_page():
    st.header("📜 История расчетов")
    # Здесь логика отображения таблицы истории из Google Sheets
    st.info("Раздел истории находится в разработке или загружается из Sheets...")

def render_main_page():
    st.header("🪟 Расчет Окон и Дверей")
    
    if 'positions' not in st.session_state:
        st.session_state.positions = []

    # Боковая панель или форма параметров заказа
    with st.expander("🛠️ Общие параметры заказа", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1: tinting = st.checkbox("Тонировка")
        with c2: assembly = st.checkbox("Сборка")
        with c3: installation = st.selectbox("Монтаж:", ["Нет", "Монтаж", "Демонтаж/Монтаж", "Сложный монтаж"])
        add_details = st.checkbox("Дополнительные детали", value=True)

    # Форма добавления позиции
    with st.form("add_win_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            p_type = st.selectbox("Тип изделия:", PRODUCT_TYPES)
            sys_id = st.selectbox("Система:", list(SYSTEM_MAPPING.keys()))
        with col2:
            width = st.number_input("Ширина (мм):", min_value=100, value=1000)
            height = st.number_input("Высота (мм):", min_value=100, value=1400)
        with col3:
            glass = st.selectbox("Стеклопакет:", list(ref2.keys()) if ref2 else ["Стандарт"])
            qty = st.number_input("Количество (шт):", min_value=1, value=1)
        
        if st.form_submit_button("➕ Добавить в список"):
            code = get_code_for_windows_doors(p_type, sys_id)
            st.session_state.positions.append({
                "product_type": p_type,
                "system": sys_id,
                "code": code,
                "data": {"width": width, "height": height, "qty": qty, "glass_type": glass},
                "sashes": [],
                "imposts": {"has_center": False, "has_tor": False}
            })
            st.rerun()

    # Отображение списка и кнопка расчета
    if st.session_state.positions:
        st.subheader("📝 Список позиций в заказе")
        for i, pos in enumerate(st.session_state.positions):
            st.write(f"{i+1}. {pos['product_type']} ({pos['system']}) - {pos['data']['width']}x{pos['data']['height']} мм — {pos['data']['qty']} шт.")
        
        if st.button("🗑️ Очистить список"):
            st.session_state.positions = []
            st.rerun()

        if st.button("🚀 РАССЧИТАТЬ СМЕТУ"):
            order_data = {
                "common": {
                    "tinting": tinting,
                    "assembly": assembly,
                    "install_type": installation,
                    "add_details": add_details
                },
                "positions": st.session_state.positions
            }
            
            # Вызов движка расчета
            result = calculate_window_smeta(order_data, ref1, ref2, ref3)
            
            # ВЫВОД РЕЗУЛЬТАТОВ (Часть 1, 2, 3)
            st.success("✅ Расчет выполнен успешно")
            
            st.subheader("📊 Итоговые метрики")
            m1, m2, m3 = st.columns(3)
            m1.metric("Общая площадь", f"{result['metrics']['total_area']:.2f} м²")
            m2.metric("Общий периметр", f"{result['metrics']['total_perimeter']:.2f} м")
            m3.metric("ИТОГО К ОПЛАТЕ", f"{result['total_with_margin']:,} ₸")

            with st.expander("📦 Детализация материалов"):
                df_mat = pd.DataFrame(result['part2_materials'])
                st.table(df_mat[['Элемент', 'Артикул', 'Количество', 'Единица', 'Цена', 'Сумма']])

            with st.expander("💰 Финансовый расчет"):
                for label, val in result['part3_final'].items():
                    st.write(f"**{label}:** {val:,} ₸")

# ========================================
# СТРАНИЦА: ФАСАДЫ
# ========================================

def render_facade_page():
    st.header("🏢 Расчет Фасадных систем")
    
    with st.form("facade_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            fw = st.number_input("Ширина фасада W (м):", value=5.0)
            fh1 = st.number_input("Высота H1 (м):", value=6.0)
        with c2:
            fh2 = st.number_input("Высота H2 (м) (для трапеции):", value=0.0)
            f_cols = st.number_input("Количество столбцов:", value=3, min_value=1)
        with c3:
            f_rows = st.number_input("Количество рядов:", value=2, min_value=1)
            f_sys = st.selectbox("Система:", FACADE_SYSTEMS)
        
        if st.form_submit_button("🚀 Рассчитать Фасад"):
            pos = {
                "width": fw, "height_left": fh1, "height_right": fh2,
                "columns": f_cols, "rows": f_rows, "system": f_sys
            }
            res = calculate_facade_materials(pos, ref1, ref2, ref3)
            
            st.subheader(f"📊 Результат: {res['area']:.2f} м²")
            
            col_a, col_b = st.columns(2)
            with col_a:
                st.write("**Каркас (Стойки/Ригели):**")
                st.dataframe(pd.DataFrame(res['skeleton_items']))
            with col_b:
                st.write("**Заполнение (Вставки):**")
                st.dataframe(pd.DataFrame(res['insert_items']))

# ========================================
# СТРАНИЦА: ТАМБУР
# ========================================

def render_tambour_page():
    st.header("🚪 Оконный тамбур")
    st.write("Используйте этот раздел для расчета входных групп (тамбуров).")
    # Здесь логика для calculate_tambour_materials

# ========================================
# ГЛАВНОЕ МЕНЮ
# ========================================

with st.sidebar:
    st.title("📍 AXIS PRO")
    if 'menu_selection' not in st.session_state:
        st.session_state.menu_selection = "Главная (Окна/Двери)"
    
    choice = st.radio("Разделы:", ["Главная (Окна/Двери)", "Фасады", "Оконный тамбур", "История"])
    st.session_state.menu_selection = choice

# Роутинг страниц
if st.session_state.menu_selection == "Главная (Окна/Двери)":
    render_main_page()
elif st.session_state.menu_selection == "Фасады":
    render_facade_page()
elif st.session_state.menu_selection == "Оконный тамбур":
    render_tambour_page()
elif st.session_state.menu_selection == "История":
    render_history_page()
