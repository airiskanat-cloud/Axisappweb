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
from calculations.material_basket_V9 import MaterialAggregator
from calculations.mapping import get_code_for_windows_doors, get_code_for_facade
from export.export_kp import export_to_excel, export_facade_to_excel
from history.save_history import save_history

# --- КОНСТАНТЫ ИЗ ТЗ ---
PRODUCT_TYPES = ["Окно с откр.", "Окно глух.", "Дверь 1 створч.", "Дверь 2-х створч."]
FACADE_SYSTEMS = ["Ruit 50F"]

# Настройка страницы
st.set_page_config(page_title="Axis Pro - Калькулятор систем", layout="wide", page_icon="🏗️")

# Проверка авторизации
if not authenticate():
    st.stop()

# Загрузка данных
@st.cache_data(ttl=600)
def get_references():
    r1 = load_reference_1()
    r2 = load_reference_2()
    r3 = load_reference_3()
    rf = load_facade_reference()
    return r1, r2, r3, rf

ref1, ref2, ref3, ref_facade = get_references()

# ========================================
# Вспомогательные функции V.9
# ========================================

def render_v9_financial_table(totals, aggregator):
    """Отрисовка финальной таблицы Часть 4 по ТЗ V.9"""
    final_items = []
    
    # Стеклопакет
    if totals['breakdown'].get('glass', 0) > 0:
        final_items.append({'Наименование': 'Стеклопакет', 'Сумма (₸)': f"{totals['breakdown']['glass']:,}"})
    
    # Ламбри
    if totals['breakdown'].get('lambri', 0) > 0:
        final_items.append({'Наименование': 'Ламбри', 'Сумма (₸)': f"{totals['breakdown']['lambri']:,}"})
        
    # Тонировка
    if totals['breakdown'].get('tinting', 0) > 0:
        final_items.append({'Наименование': 'Тонировка', 'Сумма (₸)': f"{totals['breakdown']['tinting']:,}"})
        
    # Сборка
    if totals['breakdown'].get('assembly', 0) > 0:
        final_items.append({'Наименование': 'Сборка', 'Сумма (₸)': f"{totals['breakdown']['assembly']:,}"})
        
    # Монтаж
    if totals['breakdown'].get('installation', 0) > 0:
        final_items.append({'Наименование': 'Монтаж', 'Сумма (₸)': f"{totals['breakdown']['installation']:,}"})
        
    # Дополнительные детали (Периметр/3 * цена)
    if totals['breakdown'].get('additional_details', 0) > 0:
        final_items.append({'Наименование': 'Дополнительные детали', 'Сумма (₸)': f"{totals['breakdown']['additional_details']:,}"})

    # Материалы (Каркас + Вставки или просто Окна)
    if 'facade_frame' in totals['breakdown']:
        mat_sum = totals['breakdown']['facade_frame'] + totals['breakdown']['facade_inserts']
        final_items.append({'Наименование': 'Материалы (Каркас + Вставки)', 'Сумма (₸)': f"{mat_sum:,}"})
    else:
        final_items.append({'Наименование': 'Материалы', 'Сумма (₸)': f"{totals['materials_total']:,}"})

    # Обеспечение (81%)
    final_items.append({'Наименование': 'Обеспечение', 'Сумма (₸)': f"{totals['margin']:,}"})

    st.table(pd.DataFrame(final_items))
    st.subheader(f"✅ ИТОГО: {totals['total']:,} ₸")

# ========================================
# СТРАНИЦА: ОКНА / ДВЕРИ
# ========================================

def render_main_page():
    st.header("🪟 Окна и Двери (V.9 Проектный метод)")
    
    if 'positions' not in st.session_state:
        st.session_state.positions = []

    # Общие настройки на весь заказ
    with st.expander("🛠️ Общие параметры заказа", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1: t = st.checkbox("Тонировка")
        with c2: a = st.checkbox("Сборка")
        with c3: i = st.selectbox("Монтаж:", ["Нет", "Монтаж", "Демонтаж/Монтаж", "Сложный монтаж"])
        add_d = st.checkbox("Дополнительные детали (Мувиль)", value=True)

    # Форма добавления
    with st.form("win_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            p_type = st.selectbox("Тип:", PRODUCT_TYPES)
            sys_type = st.selectbox("Система:", list(SYSTEM_MAPPING.keys()))
        with col2:
            w = st.number_input("W (мм):", value=1000)
            h = st.number_input("H (мм):", value=1400)
        with col3:
            gl = st.selectbox("Стекло:", list(ref2.keys()))
            q = st.number_input("Кол-во:", value=1)
        if st.form_submit_button("Добавить позицию"):
            code = get_code_for_windows_doors(p_type, sys_type)
            st.session_state.positions.append({
                "product_type": p_type, "system": sys_type, "code": code,
                "data": {"width": w, "height": h, "qty": q, "glass_type": gl},
                "sashes": [], "imposts": {"has_center": False, "has_tor": False}
            })
            st.rerun()

    if st.session_state.positions:
        st.markdown("---")
        # РАССЧЕТ
        if st.button("🚀 РАССЧИТАТЬ ВЕСЬ ЗАКАЗ"):
            aggregator = MaterialAggregator(ref1)
            pos_details = []
            
            for idx, pos in enumerate(st.session_state.positions, 1):
                order_data = {"common": {"tinting": t, "assembly": a, "install_type": i, "add_details": add_d}, "positions": [pos]}
                res = calculate_window_smeta(order_data, ref1, ref2, ref3)
                
                # Агрегация материалов (V.9)
                aggregator.add_metrics(res['metrics']['total_area'], res['metrics']['total_perimeter'])
                aggregator.add_materials(res.get('part2_materials', []), category='windows_doors')
                
                pos_details.append({
                    "Позиция": idx, "Размеры": f"{pos['data']['width']}x{pos['data']['height']}",
                    "Площадь": f"{res['metrics']['total_area']:.2f} м²", "Периметр": f"{res['metrics']['total_perimeter']:.2f} м"
                })

            # Вывод Часть 1-2
            st.subheader("📊 ЧАСТЬ 1 & 2: Общие показатели и Список изделий")
            st.table(pd.DataFrame(pos_details))
            
            # Часть 3: Материалы
            st.subheader("📦 ЧАСТЬ 3: Сводная спецификация материалов (ОКНА)")
            basket = aggregator.get_basket()
            df_mat = pd.DataFrame([
                {"Артикул": k, "Товар": v['name'], "Расход": f"{v['quantity_raw']:.2f}", "К отгрузке": v['quantity_ship'], "Сумма": v['row_sum']}
                for k, v in basket.items()
            ])
            st.dataframe(df_mat, use_container_width=True)
            
            # Часть 4: Итог
            st.subheader("💰 ЧАСТЬ 4: Итоговый расчет")
            aggregator.common_params.update({"tinting": t, "assembly": a, "install_type": i, "add_details": add_d})
            render_v9_financial_table(aggregator.calculate_totals(ref2), aggregator)

# ========================================
# СТРАНИЦА: ФАСАДЫ
# ========================================

def render_facade_page():
    st.header("🏢 Фасадные системы (V.9 Глобальный итог)")
    
    if 'facade_positions' not in st.session_state:
        st.session_state.facade_positions = []

    with st.form("facade_v9_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            fw = st.number_input("Ширина W (м):", value=5.5)
            fh1 = st.number_input("Высота H1 (м):", value=6.0)
        with c2:
            fh2 = st.number_input("Высота H2 (м):", value=0.0)
            fcols = st.number_input("Кол-во столбцов:", value=5)
        with c3:
            frows = st.number_input("Кол-во рядов:", value=4)
            f_sys = st.selectbox("Система:", FACADE_SYSTEMS)
        
        if st.form_submit_button("Добавить фасад"):
            st.session_state.facade_positions.append({
                "width": fw, "height_left": fh1, "height_right": fh2,
                "columns": fcols, "rows": frows, "system": f_sys
            })
            st.rerun()

    if st.session_state.facade_positions:
        if st.button("🚀 РАССЧИТАТЬ ФАСАДЫ (ЕДИНЫМ ИТОГОМ)"):
            aggregator = MaterialAggregator(ref1)
            facade_info = []
            
            for idx, pos in enumerate(st.session_state.facade_positions, 1):
                res = calculate_facade_materials(pos, ref1, ref2, ref3)
                
                # Разделение Каркас / Вставки по ТЗ
                aggregator.add_metrics(res['area'], res['perimeter'])
                aggregator.add_materials(res.get('skeleton_items', []), category='facade_frame')
                aggregator.add_materials(res.get('insert_items', []), category='facade_inserts')
                
                facade_info.append({
                    "Позиция": idx, "Размеры": f"{pos['width']}x{pos['height_left']} м",
                    "Сетка": f"{pos['columns']}x{pos['rows']}", "Площадь": f"{res['area']:.2f} м²"
                })

            st.subheader("🔹 ЧАСТЬ 1 & 2: Детализация проекта")
            st.table(pd.DataFrame(facade_info))

            st.subheader("📦 ЧАСТЬ 3.1: Материалы КАРКАСА (Общий)")
            df_skeleton = pd.DataFrame([
                {"Артикул": k, "Товар": v['name'], "Брутто": v['quantity_ship'], "Сумма": v['row_sum']}
                for k, v in aggregator.get_basket('facade_frame').items()
            ])
            st.dataframe(df_skeleton, use_container_width=True)

            st.subheader("📦 ЧАСТЬ 3.2: Материалы ВСТАВОК (Общий)")
            df_inserts = pd.DataFrame([
                {"Артикул": k, "Товар": v['name'], "Кол-во": v['quantity_ship'], "Сумма": v['row_sum']}
                for k, v in aggregator.get_basket('facade_inserts').items()
            ])
            st.dataframe(df_inserts, use_container_width=True)

            st.subheader("💰 ЧАСТЬ 4: Итоговый расчет (ФАСАД)")
            aggregator.common_params.update({"add_details": True})
            render_v9_financial_table(aggregator.calculate_totals(ref2), aggregator)

# ========================================
# ОКОННЫЙ ТАМБУР
# ========================================

def render_tambour_page():
    st.header("🚪 Оконный тамбур (V.9)")
    st.info("Расчет ведется по методу агрегации материалов для всего тамбура целиком.")
    # Аналогичная логика агрегации для тамбура

# ========================================
# РОУТИНГ И МЕНЮ
# ========================================

with st.sidebar:
    st.title("📍 AXIS PRO V.9")
    nav = st.radio("Разделы:", ["Главная (Окна/Двери)", "Фасады", "Оконный тамбур", "История"])

if nav == "Главная (Окна/Двери)":
    render_main_page()
elif nav == "Фасады":
    render_facade_page()
elif nav == "Оконный тамбур":
    render_tambour_page()
else:
    render_history_page()
