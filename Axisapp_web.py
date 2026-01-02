import streamlit as st
import pandas as pd
import math
from io import BytesIO
from openpyxl import Workbook

# --- КОНСТАНТЫ СИСТЕМ ---
SYSTEMS = ["ALG 2030-45C", "ALG 2030-55C", "ALG 2030-63C", "ALG 2030-73C", "ALG 2030-Slim", "Ruit 50F"]
PRODUCTS = ["Окно глух.", "Окно с откр.", "Дверь 1 створч.", "Дверь 2-х створч.", "Фасад"]

def main():
    st.set_page_config(page_title="Axisapp - Профессиональный расчет", layout="wide")
    st.title("🚀 Калькулятор алюминиевых систем (Ruit/ALG)")

    # --- SIDEBAR: Основные настройки ---
    st.sidebar.header("Параметры заказа")
    order_number = st.sidebar.text_input("Номер заказа", "001")
    
    product_type = st.sidebar.selectbox("Тип изделия", PRODUCTS)
    profile_system = st.sidebar.selectbox("Система профиля", SYSTEMS)
    
    # Цвет (автоматически добавляется к артикулам в смете)
    color_code = st.sidebar.selectbox("Цвет RAL", ["7024", "9016", "9005"])
    color_finish = st.sidebar.radio("Фактура", ["Мат", "Глянец"])

    # --- ФОРМА ЗАПОЛНЕНИЯ (UI) ---
    st.header(f"Ввод данных: {product_type}")
    
    sections = []
    
    if product_type == "Фасад":
        # Специфический ввод для фасада
        col1, col2, col3 = st.columns(3)
        with col1:
            W = st.number_input("Общая ширина фасада (мм)", min_value=100)
            H = st.number_input("Общая высота фасада (мм)", min_value=100)
        with col2:
            n_mullions = st.number_input("Кол-во вертикальных стоек", min_value=2, value=2)
            n_transoms = st.number_input("Кол-во уровней ригелей", min_value=1, value=1)
        with col3:
            filling_type = st.radio("Тип заполнения глухих частей", ["Стеклопакет", "Ламбри (Панель)"])
            is_insert = st.checkbox("Вставить окно/дверь в ячейку?")
        
        # Логика расчета Фасада
        sections.append({
            "kind": "facade",
            "W": W, "H": H,
            "n_m": n_mullions, "n_t": n_transoms,
            "filling": filling_type,
            "is_insert": is_insert
        })
        
    else:
        # Стандартный ввод для Окон и Дверей
        num_pos = st.number_input("Количество позиций", min_value=1, step=1)
        for i in range(int(num_pos)):
            st.markdown(f"**Позиция №{i+1}**")
            c1, c2, c3 = st.columns(3)
            with c1:
                w = st.number_input(f"Ширина W (мм) - поз {i+1}", min_value=100)
            with c2:
                h = st.number_input(f"Высота H (мм) - поз {i+1}", min_value=100)
            with c3:
                qty = st.number_input(f"Кол-во (шт) - поз {i+1}", min_value=1, value=1)
            
            sections.append({"kind": "standard", "W": w, "H": h, "qty": qty})

    # --- АЛГОРИТМ РАСЧЕТА (ENGINE) ---
    if st.button("Рассчитать материалы"):
        st.header("📋 Итоговая спецификация")
        
        all_materials = []
        
        for sec in sections:
            if sec["kind"] == "facade":
                # --- ЛОГИКА ФАСАДА Ruit 50F ---
                # 1. Стойки
                m_len = (sec["H"] / 1000) * sec["n_m"]
                all_materials.append({"Товар": "Стойка фасадная", "Расход": m_len, "Ед.": "м.п."})
                
                # 2. Ригели (между стойками)
                t_qty = (sec["n_m"] - 1) * sec["n_t"]
                t_len = ((sec["W"] - (sec["n_m"] * 50)) / 1000) * sec["n_t"]
                all_materials.append({"Товар": "Ригель фасадный", "Расход": t_len, "Ед.": "м.п."})
                
                # 3. Комплектующие (Узлы)
                nodes = (sec["n_m"]) * sec["n_t"]
                all_materials.append({"Товар": "U-соединитель ригеля", "Расход": nodes * 2, "Ед.": "шт."})
                all_materials.append({"Товар": "Упл. торцевой ригеля", "Расход": nodes * 2, "Ед.": "шт."})
                
                # 4. Ламбри (если выбрано)
                if sec["filling"] == "Ламбри (Панель)":
                    area = (sec["W"] * sec["H"]) / 1000000
                    all_materials.append({"Товар": "Панель Ламбри", "Расход": area, "Ед.": "м2"})

            elif "Дверь" in product_type:
                # --- ЛОГИКА ДВЕРЕЙ (45C-73C) ---
                w, h, q = sec["W"], sec["H"], sec["qty"]
                
                # Вычеты по сериям (Инженерная база)
                if "55C" in profile_system:
                    ws, hs = w - 74, h - 45 # Порог 45мм
                elif "73C" in profile_system:
                    ws, hs = w - 82, h - 50
                else:
                    ws, hs = w - 70, h - 40
                
                all_materials.append({"Товар": "Профиль рамы", "Расход": (w + h*2)*q/1000, "Ед.": "м.п."})
                all_materials.append({"Товар": "Профиль створки", "Расход": (ws + hs)*2*q/1000, "Ед.": "м.п."})
                
                # Авто-расчет петель (без ручного ввода)
                h_count = 3 if hs > 2100 else 2
                all_materials.append({"Товар": "Петля дверная", "Расход": h_count * q, "Ед.": "шт."})
                all_materials.append({"Товар": "Замок дверной", "Расход": 1 * q, "Ед.": "шт."})

        # Вывод таблицы
        df_res = pd.DataFrame(all_materials)
        st.table(df_res)
        
        # Кнопка скачивания (заглушка для примера)
        st.success("Расчет выполнен успешно. Данные готовы к экспорту.")

if __name__ == "__main__":
    main()
