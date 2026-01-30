import streamlit as st
import sys
import os
import pandas as pd
from pathlib import Path
import datetime
import tempfile
import math  # ДОБАВЛЕНО

# --- ФИКСАЦИЯ ПУТЕЙ (Стандарт Axis Pro GF) ---
current_file = Path(__file__).resolve()
root_dir = current_file.parents[1] 
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

# Импорты внутренних модулей
from auth.auth import authenticate
from config.settings import SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH
from references.sheets_reader import load_reference_1, load_reference_2, load_reference_3, load_facade_reference  # ДОБАВЛЕНО load_facade_reference
from calculations.engine_windows import calculate_window_smeta, calculate_impost_length, SYSTEM_MAPPING
from calculations.engine_facade import calculate_facade_materials, calculate_tambour_materials, calculate_tambour_materials_v2  # ДОБАВЛЕНО
from calculations.material_basket import MaterialAggregator as MaterialBasket
from calculations.mapping import get_code_for_windows_doors, get_code_for_facade
from export.export_kp import export_to_excel, export_facade_to_excel
from history.save_history import save_history

# --- КОНСТАНТЫ ИЗ ТЗ ---
# ИСПРАВЛЕНО: "Окно глух." теперь с ОДНИМ пробелом (как в Справочнике-1)
PRODUCT_TYPES = ["Окно с откр.", "Окно глух.", "Дверь 2-х створч.", "Дверь 1 створч.", "Фасад"]
PROFILE_SYSTEMS = [
    "ALG RUIT 73i 22MM",
    "ALG RUIT 63i", 
    "ALG RUIT 55i", 
    "ALG RUIT 45i",
    "ALG 2030-73C", 
    "ALG 2030-63C", 
    "ALG 2030-55C", 
    "ALG 2030-45C", 
    "ALG 2030-Slim", 
    "Ruit 50F"
]
# GLASS_TYPES теперь загружаются динамически из ref2 (удалён хардкод)
PANELS = ["Стеклопакет", "Ламбри без термо", "Ламбри с термо"]
TONING = ["Есть", "Нет"]
ASSEMBLY = ["Есть", "Нет"]
INSTALLATION = ["Нет", "Монтаж", "Демонтаж", "Демонтаж / Монтаж", "Сложный монтаж"]

# --- ДИЗАЙН И СТИЛИЗАЦИЯ ---
st.set_page_config(page_title="Axis Pro GF", layout="wide")

page_bg_img = """
<style>
[data-testid="stAppViewContainer"] {
    background-image: linear-gradient(rgba(255,255,255,0.85), rgba(255,255,255,0.85)), 
    url("https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?q=80&w=2070&auto=format&fit=crop");
    background-size: cover;
    background-attachment: fixed;
}
[data-testid="stHeader"] {
    background: rgba(0,0,0,0);
}
.stExpander, .stContainer, div[data-testid="stVerticalBlock"] > div {
    background-color: rgba(255, 255, 255, 0.7);
    border-radius: 10px;
    padding: 10px;
}
</style>
"""
st.markdown(page_bg_img, unsafe_allow_html=True)

# --- 1. АВТОРИЗАЦИЯ ---
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔐 Вход в Axis Pro GF")
    col1, col2 = st.columns(2)
    login = col1.text_input("Логин")
    password = col2.text_input("Пароль", type="password")
    if st.button("Войти"):
        user = authenticate(login, password, GOOGLE_CREDENTIALS_PATH, SPREADSHEET_ID)
        if user:
            st.session_state.authenticated = True
            st.session_state.current_user = {"login": login, "data": user}
            st.rerun()
        else:
            st.error("Ошибка входа")
    st.stop()

# --- 2. ЗАГРУЗКА ДАННЫХ ---

@st.cache_data(ttl=60)  # Кеш на 60 секунд (обновляется каждую минуту)
def get_data():
    r1 = load_reference_1(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r2_raw = load_reference_2(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r3 = load_reference_3(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r_facade = load_facade_reference(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)  # ДОБАВЛЕНО
    
    # КРИТИЧНО: Нормализуем ВСЕ ключи ref2 в lowercase
    r2 = {k.lower(): v for k, v in r2_raw.items()}
    
    return r1, r2, r3, r_facade  # ДОБАВЛЕН r_facade

ref1, ref2, ref3, ref_facade = get_data()  # ДОБАВЛЕН ref_facade

# ДИНАМИЧЕСКАЯ загрузка типов стеклопакетов из ref2
def get_glass_types():
    """Получает список типов стеклопакетов из ref2"""
    glass_types = []
    for key in ref2.keys():
        # Пропускаем служебные ключи
        if key not in ['тонировка', 'сборка', 'монтаж', 'демонтаж/монтаж', 'сложный монтаж', 'нащельник']:
            # Капитализируем первую букву
            glass_types.append(key.capitalize())
    # Сортируем для стабильности
    return sorted(glass_types, key=lambda x: (x == 'Нет', x))

GLASS_TYPES = get_glass_types()

# --- 3. ФУНКЦИЯ КОНСТРУКТОР ОКНА ---
def window_door_ui(prefix, pos_idx, system_id, initial_data=None):
    """
    Форма для заполнения данных окна/двери
    
    Args:
        prefix: префикс для ключей виджетов
        pos_idx: индекс позиции (не используется)
        system_id: ID системы профиля
        initial_data: начальные данные для заполнения формы (для сохранения состояния)
    """
    st.markdown("---")
    
    # Получаем начальные значения из initial_data или используем дефолтные
    if initial_data is None:
        initial_data = {}
    
    # Габариты
    st.markdown("### 📐 Габариты изделия")
    c1, c2 = st.columns(2)
    w = c1.number_input(
        "Ширина (мм)", 
        min_value=0.0, 
        value=float(initial_data.get("width", 2000.0)), 
        step=50.0, 
        key=f"{prefix}_w"
    )
    h = c2.number_input(
        "Высота (мм)", 
        min_value=0.0, 
        value=float(initial_data.get("height", 1560.0)), 
        step=50.0, 
        key=f"{prefix}_h"
    )

    # Импосты
    st.markdown("### 🔲 Импосты")
    
    initial_imposts = initial_data.get("imposts", {})
    auto_default = initial_imposts.get("auto_calculate", True) if initial_imposts else True
    
    auto_imposts = st.checkbox(
        "✅ Автоматический расчет (рекомендуется)", 
        value=auto_default, 
        key=f"{prefix}_auto_imp"
    )
    
    if auto_imposts:
        st.caption("💡 Длина импостов рассчитывается автоматически по системе профиля")
        ic1, ic2, ic3, ic4 = st.columns(4)
        
        has_left = ic1.checkbox(
            "Левый", 
            value=initial_imposts.get("has_left", False),
            key=f"{prefix}_has_il"
        )
        has_center = ic2.checkbox(
            "Центральный", 
            value=initial_imposts.get("has_center", False),
            key=f"{prefix}_has_ic"
        )
        has_right = ic3.checkbox(
            "Правый", 
            value=initial_imposts.get("has_right", False),
            key=f"{prefix}_has_ir"
        )
        has_tor = ic4.checkbox(
            "ТОР (гориз.)", 
            value=initial_imposts.get("has_tor", False),
            key=f"{prefix}_has_it"
        )
        
        imposts_data = {
            "auto_calculate": True,
            "has_left": has_left,
            "has_center": has_center,
            "has_right": has_right,
            "has_tor": has_tor
        }
        
        # Показываем рассчитанные значения
        if any([has_left, has_center, has_right, has_tor]):
            st.caption("**Рассчитанные длины:**")
            calc_cols = st.columns(4)
            if has_left:
                imp_len = calculate_impost_length(w, h, system_id, "vertical")
                calc_cols[0].info(f"Левый: {imp_len:.0f} мм")
            if has_center:
                imp_len = calculate_impost_length(w, h, system_id, "vertical")
                calc_cols[1].info(f"Центр: {imp_len:.0f} мм")
            if has_right:
                imp_len = calculate_impost_length(w, h, system_id, "vertical")
                calc_cols[2].info(f"Правый: {imp_len:.0f} мм")
            if has_tor:
                imp_len = calculate_impost_length(w, h, system_id, "horizontal")
                calc_cols[3].info(f"ТОР: {imp_len:.0f} мм")
    else:
        st.caption("✋ Ручной ввод длин импостов (для нестандартных конструкций)")
        i1, i2, i3, i4 = st.columns(4)
        il = i1.number_input(
            "Левый (мм)", 
            min_value=0, 
            value=int(initial_imposts.get("left", 0)), 
            step=50, 
            key=f"{prefix}_il"
        )
        ic = i2.number_input(
            "Центр (мм)", 
            min_value=0, 
            value=int(initial_imposts.get("center", 0)), 
            step=50, 
            key=f"{prefix}_ic"
        )
        ir = i3.number_input(
            "Правый (мм)", 
            min_value=0, 
            value=int(initial_imposts.get("right", 0)), 
            step=50, 
            key=f"{prefix}_ir"
        )
        it = i4.number_input(
            "ТОР (мм)", 
            min_value=0, 
            value=int(initial_imposts.get("tor", 0)), 
            step=50, 
            key=f"{prefix}_it"
        )
        
        imposts_data = {
            "auto_calculate": False,
            "left": il,
            "center": ic,
            "right": ir,
            "tor": it
        }
    
    # Створки
    st.markdown("### 🚪 Створки")
    
    initial_sashes = initial_data.get("sashes", [])
    s_count = st.number_input(
        "Количество створок", 
        min_value=0, 
        max_value=10, 
        value=len(initial_sashes) if initial_sashes else 1, 
        step=1, 
        key=f"{prefix}_sc"
    )
    sashes = []
    
    if s_count > 0:
        st.caption("💡 Для расчета точек запирания и фурнитуры используется первая створка")
        for s in range(s_count):
            # Получаем начальные значения для этой створки
            if s < len(initial_sashes):
                initial_sash = initial_sashes[s]
                default_w = int(initial_sash.get("w", 952))
                default_h = int(initial_sash.get("h", 512))
            else:
                default_w = 952
                default_h = 512
            
            with st.expander(f"Створка №{s+1}", expanded=(s==0)):
                sc1, sc2 = st.columns(2)
                sw = sc1.number_input(
                    f"Ширина", 
                    min_value=0, 
                    value=default_w, 
                    step=50, 
                    key=f"{prefix}_sw{s}"
                )
                sh = sc2.number_input(
                    f"Высота", 
                    min_value=0, 
                    value=default_h, 
                    step=50, 
                    key=f"{prefix}_sh{s}"
                )
                sashes.append({"w": sw, "h": sh})
    
    # Заполнение
    st.markdown("### 🖼 Заполнение")
    
    initial_fill = initial_data.get("fill_category", "Стеклопакет")
    try:
        fill_index = PANELS.index(initial_fill)
    except ValueError:
        fill_index = 0
    
    fill_cat = st.selectbox(
        "Тип заполнения", 
        PANELS, 
        index=fill_index,
        key=f"{prefix}_fill_cat"
    )
    
    selected_glass = "Нет"
    if fill_cat == "Стеклопакет":
        initial_glass = initial_data.get("glass_type", "Двойной")
        try:
            glass_index = GLASS_TYPES.index(initial_glass)
        except ValueError:
            glass_index = 0
            
        selected_glass = st.selectbox(
            "Тип стеклопакета", 
            GLASS_TYPES, 
            index=glass_index,
            key=f"{prefix}_glass"
        )
        
        # Показываем информацию о световом проеме
        normalized_sys = SYSTEM_MAPPING.get(system_id, system_id)
        offset = {"ALG 2030-73C": 73, "ALG 2030-63C": 63, "ALG 2030-55C": 55, "ALG 2030-45C": 45}.get(normalized_sys, 73)
        w_g = w - (offset * 2)
        h_g = h - (offset * 2)
        st.caption(f"💡 Световой проем (автоматически): **{w_g:.0f} × {h_g:.0f} мм** (габарит - {offset*2} мм)")
    
    return {
        "width": w, 
        "height": h, 
        "imposts": imposts_data, 
        "sashes": sashes, 
        "fill_category": fill_cat, 
        "glass_type": selected_glass
    }

# ========================================
# ФУНКЦИЯ ДЛЯ ГЛАВНОЙ СТРАНИЦЫ (ОКНА/ДВЕРИ)
# ========================================
def render_windows_doors_page():
    """Главная страница - расчет окон и дверей"""
    
    # --- 4. ШАПКА И КНОПКИ ---
    header_col1, header_col2 = st.columns([3, 1])
    with header_col1:
        st.title("🚀 Axis Pro GF - Калькулятор окон V2")
        st.markdown("""
        **Компания «AXIS»** 📍 Город: Астана  
        📞 Тел.: +7 707 504 4040 | 📧 E-mail: Axisokna.kz@mail.ru | 🌐 Сайт: www.axis.kz
        """)
    with header_col2:
        if st.button("🔄 Очистить и Новый расчет", width="stretch"):
            for key in list(st.session_state.keys()):
                if key not in ['authenticated', 'current_user', 'menu_selection']:
                    del st.session_state[key]
            st.rerun()

    st.divider()

    # --- 5. ОСНОВНОЙ ИНТЕРФЕЙС ---
    col_left, col_right = st.columns([1, 2.5])

    with col_left:
        st.subheader("📋 Данные заказа")
        with st.container():
            order_num = st.text_input("Номер заказа", value="001", key="main_order_num")
            
            st.markdown("#### Общие параметры:")
            toning_id = st.selectbox("Тонировка", TONING, key="main_toning")
            assembly_id = st.selectbox("Сборка", ASSEMBLY, key="main_assembly")
            install_id = st.selectbox("Монтаж", INSTALLATION, key="main_install")
            
            # ДОБАВЛЕНО: Дополнительные детали
            additional_options = ["Нет"] + [k.capitalize() for k in ref2.keys() if "нащельник" in k.lower()]
            additional_id = st.selectbox("Дополнительные детали", additional_options, key="main_additional")

    with col_right:
        st.subheader(f"🪟 Список позиций")
        
        if "positions" not in st.session_state: 
            st.session_state.positions = []
        
        if st.button("➕ Добавить позицию", width="stretch"):
            st.session_state.positions.append({
                "count": 1,
                "product_type": "Окно с откр.",
                "system_id": "ALG RUIT 73i 22MM"
            })
            st.rerun()
        
        if not st.session_state.positions:
            st.info("👆 Нажмите кнопку выше, чтобы добавить первую позицию")
        
        for idx, pos in enumerate(st.session_state.positions):
            with st.expander(f"📦 Позиция №{idx+1}", expanded=True):
                # Кнопка удаления позиции
                if st.button(f"🗑️ Удалить позицию", key=f"del_pos_{idx}"):
                    st.session_state.positions.pop(idx)
                    st.rerun()
                
                # Тип изделия и система на уровне позиции
                pc1, pc2 = st.columns(2)
                
                # Определяем текущий индекс для типа изделия
                current_type = pos.get("product_type", "Окно с откр.")
                try:
                    type_index = PRODUCT_TYPES.index(current_type)
                except ValueError:
                    type_index = 0
                
                product_type = pc1.selectbox(
                    "Тип изделия", 
                    PRODUCT_TYPES,
                    key=f"pc_type{idx}",
                    index=type_index
                )
                pos["product_type"] = product_type
                st.session_state.positions[idx]["product_type"] = product_type
                
                system_id = pc2.selectbox(
                    "Система профиля", 
                    PROFILE_SYSTEMS, 
                    key=f"pc_sys{idx}",
                    index=0
                )
                pos["system_id"] = system_id
                st.session_state.positions[idx]["system_id"] = system_id
                
                # Генерация CODE для расчётов (mapping UI → CODE)
                if pos["product_type"] != "Фасад":
                    code = get_code_for_windows_doors(
                        pos["product_type"],
                        pos["system_id"]
                    )
                    pos["code"] = code
                    # КРИТИЧНО: сохраняем CODE обратно в session_state
                    st.session_state.positions[idx]["code"] = code
                
                # 1 позиция = 1 изделие
                pos["count"] = 1
                st.session_state.positions[idx]["count"] = 1
                
                data = window_door_ui(f"main_pos_{idx}", idx, pos["system_id"])
                pos["data"] = data
                # КРИТИЧНО: сохраняем data обратно в session_state
                st.session_state.positions[idx]["data"] = data


    # --- 6. РАСЧЕТ И ВЫВОД ---
    st.divider()

    if st.button("🚀 РАССЧИТАТЬ", type="primary", width="stretch"):
        if not st.session_state.positions:
            st.error("❌ Добавьте хотя бы одну позицию!")
        else:
            # Формирование данных заказа
            # ИСПРАВЛЕНО: Конвертируем _id в правильные ключи
            order_data = {
                "common": {
                    "order_number": order_num,
                    "toning": toning_id,         # "Есть" или "Нет"
                    "assembly": assembly_id,     # "Есть" или "Нет"
                    "installation": install_id   # "Монтаж простой" или "Нет"
                },
                "positions": st.session_state.get("positions", [])
            }
            
            # Расчет через новый движок для окон V2 + ГЛОБАЛЬНАЯ КОРЗИНА
            try:
                # === ГЛОБАЛЬНАЯ КОРЗИНА МАТЕРИАЛОВ ===
                # Устраняет 5× перерасход профилей за счёт округления один раз
                basket = MaterialBasket(ref1)
                all_results = []
                
                # Рассчитываем каждую позицию отдельно
                for position in st.session_state.positions:
                    pos_order_data = {
                        "common": order_data["common"],
                        "positions": [position]
                    }
                    pos_result = calculate_window_smeta(pos_order_data, ref1, ref2, ref3)
                    all_results.append(pos_result)
                    
                    # Добавляем материалы в корзину (БЕЗ округления!)
                    for material in pos_result.get("part2_materials", []):
                        basket.add_material(
                            article=material.get("Артикул", ""),
                            quantity_raw=material.get("Количество_raw", material.get("Количество", 0)),
                            unit=material.get("Единица", "шт"),
                            price=material.get("Цена", 0),
                            name=material.get("Элемент", "")
                        )
                
                # Округляем материалы ОДИН РАЗ
                basket.round_all_materials()
                basket_costs = basket.calculate_costs()
                
                # Объединяем результаты
                if not all_results:
                    raise Exception("Нет результатов расчёта")
                
                # Суммируем метрики из всех результатов
                total_area = sum(r.get("metrics", {}).get("total_area", 0) for r in all_results)
                total_perimeter = sum(r.get("metrics", {}).get("total_perimeter", 0) for r in all_results)
                
                # Используем материалы из корзины (с правильным округлением!)
                materials_list = basket.get_materials_list()
                materials_cost = basket_costs["total_materials_cost"]
                
                # Стекло, услуги берём из суммы всех результатов
                total_glass = sum(r.get("part3_final", {}).get("Стеклопакет", 0) for r in all_results)
                total_lambri = sum(r.get("part3_final", {}).get("Ламбри", 0) for r in all_results)
                total_toning = sum(r.get("part3_final", {}).get("Тонировка", 0) for r in all_results)
                total_assembly = sum(r.get("part3_final", {}).get("Сборка", 0) for r in all_results)
                total_install = sum(r.get("part3_final", {}).get("Монтаж", 0) for r in all_results)
                total_additional = sum(r.get("part3_final", {}).get("Дополнительные детали", 0) for r in all_results)
                
                # Себестоимость
                subtotal = materials_cost + total_glass + total_lambri + total_toning + total_assembly + total_install + total_additional
                
                # Обеспечение ОДИН РАЗ
                margin = subtotal * 0.81
                total_with_margin = subtotal + margin
                
                # Формируем итоговый результат
                res = {
                    "part1_summary": [],  # Габариты из всех результатов
                    "part2_materials": materials_list,  # Материалы из корзины
                    "part3_final": {
                        "Стеклопакет": round(total_glass, 0),
                        "Ламбри": round(total_lambri, 0),
                        "Тонировка": round(total_toning, 0),
                        "Сборка": round(total_assembly, 0),
                        "Монтаж": round(total_install, 0),
                        "Дополнительные детали": round(total_additional, 0),
                        "Материалы": round(materials_cost, 0),
                        "Обеспечение (81%)": round(margin, 0)
                    },
                    "materials_cost": round(subtotal, 0),
                    "total_with_margin": round(total_with_margin, 0),
                    "metrics": {
                        "total_area": total_area,
                        "total_perimeter": total_perimeter
                    },
                    "basket_savings": basket_costs.get("total_saved_quantity", 0)
                }
                
                # Собираем габариты из всех результатов
                for result in all_results:
                    res["part1_summary"].extend(result.get("part1_summary", []))
                
                # Сохранение результата в session_state для экспорта
                st.session_state.last_result = res
                st.session_state.last_order_data = order_data
                
                # СОХРАНЕНИЕ ИСТОРИИ В GOOGLE SHEETS
                try:
                    current_user = st.session_state.get("current_user", {})
                    user_login = current_user.get("login", "unknown")
                    save_history(
                        GOOGLE_CREDENTIALS_PATH,
                        SPREADSHEET_ID,
                        user_login,
                        order_data,
                        res
                    )
                except Exception as e:
                    st.warning(f"⚠️ История не сохранена: {e}")

                # ============================================================
                # ВЫВОД РЕЗУЛЬТАТОВ ПО ТЗ V.9
                # ============================================================
                
                st.success("✅ Расчёт выполнен!")
                
                # Главная метрика
                st.metric(
                    "💰 ИТОГО К ОПЛАТЕ",
                    f"{res['total_with_margin']:,} ₸",
                    delta="Проектный метод V.9"
                )
                
                st.divider()
                
                # ============================================================
                # ЧАСТЬ 1: ОБЩИЕ МЕТРИКИ
                # ============================================================
                st.header("📊 ЧАСТЬ 1: Общие показатели")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Общая площадь", f"{res['metrics']['total_area']:.2f} м²")
                col2.metric("Общий периметр", f"{res['metrics']['total_perimeter']:.2f} м")
                col3.metric("Позиций в заказе", len(st.session_state.positions))
                
                st.divider()
                
                # ============================================================
                # ЧАСТЬ 2: ИНФОРМАЦИОННАЯ ДЕТАЛИЗАЦИЯ (БЕЗ ЦЕН!)
                # ============================================================
                with st.expander("🔹 ЧАСТЬ 2: Список изделий (информация)", expanded=False):
                    st.info("ℹ️ Справочная информация для контроля состава заказа. Цены в этом блоке НЕ указаны.")
                    
                    # Собираем детали позиций
                    position_details = []
                    for idx, position in enumerate(st.session_state.positions, 1):
                        position_details.append({
                            'Позиция': idx,
                            'Тип': position.get('product_type', ''),
                            'Ширина (мм)': position['data']['width'],
                            'Высота (мм)': position['data']['height'],
                            'Площадь (м²)': round(
                                position['data']['width'] * position['data']['height'] / 1_000_000, 3
                            ),
                            'Периметр (м)': round(
                                2 * (position['data']['width'] + position['data']['height']) / 1000, 2
                            )
                        })
                    
                    if position_details:
                        df_positions = pd.DataFrame(position_details)
                        st.dataframe(df_positions, use_container_width=True, hide_index=True)
                    else:
                        st.warning("Нет данных о позициях")
                
                st.divider()
                
                # ============================================================
                # ЧАСТЬ 3: АГРЕГИРОВАННАЯ СПЕЦИФИКАЦИЯ МАТЕРИАЛОВ
                # ============================================================
                st.header("📦 ЧАСТЬ 3: Спецификация материалов")
                
                st.info(
                    "✨ **Проектный метод:** Материалы из всех позиций суммированы и округлены ОДИН РАЗ. "
                    "Это устраняет перерасход профилей и позволяет сравнить с заводским расчётом поартикульно."
                )
                
                materials = res.get("part2_materials", [])
                
                if materials:
                    df_materials = pd.DataFrame(materials)
                    
                    # Показываем таблицу
                    st.dataframe(
                        df_materials,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Количество_raw": st.column_config.NumberColumn(
                                "Кол-во нетто",
                                help="Точное количество ДО округления",
                                format="%.3f"
                            ) if "Количество_raw" in df_materials.columns else None,
                            "Количество": st.column_config.NumberColumn(
                                "Кол-во брутто",
                                help="Количество ПОСЛЕ округления до упаковок",
                                format="%.2f"
                            ),
                            "Сумма": st.column_config.NumberColumn(
                                "Сумма (₸)",
                                format="%d"
                            )
                        }
                    )
                    
                    # Итого по материалам
                    materials_cost = res.get("part3_final", {}).get("Материалы", 0)
                    st.metric(
                        "💼 ИТОГО материалы",
                        f"{materials_cost:,} ₸",
                        help="Стоимость всех материалов после округления"
                    )
                    
                    # Показываем экономию от глобальной корзины
                    if res.get("basket_savings", 0) > 0:
                        st.success(
                            f"📉 Экономия на профилях: {res['basket_savings']:.1f}м благодаря округлению ОДИН РАЗ "
                            f"(вместо поштучного округления)"
                        )
                else:
                    st.warning("⚠️ Материалы не найдены. Возможно, система не определена в Справочнике-1.")
                
                st.divider()
                
                # ============================================================
                # ЧАСТЬ 4: ФИНАНСОВЫЙ ИТОГ (ПРОЕКТНЫЙ МЕТОД)
                # ============================================================
                st.header("💰 ЧАСТЬ 4: Финансовый итог")
                
                st.markdown("**Расчёт ведётся ОДИН РАЗ для всего блока окон/дверей:**")
                
                # Таблица итогов
                final_items = []
                part3 = res.get("part3_final", {})
                
                if part3.get('Стеклопакет', 0) > 0:
                    final_items.append({
                        'Наименование': 'Стеклопакеты',
                        'Площадь (м²)': f"{res['metrics']['total_area']:.2f}",
                        'Сумма (₸)': f"{part3['Стеклопакет']:,}"
                    })
                
                if part3.get('Ламбри', 0) > 0:
                    final_items.append({
                        'Наименование': 'Ламбри',
                        'Площадь (м²)': '-',
                        'Сумма (₸)': f"{part3['Ламбри']:,}"
                    })
                
                if part3.get('Тонировка', 0) > 0:
                    final_items.append({
                        'Наименование': 'Тонировка',
                        'Площадь (м²)': '-',
                        'Сумма (₸)': f"{part3['Тонировка']:,}"
                    })
                
                if part3.get('Сборка', 0) > 0:
                    final_items.append({
                        'Наименование': 'Сборка',
                        'Площадь (м²)': f"{res['metrics']['total_area']:.2f}",
                        'Сумма (₸)': f"{part3['Сборка']:,}"
                    })
                
                if part3.get('Монтаж', 0) > 0:
                    final_items.append({
                        'Наименование': 'Монтаж',
                        'Площадь (м²)': f"{res['metrics']['total_area']:.2f}",
                        'Сумма (₸)': f"{part3['Монтаж']:,}"
                    })
                
                if part3.get('Дополнительные детали', 0) > 0:
                    final_items.append({
                        'Наименование': 'Дополнительные детали',
                        'Площадь (м²)': '-',
                        'Сумма (₸)': f"{part3['Дополнительные детали']:,}"
                    })
                
                final_items.append({
                    'Наименование': 'Материалы',
                    'Площадь (м²)': '-',
                    'Сумма (₸)': f"{part3.get('Материалы', 0):,}"
                })
                
                # Показываем таблицу
                if final_items:
                    df_final = pd.DataFrame(final_items)
                    st.dataframe(df_final, use_container_width=True, hide_index=True)
                
                st.divider()
                
                # Обеспечение и итого
                col_a, col_b = st.columns(2)
                with col_a:
                    st.metric(
                        "Обеспечение",
                        f"{part3.get('Обеспечение (81%)', 0):,} ₸",
                        help="Наценка 81% на себестоимость (начислена ОДИН РАЗ)"
                    )
                with col_b:
                    st.metric(
                        "💰 К ОПЛАТЕ",
                        f"{res['total_with_margin']:,} ₸",
                        delta="Финальная сумма"
                    )
                
                # Детализация расчёта обеспечения
                with st.expander("ℹ️ Как рассчитано обеспечение", expanded=False):
                    materials_sum = part3.get('Материалы', 0)
                    services_sum = (
                        part3.get('Стеклопакет', 0) +
                        part3.get('Ламбри', 0) +
                        part3.get('Тонировка', 0) +
                        part3.get('Сборка', 0) +
                        part3.get('Монтаж', 0) +
                        part3.get('Дополнительные детали', 0)
                    )
                    subtotal = materials_sum + services_sum
                    
                    st.write(f"**Материалы:** {materials_sum:,} ₸")
                    st.write(f"**Услуги:** {services_sum:,} ₸")
                    st.write(f"**Себестоимость:** {subtotal:,} ₸")
                    st.write(f"**Обеспечение (81%):** {part3.get('Обеспечение (81%)', 0):,} ₸")
                    st.divider()
                    st.write(f"**ИТОГО:** {res['total_with_margin']:,} ₸")
                

            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {e}")
                st.exception(e)

    # Кнопка экспорта в Excel
    if 'last_result' in st.session_state and 'last_order_data' in st.session_state:
        st.divider()
        if st.button("📥 Скачать КП в Excel", type="secondary", width="stretch"):
            try:
                # Используем временную директорию
                temp_dir = tempfile.gettempdir()
                
                excel_file = export_to_excel(
                    st.session_state.last_order_data, 
                    st.session_state.last_result,
                    output_dir=temp_dir
                )
                
                order_num = st.session_state.last_order_data["common"]["order_number"]
                with open(excel_file, 'rb') as f:
                    st.download_button(
                        label="⬇️ Загрузить файл",
                        data=f,
                        file_name=f"KP_AXIS_{order_num}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("✅ Файл готов к скачиванию!")
            except Exception as e:
                st.error(f"Ошибка при создании файла: {e}")
                st.exception(e)


# ========================================
# ФУНКЦИЯ ДЛЯ СТРАНИЦЫ ФАСАДОВ
# ========================================
def render_facade_page():
    """Страница расчета фасадов"""
    
    st.title("🏢 Расчет Фасадов")
    st.markdown("---")
    
    # === ОБЩИЕ ПАРАМЕТРЫ ЗАКАЗА (ПЕРЕД ПОЗИЦИЯМИ!) ===
    st.subheader("📋 Общие параметры заказа")
    
    col_order, col_ton = st.columns(2)
    facade_order_num = col_order.text_input(
        "Номер заказа", 
        value=f"FAC-{datetime.datetime.now().strftime('%Y%m%d')}",
        key="facade_order_num"
    )
    
    facade_toning = col_ton.selectbox("Тонировка", TONING, key="facade_toning")
    
    col_asm, col_inst, col_add = st.columns(3)
    facade_assembly = col_asm.selectbox("Сборка", ASSEMBLY, key="facade_assembly")
    facade_installation = col_inst.selectbox("Монтаж", INSTALLATION, key="facade_installation")
    
    # ДОБАВЛЕНО: Дополнительные детали
    additional_options = ["Нет"] + [k.capitalize() for k in ref2.keys() if "нащельник" in k.lower()]
    facade_additional = col_add.selectbox("Дополнительные детали", additional_options, key="facade_additional")
    
    st.markdown("---")
    
    # Только Ruit 50F (тамбур в отдельном разделе)
    facade_type_value = "Фасадная система (Ruit 50F)"
    # ========== ФАСАДНАЯ СИСТЕМА (Ruit 50F) ==========
    st.subheader("📦 Позиции фасада")
    
    if "facade_positions" not in st.session_state:
        st.session_state.facade_positions = []
    
    st.subheader(f"Позиции фасада ({len(st.session_state.facade_positions)})")
    
    col_add, col_clear, col_new = st.columns(3)
    
    if col_add.button("➕ Добавить позицию", width="stretch"):
        # Генерация CODE для фасада
        facade_code = get_code_for_facade(facade_type_value)
        
        st.session_state.facade_positions.append({
            "code": facade_code,  # Добавляем CODE
            "facade_type": facade_type_value,  # Сохраняем тип для отображения
            "width": 6.0,
            "height_left": 3.0,        # ИЗМЕНЕНО: вместо "height"
            "height_right": 0.0,       # НОВОЕ: 0 = прямоугольник
            "columns": 3,
            "rows": 2,
            "mullion_size": 130,       # НОВОЕ: по умолчанию 130мм
            "transom_size": 85,        # НОВОЕ: по умолчанию 85мм
            "brackets_per_mullion": 2, # НОВОЕ: по умолчанию 2 кронштейна
            "filling_type": "blind",
            "cells_data": []  # Данные для каждой ячейки
        })
        st.rerun()
    
    if col_clear.button("🗑️ Очистить всё", width="stretch"):
        st.session_state.facade_positions = []
        if "last_facade_result" in st.session_state:
            del st.session_state.last_facade_result
        st.rerun()
    
    if col_new.button("🔄 Новый расчёт", width="stretch"):
        if "last_facade_result" in st.session_state:
            del st.session_state.last_facade_result
        st.rerun()
    
    if not st.session_state.facade_positions:
        st.info("👆 Нажмите кнопку выше, чтобы добавить первую позицию фасада")
    
    # === ОТОБРАЖЕНИЕ ПОЗИЦИЙ ===
    for idx, pos in enumerate(st.session_state.facade_positions):
        with st.expander(f"📦 Позиция фасада №{idx+1}", expanded=True):
            
            # Кнопка удаления
            if st.button(f"🗑️ Удалить", key=f"del_fac_{idx}"):
                st.session_state.facade_positions.pop(idx)
                st.rerun()
            
            # === ГАБАРИТЫ ФАСАДА ===
            st.markdown("### Габариты фасада")
            col1, col2, col3 = st.columns(3)  # ИЗМЕНЕНО: 3 колонки вместо 2
            
            pos["width"] = col1.number_input(
                "Ширина (м)", 
                min_value=0.5, 
                max_value=50.0,
                value=pos.get("width", 6.0), 
                step=0.1, 
                key=f"fac_w_{idx}"
            )
            
            pos["height_left"] = col2.number_input(  # ИЗМЕНЕНО: height → height_left
                "Высота слева (м)", 
                min_value=0.5, 
                max_value=20.0,
                value=pos.get("height_left", pos.get("height", 3.0)),  # fallback для старых данных
                step=0.1, 
                key=f"fac_h1_{idx}",  # ИЗМЕНЕНО: fac_h → fac_h1
                help="Высота фасада с левой стороны"
            )
            
            pos["height_right"] = col3.number_input(  # НОВОЕ ПОЛЕ
                "Высота справа (м)", 
                min_value=0.0,  # можно 0!
                max_value=20.0,
                value=pos.get("height_right", 0.0),
                step=0.1, 
                key=f"fac_h2_{idx}",
                help="Оставьте 0 для прямоугольного фасада"
            )
            
            # НОВОЕ: Предупреждение для высоких фасадов
            max_h = max(pos["height_left"], pos["height_right"] if pos["height_right"] > 0 else pos["height_left"])
            if max_h > 4.5:
                st.warning(
                    "⚠️ Внимание: Данный калькулятор предназначен для предварительного расчета "
                    "конструкций высотой до 4.5 метров. При высоте свыше 4.5 м требуется "
                    "подтверждение заводской программой Logikal."
                )
            
            # === СЕТКА ===
            st.markdown("### Разбивка на ячейки")
            col3, col4 = st.columns(2)
            pos["columns"] = col3.number_input(
                "Количество столбцов", 
                min_value=1, 
                max_value=20,
                value=pos.get("columns", 3), 
                step=1, 
                key=f"fac_col_{idx}"
            )
            pos["rows"] = col4.number_input(
                "Количество рядов", 
                min_value=1, 
                max_value=10,
                value=pos.get("rows", 2), 
                step=1, 
                key=f"fac_row_{idx}"
            )
            
            # === РАСЧЁТ РАЗМЕРА ЯЧЕЙКИ ===
            # ИЗМЕНЕНО: используем среднюю высоту для трапеции
            h_left = pos.get("height_left", pos.get("height", 3.0))  # fallback для старых данных
            h_right = pos.get("height_right", 0.0)
            h_avg = (h_left + h_right) / 2 if h_right > 0 else h_left
            
            cell_w_m = pos["width"] / pos["columns"]
            cell_h_m = h_avg / pos["rows"]  # ИЗМЕНЕНО: используем h_avg
            cell_w_mm = cell_w_m * 1000
            cell_h_mm = cell_h_m * 1000
            
            # НОВОЕ: Показываем форму фасада
            if h_right > 0 and abs(h_left - h_right) > 0.01:
                st.info(f"📐 Форма: Трапеция ({h_left:.2f}м → {h_right:.2f}м) | Размер ячейки: {cell_w_m:.2f} × {cell_h_m:.2f} м ({cell_w_mm:.0f} × {cell_h_mm:.0f} мм)")
            else:
                st.info(f"📐 Форма: Прямоугольник | Размер ячейки: {cell_w_m:.2f} × {cell_h_m:.2f} м ({cell_w_mm:.0f} × {cell_h_mm:.0f} мм)")
            
            # === НОВОЕ: ВЫБОР ПРОФИЛЕЙ КАРКАСА ===
            st.markdown("---")
            st.markdown("### 🔧 Выбор профилей каркаса")
            
            col_mullion, col_transom, col_bracket = st.columns(3)
            
            # === АВТОМАТИЧЕСКИЕ РЕКОМЕНДАЦИИ (ЭТАП 3) ===
            # Расчёт рекомендуемых размеров на основе габаритов
            W = pos.get("width", 6.0)
            H1 = pos.get("height_left", pos.get("height", 3.5))
            H2 = pos.get("height_right", 0.0)
            cols = pos.get("columns", 3)
            
            h_avg = (H1 + H2) / 2 if H2 > 0 else H1
            width_cell = W / cols if cols > 0 else W
            
            # Рекомендуемая стойка по высоте
            if h_avg <= 2.5:
                recommended_mullion = 90
                mullion_reason = "высота до 2.5м"
            elif h_avg <= 3.5:
                recommended_mullion = 110
                mullion_reason = "высота 2.5-3.5м"
            elif h_avg <= 4.5:
                recommended_mullion = 130
                mullion_reason = "высота 3.5-4.5м"
            elif h_avg <= 6.0:
                recommended_mullion = 150
                mullion_reason = "высота 4.5-6.0м"
            elif h_avg <= 8.0:
                recommended_mullion = 180
                mullion_reason = "высота 6.0-8.0м"
            else:
                recommended_mullion = 210
                mullion_reason = "высота более 8.0м"
            
            # Рекомендуемый ригель по ширине ячейки
            if width_cell <= 0.8:
                recommended_transom = 50
                transom_reason = "ячейка до 0.8м"
            elif width_cell <= 1.2:
                recommended_transom = 70
                transom_reason = "ячейка 0.8-1.2м"
            elif width_cell <= 1.5:
                recommended_transom = 85
                transom_reason = "ячейка 1.2-1.5м"
            elif width_cell <= 2.0:
                recommended_transom = 105
                transom_reason = "ячейка 1.5-2.0м"
            elif width_cell <= 2.5:
                recommended_transom = 135
                transom_reason = "ячейка 2.0-2.5м"
            else:
                recommended_transom = 155
                transom_reason = "ячейка более 2.5м"
            
            # Рекомендуемое количество кронштейнов
            recommended_brackets = max(2, int(h_avg / 1.5) + 1)
            
            # Показываем рекомендации
            st.info(
                f"💡 **Автоматические рекомендации для вашего фасада:**\n\n"
                f"• **Стойка {recommended_mullion}мм** — {mullion_reason}\n"
                f"• **Ригель {recommended_transom}мм** — {transom_reason} (ширина ячейки {width_cell:.2f}м)\n"
                f"• **Кронштейны {recommended_brackets} шт** — для высоты {h_avg:.2f}м"
            )
            
            # Добавляем таблицу рекомендаций
            with st.expander("📊 Полная таблица подбора профилей"):
                st.markdown("### Стойки (вертикальные профили)")
                st.markdown("""
| Высота фасада (H) | Сечение стойки | Точки крепления (кронштейны) |
|-------------------|----------------|------------------------------|
| до 3.0 м          | 90 – 110 мм    | 2 шт. (верх / низ)          |
| 3.0 – 4.5 м       | 130 мм         | 2 - 3 шт.                   |
| 4.5 – 6.0 м       | 150 мм         | 3 шт. (обязателен расчёт)   |
| свыше 6.0 м       | 180 – 210 мм   | спец. кронштейны + статика  |
                """)
                
                st.markdown("### Ригели (горизонтальные профили)")
                st.markdown("""
| Ширина ячейки | Сечение ригеля |
|---------------|----------------|
| до 1.2 м      | 50 – 70 мм     |
| 1.2 – 1.8 м   | 85 – 105 мм    |
| свыше 1.8 м   | 135 – 155 мм   |
                """)
                
                st.warning(
                    "⚠️ **Внимание!**\n\n"
                    "Данный калькулятор предназначен для предварительного расчета конструкций "
                    "высотой **до 4.5 метров**. При высоте свыше 4.5 м требуется **обязательная "
                    "проверка статических нагрузок** и подтверждение спецификации заводской программой."
                )
            
            # Дополнительное предупреждение если высота > 4.5м
            if h_avg > 4.5:
                st.error(
                    f"🚨 **ТРЕБУЕТСЯ РАСЧЁТ СТАТИКИ!**\n\n"
                    f"Высота фасада {h_avg:.2f}м превышает 4.5м. "
                    f"Необходима обязательная проверка в заводской программе!"
                )
            
            # Стойка
            mullion_options = [90, 110, 130, 150, 180, 210]
            default_mullion = pos.get("mullion_size", 130)
            if default_mullion not in mullion_options:
                default_mullion = 130
            mullion_index = mullion_options.index(default_mullion)
            
            pos["mullion_size"] = col_mullion.selectbox(
                "Сечение стойки (мм)",
                options=mullion_options,
                index=mullion_index,
                key=f"fac_mullion_{idx}",
                help=(
                    "Рекомендации по высоте:\n"
                    "• 90–110 мм: до 3.0 м\n"
                    "• 130 мм: 3.0 – 4.5 м\n"
                    "• 150 мм: 4.5 – 6.0 м\n"
                    "• 180–210 мм: более 6.0 м (требуется расчет статики)"
                )
            )
            
            # Ригель
            transom_options = [50, 70, 85, 105, 135, 155]
            default_transom = pos.get("transom_size", 85)
            if default_transom not in transom_options:
                default_transom = 85
            transom_index = transom_options.index(default_transom)
            
            pos["transom_size"] = col_transom.selectbox(
                "Сечение ригеля (мм)",
                options=transom_options,
                index=transom_index,
                key=f"fac_transom_{idx}",
                help=(
                    "Рекомендации по ширине ячейки:\n"
                    "• 50–70 мм: ширина ячейки до 1.2 м\n"
                    "• 85–105 мм: 1.2 – 1.8 м\n"
                    "• 135–155 мм: панорамные ячейки / тяжелые СП"
                )
            )
            
            # Кронштейны
            pos["brackets_per_mullion"] = col_bracket.number_input(
                "Кронштейнов на 1 стойку",
                min_value=1,
                max_value=10,
                value=pos.get("brackets_per_mullion", 2),
                step=1,
                key=f"fac_brackets_{idx}",
                help="Количество кронштейнов на каждую вертикальную стойку"
            )
            
            # Динамическая проверка
            if max_h > 4.5 and pos["mullion_size"] < 150:
                st.info("💡 Рекомендуется сечение стойки не менее 150 мм для высоты более 4.5 м")
            
            # === ДОБАВЛЕНО: ВЫБОР ОСНОВНОГО ЗАПОЛНЕНИЯ ФАСАДА ===
            st.markdown("---")
            st.markdown("### 🎨 Основное заполнение фасада")
            st.caption("Это заполнение применяется ко ВСЕМ глухим ячейкам фасада")
            
            # Сохраняем в session_state для использования
            if 'facade_main_filling' not in st.session_state:
                st.session_state.facade_main_filling = {}
            
            main_panel_category = st.selectbox(
                "Категория заполнения основного фасада",
                ["Стеклопакет", "Ламбри"],
                key=f"main_panel_cat_{idx}",
                help="Выберите чем будут заполнены основные (глухие) ячейки"
            )
            
            if main_panel_category == "Стеклопакет":
                # Динамически из Справочника-2
                main_glass_type = st.selectbox(
                    "Тип стеклопакета основного фасада",
                    GLASS_TYPES,
                    key=f"main_glass_{idx}",
                    help="Тип стеклопакета для основных ячеек"
                )
                
                st.session_state.facade_main_filling[idx] = {
                    "category": "Стеклопакет",
                    "type": main_glass_type
                }
                
                st.success(f"✅ Основное заполнение: Стеклопакет - {main_glass_type}")
            
            else:
                # Ламбри - загружаем типы из ref2
                lambri_types = []
                for key in ref2.keys():
                    if "ламбри" in key.lower():
                        lambri_types.append(key)
                
                if not lambri_types:
                    lambri_types = ["Ламбри без термо", "Ламбри с термо"]
                
                main_lambri_type = st.selectbox(
                    "Тип ламбри основного фасада",
                    lambri_types,
                    key=f"main_lambri_{idx}",
                    help="Тип ламбри для основных ячеек"
                )
                
                st.session_state.facade_main_filling[idx] = {
                    "category": "Ламбри",
                    "type": main_lambri_type
                }
                
                st.success(f"✅ Основное заполнение: Ламбри - {main_lambri_type}")
            
            st.markdown("---")
            
            # === ЗАПОЛНЕНИЕ ===
            st.markdown("### Заполнение ячеек")
            
            fill_type = st.radio(
                "Тип заполнения для ВСЕХ ячеек",
                ["Глухое остекление", "Окно", "Дверь"],
                key=f"fac_fill_{idx}"
            )
            
            # === ГЛУХОЕ ОСТЕКЛЕНИЕ ===
            if fill_type == "Глухое остекление":
                pos["filling_type"] = "blind"
                
                st.info("✨ Используется основное заполнение фасада (выбрано выше)")
                
                # Берём из session_state
                main_filling = st.session_state.facade_main_filling.get(idx, {
                    "category": "Стеклопакет",
                    "type": GLASS_TYPES[0] if GLASS_TYPES else "двойной"
                })
                
                if main_filling["category"] == "Стеклопакет":
                    pos["blind_data"] = {
                        "panel_type": "glass",
                        "glass_type": main_filling["type"]
                    }
                    st.write(f"🔹 Заполнение: **Стеклопакет** - {main_filling['type']}")
                else:
                    pos["blind_data"] = {
                        "panel_type": main_filling["type"],  # "Ламбри без термо" или "Ламбри с термо"
                        "glass_type": None
                    }
                    st.write(f"🔹 Заполнение: **Ламбри** - {main_filling['type']}")
            
            # === ОКНО ИЛИ ДВЕРЬ (ВСТАВКА) ===
            elif fill_type in ["Окно", "Дверь"]:
                pos["filling_type"] = "window" if fill_type == "Окно" else "door"
                
                # === ДОБАВЛЕНО: ВЫБОР ЯЧЕЙКИ ДЛЯ ВСТАВКИ ===
                st.markdown("### 📍 Размещение вставки в фасаде")
                st.caption("Укажите в какой ячейке разместить вставку (дверь/окно)")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    insert_col = st.number_input(
                        "Столбец (от 1)",
                        min_value=1,
                        max_value=pos.get("columns", 3),
                        value=pos.get("insert_col", 1),
                        step=1,
                        key=f"insert_col_{idx}",
                        help="Номер столбца слева направо"
                    )
                
                with col2:
                    insert_row = st.number_input(
                        "Ряд (от 1)",
                        min_value=1,
                        max_value=pos.get("rows", 2),
                        value=pos.get("insert_row", 1),
                        step=1,
                        key=f"insert_row_{idx}",
                        help="Номер ряда сверху вниз"
                    )
                
                # Визуализация
                st.info(f"🎯 Вставка будет размещена: **Столбец {insert_col}, Ряд {insert_row}**")
                
                # Сохраняем
                pos["insert_col"] = insert_col
                pos["insert_row"] = insert_row
                
                st.markdown("---")
                st.info(f"🔧 Настройка вставки ({fill_type})")
                
                # ТИП ИЗДЕЛИЯ (ДОБАВЛЕНО)
                if fill_type == "Дверь":
                    product_types = ["Дверь 1 створч.", "Дверь 2-х створч."]
                    saved_product = pos.get("insert_product_type", "Дверь 2-х створч.")
                    try:
                        product_index = product_types.index(saved_product)
                    except ValueError:
                        product_index = 1
                    
                    product_type = st.selectbox(
                        "Тип изделия",
                        product_types,
                        index=product_index,
                        key=f"fac_product_{idx}"
                    )
                    pos["insert_product_type"] = product_type
                    sash_count = 1 if "1 створч" in product_type else 2
                else:
                    product_type = "Окно с откр."
                    pos["insert_product_type"] = product_type
                    sash_count = 2
                
                # Система профиля для вставки
                # Получаем сохраненную систему или используем дефолтную
                saved_system = pos.get("insert_system", "ALG 2030-73C")
                system_options = ["ALG 2030-73C", "ALG 2030-63C", "ALG 2030-55C", "ALG 2030-45C"]
                
                try:
                    system_index = system_options.index(saved_system)
                except ValueError:
                    system_index = 0  # Дефолт на 73C если не найдено
                
                insert_system = st.selectbox(
                    "Система профиля вставки",
                    system_options,
                    index=system_index,  # ← Используем сохраненный индекс
                    key=f"fac_ins_sys_{idx}"
                )
                
                # Вызываем форму окна/двери для вставки
                with st.container():
                    st.caption(f"⚠️ Максимальные размеры вставки: {cell_w_mm:.0f} × {cell_h_mm:.0f} мм")
                    
                    # Получаем ранее сохраненные данные (если есть)
                    initial_insert_data = pos.get("insert_data", None)
                    
                    # Вызываем форму с начальными данными
                    insert_data = window_door_ui(
                        f"fac_insert_{idx}", 
                        idx, 
                        insert_system,
                        initial_data=initial_insert_data  # ← Передаем сохраненные данные
                    )
                    
                    # ИСПРАВЛЕНО: Добавляем тип изделия и количество створок
                    insert_data["product_type"] = product_type
                    insert_data["sash_count"] = sash_count
                    
                    # Заполнение УЖЕ установлено в window_door_ui (форма спросила)
                    
                    # Сохраняем данные вставки
                    pos["insert_data"] = insert_data
                    pos["insert_system"] = insert_system
    
    # ========== КОНЕЦ РАЗДЕЛЕНИЯ ==========
    
    # === КНОПКИ УПРАВЛЕНИЯ ===
    st.markdown("---")
    
    col_calc, col_clear = st.columns([3, 1])
    
    # Определяем какие позиции используются
    positions_list = st.session_state.get("facade_positions", [])
    
    print(f"DEBUG: facade_type_value = {facade_type_value}")
    print(f"DEBUG: positions count = {len(positions_list)}")
    
    with col_calc:
        calc_button = st.button(
            "🚀 РАССЧИТАТЬ ФАСАД","🚀 РАССЧИТАТЬ ФАСАД",
            type="primary",
            width="stretch"
        )
    
    with col_clear:
        if st.button("🗑️ Очистить", type="secondary", width="stretch"):
            st.session_state.facade_positions = []
            if 'last_facade_result' in st.session_state:
                del st.session_state.last_facade_result
            st.rerun()
    
    if calc_button:
        if not positions_list:
            st.error("❌ Добавьте хотя бы одну позицию!")
            # ========== РАСЧЁТ ТАМБУРА ==========
            try:
                # Импорт уже в начале файла
                
                print("\n🏗️ Вызов calculate_tambour_materials_v2:")
                print(f"   Позиций: {len(st.session_state.tambour_positions)}")
                
                tambour_calc = calculate_tambour_materials_v2(
                    positions=st.session_state.tambour_positions,
                    ref1=ref1,
                    ref2=ref2,
                    ref3=ref3
                )
                
                materials_cost = tambour_calc.get("total_cost", 0)
                
                # ✅ ИСПРАВЛЕНО: Берём метрики ИЗ РЕЗУЛЬТАТА (как в окнах/дверях)
                total_area = tambour_calc["metrics"]["total_area"]
                total_perimeter = tambour_calc["metrics"]["total_perimeter"]
                
                # ✅ ДОБАВЛЕНО: Сохраняем для экспорта (будет создан ниже)
                # tambour_order_data будет создан в блоке истории
                
                # Тонировка, Сборка, Монтаж
                toning_cost = 0
                if facade_toning == "Есть":
                    toning_cost = total_area * ref2.get("тонировка", 0)
                
                assembly_cost = 0
                if facade_assembly == "Есть":
                    assembly_cost = total_area * ref2.get("сборка", 0)
                
                installation_cost = 0
                if facade_installation != "Нет":
                    install_key = facade_installation.lower().replace(" / ", "/")
                    installation_cost = total_area * ref2.get(install_key, 0)
                
                # Дополнительные детали
                additional_cost = 0
                if facade_additional != "Нет":
                    price_additional = ref2.get(facade_additional.lower(), 0)
                    additional_cost = math.ceil(total_perimeter / 3) * price_additional
                
                # Сумма без обеспечения
                subtotal = materials_cost + toning_cost + assembly_cost + installation_cost + additional_cost
                
                # Обеспечение 81%
                margin = subtotal * 0.81
                total_cost = subtotal + margin
                
                # Сохраняем результат
                st.session_state.last_facade_result = {
                    "facade_type": "Оконный тамбур",
                    "order_number": facade_order_num,
                    "metrics": {
                        "total_area": total_area,
                        "total_perimeter": total_perimeter
                    },
                    "tambour_calc": tambour_calc,
                    "part3_final": {
                        "Тонировка": round(toning_cost, 0),
                        "Сборка": round(assembly_cost, 0),
                        "Монтаж": round(installation_cost, 0),
                        "Дополнительные детали": round(additional_cost, 0),
                        "Материалы": round(materials_cost, 0),
                        "Обеспечение": round(margin, 0)
                    },
                    "total_cost": total_cost
                }
                
                # Сохранение истории
                try:
                    current_user = st.session_state.get("current_user", {})
                    user_login = current_user.get("login", "unknown")
                    
                    tambour_order_data = {
                        "common": {
                            "order_number": facade_order_num,
                            "facade_type": "Оконный тамбур"
                        },
                        "positions": st.session_state.tambour_positions
                    }
                    
                    # ✅ ДОБАВЛЕНО: Сохраняем для экспорта (как в окнах/дверях)
                    st.session_state.last_tambour_order_data = tambour_order_data
                    st.session_state.last_tambour_result = tambour_calc
                    
                    save_history(
                        GOOGLE_CREDENTIALS_PATH,
                        SPREADSHEET_ID,
                        user_login,
                        tambour_order_data,
                        st.session_state.last_facade_result
                    )
                except Exception as e:
                    print(f"⚠️ История тамбура не сохранена: {e}")
                
                # Вывод результатов
                st.success("✅ Расчет тамбура выполнен!")
                
                # Метрики
                col1, col2, col3 = st.columns(3)
                col1.metric("Общая площадь", f"{total_area:.2f} м²")
                col2.metric("Суммарный периметр", f"{total_perimeter:.2f} м.п.")
                col3.metric("💰 ИТОГО К ОПЛАТЕ", f"{total_cost:,.0f} ₸")
                
                st.markdown("---")
                
                # Детализация изделий
                st.subheader("Изделия тамбура:")
                products_data = []
                for prod in tambour_calc["products"]:
                    products_data.append({
                        "Изделие": prod["name"],
                        "Размер": prod["size"],
                        "Стоимость": f"{prod['cost']:,.0f} ₸"
                    })
                st.dataframe(pd.DataFrame(products_data), width="stretch", hide_index=True)
                st.write(f"**Итого изделия:** {tambour_calc['total_products_cost']:,.0f} ₸")
                
                st.markdown("---")
                
                # Детализация соединительных элементов
                st.subheader("Соединительные элементы:")
                conn_data = []
                for elem, data in tambour_calc["connecting"].items():
                    conn_data.append({
                        "Элемент": elem,
                        "Количество": f"{data['quantity']:.2f} {data['unit']}",
                        "Цена": f"{data['price']:,.0f} ₸",
                        "Стоимость": f"{data['cost']:,.0f} ₸"
                    })
                st.dataframe(pd.DataFrame(conn_data), width="stretch", hide_index=True)
                st.write(f"**Итого соединения:** {tambour_calc['total_connecting_cost']:,.0f} ₸")
                
                st.markdown("---")
                
                # Итоговый расчет
                st.subheader("💰 Итоговый расчет")
                part3_data = []
                for key, value in st.session_state.last_facade_result["part3_final"].items():
                    part3_data.append({"Наименование": key, "Сумма (₸)": f"{value:,.0f}"})
                
                df_part3 = pd.DataFrame(part3_data)
                st.dataframe(df_part3, width="stretch", hide_index=True)
                
                st.metric("🎯 ИТОГО К ОПЛАТЕ", f"{total_cost:,.0f} ₸", help="С учетом обеспечения 81%")
                
            except Exception as e:
                st.error(f"❌ Ошибка при расчете тамбура: {e}")
                with st.expander("🔍 Детали ошибки"):
                    import traceback
                    st.code(traceback.format_exc())
        
        else:
            # ========== РАСЧЁТ ФАСАДА (Ruit 50F) ==========
            try:
                if not st.session_state.facade_positions:
                    st.error("❌ Добавьте хотя бы одну позицию фасада!")
                else:
                    # ✅ ИСПРАВЛЕНО: Площадь и периметр будут взяты из facade_calc
                    # (не пересчитываем вручную)
                    results = []
                
                    for idx, pos in enumerate(st.session_state.facade_positions):
                        # ИЗМЕНЕНО: используем среднюю высоту для трапеции
                        h_left = pos.get("height_left", pos.get("height", 3.0))  # fallback для старых данных
                        h_right = pos.get("height_right", 0.0)
                        h_avg = (h_left + h_right) / 2 if h_right > 0 else h_left
                        area = pos["width"] * h_avg  # ИЗМЕНЕНО: используем h_avg
                    
                        n_cells = pos["columns"] * pos["rows"]
                        
                        fill_name = {
                            "blind": "Глухое остекление",
                            "window": "Окно",
                            "door": "Дверь"
                        }.get(pos.get("filling_type", "blind"), "Неизвестно")
                        
                        # ИЗМЕНЕНО: Показываем форму в габаритах
                        if h_right > 0 and abs(h_left - h_right) > 0.01:
                            gabarity = f"{pos['width']:.2f} × ({h_left:.2f}→{h_right:.2f})"
                        else:
                            gabarity = f"{pos['width']:.2f} × {h_left:.2f}"
                        
                        results.append({
                            "Позиция": idx + 1,
                            "Габариты (м)": gabarity,  # ИЗМЕНЕНО
                            "Площадь (м²)": f"{area:.2f}",
                            "Ячейки": f"{pos['columns']} × {pos['rows']} = {n_cells} шт",
                            "Тип заполнения": fill_name
                        })
                
                # ===== РАСЧЁТ МАТЕРИАЛОВ ФАСАДА =====
                # НОВОЕ: Считаем каждую позицию отдельно!
                
                total_materials_cost = 0
                total_area = 0
                total_perimeter = 0
                total_cost_per_sqm = 0
                all_positions_calcs = []  # Сохраняем результаты каждой позиции
                
                print(f"\n{'='*70}")
                print(f"РАСЧЁТ {len(st.session_state.facade_positions)} ПОЗИЦИЙ ФАСАДА")
                print(f"{'='*70}")
                
                for idx, pos in enumerate(st.session_state.facade_positions, 1):
                    print(f"\n--- ПОЗИЦИЯ {idx} ---")
                    
                    # Габариты ЭТОЙ позиции
                    W = pos.get("width", 6.0)
                    H1 = pos.get("height_left", pos.get("height", 3.5))
                    H2 = pos.get("height_right", 0.0)
                    cols = pos.get("columns", 3)
                    rows = pos.get("rows", 2)
                    
                    # Профили ЭТОЙ позиции
                    mullion_size = pos.get("mullion_size", 130)
                    transom_size = pos.get("transom_size", 85)
                    brackets_per_mullion = pos.get("brackets_per_mullion", 2)
                    
                    # Вставка ЭТОЙ позиции (если есть)
                    inserts_for_this_pos = []
                    insert_materials_cost = 0  # Стоимость материалов вставки
                    insert_calc_details = None  # Детализация вставки для UI
                    filling_type = pos.get("filling_type", "blind")
                    
                    if filling_type in ["window", "door"]:
                        insert_data = pos.get("insert_data", {})
                        
                        # Данные для передачи в calculate_facade_materials (для адаптера рамы)
                        inserts_for_this_pos.append({
                            "type": filling_type,
                            "cell_col": pos.get("insert_col", 1),
                            "cell_row": pos.get("insert_row", 1),
                            "width": insert_data.get("width", 1800) / 1000,
                            "height": insert_data.get("height", 2200) / 1000,
                            "system": pos.get("insert_system", "ALG 2030-63C"),
                            "product_type": insert_data.get("product_type", "Дверь 2-х створч."),
                            "data": {
                                "glass_type": insert_data.get("glass_type", "двойной"),
                                "fill_category": insert_data.get("fill_category", "Стеклопакет"),
                                "lambri_type": insert_data.get("lambri_type", "Ламбри без термо"),
                                "toning": "Нет",
                                "assembly": "Нет",
                                "installation": "Нет",
                                "sash_count": insert_data.get("sash_count", 2)
                            }
                        })
                        
                        print(f"   Вставка: {filling_type} в ячейке ({pos.get('insert_col', 1)}, {pos.get('insert_row', 1)})")
                        
                        # ============================================================================
                        # НОВОЕ: РАСЧЁТ МАТЕРИАЛОВ ВСТАВКИ через calculate_window_smeta()
                        # ============================================================================
                        
                        try:
                            # Формируем order_data для calculate_window_smeta
                            insert_system = pos.get("insert_system", "ALG 2030-63C")
                            product_type = insert_data.get("product_type", "Дверь 2-х створч.")
                            
                            # КРИТИЧНО: Используем правильную функцию для получения CODE
                            code = get_code_for_windows_doors(product_type, insert_system)
                            
                            print(f"   📋 Система: {insert_system}, Тип: {product_type}")
                            print(f"   🔑 CODE: {code}")
                            
                            insert_order_data = {
                                "common": {
                                    "order_number": f"INS-{idx}",
                                    "toning": "Нет",
                                    "assembly": "Нет",
                                    "installation": "Нет"
                                },
                                "positions": [{
                                    "product_type": product_type,
                                    "system_id": insert_system,  # ИСПРАВЛЕНО: system_id вместо system
                                    "code": code,                 # ИСПРАВЛЕНО: правильный CODE
                                    "count": 1,
                                    "data": insert_data  # КРИТИЧНО: передаём ВСЕ данные напрямую!
                                }]
                            }
                            
                            # Вызываем расчёт
                            print(f"   🔧 Расчёт материалов вставки через calculate_window_smeta...")
                            insert_result = calculate_window_smeta(insert_order_data, ref1, ref2, ref3)
                            
                            # КРИТИЧНО: Для вставки в фасад берём ТОЛЬКО профили и фурнитуру
                            # Стеклопакет НЕ включаем - он считается в общем итоге фасада!
                            # Нащельник вставки ТОЖЕ НЕ включаем - он часть общего нащельника фасада!
                            
                            # ИСПРАВЛЕНО: Извлекаем материалы из part2_materials (новый формат от adapter.py)
                            part2_materials = insert_result.get("part2_materials", [])
                            
                            # Разделяем материалы по типам элементов
                            part1_cost = 0  # Профили
                            part2_cost = 0  # Фурнитура + Комплектующие + Уплотнители
                            
                            for material in part2_materials:
                                material_type = material.get("Тип элемента", "")
                                material_sum = material.get("Сумма", 0)
                                
                                if material_type == "Профиль":
                                    part1_cost += material_sum
                                elif material_type in ["Фурнитура", "Комплектующие", "Уплотнитель"]:
                                    part2_cost += material_sum
                            
                            # ИТОГО: ТОЛЬКО Профили + Фурнитура (БЕЗ стеклопакета, БЕЗ нащельника, БЕЗ обеспечения)
                            insert_materials_cost = part1_cost + part2_cost
                            insert_calc_details = insert_result  # Сохраняем для детализации
                            
                            print(f"   💎 Детализация вставки:")
                            print(f"      Профили: {part1_cost:,.0f}₸")
                            print(f"      Фурнитура: {part2_cost:,.0f}₸")
                            print(f"      (Нащельник считается в общем нащельнике фасада)")
                            print(f"      (Стеклопакет считается в общем итоге)")
                            print(f"   ✅ ИТОГО материалы вставки: {insert_materials_cost:,.0f}₸")
                            
                        except Exception as e:
                            print(f"   ⚠️ Ошибка расчёта вставки: {e}")
                            insert_materials_cost = 0
                            insert_calc_details = None
                    
                    # Расчёт ЭТОЙ позиции (count=1!)
                    pos_calc = calculate_facade_materials(
                        W=W,
                        H1=H1,
                        H2=H2,
                        cols=cols,
                        rows=rows,
                        count=1,  # ИСПРАВЛЕНО: каждая позиция считается отдельно!
                        mullion_size=mullion_size,
                        transom_size=transom_size,
                        brackets_per_mullion=brackets_per_mullion,
                        inserts=inserts_for_this_pos,
                        facade_profiles_ref=ref_facade,
                        ref1=ref1,
                        ref2=ref2,
                        ref3=ref3
                    )
                    
                    # Суммируем результаты
                    pos_cost = pos_calc.get("total_cost", 0)
                    pos_cost += insert_materials_cost  # НОВОЕ: Добавляем стоимость вставки!
                    pos_area = pos_calc.get("metrics", {}).get("total_area", 0)
                    pos_perimeter = pos_calc.get("metrics", {}).get("total_perimeter", 0)
                    
                    total_materials_cost += pos_cost
                    total_area += pos_area
                    total_perimeter += pos_perimeter
                    
                    # Сохраняем результаты с деталями вставки
                    pos_calc["insert_materials_cost"] = insert_materials_cost
                    pos_calc["insert_calc_details"] = insert_calc_details
                    all_positions_calcs.append(pos_calc)
                    
                    print(f"   ✅ Позиция {idx}: {pos_area:.2f}м², {pos_perimeter:.2f}м, {pos_cost:,.0f}₸")
                    if insert_materials_cost > 0:
                        print(f"      (в т.ч. вставка: {insert_materials_cost:,.0f}₸)")
                
                # Средняя стоимость за м²
                if total_area > 0:
                    total_cost_per_sqm = total_materials_cost / total_area
                
                print(f"\n{'='*70}")
                print(f"ИТОГО ПО ВСЕМ ПОЗИЦИЯМ:")
                print(f"   Площадь: {total_area:.2f} м²")
                print(f"   Периметр: {total_perimeter:.2f} м")
                print(f"   Стоимость: {total_materials_cost:,.0f}₸")
                print(f"   Средняя ₸/м²: {total_cost_per_sqm:,.0f}₸/м²")
                print(f"{'='*70}")
                
                materials_cost = total_materials_cost
                
                # Создаём объединённый facade_calc для совместимости со старым кодом
                # (содержит агрегированные данные по ВСЕМ позициям)
                facade_calc_combined = {
                    "total_cost": total_materials_cost,
                    "metrics": {
                        "total_area": total_area,
                        "total_perimeter": total_perimeter,
                        "cost_per_sqm": total_cost_per_sqm
                    },
                    "all_positions": all_positions_calcs,  # Массив результатов каждой позиции
                    "positions_count": len(all_positions_calcs)
                }
                
                # Для детализации берём первую позицию (если есть)
                facade_calc_saved = all_positions_calcs[0] if all_positions_calcs else None
                
                # Стеклопакеты/ламбри - собираем данные ПО ТИПАМ
                glass_areas = {"Двойной": 0, "Тройной": 0, "Энергодвойной": 0}
                lambri_areas = {}  # ИСПРАВЛЕНО: по типам
                
                for pos in st.session_state.facade_positions:
                    # ИЗМЕНЕНО: используем среднюю высоту
                    h_left = pos.get("height_left", pos.get("height", 3.0))  # fallback для старых данных
                    h_right = pos.get("height_right", 0.0)
                    h_avg = (h_left + h_right) / 2 if h_right > 0 else h_left
                    area = pos["width"] * h_avg  # ИЗМЕНЕНО
                    
                    filling_type = pos.get("filling_type", "blind")
                    
                    # ГЛУХОЕ ОСТЕКЛЕНИЕ
                    if filling_type == "blind":
                        blind_data = pos.get("blind_data", {})
                        panel_type = blind_data.get("panel_type", "glass")
                        
                        if panel_type == "glass":
                            # Собираем площадь по типам стеклопакетов
                            glass_type = blind_data.get("glass_type", "Двойной")
                            glass_areas[glass_type] = glass_areas.get(glass_type, 0) + area
                        else:
                            # ИСПРАВЛЕНО: Собираем ламбри ПО ТИПАМ
                            lambri_type = panel_type
                            lambri_areas[lambri_type] = lambri_areas.get(lambri_type, 0) + area
                    
                    # ВСТАВКИ (ОКНА/ДВЕРИ) - собираем площадь стеклопакетов И ЛАМБРИ
                    elif filling_type in ["window", "door"]:
                        # ✅ ИСПРАВЛЕНО: Сначала считаем ОСНОВНОЙ ФАСАД (глухие ячейки)
                        blind_data = pos.get("blind_data", {})
                        panel_type = blind_data.get("panel_type", "glass")
                        
                        # Площадь ВСЕГО ФАСАДА (используем h_avg)
                        total_facade_area = area  # ИЗМЕНЕНО: уже посчитано выше
                        
                        # Площадь ВСТАВКИ
                        insert_data = pos.get("insert_data", {})
                        insert_w = insert_data.get("width", 1800) / 1000
                        insert_h = insert_data.get("height", 2200) / 1000
                        insert_area = insert_w * insert_h
                        
                        # Площадь ОСНОВНОГО ФАСАДА (всё минус вставка)
                        main_facade_area = total_facade_area - insert_area
                        
                        # ✅ СЧИТАЕМ ОСНОВНОЙ ФАСАД (глухие ячейки):
                        if panel_type == "glass":
                            # Основной фасад - стеклопакет
                            glass_type = blind_data.get("glass_type", "Двойной")
                            glass_areas[glass_type] = glass_areas.get(glass_type, 0) + main_facade_area
                        else:
                            # Основной фасад - ламбри
                            lambri_type = panel_type
                            lambri_areas[lambri_type] = lambri_areas.get(lambri_type, 0) + main_facade_area
                        
                        # ✅ ДОБАВЛЯЕМ СТЕКЛОПАКЕТ/ЛАМБРИ ВСТАВКИ (окна/двери):
                        # Материалы вставки (профили+фурнитура) посчитаны отдельно через calculate_window_smeta()
                        # Но СТЕКЛОПАКЕТ/ЛАМБРИ вставки нужно добавить в общий расчёт!
                        
                        insert_fill_category = insert_data.get("fill_category", "Стеклопакет")
                        
                        # ПРОВЕРЯЕМ ТИП ЗАПОЛНЕНИЯ ВСТАВКИ:
                        if "ламбри" in insert_fill_category.lower():
                            # ЛАМБРИ ВСТАВКИ
                            lambri_type_insert = insert_fill_category  # "Ламбри без термо" или "Ламбри с термо"
                            lambri_areas[lambri_type_insert] = lambri_areas.get(lambri_type_insert, 0) + insert_area
                            
                            print(f"   📊 Площади заполнения:")
                            print(f"      Основной фасад ({panel_type}): {main_facade_area:.2f}м²")
                            print(f"      Вставка ({lambri_type_insert}): {insert_area:.2f}м²")
                        else:
                            # СТЕКЛОПАКЕТ ВСТАВКИ
                            insert_glass_type = insert_data.get("glass_type", "Двойной")
                            
                            # КРИТИЧНО: Проверяем что выбран стеклопакет, а не "Нет"
                            if insert_glass_type and insert_glass_type.lower() != "нет":
                                insert_glass_type_normalized = insert_glass_type.capitalize()  # Нормализуем
                                glass_areas[insert_glass_type_normalized] = glass_areas.get(insert_glass_type_normalized, 0) + insert_area
                                
                                print(f"   📊 Площади стеклопакетов:")
                                print(f"      Основной фасад ({glass_type if panel_type == 'glass' else 'нет'}): {main_facade_area:.2f}м²")
                                print(f"      Вставка ({insert_glass_type_normalized}): {insert_area:.2f}м²")
                            else:
                                print(f"   📊 Площади стеклопакетов:")
                                print(f"      Основной фасад ({glass_type if panel_type == 'glass' else 'нет'}): {main_facade_area:.2f}м²")
                                print(f"      Вставка: БЕЗ стеклопакета (выбрано: {insert_glass_type})")
                
                # РАСЧЕТ СТЕКЛОПАКЕТОВ (по общей площади каждого типа)
                glass_cost = 0
                print(f"\n📊 РАСЧЁТ СТЕКЛОПАКЕТОВ:")
                for glass_type, total_glass_area in glass_areas.items():
                    if total_glass_area > 0:
                        # ИСПРАВЛЕНО: берём из ref2 с нормализацией регистра
                        price_per_m2 = ref2.get(glass_type.lower(), 9000)
                        cost = total_glass_area * price_per_m2
                        glass_cost += cost
                        print(f"   {glass_type}: {total_glass_area:.2f}м² × {price_per_m2:,.0f}₸/м² = {cost:,.0f}₸")
                print(f"   ИТОГО стеклопакет: {glass_cost:,.0f}₸")
                
                # РАСЧЕТ ЛАМБРИ (по общей площади КАЖДОГО ТИПА)
                lambri_cost = 0
                if any(lambri_areas.values()):
                    print(f"\n📊 РАСЧЁТ ЛАМБРИ:")
                for lambri_type, lambri_area in lambri_areas.items():
                    if lambri_area > 0:
                        # Кол-во к отгрузке = ceil(площадь / 6)
                        q_otgr = math.ceil(lambri_area / 6.0)
                        
                        # ИСПРАВЛЕНО: берём из ref2 по типу, нормализация регистра
                        price_per_m_lambri = ref2.get(lambri_type.lower(), 2248)
                        
                        # Сумма = цена_за_метр * (кол-во_хлыстов * 6м)
                        cost = price_per_m_lambri * (q_otgr * 6)
                        lambri_cost += cost
                        print(f"   {lambri_type}: {lambri_area:.2f}м² → {q_otgr} хлыстов × {price_per_m_lambri:,.0f}₸/м = {cost:,.0f}₸")
                if lambri_cost > 0:
                    print(f"   ИТОГО ламбри: {lambri_cost:,.0f}₸")
                
                # Тонировка
                toning_cost = 0
                if facade_toning == "Есть":
                    price_toning = ref2.get("тонировка", 2000)
                    toning_cost = total_area * price_toning
                
                # Сборка
                assembly_cost = 0
                if facade_assembly == "Есть":
                    price_assembly = ref2.get("сборка", 10000)
                    assembly_cost = total_area * price_assembly
                
                # Монтаж - ИСПРАВЛЕНО: берём из ref2 с нормализацией
                installation_cost = 0
                if facade_installation != "Нет":
                    # Нормализуем: убираем пробелы вокруг / и приводим к нижнему регистру
                    install_key = facade_installation.lower().replace(" / ", "/")
                    price_installation = ref2.get(install_key, 10000)
                    installation_cost = total_area * price_installation
                
                # ДОБАВЛЕНО: Дополнительные детали (НАЩЕЛЬНИК)
                additional_cost = 0
                additional_length = 0  # Длина нащельника
                
                # Ищем "Нащельник" в ref2
                additional_name = None
                for key in ref2.keys():
                    if "нащельник" in key.lower():
                        additional_name = key
                        break
                
                if additional_name and facade_additional != "Нет":
                    price_additional = ref2.get(additional_name, 0)
                    
                    # ПРАВИЛЬНАЯ ФОРМУЛА НАЩЕЛЬНИКА:
                    # Нащельник ставится по внешнему контуру (слева, справа, сверху)
                    # Формула: ceil((H1 + H2 + W) * count / 3) * цена
                    
                    for pos_calc in all_positions_calcs:
                        # Берём габариты из расчёта позиции
                        pos_metrics = pos_calc.get("metrics", {})
                        pos_h1 = pos_metrics.get("H1", 0)
                        pos_h2 = pos_metrics.get("H2", 0)
                        pos_w = pos_metrics.get("W", 0)
                        
                        # L = (H1 + H2 + W) * count
                        # count = 1 для каждой позиции (каждая позиция считается отдельно)
                        pos_count = 1  # У нас каждая позиция = 1 фасад
                        pos_length = (pos_h1 + pos_h2 + pos_w) * pos_count
                        additional_length += pos_length
                    
                    # Округление вверх: ceil(длина / 3)
                    # Формула: ceil((H1+H2+W)*count / 3) * цена
                    additional_qty = math.ceil(additional_length / 3)
                    additional_cost = additional_qty * price_additional
                    
                    print(f"\n💎 Нащельник:")
                    print(f"   Формула: ceil((H1+H2+W)*count / 3) × цена")
                    print(f"   Длина по контуру: {additional_length:.2f}м")
                    print(f"   Упаковок (по 3м): {additional_qty} шт")
                    print(f"   Цена: {price_additional:,.0f}₸ за упаковку")
                    print(f"   Стоимость: {additional_cost:,.0f}₸")
                
                # Сумма без обеспечения
                subtotal = materials_cost + glass_cost + lambri_cost + toning_cost + assembly_cost + installation_cost + additional_cost
                
                print(f"\n{'='*70}")
                print(f"📊 ИТОГОВЫЙ РАСЧЁТ ФАСАДА:")
                print(f"   Материалы каркаса: {materials_cost:,.0f}₸")
                print(f"   Стеклопакет: {glass_cost:,.0f}₸")
                print(f"   Ламбри: {lambri_cost:,.0f}₸")
                print(f"   Тонировка: {toning_cost:,.0f}₸")
                print(f"   Сборка: {assembly_cost:,.0f}₸")
                print(f"   Монтаж: {installation_cost:,.0f}₸")
                print(f"   Дополнительные детали: {additional_cost:,.0f}₸")
                print(f"   {'─'*68}")
                print(f"   СУММА БЕЗ ОБЕСПЕЧЕНИЯ: {subtotal:,.0f}₸")
                
                # Обеспечение 81% (было 65%)
                margin = subtotal * 0.81
                total_cost = subtotal + margin
                
                print(f"   Обеспечение (81%): {margin:,.0f}₸")
                print(f"{'='*70}")
                print(f"К ОПЛАТЕ: {total_cost:,.0f}₸")
                print(f"{'='*70}")
                
                # ИСПРАВЛЕНО: Сохраняем результат - стеклопакет ПО ТИПАМ
                st.session_state.last_facade_result = {
                    "facade_type": "Фасад",
                    "order_number": facade_order_num,
                    "metrics": {
                        "total_area": total_area,
                        "total_perimeter": total_perimeter
                    },
                    "part3_final": {}
                }
                
                # === ДОБАВЛЯЕМ В ИТОГИ ОБЯЗАТЕЛЬНО (ДАЖЕ ЕСЛИ 0) ===
                
                # Стеклопакеты - ОБЩАЯ СУММА (БЕЗ РАЗБИВКИ ПО ТИПАМ)
                total_glass_cost_all = 0
                for glass_type, glass_area in glass_areas.items():
                    price_glass = ref2.get(glass_type.lower(), 9000)
                    cost_glass_type = glass_area * price_glass if glass_area > 0 else 0
                    total_glass_cost_all += cost_glass_type
                
                # ОДНА строка для всего стеклопакета
                st.session_state.last_facade_result["part3_final"]["Стеклопакет"] = round(total_glass_cost_all, 0)
                
                # Ламбри - ОБЩАЯ СУММА (БЕЗ РАЗБИВКИ ПО ТИПАМ)
                total_lambri_cost_all = 0
                for lambri_type, lambri_area in lambri_areas.items():
                    q_otgr = math.ceil(lambri_area / 6.0) if lambri_area > 0 else 0
                    price_lambri = ref2.get(lambri_type.lower(), 2248)
                    cost_lambri_type = price_lambri * (q_otgr * 6) if lambri_area > 0 else 0
                    total_lambri_cost_all += cost_lambri_type
                
                # ОДНА строка для всего ламбри
                st.session_state.last_facade_result["part3_final"]["Ламбри"] = round(total_lambri_cost_all, 0)
                
                # Остальные статьи
                st.session_state.last_facade_result["part3_final"].update({
                    "Тонировка": round(toning_cost, 0),
                    "Сборка": round(assembly_cost, 0),
                    "Монтаж": round(installation_cost, 0),
                    "Дополнительные детали": round(additional_cost, 0),
                    "Материалы": round(materials_cost, 0),
                    "Обеспечение": round(margin, 0)
                })
                
                st.session_state.last_facade_result.update({
                    "total_cost": round(total_cost, 0),
                    "materials_cost": round(materials_cost, 0),
                    "positions": results,
                    "facade_calc": facade_calc_combined  # ИСПРАВЛЕНО: Сохраняем объединённый результат
                })
                
                # ИСПРАВЛЕНО: Сохранение в историю
                try:
                    current_user = st.session_state.get("current_user", {})
                    user_login = current_user.get("login", "unknown")
                    
                    facade_order_data = {
                        "common": {
                            "order_number": facade_order_num,
                            "facade_type": "Фасад"
                        },
                        "positions": st.session_state.facade_positions  # ИСПРАВЛЕНО: передаём реальные позиции
                    }
                    
                    save_history(
                        GOOGLE_CREDENTIALS_PATH,
                        SPREADSHEET_ID,
                        user_login,
                        facade_order_data,
                        st.session_state.last_facade_result
                    )
                except Exception as e:
                    print(f"⚠️ История фасада не сохранена: {e}")
                
                # Вывод результатов
                st.success("✅ Расчет выполнен!")
                
                # Метрики
                col1, col2, col3 = st.columns(3)
                col1.metric("Общая площадь", f"{total_area:.2f} м²")
                col2.metric("Суммарный периметр", f"{total_perimeter:.2f} м.п.")
                col3.metric("К оплате", f"{total_cost:,.0f} ₸")
                
                st.markdown("---")
                
                # Таблица позиций
                st.subheader("Детализация по позициям")
                df = pd.DataFrame(results)
                st.dataframe(df, width="stretch", hide_index=True)
                
                # Итоговый расчет
                st.markdown("---")
                st.subheader("💰 ЧАСТЬ 3: Итоговый расчет")
                
                # ИСПРАВЛЕНО: Показываем отдельно материалы каркаса и вставок
                if "ALG" in facade_type_value or facade_type_value == "Оконный тамбур (ALG)":
                    # Для тамбура показываем детализацию
                    st.write("**Материалы тамбура (ALG):**")
                    if 'tambour_calc' in locals() and tambour_calc.get("skeleton"):
                        tambour_data = []
                        for elem, data in tambour_calc["skeleton"].items():
                            tambour_data.append({
                                "Элемент": elem,
                                "Количество": f"{data['quantity']:.2f} {data['unit']}",
                                "Цена": f"{data['price']:,.0f} ₸",
                                "Стоимость": f"{data['cost']:,.0f} ₸"
                            })
                        st.dataframe(pd.DataFrame(tambour_data), width="stretch", hide_index=True)
                        st.write(f"**Итого материалы:** {tambour_calc.get('total_cost', 0):,.0f} ₸")
                    st.markdown("---")
                else:
                    # Для Ruit 50F показываем детализацию
                    facade_calc = st.session_state.get("last_facade_result", {}).get("facade_calc")
                    
                    if facade_calc and facade_calc.get("all_positions"):
                        # НОВОЕ: Показываем детализацию ПО КАЖДОЙ ПОЗИЦИИ
                        positions_count = facade_calc.get("positions_count", 0)
                        st.write(f"**Детализация по {positions_count} позициям:**")
                        
                        for idx, pos_calc in enumerate(facade_calc["all_positions"], 1):
                            with st.expander(f"📋 Позиция {idx} - детали"):
                                # Каркас этой позиции
                                if pos_calc.get("frame"):
                                    st.write("**Каркас:**")
                                    frame_data = []
                                    for elem, data in pos_calc["frame"].items():
                                        if isinstance(data, dict) and "quantity" in data:
                                            frame_data.append({
                                                "Элемент": elem,
                                                "Количество": f"{data.get('quantity', 0):.2f} {data.get('unit', '')}",
                                                "Цена": f"{data.get('price', 0):,.0f} ₸",
                                                "Стоимость": f"{data.get('cost', 0):,.0f} ₸"
                                            })
                                    if frame_data:
                                        st.dataframe(pd.DataFrame(frame_data), width="stretch", hide_index=True)
                                        st.write(f"**Итого каркас:** {pos_calc['frame'].get('total_cost', 0):,.0f} ₸")
                                
                                # Вставки этой позиции (ЭТАП 2: ДЕТАЛИЗАЦИЯ)
                                if pos_calc.get("insert_calc_details"):
                                    st.write("**Материалы вставки (окно/дверь):**")
                                    
                                    insert_details = pos_calc["insert_calc_details"]
                                    
                                    # Профили вставки
                                    if insert_details.get("part1_final"):
                                        st.write("*Профили:*")
                                        profile_data = []
                                        for name, cost in insert_details["part1_final"].items():
                                            if cost > 0:
                                                profile_data.append({
                                                    "Элемент": name,
                                                    "Стоимость": f"{cost:,.0f} ₸"
                                                })
                                        if profile_data:
                                            st.dataframe(pd.DataFrame(profile_data), width="stretch", hide_index=True)
                                    
                                    # Фурнитура вставки
                                    if insert_details.get("part2_final"):
                                        st.write("*Фурнитура:*")
                                        furn_data = []
                                        for name, cost in insert_details["part2_final"].items():
                                            if cost > 0:
                                                furn_data.append({
                                                    "Элемент": name,
                                                    "Стоимость": f"{cost:,.0f} ₸"
                                                })
                                        if furn_data:
                                            st.dataframe(pd.DataFrame(furn_data), width="stretch", hide_index=True)
                                    
                                    # Доп.детали вставки
                                    if insert_details.get("part3_final"):
                                        part3 = insert_details["part3_final"]
                                        if "Дополнительные детали" in part3 and part3["Дополнительные детали"] > 0:
                                            st.write("*Дополнительные детали:*")
                                            st.write(f"- Нащельник: {part3['Дополнительные детали']:,.0f} ₸")
                                    
                                    # Итого вставка
                                    insert_cost = pos_calc.get("insert_materials_cost", 0)
                                    if insert_cost > 0:
                                        st.write(f"**Итого вставка:** {insert_cost:,.0f} ₸")
                                        st.caption("(Стеклопакет вставки считается в общем итоге)")
                                
                                # Адаптер рамы
                                if pos_calc.get("inserts") and pos_calc["inserts"].get("adapter_frames"):
                                    st.write("**Адаптер рамы:**")
                                    adapter = pos_calc["inserts"]["adapter_frames"]
                                    st.write(f"- Количество: {adapter.get('quantity', 0):.2f} {adapter.get('unit', 'м')}")
                                    st.write(f"- Стоимость: {adapter.get('cost', 0):,.0f} ₸")
                        
                        st.markdown("---")
                    else:
                        st.info("ℹ️ Детализация по позициям недоступна")
                    
                    st.markdown("---")
                
                part3_data = []
                for key, value in st.session_state.last_facade_result["part3_final"].items():
                    part3_data.append({"Наименование": key, "Сумма (₸)": f"{value:,.0f}"})
                
                df_part3 = pd.DataFrame(part3_data)
                st.dataframe(df_part3, width="stretch", hide_index=True)
                
                st.metric("🎯 ИТОГО К ОПЛАТЕ", f"{total_cost:,.0f} ₸", help="С учетом обеспечения 81%")
                
            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {e}")
                with st.expander("🔍 Детали ошибки"):
                    import traceback
                    st.code(traceback.format_exc())
    
                # Таблица результатов
                st.subheader("Детализация по позициям")
                df = pd.DataFrame(results)
                st.dataframe(df, width="stretch", hide_index=True)
                
                # Предупреждение
                st.warning("⚠️ Это упрощенный расчет. Для точной стоимости необходимо добавить расчет профилей, стеклопакетов и фурнитуры из справочников.")
                
            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {e}")
                st.exception(e)
    
    # === КНОПКА ЭКСПОРТА ===
    if 'last_facade_result' in st.session_state:
        st.divider()
        if st.button("📥 Скачать КП фасада в Excel", type="secondary", width="stretch"):
            try:
                # Импорт уже в начале файла
                
                temp_dir = tempfile.gettempdir()
                order_num = f"FAC-{datetime.datetime.now().strftime('%Y%m%d%H%M')}"
                
                excel_file = export_facade_to_excel(
                    st.session_state.last_facade_result,
                    order_number=order_num,
                    output_dir=temp_dir
                )
                
                with open(excel_file, 'rb') as f:
                    st.download_button(
                        label="⬇️ Загрузить файл",
                        data=f,
                        file_name=f"KP_FACADE_{order_num}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("✅ Файл готов к скачиванию!")
            except Exception as e:
                st.error(f"Ошибка при создании файла: {e}")
                st.exception(e)


# ========================================
# ФУНКЦИЯ ДЛЯ СТРАНИЦЫ ИСТОРИИ (ЗАГЛУШКА)
# ========================================
def render_tambour_page():
    """Оконный тамбур - как Окна/Двери + направляющий"""
    st.header("🪟 Оконный тамбур")
    st.info("Тамбур = готовые двери/окна + направляющий профиль")
    
    # === ОБЩИЕ ДАННЫЕ ===
    st.subheader("📋 Общие данные заказа")
    col1, col2, col3, col4 = st.columns(4)
    
    order_number = col1.text_input("Номер заказа", value="", key="tambour_order_number")
    
    toning = col2.selectbox("Тонировка", ["Нет", "Есть"], key="tambour_toning")
    assembly = col3.selectbox("Сборка", ["Нет", "Есть"], key="tambour_assembly")
    
    installation_options = ["Нет", "Монтаж", "Демонтаж/Монтаж", "Сложный монтаж"]
    installation = col4.selectbox("Монтаж", installation_options, key="tambour_installation")
    
    # Дополнительные детали
    additional_options = ["Нет"] + [k.capitalize() for k in ref2.keys() if "нащельник" in k.lower()]
    additional = st.selectbox("Дополнительные детали", additional_options, key="tambour_additional")
    
    st.markdown("---")
    
    # === ПОЗИЦИИ ===
    if "tambour_positions" not in st.session_state:
        st.session_state.tambour_positions = []
    
    st.subheader(f"📦 Позиции тамбура ({len(st.session_state.tambour_positions)})")
    
    col_add, col_clear, col_new = st.columns(3)
    
    if col_add.button("➕ Добавить позицию", width="stretch", key="tambour_add_btn"):
        st.session_state.tambour_positions.append({
            "product_type": "Дверь 2-х створч.",
            "system_id": "ALG 2030-63C",
            "width": 1800,
            "height": 2200,
            "count": 1,
            "fill_category": "Стеклопакет",
            "glass_type": "Двойной",
            "opening_type": "Откр.",
            "horizontal_imposts": 0,
            "vertical_imposts": 0
        })
        st.rerun()
    
    if col_clear.button("🗑️ Очистить всё", width="stretch", key="tambour_clear_btn"):
        st.session_state.tambour_positions = []
        st.rerun()
    
    # Формы позиций - ИСПОЛЬЗУЕМ window_door_ui
    for idx, pos in enumerate(st.session_state.tambour_positions):
        with st.expander(f"📦 Позиция {idx+1}: {pos.get('product_type', 'Изделие')}", expanded=True):
            if st.button(f"🗑️ Удалить позицию", key=f"tambour_del_{idx}"):
                st.session_state.tambour_positions.pop(idx)
                st.rerun()
            
            # Тип изделия и система
            col_type, col_sys = st.columns(2)
            
            product_type = col_type.selectbox(
                "Тип изделия",
                ["Дверь 1 створч.", "Дверь 2-х створч.", "Окно с откр.", "Окно глухое"],
                index=["Дверь 1 створч.", "Дверь 2-х створч.", "Окно с откр.", "Окно глухое"].index(pos.get("product_type", "Дверь 2-х створч.")),
                key=f"tambour_type_{idx}"
            )
            pos["product_type"] = product_type
            
            system_id = col_sys.selectbox(
                "Система профиля",
                ["ALG 2030-73C", "ALG 2030-63C", "ALG 2030-55C", "ALG 2030-45C"],
                index=["ALG 2030-73C", "ALG 2030-63C", "ALG 2030-55C", "ALG 2030-45C"].index(pos.get("system_id", "ALG 2030-63C")),
                key=f"tambour_sys_{idx}"
            )
            pos["system_id"] = system_id
            
            # Используем window_door_ui для остальных параметров
            initial_data = {
                "width": pos.get("width", 1800),
                "height": pos.get("height", 2200),
                "fill_category": pos.get("fill_category", "Стеклопакет"),
                "glass_type": pos.get("glass_type", "Двойной"),
                "opening_type": pos.get("opening_type", "Откр."),
                "horizontal_imposts": pos.get("horizontal_imposts", 0),
                "vertical_imposts": pos.get("vertical_imposts", 0)
            }
            
            # Вызываем форму
            form_data = window_door_ui(f"tambour_form_{idx}", idx, system_id, initial_data)
            
            # Обновляем позицию
            pos.update(form_data)
            pos["count"] = 1  # Всегда 1 для тамбура
    
    # === КНОПКА РАСЧЁТА ===
    st.markdown("---")
    
    if st.button("🚀 РАССЧИТАТЬ ТАМБУР", type="primary", width="stretch", key="tambour_calc_btn"):
        if not st.session_state.tambour_positions:
            st.error("❌ Добавьте хотя бы одну позицию!")
        else:
            try:
                # Импорты уже в начале файла
                
                # Формируем order_data КАК В ОКНАХ
                order_data = {
                    "common": {
                        "order_number": order_number,
                        "toning": toning,
                        "assembly": assembly,
                        "installation": installation
                    },
                    "positions": []
                }
                
                # Добавляем CODE к позициям И ОБОРАЧИВАЕМ В "data"
                for pos in st.session_state.tambour_positions:
                    # ВАЖНО: Создаём КОПИЮ чтобы не менять session_state!
                    pos_copy = pos.copy()
                    pos_copy["code"] = get_code_for_windows_doors(pos["product_type"], pos["system_id"])
                    
                    # ✅ ИСПРАВЛЕНО: engine_windows ожидает данные В "data"!
                    pos_copy["data"] = {
                        "width": pos["width"],  # В ММ, engine_windows конвертирует
                        "height": pos["height"],
                        "count": pos.get("count", 1),
                        "fill_category": pos.get("fill_category", "Стеклопакет"),
                        "glass_type": pos.get("glass_type", "Двойной"),
                        "product_type": pos["product_type"],
                        "imposts": {
                            "auto_calculate": True,
                            "has_left": False,
                            "has_center": False,
                            "has_right": False,
                            "has_tor": False
                        },
                        "sashes": []
                    }
                    
                    order_data["positions"].append(pos_copy)
                
                # РАСЧЁТ ИЗДЕЛИЙ (как в окнах)
                result = calculate_window_smeta(order_data, ref1, ref2, ref3)
                
                # === ДОБАВЛЯЕМ НАПРАВЛЯЮЩИЙ ПО ФОРМУЛЕ ИЗ СПРАВОЧНИКА-1 ===
                # ПРАВИЛО: Количество стыков = max(1, количество позиций - 1)
                count_joints = max(1, len(st.session_state.tambour_positions) - 1)
                
                # Берём максимальную высоту из позиций
                max_height = max(pos["height"] for pos in st.session_state.tambour_positions) / 1000  # в метры
                
                # ФОРМУЛА ИЗ СПРАВОЧНИКА-1: H × 2 × count_joints
                L_guide_raw = max_height * 2 * count_joints
                
                # Округление до хлыстов 6м (как весь алюминий)
                sticks_guide = math.ceil(L_guide_raw / 6.0)
                L_guide = sticks_guide * 6.0
                
                print(f"\n🔧 DEBUG Направляющий:")
                print(f"  Позиций в тамбуре: {len(st.session_state.tambour_positions)}")
                print(f"  Стыков (по правилу): {count_joints}")
                print(f"  Высота: {max_height:.2f}м")
                print(f"  Формула: {max_height:.2f} × 2 × {count_joints} = {L_guide_raw:.2f}м")
                print(f"  Округление: ⌈{L_guide_raw:.2f}/6⌉ = {sticks_guide} хлыстов × 6м = {L_guide:.2f}м")
                
                # Ищем цену в ref1 (Справочник-1)
                price_guide = 3846  # Запасное из ТЗ
                
                for item in ref1:
                    art = item.get("Артикул", "")
                    if "2-00-5581-60-0000" in art or "2-00-5581" in art:
                        price_guide = item.get("Цена за единицу", 3846)
                        if isinstance(price_guide, str):
                            # Парсим если строка
                            try:
                                price_guide = float(price_guide.replace(" ", "").replace(",", "."))
                            except:
                                price_guide = 3846
                        print(f"DEBUG: Найден направляющий - арт: {art}, цена: {price_guide}₸/м")
                        break
                
                print(f"DEBUG: Используемая цена направляющего = {price_guide}₸/м")
                
                cost_guide = L_guide * price_guide
                
                # Добавляем в МАТЕРИАЛЫ (part2), не в итоги!
                result["part2_materials"].append({
                    "Артикул": "2-00-5581-60-0000",
                    "Наименование": "Направляющий профиль",
                    "Количество": f"{L_guide:.2f} м",
                    "Цена": round(price_guide, 0),
                    "Сумма": round(cost_guide, 0)
                })
                
                # Обновляем "Материалы" в итогах
                old_materials = result["part3_final"].get("Материалы", 0)
                result["part3_final"]["Материалы"] = round(old_materials + cost_guide, 0)
                
                # Пересчитываем обеспечение и итого
                old_margin = result["part3_final"].pop("Обеспечение", 0)
                subtotal = sum(result["part3_final"].values())
                margin = subtotal * 0.81
                result["part3_final"]["Обеспечение"] = round(margin, 0)
                result["total_with_margin"] = round(subtotal + margin, 0)
                
                # ✅ СОХРАНЯЕМ ДЛЯ ЭКСПОРТА
                st.session_state.last_tambour_order_data = order_data
                st.session_state.last_tambour_result = result
                
                # === ВЫВОД РЕЗУЛЬТАТОВ ===
                st.success("✅ Расчет выполнен!")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Общая площадь", f"{result['metrics']['total_area']:.3f} м²")
                col2.metric("Периметр", f"{result['metrics']['total_perimeter']:.3f} м.п.")
                col3.metric("💰 ИТОГО", f"{result['total_with_margin']:,.0f} ₸")
                
                st.markdown("---")
                
                # Таблица позиций
                with st.expander("📊 Позиции", expanded=True):
                    positions_df = []
                    for i, p in enumerate(st.session_state.tambour_positions, 1):
                        positions_df.append({
                            "№": i,
                            "Изделие": p["product_type"],
                            "Система": p["system_id"],
                            "Размер": f"{p['width']}×{p['height']} мм",
                            "Стекло": p["glass_type"]
                        })
                    st.dataframe(pd.DataFrame(positions_df), width="stretch", hide_index=True)
                
                # Материалы
                with st.expander("📦 Материалы (Артикулы)", expanded=True):
                    if result["part2_materials"]:
                        df2 = pd.DataFrame(result["part2_materials"])
                        st.dataframe(df2, width="stretch", hide_index=True)
                
                # Итоговый расчёт
                with st.expander("💰 Итоговый расчет", expanded=True):
                    df3 = pd.DataFrame(result["part3_final"].items(), columns=["Наименование", "Сумма (₸)"])
                    st.dataframe(df3, width="stretch", hide_index=True)
                    st.metric("🎯 ИТОГО К ОПЛАТЕ", f"{result['total_with_margin']:,.0f} ₸")
                
                # ДОБАВЛЕНО: Кнопка скачать КП
                st.markdown("---")
                if st.button("📥 Скачать КП в Excel", type="secondary", width="stretch", key="tambour_export_btn"):
                    try:
                        # ✅ ИСПРАВЛЕНО: Используем session_state (как в окнах/дверях)
                        temp_dir = tempfile.gettempdir()
                        order_num = st.session_state.last_tambour_order_data["common"]["order_number"]
                        excel_path = os.path.join(temp_dir, f"KP_{order_num}.xlsx")
                        
                        export_to_excel(
                            st.session_state.last_tambour_order_data,
                            st.session_state.last_tambour_result,
                            excel_path
                        )
                        
                        with open(excel_path, "rb") as f:
                            st.download_button(
                                label="💾 Сохранить файл",
                                data=f,
                                file_name=f"KP_TAMBOUR_{order_num}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                width="stretch"
                            )
                        st.success(f"✅ Файл KP_TAMBOUR_{order_num}.xlsx готов к скачиванию")
                    except Exception as e:
                        st.error(f"❌ Ошибка экспорта: {e}")
                        import traceback
                        st.code(traceback.format_exc())
                
            except Exception as e:
                st.error(f"❌ Ошибка: {e}")
                import traceback
                st.code(traceback.format_exc())



def render_history_page():
    """Страница истории заказов"""
    
    st.title("📚 История заказов")
    
    try:
        # Загружаем историю из Google Sheets
        from google.oauth2.service_account import Credentials
        import gspread
        
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_PATH, scopes=scopes)
        gc = gspread.authorize(creds)
        
        sh = gc.open_by_key(SPREADSHEET_ID)
        ws = sh.worksheet("ИСТОРИЯ")
        
        # Получаем все данные
        data = ws.get_all_values()
        
        if len(data) <= 1:
            st.info("📭 История пуста")
            return
        
        # Заголовки
        headers = data[0]
        rows = data[1:]
        
        # ИСПРАВЛЕНО: Обработка дублирующихся колонок
        # Добавляем счётчик к дублям
        seen = {}
        unique_headers = []
        for h in headers:
            if h in seen:
                seen[h] += 1
                unique_headers.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 0
                unique_headers.append(h)
        
        # Создаём DataFrame с уникальными заголовками
        import pandas as pd
        df = pd.DataFrame(rows, columns=unique_headers)
        
        # Сортируем по дате (новые сверху)
        if len(df) > 0:
            df = df.iloc[::-1]  # Разворачиваем
            
            st.write(f"**Всего записей:** {len(df)}")
            
            # Фильтр по пользователю
            if "Пользователь" in df.columns:
                users = df["Пользователь"].unique()
                selected_user = st.selectbox("Фильтр по пользователю:", ["Все"] + list(users))
                
                if selected_user != "Все":
                    df = df[df["Пользователь"] == selected_user]
            
            # Показываем таблицу
            st.dataframe(df, width="stretch", hide_index=True)
    
    except Exception as e:
        st.error(f"❌ Ошибка загрузки истории: {e}")
        st.info("💡 Убедитесь что в Google Sheets есть лист 'ИСТОРИЯ' с колонками: Дата, Пользователь, Номер заказа, Позиций, Площадь, Стоимость")


# ========================================
# ГЛАВНОЕ МЕНЮ НАВИГАЦИИ
# ========================================

# Инициализация меню
if 'menu_selection' not in st.session_state:
    st.session_state.menu_selection = "Главная (Окна/Двери)"

# Сайдбар с навигацией
with st.sidebar:
    st.title("📍 Навигация")
    
    menu_selection = st.radio(
        "Выберите раздел:",
        ["Главная (Окна/Двери)", "Фасады", "Оконный тамбур", "История"],
        index=["Главная (Окна/Двери)", "Фасады", "Оконный тамбур", "История"].index(st.session_state.menu_selection) if st.session_state.menu_selection in ["Главная (Окна/Двери)", "Фасады", "Оконный тамбур", "История"] else 0,
        key="sidebar_navigation"
    )
    
    # Сохраняем выбор
    st.session_state.menu_selection = menu_selection

# Роутинг
if st.session_state.menu_selection == "Главная (Окна/Двери)":
    render_windows_doors_page()
elif st.session_state.menu_selection == "Фасады":
    render_facade_page()
elif st.session_state.menu_selection == "Оконный тамбур":
    render_tambour_page()
elif st.session_state.menu_selection == "История":
    render_history_page()
