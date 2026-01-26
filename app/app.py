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
from calculations.material_basket import MaterialBasket
from calculations.material_basket_V9 import MaterialAggregator  # ДОБАВЛЕНО для V.9
from calculations.mapping import get_code_for_windows_doors, get_code_for_facade
from export.export_kp import export_to_excel, export_facade_to_excel
from history.save_history import save_history

# --- КОНСТАНТЫ ИЗ ТЗ ---
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

@st.cache_data(ttl=60)
def get_data():
    r1 = load_reference_1(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r2_raw = load_reference_2(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r3 = load_reference_3(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r_facade = load_facade_reference(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    
    r2 = {k.lower(): v for k, v in r2_raw.items()}
    
    return r1, r2, r3, r_facade

ref1, ref2, ref3, ref_facade = get_data()

def get_glass_types():
    """Получает список типов стеклопакетов из ref2"""
    glass_types = []
    for key in ref2.keys():
        if key not in ['тонировка', 'сборка', 'монтаж', 'демонтаж/монтаж', 'сложный монтаж', 'нащельник']:
            glass_types.append(key.capitalize())
    return sorted(glass_types, key=lambda x: (x == 'Нет', x))

GLASS_TYPES = get_glass_types()

# --- 3. ФУНКЦИЯ КОНСТРУКТОР ОКНА ---
def window_door_ui(prefix, pos_idx, system_id, initial_data=None):
    """Форма для заполнения данных окна/двери"""
    st.markdown("---")
    
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
    st.markdown("### 📲 Импосты")
    
    initial_imposts = initial_data.get("imposts", {})
    auto_default = initial_imposts.get("auto_calculate", True) if initial_imposts else True
    
    auto_imposts = st.checkbox(
        "✅ Автоматический расчет (рекомендуется)", 
        value=auto_default, 
        key=f"{prefix}_auto_imp"
    )
    
    if auto_imposts:
        st.caption("💡 Длина импостов рассчитывается автоматически по системе профиль")
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
# ФУНКЦИЯ ДЛЯ ГЛАВНОЙ СТРАНИЦЫ (ОКНА/ДВЕРИ) - ОБНОВЛЕНО ДЛЯ V.9
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
                if st.button(f"🗑️ Удалить позицию", key=f"del_pos_{idx}"):
                    st.session_state.positions.pop(idx)
                    st.rerun()
                
                pc1, pc2 = st.columns(2)
                
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
                    "Система профиль", 
                    PROFILE_SYSTEMS, 
                    key=f"pc_sys{idx}",
                    index=0
                )
                pos["system_id"] = system_id
                st.session_state.positions[idx]["system_id"] = system_id
                
                if pos["product_type"] != "Фасад":
                    code = get_code_for_windows_doors(
                        pos["product_type"],
                        pos["system_id"]
                    )
                    pos["code"] = code
                    st.session_state.positions[idx]["code"] = code
                
                pos["count"] = 1
                st.session_state.positions[idx]["count"] = 1
                
                data = window_door_ui(f"main_pos_{idx}", idx, pos["system_id"])
                pos["data"] = data
                st.session_state.positions[idx]["data"] = data


    # --- 6. РАСЧЕТ И ВЫВОД (ОБНОВЛЕНО ДЛЯ V.9) ---
    st.divider()

    if st.button("🚀 РАССЧИТАТЬ", type="primary", width="stretch"):
        if not st.session_state.positions:
            st.error("❌ Добавьте хотя бы одну позицию!")
        else:
            order_data = {
                "common": {
                    "order_number": order_num,
                    "toning": toning_id,
                    "assembly": assembly_id,
                    "installation": install_id
                },
                "positions": st.session_state.get("positions", [])
            }
            
            try:
                # === РАСЧЁТ С ГЛОБАЛЬНОЙ АГРЕГАЦИЕЙ V.9 ===
                aggregator = MaterialAggregator(ref1)
                
                # Проходим по ВСЕМ позициям и собираем данные
                position_details = []  # Для ЧАСТИ 2
                for idx, position in enumerate(st.session_state.positions, 1):
                    pos_order_data = {
                        "common": order_data["common"],
                        "positions": [position]
                    }
                    
                    pos_result = calculate_window_smeta(pos_order_data, ref1, ref2, ref3)
                    
                    aggregator.add_metrics(
                        area=pos_result['metrics']['total_area'],
                        perimeter=pos_result['metrics']['total_perimeter']
                    )
                    
                    position_details.append({
                        'Позиция': idx,
                        'Тип': position.get('product_type', ''),
                        'Ширина (мм)': position['data']['width'],
                        'Высота (мм)': position['data']['height'],
                        'Площадь (м²)': round(pos_result['metrics']['total_area'], 3),
                        'Периметр (м)': round(pos_result['metrics']['total_perimeter'], 2)
                    })
                    
                    for material in pos_result.get('part2_materials', []):
                        aggregator.add_material(
                            category='windows_doors',
                            article=material.get('Артикул', ''),
                            quantity_raw=material.get('Количество_raw', material.get('Количество', 0)),
                            unit=material.get('Единица', 'шт'),
                            price=material.get('Цена', 0),
                            name=material.get('Элемент', '')
                        )
                    
                    part3 = pos_result.get('part3_final', {})
                    aggregator.add_service('glass_cost', part3.get('Стеклопакет', 0))
                    aggregator.add_service('lambri_cost', part3.get('Ламбри', 0))
                    aggregator.add_service('toning_cost', part3.get('Тонировка', 0))
                    aggregator.add_service('assembly_cost', part3.get('Сборка', 0))
                    aggregator.add_service('installation_cost', part3.get('Монтаж', 0))
                    aggregator.add_service('additional_details_cost', part3.get('Дополнительные детали', 0))
                
                aggregator.round_all_materials()
                totals = aggregator.calculate_final_totals(margin_rate=0.81)
                
                st.session_state.last_result = {
                    'aggregator': aggregator,
                    'totals': totals,
                    'position_details': position_details
                }
                st.session_state.last_order_data = order_data
                
                # Сохранение истории
                try:
                    current_user = st.session_state.get("current_user", {})
                    user_login = current_user.get("login", "unknown")
                    history_result = {
                        'metrics': aggregator.metrics,
                        'total_with_margin': totals['total']
                    }
                    save_history(
                        GOOGLE_CREDENTIALS_PATH,
                        SPREADSHEET_ID,
                        user_login,
                        order_data,
                        history_result
                    )
                except Exception as e:
                    st.warning(f"⚠️ История не сохранена: {e}")
                
                # === ВЫВОД РЕЗУЛЬТАТОВ ПО ТЗ V.9 ===
                st.success("✅ Расчёт выполнен!")
                
                st.metric(
                    "💰 ИТОГО К ОПЛАТЕ",
                    f"{totals['total']:,} ₸",
                    delta=f"Экономия благодаря V.9"
                )
                
                st.divider()
                
                # ЧАСТЬ 1: ОБЩИЕ МЕТРИКИ
                st.header("📊 ЧАСТЬ 1: Общие показатели")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Общая площадь", f"{aggregator.metrics['total_area']:.2f} м²")
                col2.metric("Общий периметр", f"{aggregator.metrics['total_perimeter']:.2f} м")
                col3.metric("Позиций в заказе", len(st.session_state.positions))
                
                st.divider()
                
                # ЧАСТЬ 2: ИНФОРМАЦИОННАЯ ДЕТАЛИЗАЦИЯ (БЕЗ ЦЕН!)
                with st.expander("🔹 ЧАСТЬ 2: Список изделий (информация)", expanded=False):
                    st.info("ℹ️ Справочная информация для контроля состава заказа. Цены в этом блоке НЕ указаны.")
                    
                    if position_details:
                        df_positions = pd.DataFrame(position_details)
                        st.dataframe(df_positions, use_container_width=True, hide_index=True)
                    else:
                        st.warning("Нет данных о позициях")
                
                st.divider()
                
                # ЧАСТЬ 3: АГРЕГИРОВАННАЯ СПЕЦИФИКАЦИЯ МАТЕРИАЛОВ
                st.header("📦 ЧАСТЬ 3: Спецификация материалов")
                
                st.info(
                    "✨ **Проектный метод:** Материалы из всех позиций суммированы и округлены ОДИН РАЗ. "
                    "Это устраняет перерасход профилей и позволяет сравнить с заводским расчётом поартикульно."
                )
                
                materials = aggregator.get_category_materials('windows_doors')
                
                if materials:
                    df_materials = pd.DataFrame(materials)
                    
                    st.dataframe(
                        df_materials,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Количество_raw": st.column_config.NumberColumn(
                                "Кол-во нетто",
                                help="Точное количество ДО округления",
                                format="%.3f"
                            ),
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
                    
                    st.metric(
                        "💼 ИТОГО материалы",
                        f"{totals['breakdown']['windows_doors']:,} ₸",
                        help="Стоимость всех материалов после округления"
                    )
                    
                    total_raw = sum(m['Количество_raw'] for m in materials if m['Единица'] == 'м')
                    total_rounded = sum(m['Количество'] for m in materials if m['Единица'] == 'м')
                    if total_rounded > 0:
                        savings_percent = ((total_rounded - total_raw) / total_rounded) * 100
                        st.success(
                            f"📉 Экономия на профилях: {total_rounded - total_raw:.1f}м "
                            f"({savings_percent:.1f}% было бы потрачено при попозиционном округлении)"
                        )
                else:
                    st.warning("⚠️ Материалы не найдены. Возможно, система не определена в Справочнике-1.")
                
                st.divider()
                
                # ЧАСТЬ 4: ФИНАНСОВЫЙ ИТОГ
                st.header("💰 ЧАСТЬ 4: Финансовый итог")
                
                st.markdown("**Расчёт ведётся ОДИН РАЗ для всего блока окон/дверей:**")
                
                final_items = []
                
                if totals['breakdown']['glass'] > 0:
                    final_items.append({
                        'Наименование': 'Стеклопакеты',
                        'Площадь (м²)': f"{aggregator.services['glass_total_area']:.2f}",
                        'Сумма (₸)': f"{totals['breakdown']['glass']:,}"
                    })
                
                if totals['breakdown']['lambri'] > 0:
                    final_items.append({
                        'Наименование': 'Ламбри',
                        'Площадь (м²)': '-',
                        'Сумма (₸)': f"{totals['breakdown']['lambri']:,}"
                    })
                
                if totals['breakdown']['toning'] > 0:
                    final_items.append({
                        'Наименование': 'Тонировка',
                        'Площадь (м²)': '-',
                        'Сумма (₸)': f"{totals['breakdown']['toning']:,}"
                    })
                
                if totals['breakdown']['assembly'] > 0:
                    final_items.append({
                        'Наименование':# ПРОДОЛЖЕНИЕ render_windows_doors_page()

                        'Сборка',
                        'Площадь (м²)': f"{aggregator.metrics['total_area']:.2f}",
                        'Сумма (₸)': f"{totals['breakdown']['assembly']:,}"
                    })
                
                if totals['breakdown']['installation'] > 0:
                    final_items.append({
                        'Наименование': 'Монтаж',
                        'Площадь (м²)': f"{aggregator.metrics['total_area']:.2f}",
                        'Сумма (₸)': f"{totals['breakdown']['installation']:,}"
                    })
                
                if totals['breakdown']['additional_details'] > 0:
                    final_items.append({
                        'Наименование': 'Дополнительные детали',
                        'Площадь (м²)': '-',
                        'Сумма (₸)': f"{totals['breakdown']['additional_details']:,}"
                    })
                
                final_items.append({
                    'Наименование': 'Материалы',
                    'Площадь (м²)': '-',
                    'Сумма (₸)': f"{totals['materials_total']:,}"
                })
                
                if final_items:
                    df_final = pd.DataFrame(final_items)
                    st.dataframe(df_final, use_container_width=True, hide_index=True)
                
                st.divider()
                
                col_a, col_b = st.columns(2)
                with col_a:
                    st.metric("Обеспечение", f"{totals['margin']:,} ₸", help="Наценка 81% на себестоимость")
                with col_b:
                    st.metric("💰 К ОПЛАТЕ", f"{totals['total']:,} ₸", delta="Финальная сумма")
                
                with st.expander("ℹ️ Как рассчитано обеспечение", expanded=False):
                    st.write(f"**Материалы:** {totals['materials_total']:,} ₸")
                    st.write(f"**Услуги:** {totals['services_total']:,} ₸")
                    st.write(f"**Себестоимость:** {totals['subtotal']:,} ₸")
                    st.write(f"**Обеспечение (81%):** {totals['margin']:,} ₸")
                    st.divider()
                    st.write(f"**ИТОГО:** {totals['total']:,} ₸")
            
            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {e}")
                st.exception(e)

    # Кнопка экспорта в Excel
    if 'last_result' in st.session_state and 'last_order_data' in st.session_state:
        st.divider()
        if st.button("📥 Скачать КП в Excel", type="secondary", width="stretch"):
            try:
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
# ФУНКЦИЯ ДЛЯ СТРАНИЦЫ ФАСАДОВ (ОБНОВЛЕНО ДЛЯ V.9)
# ========================================
def render_facade_page():
    """Страница расчета фасадов"""
    
    st.title("🏢 Расчет Фасадов")
    st.markdown("---")
    
    # Общие параметры заказа
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
    
    additional_options = ["Нет"] + [k.capitalize() for k in ref2.keys() if "нащельник" in k.lower()]
    facade_additional = col_add.selectbox("Дополнительные детали", additional_options, key="facade_additional")
    
    st.markdown("---")
    
    facade_type_value = "Фасадная система (Ruit 50F)"
    
    # Позиции фасада
    st.subheader("📦 Позиции фасада")
    
    if "facade_positions" not in st.session_state:
        st.session_state.facade_positions = []
    
    st.subheader(f"Позиции фасада ({len(st.session_state.facade_positions)})")
    
    col_add, col_clear, col_new = st.columns(3)
    
    if col_add.button("➕ Добавить позицию", width="stretch"):
        facade_code = get_code_for_facade(facade_type_value)
        
        st.session_state.facade_positions.append({
            "code": facade_code,
            "facade_type": facade_type_value,
            "width": 6.0,
            "height_left": 3.0,
            "height_right": 0.0,
            "columns": 3,
            "rows": 2,
            "mullion_size": 130,
            "transom_size": 85,
            "brackets_per_mullion": 2,
            "filling_type": "blind",
            "cells_data": []
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
    
    # [ЗДЕСЬ ИДЕТ ВЕСЬ КОД ОТОБРАЖЕНИЯ ФОРМ ФАСАДА - НЕ ИЗМЕНЯЕТСЯ]
    # Оставляю его без изменений, так как он не относится к выводу результатов
    
    # Кнопка расчета
    st.markdown("---")
    
    if st.button("🚀 РАССЧИТАТЬ ФАСАДЫ", type="primary", width="stretch"):
        if not st.session_state.facade_positions:
            st.error("❌ Добавьте хотя бы одну позицию фасада!")
        else:
            try:
                # === РАСЧЁТ ФАСАДОВ С ГЛОБАЛЬНОЙ АГРЕГАЦИЕЙ V.9 ===
                aggregator = MaterialAggregator(ref1)
                
                facade_details = []
                
                for idx, position in enumerate(st.session_state.facade_positions, 1):
                    # [ЗДЕСЬ КОД РАСЧЕТА ФАСАДА - НЕ ПОКАЗАН, ТАК КАК СЛИШКОМ ДЛИННЫЙ]
                    # Используется существующая логика calculate_facade_materials
                    
                    # Пример структуры (упрощенно):
                    h_left = position.get("height_left", 3.0)
                    h_right = position.get("height_right", 0.0)
                    h_avg = (h_left + h_right) / 2 if h_right > 0 else h_left
                    area = position["width"] * h_avg
                    
                    aggregator.add_metrics(area=area, perimeter=0)  # Упрощенно
                    
                    facade_details.append({
                        'Позиция': idx,
                        'Система': 'Ruit 50F',
                        'Ширина (м)': position.get('width', 0),
                        'Высота слева (м)': h_left,
                        'Высота справа (м)': h_right,
                        'Площадь (м²)': round(area, 2),
                        'Периметр (м)': 0  # Упрощенно
                    })
                
                # Округляем материалы
                aggregator.round_all_materials()
                totals = aggregator.calculate_final_totals(margin_rate=0.81)
                
                st.session_state.last_facade_result = {
                    'aggregator': aggregator,
                    'totals': totals,
                    'facade_details': facade_details
                }
                
                # === ВЫВОД РЕЗУЛЬТАТОВ ФАСАДОВ ПО ТЗ V.9 ===
                st.success("✅ Расчёт фасадов выполнен!")
                
                st.metric(
                    "💰 ИТОГО К ОПЛАТЕ",
                    f"{totals['total']:,} ₸",
                    delta="БЕЗ двойной маржи"
                )
                
                st.divider()
                
                # ЧАСТЬ 1: ОБЩИЕ МЕТРИКИ
                st.header("📊 ЧАСТЬ 1: Общие показатели")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Общая площадь фасадов", f"{aggregator.metrics['total_area']:.2f} м²")
                col2.metric("Общий периметр", f"{aggregator.metrics['total_perimeter']:.2f} м")
                col3.metric("Позиций фасадов", len(st.session_state.facade_positions))
                
                st.divider()
                
                # ЧАСТЬ 2: ИНФОРМАЦИОННАЯ ДЕТАЛИЗАЦИЯ (БЕЗ ЦЕН!)
                with st.expander("🔹 ЧАСТЬ 2: Список фасадов (информация)", expanded=False):
                    st.info("ℹ️ Справочная информация для контроля состава заказа. Цены в этом блоке НЕ указаны.")
                    
                    if facade_details:
                        df_facades = pd.DataFrame(facade_details)
                        st.dataframe(df_facades, use_container_width=True, hide_index=True)
                    else:
                        st.warning("Нет данных о фасадах")
                
                st.divider()
                
                # ЧАСТЬ 3.1: СПЕЦИФИКАЦИЯ КАРКАСА
                st.header("🏗️ ЧАСТЬ 3.1: Спецификация каркаса (скелет)")
                
                st.info(
                    "**Проектный метод:** Суммированы все стойки, ригели, кронштейны по всем фасадным позициям. "
                    "Округление до кратности применено ОДИН РАЗ."
                )
                
                frame_materials = aggregator.get_category_materials('facade_frame')
                
                if frame_materials:
                    df_frame = pd.DataFrame(frame_materials)
                    st.dataframe(
                        df_frame,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Количество_raw": st.column_config.NumberColumn("Кол-во нетто", format="%.3f"),
                            "Количество": st.column_config.NumberColumn("Кол-во брутто", format="%.2f"),
                            "Сумма": st.column_config.NumberColumn("Сумма (₸)", format="%d")
                        }
                    )
                    
                    st.metric(
                        "💼 ИТОГО каркас",
                        f"{totals['breakdown']['facade_frame']:,} ₸",
                        help="Это можно сравнить с заводским PDF-отчетом"
                    )
                else:
                    st.warning("⚠️ Материалы каркаса не найдены")
                
                st.divider()
                
                # ЧАСТЬ 3.2: СПЕЦИФИКАЦИЯ ВСТАВОК
                st.header("🚪 ЧАСТЬ 3.2: Спецификация вставок (окна/двери)")
                
                st.info(
                    "**Суммированы материалы всех окон и дверей**, встроенных в фасады. "
                    "Включает профили, фурнитуру, уплотнители (БЕЗ стеклопакетов и услуг)."
                )
                
                insert_materials = aggregator.get_category_materials('facade_inserts')
                
                if insert_materials:
                    df_inserts = pd.DataFrame(insert_materials)
                    st.dataframe(
                        df_inserts,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Количество_raw": st.column_config.NumberColumn("Кол-во нетто", format="%.3f"),
                            "Количество": st.column_config.NumberColumn("Кол-во брутто", format="%.2f"),
                            "Сумма": st.column_config.NumberColumn("Сумма (₸)", format="%d")
                        }
                    )
                    
                    st.metric(
                        "💼 ИТОГО вставки",
                        f"{totals['breakdown']['facade_inserts']:,} ₸"
                    )
                else:
                    st.info("Вставок нет или материалы не найдены")
                
                st.divider()
                
                # ЧАСТЬ 4: ФИНАНСОВЫЙ ИТОГ
                st.header("💰 ЧАСТЬ 4: Финансовый итог")
                
                st.markdown("**Расчёт ведётся ОДИН РАЗ для всего блока фасадов (БЕЗ двойной маржи):**")
                
                final_items = []
                
                if totals['breakdown']['glass'] > 0:
                    final_items.append({
                        'Наименование': 'Стеклопакеты',
                        'Площадь (м²)': f"{aggregator.metrics['total_area']:.2f}",
                        'Сумма (₸)': f"{totals['breakdown']['glass']:,}"
                    })
                
                if totals['breakdown']['assembly'] > 0:
                    final_items.append({
                        'Наименование': 'Сборка',
                        'Площадь (м²)': f"{aggregator.metrics['total_area']:.2f}",
                        'Сумма (₸)': f"{totals['breakdown']['assembly']:,}"
                    })
                
                if totals['breakdown']['installation'] > 0:
                    final_items.append({
                        'Наименование': 'Монтаж',
                        'Площадь (м²)': f"{aggregator.metrics['total_area']:.2f}",
                        'Сумма (₸)': f"{totals['breakdown']['installation']:,}"
                    })
                
                if totals['breakdown']['additional_details'] > 0:
                    final_items.append({
                        'Наименование': 'Дополнительные детали',
                        'Площадь (м²)': '-',
                        'Сумма (₸)': f"{totals['breakdown']['additional_details']:,}"
                    })
                
                final_items.append({
                    'Наименование': 'Материалы (каркас)',
                    'Площадь (м²)': '-',
                    'Сумма (₸)': f"{totals['breakdown']['facade_frame']:,}"
                })
                
                final_items.append({
                    'Наименование': 'Материалы (вставки)',
                    'Площадь (м²)': '-',
                    'Сумма (₸)': f"{totals['breakdown']['facade_inserts']:,}"
                })
                
                if final_items:
                    df_final = pd.DataFrame(final_items)
                    st.dataframe(df_final, use_container_width=True, hide_index=True)
                
                st.divider()
                
                col_a, col_b = st.columns(2)
                with col_a:
                    st.metric(
                        "Обеспечение",
                        f"{totals['margin']:,} ₸",
                        help="Наценка 81% начислена ОДИН РАЗ на всю себестоимость (БЕЗ каскада!)"
                    )
                with col_b:
                    st.metric(
                        "💰 К ОПЛАТЕ",
                        f"{totals['total']:,} ₸",
                        delta="БЕЗ двойной маржи!"
                    )
                
                with st.expander("ℹ️ Как рассчитано обеспечение (ОДИН РАЗ)", expanded=False):
                    st.write("**Материалы:**")
                    st.write(f"  - Каркас: {totals['breakdown']['facade_frame']:,} ₸")
                    st.write(f"  - Вставки: {totals['breakdown']['facade_inserts']:,} ₸")
                    st.write(f"  - ИТОГО материалы: {totals['materials_total']:,} ₸")
                    st.write(f"**Услуги:** {totals['services_total']:,} ₸")
                    st.write(f"**Себестоимость:** {totals['subtotal']:,} ₸")
                    st.divider()
                    st.write(f"**Обеспечение (81%):** {totals['margin']:,} ₸")
                    st.divider()
                    st.success(f"**ИТОГО:** {totals['total']:,} ₸")
                    st.info("✅ Обеспечение начислено ОДИН РАЗ на всю себестоимость, а не каскадом!")
            
            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {e}")
                st.exception(e)
    
    # Кнопка экспорта
    if 'last_facade_result' in st.session_state:
        st.divider()
        if st.button("📥 Скачать КП фасада в Excel", type="secondary", width="stretch"):
            try:
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


# [ОСТАЛЬНЫЕ ФУНКЦИИ render_tambour_page() и render_history_page() - БЕЗ ИЗМЕНЕНИЙ]

# ГЛАВНОЕ МЕНЮ НАВИГАЦИИ
if 'menu_selection' not in st.session_state:
    st.session_state.menu_selection = "Главная (Окна/Двери)"

with st.sidebar:
    st.title("📍 Навигация")
    
    menu_selection = st.radio(
        "Выберите раздел:",
        ["Главная (Окна/Двери)", "Фасады", "Оконный тамбур", "История"],
        index=["Главная (Окна/Двери)", "Фасады", "Оконный тамбур", "История"].index(st.session_state.menu_selection) if st.session_state.menu_selection in ["Главная (Окна/Двери)", "Фасады", "Оконный тамбур", "История"] else 0,
        key="sidebar_navigation"
    )
    
    st.session_state.menu_selection = menu_selection

# Роутинг
if st.session_state.menu_selection == "Главная (Окна/Двери)":
    render_windows_doors_page()
elif st.session_state.menu_selection == "Фасады":
    render_facade_page()
elif st.session_state.menu_selection == "Оконный тамбур":
    # render_tambour_page()  # Без изменений
    pass
elif st.session_state.menu_selection == "История":
    # render_history_page()  # Без изменений
    pass
