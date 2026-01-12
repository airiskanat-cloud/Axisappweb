import streamlit as st
import sys
import os
import pandas as pd
from pathlib import Path
import datetime
import tempfile

# --- ФИКСАЦИЯ ПУТЕЙ (Стандарт Axis Pro GF) ---
current_file = Path(__file__).resolve()
root_dir = current_file.parents[1] 
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

# Импорты внутренних модулей
from auth.auth import authenticate
from config.settings import SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH
from references.sheets_reader import load_reference_1, load_reference_2, load_reference_3
from calculations.engine_windows import calculate_window_smeta, calculate_impost_length, SYSTEM_MAPPING
from calculations.mapping import get_code_for_windows_doors, get_code_for_facade
from export.export_kp import export_to_excel
from history.save_history import save_history
from calculations.engine_facade import calculate_facade_smeta


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
GLASS_TYPES = ["Двойной", "Тройной", "Энергодвойной", "Энерготройной", "Одинарный 4мм", "Одинарный 6мм", "Одинарный 4мм закал", "Одинарный 6мм закал", "Нет"]
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
@st.cache_data(ttl=3600)
def get_data():
    r1 = load_reference_1(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r2 = load_reference_2(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r3 = load_reference_3(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    return r1, r2, r3

# Вызов с индикатором
with st.spinner('⏳ Загрузка справочников из Google Sheets...'):
    ref1, ref2, ref3 = get_data()

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
        if st.button("🔄 Очистить и Новый расчет", use_container_width=True):
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

    with col_right:
        st.subheader(f"🪟 Список позиций")
        
        if "positions" not in st.session_state: 
            st.session_state.positions = []
        
        if st.button("➕ Добавить позицию", use_container_width=True):
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

    if st.button("🚀 РАССЧИТАТЬ", type="primary", use_container_width=True):
        if not st.session_state.positions:
            st.error("❌ Добавьте хотя бы одну позицию!")
        else:
            order_data = {
                "common": {
                    "order_number": order_num,
                    "toning_id": toning_id,
                    "assembly_id": assembly_id,
                    "installation_id": install_id
                },
                "positions": st.session_state.get("positions", [])
            }

            try:
                if st.session_state.menu_selection == "Фасады":
                    res = calculate_facade_smeta(order_data, ref2)
                else:
                    res = calculate_window_smeta(order_data, ref1, ref2, ref3)

                st.session_state.last_result = res
                st.session_state.last_order_data = order_data

            except Exception as e:
                st.error(f"Ошибка расчета: {e}")
                st.exception(e)

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
                
    # ===== ПОКАЗ РЕЗУЛЬТАТОВ =====
    if "last_result" in st.session_state:
        res = st.session_state.last_result

        st.header("📊 Детальная смета AXIS")

        m_col1, m_col2, m_col3 = st.columns(3)
        m_col1.metric("Общая площадь", f"{res['metrics']['total_area']:.3f} м²")
        m_col2.metric("Суммарный периметр", f"{res['metrics']['total_perimeter']:.3f} м")
        m_col3.metric("💰 ИТОГО К ОПЛАТЕ", f"{res['total_with_margin']:,} ₸")

        st.divider()
        
        # ЧАСТЬ 1: Габаритная ведомость (СВЕРНУТАЯ)
        with st.expander("🔹 ЧАСТЬ 1: Габаритная ведомость (общая по типам)", expanded=False):
            if res["part1_summary"]:
                st.markdown("#### 📊 Общий расчет элементов по типам изделия:")
                
                # Группируем по типу изделия
                by_type = {}
                for item in res["part1_summary"]:
                    prod_type = item["Тип изделия"]
                    if prod_type not in by_type:
                        by_type[prod_type] = []
                    by_type[prod_type].append(item)
                
                for prod_type, items in by_type.items():
                    st.markdown(f"**{prod_type}:**")
                    df = pd.DataFrame(items)
                    df = df[["Категория", "Элемент", "Значение"]]
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    st.markdown("---")
            else:
                st.info("Данные для габаритной ведомости отсутствуют")

        # ЧАСТЬ 2: Ведомость материалов
        with st.expander("🔹 ЧАСТЬ 2: Ведомость материалов (Артикулы)", expanded=True):
            if res["part2_materials"]:
                df2 = pd.DataFrame(res["part2_materials"])
                st.dataframe(df2, use_container_width=True, hide_index=True)
                
                # Итого по материалам
                total_mat = sum(m["Сумма"] for m in res["part2_materials"])
                st.metric("Итого материалы", f"{total_mat:,} ₸")
            else:
                st.warning("⚠️ Данные для ведомости материалов отсутствуют")
                st.info("Возможные причины: не найдены материалы в Справочнике-1 для выбранной системы")

        # ЧАСТЬ 3: Итоговый расчет
        with st.expander("🔹 ЧАСТЬ 3: Итоговый расчет", expanded=True):
            df3 = pd.DataFrame(res["part3_final"].items(), columns=["Наименование", "Сумма (₸)"])
            st.dataframe(df3, use_container_width=True, hide_index=True)
        
        # Отладочная информация (скрытая)
        with st.expander("🔍 Отладочная информация", expanded=False):
            st.json(res["debug_info"])

    # Кнопка экспорта в Excel
    if 'last_result' in st.session_state and 'last_order_data' in st.session_state:
        st.divider()
        if st.button("📥 Скачать КП в Excel", type="secondary", use_container_width=True):
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

                
