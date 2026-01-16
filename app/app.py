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

@st.cache_data
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
            
            # ДОБАВЛЕНО: Дополнительные детали
            additional_options = ["Нет"] + [k.capitalize() for k in ref2.keys() if "нащельник" in k.lower()]
            additional_id = st.selectbox("Дополнительные детали", additional_options, key="main_additional")

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
            # Формирование данных заказа
            order_data = {
                "common": {
                    "order_number": order_num, 
                    "toning_id": toning_id, 
                    "assembly_id": assembly_id, 
                    "installation_id": install_id
                },
                "positions": st.session_state.get("positions", [])
            }
            
            # Расчет через новый движок для окон V2
            try:
                res = calculate_window_smeta(order_data, ref1, ref2, ref3)
                
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

                st.header("📊 Детальная смета AXIS")
                
                # Метрики вверху
                m_col1, m_col2, m_col3 = st.columns(3)
                m_col1.metric("Общая площадь", f"{res['metrics']['total_area']:.3f} м²")
                m_col2.metric("Суммарный периметр", f"{res['metrics']['total_perimeter']:.3f} м.п.")
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
            
            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {e}")
                st.exception(e)

    # Кнопка экспорта в Excel
    if 'last_result' in st.session_state and 'last_order_data' in st.session_state:
        st.divider()
        if st.button("📥 Скачать КП в Excel", type="secondary", use_container_width=True):
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
    
    if col_add.button("➕ Добавить позицию", use_container_width=True):
        # Генерация CODE для фасада
        facade_code = get_code_for_facade(facade_type_value)
        
        st.session_state.facade_positions.append({
            "code": facade_code,  # Добавляем CODE
            "facade_type": facade_type_value,  # Сохраняем тип для отображения
            "width": 6.0,
            "height": 3.0,
            "columns": 3,
            "rows": 2,
            "filling_type": "blind",
            "cells_data": []  # Данные для каждой ячейки
        })
        st.rerun()
    
    if col_clear.button("🗑️ Очистить всё", use_container_width=True):
        st.session_state.facade_positions = []
        if "last_facade_result" in st.session_state:
            del st.session_state.last_facade_result
        st.rerun()
    
    if col_new.button("🔄 Новый расчёт", use_container_width=True):
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
            col1, col2 = st.columns(2)
            pos["width"] = col1.number_input(
                "Ширина (м)", 
                min_value=0.5, 
                max_value=50.0,
                value=pos.get("width", 6.0), 
                step=0.1, 
                key=f"fac_w_{idx}"
            )
            pos["height"] = col2.number_input(
                "Высота (м)", 
                min_value=0.5, 
                max_value=20.0,
                value=pos.get("height", 3.0), 
                step=0.1, 
                key=f"fac_h_{idx}"
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
            
            # Расчет размера ячейки
            cell_w_m = pos["width"] / pos["columns"]
            cell_h_m = pos["height"] / pos["rows"]
            cell_w_mm = cell_w_m * 1000
            cell_h_mm = cell_h_m * 1000
            
            st.info(f"📐 Размер одной ячейки: {cell_w_m:.2f} × {cell_h_m:.2f} м ({cell_w_mm:.0f} × {cell_h_mm:.0f} мм)")
            
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
                
                # === ДОБАВЛЕНО: ВЫБОР ЗАПОЛНЕНИЯ ВСТАВКИ ===
                st.markdown("### 🎨 Заполнение вставки")
                st.caption("Выберите чем будет заполнена вставка (может отличаться от основного фасада)")
                
                insert_panel_category = st.selectbox(
                    "Категория заполнения вставки",
                    ["Стеклопакет", "Ламбри"],
                    key=f"insert_panel_cat_{idx}",
                    help="Заполнение для этой конкретной вставки"
                )
                
                if insert_panel_category == "Стеклопакет":
                    insert_glass_type = st.selectbox(
                        "Тип стеклопакета вставки",
                        GLASS_TYPES,
                        key=f"insert_glass_{idx}"
                    )
                    insert_fill_category = "Стеклопакет"
                    insert_fill_type = insert_glass_type
                else:
                    # Ламбри
                    lambri_types = []
                    for key in ref2.keys():
                        if "ламбри" in key.lower():
                            lambri_types.append(key)
                    if not lambri_types:
                        lambri_types = ["Ламбри без термо", "Ламбри с термо"]
                    
                    insert_lambri_type = st.selectbox(
                        "Тип ламбри вставки",
                        lambri_types,
                        key=f"insert_lambri_{idx}"
                    )
                    insert_fill_category = "Ламбри"
                    insert_fill_type = insert_lambri_type
                
                st.success(f"✅ Заполнение вставки: {insert_fill_category} - {insert_fill_type}")
                
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
                    
                    # ДОБАВЛЕНО: Сохраняем заполнение вставки
                    insert_data["fill_category"] = insert_fill_category
                    if insert_fill_category == "Стеклопакет":
                        insert_data["glass_type"] = insert_fill_type
                    else:
                        insert_data["lambri_type"] = insert_fill_type
                    
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
            use_container_width=True
        )
    
    with col_clear:
        if st.button("🗑️ Очистить", type="secondary", use_container_width=True):
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
                
                # Считаем общую площадь и периметр
                total_area = sum(
                    (p["width"] * p["height"]) / 1000000
                    for p in st.session_state.tambour_positions
                )
                total_perimeter = sum(
                    2 * ((p["width"] + p["height"]) / 1000)
                    for p in st.session_state.tambour_positions
                )
                
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
                st.dataframe(pd.DataFrame(products_data), use_container_width=True, hide_index=True)
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
                st.dataframe(pd.DataFrame(conn_data), use_container_width=True, hide_index=True)
                st.write(f"**Итого соединения:** {tambour_calc['total_connecting_cost']:,.0f} ₸")
                
                st.markdown("---")
                
                # Итоговый расчет
                st.subheader("💰 Итоговый расчет")
                part3_data = []
                for key, value in st.session_state.last_facade_result["part3_final"].items():
                    part3_data.append({"Наименование": key, "Сумма (₸)": f"{value:,.0f}"})
                
                df_part3 = pd.DataFrame(part3_data)
                st.dataframe(df_part3, use_container_width=True, hide_index=True)
                
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
                    # Расчет площади и периметра
                    total_area = 0
                    total_perimeter = 0
                    results = []
                
                    for idx, pos in enumerate(st.session_state.facade_positions):
                        area = pos["width"] * pos["height"]
                        perimeter = 2 * (pos["width"] + pos["height"])
                        total_area += area
                        total_perimeter += perimeter
                    
                        n_cells = pos["columns"] * pos["rows"]
                        
                        fill_name = {
                            "blind": "Глухое остекление",
                            "window": "Окно",
                            "door": "Дверь"
                        }.get(pos.get("filling_type", "blind"), "Неизвестно")
                        
                        results.append({
                            "Позиция": idx + 1,
                            "Габариты (м)": f"{pos['width']:.2f} × {pos['height']:.2f}",
                            "Площадь (м²)": f"{area:.2f}",
                            "Ячейки": f"{pos['columns']} × {pos['rows']} = {n_cells} шт",
                            "Тип заполнения": fill_name
                        })
                
                # УПРОЩЕННЫЙ РАСЧЕТ (т.к. нет справочника Ruit 50F)
                # На основе реального проекта из PDF:
                # 45.3 м² = 2,076,961 ₸ материалов (только профили + фурнитура)
                # Это ~46,000 ₸/м²
                
                # Определяем тип системы
                # Базовая стоимость профилей и материалов
                if "ALG" in facade_type_value or facade_type_value == "Оконный тамбур (ALG)":
                    # ОКОННАЯ СИСТЕМА (ALG 2030-45C и т.д.)
                    # ИСПРАВЛЕНО: Используем точный расчёт из engine_facade
                    
                    first_pos = st.session_state.facade_positions[0] if st.session_state.facade_positions else {}
                    W = first_pos.get("width", 6.0)
                    H = first_pos.get("height", 3.5)
                    cols = first_pos.get("columns", 3)
                    rows = first_pos.get("rows", 2)
                    count = len(st.session_state.facade_positions)
                    
                    print(f"\n🏗️ Вызов calculate_tambour_materials:")
                    print(f"   W={W}, H={H}, cols={cols}, rows={rows}, count={count}")
                    
                    tambour_calc = calculate_tambour_materials(
                        W=W,
                        H=H,
                        cols=cols,
                        rows=rows,
                        count=count,
                        ref1=ref1,
                        ref2=ref2,
                        ref3=ref3
                    )
                    
                    materials_cost = tambour_calc.get("total_cost", 0)
                    print(f"   ✅ Материалы рассчитаны: {materials_cost:,.0f}₸")
                else:
                    # ФАСАДНАЯ СИСТЕМА (Ruit 50F)
                    # ИСПРАВЛЕНО: Используем точный расчёт из engine_facade
                    
                    # Собираем данные о вставках (окна/двери)
                    facade_inserts = []
                    facade_calc_saved = None  # ИСПРАВЛЕНО: Инициализируем заранее
                    for pos in st.session_state.facade_positions:
                        filling_type = pos.get("filling_type", "blind")
                        
                        # Если это окно или дверь
                        if filling_type in ["window", "door"]:
                            insert_data = pos.get("insert_data", {})
                            
                            # ИСПРАВЛЕНО: Создаём ОДНУ вставку в указанной ячейке (не цикл!)
                            facade_inserts.append({
                                "type": filling_type,
                                "cell_col": pos.get("insert_col", 1),
                                "cell_row": pos.get("insert_row", 1),
                                "width": insert_data.get("width", 1800) / 1000,  # в метры
                                "height": insert_data.get("height", 2200) / 1000,
                                "system": insert_data.get("system", "ALG 2030-63C"),
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
                            
                            print(f"📍 Вставка создана: тип={filling_type}, ячейка=({pos.get('insert_col', 1)}, {pos.get('insert_row', 1)})")
                    
                    # Для первой позиции берём габариты
                    first_pos = st.session_state.facade_positions[0] if st.session_state.facade_positions else {}
                    W = first_pos.get("width", 6.0)
                    H = first_pos.get("height", 3.5)
                    cols = first_pos.get("columns", 3)
                    rows = first_pos.get("rows", 2)
                    count = len(st.session_state.facade_positions)
                    
                    print(f"\n🏗️ Вызов calculate_facade_materials:")
                    print(f"   W={W}, H={H}, cols={cols}, rows={rows}, count={count}")
                    print(f"   Вставок: {len(facade_inserts)}")
                    
                    # Вызываем расчёт
                    facade_calc = calculate_facade_materials(
                        W=W,
                        H=H,
                        cols=cols,
                        rows=rows,
                        count=count,
                        inserts=facade_inserts,
                        facade_profiles_ref=ref_facade,
                        ref1=ref1,
                        ref2=ref2,
                        ref3=ref3
                    )
                    
                    materials_cost = facade_calc.get("total_cost", 0)
                    print(f"   ✅ Материалы рассчитаны: {materials_cost:,.0f}₸")
                    
                    # ИСПРАВЛЕНО: Сохраняем facade_calc для отображения детализации
                    facade_calc_saved = facade_calc
                
                # Стеклопакеты/ламбри - собираем данные ПО ТИПАМ
                glass_areas = {"Двойной": 0, "Тройной": 0, "Энергодвойной": 0}
                lambri_areas = {}  # ИСПРАВЛЕНО: по типам
                
                for pos in st.session_state.facade_positions:
                    area = pos["width"] * pos["height"]
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
                        insert_data = pos.get("insert_data", {})
                        fill_category = insert_data.get("fill_category", "Стеклопакет")
                        
                        if fill_category == "Стеклопакет":
                            glass_type_raw = insert_data.get("glass_type", "двойной")
                            glass_type = glass_type_raw.capitalize()  # двойной → Двойной
                            
                            # Площадь вставок
                            n_cells = pos["columns"] * pos["rows"]
                            cell_w = pos["width"] / pos["columns"]
                            cell_h = pos["height"] / pos["rows"]
                            insert_w = insert_data.get("width", cell_w * 1000) / 1000
                            insert_h = insert_data.get("height", cell_h * 1000) / 1000
                            insert_area = insert_w * insert_h * n_cells
                            
                            glass_areas[glass_type] = glass_areas.get(glass_type, 0) + insert_area
                        
                        # ИСПРАВЛЕНО: Добавлена обработка ламбри из вставок
                        elif "Ламбри" in fill_category:
                            lambri_type = fill_category  # "Ламбри без термо" или "Ламбри с термо"
                            
                            # Площадь вставок с ламбри
                            n_cells = pos["columns"] * pos["rows"]
                            cell_w = pos["width"] / pos["columns"]
                            cell_h = pos["height"] / pos["rows"]
                            insert_w = insert_data.get("width", cell_w * 1000) / 1000
                            insert_h = insert_data.get("height", cell_h * 1000) / 1000
                            insert_area = insert_w * insert_h * n_cells
                            
                            lambri_areas[lambri_type] = lambri_areas.get(lambri_type, 0) + insert_area
                
                # РАСЧЕТ СТЕКЛОПАКЕТОВ (по общей площади каждого типа)
                glass_cost = 0
                for glass_type, total_glass_area in glass_areas.items():
                    if total_glass_area > 0:
                        # ИСПРАВЛЕНО: берём из ref2 с нормализацией регистра
                        price_per_m2 = ref2.get(glass_type.lower(), 9000)
                        glass_cost += total_glass_area * price_per_m2
                
                # РАСЧЕТ ЛАМБРИ (по общей площади КАЖДОГО ТИПА)
                lambri_cost = 0
                for lambri_type, lambri_area in lambri_areas.items():
                    if lambri_area > 0:
                        # Кол-во к отгрузке = ceil(площадь / 6)
                        q_otgr = math.ceil(lambri_area / 6.0)
                        
                        # ИСПРАВЛЕНО: берём из ref2 по типу, нормализация регистра
                        price_per_m_lambri = ref2.get(lambri_type.lower(), 2248)
                        
                        # Сумма = цена_за_метр * (кол-во_хлыстов * 6м)
                        lambri_cost += price_per_m_lambri * (q_otgr * 6)
                
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
                
                # ДОБАВЛЕНО: Дополнительные детали
                additional_cost = 0
                # Ищем "Нащельник" в ref2
                additional_name = None
                for key in ref2.keys():
                    if "нащельник" in key.lower():
                        additional_name = key
                        break
                
                if additional_name:
                    price_additional = ref2.get(additional_name, 0)
                    # Формула: ОКРУГЛЕНИЕ ВВЕРХ (периметр / 3) * цена
                    additional_cost = math.ceil(total_perimeter / 3) * price_additional
                
                # Сумма без обеспечения
                subtotal = materials_cost + glass_cost + lambri_cost + toning_cost + assembly_cost + installation_cost + additional_cost
                
                # Обеспечение 81% (было 65%)
                margin = subtotal * 0.81
                total_cost = subtotal + margin
                
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
                
                # Стеклопакеты ПО ТИПАМ (ВСЕГДА показываем)
                total_glass_cost_all = 0
                for glass_type, glass_area in glass_areas.items():
                    price_glass = ref2.get(glass_type.lower(), 9000)
                    cost_glass_type = glass_area * price_glass if glass_area > 0 else 0
                    st.session_state.last_facade_result["part3_final"][f"Стеклопакет ({glass_type})"] = round(cost_glass_type, 0)
                    total_glass_cost_all += cost_glass_type
                
                # Если НИ ОДНОГО типа стеклопакета не было, добавим общую строку
                if not glass_areas:
                    st.session_state.last_facade_result["part3_final"]["Стеклопакет"] = 0
                
                # Ламбри ПО ТИПАМ (ВСЕГДА показываем)
                total_lambri_cost_all = 0
                for lambri_type, lambri_area in lambri_areas.items():
                    q_otgr = math.ceil(lambri_area / 6.0) if lambri_area > 0 else 0
                    price_lambri = ref2.get(lambri_type.lower(), 2248)
                    cost_lambri_type = price_lambri * (q_otgr * 6) if lambri_area > 0 else 0
                    st.session_state.last_facade_result["part3_final"][f"Ламбри ({lambri_type})"] = round(cost_lambri_type, 0)
                    total_lambri_cost_all += cost_lambri_type
                
                # Если НИ ОДНОГО типа ламбри не было, добавим общую строку
                if not lambri_areas:
                    st.session_state.last_facade_result["part3_final"]["Ламбри"] = 0
                
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
                    "facade_calc": facade_calc_saved  # ИСПРАВЛЕНО: Сохраняем для детализации
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
                col3.metric("💰 ИТОГО К ОПЛАТЕ", f"{total_cost:,.0f} ₸")
                
                st.markdown("---")
                
                # Таблица позиций
                st.subheader("Детализация по позициям")
                df = pd.DataFrame(results)
                st.dataframe(df, use_container_width=True, hide_index=True)
                
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
                        st.dataframe(pd.DataFrame(tambour_data), use_container_width=True, hide_index=True)
                        st.write(f"**Итого материалы:** {tambour_calc.get('total_cost', 0):,.0f} ₸")
                    st.markdown("---")
                else:
                    # Для Ruit 50F показываем детализацию
                    st.write("**Материалы каркаса (Ruit 50F):**")
                    # ИСПРАВЛЕНО: Берём из session_state вместо locals()
                    facade_calc = st.session_state.get("last_facade_result", {}).get("facade_calc")
                    if facade_calc and facade_calc.get("skeleton"):
                        skeleton_data = []
                        for elem, data in facade_calc["skeleton"].items():
                            skeleton_data.append({
                                "Элемент": elem,
                                "Количество": f"{data['quantity']} {data['unit']}",
                                "Цена": f"{data['price']:,.0f} ₸",
                                "Стоимость": f"{data['cost']:,.0f} ₸"
                            })
                        st.dataframe(pd.DataFrame(skeleton_data), use_container_width=True, hide_index=True)
                        st.write(f"**Итого каркас:** {facade_calc.get('skeleton_cost', 0):,.0f} ₸")
                    
                    # ИСПРАВЛЕНО: Проверяем через session_state
                    if facade_calc and facade_calc.get("inserts_details"):
                        st.write("**Материалы вставок (двери/окна):**")
                        inserts_data = []
                        for insert in facade_calc["inserts_details"]:
                            inserts_data.append({
                                "Изделие": insert["name"],
                                "Размер": insert["size"],
                                "Стоимость": f"{insert['cost']:,.0f} ₸"
                            })
                        st.dataframe(pd.DataFrame(inserts_data), use_container_width=True, hide_index=True)
                        st.write(f"**Итого вставки:** {facade_calc.get('inserts_cost', 0):,.0f} ₸")
                    
                    st.markdown("---")
                
                part3_data = []
                for key, value in st.session_state.last_facade_result["part3_final"].items():
                    part3_data.append({"Наименование": key, "Сумма (₸)": f"{value:,.0f}"})
                
                df_part3 = pd.DataFrame(part3_data)
                st.dataframe(df_part3, use_container_width=True, hide_index=True)
                
                st.metric("🎯 ИТОГО К ОПЛАТЕ", f"{total_cost:,.0f} ₸", help="С учетом обеспечения 81%")
                
            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {e}")
                with st.expander("🔍 Детали ошибки"):
                    import traceback
                    st.code(traceback.format_exc())
    
                # Таблица результатов
                st.subheader("Детализация по позициям")
                df = pd.DataFrame(results)
                st.dataframe(df, use_container_width=True, hide_index=True)
                
                # Предупреждение
                st.warning("⚠️ Это упрощенный расчет. Для точной стоимости необходимо добавить расчет профилей, стеклопакетов и фурнитуры из справочников.")
                
            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {e}")
                st.exception(e)
    
    # === КНОПКА ЭКСПОРТА ===
    if 'last_facade_result' in st.session_state:
        st.divider()
        if st.button("📥 Скачать КП фасада в Excel", type="secondary", use_container_width=True):
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
    
    if col_add.button("➕ Добавить позицию", use_container_width=True, key="tambour_add_btn"):
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
    
    if col_clear.button("🗑️ Очистить всё", use_container_width=True, key="tambour_clear_btn"):
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
    
    if st.button("🚀 РАССЧИТАТЬ ТАМБУР", type="primary", use_container_width=True, key="tambour_calc_btn"):
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
                
                # Добавляем CODE к позициям И КОНВЕРТИРУЕМ ММ В МЕТРЫ
                for pos in st.session_state.tambour_positions:
                    # ВАЖНО: Создаём КОПИЮ чтобы не менять session_state!
                    pos_copy = pos.copy()
                    pos_copy["code"] = get_code_for_windows_doors(pos["product_type"], pos["system_id"])
                    # engine_windows ожидает метры, в session_state хранятся мм
                    pos_copy["width"] = pos["width"] / 1000.0
                    pos_copy["height"] = pos["height"] / 1000.0
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
                    st.dataframe(pd.DataFrame(positions_df), use_container_width=True, hide_index=True)
                
                # Материалы
                with st.expander("📦 Материалы (Артикулы)", expanded=True):
                    if result["part2_materials"]:
                        df2 = pd.DataFrame(result["part2_materials"])
                        st.dataframe(df2, use_container_width=True, hide_index=True)
                
                # Итоговый расчёт
                with st.expander("💰 Итоговый расчет", expanded=True):
                    df3 = pd.DataFrame(result["part3_final"].items(), columns=["Наименование", "Сумма (₸)"])
                    st.dataframe(df3, use_container_width=True, hide_index=True)
                    st.metric("🎯 ИТОГО К ОПЛАТЕ", f"{result['total_with_margin']:,.0f} ₸")
                
                # ДОБАВЛЕНО: Кнопка скачать КП
                st.markdown("---")
                if st.button("📥 Скачать КП в Excel", type="secondary", use_container_width=True, key="tambour_export_btn"):
                    try:
                        # Экспорт уже в начале файла импортирован
                        temp_dir = tempfile.gettempdir()
                        order_num = order_number
                        excel_path = os.path.join(temp_dir, f"KP_{order_num}.xlsx")
                        
                        export_to_excel(order_data, result, excel_path)
                        
                        with open(excel_path, "rb") as f:
                            st.download_button(
                                label="💾 Сохранить файл",
                                data=f,
                                file_name=f"KP_{order_num}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        st.success(f"✅ Файл {order_num}.xlsx готов к скачиванию")
                    except Exception as e:
                        st.error(f"❌ Ошибка экспорта: {e}")
                
                # Сохраняем в историю
                save_history(
                    credentials_path=GOOGLE_CREDENTIALS_PATH,
                    spreadsheet_id=SPREADSHEET_ID,
                    user_login=st.session_state.get("user_email", "unknown"),
                    order_data=order_data,
                    result=result
                )
                
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
        
        # Создаём DataFrame
        import pandas as pd
        df = pd.DataFrame(rows, columns=headers)
        
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
            st.dataframe(df, use_container_width=True, hide_index=True)
    
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
