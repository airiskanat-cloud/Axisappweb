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
from export.export_kp import export_to_excel
from history.save_history import save_history


def facade_ui(prefix, pos_idx):
    st.markdown(f"#### 🏗️ Настройка фасада №{pos_idx+1}")
    
    col1, col2 = st.columns(2)
    with col1:
        width = st.number_input("Общая ширина (мм)", min_value=100, value=3000, key=f"{prefix}_w")
        height = st.number_input("Общая высота (мм)", min_value=100, value=4000, key=f"{prefix}_h")
    
    with col2:
        cols = st.number_input("Кол-во вертикальных стоек", min_value=1, value=3, key=f"{prefix}_cols")
        rows = st.number_input("Кол-во горизонтальных ригелей", min_value=1, value=2, key=f"{prefix}_rows")
        
    wind_load = st.select_slider("Ветровая нагрузка (кг/м²)", options=[30, 40, 50, 60, 80], value=50, key=f"{prefix}_wind")
    
    # Расчет требуемой инерции
    span_width = (width / cols) / 1000 if cols > 0 else width / 1000
    span_height = height / 1000
    
    # Расчет требуемой инерции (временно отключен - требует facade_pro)
    # req_jx = calculate_required_jx(span_height, span_width, wind_load)
    # best_mullion = find_best_mullion(req_jx)
    # st.info(f"📊 **Тех. расчет:** Требуемый Jx = {req_jx} см⁴. \n\n "
    #         f"✅ **Рекомендуемая стойка:** {best_mullion['art']} (Jx={best_mullion['jx']})")
    
    return {
        "type": "Фасад",
        "width": width,
        "height": height,
        "grid": {"cols": cols, "rows": rows},
        "mullion_art": best_mullion['art'],
        "wind_load": wind_load
    }

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
            st.session_state.current_user = {"login": login, "data": user}  # Сохраняем пользователя
            st.rerun()
        else:
            st.error("Ошибка входа")
    st.stop()

# --- 2. ЗАГРУЗКА ДАННЫХ ---

@st.cache_data
def get_data():
    r1 = load_reference_1(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r2 = load_reference_2(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    r3 = load_reference_3(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
    return r1, r2, r3

ref1, ref2, ref3 = get_data()

# --- 3. ФУНКЦИЯ КОНСТРУКТОР ОКНА ---
def window_door_ui(prefix, pos_idx, system_id):
    """Форма для заполнения данных окна/двери"""
    st.markdown("---")
    
    # Габариты
    st.markdown("### 📐 Габариты изделия")
    c1, c2 = st.columns(2)
    w = c1.number_input("Ширина (мм)", min_value=0.0, value=2000.0, step=50.0, key=f"{prefix}_w")
    h = c2.number_input("Высота (мм)", min_value=0.0, value=1560.0, step=50.0, key=f"{prefix}_h")

    # Импосты
    st.markdown("### 🔲 Импосты")
    
    auto_imposts = st.checkbox("✅ Автоматический расчет (рекомендуется)", value=True, key=f"{prefix}_auto_imp")
    
    if auto_imposts:
        st.caption("💡 Длина импостов рассчитывается автоматически по системе профиля")
        ic1, ic2, ic3, ic4 = st.columns(4)
        has_left = ic1.checkbox("Левый", key=f"{prefix}_has_il")
        has_center = ic2.checkbox("Центральный", key=f"{prefix}_has_ic")
        has_right = ic3.checkbox("Правый", key=f"{prefix}_has_ir")
        has_tor = ic4.checkbox("ТОР (гориз.)", key=f"{prefix}_has_it")
        
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
        il = i1.number_input("Левый (мм)", min_value=0, value=0, step=50, key=f"{prefix}_il")
        ic = i2.number_input("Центр (мм)", min_value=0, value=0, step=50, key=f"{prefix}_ic")
        ir = i3.number_input("Правый (мм)", min_value=0, value=0, step=50, key=f"{prefix}_ir")
        it = i4.number_input("ТОР (мм)", min_value=0, value=0, step=50, key=f"{prefix}_it")
        
        imposts_data = {
            "auto_calculate": False,
            "left": il,
            "center": ic,
            "right": ir,
            "tor": it
        }
    
    # Створки
    st.markdown("### 🚪 Створки")
    s_count = st.number_input("Количество створок", min_value=0, max_value=10, value=1, step=1, key=f"{prefix}_sc")
    sashes = []
    
    if s_count > 0:
        st.caption("💡 Для расчета точек запирания и фурнитуры используется первая створка")
        for s in range(s_count):
            with st.expander(f"Створка №{s+1}", expanded=(s==0)):
                sc1, sc2 = st.columns(2)
                sw = sc1.number_input(f"Ширина", min_value=0, value=952, step=50, key=f"{prefix}_sw{s}")
                sh = sc2.number_input(f"Высота", min_value=0, value=512, step=50, key=f"{prefix}_sh{s}")
                sashes.append({"w": sw, "h": sh})
    
    # Заполнение
    st.markdown("### 🖼 Заполнение")
    fill_cat = st.selectbox("Тип заполнения", PANELS, key=f"{prefix}_fill_cat")
    
    selected_glass = "Нет"
    if fill_cat == "Стеклопакет":
        selected_glass = st.selectbox("Тип стеклопакета", GLASS_TYPES, key=f"{prefix}_glass")
        
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

# --- 4. ШАПКА КОМПАНИИ ---
header_col1, header_col2 = st.columns([2, 1])
with header_col1:
    st.title("🚀 Axis Pro GF - Калькулятор окон V2")
    st.markdown("""
    **Компания «AXIS»** 📍 Город: Астана  
    📞 Тел.: +7 707 504 4040 | 📧 E-mail: Axisokna.kz@mail.ru | 🌐 Сайт: www.axis.kz
    """)
with header_col2:
    if st.button("🔄 Очистить и Новый расчет", use_container_width=True):
        for key in list(st.session_state.keys()):
            if key != 'authenticated':
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
            
            # НОВОЕ: Тип изделия и система на уровне позиции
            pc1, pc2 = st.columns(2)
            
            # Определяем текущий индекс для типа изделия
            current_type = pos.get("product_type", "Окно с откр.")
            try:
                type_index = PRODUCT_TYPES.index(current_type)
            except ValueError:
                type_index = 0
            
            pos["product_type"] = pc1.selectbox(
                "Тип изделия", 
                PRODUCT_TYPES,  # ИСПРАВЛЕНО: Используем полный список (окна + двери + фасад)
                key=f"pc_type{idx}",
                index=type_index
            )
            
            pos["system_id"] = pc2.selectbox(
                "Система профиля", 
                PROFILE_SYSTEMS, 
                key=f"pc_sys{idx}",
                index=0
            )
            
            # ИСПРАВЛЕНО: Убрано поле "Количество одинаковых изделий"
            # Теперь: 1 позиция = 1 изделие
            pos["count"] = 1
            
            pos["data"] = window_door_ui(f"main_pos_{idx}", idx, pos["system_id"])

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
