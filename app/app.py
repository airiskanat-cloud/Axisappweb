"""
НОВЫЙ БЛОК ВЫВОДА ДЛЯ ФАСАДОВ (V.9)
Вставить в app.py в блок render_facade_page()
"""

# После кнопки "РАССЧИТАТЬ ФАСАДЫ" и try:

# === РАСЧЁТ ФАСАДОВ С ГЛОБАЛЬНОЙ АГРЕГАЦИЕЙ V.9 ===
from calculations.material_basket_V9 import MaterialAggregator

# Создаём агрегатор
aggregator = MaterialAggregator(ref1)

# Проходим по ВСЕМ фасадным позициям
facade_details = []  # Для ЧАСТИ 2

for idx, position in enumerate(st.session_state.facade_positions, 1):
    # Рассчитываем фасад
    facade_result = calculate_facade_smeta(position, ref1, ref2, ref3)
    
    # Добавляем метрики
    area = facade_result.get('area', 0)
    perimeter = facade_result.get('perimeter', 0)
    aggregator.add_metrics(area=area, perimeter=perimeter)
    
    # Сохраняем детали для ЧАСТИ 2
    facade_details.append({
        'Позиция': idx,
        'Система': position.get('system', ''),
        'Ширина (м)': position.get('width', 0),
        'Высота слева (м)': position.get('height_left', 0),
        'Высота справа (м)': position.get('height_right', 0),
        'Площадь (м²)': round(area, 2),
        'Периметр (м)': round(perimeter, 2)
    })
    
    # === КАРКАС (facade_frame) ===
    frame_materials = facade_result.get('frame_materials', [])
    for material in frame_materials:
        aggregator.add_material(
            category='facade_frame',
            article=material.get('Артикул', ''),
            quantity_raw=material.get('quantity_raw', material.get('Количество', 0)),
            unit=material.get('Единица', 'м'),
            price=material.get('Цена', 0),
            name=material.get('Элемент', '')
        )
    
    # === ВСТАВКИ (facade_inserts) ===
    insert_materials = facade_result.get('insert_materials', [])
    for material in insert_materials:
        aggregator.add_material(
            category='facade_inserts',
            article=material.get('Артикул', ''),
            quantity_raw=material.get('quantity_raw', material.get('Количество', 0)),
            unit=material.get('Единица', 'шт'),
            price=material.get('Цена', 0),
            name=material.get('Элемент', '')
        )
    
    # === УСЛУГИ ===
    services = facade_result.get('services', {})
    aggregator.add_service('glass_cost', services.get('glass', 0))
    aggregator.add_service('assembly_cost', services.get('assembly', 0))
    aggregator.add_service('installation_cost', services.get('installation', 0))
    aggregator.add_service('additional_details_cost', services.get('additional_details', 0))

# Округляем ВСЕ материалы ОДИН РАЗ
aggregator.round_all_materials()

# Рассчитываем финальные итоги (БЕЗ двойной маржи!)
totals = aggregator.calculate_final_totals(margin_rate=0.81)

# Сохраняем для экспорта
st.session_state.last_facade_result = {
    'aggregator': aggregator,
    'totals': totals,
    'facade_details': facade_details
}

# СОХРАНЕНИЕ ИСТОРИИ
try:
    current_user = st.session_state.get("current_user", {})
    user_login = current_user.get("login", "unknown")
    # ... сохранение истории
except Exception as e:
    st.warning(f"⚠️ История не сохранена: {e}")

# ============================================================
# ВЫВОД РЕЗУЛЬТАТОВ ФАСАДОВ ПО ТЗ V.9
# ============================================================

st.success("✅ Расчёт фасадов выполнен!")

# Главная метрика
st.metric(
    "💰 ИТОГО К ОПЛАТЕ",
    f"{totals['total']:,} ₸",
    delta="БЕЗ двойной маржи"
)

st.divider()

# ============================================================
# ЧАСТЬ 1: ОБЩИЕ МЕТРИКИ
# ============================================================
st.header("📊 ЧАСТЬ 1: Общие показатели")

col1, col2, col3 = st.columns(3)
col1.metric("Общая площадь фасадов", f"{aggregator.metrics['total_area']:.2f} м²")
col2.metric("Общий периметр", f"{aggregator.metrics['total_perimeter']:.2f} м")
col3.metric("Позиций фасадов", len(st.session_state.facade_positions))

st.divider()

# ============================================================
# ЧАСТЬ 2: ИНФОРМАЦИОННАЯ ДЕТАЛИЗАЦИЯ (БЕЗ ЦЕН!)
# ============================================================
with st.expander("🔹 ЧАСТЬ 2: Список фасадов (информация)", expanded=False):
    st.info("ℹ️ Справочная информация для контроля состава заказа. Цены в этом блоке НЕ указаны.")
    
    if facade_details:
        df_facades = pd.DataFrame(facade_details)
        st.dataframe(df_facades, use_container_width=True, hide_index=True)
    else:
        st.warning("Нет данных о фасадах")

st.divider()

# ============================================================
# ЧАСТЬ 3.1: СПЕЦИФИКАЦИЯ КАРКАСА
# ============================================================
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

# ============================================================
# ЧАСТЬ 3.2: СПЕЦИФИКАЦИЯ ВСТАВОК
# ============================================================
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

# ============================================================
# ЧАСТЬ 4: ФИНАНСОВЫЙ ИТОГ (БЕЗ ДВОЙНОЙ МАРЖИ!)
# ============================================================
st.header("💰 ЧАСТЬ 4: Финансовый итог")

st.markdown("**Расчёт ведётся ОДИН РАЗ для всего блока фасадов (БЕЗ двойной маржи):**")

# Таблица итогов
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
        f"{totals['margin']:,} ₸",
        help="Наценка 81% начислена ОДИН РАЗ на всю себестоимость (БЕЗ каскада!)"
    )
with col_b:
    st.metric(
        "💰 К ОПЛАТЕ",
        f"{totals['total']:,} ₸",
        delta="БЕЗ двойной маржи!"
    )

# Детализация
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

