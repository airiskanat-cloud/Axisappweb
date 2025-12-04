import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from zipfile import BadZipFile
from pathlib import Path

# ================== НАСТРОЙКИ ==================

# Имя Excel-файла в корне проекта (рядом с этим .py)
EXCEL_FILE = "axis_pro_gf.xlsx"


# ================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==================

def make_unique_columns(cols):
    """
    Делаем имена колонок уникальными и без None.
    Если в Excel есть одинаковые заголовки или пустые – здесь они исправятся.
    """
    seen = {}
    new = []
    for i, c in enumerate(cols):
        base = str(c) if c is not None else f"col_{i+1}"
        count = seen.get(base, 0)
        if count == 0:
            new_name = base
        else:
            new_name = f"{base}_{count}"
        seen[base] = count + 1
        new.append(new_name)
    return new


class ExcelClient:
    """Класс для чтения/записи Excel через openpyxl + pandas."""

    def __init__(self, filename: str):
        self.filename = filename
        self.wb = None
        self.load()

    def load(self):
        """Загружаем книгу Excel с вычисленными значениями (data_only=True)."""
        self.wb = load_workbook(self.filename, data_only=True)

    @property
    def sheets(self):
        """Список имён листов книги."""
        return self.wb.sheetnames

    def get_sheet_df(self, sheet_name: str) -> pd.DataFrame:
        """
        Читаем лист в DataFrame.
        Первая строка — заголовки колонок, остальные — данные.
        """
        ws = self.wb[sheet_name]
        data = list(ws.values)

        if not data:
            return pd.DataFrame()

        header = list(data[0])
        rows = data[1:]
        df = pd.DataFrame(rows, columns=header)

        # Делаем названия колонок безопасными (без дубликатов и None)
        df.columns = make_unique_columns(df.columns)

        return df

    def save_df_to_sheet(self, df: pd.DataFrame, sheet_name: str):
        """
        Полностью перезаписываем лист содержимым df.
        ВНИМАНИЕ: старый лист будет удалён.
        """
        # если лист есть — удаляем и создаём заново
        if sheet_name in self.wb.sheetnames:
            ws_old = self.wb[sheet_name]
            self.wb.remove(ws_old)
        ws = self.wb.create_sheet(sheet_name)

        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        self.wb.save(self.filename)


# ================== ОСНОВНОЕ ПРИЛОЖЕНИЕ ==================

def main():
    st.set_page_config(page_title="Axis Web", layout="wide")

    st.title("Axis Web – работа с Excel")
    st.caption(f"Файл данных: `{EXCEL_FILE}`")

    # --- Проверяем, что файл существует физически ---
    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists():
        st.error(
            f"Файл **{EXCEL_FILE}** не найден в папке проекта.\n\n"
            "Проверьте, что Excel-файл лежит рядом с Axisapp_web.py, "
            "затем выполните `git add`, `git commit`, `git push` и заново Deploy на Render."
        )
        st.stop()

    # --- Пытаемся загрузить книгу Excel ---
    try:
        excel = ExcelClient(EXCEL_FILE)
    except BadZipFile:
        st.error(
            f"Файл **{EXCEL_FILE}** повреждён или не является настоящим `.xlsx`.\n\n"
            "Откройте его в Excel и сохраните как **Книга Excel (*.xlsx)**, "
            "замените файл в проекте и снова задеплойте приложение."
        )
        st.stop()
    except Exception as e:
        st.error(f"Ошибка при открытии Excel: {e}")
        st.stop()

    # --- Выбор листа слева ---
    sheet_name = st.sidebar.selectbox("Выберите лист Excel", excel.sheets)

    # --- Загружаем данные выбранного листа ---
    df = excel.get_sheet_df(sheet_name)

    st.subheader(f"Данные листа: {sheet_name}")
    if df.empty:
        st.info("На этом листе пока нет данных.")
    else:
        # Просто показываем таблицу; ширину Streamlit подберёт сам
        st.dataframe(df)

    # --- Форма добавления строки ---
    st.subheader("Добавить новую строку")

    if df.empty:
        st.info(
            "Лист пустой: нет ни одной строки с данными.\n"
            "Добавьте хотя бы одну строку в Excel вручную, чтобы появились колонки."
        )
        return

    with st.form("add_row_form"):
        inputs = {}
        for col in df.columns:
            # простое текстовое поле для каждого столбца
            val = st.text_input(str(col))
            inputs[col] = val

        submitted = st.form_submit_button("Сохранить строку")

        if submitted:
            new_row = pd.DataFrame([inputs])
            new_df = pd.concat([df, new_row], ignore_index=True)
            try:
                excel.save_df_to_sheet(new_df, sheet_name)
                st.success("Строка сохранена в Excel. Обновите страницу (Ctrl+R), чтобы увидеть изменения.")
            except Exception as e:
                st.error(f"Ошибка при сохранении в Excel: {e}")


if __name__ == "__main__":
    main()
