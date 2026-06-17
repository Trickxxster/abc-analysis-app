# app.py — финальная версия
# Автоматизация расчёта заказа и ABC-анализа (Streamlit)
# Для работы требуется Python 3.8+, библиотеки streamlit, pandas, numpy, plotly, openpyxl

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

# ------------------------------------------------------------
# Настройка страницы (должна быть первой командой Streamlit)
# ------------------------------------------------------------
st.set_page_config(page_title="📊 Расчёт заказа и ABC-анализ", layout="wide")
st.title("📊 Автоматизация расчёта заказа и ABC-анализ")

# ------------------------------------------------------------
# Города, которые будут обрабатываться
# ------------------------------------------------------------
EXPECTED_CITIES = [
    "Иркутск", "Улан-Удэ", "Хабаровск", "Красноярск",
    "Абакан", "Чита", "Кемерово", "Барнаул", "Новокузнецк",
    "Омск", "Томск"
]

# ------------------------------------------------------------
# Вспомогательные функции
# ------------------------------------------------------------
def safe_convert_to_float(series: pd.Series) -> pd.Series:
    """
    Преобразует значения колонки в числа (float).
    Удаляет пробелы, заменяет запятую на точку (русская локаль).
    Пустые значения и ошибки заменяются на 0.
    """
    if series.dtype == object:
        # Убираем все пробелы (обычные и неразрывные) и меняем запятую на точку
        series = series.astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.')
    return pd.to_numeric(series, errors='coerce').fillna(0)


def find_city_columns(df: pd.DataFrame, city: str):
    """
    Ищет в заголовках столбцов те, которые относятся к указанному городу
    и содержат ключевые слова: "остаток", "продажи", "в пути".
    Возвращает словарь с именами найденных колонок или None, если какого-то ключа нет.
    """
    city_lower = city.lower()
    col_map = {}
    keywords = {"остаток": "остаток", "продажи": "продажи", "в пути": "в пути"}
    for key, word in keywords.items():
        candidates = [
            col for col in df.columns
            # Приводим название колонки к строке (на случай чисел в заголовках)
            if city_lower in str(col).lower() and word in str(col).lower()
        ]
        if candidates:
            col_map[key] = candidates[0]  # берём первый подходящий
        else:
            return None  # не хватает данных для города
    return col_map


# ------------------------------------------------------------
# Боковая панель – загрузка файла и настройки
# ------------------------------------------------------------
st.sidebar.header("⚙️ Загрузка данных и настройки")

uploaded_file = st.sidebar.file_uploader(
    "Загрузите Excel-файл с данными", type=["xlsx", "xls"]
)

days_in_report = st.sidebar.number_input(
    "Количество дней в выгрузке продаж (из 1С)",
    min_value=1, max_value=365, value=30, step=1,
    help="За какой период собрана статистика продаж в файле."
)

target_days = st.sidebar.slider(
    "Целевой период обеспечения (дней)",
    min_value=1, max_value=90, value=14, step=1,
    help="На сколько дней вперёд нужно обеспечить запас."
)

st.sidebar.markdown("---")
st.sidebar.info(
    "**Формула расчёта:**\n\n"
    "1. Среднедневные продажи = Продажи / Дни выгрузки\n"
    "2. Потребность = (Среднедневные продажи × Целевой период) – Остаток – В пути\n"
    "3. Если потребность < 0, то ставится 0.\n"
    "4. Округление до целого вверх."
)

# ------------------------------------------------------------
# Основная логика обработки
# ------------------------------------------------------------
if uploaded_file is not None:
    # Чтение файла
    try:
        df_raw = pd.read_excel(uploaded_file, header=0)
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")
        st.stop()

    # Первый столбец считаем названиями товаров
    if df_raw.shape[1] < 2:
        st.error("Файл должен содержать как минимум два столбца: товары и данные по городам.")
        st.stop()

    product_col = df_raw.columns[0]  # исходное название первого столбца
    df = df_raw.rename(columns={product_col: "Товар"}).copy()
    df["Товар"] = df["Товар"].astype(str).fillna("Без названия")

    # Словарь для хранения данных по городам
    city_data = {}

    # Сначала найдём все столбцы для каждого города
    for city in EXPECTED_CITIES:
        cols = find_city_columns(df, city)
        if cols is None:
            st.warning(f"Для города **{city}** не найдены все необходимые столбцы (Остаток, Продажи, В пути). Пропускаем.")
            continue

        # Создаём локальный DataFrame с нужными столбцами
        city_df = df[["Товар", cols["остаток"], cols["продажи"], cols["в пути"]]].copy()
        city_df.columns = ["Товар", "Остаток", "Продажи", "В пути"]

        # Преобразуем числовые колонки
        for col in ["Остаток", "Продажи", "В пути"]:
            city_df[col] = safe_convert_to_float(city_df[col])

        # --------------------------------------------------------
        # Расчёт потребности
        # --------------------------------------------------------
        # Среднедневные продажи
        daily_sales = city_df["Продажи"] / days_in_report
        # Потребность по формуле
        need = (daily_sales * target_days) - city_df["Остаток"] - city_df["В пути"]
        # Обрезаем снизу нулём и округляем вверх до целого
        city_df["К заказу"] = np.ceil(np.maximum(need, 0)).astype(int)
        city_df["Среднедневные продажи"] = daily_sales.round(4)

        # --------------------------------------------------------
        # ABC-анализ внутри города
        # --------------------------------------------------------
        total_sales = city_df["Продажи"].sum()
        if total_sales == 0:
            city_df["Категория ABC"] = "C"
        else:
            # Сортируем по убыванию продаж
            sorted_df = city_df.sort_values("Продажи", ascending=False).copy()
            sorted_df["cum_sales"] = sorted_df["Продажи"].cumsum()
            sorted_df["cum_perc"] = sorted_df["cum_sales"] / total_sales

            # Определяем категорию
            def assign_abc(row):
                if row["Продажи"] == 0:
                    return "C"
                if row["cum_perc"] <= 0.8:
                    return "A"
                elif row["cum_perc"] <= 0.95:
                    return "B"
                else:
                    return "C"

            sorted_df["Категория ABC"] = sorted_df.apply(assign_abc, axis=1)
            # Возвращаем исходный порядок строк
            city_df = city_df.merge(
                sorted_df[["Товар", "Категория ABC"]], on="Товар", how="left"
            )

        # Устанавливаем индекс для быстрого доступа
        city_df.set_index("Товар", inplace=True)
        city_data[city] = city_df

    if not city_data:
        st.error("Не удалось распознать ни одного города. Проверьте названия столбцов в файле.")
        st.stop()

    st.success(f"Успешно загружены данные по {len(city_data)} городам: {', '.join(city_data.keys())}")

    # ------------------------------------------------------------
    # Сводная матрица (для вкладки 1)
    # ------------------------------------------------------------
    all_products = sorted(df["Товар"].unique())
    matrix = pd.DataFrame(index=all_products, columns=list(city_data.keys()))
    for city, cdf in city_data.items():
        matrix[city] = matrix.index.map(lambda p: cdf.loc[p, "К заказу"] if p in cdf.index else 0)

    # ------------------------------------------------------------
    # Генерация Excel для скачивания
    # ------------------------------------------------------------
    def generate_excel(matrix, city_data):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Лист 1: сводная матрица
            matrix.to_excel(writer, sheet_name="Сводная матрица", index=True)
            # Листы по городам
            for city, cdf in city_data.items():
                city_sheet = cdf.reset_index()
                cols_to_save = [
                    "Товар", "Остаток", "Продажи", "В пути",
                    "Среднедневные продажи", "К заказу", "Категория ABC"
                ]
                city_sheet = city_sheet[cols_to_save]
                # Обрезаем длину названия листа (ограничение Excel – 31 символ)
                sheet_name = f"{city}"[:31]
                city_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        return output

    excel_data = generate_excel(matrix, city_data)

    # ------------------------------------------------------------
    # Организация вкладок
    # ------------------------------------------------------------
    tab1, tab2, tab3 = st.tabs(["📋 Сводная матрица", "🏙️ Детализация по городам", "📈 Интерактивные графики"])

    # Вкладка 1: Сводная матрица
    with tab1:
        st.subheader("Общая потребность в дозаказе по городам")
        st.dataframe(matrix, use_container_width=True)
        st.download_button(
            label="📥 Скачать итоговый Excel-файл",
            data=excel_data,
            file_name="ABC_заказ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Вкладка 2: Детализация по конкретному городу
    with tab2:
        st.subheader("Детализация по выбранному городу")
        selected_city = st.selectbox("Выберите город", list(city_data.keys()))
        if selected_city:
            cdf = city_data[selected_city].reset_index()
            display_cols = [
                "Товар", "Остаток", "Продажи", "В пути",
                "Среднедневные продажи", "К заказу", "Категория ABC"
            ]
            st.dataframe(cdf[display_cols], use_container_width=True)
            # Краткая сводка
            total_to_order = cdf["К заказу"].sum()
            st.metric(f"Всего единиц к заказу в г. {selected_city}", total_to_order)

    # Вкладка 3: График продаж по товару
    with tab3:
        st.subheader("Продажи товара по городам")
        selected_product = st.selectbox("Выберите товар", all_products)
        if selected_product:
            # Соберём продажи по всем доступным городам
            sales_data = {}
            for city, cdf in city_data.items():
                if selected_product in cdf.index:
                    sales_data[city] = cdf.loc[selected_product, "Продажи"]
                else:
                    sales_data[city] = 0
            sales_df = pd.DataFrame(list(sales_data.items()), columns=["Город", "Продажи, шт."])
            # Строим столбчатую диаграмму
            fig = px.bar(
                sales_df,
                x="Город",
                y="Продажи, шт.",
                title=f"Продажи: {selected_product}",
                labels={"Продажи, шт.": "Продано за период"},
                text_auto=True,
                color="Продажи, шт.",
                color_continuous_scale="Viridis"
            )
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)

else:
    # Если файл не загружен
    st.info("👈 Загрузите Excel-файл через боковую панель, чтобы начать работу.")
    st.markdown("""
    ### Ожидаемая структура данных
    - Первый столбец – наименования товаров.
    - Далее идут блоки колонок для каждого из 11 городов: **Иркутск, Улан-Удэ, Хабаровск, Красноярск, Абакан, Чита, Кемерово, Барнаул, Новокузнецк, Омск, Томск**.
    - Для каждого города должны присутствовать колонки, содержащие слова **Остаток**, **Продажи**, **В пути** (регистр не важен).
    """)
