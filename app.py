# app.py — финальная версия (обработка объединённых заголовков городов)
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="📊 Расчёт заказа и ABC-анализ", layout="wide")
st.title("📊 Автоматизация расчёта заказа и ABC-анализ")

# ------------------------------------------------------------
# Города, которые будут обрабатываться (ищем их в заголовках)
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
    """Преобразует значения в числа, убирая пробелы и меняя запятую на точку."""
    if series.dtype == object:
        series = series.astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.')
    return pd.to_numeric(series, errors='coerce').fillna(0)


def extract_city_from_string(s: str):
    """Извлекает название города из строки, если оно там есть (например 'Иркутск Like38' -> 'Иркутск')."""
    s_lower = s.lower()
    for city in EXPECTED_CITIES:
        if city.lower() in s_lower:
            return city
    return None


def process_merged_headers(df_raw: pd.DataFrame, product_col: str):
    """
    Обрабатывает Excel с объединёнными заголовками городов.
    Возвращает DataFrame, где столбцы переименованы в формат 'Город | Показатель'.
    """
    # Список новых названий столбцов
    new_columns = [product_col]  # первый столбец – товар
    current_city = None

    for col in df_raw.columns[1:]:  # пропускаем столбец товара
        col_str = str(col)
        # Проверяем, не является ли значение ячейки названием города
        city = extract_city_from_string(col_str)
        if city:
            current_city = city
        # Формируем новое имя: "Город | Показатель"
        if current_city:
            new_name = f"{current_city} | {col_str}"
        else:
            new_name = col_str  # если город ещё не определён (редкий случай)
        new_columns.append(new_name)

    # Проверяем, что длина списка совпадает
    if len(new_columns) != len(df_raw.columns):
        # Если не совпадает, возвращаем как есть (маловероятно)
        return df_raw

    df_raw.columns = new_columns
    return df_raw


def find_city_columns(df: pd.DataFrame, city: str):
    """Ищет столбцы, относящиеся к городу, по ключевым словам."""
    city_lower = city.lower()
    col_map = {}
    keywords = {"остаток": "остаток", "продажи": "продажи", "в пути": "в пути"}
    for key, word in keywords.items():
        candidates = [
            col for col in df.columns
            if city_lower in str(col).lower() and word in str(col).lower()
        ]
        if candidates:
            col_map[key] = candidates[0]
        else:
            return None
    return col_map


# ------------------------------------------------------------
# Боковая панель
# ------------------------------------------------------------
st.sidebar.header("⚙️ Загрузка данных и настройки")
uploaded_file = st.sidebar.file_uploader("Загрузите Excel-файл с данными", type=["xlsx", "xls"])
days_in_report = st.sidebar.number_input(
    "Количество дней в выгрузке продаж (из 1С)",
    min_value=1, max_value=365, value=30, step=1
)
target_days = st.sidebar.slider(
    "Целевой период обеспечения (дней)",
    min_value=1, max_value=90, value=14, step=1
)
st.sidebar.markdown("---")
st.sidebar.info(
    "**Формула расчёта:**\n\n"
    "1. Среднедневные продажи = Продажи / Дни выгрузки\n"
    "2. Потребность = (Среднедневные продажи × Целевой период) – Остаток – В пути\n"
    "3. Если потребность < 0 → 0\n"
    "4. Округление до целого вверх"
)

# ------------------------------------------------------------
# Основная логика
# ------------------------------------------------------------
if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, header=0)
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")
        st.stop()

    if df_raw.shape[1] < 2:
        st.error("Файл должен содержать как минимум два столбца: товары и данные по городам.")
        st.stop()

    product_col = df_raw.columns[0]
    # Обрабатываем объединённые заголовки городов
    df_raw = process_merged_headers(df_raw, product_col)

    df = df_raw.rename(columns={product_col: "Товар"}).copy()
    df["Товар"] = df["Товар"].astype(str).fillna("Без названия")

    city_data = {}

    for city in EXPECTED_CITIES:
        cols = find_city_columns(df, city)
        if cols is None:
            st.warning(f"Для города **{city}** не найдены все необходимые столбцы. Пропускаем.")
            continue

        city_df = df[["Товар", cols["остаток"], cols["продажи"], cols["в пути"]]].copy()
        city_df.columns = ["Товар", "Остаток", "Продажи", "В пути"]

        for col in ["Остаток", "Продажи", "В пути"]:
            city_df[col] = safe_convert_to_float(city_df[col])

        daily_sales = city_df["Продажи"] / days_in_report
        need = (daily_sales * target_days) - city_df["Остаток"] - city_df["В пути"]
        city_df["К заказу"] = np.ceil(np.maximum(need, 0)).astype(int)
        city_df["Среднедневные продажи"] = daily_sales.round(4)

        # ABC-анализ
        total_sales = city_df["Продажи"].sum()
        if total_sales == 0:
            city_df["Категория ABC"] = "C"
        else:
            sorted_df = city_df.sort_values("Продажи", ascending=False).copy()
            sorted_df["cum_sales"] = sorted_df["Продажи"].cumsum()
            sorted_df["cum_perc"] = sorted_df["cum_sales"] / total_sales

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
            city_df = city_df.merge(sorted_df[["Товар", "Категория ABC"]], on="Товар", how="left")

        city_df.set_index("Товар", inplace=True)
        city_data[city] = city_df

    if not city_data:
        st.error("Не удалось распознать ни одного города. Проверьте названия столбцов в файле.")
        st.stop()

    st.success(f"Успешно загружены данные по {len(city_data)} городам: {', '.join(city_data.keys())}")

    all_products = sorted(df["Товар"].unique())
    matrix = pd.DataFrame(index=all_products, columns=list(city_data.keys()))
    for city, cdf in city_data.items():
        matrix[city] = matrix.index.map(lambda p: cdf.loc[p, "К заказу"] if p in cdf.index else 0)

    def generate_excel(matrix, city_data):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            matrix.to_excel(writer, sheet_name="Сводная матрица", index=True)
            for city, cdf in city_data.items():
                city_sheet = cdf.reset_index()[["Товар", "Остаток", "Продажи", "В пути",
                                                "Среднедневные продажи", "К заказу", "Категория ABC"]]
                city_sheet.to_excel(writer, sheet_name=str(city)[:31], index=False)
        output.seek(0)
        return output

    excel_data = generate_excel(matrix, city_data)

    tab1, tab2, tab3 = st.tabs(["📋 Сводная матрица", "🏙️ Детализация по городам", "📈 Интерактивные графики"])

    with tab1:
        st.subheader("Общая потребность в дозаказе по городам")
        st.dataframe(matrix, use_container_width=True)
        st.download_button(label="📥 Скачать итоговый Excel-файл",
                           data=excel_data,
                           file_name="ABC_заказ.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        st.subheader("Детализация по выбранному городу")
        selected_city = st.selectbox("Выберите город", list(city_data.keys()))
        if selected_city:
            cdf = city_data[selected_city].reset_index()
            st.dataframe(cdf[["Товар", "Остаток", "Продажи", "В пути",
                              "Среднедневные продажи", "К заказу", "Категория ABC"]],
                         use_container_width=True)
            st.metric(f"Всего единиц к заказу в г. {selected_city}", cdf["К заказу"].sum())

    with tab3:
        st.subheader("Продажи товара по городам")
        selected_product = st.selectbox("Выберите товар", all_products)
        if selected_product:
            sales_data = {}
            for city, cdf in city_data.items():
                if selected_product in cdf.index:
                    sales_data[city] = cdf.loc[selected_product, "Продажи"]
                else:
                    sales_data[city] = 0
            sales_df = pd.DataFrame(list(sales_data.items()), columns=["Город", "Продажи, шт."])
            fig = px.bar(sales_df, x="Город", y="Продажи, шт.",
                         title=f"Продажи: {selected_product}",
                         labels={"Продажи, шт.": "Продано за период"},
                         text_auto=True, color="Продажи, шт.",
                         color_continuous_scale="Viridis")
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)

else:
    st.info("👈 Загрузите Excel-файл через боковую панель, чтобы начать работу.")
    st.markdown("""
    ### Ожидаемая структура данных
    - Первый столбец – наименования товаров.
    - Далее идут блоки колонок для каждого города. 
      Заголовки могут быть объединены, но должны содержать название города (например, «Иркутск Like38»).
    - Внутри блока города обязательно должны быть колонки со словами **Остаток**, **Продажи**, **В пути**.
    """)
