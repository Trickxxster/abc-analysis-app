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
    """Извлекает название города из строки, если оно там есть."""
    s_lower = s.lower()
    for city in EXPECTED_CITIES:
        if city.lower() in s_lower:
            return city
    return None

def find_city_columns(df: pd.DataFrame, city: str):
    """
    Ищет столбцы для города по уточнённым ключевым словам.
    Возвращает словарь {'остаток': col, 'продажи': col, 'в пути': col}
    """
    city_lower = city.lower()
    # Уточнённые ключевые слова для поиска
    patterns = {
        'остаток': 'Итоговый остаток',
        'продажи': 'Ср. продажа в день за период',
        'в пути': 'Товары в пути'
    }
    col_map = {}
    for key, pattern in patterns.items():
        candidates = [
            col for col in df.columns
            if city_lower in str(col).lower() and pattern.lower() in str(col).lower()
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
    "1. Среднедневные продажи берутся напрямую из столбца 'Ср. продажа в день за период, шт'\n"
    "2. Потребность = (Среднедневные продажи × Целевой период) – Остаток – В пути\n"
    "3. Если потребность < 0 → 0\n"
    "4. Округление до целого вверх"
)

# ------------------------------------------------------------
# Основная логика
# ------------------------------------------------------------
if uploaded_file is not None:
    try:
        # Читаем первые две строки как заголовки (без данных)
        header_rows = pd.read_excel(uploaded_file, header=None, nrows=2)
        cities_row = header_rows.iloc[0].values
        metrics_row = header_rows.iloc[1].values

        # Читаем данные, начиная с 3-й строки (индекс 2 в pandas)
        data_df = pd.read_excel(uploaded_file, header=None, skiprows=2)

        # Формируем новые имена столбцов
        new_columns = []
        current_city = None
        for j in range(len(cities_row)):
            city_val = str(cities_row[j]) if pd.notna(cities_row[j]) else ''
            metric_val = str(metrics_row[j]) if pd.notna(metrics_row[j]) else ''

            # Для первого столбца – это номенклатура
            if j == 0:
                new_columns.append("Товар")
                continue

            # Пытаемся извлечь город из значения
            extracted_city = extract_city_from_string(city_val)
            if extracted_city:
                current_city = extracted_city

            if current_city:
                # Если метрика пустая – оставляем только город (редко)
                col_name = f"{current_city} | {metric_val}" if metric_val else current_city
            else:
                col_name = metric_val

            new_columns.append(col_name)

        # Проверяем, что длина совпадает
        if len(new_columns) != data_df.shape[1]:
            st.error("Не совпадает количество столбцов после обработки заголовков.")
            st.stop()

        data_df.columns = new_columns
        df = data_df.copy()

        # Первый столбец – товар
        df["Товар"] = df["Товар"].astype(str).fillna("Без названия")

        # Обработка по городам
        city_data = {}
        for city in EXPECTED_CITIES:
            cols = find_city_columns(df, city)
            if cols is None:
                st.warning(f"Для города **{city}** не найдены все необходимые столбцы. Пропускаем.")
                continue

            city_df = df[["Товар", cols["остаток"], cols["продажи"], cols["в пути"]]].copy()
            city_df.columns = ["Товар", "Остаток", "Среднедневные продажи", "В пути"]

            for col in ["Остаток", "Среднедневные продажи", "В пути"]:
                city_df[col] = safe_convert_to_float(city_df[col])

            # Используем среднедневные продажи напрямую (без деления на дни)
            daily_sales = city_df["Среднедневные продажи"]
            need = (daily_sales * target_days) - city_df["Остаток"] - city_df["В пути"]
            city_df["К заказу"] = np.ceil(np.maximum(need, 0)).astype(int)

            # ABC-анализ по общим продажам (здесь продажи = среднедневные × дни отчёта)
            city_df["Продажи за период"] = daily_sales * days_in_report
            total_sales = city_df["Продажи за период"].sum()
            if total_sales == 0:
                city_df["Категория ABC"] = "C"
            else:
                sorted_df = city_df.sort_values("Продажи за период", ascending=False).copy()
                sorted_df["cum_sales"] = sorted_df["Продажи за период"].cumsum()
                sorted_df["cum_perc"] = sorted_df["cum_sales"] / total_sales

                def assign_abc(row):
                    if row["Продажи за период"] == 0:
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
                    city_sheet = cdf.reset_index()[["Товар", "Остаток", "Среднедневные продажи", "В пути",
                                                    "Продажи за период", "К заказу", "Категория ABC"]]
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
                st.dataframe(cdf[["Товар", "Остаток", "Среднедневные продажи", "В пути",
                                  "Продажи за период", "К заказу", "Категория ABC"]],
                             use_container_width=True)
                st.metric(f"Всего единиц к заказу в г. {selected_city}", cdf["К заказу"].sum())

        with tab3:
            st.subheader("Продажи товара по городам")
            selected_product = st.selectbox("Выберите товар", all_products)
            if selected_product:
                sales_data = {}
                for city, cdf in city_data.items():
                    if selected_product in cdf.index:
                        sales_data[city] = cdf.loc[selected_product, "Продажи за период"]
                    else:
                        sales_data[city] = 0
                sales_df = pd.DataFrame(list(sales_data.items()), columns=["Город", "Продажи, шт."])
                fig = px.bar(sales_df, x="Город", y="Продажи, шт.",
                             title=f"Продажи за период: {selected_product}",
                             labels={"Продажи, шт.": "Продано за период"},
                             text_auto=True, color="Продажи, шт.",
                             color_continuous_scale="Viridis")
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
        st.stop()

else:
    st.info("👈 Загрузите Excel-файл через боковую панель, чтобы начать работу.")
    st.markdown("""
    ### Ожидаемая структура данных
    - Первый столбец – наименования товаров.
    - Первая строка – названия городов (объединённые ячейки).
    - Вторая строка – названия показателей (Остаток, Продажи, В пути).
    - Данные начинаются с третьей строки.
    """)
