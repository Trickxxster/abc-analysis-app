import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="📊 Расчёт заказа и ABC-анализ", layout="wide")
st.title("📊 Автоматизация расчёта заказа и ABC-анализ")

# ------------------------------------------------------------
# Жёсткий порядок городов (из Заголовки.xlsx)
# ------------------------------------------------------------
CITY_ORDER = [
    "Абакан", "Архангельск", "Барнаул", "Иркутск", "Кемерово",
    "Красноярск", "Новокузнецк", "Омск", "Томск", "Хабаровск", "Чита"
]

# ------------------------------------------------------------
# Вспомогательные функции
# ------------------------------------------------------------
def safe_convert_to_float(series: pd.Series) -> pd.Series:
    """Преобразует значения в числа, убирая пробелы и заменяя запятую на точку."""
    if series.dtype == object:
        series = series.astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.')
    return pd.to_numeric(series, errors='coerce').fillna(0)

def find_class_columns(df: pd.DataFrame) -> list:
    """
    Находит индексы столбцов, в которых в первой строке данных стоит класс (A, B или C).
    Возвращает список индексов (начиная с 0) в порядке возрастания.
    """
    class_set = {'A', 'B', 'C', 'a', 'b', 'c'}
    class_cols = []
    first_row = df.iloc[0]  # первая строка данных
    for i in range(2, len(df.columns)):  # начинаем с третьего столбца (индекс 2)
        val = first_row.iloc[i]
        if pd.notna(val):
            s = str(val).strip()
            if s in class_set:
                class_cols.append(i)
    return class_cols

# ------------------------------------------------------------
# Боковая панель
# ------------------------------------------------------------
st.sidebar.header("⚙️ Загрузка данных и настройки")
uploaded_file = st.sidebar.file_uploader("Загрузите Excel-файл с данными (без заголовков)", type=["xlsx", "xls"])
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
        # Читаем файл без заголовков
        df_raw = pd.read_excel(uploaded_file, header=None)

        # Создаём товар из двух первых столбцов (номенклатура + характеристика)
        product_col = df_raw[0].astype(str).fillna('')
        char_col = df_raw[1].astype(str).fillna('')
        df_raw['Товар'] = product_col + ' ' + char_col
        df_raw['Товар'] = df_raw['Товар'].str.strip()

        # Находим столбцы с классами
        class_indices = find_class_columns(df_raw)
        if not class_indices:
            st.error("Не найдено ни одного блока с классом (A/B/C). Проверьте структуру файла.")
            st.stop()

        # Обрабатываем каждый найденный блок
        city_data = {}
        for idx, city in enumerate(CITY_ORDER):
            if idx >= len(class_indices):
                break  # городов в файле меньше, чем в списке
            start_col = class_indices[idx]
            # Проверяем, что блок содержит данные (не все пустые)
            block = df_raw.iloc[:, start_col:start_col+10]  # 10 столбцов показателей
            if block.isnull().all().all():
                # блок полностью пуст – пропускаем
                continue

            # Извлекаем нужные столбцы:
            # 0: Класс, 1: Количество, 2: Остаток, 3: Свободный, 4: Товары в пути,
            # 5: Итоговый остаток, 6: Ср. продажи, 7: Себестоимость, 8: Остаток по себест., 9: Расчет к заказу
            col_ost = start_col + 5   # Итоговый остаток
            col_sales = start_col + 6 # Среднедневные продажи
            col_in_transit = start_col + 4 # Товары в пути

            city_df = df_raw[['Товар', col_ost, col_sales, col_in_transit]].copy()
            city_df.columns = ['Товар', 'Остаток', 'Среднедневные продажи', 'В пути']

            # Преобразуем в числа
            for col in ['Остаток', 'Среднедневные продажи', 'В пути']:
                city_df[col] = safe_convert_to_float(city_df[col])

            # Расчёт потребности
            daily_sales = city_df['Среднедневные продажи']
            need = (daily_sales * target_days) - city_df['Остаток'] - city_df['В пути']
            city_df['К заказу'] = np.ceil(np.maximum(need, 0)).astype(int)

            # Продажи за период для ABC-анализа
            city_df['Продажи за период'] = daily_sales * days_in_report
            total_sales = city_df['Продажи за период'].sum()

            # ABC-анализ
            if total_sales == 0:
                city_df['Категория ABC'] = 'C'
            else:
                sorted_df = city_df.sort_values('Продажи за период', ascending=False).copy()
                sorted_df['cum_sales'] = sorted_df['Продажи за период'].cumsum()
                sorted_df['cum_perc'] = sorted_df['cum_sales'] / total_sales

                def assign_abc(row):
                    if row['Продажи за период'] == 0:
                        return 'C'
                    if row['cum_perc'] <= 0.8:
                        return 'A'
                    elif row['cum_perc'] <= 0.95:
                        return 'B'
                    else:
                        return 'C'

                sorted_df['Категория ABC'] = sorted_df.apply(assign_abc, axis=1)
                city_df = city_df.merge(sorted_df[['Товар', 'Категория ABC']], on='Товар', how='left')

            city_df.set_index('Товар', inplace=True)
            city_data[city] = city_df

        if not city_data:
            st.error("Не удалось обработать ни одного города. Проверьте структуру файла.")
            st.stop()

        st.success(f"Успешно загружены данные по {len(city_data)} городам: {', '.join(city_data.keys())}")

        # Сводная матрица
        all_products = sorted(df_raw['Товар'].unique())
        matrix = pd.DataFrame(index=all_products, columns=list(city_data.keys()))
        for city, cdf in city_data.items():
            matrix[city] = matrix.index.map(lambda p: cdf.loc[p, 'К заказу'] if p in cdf.index else 0)

        # --- Добавляем столбец "Итого" (сумма по строке) ---
        matrix['Итого'] = matrix.sum(axis=1)

        # (Опционально) Можно добавить итоговую строку "Всего" по городам
        # matrix.loc['Всего'] = matrix.sum(axis=0)

        # Генерация Excel-файла (теперь с Итого)
        def generate_excel(matrix, city_data):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                matrix.to_excel(writer, sheet_name='Сводная матрица', index=True)
                for city, cdf in city_data.items():
                    sheet = cdf.reset_index()[['Товар', 'Остаток', 'Среднедневные продажи', 'В пути',
                                               'Продажи за период', 'К заказу', 'Категория ABC']]
                    sheet.to_excel(writer, sheet_name=str(city)[:31], index=False)
            output.seek(0)
            return output

        excel_data = generate_excel(matrix, city_data)

        # Вкладки
        tab1, tab2, tab3 = st.tabs(['📋 Сводная матрица', '🏙️ Детализация по городам', '📈 Интерактивные графики'])

        with tab1:
            st.subheader('Общая потребность в дозаказе по городам')
            st.dataframe(matrix, use_container_width=True)
            st.download_button(
                label='📥 Скачать итоговый Excel-файл',
                data=excel_data,
                file_name='ABC_заказ.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        with tab2:
            st.subheader('Детализация по выбранному городу')
            selected_city = st.selectbox('Выберите город', list(city_data.keys()))
            if selected_city:
                cdf = city_data[selected_city].reset_index()
                st.dataframe(
                    cdf[['Товар', 'Остаток', 'Среднедневные продажи', 'В пути',
                         'Продажи за период', 'К заказу', 'Категория ABC']],
                    use_container_width=True
                )
                st.metric(f'Всего единиц к заказу в г. {selected_city}', cdf['К заказу'].sum())

        with tab3:
            st.subheader('Продажи товара по городам')
            selected_product = st.selectbox('Выберите товар', all_products)
            if selected_product:
                sales_data = {}
                for city, cdf in city_data.items():
                    sales_data[city] = cdf.loc[selected_product, 'Продажи за период'] if selected_product in cdf.index else 0
                sales_df = pd.DataFrame(list(sales_data.items()), columns=['Город', 'Продажи, шт.'])
                fig = px.bar(
                    sales_df, x='Город', y='Продажи, шт.',
                    title=f'Продажи за период: {selected_product}',
                    labels={'Продажи, шт.': 'Продано за период'},
                    text_auto=True, color='Продажи, шт.',
                    color_continuous_scale='Viridis'
                )
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f'Ошибка при обработке файла: {e}')
        st.stop()

else:
    st.info('👈 Загрузите Excel-файл через боковую панель, чтобы начать работу.')
    st.markdown("""
    ### Ожидаемая структура данных (без заголовков)
    - **Столбец A** – номенклатура товара  
    - **Столбец B** – характеристика (модель, цвет и т.п.)  
    - Далее идут блоки по **10 столбцов** для каждого города в строгом порядке:  
      Абакан, Архангельск, Барнаул, Иркутск, Кемерово, Красноярск, Новокузнецк, Омск, Томск, Хабаровск, Чита  
    - Внутри блока: **Класс (A/B/C), Количество, Остаток, Свободный, Товары в пути, Итоговый остаток, Ср. продажи, Себестоимость, Остаток по себест., Расчёт к заказу**  
    - Программа автоматически найдёт начала блоков по столбцам с классами.
    """)
