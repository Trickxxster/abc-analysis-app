import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="📊 Расчёт заказа и ABC-анализ", layout="wide")
st.title("📊 Автоматизация расчёта заказа и ABC-анализ (с ручной корректировкой)")

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
    if series.dtype == object:
        series = series.astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.')
    return pd.to_numeric(series, errors='coerce').fillna(0)

def find_class_columns(df: pd.DataFrame) -> list:
    class_set = {'A', 'B', 'C', 'a', 'b', 'c'}
    class_cols = []
    for i in range(2, len(df.columns)):
        col_data = df.iloc[:, i]
        found = False
        for val in col_data:
            if pd.notna(val):
                s = str(val).strip()
                if s in class_set:
                    found = True
                    break
        if found:
            class_cols.append(i)
    return class_cols

def recalculate_city_data(df_raw, class_indices, target_days, days_in_report, manual_overrides):
    """Пересчитывает все города с учётом ручных правок."""
    city_data = {}
    for idx, city in enumerate(CITY_ORDER):
        if idx >= len(class_indices):
            break
        start_col = class_indices[idx]
        block = df_raw.iloc[:, start_col:start_col+10]
        if block.isnull().all().all():
            continue

        col_ost = start_col + 5
        col_sales = start_col + 6
        col_in_transit = start_col + 4

        city_df = df_raw[['Товар', col_ost, col_sales, col_in_transit]].copy()
        city_df.columns = ['Товар', 'Остаток', 'Среднедневные продажи', 'В пути']

        for col in ['Остаток', 'Среднедневные продажи', 'В пути']:
            city_df[col] = safe_convert_to_float(city_df[col])

        daily_sales = city_df['Среднедневные продажи']
        need = (daily_sales * target_days) - city_df['Остаток'] - city_df['В пути']
        raw_need = np.ceil(np.maximum(need, 0)).astype(int)

        # Применяем ручные правки
        order_col = []
        for idx_row, product in enumerate(city_df['Товар']):
            key = (city, product)
            if key in manual_overrides and pd.notna(manual_overrides[key]):
                order_col.append(int(manual_overrides[key]))
            else:
                order_col.append(int(raw_need[idx_row]))
        city_df['К заказу'] = order_col

        # Продажи за период для ABC
        city_df['Продажи за период'] = daily_sales * days_in_report
        total_sales = city_df['Продажи за период'].sum()

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
        city_df = city_df[['Остаток', 'Среднедневные продажи', 'В пути',
                           'Продажи за период', 'К заказу', 'Категория ABC']]
        city_data[city] = city_df
    return city_data

def build_matrix(city_data):
    all_products = sorted({p for cdf in city_data.values() for p in cdf.index})
    matrix = pd.DataFrame(index=all_products, columns=list(city_data.keys()))
    for city, cdf in city_data.items():
        matrix[city] = matrix.index.map(lambda p: cdf.loc[p, 'К заказу'] if p in cdf.index else 0)
    matrix['Итого'] = matrix.sum(axis=1)
    return matrix

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

# Кнопка сброса ручных правок
reset_overrides = st.sidebar.button("🔄 Сбросить все ручные правки")

st.sidebar.markdown("---")
st.sidebar.info(
    "**Формула расчёта:**\n\n"
    "1. Среднедневные продажи берутся из столбца 'Ср. продажа в день за период, шт'\n"
    "2. Потребность = (Среднедневные продажи × Целевой период) – Остаток – В пути\n"
    "3. Если потребность < 0 → 0\n"
    "4. Округление до целого вверх\n\n"
    "💡 Вы можете **редактировать** значения «К заказу» во вкладке «Детализация по городам». "
    "Правки сбрасываются при изменении слайдера или количества дней, а также по кнопке сброса."
)

# ------------------------------------------------------------
# Основная логика
# ------------------------------------------------------------
if uploaded_file is not None:
    try:
        # Читаем файл
        df_raw = pd.read_excel(uploaded_file, header=None)
        product_col = df_raw[0].astype(str).fillna('')
        char_col = df_raw[1].astype(str).fillna('')
        df_raw['Товар'] = (product_col + ' ' + char_col).str.strip()

        class_indices = find_class_columns(df_raw)
        if not class_indices:
            st.error("Не найдено ни одного блока с классом (A/B/C). Проверьте структуру файла.")
            st.stop()

        # Инициализация session_state
        if 'manual_overrides' not in st.session_state:
            st.session_state.manual_overrides = {}
        if 'city_data' not in st.session_state:
            st.session_state.city_data = None
        if 'matrix' not in st.session_state:
            st.session_state.matrix = None
        if 'last_target_days' not in st.session_state:
            st.session_state.last_target_days = target_days
        if 'last_days_in_report' not in st.session_state:
            st.session_state.last_days_in_report = days_in_report

        # Проверяем, изменился ли слайдер или дни – если да, сбрасываем правки
        if (target_days != st.session_state.last_target_days or 
            days_in_report != st.session_state.last_days_in_report):
            st.session_state.manual_overrides = {}
            st.session_state.last_target_days = target_days
            st.session_state.last_days_in_report = days_in_report

        # Если нажата кнопка сброса – очищаем правки
        if reset_overrides:
            st.session_state.manual_overrides = {}

        # Пересчитываем данные с учётом правок (если они изменились или первый раз)
        city_data = recalculate_city_data(
            df_raw, class_indices, target_days, days_in_report, st.session_state.manual_overrides
        )
        matrix = build_matrix(city_data)

        # Сохраняем в session_state для использования в других частях
        st.session_state.city_data = city_data
        st.session_state.matrix = matrix

        if not city_data:
            st.error("Не удалось обработать ни одного города. Проверьте структуру файла.")
            st.stop()

        st.success(f"Успешно загружены данные по {len(city_data)} городам: {', '.join(city_data.keys())}")

        # Вкладки
        tab1, tab2, tab3 = st.tabs(['📋 Сводная матрица', '🏙️ Детализация по городам (с редактированием)', '📈 Интерактивные графики'])

        with tab1:
            st.subheader('Общая потребность в дозаказе по городам')
            st.dataframe(matrix, use_container_width=True)

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
            st.download_button(
                label='📥 Скачать итоговый Excel-файл',
                data=excel_data,
                file_name='ABC_заказ.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        with tab2:
            st.subheader('Детализация по выбранному городу (редактируйте «К заказу»)')
            selected_city = st.selectbox('Выберите город', list(city_data.keys()))

            if selected_city:
                cdf = city_data[selected_city].reset_index()
                display_cols = ['Товар', 'Остаток', 'Среднедневные продажи', 'В пути',
                                'Продажи за период', 'К заказу', 'Категория ABC']
                edit_df = cdf[display_cols].copy()

                edited_df = st.data_editor(
                    edit_df,
                    column_config={
                        "К заказу": st.column_config.NumberColumn(
                            "К заказу",
                            help="Введите нужное количество. Изменения сохраняются до сброса или изменения слайдера.",
                            min_value=0,
                            step=1
                        )
                    },
                    disabled=['Товар', 'Остаток', 'Среднедневные продажи', 'В пути', 'Продажи за период', 'Категория ABC'],
                    use_container_width=True,
                    key=f"editor_{selected_city}"
                )

                # Обновляем ручные правки на основе изменений в редакторе
                # Сравниваем edited_df с исходным cdf (до редактирования)
                # Замечаем, что cdf уже содержит текущие значения (с учётом правок), 
                # поэтому сравнение нужно делать с расчётными значениями? 
                # Мы будем обновлять правки для каждой строки, где значение отличается от расчётного.
                # Для этого нужно знать расчётное значение. Его можно получить из пересчитанного city_data, 
                # но проще сохранить расчётное значение в отдельной колонке при формировании.
                # В нашей функции recalculate_city_data мы не сохраняем расчётное значение отдельно.
                # Поэтому мы будем проверять: если значение в edited_df отличается от значения в cdf (которое уже с учётом правок) - это изменение.
                # Но если пользователь изменил значение, оно будет отличаться от cdf.
                # Если же пользователь вернул значение обратно к расчётному, оно станет равно cdf, и мы удалим правку.
                # Для этого нужно знать расчётное значение. Переделаем recalculate_city_data так, чтобы возвращать и расчётное значение.
                # Вместо этого пересчитаем заново city_data без применения правок и сравним.
                # Это сложно. Проще при каждом изменении редактора обновлять manual_overrides для всех строк,
                # записывая введённое значение как правку, даже если оно совпадает с расчётным – это допустимо.
                # При сбросе или изменении слайдера правки будут очищены.
                # Единственный минус – если пользователь ввел то же число, что и расчётное, оно всё равно будет считаться правкой.
                # При следующем пересчёте (без изменения слайдера) оно останется таким же, что не страшно.
                for idx, row in edited_df.iterrows():
                    product = row['Товар']
                    new_val = row['К заказу']
                    if pd.isna(new_val):
                        key = (selected_city, product)
                        if key in st.session_state.manual_overrides:
                            del st.session_state.manual_overrides[key]
                    else:
                        st.session_state.manual_overrides[(selected_city, product)] = int(new_val)

                # После обновления правок пересчитываем данные и матрицу
                # Обновляем city_data и matrix в session_state
                new_city_data = recalculate_city_data(
                    df_raw, class_indices, target_days, days_in_report, st.session_state.manual_overrides
                )
                new_matrix = build_matrix(new_city_data)
                st.session_state.city_data = new_city_data
                st.session_state.matrix = new_matrix
                # Обновляем переменные для текущего отображения
                city_data = new_city_data
                matrix = new_matrix

                total_order = cdf['К заказу'].sum() if 'К заказу' in cdf else 0
                st.metric(f'Всего единиц к заказу в г. {selected_city}', total_order)

        with tab3:
            st.subheader('Продажи товара по городам')
            # Берём список товаров из matrix (все уникальные)
            all_products = sorted(matrix.index.tolist())
            selected_product = st.selectbox('Выберите товар', all_products, key='product_select')
            if selected_product:
                # Используем city_data из session_state (актуальное)
                sales_data = {}
                for city, cdf in st.session_state.city_data.items():
                    if selected_product in cdf.index:
                        sales_data[city] = cdf.loc[selected_product, 'Продажи за период']
                    else:
                        sales_data[city] = 0
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
