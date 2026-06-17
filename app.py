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
    """Преобразует значения в числа, убирая пробелы и заменяя запятую на точку."""
    if series.dtype == object:
        series = series.astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.')
    return pd.to_numeric(series, errors='coerce').fillna(0)

def find_class_columns(df: pd.DataFrame) -> list:
    """
    Находит индексы столбцов, в которых в любой строке данных стоит класс (A, B или C).
    Возвращает список индексов (начиная с 0) в порядке возрастания.
    """
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
keep_overrides = st.sidebar.checkbox(
    "Сохранять ручные правки при пересчёте",
    value=True,
    help="Если включено, ваши изменения в столбце «К заказу» сохраняются и переопределяют расчёт. Если выключено, правки сбрасываются и используются только расчётные значения."
)
if st.sidebar.button("Сбросить все ручные правки"):
    st.session_state.manual_overrides = {}
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.info(
    "**Формула расчёта:**\n\n"
    "1. Среднедневные продажи берутся напрямую из столбца 'Ср. продажа в день за период, шт'\n"
    "2. Потребность = (Среднедневные продажи × Целевой период) – Остаток – В пути\n"
    "3. Если потребность < 0 → 0\n"
    "4. Округление до целого вверх\n\n"
    "💡 Вы можете **редактировать** значения «К заказу» во вкладке «Детализация по городам». "
    "Отредактированные значения сохраняются и используются в сводной матрице и при выгрузке."
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

        # Находим столбцы с классами (по всем строкам)
        class_indices = find_class_columns(df_raw)
        if not class_indices:
            st.error("Не найдено ни одного блока с классом (A/B/C). Проверьте структуру файла.")
            st.stop()

        # Инициализируем хранение ручных правок в session_state
        if 'manual_overrides' not in st.session_state:
            st.session_state.manual_overrides = {}

        # Функция пересчёта всех городов (возвращает city_data)
        def recalculate_city_data(target_days, days_in_report, keep_overrides):
            city_data = {}
            # Если режим сохранения правок выключен, очищаем все правки
            if not keep_overrides:
                st.session_state.manual_overrides = {}

            overrides = st.session_state.manual_overrides
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
                city_df['К заказу (расчёт)'] = raw_need

                # Применяем ручные правки, если режим включён
                if keep_overrides:
                    for idx_row, product in enumerate(city_df['Товар']):
                        key = (city, product)
                        if key in overrides and pd.notna(overrides[key]):
                            city_df.loc[idx_row, 'К заказу'] = overrides[key]
                        else:
                            city_df.loc[idx_row, 'К заказу'] = raw_need[idx_row]
                else:
                    # Режим выключен – используем только расчёт
                    city_df['К заказу'] = raw_need

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

        # Пересчитываем данные
        city_data = recalculate_city_data(target_days, days_in_report, keep_overrides)

        if not city_data:
            st.error("Не удалось обработать ни одного города. Проверьте структуру файла.")
            st.stop()

        # Функция построения сводной матрицы с учётом правок
        def build_matrix(city_data):
            all_products = sorted(df_raw['Товар'].unique())
            matrix = pd.DataFrame(index=all_products, columns=list(city_data.keys()))
            for city, cdf in city_data.items():
                matrix[city] = matrix.index.map(lambda p: cdf.loc[p, 'К заказу'] if p in cdf.index else 0)
            matrix['Итого'] = matrix.sum(axis=1)
            return matrix

        matrix = build_matrix(city_data)

        st.success(f"Успешно загружены данные по {len(city_data)} городам: {', '.join(city_data.keys())}")

        # --- Вкладки ---
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
                # Получаем DataFrame для города и сбрасываем индекс
                cdf = city_data[selected_city].reset_index()
                display_cols = ['Товар', 'Остаток', 'Среднедневные продажи', 'В пути',
                                'Продажи за период', 'К заказу', 'Категория ABC']
                edit_df = cdf[display_cols].copy()

                # Используем data_editor с разрешением редактирования только колонки "К заказу"
                edited_df = st.data_editor(
                    edit_df,
                    column_config={
                        "К заказу": st.column_config.NumberColumn(
                            "К заказу",
                            help="Введите нужное количество. Изменения сохраняются и влияют на сводную матрицу.",
                            min_value=0,
                            step=1
                        )
                    },
                    disabled=['Товар', 'Остаток', 'Среднедневные продажи', 'В пути', 'Продажи за период', 'Категория ABC'],
                    use_container_width=True,
                    key=f"editor_{selected_city}"
                )

                # Обновляем ручные правки на основе изменений в редакторе
                # Сравниваем с исходным cdf, чтобы найти изменённые ячейки
                # Для простоты обновляем все значения, которые были изменены
                # Но важно не перезаписывать правки, если они не изменились
                # Мы можем просто сохранить все значения из edited_df как правки
                # (даже если они совпадают с расчётными – это нормально)
                # Однако нужно удалять правку, если пользователь очистил ячейку (NaN)
                overrides = st.session_state.manual_overrides
                for idx, row in edited_df.iterrows():
                    product = row['Товар']
                    new_val = row['К заказу']
                    key = (selected_city, product)
                    if pd.isna(new_val):
                        # Удаляем правку
                        if key in overrides:
                            del overrides[key]
                    else:
                        # Сохраняем как правку (даже если равно расчётному)
                        overrides[key] = int(new_val)

                # Если режим сохранения правок выключен, принудительно очищаем все правки
                if not keep_overrides:
                    st.session_state.manual_overrides = {}
                    # И пересчитываем данные, чтобы сбросить все к расчётным
                    city_data = recalculate_city_data(target_days, days_in_report, keep_overrides)
                    matrix = build_matrix(city_data)
                    st.rerun()

                # В любом случае, после редактирования пересчитываем данные,
                # чтобы обновить матрицу и другие вкладки
                # Но чтобы не было бесконечного цикла, вызываем rerun только если были изменения
                # Для простоты будем вызывать всегда, но проверяем, изменились ли правки
                # Можно сравнить текущие правки с предыдущими, но для демонстрации оставим так:
                city_data = recalculate_city_data(target_days, days_in_report, keep_overrides)
                matrix = build_matrix(city_data)
                # Обновляем сессионные переменные (если нужно)
                st.session_state['city_data'] = city_data
                st.session_state['matrix'] = matrix
                # Вызов rerun для обновления интерфейса (если изменения были)
                # Но мы не можем вызывать rerun внутри блока обработки редактора, это приведёт к циклу.
                # Вместо этого мы можем использовать st.rerun() только если были изменения,
                # но проще положиться на то, что Streamlit сам перерисует при следующем взаимодействии.
                # Однако после редактирования данные в матрице обновятся только после следующего rerun.
                # Чтобы обновить их сразу, можно использовать st.rerun().
                # Но чтобы избежать бесконечного цикла, добавим условие: если были изменения в правках.
                # Просто проверим, изменилось ли что-то в overrides по сравнению с предыдущим состоянием.
                # Для упрощения: будем считать, что любое редактирование вызывает rerun.
                # Это нормально, т.к. пользователь ожидает обновления.
                # Но нужно избежать rerun, если ничего не изменилось.
                # Мы можем сохранить предыдущие правки и сравнить.
                # Для надёжности добавим флаг в session_state.
                if 'last_overrides_hash' not in st.session_state:
                    st.session_state.last_overrides_hash = hash(frozenset(overrides.items()))
                current_hash = hash(frozenset(overrides.items()))
                if current_hash != st.session_state.last_overrides_hash:
                    st.session_state.last_overrides_hash = current_hash
                    st.rerun()

                # Отображаем метрику
                total_order = cdf['К заказу'].sum() if 'К заказу' in cdf else 0
                st.metric(f'Всего единиц к заказу в г. {selected_city}', total_order)

        with tab3:
            st.subheader('Продажи товара по городам')
            all_products = sorted(df_raw['Товар'].unique())
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
