import pandas as pd
import numpy as np
import os
import re
from datetime import datetime

# === НАСТРОЙКИ ===
FOLDER_PATH = r"C:\Users\rosto\Documents\Маркетплейсы\Выгрузки из SF\Категории"  # Windows-путь с raw-строкой

# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===
def parse_filename(filename):
    """Извлекаем дату начала из имени файла (категория)."""
    date_match = re.search(r"(\d{4}-\d{2}-\d{2})\s+(\d{4}-\d{2}-\d{2})", filename)
    start_date = None
    if date_match:
        start_date = datetime.strptime(date_match.group(1), "%Y-%m-%d")
    return start_date

# === ЗАГРУЗКА И ОБЪЕДИНЕНИЕ ===
all_data = []

# Рекурсивный обход всех подпапок
for root, dirs, files in os.walk(FOLDER_PATH):
    for fname in files:
        if fname.endswith(".xlsx"):
            file_path = os.path.join(root, fname)
            date = parse_filename(fname)

            df = pd.read_excel(file_path, dtype={"SKU": str})  # читаем SKU как строку

            # Только нужные столбцы
            required_cols = ['SKU', 'Товар', 'Цена', 'Отзывы', 'Выручка за 7 дн', 'Категория', 'Продавец', 'Текущий остаток (шт)' ]
            existing_cols = [col for col in required_cols if col in df.columns]
            df = df[existing_cols].copy()

            df['Дата'] = date
            all_data.append(df)

# Объединяем всё в один датафрейм
full_df = pd.concat(all_data, ignore_index=True)

# === ФИЛЬТРАЦИЯ ===
filtered_df = full_df[full_df['Отзывы'] < 200].copy()

# === АГРЕГАЦИЯ ДЛЯ УСТРАНЕНИЯ ДУБЛЕЙ И ДОБАВЛЕНИЕ ОСТАТКОВ ===
aggregated = (
    filtered_df
    .groupby(['SKU', 'Дата'], as_index=False)
    .agg({
        'Цена': 'mean',
        'Выручка за 7 дн': 'sum',
        'Товар': 'first',
        'Категория': 'first',
        'Продавец': 'first',
        'Отзывы': 'max',
        'Текущий остаток (шт)': 'mean',
        
    })
)

# === СВОДНЫЕ ТАБЛИЦЫ (Цены, Обороты, Остатки) ===
pivot_price = aggregated.pivot(index='SKU', columns='Дата', values='Цена')
pivot_revenue = aggregated.pivot(index='SKU', columns='Дата', values='Выручка за 7 дн')

# Остатки по неделям
pivot_stock = aggregated.pivot(index='SKU', columns='Дата', values='Текущий остаток (шт)')
stock_cols_sorted = sorted(pivot_stock.columns.tolist())
stock_rename = {col: f"Остаток Week -{len(stock_cols_sorted)-i}" for i, col in enumerate(stock_cols_sorted)}
pivot_stock.rename(columns=stock_rename, inplace=True)

# === ПЕРЕИМЕНОВАНИЕ КОЛОНОК ДО ОБЪЕДИНЕНИЯ ===
price_cols_sorted = sorted(pivot_price.columns.tolist())
revenue_cols_sorted = sorted(pivot_revenue.columns.tolist())

price_rename = {col: f"Цена Week -{len(price_cols_sorted)-i}" for i, col in enumerate(price_cols_sorted)}
revenue_rename = {col: f"Оборот Week -{len(revenue_cols_sorted)-i}" for i, col in enumerate(revenue_cols_sorted)}

pivot_price.rename(columns=price_rename, inplace=True)
pivot_revenue.rename(columns=revenue_rename, inplace=True)

# === ДОП ИНФОРМАЦИЯ ===
info_cols = ['SKU', 'Товар']
if 'Категория' in filtered_df.columns:
    info_cols.append('Категория')
if 'Продавец' in filtered_df.columns:
    info_cols.append('Продавец')
base_info = aggregated.drop_duplicates(subset='SKU')[info_cols].set_index('SKU')

# === ОБЪЕДИНЕНИЕ С OUTER JOIN ===
combined = base_info.join(pivot_price, how="outer").join(pivot_revenue, how="outer").join(pivot_stock, how="outer")

# === ЗАПОЛНЕНИЕ ПРОПУСКОВ НУЛЯМИ ТОЛЬКО ДЛЯ ЧИСЕЛ (Цена, Оборот, Остаток) ===
num_cols = [col for col in combined.columns if col.startswith('Цена') or col.startswith('Оборот') or col.startswith('Остаток')]
combined[num_cols] = combined[num_cols].fillna(0)

# === ФИЛЬТРАЦИЯ: ОСТАВИТЬ ТОЛЬКО ТЕ, КТО ПРОДАВАЛСЯ ДВЕ ПОСЛЕДНИЕ НЕДЕЛИ ===
last_two_revenue_cols = [f"Оборот Week -{i}" for i in [2, 1] if f"Оборот Week -{i}" in combined.columns]
if len(last_two_revenue_cols) == 2:
    combined = combined[(combined[last_two_revenue_cols[0]] > 0) & (combined[last_two_revenue_cols[1]] > 0)]

# === РАСЧЕТЫ ===
# Динамика оборота: последние 2 недели
last_rev_2, last_rev_1 = last_two_revenue_cols if len(last_two_revenue_cols) == 2 else (None, None)
if last_rev_1 and last_rev_2:
    combined['Динамика оборота (%)'] = ((combined[last_rev_1] - combined[last_rev_2]) / combined[last_rev_2] * 100).round(2)

# Динамика цены: стандартное отклонение от всех ненулевых значений цены
price_week_cols = [col for col in combined.columns if col.startswith("Цена Week")]
combined['Динамика цены (%)'] = combined[price_week_cols].replace(0, np.nan).std(axis=1).round(2)


# === РЕЗУЛЬТАТ ===
combined.reset_index(inplace=True)

# СОРТИРОВКА ПО ДИНАМИКЕ ОБОРОТА
if 'Динамика оборота (%)' in combined.columns:
    combined.sort_values(by='Динамика оборота (%)', ascending=False, inplace=True)

# Сохраняем как Excel в ту же папку, где исходные файлы
output_path = os.path.join(FOLDER_PATH, "итоговая_таблица.xlsx")
combined.to_excel(output_path, index=False)

# === ВЫЯВЛЕНИЕ НОВЫХ УСПЕШНЫХ ТОВАРОВ ПО ПРОДАВЦАМ ===
# Порог успешности
min_revenue = 15000

# Определяем недели оборота
revenue_weeks = sorted([col for col in combined.columns if col.startswith("Оборот Week")], key=lambda x: int(x.split('-')[-1]))
if len(revenue_weeks) >= 2:
    last_week = revenue_weeks[-1]
    prev_week = revenue_weeks[-2]

    # Фильтрация успешных новых товаров: продажи только в последнюю неделю, и выше чем в предыдущую
    is_success = (combined[last_week] > min_revenue) & (combined[last_week] > combined[prev_week])

    new_successful = combined[is_success]

    # Группировка по продавцам
    top_sellers = new_successful.groupby('Продавец').size().reset_index(name='Новых успешных товаров')
    top_sellers = top_sellers.sort_values(by='Новых успешных товаров', ascending=False)

    # Сохраняем как отдельный файл
    top_sellers.to_excel(os.path.join(FOLDER_PATH, "топ_продавцов_по_новым_товарам.xlsx"), index=False)