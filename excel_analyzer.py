"""
Автоматизация Excel: объединение файлов и создание аналитики
Полезно для бухгалтеров, маркетологов, менеджеров по продажам
"""

import pandas as pd
import glob
import os
from datetime import datetime


def merge_and_analyze_excel(folder_path, output_file="аналитика_продаж.xlsx"):
    """
    Объединяет все Excel-файлы из папки, добавляет аналитику и сохраняет отчёт

    Параметры:
    - folder_path: путь к папке с Excel-файлами
    - output_file: имя файла для сохранения результата
    """

    # 1. Находим все Excel-файлы в папке
    all_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

    if not all_files:
        print("❌ Excel-файлы не найдены в указанной папке!")
        return None

    print(f"🔍 Найдено файлов: {len(all_files)}")

    # 2. Объединяем все файлы
    merged_data = []

    for file in all_files:
        # Читаем файл
        df = pd.read_excel(file)

        # Добавляем колонку с именем файла (чтобы знать источник)
        df['источник'] = os.path.basename(file)

        # Пытаемся извлечь месяц из имени файла (если есть дата)
        file_name = os.path.basename(file)
        if 'янв' in file_name.lower() or '01' in file_name:
            df['месяц'] = 'Январь'
        elif 'фев' in file_name.lower() or '02' in file_name:
            df['месяц'] = 'Февраль'
        elif 'мар' in file_name.lower() or '03' in file_name:
            df['месяц'] = 'Март'
        else:
            df['месяц'] = 'Не указан'

        merged_data.append(df)
        print(f"   ✅ Обработан: {os.path.basename(file)} — {len(df)} строк")

    # 3. Объединяем все данные
    result = pd.concat(merged_data, ignore_index=True)
    print(f"\n📊 Всего строк после объединения: {len(result)}")

    # 4. Создаём аналитику
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

        # Лист 1: Все данные
        result.to_excel(writer, sheet_name='Все данные', index=False)

        # Лист 2: Сводка по месяцам (если есть колонка с суммой или количеством)
        if 'сумма' in result.columns or 'sales' in result.columns or 'выручка' in result.columns:
            # Ищем колонку с деньгами
            money_col = None
            for col in ['сумма', 'sales', 'выручка', 'amount', 'total']:
                if col in result.columns:
                    money_col = col
                    break

            if money_col:
                pivot = pd.pivot_table(
                    result,
                    values=money_col,
                    index='месяц',
                    aggfunc=['sum', 'count', 'mean']
                )
                pivot.to_excel(writer, sheet_name='Сводка по месяцам')

        # Лист 3: Статистика по файлам
        file_stats = result.groupby('источник').size().reset_index(name='количество_строк')
        file_stats.to_excel(writer, sheet_name='Статистика по файлам', index=False)

        # Лист 4: Отчёт об обработке
        report = pd.DataFrame({
            'параметр': ['Дата обработки', 'Количество файлов', 'Всего строк', 'Исходная папка'],
            'значение': [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                len(all_files),
                len(result),
                folder_path
            ]
        })
        report.to_excel(writer, sheet_name='Отчёт', index=False)

    print(f"\n✅ Готово! Файл сохранён: {output_file}")
    print(f"   Создано листов: 4 (Все данные, Сводка, Статистика, Отчёт)")

    return result


def create_sample_data():
    """
    Создаёт тестовые данные для демонстрации работы
    """
    # Создаём папку sample_data, если её нет
    if not os.path.exists("sample_data"):
        os.makedirs("sample_data")

    # Январь
    jan_data = pd.DataFrame({
        'дата': ['2025-01-15', '2025-01-20', '2025-01-25'],
        'товар': ['Ноутбук', 'Мышь', 'Клавиатура'],
        'сумма': [50000, 1500, 3000],
        'менеджер': ['Анна', 'Иван', 'Анна']
    })
    jan_data.to_excel("sample_data/январь_продажи.xlsx", index=False)

    # Февраль
    feb_data = pd.DataFrame({
        'дата': ['2025-02-10', '2025-02-18', '2025-02-25', '2025-02-28'],
        'товар': ['Монитор', 'Ноутбук', 'Мышь', 'Клавиатура'],
        'сумма': [25000, 55000, 1500, 3500],
        'менеджер': ['Иван', 'Анна', 'Анна', 'Сергей']
    })
    feb_data.to_excel("sample_data/февраль_продажи.xlsx", index=False)

    # Март
    mar_data = pd.DataFrame({
        'дата': ['2025-03-05', '2025-03-12', '2025-03-20', '2025-03-28'],
        'товар': ['Ноутбук', 'Планшет', 'Мышь', 'Монитор'],
        'сумма': [52000, 30000, 1500, 27000],
        'менеджер': ['Анна', 'Сергей', 'Иван', 'Анна']
    })
    mar_data.to_excel("sample_data/март_продажи.xlsx", index=False)

    print("✅ Созданы тестовые файлы в папке 'sample_data'")
    print("   - январь_продажи.xlsx (3 строки)")
    print("   - февраль_продажи.xlsx (4 строки)")
    print("   - март_продажи.xlsx (4 строки)")


if __name__ == "__main__":
    print("=" * 50)
    print("📊 АВТОМАТИЗАЦИЯ EXCEL — ОБЪЕДИНЕНИЕ ОТЧЁТОВ")
    print("=" * 50)

    print("\n1. Создаю тестовые данные...")
    create_sample_data()

    print("\n2. Объединяю файлы и создаю аналитику...")
    result = merge_and_analyze_excel("sample_data", "итоговая_аналитика.xlsx")

    print("\n" + "=" * 50)
    print("✨ ПРИМЕР РЕЗУЛЬТАТА (первые 5 строк):")
    print("=" * 50)
    if result is not None:
        print(result.head().to_string())