import pandas as pd
import sqlite3 as sq
import os
import warnings
warnings.simplefilter("ignore", UserWarning)

import dictionaries as dct
import parsers as prs

# 🔹 Ввод отчета
valid_reports = dct.report_paths.keys()
while True:
    report_name = input(f"Укажите название отчета {list(valid_reports)}: ").strip()
    if report_name in valid_reports:
        break
    print("❗ Неверное значение. Укажите правильное имя отчёта.")

# 🔹 Выбор периода обновления
while True:
    if report_name == "StockCRM":
        update_period = input("Обновлять за год, за месяц или за дату? (пишите 'Y' для года, 'M' для месяца, 'D' для даты): ").lower()
        if update_period in ["y", "m", "d"]:
            break
        print("❗ Неверное значение. Укажите 'Y','M' или 'D'.")
    else:
        update_period = input("Обновлять за год или за месяц? ('Y' для года, 'M' для месяца): ").lower()
        if update_period in ["y", "m"]:
            break
        print("❗ Неверное значение. Укажите 'Y' или 'M'.")

# 🔹 Инициализация переменных
year = month = date = None

# 🔹 Определяем нужные переменные по выбору
# Обновление по "Год"
if update_period == "y":
    while True:
        year = input("Введите год (например: 2025): ").strip()
        if len(year) == 4:
            break
        else:
            print("❗ Неверное значение. Укажите значение правильно.")
# Обновление по "Месяц"
elif update_period == "m":
    while True:
        year = input("Введите год (например: 2025): ").strip()
        if len(year) == 4:
            break
        else:
            print("❗ Неверное значение. Укажите значение правильно.")
    while True:
        month = input("Введите номер месяца (например: 01): ").strip()
        if 2 >= len(month) > 0:
            break
        else:
            print("❗ Неверное значение. Укажите значение правильно.")
# Обновление по "Дата"
elif update_period == "d":
    while True:
        date = input("Введите дату в формате ДД.ММ.ГГГГ (например: 06.11.2024 или 23.06.2025): ").replace("/", ".").replace("-", ".")
        if len(date) == 10:
            break
        else:
            print("❗ Неверное значение. Укажите формат даты правильно.")



def db_table_updater(report_name, year=None, month=None, date=None):
    conn = sq.connect("../../02_Database/01_Sales_DB/Sales.db")
    cursor = conn.cursor()

    # 🔹 Путь к папке с файлами
    base_path = "../../01_Data_Source/01_Exported_data"
    end_folder = dct.report_paths.get(report_name)
    folder_path = os.path.join(base_path, end_folder)

    # 🔹 Собираем все Excel-файлы из папки и подпапок
    files = []
    for root, dirs, file_names in os.walk(folder_path):
        for f in file_names:
            if f.lower().endswith((".xls", ".xlsx")):
                files.append(os.path.join(root, f))

    # 🔹 Универсальная фильтрация по году/месяцу
    filtered_files = []
    for f in files:
        fname = os.path.basename(f)
        if report_name == "StockCRM":
            file_year = fname[-9:-5]       # "2023"
            file_month = fname[-12:-10]    # "08"
            file_day = fname[-15:-13]      # '07'
        else:
            file_year = fname[0:4]         # "2023"
            file_month = fname[5:7]        # "01"
            file_day = fname[8:10]

        file_date = f"{file_day}.{file_month}.{file_year}"

        if date:  # если указана дата
            if file_date == date:
                filtered_files.append(f)
        elif year and month:
            if file_year == year and file_month == month:
                filtered_files.append(f)
        elif year:
            if file_year == year:
                filtered_files.append(f)
        else:
            filtered_files.append(f)  # если не указан год/месяц/дата, берем все файлы

    files = filtered_files

    if not files:
        print("⚠️ Нет файлов для обработки после фильтрации!")
        conn.close()
        return

    # 🔹 Проверяем наличие таблицы в БД
    table_name = report_name
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
    table_exists = cursor.fetchone()

    # 🔹 Создаем DataFrame
    all_data = pd.DataFrame()

    # 🔹 Вызываем нужный парсер
    if report_name == "630":
        all_data = prs.sales630_parser(files=files, all_data=all_data)
    elif report_name == "629":
        all_data = prs.sales629_parser(files=files, all_data=all_data)
    elif report_name == "StockCRM":
        all_data = prs.stock_parser(files=files, all_data=all_data)
    elif report_name == "Sales102":
        all_data = prs.sales102_parser(files=files, all_data=all_data)
    else:
        print("⚠️ Файлы для этого отчета не найдены. Обновление не произошло.")
        conn.close()
        return

    if table_exists:
        print(f"🗑 Удаляем старые данные из таблицы {table_name}...")

        # 🔹 Удаление по конкретной дате
        if date:
            try:
                # Преобразуем пользовательский формат DD.MM.YYYY → YYYY-MM-DD
                date_sql = pd.to_datetime(date, dayfirst=True).strftime("%Y-%m-%d")

                cursor.execute(f'''
                    DELETE FROM "{table_name}"
                    WHERE date("Дата") = ?;
                ''', (date_sql,))
            except Exception as e:
                print("❗ Ошибка обработки даты:", e)

        # 🔹 Удаление по году + месяцу
        elif year and month:
            cursor.execute(f'''
                DELETE FROM "{table_name}"
                WHERE strftime("%Y", "Дата") = ?
                  AND strftime("%m", "Дата") = ?;
            ''', (year, month))

        # 🔹 Удаление только по году
        elif year:
            cursor.execute(f'''
                DELETE FROM "{table_name}"
                WHERE strftime("%Y", "Дата") = ?;
            ''', (year,))

        conn.commit()


    else:
        print(f"ℹ️ Таблица {table_name} в БД Sales.db пока не существует. Она будет создана автоматически.")

    # 🔹 Добавляем новые данные
    if not all_data.empty:
        all_data.to_sql(table_name, conn, if_exists='append', index=False)
        print(f"✅ Таблица {table_name} в БД Sales.db обновлена ({len(all_data)} строк).")
    else:
        print("⚠️ Нет данных для обновления!")

    conn.close()
    print("🔒 Соединение с БД закрыто.")


# 🔹 Запуск обновления
db_table_updater(report_name=report_name, year=year, month=month, date=date)
