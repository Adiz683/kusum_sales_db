import pandas as pd
import numpy as np
import xlrd
import os

import dictionaries as dct

def sales630_parser (files, all_data):

    for file in files:
        df = pd.read_excel(file)

        df["Месяц"] = df["Месяц, Год"].str.split(' ').str[0].map(dct.months)
        df["Год"] = df["Месяц, Год"].str.split(' ').str[1].str.replace("(", "").str.replace(")", "").astype(int)
        df["Дата"] = pd.to_datetime(df["Год"].astype(str) + "-" + df["Месяц"] + "-01", format="%Y-%B-%d")

        df["Область ID"] = df["Область"].map(dct.regions_dict)

        df["Товар ID"] = dct.create_drug_id(left_table=df, drug_name="Товар", year="Год")

        df = df.drop(["Регион", "Область", "Товар", "Месяц, Год", "Месяц", "Год", "Цена", "% Продажи шт.", "% Прогноз шт.", "% Продажи деньги", "% Прогноз деньги"],
                     axis=1, errors='ignore')
        df = df[~((df["План шт."] == 0) & (df["Факт шт."] == 0))]
        df = df.groupby(by=["Дата", "Область ID", "Товар ID"], as_index=False, dropna=False).sum()

        col_order = ["Дата", "Область ID", "Товар ID", "План шт.", "Факт шт.", "План деньги", "Факт деньги"]

        df = df[col_order]

        all_data = pd.concat([all_data, df], ignore_index=True)
        print(f"Файл {file} добавлен в DataFrame all_data")

    return all_data




def sales629_parser (files, all_data):
    for file in files:
        df = pd.read_excel(file, header=1)
        df["Область ID"] = df["Область"].map(dct.regions_dict)
        df["Товар ID"] = dct.create_drug_id(left_table=df, drug_name="Товар", year="Год")
        df.rename(columns={"Организация(эталонId)": "Орг ID", "ОКПО": "ИНН", "Количество (IN)":"Факт шт.", "Сумма (IN)": "Факт сумма"}, inplace=True)
        df["Дистрибьютор ID"] = df["Дистрибьютор (эталон)"].map(dct.distrs_dict)
        df["Месяц"] = df["Месяц"].map(dct.months)
        df["Дата"] = pd.to_datetime(df["Год"].astype(str) + "-" + df["Месяц"] + "-01", format="%Y-%B-%d")
        df = df.drop(["Год", "Месяц", "Неделя", "Категория товарооборота", "Дата отгрузки", "Область", "Населенный пункт", "Территория", "Полигон", "Организация",
                      "Почтовый адрес (организация)", "Вид организации", "Тип Организации", "Гл. Офис Сети", "Гл. Офис ОКПО", "Группа товара", "Группа товара 2",
                      "Бренд", "Товар", "Адрес, Название, ОКПО(отчет дистрибьютора)", "ОКПО (дистрибъютор)",
                      "Имя организации (дистрибъютор)", "Дистрибьютор (эталон)", "Адрес (дистрибъютор)", "Сотрудник", "Месяц, Год"],
                      axis=1, errors='ignore')
        df = df.groupby(by=["Дата", "Область ID", "Орг ID", "ИНН", "Товар ID", "Дистрибьютор ID"], as_index=False).sum()
        col_order = ["Дата", "Область ID", "Орг ID", "ИНН", "Товар ID", "Факт шт.", "Факт сумма", "Дистрибьютор ID"]
        df = df[col_order]

        all_data = pd.concat([all_data, df], ignore_index=True)
        print(f"Файл {file} добавлен в DataFrame all_data")
    return all_data




def sales102_parser (files, all_data):
    for file in files:
        df = pd.read_excel(file, header=7, engine="xlrd")
        df["Год"] = df["Дата"].str[-4:].astype(int)
        df["Товар ID"] = dct.create_drug_id(left_table=df, drug_name="Название товара (кратко)", year="Год")
        df["Дистрибьютор ID"] = df["Контрагент (плательщик)(кратко)"].map(dct.distrs_dict)
        df["Дата"] = pd.to_datetime(df["Дата"], dayfirst=True)
        # df["Количество"].astype("float64")
        df = df.drop(['Документ', 'Создан на основании', 'Название товара (кратко)',
                      'Код SAP', 'Производитель (кратко)', 'Серия', 'Годен до',
                      'Валюта', 'Скр.%', 'Скр. скидка', 'Сумма после наценки',
                      'Откр.%', 'Откр. скидка', 'Сумма без НДС', 'Сумма скидки', 'Сумма НДС',
                      'Сумма с НДС', 'Контрагент (плательщик)(кратко)', 'Контрагент (получатель)(кратко)',
                      'Адрес', 'Область', 'ID товара', 'ID фирмы', 'ОКПО', 'ИНН', 'Поставщик',
                      'Товарные акции', 'Ассортимент (пакет)', 'Год'], axis=1, errors='ignore')
        df.rename(columns={"Сумма продажи по базовой цене": "Сумма"}, inplace=True)
        col_order = ["Дата", "Дистрибьютор ID", "Товар ID", "Количество", "Сумма"]
        df = df[col_order]
        all_data = pd.concat([all_data, df], ignore_index=True)
        print(f"Файл {file} добавлен в DataFrame all_data")
    return all_data




def nd_weekly_parser(files, all_data, conn):
    """
    Парсер для еженедельных/ежемесячных файлов "ND_Weekly" (Qlik), например:
    "2025.01.01_SS_Weekly.xlsx"

    Ожидаемые столбцы в файле (некоторые с лишними пробелами в исходнике,
    поэтому сначала обрезаем пробелы в названиях столбцов):
    "Организация", "Область", " Город", "Адрес", "Препарат", " Дистрибутор",
    " Дата ", "Упаковки Факт ", "Суммы Факт", "EDRPO" (это ОКПО)
    """
    for file in files:
        file_date = dct.extract_file_date(file)
        df = pd.read_excel(file)

        # убираем лишние/неразрывные пробелы в названиях столбцов
        df.columns = (
            df.columns.astype(str)
            .str.replace("\xa0", " ", regex=False)
            .str.strip()
        )

        # ищем столбец с ОКПО без учёта регистра и возможных вариантов написания
        # (в разных выгрузках встречается "EDRPO", "Edrpo", "ЕДРПОУ" и т.п.)
        edrpo_candidates = ["EDRPO", "EDRPOU", "ЕДРПОУ", "ОКПО"]
        edrpo_col = next(
            (c for c in df.columns if c.strip().upper() in [x.upper() for x in edrpo_candidates]),
            None
        )
        if edrpo_col is None:
            raise KeyError(
                f"Не найден столбец с ОКПО (искали {edrpo_candidates}) в файле {file}.\n"
                f"Реальные столбцы в файле: {list(df.columns)}"
            )

        # валидация ОКПО (9 цифр, начинается на 2 или 3, иначе - None)
        df["ОКПО"] = dct.validate_okpo(df[edrpo_col])

        # справочники
        df["Область ID"] = df["Область"].map(dct.regions_dict)
        df["Дистрибьютор ID"] = df["Дистрибутор"].map(dct.distrs_dict)

        df["Дата"] = pd.to_datetime(df["Дата"], dayfirst=True)
        df["Год"] = df["Дата"].dt.year
        df["Товар ID"] = dct.create_drug_id(left_table=df, drug_name="Препарат", year="Год")

        # получаем/создаём Орг ID клиента (по связке ОКПО + Область, см. dct.get_or_create_org_id)
        df = dct.get_or_create_org_id(df, conn, record_date=file_date)

        df.rename(columns={"Упаковки Факт": "Факт шт.", "Суммы Факт": "Факт сумма"}, inplace=True)

        df = df.drop(
            ["Организация", "Область", "Город", "Адрес", "Препарат", "Дистрибутор", edrpo_col, "Год"],
            axis=1, errors="ignore"
        )

        df = df.groupby(
            by=["Дата", "Область ID", "Орг ID", "ОКПО", "Товар ID", "Дистрибьютор ID"], as_index=False,
            dropna=False
        ).sum()

        col_order = ["Дата", "Область ID", "Орг ID", "ОКПО", "Товар ID", "Дистрибьютор ID", "Факт шт.", "Факт сумма"]
        df = df[col_order]

        all_data = pd.concat([all_data, df], ignore_index=True)
        print(f"Файл {file} добавлен в DataFrame all_data")

    return all_data


def stock_parser (files, all_data):
    for file in files:
        df = pd.read_excel(file)
        df["Область ID"] = df["Область"].map(dct.regions_dict)
        df["Дистрибьютор ID"] = df["Дистрибьютор"].map(dct.distrs_dict)
        df.rename(columns={"Код": "Товар ID"}, inplace=True)
        df = df.drop(
            ["Год", "Неделя", "Регион", "Область", "Населенный пункт", "Дистрибьютор", "Товар", "Дистрибьютор (склад)",
             "Сумма", "Полное название", "Группа товара", "Группа товара 2", "Бренд",
             ],
            axis=1, errors='ignore')
        col_order = ["Область ID", "Дистрибьютор ID", "Товар ID", "Количество", "Цена", "Дата"]
        df = df[col_order]

        all_data = pd.concat([all_data, df], ignore_index=True)
        print(f"Файл {file} добавлен в DataFrame all_data")
    return all_data