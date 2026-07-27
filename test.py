import pandas as pd
import numpy as np
from pathlib import Path

pd.set_option('display.max_columns', None)
pd.set_option('display.width', 0)

def create_drugstore_list():
       folder = Path("../../01_Data_Source/01_Exported_data/02_629_H3_CRM")
       files = list(folder.glob("*xlsx"))
       # print(files)
       # for file in files:
       df = pd.read_excel(files[0], header=1)

       df = df.drop(
              ['Год', 'Месяц', 'Неделя', 'Дата отгрузки', 'Территория', 'Полигон',
              'Категория товарооборота', 'Вид организации', 'Тип Организации',
              'Гл. Офис Сети', 'Гл. Офис ОКПО', 'Группа товара', 'Группа товара 2',
              'Бренд', 'Товар', 'Количество (IN)', 'Сумма (IN)',
              'Дистрибьютор (эталон)', 'Адрес, Название, ОКПО(отчет дистрибьютора)',
              'Сотрудник'],
              axis=1, errors='ignore').drop_duplicates(ignore_index=True)

       df_without_na = df[df["Организация"].notna()]

       df_without_na["ИНН"] = df["ОКПО"]

       df_without_na["ИНН"] = df_without_na["ИНН"].where(
              df_without_na["ИНН"].str.len() == 9,
              df_without_na["ОКПО (дистрибъютор)"])

       # df_without_na["ИНН"] = df_without_na["ИНН"].where(
       #        df_without_na["ИНН"].str.len() == 9,
       #        "не определено")


       # print(df.columns)
       # df_new = df[['Организация(эталонId)', 'Организация', 'Почтовый адрес (организация)', 'ОКПО']].drop_duplicates(ignore_index=True)
       print(df_without_na.head(50))

create_drugstore_list()