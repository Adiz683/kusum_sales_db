import pandas as pd
import numpy as np
import sqlite3 as sq

pd.set_option('display.max_columns', None)
pd.set_option('display.width', 0)
pd.set_option('future.no_silent_downcasting', True)




df = pd.read_excel(
    "../../01_Data_Source/01_Exported_data/02_629_H3_CRM/2021.01.01_export_regional_sales.xlsx",
    header=1,
    dtype={"ОКПО": str, "ОКПО (дистрибъютор)": str})



# БЛОК ДЛЯ ОЧИСТКИ ИЗ ЛИШНЫХ СИМВОЛОВ СТОЛБЦЫ "ОКПО" И "ОКПО (дистрибъютор)"
df["ОКПО"] = (
    df["ОКПО"]
    .astype(str)
    .str.replace("\xa0", "", regex=False)
    .str.replace(" ", "", regex=False)
    .str.strip()
)
df.loc[df["ОКПО"].isin(["nan", "None", ""]), "ОКПО"] = None

df["ОКПО (дистрибъютор)"] = (
    df["ОКПО (дистрибъютор)"]
    .astype(str)
    .str.replace("\xa0", "", regex=False)
    .str.replace(" ", "", regex=False)
    .str.strip()
)
df.loc[df["ОКПО (дистрибъютор)"].isin(["nan", "None", ""]), "ОКПО (дистрибъютор)"] = None
# БЛОК ДЛЯ ОЧИСТКИ ИЗ ЛИШНЫХ СИМВОЛОВ СТОЛБЦЫ "ОКПО" И "ОКПО (дистрибъютор)"




def is_valid_okpo(val):
    """9 знаков и начинается на 2 или 3."""
    if not isinstance(val, str):
        return False
    val = val.strip()
    return len(val) == 9 and val[0] in "23"

df["ОКПО"] = df["ОКПО"].str.replace(" ", "", regex=False)

def split_okpo(val):
    """Если есть '|', оставляем первое подходящее значение, иначе не трогаем."""
    if not isinstance(val, str) or "|" not in val:
        return val
    for part in val.split("|"):
        part = part.strip()
        if len(part) == 9 and part[0] in "23":
            return part
    return val


def fix_okpo(row):
    """Если ОКПО некорректный, берём из дистрибьютора (если тот корректный)."""
    if is_valid_okpo(row["ОКПО"]):
        return row["ОКПО"].strip()
    if is_valid_okpo(row["ОКПО (дистрибъютор)"]):
        return row["ОКПО (дистрибъютор)"].strip()
    return row["ОКПО"]


# Шаг 1: разделение по "|"
df["ОКПО"] = df["ОКПО"].apply(split_okpo)

# Шаг 2: замена из "ОКПО (дистрибъютор)"
df["ОКПО"] = df.apply(fix_okpo, axis=1)

# Шаг 3: невалидные значения ОКПО -> NULL
df["ОКПО"] = df["ОКПО"].where(df["ОКПО"].apply(is_valid_okpo), None)



# Заполнение NULL в "Организация" и "Почтовый адрес (организация)"
mask_null_org = df["Организация"].isna()

df.loc[mask_null_org, "Организация"] = df.loc[mask_null_org, "Имя организации (дистрибъютор)"]
df.loc[mask_null_org, "Почтовый адрес (организация)"] = df.loc[mask_null_org, "Адрес (дистрибъютор)"]










# ЗАПОЛНЕНИЕ NULL ЗНАЧЕНИЙ ОКПО С ВАЛИДНЫМ ОКПО ИЗ ЭТОГО ЭТАЛОН ID
group_cols = ["Организация(эталонId)", "Организация", "Почтовый адрес (организация)"]

df["ОКПО"] = df.groupby(group_cols)["ОКПО"].transform(lambda s: s.ffill().bfill())
# ЗАПОЛНЕНИЕ NULL ЗНАЧЕНИЙ ОКПО С ВАЛИДНЫМ ОКПО ИЗ ЭТОГО ЭТАЛОН ID



group_cols = ["Область", "Населенный пункт", "Организация", "Почтовый адрес (организация)", "ОКПО"]

# Уникальные комбинации для каждого эталонId
df_unique_combos = df.drop_duplicates(subset=["Организация(эталонId)"] + group_cols)

# Считаем сколько уникальных комбинаций у каждого эталонId
combo_counts = df_unique_combos.groupby("Организация(эталонId)").size()

# ID, у которых больше одной уникальной комбинации -> нужен mock_id
ids_need_mock = combo_counts[combo_counts > 1].index

def assign_mock_id(sub_df):
    """Присваивает mock_id внутри одного эталонId, если у него несколько уникальных комбинаций."""
    etalon_id = sub_df.name
    if etalon_id not in ids_need_mock:
        sub_df = sub_df.copy()
        sub_df["mock_id"] = None
        return sub_df

    sub_df = sub_df.copy()
    # определяем номер группы по уникальной комбинации внутри этого эталонId
    combo_keys = sub_df[group_cols].astype(str).agg("|".join, axis=1)
    unique_combos = combo_keys.drop_duplicates().tolist()
    combo_to_num = {combo: i + 1 for i, combo in enumerate(unique_combos)}
    sub_df["mock_id"] = combo_keys.map(combo_to_num).apply(
        lambda n: f"{etalon_id}-{n}"
    )
    return sub_df

df = df.groupby("Организация(эталонId)", group_keys=False).apply(assign_mock_id)

df["pharm_id"] = df["mock_id"].fillna(df["Организация(эталонId)"].astype(str))







# Дата отгрузки -> дата без времени
df["Дата отгрузки"] = pd.to_datetime(df["Дата отгрузки"]).dt.date

# created_at - копия Дата отгрузки
df["created_at"] = df["Дата отгрузки"]







#ВЫБИРАЕМ ИМПОРТИРУЕМЫЕ СТОЛБЦЫ ДЛЯ СУБД
cols_for_pharms = [
    "pharm_id",
    "mock_id",
    "Организация(эталонId)",
    "Область",
    "Населенный пункт",
    "Организация",
    "Почтовый адрес (организация)",
    "ОКПО",
    "created_at"]
#ВЫБИРАЕМ ИМПОРТИРУЕМЫЕ СТОЛБЦЫ ДЛЯ СУБД




df_orgs = df[cols_for_pharms].copy().drop_duplicates(ignore_index=True)





conn = sq.connect('../../02_Database/01_Sales_DB/Sales.db')

df_orgs.to_sql('Pharms', conn, if_exists='replace', index=False)

conn.close()

print("Готово: таблица Pharms записана в Sales.db")