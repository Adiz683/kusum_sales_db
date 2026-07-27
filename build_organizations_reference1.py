import pandas as pd
import numpy as np
import sqlite3 as sq

pd.set_option('display.max_columns', None)
pd.set_option('display.width', 0)



df = pd.read_excel(
    "../../01_Data_Source/01_Exported_data/02_629_H3_CRM/2021.01.01_export_regional_sales.xlsx",
    header=1
)



df["ОКПО"] = df["ОКПО"].str.replace(" ", "", regex=False)

def okpo_split(inn, inn2):
    def clean(value):
        if pd.isna(value):
            return None

        value = str(value)

        if "|" in value:
            parts = value.split("|")
            for part in parts:
                if len(part) == 9 and part.startswith(("2", "3")):
                    return part
            return None

        if len(value) != 9 or not value.startswith(("2", "3")):
            return None

        return value

    result = clean(inn)

    if result is not None:
        return result

    return clean(inn2)  # проверяется тем же правилом; если не пройдёт — вернёт None



# Применяем функцию okpo_split к нужной колонке
df["ОКПО"] = [okpo_split(a, b) for a, b in zip(df["ОКПО"], df["ОКПО (дистрибъютор)"])]




# "n/a" из Excel -> NaN
df["Организация"] = df["Организация"].replace("n/a", np.nan)
df["Почтовый адрес (организация)"] = df["Почтовый адрес (организация)"].replace("n/a", np.nan)

# mock_id по умолчанию = эталонный id (для строк без ошибок и для "чистых" групп)
df["mock_id"] = df["Организация(эталонId)"]

# Маска: строки, где "Организация" пустая (признак ошибки)
mask_na_org = df["Организация"].isna()

# Среди этих строк считаем кол-во уникальных донор-имён на каждый эталонId
nunique_names = df.loc[mask_na_org].groupby("Организация(эталонId)")["Имя организации (дистрибъютор)"] \
                   .transform(lambda x: x.dropna().nunique())

# Разделяем на "одно имя" и "несколько имён"
idx_single = nunique_names[nunique_names == 1].index
idx_multi = nunique_names[nunique_names > 1].index


def pick_name(sub):
    """Донор для Организация: первое значение из Имя организации (дистрибъютор)"""
    names = sub["Имя организации (дистрибъютор)"].dropna()
    return names.iloc[0] if not names.empty else np.nan


def pick_address(sub):
    """Донор для адреса: самое длинное значение из Адрес (дистрибъютор)"""
    addrs = sub["Адрес (дистрибъютор)"].dropna()
    return addrs.loc[addrs.str.len().idxmax()] if not addrs.empty else np.nan


# --- Случай 1: одно уникальное донор-имя на эталонId ---
def fill_single_group(g):
    g = g.copy()
    g["Организация"] = pick_name(g)
    g["Почтовый адрес (организация)"] = pick_address(g)
    return g

df.loc[idx_single] = df.loc[idx_single].groupby("Организация(эталонId)", group_keys=False).apply(fill_single_group)


# --- Случай 2: несколько уникальных донор-имён на один эталонId ---
mock_id_map = {}
current_id = 999999

def fill_multi_group(g):
    global current_id
    parts = []

    has_okpo = g["ОКПО"].notna()

    # Строки С известным ОКПО: группируем по значению ОКПО (единый ОКПО = одна организация)
    for okpo_value, sub in g[has_okpo].groupby("ОКПО"):
        sub = sub.copy()
        sub["Организация"] = pick_name(sub)
        sub["Почтовый адрес (организация)"] = pick_address(sub)

        key = (sub["Организация(эталонId)"].iloc[0], "okpo", okpo_value)
        if key not in mock_id_map:
            mock_id_map[key] = current_id
            current_id -= 1
        sub["mock_id"] = mock_id_map[key]

        parts.append(sub)

    # Строки БЕЗ ОКПО: группируем по уникальному имени донора
    for name_value, sub in g[~has_okpo].groupby("Имя организации (дистрибъютор)"):
        sub = sub.copy()
        sub["Организация"] = pick_name(sub)
        sub["Почтовый адрес (организация)"] = pick_address(sub)

        key = (sub["Организация(эталонId)"].iloc[0], "name", name_value)
        if key not in mock_id_map:
            mock_id_map[key] = current_id
            current_id -= 1
        sub["mock_id"] = mock_id_map[key]

        parts.append(sub)

    return pd.concat(parts)

df.loc[idx_multi] = df.loc[idx_multi].groupby("Организация(эталонId)", group_keys=False).apply(fill_multi_group)


# NULL/пустые ячейки Excel -> NaN, если ещё не приведено
df["Почтовый адрес (организация)"] = df["Почтовый адрес (организация)"].replace("NULL", np.nan)
df["ОКПО"] = df["ОКПО"].replace("NULL", np.nan)

# Заполняем пропуски внутри каждой группы (по mock_id) значением, которое уже есть в этой же группе
for col in ["Почтовый адрес (организация)", "ОКПО"]:
    df[col] = df.groupby("mock_id")[col].transform(
        lambda x: x.ffill().bfill()
    )




# ============================================================
# Разрешение конфликтов на уровне "mock_id"
# ============================================================

# Кейс 1: несколько разных ОКПО на один mock_id -> оставляем первое
okpo_conflicts = df.groupby("mock_id")["ОКПО"] \
                    .apply(lambda x: x.dropna().unique())
okpo_conflicts = okpo_conflicts[okpo_conflicts.apply(len) > 1]
print("mock_id с разными ОКПО:")
print(okpo_conflicts)

df["ОКПО"] = df.groupby("mock_id")["ОКПО"].transform(
    lambda x: x.dropna().iloc[0] if x.dropna().any() else np.nan
)

# Кейс 2: несколько разных комбинаций (Область, Населенный пункт) на один mock_id -> оставляем первую
region_conflict_counts = df.groupby("mock_id")[["Область", "Населенный пункт"]] \
                            .apply(lambda g: g.drop_duplicates().shape[0])
conflict_ids = region_conflict_counts[region_conflict_counts > 1].index
print("mock_id с разными комбинациями Область/Населенный пункт:")
print(conflict_ids.tolist())

def first_region_combo(g):
    sub = g[["Область", "Населенный пункт"]].dropna(how="all")
    if sub.empty:
        first_vals = pd.Series({"Область": np.nan, "Населенный пункт": np.nan})
    else:
        first_vals = sub.iloc[0]
    return pd.DataFrame([first_vals] * len(g), index=g.index)

df[["Область", "Населенный пункт"]] = df.groupby("mock_id", group_keys=False).apply(
    lambda g: first_region_combo(g)
)







conn = sq.connect('../../02_Database/01_Sales_DB/Sales.db')



cols_for_sales = [
    "mock_id",
    "Товар",
    "Количество (IN)",
    "Сумма (IN)",
    "Дата отгрузки"
]

cols_for_orgs_reference = [
    "mock_id",
    "Организация(эталонId)",
    "Область",
    "Населенный пункт",
    "Организация",
    "Почтовый адрес (организация)",
    "ОКПО",
    "Дата отгрузки"
]

df_sales = df[cols_for_sales].copy()
df_orgs = df[cols_for_orgs_reference].copy().drop_duplicates(ignore_index=True)

# print(df_sales.head())





df_sales.to_sql('629Sales', conn, if_exists='replace', index=False)
df_orgs.to_sql('Organizations', conn, if_exists='replace', index=False)

conn.close()

print("Готово: таблица записана в Sales.db")

