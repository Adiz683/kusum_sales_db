import pandas as pd
import sqlite3 as sq
import re

# ---------- 1. Загрузка исходных данных ----------
df_org = pd.read_excel(
    "../../01_Data_Source/01_Exported_data/02_629_H3_CRM/2021.01.01_export_regional_sales.xlsx",
    header=1
)

id_col = "Организация(эталонId)"
cols = [
    "Область",
    "Населенный пункт",
    "Организация",
    "Почтовый адрес (организация)",
    "ОКПО",
]

df = df_org[[id_col] + cols].copy()

# Без ID строка бесполезна как справочная запись
df = df.dropna(subset=[id_col])


OKPO_RE = re.compile(r"[23]\d{8}")  # 9 цифр, первая цифра 2 или 3


# ---------- 2. Нормализация ОКПО ----------
# Важно делать это ДО группировки/проверки на конфликты, а не только перед
# проверкой валидности: иначе "200976426" и "200 976 426" по одному и тому же
# ID будут считаться двумя разными значениями и ID ошибочно попадёт в "плохие".
def normalize_okpo(value):
    if pd.isna(value):
        return value

    s = str(value).strip()
    if s.endswith(".0"):  # Excel иногда отдаёт число как float
        s = s[:-2]

    s = re.sub(r"\s+", "", s)  # убираем пробелы внутри и по краям, напр. "200 976 426"

    if "|" in s:
        # "двойные" значения вида "04899026|200000034" — берём часть,
        # которая проходит проверку на валидный ОКПО (если валидна ровно одна).
        parts = [re.sub(r"\s+", "", p) for p in s.split("|")]
        valid_parts = [p for p in parts if OKPO_RE.fullmatch(p)]
        if len(valid_parts) == 1:
            return valid_parts[0]
        # если валидных частей 0 или больше 1 — не угадываем, оставляем как есть
        # (дальше это всё равно будет забраковано как невалидный ОКПО)
        return s

    return s


df["ОКПО"] = df["ОКПО"].apply(normalize_okpo)


# ---------- 3. Проверка валидности ОКПО ----------
def is_valid_okpo(value) -> bool:
    if pd.isna(value):
        return False
    return bool(OKPO_RE.fullmatch(str(value)))


df["ОКПО_valid"] = df["ОКПО"].apply(is_valid_okpo)


# ---------- 4. Классификация ID на хорошие/плохие ----------
good_ids = []
bad_ids_info = []  # для отчёта, почему ID забракован

for org_id, group in df.groupby(id_col):
    reasons = []

    if group[cols].isnull().any().any():
        reasons.append("есть null в одном из полей")

    dup_cols = group[cols].nunique()
    conflicting = dup_cols[dup_cols > 1].index.tolist()
    if conflicting:
        reasons.append(f"конфликт значений в столбцах: {conflicting}")

    if not group["ОКПО_valid"].iloc[0]:
        reasons.append("невалидный ОКПО")

    if reasons:
        bad_ids_info.append((org_id, "; ".join(reasons)))
    else:
        good_ids.append(org_id)


# ---------- 5. Формирование итоговой таблицы ----------
df_good = (
    df[df[id_col].isin(good_ids)]
    .drop_duplicates(subset=[id_col])[[id_col] + cols]
    .reset_index(drop=True)
)

print(f"Всего уникальных ID: {df[id_col].nunique()}")
print(f"Хороших ID: {len(good_ids)}")
print(f"Плохих ID: {len(bad_ids_info)}")

# при желании можно посмотреть причины брака:
# for org_id, reason in bad_ids_info:
#     print(org_id, "->", reason)


# ---------- 6. Запись в базу ----------
conn = sq.connect('../../02_Database/01_Sales_DB/Sales.db')
df_good.to_sql("Organizations_Reference", conn, if_exists="replace", index=False)
conn.commit()
conn.close()

print("Готово: таблица Organizations_Reference записана в Sales.db")