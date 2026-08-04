import pandas as pd
import os

# СЛОВАРЬ ПАПОК ОТЧЕТОВ
report_paths = {
    "630": "01_630_CRM",
    "629": "02_629_H3_CRM",
    "SalesQlik": "03_2.21_Qlik",
    "Sales102": "04_102_Torgovaja",
    "StockCRM": "05_1603_stock_CRM",
    "ND": "06_Parser_ND_Qlik"
}

# СЛОВАРЬ МЕСЯЦА
months = {
    'Январь': 'January', 'Февраль': 'February', 'Март': 'March',
    'Апрель': 'April', 'Май': 'May', 'Июнь': 'June',
    'Июль': 'July', 'Август': 'August', 'Сентябрь': 'September',
    'Октябрь': 'October', 'Ноябрь': 'November', 'Декабрь': 'December'}




# СЛОВАРЬ ОБЛАСТЕЙ
df_regions = pd.read_excel(io="../../01_Data_Source/02_Reference/01_Regions.xlsx",
                         sheet_name="Regions")
regions_dict = dict(zip(df_regions['Region'], df_regions['Region ID']))






# СЛОВАРЬ ДИСТРИБЬЮТОРОВ
df_distrs = pd.read_excel(io="../../01_Data_Source/02_Reference/03_Distributors.xlsx",
                         sheet_name="Distr_mapping")
distrs_dict = dict(zip(df_distrs['Имя дистрибьютора (отчет)'], df_distrs['Distr ID']))







# СЛОВАРЬ ПРЕПАРАТОВ (для замены имя препаратов с ID для препаратов внутри Справочника препаратов по годам)
def create_drug_id (left_table, drug_name, year):
    df_drugs = pd.read_excel(io="../../01_Data_Source/02_Reference/02_Drugs.xlsx",
                             sheet_name="Drugs")
    df_drugs['Год'] = df_drugs['Год'].astype(int)
    drug_id = left_table.merge(
        df_drugs[['Название препарата', 'Год', 'Drug ID']],
        left_on=[drug_name, year],
        right_on=['Название препарата', 'Год'],
        how='left'
    )
    return drug_id["Drug ID"]


# ВАЛИДАЦИЯ ОКПО (EDRPO)
# Валидное ОКПО - это 9-значное число, начинающееся на 2 или 3.
# Всё остальное (пусто, другая длина, другая первая цифра, мусор) превращаем в None.
def validate_okpo(series):
    s = series.astype(str).str.strip()
    # убираем ".0", если ОКПО было прочитано как float (например 234567890.0)
    s = s.str.replace(r"\.0$", "", regex=True)
    # берём первые 9 символов - в выгрузке встречаются значения с мусором после
    # валидного ОКПО (например "307870693f4" -> нужно получить "307870693")
    s = s.str[:9]
    is_valid = s.str.fullmatch(r"[23]\d{8}")
    return s.where(is_valid.fillna(False), other=None)


def extract_file_date(file_path):
    """
    Извлекает дату из названия файла вида "2025.01.01_SS_Weekly.xlsx" (или с "_" вместо ".").
    Позиции в имени файла фиксированные: год - символы 0-3, месяц - 5-6, день - 8-9
    (разделитель между ними не важен, работает и с точками, и с подчёркиваниями).
    Возвращает дату в формате "YYYY-MM-DD" (строкой) или None, если распарсить не удалось.
    """
    fname = os.path.basename(file_path)
    try:
        year, month, day = int(fname[0:4]), int(fname[5:7]), int(fname[8:10])
        return pd.Timestamp(year=year, month=month, day=day).date().isoformat()
    except (ValueError, TypeError):
        return None


# СПРАВОЧНИК (ТАБЛИЦА ИЗМЕРЕНИЙ) КЛИЕНТОВ ДЛЯ ND_Weekly
def get_or_create_org_id(df, conn, dim_table="ND_Organizations", record_date=None):
    """
    Сопоставляет клиентов из текущего файла с их "Орг ID" из таблицы измерений (справочника
    клиентов) dim_table. Если клиент новый - создаёт новый "Орг ID" и добавляет клиента
    в справочник.

    df должен содержать столбцы: "ОКПО", "Организация", "Область".

    record_date - дата обрабатываемого файла (обычно берётся из его названия, см.
    parsers.extract_file_date), в формате "YYYY-MM-DD". Записывается в столбцы
    "created_at" (при создании нового клиента) и "updated_at" (при создании, а также при
    любом последующем изменении данных клиента - апгрейде ОКПО или смене названия). Если
    клиент ни разу не менялся с момента создания, created_at и updated_at совпадают.

    Основной ключ сопоставления - связка ОКПО + Область. Если ОКПО невалидный (после
    validate_okpo стал None), используется запасной ключ - связка Организация + Область
    (null не равен null-у ни в pandas, ни в SQL, поэтому сопоставлять по одному только
    пустому ОКПО нельзя).

    Дополнительно обрабатываются два случая, когда основной ключ не совпал напрямую:

    1. АПГРЕЙД: клиент раньше был записан в справочник без ОКПО (по ключу Организация+Область),
       а в новом файле у него появился валидный ОКПО. Тогда ищем в справочнике клиента с тем же
       названием+областью, у которого ОКПО ещё не заполнен, и обновляем ему ОКПО и ключ
       (Орг ID при этом не меняется, новая строка не создаётся).
    2. ОБРАТНЫЙ СЛУЧАЙ: у клиента в справочнике уже есть ОКПО (был найден ранее), а в текущем
       файле для него почему-то указан пустой/невалидный ОКПО. Находим его по названию+области
       и используем существующий Орг ID, ничего не перезаписывая в справочнике (новых данных нет).

    Если название "Организация" в новом файле отличается от того, что сохранено в справочнике
    (например, компанию переименовали) - название в ND_Organizations актуализируется.

    ЗАЩИТА ОТ "ОТКАТА" СТАРЫМИ ФАЙЛАМИ: если запись в справочнике уже обновлялась файлом
    с более новой датой (updated_at), чем record_date текущего файла - изменения (апгрейд ОКПО
    или переименование) из текущего (более старого) файла НЕ применяются. Смысл: если,
    например, сначала обработали 2026.02.01, а потом (довобновление/пересчёт) обрабатывают
    2025.03.01 - более старый файл не должен затирать более свежие данные о клиенте. Сам
    Орг ID при этом всё равно возвращается корректно - блокируется только запись в справочник.
    """
    def is_stale(candidate_date, existing_updated_at):
        # не можем сравнить (одна из дат неизвестна) - не блокируем, применяем изменение
        if candidate_date is None or existing_updated_at is None or pd.isna(existing_updated_at):
            return False
        return candidate_date < existing_updated_at

    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (dim_table,))
    table_exists = cursor.fetchone()

    if table_exists:
        # миграция схемы: если таблица создана раньше, до появления created_at/updated_at -
        # добавляем эти столбцы (для уже существующих клиентов дата создания неизвестна,
        # поэтому останется NULL - это ожидаемо и безопасно)
        cursor.execute(f'PRAGMA table_info("{dim_table}")')
        existing_cols = [col[1] for col in cursor.fetchall()]
        if "created_at" not in existing_cols:
            cursor.execute(f'ALTER TABLE "{dim_table}" ADD COLUMN "created_at" TEXT')
        if "updated_at" not in existing_cols:
            cursor.execute(f'ALTER TABLE "{dim_table}" ADD COLUMN "updated_at" TEXT')
        conn.commit()
        dim_df = pd.read_sql(f'SELECT * FROM "{dim_table}"', conn)
    else:
        dim_df = pd.DataFrame(columns=[
            "Орг ID", "ОКПО", "Организация", "Область", "match_key", "created_at", "updated_at"
        ])

    def build_key(okpo, name, area):
        if pd.notna(okpo):
            return f'OKPO_{okpo}_{area}'
        else:
            return f'NAME_{str(name).strip().upper()}_{area}'

    def build_area_key(name, area):
        # ключ по названию+области без учёта ОКПО - используется как "мостик" между
        # старой записью без ОКПО и новой строкой, где ОКПО уже появился (и наоборот)
        return f'{str(name).strip().upper()}_{area}'

    # быстрые справочники по уже известным клиентам
    key_to_id = dict(zip(dim_df["match_key"], dim_df["Орг ID"])) if not dim_df.empty else {}
    key_to_name = dict(zip(dim_df["match_key"], dim_df["Организация"])) if not dim_df.empty else {}
    id_to_updated_at = dict(zip(dim_df["Орг ID"], dim_df["updated_at"])) if not dim_df.empty else {}

    if not dim_df.empty:
        dim_area_keys = [build_area_key(n, a) for n, a in zip(dim_df["Организация"], dim_df["Область"])]
        area_key_to_id = dict(zip(dim_area_keys, dim_df["Орг ID"]))
        area_key_to_okpo = dict(zip(dim_area_keys, dim_df["ОКПО"]))
    else:
        area_key_to_id = {}
        area_key_to_okpo = {}

    max_id = int(dim_df["Орг ID"].max()) if not dim_df.empty and dim_df["Орг ID"].notna().any() else 0

    df["match_key"] = [
        build_key(okpo, name, area)
        for okpo, name, area in zip(df["ОКПО"], df["Организация"], df["Область"])
    ]

    current_orgs = df[["match_key", "ОКПО", "Организация", "Область"]].drop_duplicates(subset=["match_key"])

    resolved_ids = {}       # match_key текущего файла -> Орг ID
    new_rows = []           # новые клиенты, которых нет в справочнике вообще
    upgrade_updates = []    # (ОКПО, Организация, match_key, updated_at, Орг ID) - клиенту "нашёлся" ОКПО
    rename_updates = []     # (Организация, updated_at, Орг ID) - у существующего клиента поменялось название
    skipped_stale = 0       # счётчик изменений, пропущенных из-за более старой даты файла

    for _, row in current_orgs.iterrows():
        mk = row["match_key"]
        okpo, name, area = row["ОКПО"], row["Организация"], row["Область"]

        # 1) точное совпадение по основному ключу
        if mk in key_to_id:
            org_id = key_to_id[mk]
            if name != key_to_name.get(mk):
                if is_stale(record_date, id_to_updated_at.get(org_id)):
                    skipped_stale += 1
                else:
                    rename_updates.append((name, record_date, org_id))
                    key_to_name[mk] = name
                    id_to_updated_at[org_id] = record_date
            resolved_ids[mk] = org_id
            continue

        area_key = build_area_key(name, area)

        # 2) совпадение по названию+области (мостик между "без ОКПО" и "с ОКПО")
        if area_key in area_key_to_id:
            org_id = area_key_to_id[area_key]
            known_okpo = area_key_to_okpo.get(area_key)

            if pd.notna(okpo) and pd.isna(known_okpo):
                # апгрейд: клиенту раньше не хватало ОКПО, теперь он есть
                if is_stale(record_date, id_to_updated_at.get(org_id)):
                    skipped_stale += 1
                else:
                    upgrade_updates.append((okpo, name, mk, record_date, org_id))
                    area_key_to_okpo[area_key] = okpo
                    key_to_id[mk] = org_id
                    key_to_name[mk] = name
                    id_to_updated_at[org_id] = record_date
            elif pd.notna(okpo) and pd.notna(known_okpo) and str(known_okpo) != str(okpo):
                # у одного и того же названия+области - два РАЗНЫХ валидных ОКПО.
                # это подозрительно (скорее два разных клиента с похожим названием),
                # поэтому НЕ объединяем - создаём отдельного клиента ниже, а не тут
                org_id = None
            else:
                # либо текущая строка без ОКПО (обратный случай), либо ОКПО совпадает -
                # просто переиспользуем существующий Орг ID, в справочник ничего не пишем
                if name != key_to_name.get(mk, name):
                    if is_stale(record_date, id_to_updated_at.get(org_id)):
                        skipped_stale += 1
                    else:
                        rename_updates.append((name, record_date, org_id))
                        key_to_name[mk] = name
                        id_to_updated_at[org_id] = record_date

            if org_id is not None:
                resolved_ids[mk] = org_id
                continue

        # 3) действительно новый клиент
        max_id += 1
        org_id = max_id
        resolved_ids[mk] = org_id
        key_to_id[mk] = org_id
        key_to_name[mk] = name
        area_key_to_id[area_key] = org_id
        area_key_to_okpo[area_key] = okpo
        id_to_updated_at[org_id] = record_date
        new_rows.append((org_id, okpo, name, area, mk, record_date, record_date))

    if new_rows:
        new_df = pd.DataFrame(new_rows, columns=[
            "Орг ID", "ОКПО", "Организация", "Область", "match_key", "created_at", "updated_at"
        ])
        new_df.to_sql(dim_table, conn, if_exists="append", index=False)

    if upgrade_updates:
        cursor.executemany(
            f'UPDATE "{dim_table}" SET "ОКПО" = ?, "Организация" = ?, "match_key" = ?, "updated_at" = ? '
            f'WHERE "Орг ID" = ?',
            upgrade_updates
        )
        print(f"ℹ️ Найден ОКПО для {len(upgrade_updates)} клиент(ов), ранее записанных без ОКПО - справочник обновлён")

    if rename_updates:
        cursor.executemany(
            f'UPDATE "{dim_table}" SET "Организация" = ?, "updated_at" = ? WHERE "Орг ID" = ?',
            rename_updates
        )
        print(f"ℹ️ Обновлено названий организаций в {dim_table}: {len(rename_updates)}")

    if skipped_stale:
        print(
            f"⚠️ Пропущено изменений в {dim_table}: {skipped_stale} - "
            f"обрабатываемый файл (дата {record_date}) старее, чем последнее обновление записи в справочнике"
        )

    if new_rows or upgrade_updates or rename_updates:
        conn.commit()

    # проставляем Орг ID обратно во все строки исходного df
    df["Орг ID"] = df["match_key"].map(resolved_ids).astype("Int64")
    df = df.drop(columns=["match_key"])

    return df