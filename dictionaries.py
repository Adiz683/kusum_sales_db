import pandas as pd
import os

# СЛОВАРЬ ПАПОК ОТЧЕТОВ
report_paths = {
    "630": "01_630_CRM",
    "629": "02_629_H3_CRM",
    "SalesQlik": "03_2.21_Qlik",
    "Sales102": "04_102_Torgovaja",
    "StockCRM": "05_1603_stock_CRM",
    "Parser_ND": "06_Parser_ND_Qlik"
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



