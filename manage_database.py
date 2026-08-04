import pandas as pd
import sqlite3 as sq
import os
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 0)

from dictionaries import df_regions, create_drug_id

conn = sq.connect('../../02_Database/01_Sales_DB/Sales.db') #Создание общей базы данных
cursor = conn.cursor()
# cursor.execute(f'''DROP TABLE IF EXISTS "ND"''')






# # SALES PLAN UPDATE
# sheets = pd.read_excel(io="../../01_Data_Source/02_Reference/05_Sales_Plan.xlsx", sheet_name=None)
#
# df_plan = pd.concat(
#     [df.melt(id_vars='Препарат', var_name='Дата', value_name='План')
#         for df in sheets.values()
#     ], ignore_index=True
# )
# df_plan['Дата'] = pd.to_datetime(df_plan['Дата'])
# df_plan['product_id'] = create_drug_id(df_plan, df_plan['Препарат'], df_plan["Дата"].dt.year)
#
# df_plan.to_sql("Sales_Plan", conn, if_exists="replace", index=False)
# print(df_plan)





# PRICE UPDATE
# sheets = pd.read_excel(io="../../01_Data_Source/02_Reference/06_Price.xlsx", sheet_name=None)
#
# df_price = pd.concat(
#     [df.melt(id_vars='Препарат', var_name='Дата', value_name='Цена')
#       for df in sheets.values()
#      ], ignore_index=True)
#
# df_price['Дата'] = pd.to_datetime(df_price['Дата'])
# df_price['product_id'] = create_drug_id(df_price, df_price['Препарат'], df_price["Дата"].dt.year)
#
# df_price.to_sql("Price", conn, if_exists="replace", index=False)
# print(df_price)




# EMPLOYEE UPDATE
# REGIONAL MANAGERS
# df_rm = pd.read_excel(io="../../01_Data_Source/02_Reference/04_Employees.xlsx",
#                         sheet_name="RM")
#
# df_rm.to_sql("Regional_Managers", conn, if_exists="replace", index=False)




# DEFECTURE UPDATE
# df_drugs = pd.read_excel(io="../../01_Data_Source/02_Reference/02_Drugs.xlsx",
#                         sheet_name="Defecture")
#
# df_drugs.to_sql("Defecture", conn, if_exists="replace", index=False)





# CALENDAR UPDATE
# df_calendar = pd.read_excel(io="../../01_Data_Source/02_Reference/05_Calendar.xlsx",
#                          sheet_name="Calendar")
# df_calendar.to_sql("Calendar", conn, if_exists="replace", index=False)






# REGIONS UPDATE
# df_oblast = pd.read_excel(io="../../01_Data_Source/02_Reference/01_Regions.xlsx",
#                          sheet_name="Regions")
# df_oblast.to_sql("Regions", conn, if_exists="replace", index=False)





# DRUGS UPDATE
# df_drugs = pd.read_excel(io="../../01_Data_Source/02_Reference/02_Drugs.xlsx",
#                         sheet_name="Drugs")
#
# df_drugs.to_sql("Drugs", conn, if_exists="replace", index=False)






#DISTRIBUTORS UPDATE
df_distr_by_year = pd.read_excel(io="../../01_Data_Source/02_Reference/03_Distributors.xlsx",
                        sheet_name="Distr_by_year")

df_distr_by_year.to_sql("Distributors", conn, if_exists="replace", index=False)








conn.commit()
conn.close()