import pandas as pd
import numpy as np
from pathlib import Path
import dictionaries as dct

pd.set_option('display.max_columns', None)
pd.set_option('display.width', 0)

# def nd_weekly_parser(file):
#
#        # for file in files:
#        df = pd.read_excel(file)
#        df["distr_id"] = df["Дистрибутор"].map(dct.distrs_dict)
#
#        df["product_id"] = dct.create_drug_id(left_table=df,drug_name='Препарат', year=df["Дата"].dt.year)
#
#        print(df.head(20))
#        # all_data = pd.concat([all_data, df], ignore_index=True)
#
#
#
# sec_sales_weekly_parser("../../01_Data_Source/01_Exported_data/06_Parser_ND_Qlik/2025.01.01_SS_Weekly.xlsx")