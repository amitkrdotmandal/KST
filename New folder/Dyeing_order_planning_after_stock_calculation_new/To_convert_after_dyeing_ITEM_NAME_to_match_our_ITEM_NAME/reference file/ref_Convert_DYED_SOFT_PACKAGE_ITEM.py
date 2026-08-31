import pandas as pd
import matplotlib.pyplot as plt
import re

import math
import os
from datetime import datetime


file_name_Dyed_soft_package_Item_type_conversion_DB = "Dyed_soft_package_Item_type_conversion_DB.xlsx"

df_Dyed_soft_package_Item = pd.read_excel(file_name_Dyed_soft_package_Item_type_conversion_DB)
df_Dyed_soft_package_Item['only_count'] = None
print(df_Dyed_soft_package_Item.columns)
print(df_Dyed_soft_package_Item.head())
len_of_Dyed_soft_package_Item = len(df_Dyed_soft_package_Item)

for i in range(len_of_Dyed_soft_package_Item):
    vaaro = df_Dyed_soft_package_Item['Item type according to me'][i]
    item_string = df_Dyed_soft_package_Item['Item Name  aacording to dyeing'][i]
    extract_ct_tem = re.search(r'\d+/\d+', item_string)
    if extract_ct_tem:
        extract_ct = extract_ct_tem.group()
        df_Dyed_soft_package_Item.loc[i, 'only_count'] = extract_ct
    df_Dyed_soft_package_Item.loc[i,'Item Name according to me']=str(df_Dyed_soft_package_Item['only_count'][i])+' '+str(df_Dyed_soft_package_Item['Item type according to me'][i])
# print(df_GTG_RAW)
# df_GTG_RAW.to_excel('OUTPUT_converted_GTG_foR_RAW_DB.xlsx', index=True)

####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before
# date_str = datetime.now().strftime("%d-%m-%Y")
# file_name_to_save_all_data="temporary_output_Winding_lead_time_of_"+date_str+".xlsx"
file_name_to_save_all_data='OUTPUT_converted_item_name_of_dyed_soft_pkg_DB.xlsx'  ####don't change this name connected
print(file_name_to_save_all_data)
if os.path.isfile(file_name_to_save_all_data):
    os.remove(file_name_to_save_all_data)
# df_result_2.to_excel(file_name_to_save_all_data, index=True)

df_Dyed_soft_package_Item.to_excel(file_name_to_save_all_data, index=True)

####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before