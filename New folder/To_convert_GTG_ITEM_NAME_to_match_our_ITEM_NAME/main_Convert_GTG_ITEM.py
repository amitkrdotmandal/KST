import pandas as pd
import matplotlib.pyplot as plt
import re
# from search_by_raw_material_name_function import search_by_raw_material_name_function
import math
import os
from datetime import datetime


file_name_GTG_Item_type_converion_DB = "GTG_Item_type_converion_DB.xlsx"

df_GTG_RAW = pd.read_excel(file_name_GTG_Item_type_converion_DB, sheet_name="Raw")
df_GTG_RAW['only_count'] = None
print(df_GTG_RAW.columns)
print(df_GTG_RAW.head())
len_of_data_RAW = len(df_GTG_RAW)
for i in range(len_of_data_RAW):
    vaaro = df_GTG_RAW['Item type according to me'][i]
    list_first_digit_count=['POLY POLY CORE SPUN','SPUN POLYESTER YARN','TEXTURISED POLYESTER']##### may need to change
    if vaaro not in list_first_digit_count:
        item_string = df_GTG_RAW['Item Name according to GTG'][i]
        extract_ct_tem = re.search(r'^\s*(?:D\s*)?(\d+)/', item_string)
        if extract_ct_tem:
            extract_ct = int(extract_ct_tem.group(1))
            df_GTG_RAW.loc[i, 'only_count'] = extract_ct
    else:
        item_string = df_GTG_RAW['Item Name according to GTG'][i]
        extract_ct_tem = re.search(r'\d+/\d+', item_string)
        if extract_ct_tem:
            extract_ct =extract_ct_tem.group()
            df_GTG_RAW.loc[i, 'only_count'] = extract_ct
    df_GTG_RAW.loc[i,'Item Name according to me']=str(df_GTG_RAW['only_count'][i])+' '+str(df_GTG_RAW['Item type according to me'][i])
# print(df_GTG_RAW)
df_GTG_RAW.to_excel('OUTPUT_converted_GTG_foR_RAW_DB.xlsx', index=True)



####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before
# date_str = datetime.now().strftime("%d-%m-%Y")
# file_name_to_save_all_data="temporary_output_Winding_lead_time_of_"+date_str+".xlsx"
file_name_to_save_all_data='OUTPUT_converted_GTG_foR_RAW_DB.xlsx'
print(file_name_to_save_all_data)
if os.path.isfile(file_name_to_save_all_data):
    os.remove(file_name_to_save_all_data)
# df_result_2.to_excel(file_name_to_save_all_data, index=True)

df_GTG_RAW.to_excel(file_name_to_save_all_data, index=True)

####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before
