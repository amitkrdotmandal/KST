import pandas as pd
import matplotlib.pyplot as plt
import re
import math


file_name_jobcard_pkg_streaching = "jobcard_pkg_stretching.xlsx"
file_name_jobcard_pkg_winding = "jobcard_pkg_winding.xlsx"
file_name_rptstapprovalpending = "rptstapprovalpending.xlsx"
# file_name_ProductionPlanning = "ProductionPlanning.xlsx" #sheet name should remain constant as 'Raw'

file_name_DB_of_converted_item_name = "To_convert_after_dyeing_ITEM_NAME_to_match_our_ITEM_NAME/OUTPUT_converted_item_name_of_dyed_soft_pkg_DB.xlsx" ###should not be changed


#### converted item name
#### converted item name
#### converted item name
#### converted item name
df_DB_of_converted_item_name = pd.read_excel(file_name_DB_of_converted_item_name)
print(df_DB_of_converted_item_name.columns)
# print(df_DB_of_converted_item_name.head())
#### converted item name
#### converted item name
#### converted item name
#### converted item name




#### jobcard_pkg_streaching
#### jobcard_pkg_streaching
#### jobcard_pkg_streaching
#### jobcard_pkg_streaching
print('To see in jobcard_pkg_winding : ')
df_jobcard_pkg_streaching = pd.read_excel(file_name_jobcard_pkg_streaching)
print(df_jobcard_pkg_streaching.columns)
# print(df_jobcard_pkg_streaching.head())
missing_codes = df_jobcard_pkg_streaching.loc[
    ~df_jobcard_pkg_streaching['Item Name'].isin(df_DB_of_converted_item_name['Item Name  aacording to dyeing']),
    'Item Name'].tolist()
len_missing_codes=len(missing_codes)
if len_missing_codes==0:
    print('All the items are in data base')
else:
    print('following Item Code details have to be updated manually in database "To_convert_after_dyeing_ITEM_NAME_to_match_our_ITEM_NAME/Dyed_soft_package_Item_type_conversion_DB.xlsx"')
    print(missing_codes)

#### jobcard_pkg_streaching
#### jobcard_pkg_streaching
#### jobcard_pkg_streaching
#### jobcard_pkg_streaching
# print(missing_codes.head())





#### jobcard_pkg_winding
#### jobcard_pkg_winding
#### jobcard_pkg_winding
#### jobcard_pkg_winding
# del missing_codes
print('To see in jobcard_pkg_streaching : ')
df_jobcard_pkg_winding = pd.read_excel(file_name_jobcard_pkg_winding)
print(df_jobcard_pkg_winding.columns)
# print(df_jobcard_pkg_winding.head())
missing_codes = df_jobcard_pkg_winding.loc[
    ~df_jobcard_pkg_winding['Item Name'].isin(df_DB_of_converted_item_name['Item Name  aacording to dyeing']),
    'Item Name'].tolist()

len_missing_codes=len(missing_codes)
if len_missing_codes==0:
    print('All the items are in data base')
else:
    print('following Item Code details have to be updated manually in database "To_convert_after_dyeing_ITEM_NAME_to_match_our_ITEM_NAME/Dyed_soft_package_Item_type_conversion_DB.xlsx"')
    print(missing_codes)

#### jobcard_pkg_winding
#### jobcard_pkg_winding
#### jobcard_pkg_winding
#### jobcard_pkg_winding


#### file_name_rptstapprovalpending
#### file_name_rptstapprovalpending
#### file_name_rptstapprovalpending
#### file_name_rptstapprovalpending

print('To see in rptstapprovalpending  : ')
df_rptstapprovalpending = pd.read_excel(file_name_rptstapprovalpending)
print(df_rptstapprovalpending.columns)
# print(df_rptstapprovalpending.head())
missing_codes = df_rptstapprovalpending.loc[
    ~df_rptstapprovalpending['Finish Item'].isin(df_DB_of_converted_item_name['Item Name  aacording to dyeing']),
    'Finish Item'].tolist()
missing_codes_job_card_no = df_rptstapprovalpending.loc[
    ~df_rptstapprovalpending['Finish Item'].isin(df_DB_of_converted_item_name['Item Name  aacording to dyeing']),
    'JobCard No.'].tolist()
len_missing_codes=len(missing_codes)
if len_missing_codes==0:
    print('All the items are in data base')
else:
    print('following Item Code details have to be updated manually in database "To_convert_after_dyeing_ITEM_NAME_to_match_our_ITEM_NAME/Dyed_soft_package_Item_type_conversion_DB.xlsx"')
    print(missing_codes)
    print(missing_codes_job_card_no)
# df_rptstapprovalpending.to_excel('temp_out.xlsx', index=False)
#### file_name_rptstapprovalpending
#### file_name_rptstapprovalpending
#### file_name_rptstapprovalpending
#### file_name_rptstapprovalpending































