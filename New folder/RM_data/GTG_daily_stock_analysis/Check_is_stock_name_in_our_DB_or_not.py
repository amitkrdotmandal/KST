import pandas as pd
import matplotlib.pyplot as plt
import re
# from search_by_raw_material_name_function import search_by_raw_material_name_function
import math


file_name_daily_GTG_stock_report = "daily_gtg stock report.xlsx" #sheet name should remain constant as 'Raw'
# file_name_daily_GTG_stock_report = "daily_gtg stock report.xlsx" #sheet name should remain constant as 'Raw'
file_name_DB_of_converted_item_name = "To_convert_GTG_ITEM_NAME_to_match_our_ITEM_NAME/OUTPUT_converted_GTG_foR_RAW_DB.xlsx" ###should not be changed


df_daily_GTG_stock_report_RAW = pd.read_excel(file_name_daily_GTG_stock_report, sheet_name="Raw")
# print(df_daily_GTG_stock_report_RAW.columns)
# print(df_daily_GTG_stock_report_RAW.head())


df_DB_of_converted_item_name = pd.read_excel(file_name_DB_of_converted_item_name)
# print(df_DB_of_converted_item_name.columns)
# print(df_DB_of_converted_item_name.head())

missing_codes = df_daily_GTG_stock_report_RAW.loc[
    ~df_daily_GTG_stock_report_RAW['Item Code'].isin(df_DB_of_converted_item_name['Item Code']),
    'Item Code'].tolist()

len_missing_codes=len(missing_codes)
if len_missing_codes==0:
    print('All the items are in data base')
else:
    print('following Item Code details have to be updated manually in database "To_convert_GTG_ITEM_NAME_to_match_our_ITEM_NAME/GTG_Item_type_converion_DB.xlsx"')
    print(missing_codes)
# print(missing_codes.head())
