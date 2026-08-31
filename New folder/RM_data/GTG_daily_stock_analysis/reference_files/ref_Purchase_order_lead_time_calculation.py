import pandas as pd
import matplotlib.pyplot as plt
import re
from search_by_raw_material_name_function import search_by_raw_material_name_function
import math


file_name_daily_GTG_stock_report = "daily_gtg stock report.xlsx" #sheet name should remain constant as 'Raw'
file_name_DB_Output_WithOut_ply_prod = "Item_wise_order_calculation_per_month_in_kg/Output_WithOut_ply_DB.xlsx" ###should not be changed
file_name_OUTPUT_converted_GTG_foR_RAW_DB = "To_convert_GTG_ITEM_NAME_to_match_our_ITEM_NAME/OUTPUT_converted_GTG_foR_RAW_DB.xlsx" ###should not be changed


df_daily_GTG_stock_report_RAW = pd.read_excel(file_name_daily_GTG_stock_report, sheet_name="Raw")
# print(df_daily_GTG_stock_report_RAW.columns)
# print(df_daily_GTG_stock_report_RAW.head())


df_DB_of_converted_item_name = pd.read_excel(file_name_DB_of_converted_item_name)
# print(df_DB_of_converted_item_name.columns)
# print(df_DB_of_converted_item_name.head())


