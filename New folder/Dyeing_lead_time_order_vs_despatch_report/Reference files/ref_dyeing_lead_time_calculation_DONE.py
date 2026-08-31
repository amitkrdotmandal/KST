import pandas as pd
import matplotlib.pyplot as plt
import re
# from search_by_ra import search_by_raw_material_name_function
import math
import os
from datetime import datetime

file_name_daily_dyeing_order_to_despatch_report = "ordertodspsummary.xlsx"
# file_name_daily_dyeing_order_to_despatch_report = "daily_dye_order_to_dispatch_report.xlsx"
file_name_Dyeing_capacity_DB = "Dyeing_capacity_DB.xlsx"  # sheet name should remain constant as 'Raw'


def get_series(Qty):
    match = df_Dyeing_capacity_DB[
        (df_Dyeing_capacity_DB['Corrected Minimum lot size(kg)'] <= Qty) &
        (Qty < df_Dyeing_capacity_DB['Maximum lot size(Kg)'])
        ]
    if not match.empty:
        return match.iloc[0]['Series_name']
    else:
        return None


def get_capacity_from_capacity_to_link(s_name):
    match = df_Dyeing_capacity_DB[df_Dyeing_capacity_DB['Series_name'] == s_name]
    if not match.empty:
        return match.iloc[0]['Total lot per day(24 hours running)']
    else:
        return None


def get_DV_number(s_name):
    match = df_Dyeing_capacity_DB[df_Dyeing_capacity_DB['Series_name'] == s_name]
    if not match.empty:
        return match.iloc[0]['DV number']
    else:
        return None

def get_no_of_pkg_in_DV(s_name):
    match = df_Dyeing_capacity_DB[df_Dyeing_capacity_DB['Series_name'] == s_name]
    if not match.empty:
        return str(match.iloc[0]['DV (no. of pkg)'])+'p'
    else:
        return None


df_daily_dyeing_order_to_despatch_report = pd.read_excel(file_name_daily_dyeing_order_to_despatch_report)
# print(df_daily_dyeing_order_to_despatch_report.head())
# print(df_daily_dyeing_order_to_despatch_report.columns)


#### FILTERING daily_dyeing_order_to_despatch_report
#### FILTERING daily_dyeing_order_to_despatch_report
#### FILTERING daily_dyeing_order_to_despatch_report
#### FILTERING daily_dyeing_order_to_despatch_report

##FILTER 1 Dont allow "Status" as "Close"
df_daily_dyeing_order_to_despatch_report = df_daily_dyeing_order_to_despatch_report[
    df_daily_dyeing_order_to_despatch_report['Status'] != 'Close']
# print(df_daily_dyeing_order_to_despatch_report.head())
# print(df_daily_dyeing_order_to_despatch_report.columns)
##FILTER 1 Dont allow "Status" as "Close"

##FILTER 2 allow Despatch date NaN
df_daily_dyeing_order_to_despatch_report = df_daily_dyeing_order_to_despatch_report[
    df_daily_dyeing_order_to_despatch_report['Dsp. Dt'].isna()]
# print(df_daily_dyeing_order_to_despatch_report['Dsp. Dt'].head(60))
# print(df_daily_dyeing_order_to_despatch_report.columns)
##FILTER 2 allow Despatch date NaN


#### FILTERING daily_dyeing_order_to_despatch_report
#### FILTERING daily_dyeing_order_to_despatch_report
#### FILTERING daily_dyeing_order_to_despatch_report
#### FILTERING daily_dyeing_order_to_despatch_report
# print(df_daily_dyeing_order_to_despatch_report.head())
df_daily_dyeing_order_to_despatch_report.drop(columns=['Name of Party', 'Pck Date', 'Sch. Days','Total Days','Status'], inplace=True)
# print(df_daily_dyeing_order_to_despatch_report.head())
# print(df_daily_dyeing_order_to_despatch_report.columns)

#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
df_Dyeing_capacity_DB = pd.read_excel(file_name_Dyeing_capacity_DB, sheet_name="main")
len_df_Dyeing_capacity_DB = len(df_Dyeing_capacity_DB)
for i in range(len_df_Dyeing_capacity_DB):
    if i == 0:
        df_Dyeing_capacity_DB.loc[i, 'Corrected Minimum lot size(kg)'] = df_Dyeing_capacity_DB['Minimum lot size(kg)'][
            i]
    else:
        df_Dyeing_capacity_DB.loc[i, 'Corrected Minimum lot size(kg)'] = df_Dyeing_capacity_DB['Maximum lot size(Kg)'][
            i - 1]

# print(df_Dyeing_capacity_DB.head(20))
print(df_Dyeing_capacity_DB.columns)
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB


df_daily_dyeing_order_to_despatch_report['Series_name'] = df_daily_dyeing_order_to_despatch_report['D.O. Qty'].apply(
    get_series)
# print(df_daily_dyeing_order_to_despatch_report.head(50))
# print(df_daily_dyeing_order_to_despatch_report.columns)
# df_daily_dyeing_order_to_despatch_report.to_excel('temorary_OUTPUT_updated_dye_order_despatch_report.xlsx', index=True)

result_of_main = df_daily_dyeing_order_to_despatch_report.groupby('Series_name').size().reset_index(
    name='lot_Count_to_produce')

result_of_main['DV number'] = result_of_main['Series_name'].apply(get_DV_number)

result_of_main['Total lot per day(24 hours running)'] = result_of_main['Series_name'].apply(
    get_capacity_from_capacity_to_link)

result_of_main['No_of_PKG_in_DV'] = result_of_main['Series_name'].apply(
    get_no_of_pkg_in_DV)

result_of_main = result_of_main.drop(columns=['Series_name'])

result_of_main['Estimated_lead_time_in_days'] = result_of_main['lot_Count_to_produce'] / result_of_main[
    'Total lot per day(24 hours running)']
result_of_main['Estimated_lead_time_in_days'] = result_of_main['Estimated_lead_time_in_days'].round(2)

col_tem = result_of_main.pop('DV number')
result_of_main.insert(0, 'DV number', col_tem)

col_tem = result_of_main.pop('No_of_PKG_in_DV')
result_of_main.insert(1, 'No_of_PKG_in_DV', col_tem)

print(result_of_main)






####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before
date_str = datetime.now().strftime("%d-%m-%Y")
# file_name_to_save_all_data="temporary_output_Winding_lead_time_of_"+date_str+".xlsx"
file_name_to_save_all_data="temporary_output_dyeing_lead_time_of_"+".xlsx"
print(file_name_to_save_all_data)
if os.path.isfile(file_name_to_save_all_data):
    os.remove(file_name_to_save_all_data)


with pd.ExcelWriter('Output_dyeing_lead_time.xlsx') as Writter:
    result_of_main.to_excel(Writter, sheet_name='OverAll', startrow=2, index=False)

    worksheet = Writter.sheets["OverAll"]
    worksheet["A1"] = "Dyeing_lead_time_of_" + date_str
    for i in range(len_df_Dyeing_capacity_DB):
        to_coompare = df_Dyeing_capacity_DB['Series_name'][i]
        sheet_name = "DV_no - " + str(df_Dyeing_capacity_DB['DV number'][i])
        df_tempo = df_daily_dyeing_order_to_despatch_report[
            df_daily_dyeing_order_to_despatch_report['Series_name'] == to_coompare]
        # df_tempo.drop(columns=['Series_name'], inplace=True)
        df_tempo.to_excel(Writter, sheet_name=sheet_name, index=False)

####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before