import pandas as pd
import matplotlib.pyplot as plt
import re
# from search_by_raw_material_name_function import search_by_raw_material_name_function
import math


file_name_daily_GTG_stock_report = "daily_gtg stock report.xlsx" #sheet name should remain constant as 'Raw'
file_name_combined_dry_weight_DB= "combined_dry_weight_DB.xlsx" ###should not be changed
file_name_OUTPUT_converted_GTG_foR_RAW_DB = "To_convert_GTG_ITEM_NAME_to_match_our_ITEM_NAME/OUTPUT_converted_GTG_foR_RAW_DB.xlsx" ###should not be changed
file_name_purchase_order_lead_time_DB= "purchase_order_lead_time_DB.xlsx" ###should not be changed
file_name_Output_WithOut_ply_ODR_DB= "Item_wise_order_calculation_per_month_in_kg/Output_WithOut_ply_DB.xlsx" ###should not be changed

df_daily_GTG_stock_report_RAW = pd.read_excel(file_name_daily_GTG_stock_report, sheet_name="Raw")
print(df_daily_GTG_stock_report_RAW.columns)
# print(df_daily_GTG_stock_report_RAW.head())

df_combined_dry_weight_DB = pd.read_excel(file_name_combined_dry_weight_DB)
print(df_combined_dry_weight_DB.columns)
# print(df_combined_dry_weight_DB.head())


df_OUTPUT_converted_GTG_foR_RAW_DB = pd.read_excel(file_name_OUTPUT_converted_GTG_foR_RAW_DB)
print(df_OUTPUT_converted_GTG_foR_RAW_DB.columns)
# print(df_OUTPUT_converted_GTG_foR_RAW_DB.head())

df_purchase_order_lead_time_DB = pd.read_excel(file_name_purchase_order_lead_time_DB)
print(df_purchase_order_lead_time_DB.columns)
# print(df_purchase_order_lead_time_DB.head())

df_Output_WithOut_ply_ODR_DB= pd.read_excel(file_name_Output_WithOut_ply_ODR_DB)
print(df_Output_WithOut_ply_ODR_DB.columns)
# print(df_Output_WithOut_ply_ODR_DB.head())



def get_item_mapped_for_GTG_stock_report_RAW(item_name):
    match = df_OUTPUT_converted_GTG_foR_RAW_DB[df_OUTPUT_converted_GTG_foR_RAW_DB['Item Name according to GTG'].astype(str) == str(item_name)]
    if not match.empty:
        return (match.iloc[0]['Item Name according to me'])
    else:
        return 'NOT in DB for : '+ str(item_name)


#_______________________________________
#INITIAL DATA PREPARATION STARTS
#_______________________________________

df_combined_dry_weight_DB['Item Mapped'] = df_combined_dry_weight_DB['Item Mapped'].str.replace(r' BONDED', '', regex=True)
df_combined_dry_weight_DB['Item Name according to me'] = df_combined_dry_weight_DB.apply(
    lambda row: re.sub(r'/\d', '', row['Item Mapped'])
    if str(row['Article No.']).startswith(('2','3','5','6'))
    else row['Item Mapped'],
    axis=1
)
df_combined_dry_weight_DB = df_combined_dry_weight_DB.drop_duplicates(subset=['Item Name according to me'], keep='first')
df_combined_dry_weight_DB_only_NAME=df_combined_dry_weight_DB['Item Name according to me']
# df_combined_dry_weight_DB.to_excel('temp_out.xlsx', index=False)
# df_combined_dry_weight_DB_only_NAME.to_excel('temp_out.xlsx', index=False)


df_daily_GTG_stock_report_RAW['Item Name according to me'] = df_daily_GTG_stock_report_RAW['Item Name'].apply(get_item_mapped_for_GTG_stock_report_RAW)
df_daily_GTG_stock_report_RAW_grouped = df_daily_GTG_stock_report_RAW.groupby(['Item Name according to me']).agg(
        total_Closing_Bal=('Closing Bal.', 'sum')
    ).reset_index()
# print(df_daily_GTG_stock_report_RAW.columns)
# print(df_daily_GTG_stock_report_RAW.head())
# df_daily_GTG_stock_report_RAW_grouped.to_excel('temp_out.xlsx', index=False)

df_GTG_stock_report_RAW_grpd_mergred=df_daily_GTG_stock_report_RAW_grouped.merge(df_combined_dry_weight_DB_only_NAME, on='Item Name according to me', how='outer')
df_GTG_stock_report_RAW_grpd_mergred.to_excel('temp_out1.xlsx', index=False)

df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time=df_GTG_stock_report_RAW_grpd_mergred.merge(df_purchase_order_lead_time_DB, on='Item Name according to me', how='outer')
df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time.to_excel('temp_out2.xlsx', index=False)
#_______________________________________
#INITIAL DATA PREPARATION ENDS
#_______________________________________



#_______________________________________
#MAIN COMPARISONS STARTS <here may need to be changed>
#_______________________________________
df_Output_WithOut_ply_ODR_DB = df_Output_WithOut_ply_ODR_DB.rename(columns={'item mapped': 'Item Name according to me'})
df_Output_WithOut_ply_ODR_DB_new=df_Output_WithOut_ply_ODR_DB[['Item Name according to me','total_order_for_item_per_month_in_kg']]
df_Output_WithOut_ply_ODR_DB_new.to_excel('temp_out3.xlsx', index=False)

#_______________________________________
#MAIN COMPARISONS ENDS <here may need to be changed>
#_______________________________________

#
df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time_ODR=df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time.merge(df_Output_WithOut_ply_ODR_DB_new, on='Item Name according to me', how='outer')
df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time_ODR['Stock_rqd_accdng_to_lead_time']=df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time_ODR['total_order_for_item_per_month_in_kg']/26*df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time_ODR['Lead_time(days)']
df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time_ODR['total_Closing_Bal'] = df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time_ODR['total_Closing_Bal'].fillna(0)
df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time_ODR['Action_required'] = (
    df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time_ODR.apply(
        lambda row:
            'Dont Order' if row['total_Closing_Bal'] > row['Stock_rqd_accdng_to_lead_time']
            else 'Order' if row['total_Closing_Bal'] <= row['Stock_rqd_accdng_to_lead_time']
            else 'No Action',
        axis=1
    )
)
df_GTG_stock_report_RAW_grpd_mergred_WITH_Lead_Time_ODR.to_excel('Output_Purchase_Order.xlsx', index=False)

