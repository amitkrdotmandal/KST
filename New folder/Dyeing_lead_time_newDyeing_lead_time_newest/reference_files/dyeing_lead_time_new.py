import pandas as pd
import matplotlib.pyplot as plt
import re
# from search_by_ra import search_by_raw_material_name_function
import math
import os
from datetime import datetime
from datetime import timedelta
from customer_analysis_from_order_history_func import customer_analysis_from_order_history_func

file_name_pendingordertoplan = "pendingordertoplan.xlsx"
# file_name_daily_dyeing_order_to_despatch_report = "daily_dye_order_to_dispatch_report.xlsx"
file_name_Dyeing_capacity_DB = "Dyeing_capacity_DB.xlsx"  # sheet name should remain constant as 'Raw'
file_name_DB_of_converted_item_name = "To_convert_after_dyeing_ITEM_NAME_to_match_our_ITEM_NAME/OUTPUT_converted_item_name_of_dyed_soft_pkg_DB.xlsx"  ###should not be changed






def get_series(Qty):
    match = df_Dyeing_capacity_DB[
        (df_Dyeing_capacity_DB['Corrected Minimum lot size(kg)'] <= Qty) &
        (Qty < df_Dyeing_capacity_DB['Maximum lot size(Kg)'])
        ]
    if not match.empty:
        return match.iloc[0]['Series_name']
    else:
        return None


def get_series_p(Qty):
    match = df_Dyeing_capacity_DB[
        (df_Dyeing_capacity_DB['Corrected Minimum lot size(kg)'] <= Qty) &
        (Qty < df_Dyeing_capacity_DB['Maximum lot size(Kg)'])
        ]
    if not match.empty:
        return match.iloc[0]['Series_name_p']
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


def get_item_for_me_for_streaching_or_winding_or_pending(item):
    match = df_DB_of_converted_item_name[
        df_DB_of_converted_item_name['Item Name  aacording to dyeing'].astype(str) == item]
    if not match.empty:
        return (match.iloc[0]['Item Name according to me'])
    else:
        return None


def dyeing_can_be_done_in_KG(row):
    now = datetime.now()
    if (row['Total_Frequency']>1) and (row['Total_Frequency']>1) and (row['Last_date_of_Odr']>(now - timedelta(days=30))) and (row['Avg_wt_in_KG_per_month']>row['Bal PlanQty']):
        return row['Avg_wt_in_KG_per_month']
    else:
        return None








df_pendingordertoplan = pd.read_excel(file_name_pendingordertoplan)
print(df_pendingordertoplan.head(50))
print(df_pendingordertoplan.columns)





df_pendingordertoplan = df_pendingordertoplan[df_pendingordertoplan['Order No.'] != 'Item  Total ']
group = None
groups = []

for _, row in df_pendingordertoplan.iterrows():
    # If DO No. is blank, this row is a group/header row
    if pd.isna(row['DO  No.']):
        group = row['Order No.']
    groups.append(group)
df_pendingordertoplan['Item Name'] = groups
df_pendingordertoplan = df_pendingordertoplan[
    df_pendingordertoplan['DO  No.'].notna()
].reset_index(drop=True)

# df_pendingordertoplan.to_excel('tempo_out1.xlsx', index=False)


#=================================================
#starts
#=================================================
df_DB_of_converted_item_name = pd.read_excel(file_name_DB_of_converted_item_name)
print(df_DB_of_converted_item_name.columns)
# print(df_DB_of_converted_item_name.head())

#=================================================
#ends
#=================================================

#=================================================
#starts
#=================================================
start_day = '01/04/2026'

# This party list is being for shortening the party name and make the group common under one name
parties_to_be_analysed = ['zigma', 'safira', 'jak', 'T R THREADS', 'hillson', 'SI INTERPACK',
                          'K.K. HOSIERY MATERIAL STORE', 'CHIRAG CREATIONS', 'FANTAIL']
# This party list is being for shortening the party name and make the group common under one name

df_OrderHistory_grouped2,start_day, max_date=customer_analysis_from_order_history_func(start_day,parties_to_be_analysed)
print('df_OrderHistory_grouped2',df_OrderHistory_grouped2.columns)
# print(df_OrderHistory_grouped2.head())

#=================================================
#ends
#=================================================






#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
df_Dyeing_capacity_DB = pd.read_excel(file_name_Dyeing_capacity_DB, sheet_name="main")
len_df_Dyeing_capacity_DB = len(df_Dyeing_capacity_DB)
for i in range(len_df_Dyeing_capacity_DB):
    df_Dyeing_capacity_DB.loc[i, 'Series_name_p'] = str(df_Dyeing_capacity_DB['DV (no. of pkg)'][
        i])+'p'

    if i == 0:
        df_Dyeing_capacity_DB.loc[i, 'Corrected Minimum lot size(kg)'] = df_Dyeing_capacity_DB['Minimum lot size(kg)'][
            i]
    else:
        df_Dyeing_capacity_DB.loc[i, 'Corrected Minimum lot size(kg)'] = df_Dyeing_capacity_DB['Maximum lot size(Kg)'][
            i - 1]
print(df_Dyeing_capacity_DB.head(20))
print(df_Dyeing_capacity_DB.columns)
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB
#### Include corrected minimum in CAPACITY DB


df_pendingordertoplan['Item Name'] = df_pendingordertoplan['Item Name'].apply(
        get_item_for_me_for_streaching_or_winding_or_pending)


df_OrderHistory_grouped2.rename(columns={'Item Mapped': 'Item Name'}, inplace=True)
df_OrderHistory_grouped2.rename(columns={'Color': 'Shade Name'}, inplace=True)
print('df_OrderHistory_grouped2.columns',df_OrderHistory_grouped2.columns)
print('df_pendingordertoplan.columns',df_pendingordertoplan.columns)


df_pendingordertoplan = pd.merge(
    df_pendingordertoplan,
    df_OrderHistory_grouped2,
    on=['Item Name', 'Shade Name'],
    how='left'
)

df_pendingordertoplan['No_of_PKG_in_DV'] = df_pendingordertoplan['Bal PlanQty'].apply(
    get_series_p)
df_pendingordertoplan['Series_name'] = df_pendingordertoplan['Bal PlanQty'].apply(
    get_series)
df_pendingordertoplan = df_pendingordertoplan.sort_values(by='Shade Name', ascending=False)


print('max_date:',max_date)
print('start_day:',start_day)
print('Days:',(max_date-start_day).days)

df_pendingordertoplan['Avg_Qty_per_month']=df_pendingordertoplan['actual_tot_in_Qty']/(max_date-start_day).days*30
df_pendingordertoplan['Avg_wt_in_KG_per_month']=df_pendingordertoplan['actual_tot_wt_in_KG']/(max_date-start_day).days*30
df_pendingordertoplan['Avg_Frequency_per_month']=df_pendingordertoplan['Total_Frequency']/(max_date-start_day).days*30

df_pendingordertoplan['Dyeing_can_be_done_in_KG'] = df_pendingordertoplan.apply(dyeing_can_be_done_in_KG, axis=1)










result_of_main = df_pendingordertoplan.groupby('Series_name').size().reset_index(
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




df_pendingordertoplan.to_excel('Output_files/Output_all_together_Pend_Odr_Pln.xlsx', index=False)



####For saving Output_dyeing_lead_time_new starts
####For saving Output_dyeing_lead_time_new starts
####For saving Output_dyeing_lead_time_new starts
date_str = datetime.now().strftime("%d-%m-%Y")
# file_name_to_save_all_data="temporary_output_Winding_lead_time_of_"+date_str+".xlsx"
file_name_to_save_all_data="Output_files/Output_dyeing_lead_time_new_Pend_Odr_Pln"+".xlsx"
print(file_name_to_save_all_data)
if os.path.isfile(file_name_to_save_all_data):
    os.remove(file_name_to_save_all_data)


with pd.ExcelWriter(file_name_to_save_all_data) as Writter:
    result_of_main.to_excel(Writter, sheet_name='OverAll', startrow=2, index=False)

    worksheet = Writter.sheets["OverAll"]
    worksheet["A1"] = "Dyeing_lead_time_of_" + date_str
    for i in range(len_df_Dyeing_capacity_DB):
        to_coompare = df_Dyeing_capacity_DB['Series_name'][i]
        sheet_name = "DV_no - " + str(df_Dyeing_capacity_DB['DV number'][i])
        df_tempo = df_pendingordertoplan[
            df_pendingordertoplan['Series_name'] == to_coompare]
        # df_tempo.drop(columns=['Series_name'], inplace=True)
        df_tempo.to_excel(Writter, sheet_name=sheet_name, index=False)

####For saving Output_dyeing_lead_time_new ends
####For saving Output_dyeing_lead_time_new ends
####For saving Output_dyeing_lead_time_new ends


#============================================
# item vs sheet vs Article and Shade starts
#==========================================

item_list = df_pendingordertoplan['Item Name'].astype(str).unique().tolist()

print(item_list)
with pd.ExcelWriter('Output_files/Output_item_wise_sheet_vs_Article_Shade_Pend_Odr_Pln.xlsx') as Writter:
    for item in item_list:
        filtered_df = df_pendingordertoplan[
            df_pendingordertoplan['Item Name']==item]

        #======================================================================
        #See repetation starts
        #======================================================================
        count = filtered_df.groupby('Shade Name')['Shade Name'].transform('count')
        number = filtered_df.groupby('Shade Name').cumcount() + 1
        filtered_df['Repetition_of_shade'] = number.astype(str) + ' by ' + count.astype(str)
        #======================================================================
        #See repetation ends
        #======================================================================

        filtered_df.drop(columns=['Party Name.', 'actual_tot_in_Qty', 'actual_tot_wt_in_KG','Parties','Avg_Qty_per_month','Series_name'], inplace=True)
        filtered_df = filtered_df.rename(columns={'Last_date_of_Odr': 'Last_date_of_customer_odr'})

        filtered_df = filtered_df.sort_values(by='Shade Name', ascending=True).reset_index(drop=True)
        item = item.replace('/', ' by ')
        item = item.replace(' ', '_')

        sheet_name = item
        filtered_df.to_excel(Writter, sheet_name=sheet_name, index=False)

#============================================
# item vs sheet vs Article and Shade ends
#==========================================