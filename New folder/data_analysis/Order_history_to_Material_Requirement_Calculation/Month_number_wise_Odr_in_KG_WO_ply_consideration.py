import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime
import re

month_nos_list=[4,5,6,7,8]






file_name_OrderHistory="OrderHistory.xlsx"
file_name_combined_dry_weight_DB = "combined_dry_weight_DB.xlsx"

# file_name_1="OrderHistory_all.xlsx"


def get_item_mapped_for_OdrHstry(articl):
    match = df_combined_dry_weight_DB[df_combined_dry_weight_DB['Article No.'].astype(str) == str(articl)]
    if not match.empty:
        return (match.iloc[0]['Item Mapped'])
    else:
        return None


def get_dry_wt_for_OdrHstry(articl):
    match = df_combined_dry_weight_DB[df_combined_dry_weight_DB['Article No.'].astype(str) == str(articl)]
    if not match.empty:
        return (match.iloc[0]['Dry weight(gram)'])
    else:
        return None


def month_wrapper(month_no,str_to_wrap):
    return month_dict[month_no]+str_to_wrap



def BW_weight_add(row):
    if row['Color'] == 'BW' or row['Color']  == 'KS-BLEACH':
        return row['Total order in KG']
    else:
        return None













df_OrderHistory = pd.read_excel(file_name_OrderHistory)
print(df_OrderHistory.columns)
# print(df_OrderHistory.head())
df_combined_dry_weight_DB = pd.read_excel(file_name_combined_dry_weight_DB)
print(df_combined_dry_weight_DB.columns)
# print(df_combined_dry_weight_DB.head())






df_combined_dry_weight_DB['Item Mapped'] = df_combined_dry_weight_DB['Item Mapped'].str.replace(r' BONDED', '', regex=True)
df_combined_dry_weight_DB['Item Mapped'] = df_combined_dry_weight_DB.apply(
    lambda row: re.sub(r'/\d', '', row['Item Mapped'])
    if str(row['Article No.']).startswith(('2','3','5','6'))
    else row['Item Mapped'],
    axis=1
)



month_dict={1: 'Jan',2: 'Feb',3: 'March',4: 'April',5: 'May',6: 'June',7: 'July',8: 'Aug',9: 'Sept',10:'Oct',11:'Nov',12: 'Dec'}
df_OrderHistory['Article'] = df_OrderHistory['Article'].astype(str)
df_OrderHistory['Actual order']=df_OrderHistory['Order Qty']-df_OrderHistory['Cancel Qty']
df_OrderHistory['Order Date'] = pd.to_datetime(df_OrderHistory['Order Date'], format='%d/%m/%Y')
df_OrderHistory['Item Mapped'] = df_OrderHistory['Article'].apply(get_item_mapped_for_OdrHstry)
df_OrderHistory['Dry weight(gram)'] = df_OrderHistory['Article'].apply(get_dry_wt_for_OdrHstry)
df_OrderHistory['Total order in KG']=df_OrderHistory['Actual order']*df_OrderHistory['Dry weight(gram)']/1000
# print(df_OrderHistory['Order Date'].dtype)
print(df_OrderHistory.columns)
print(df_OrderHistory.head())
df_month_combined_Odr = pd.DataFrame()
df_month_combined_Odr['Item Mapped']=None
for month_no in month_nos_list:
    df_OrderHistory_t = df_OrderHistory[df_OrderHistory['Order Date'].dt.month == month_no]
    df_OrderHistory_t = df_OrderHistory_t[df_OrderHistory_t['Total order in KG'] != 0]
    df_OrderHistory_t['Date'] = pd.to_datetime(df_OrderHistory_t['Order Date'], format="%d/%m/%Y")
    no_of_days = df_OrderHistory_t['Date'].max() - df_OrderHistory_t['Date'].min()

    df_OrderHistory_t['BW_Odr_KG']=df_OrderHistory_t.apply(BW_weight_add,axis=1)


    df_OrderHistory_t_2 = df_OrderHistory_t.groupby(['Item Mapped']).agg(
        total_Odr_Qnty_in_KG_for_item=('Total order in KG', 'sum'),
        total_BW_Odr_Qnty_in_KG_for_item=('BW_Odr_KG', 'sum')
    ).reset_index()






    Odr_KG_month_no = month_wrapper(month_no, '_Odr_KG')
    BW_Odr_KG_month_no = month_wrapper(month_no, '_BW_Odr_KG')
    df_OrderHistory_t_2 = df_OrderHistory_t_2.rename(columns={
        'total_Odr_Qnty_in_KG_for_item': Odr_KG_month_no,
        'total_BW_Odr_Qnty_in_KG_for_item':BW_Odr_KG_month_no

    })
    df_OrderHistory_t_2_Data_to_show=df_OrderHistory_t_2[['Item Mapped',Odr_KG_month_no,BW_Odr_KG_month_no]]



    df_month_combined_Odr = df_month_combined_Odr.merge(
        df_OrderHistory_t_2_Data_to_show,
        on=['Item Mapped'],
        how='outer'
    )

    temp_file_name = 'temp_out_of_monthly_' + str(month_no) + '.xlsx'

    df_OrderHistory_t_2.to_excel(temp_file_name, index=False)

df_month_combined_Odr.to_excel('OUTPUT_month_wise_Ord(KG)_WO_ply.xlsx', index=False)




















