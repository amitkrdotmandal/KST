import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime
import re


No_of_last_months_to_analyze = int(input("Enter the number of last months to analyze to get avg. consumption per month according to Make_up: "))
current_month = int(datetime.now().strftime("%m"))
month_nos_list = []
for i in range(No_of_last_months_to_analyze):
    member=(current_month - i-1) % 12 or 12
    month_nos_list.append(member)

print(month_nos_list)
print(current_month)

# month_nos_list = [6,7,8]







file_name_OrderHistory="OrderHistory.xlsx"
file_name_combined_dry_weight_DB = "combined_dry_weight_DB.xlsx"


def month_wrapper(month_no,str_to_wrap):
    return month_dict[month_no]+str_to_wrap





df_OrderHistory = pd.read_excel(file_name_OrderHistory)
print(df_OrderHistory.columns)
# print(df_OrderHistory.head())
df_combined_dry_weight_DB = pd.read_excel(file_name_combined_dry_weight_DB)
print(df_combined_dry_weight_DB.columns)
# print(df_combined_dry_weight_DB.head())




month_dict={1: 'Jan',2: 'Feb',3: 'March',4: 'April',5: 'May',6: 'June',7: 'July',8: 'Aug',9: 'Sept',10:'Oct',11:'Nov',12: 'Dec'}
df_combined_dry_weight_DB.rename(columns={'Article No.': 'Article'}, inplace=True)
df_OrderHistory['Actual order']=df_OrderHistory['Order Qty']-df_OrderHistory['Cancel Qty']
df_OrderHistory['Order Date'] = pd.to_datetime(df_OrderHistory['Order Date'], format='%d/%m/%Y')
df_OrderHistory_merged = pd.merge(df_OrderHistory, df_combined_dry_weight_DB, on='Article', how='left')
df_OrderHistory_merged['actual_tot_wt_in_KG']=df_OrderHistory_merged['Actual order']*df_OrderHistory_merged['Dry weight(gram)']/1000
print(df_OrderHistory_merged.columns)


df_month_combined_Odr = pd.DataFrame()
df_month_combined_Odr['Make Up']=None
for month_no in month_nos_list:
    df_OrderHistory_merged_t = df_OrderHistory_merged[df_OrderHistory_merged['Order Date'].dt.month == month_no]

    df_OrderHistory_merged_t_2 = df_OrderHistory_merged_t.groupby(['Make Up']).agg(
        actual_tot_wt_in_KG=('actual_tot_wt_in_KG', 'sum'),
        Actual_order_Qty=('Actual order', 'sum')
    ).reset_index()




    actual_tot_wt_in_KG_month_no = month_wrapper(month_no, '_actual_tot_wt_in_KG')
    Actual_order_Qty_month_no = month_wrapper(month_no, '_Actual_order_Qty')
    df_OrderHistory_merged_t_2 = df_OrderHistory_merged_t_2.rename(columns={
        'actual_tot_wt_in_KG': actual_tot_wt_in_KG_month_no,
        'Actual_order_Qty': Actual_order_Qty_month_no
    })

    df_OrderHistory_t_2_Data_to_show = df_OrderHistory_merged_t_2[['Make Up', actual_tot_wt_in_KG_month_no, Actual_order_Qty_month_no]]
    df_month_combined_Odr = df_month_combined_Odr.merge(
        df_OrderHistory_t_2_Data_to_show,
        on=['Make Up'],
        how='outer'
    )


    temp_file_name = 'temp_out_of_monthly_' + str(month_no) + '.xlsx'

    # df_OrderHistory_merged_t_2.to_excel(temp_file_name, index=False)



num_of_columns_in_df_month_combined_Odr=len(df_month_combined_Odr.columns)


df_month_combined_Odr['Total_order_in_KG_of_all_months'] = 0
for i in range(1, num_of_columns_in_df_month_combined_Odr, 2):
    print(i)
    # print(df_month_combined_Odr.iloc[:, i])
    df_month_combined_Odr['Total_order_in_KG_of_all_months'] = (
            df_month_combined_Odr['Total_order_in_KG_of_all_months'].fillna(0)
            + df_month_combined_Odr.iloc[:, i].fillna(0)
    )
df_month_combined_Odr['Avg_order_in_KG_par_month'] =df_month_combined_Odr['Total_order_in_KG_of_all_months']/((num_of_columns_in_df_month_combined_Odr-1)/2)




df_month_combined_Odr['Total_order_in_Qty_of_all_months'] = 0
for i in range(2, num_of_columns_in_df_month_combined_Odr, 2):
    print(i)
    # print(df_month_combined_Odr.iloc[:, i])
    df_month_combined_Odr['Total_order_in_Qty_of_all_months'] = (
            df_month_combined_Odr['Total_order_in_Qty_of_all_months'].fillna(0)
            + df_month_combined_Odr.iloc[:, i].fillna(0)
    )
df_month_combined_Odr['Avg_order_in_Qty_par_month'] =df_month_combined_Odr['Total_order_in_Qty_of_all_months']/((num_of_columns_in_df_month_combined_Odr-1)/2)




df_month_combined_Odr.to_excel('Output_Make_Up_vs_Odr_Qty_&_KG.xlsx', index=False)











