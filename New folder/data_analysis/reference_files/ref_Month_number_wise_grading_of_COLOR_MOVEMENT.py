import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime
import re

month_nos_list=[1,2,3,4,5,6,7]
show_histogram=1 # 1 or 0
save_figure=1 # 1 or 0



file_name_OrderHistory="OrderHistory.xlsx"
file_name_combined_dry_weight_DB = "combined_dry_weight_DB.xlsx"

# file_name_1="OrderHistory_all.xlsx"


def get_item_mapped_for_ProdPlan(articl):
    match = df_combined_dry_weight_DB[df_combined_dry_weight_DB['Article No.'].astype(str) == str(articl)]
    if not match.empty:
        return (match.iloc[0]['Item Mapped'])
    else:
        return None


def get_dry_wt_for_ProdPlan(articl):
    match = df_combined_dry_weight_DB[df_combined_dry_weight_DB['Article No.'].astype(str) == str(articl)]
    if not match.empty:
        return (match.iloc[0]['Dry weight(gram)'])
    else:
        return None

def monthly_grading(row):
    if row['Frqcy_Odr_Qnty']>2 and row['Percentage_of_odr_WRT_item_total']>5:
        return 'Grade A'
    elif row['Frqcy_Odr_Qnty']>2 and row['Percentage_of_odr_WRT_item_total']>3 and row['Percentage_of_odr_WRT_item_total']<=5:
        return 'Grade B'
    elif row['Frqcy_Odr_Qnty']<=2 and row['Percentage_of_odr_WRT_item_total']>3:
        return 'Grade C'
    else:
        return 'Grade D'

def month_wrapper(month_no,str_to_wrap):
    return month_dict[month_no]+str_to_wrap












df_OrderHistory = pd.read_excel(file_name_OrderHistory)
print(df_OrderHistory.columns)
# print(df_OrderHistory.head())
df_combined_dry_weight_DB = pd.read_excel(file_name_combined_dry_weight_DB)
print(df_combined_dry_weight_DB.columns)
# print(df_combined_dry_weight_DB.head())











month_dict={1: 'Jan',2: 'Feb',3: 'March',4: 'April',5: 'May',6: 'June',7: 'July',8: 'Aug',9: 'Sept',10:'Oct',11:'Nov',12: 'Dec'}
df_OrderHistory['Article'] = df_OrderHistory['Article'].astype(str)
df_OrderHistory['Actual order']=df_OrderHistory['Order Qty']-df_OrderHistory['Cancel Qty']
df_OrderHistory['Order Date'] = pd.to_datetime(df_OrderHistory['Order Date'], format='%d/%m/%Y')
df_OrderHistory['Item Mapped'] = df_OrderHistory['Article'].apply(get_item_mapped_for_ProdPlan)
df_OrderHistory['Dry weight(gram)'] = df_OrderHistory['Article'].apply(get_dry_wt_for_ProdPlan)
df_OrderHistory['Total order in KG']=df_OrderHistory['Actual order']*df_OrderHistory['Dry weight(gram)']/1000
# print(df_OrderHistory['Order Date'].dtype)
print(df_OrderHistory.columns)
print(df_OrderHistory.head())
df_combined_grading = pd.DataFrame()
df_combined_grading['Item Mapped']=None
df_combined_grading['Color']=None
for month_no in month_nos_list:
    df_OrderHistory_t = df_OrderHistory[df_OrderHistory['Order Date'].dt.month == month_no]
    df_OrderHistory_t = df_OrderHistory_t[df_OrderHistory_t['Total order in KG'] != 0]
    histogram_month_no =month_wrapper(month_no, '_Odr(KG)_histogram')
    if show_histogram==1:
        plt.hist(df_OrderHistory_t['Total order in KG'], bins=400)
        plt.xlabel('Order in KG')
        plt.ylabel('Frequency')
        plt.title(histogram_month_no)
        plt.xlim(0, 200)
        plt.grid(True, linestyle='--', alpha=0.5)
        if save_figure == 1:
            to_save_figure_name = 'Save figure/' + histogram_month_no + '.png'
            plt.savefig(to_save_figure_name, dpi=300, bbox_inches='tight')
        plt.show()



    df_OrderHistory_t_2 = df_OrderHistory_t.groupby(['Item Mapped', 'Color']).agg(
        total_Odr_Qnty_in_KG=('Total order in KG', 'sum'),
        avg_in_KG=('Total order in KG', 'mean'),
        STD_in_KG=('Total order in KG', 'std'),
        Frqcy_Odr_Qnty=('Total order in KG', 'count'),
        Batches=('Order No.', ', '.join)
    ).reset_index()
    df_OrderHistory_t_2['CV_%_of_Odr'] =df_OrderHistory_t_2['STD_in_KG']/df_OrderHistory_t_2['avg_in_KG']*100
    df_OrderHistory_t_3 = df_OrderHistory_t.groupby(['Item Mapped']).agg(
        total_Odr_Qnty_in_KG_for_item=('Total order in KG', 'sum'),
    ).reset_index()
    df_OrderHistory_t_2 = df_OrderHistory_t_2.merge(
        df_OrderHistory_t_3,
        on='Item Mapped',
        how='left'
    )
    df_OrderHistory_t_2['Percentage_of_odr_WRT_item_total']=df_OrderHistory_t_2['total_Odr_Qnty_in_KG']/df_OrderHistory_t_2['total_Odr_Qnty_in_KG_for_item']*100

    # grade_month_no='grading_of_month_'+str(month_no)
    # Percentage_Ord_month_no='% of_Odr_of_month_'+str(month_no)
    # Odr_freq_no = 'Frqcy_of_Odr_of_month_' + str(month_no)
    # Order_in_KG_month_no='Order_in_KG_of_month_'+str(month_no)
    # Avg_Odr_month_no='Avg_Odr_of_month_'+str(month_no)
    # CV_of_Odr_month_no='CV_of_Odr_of_month_'+str(month_no)

    grade_month_no = month_wrapper(month_no,'_grade')
    Percentage_Ord_month_no = month_wrapper(month_no,'_%_of_Odr')
    Odr_freq_no = month_wrapper(month_no,'_Odr_Frqcy')
    Order_in_KG_month_no = month_wrapper(month_no,'_Odr(KG)')
    Avg_Odr_month_no = month_wrapper(month_no,'_Avg_Odr(KG)')
    CV_of_Odr_month_no = month_wrapper(month_no,'_CV_of_Odr')


    df_OrderHistory_t_2[grade_month_no] =df_OrderHistory_t_2.apply(monthly_grading,axis=1)

    df_OrderHistory_t_2 = df_OrderHistory_t_2.rename(columns={
        'Frqcy_Odr_Qnty': Odr_freq_no,
        'Percentage_of_odr_WRT_item_total': Percentage_Ord_month_no,
        'total_Odr_Qnty_in_KG': Order_in_KG_month_no,
        'avg_in_KG': Avg_Odr_month_no,
        'CV_%_of_Odr': CV_of_Odr_month_no
    })

    # df_OrderHistory_t_graded=df_OrderHistory_t_2[['Item Mapped', 'Color',Order_in_KG_month_no,Percentage_Ord_month_no,Odr_freq_no,grade_month_no]]

    df_OrderHistory_t_graded=df_OrderHistory_t_2[['Item Mapped', 'Color',Order_in_KG_month_no,Avg_Odr_month_no,CV_of_Odr_month_no,Odr_freq_no]]

    df_combined_grading = df_combined_grading.merge(
        df_OrderHistory_t_graded,
        on=['Item Mapped','Color'],
        how='outer'
    )




    temp_file_name='temp_out_of_month_'+str(month_no)+'.xlsx'

    # df_OrderHistory_t.to_excel(temp_file_name, index=False)


df_combined_grading.to_excel('OUTPUT_all_graded.xlsx', index=False)

unique_value_list = df_combined_grading['Item Mapped'].unique().tolist()

# for i in range(len(unique_value_list)):











####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before
date_str = datetime.now().strftime("%d-%m-%Y")
# file_name_to_save_all_data="temporary_output_Winding_lead_time_of_"+date_str+".xlsx"
file_name_to_save_all_data="OUTPUT_graded_ITEM_wise_segregation"+".xlsx"
print(file_name_to_save_all_data)
if os.path.isfile(file_name_to_save_all_data):
    os.remove(file_name_to_save_all_data)


with pd.ExcelWriter(file_name_to_save_all_data) as Writter:
    # result_of_main.to_excel(Writter, sheet_name='OverAll', startrow=2, index=False)

    # worksheet = Writter.sheets["OverAll"]
    # worksheet["A1"] = "Dyeing_lead_time_of_" + date_str
    for i in range(len(unique_value_list)):

        sheet_name = unique_value_list[i]
        sheet_name = re.sub(r'/', 'by', sheet_name)
        df_tempo = df_combined_grading[
            df_combined_grading['Item Mapped'] == unique_value_list[i]]
        # df_tempo.drop(columns=['Series_name'], inplace=True)
        try:
            df_tempo.to_excel(Writter, sheet_name=sheet_name, index=False)
        except:
            print('Error to save:',sheet_name)

####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before



















