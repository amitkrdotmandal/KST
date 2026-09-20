import pandas as pd
import matplotlib.pyplot as plt
import re
# from search_by_raw_material_name_function import search_by_raw_material_name_function
import math
import numpy as np
import os

file_name_jobcard_pkg_streaching = "jobcard_pkg_streaching.xlsx"
file_name_jobcard_pkg_winding = "jobcard_pkg_winding.xlsx"
file_name_rptstapprovalpending = "rptstapprovalpending.xlsx"
file_name_ProductionPlanning = "ProductionPlanning.xlsx"

file_name_DB_of_converted_item_name = "To_convert_after_dyeing_ITEM_NAME_to_match_our_ITEM_NAME/OUTPUT_converted_item_name_of_dyed_soft_pkg_DB.xlsx"  ###should not be changed
file_name_combined_dry_weight_DB = "combined_dry_weight_DB.xlsx"

# nearly_same_group_count=[['2/60 SPUN POLYESTER YARN','2/57 SPUN POLYESTER YARN'],['250/3 HT CONTINOUS POLYESTER FILAMENT', '210/3 HT CONTINOUS POLYESTER FILAMENT','840/3 HI-TENACITY NYLON 66 BONDED']]
nearly_same_group_count=[['2/60 SPUN POLYESTER YARN','2/57 SPUN POLYESTER YARN']]


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


def get_item_for_me_for_streaching_or_winding_or_pending(item):
    match = df_DB_of_converted_item_name[
        df_DB_of_converted_item_name['Item Name  aacording to dyeing'].astype(str) == item]
    if not match.empty:
        return (match.iloc[0]['Item Name according to me'])
    else:
        return None


#### converted item name and prd.
#### converted item name and prd.
#### converted item name and prd.
#### converted item name and prd.
df_ProductionPlanning = pd.read_excel(file_name_ProductionPlanning)
print(df_ProductionPlanning.columns)
# print(df_ProductionPlanning['Article'].head())
# print(df_ProductionPlanning['Article'].dtype)


df_DB_of_converted_item_name = pd.read_excel(file_name_DB_of_converted_item_name)
print(df_DB_of_converted_item_name.columns)
# print(df_DB_of_converted_item_name.head())

df_combined_dry_weight_DB = pd.read_excel(file_name_combined_dry_weight_DB)
print(df_combined_dry_weight_DB.columns)
# print(df_combined_dry_weight_DB['Article No.'].head())
# print(df_combined_dry_weight_DB['Article No.'].dtype)
#### converted item name and prd.
#### converted item name and prd.
#### converted item name and prd.
#### converted item name and prd.


#### jobcard_pkg_streaching, winding and pending
#### jobcard_pkg_streaching, winding and pending
#### jobcard_pkg_streaching, winding and pending
#### jobcard_pkg_streaching, winding and pending
print('streaching')
df_jobcard_pkg_streaching = pd.read_excel(file_name_jobcard_pkg_streaching)
print(df_jobcard_pkg_streaching.columns)
# print(df_jobcard_pkg_streaching.head())
print('winding')
df_jobcard_pkg_winding = pd.read_excel(file_name_jobcard_pkg_winding)
print(df_jobcard_pkg_winding.columns)
# print(df_jobcard_pkg_winding.head())
print('pending')
df_rptstapprovalpending = pd.read_excel(file_name_rptstapprovalpending)
print(df_rptstapprovalpending.columns)
# print(df_rptstapprovalpending.head())
#### jobcard_pkg_streaching, winding and pending
#### jobcard_pkg_streaching, winding and pending
#### jobcard_pkg_streaching, winding and pending
#### jobcard_pkg_streaching, winding and pending
# print(missing_codes.head())


##########
##########
##########
##########

df_ProductionPlanning['Item Mapped'] = df_ProductionPlanning['Article'].apply(get_item_mapped_for_ProdPlan)
df_ProductionPlanning['Dry weight(gram)'] = df_ProductionPlanning['Article'].apply(get_dry_wt_for_ProdPlan)
#
# df_ProductionPlanning['prod req in KG'] = df_ProductionPlanning['Dry weight(gram)'] * df_ProductionPlanning[
#     'Prod. Req. Cone'] / 1000
df_ProductionPlanning['prod req in KG'] = np.select(
    [
        df_ProductionPlanning['StockOn'] == 'Cones',
        df_ProductionPlanning['StockOn'] == 'Boxes'
    ],
    [
        df_ProductionPlanning['Dry weight(gram)'] *df_ProductionPlanning['Prod. Req. Cone'] / 1000,
        df_ProductionPlanning['Dry weight(gram)'] *df_ProductionPlanning['Prod. Req. Box'] / 1000
    ],
    default=None
)

print(df_ProductionPlanning.columns)
df_ProductionPlanning.to_excel('tempo_output_prodd_plan_extended.xlsx', index=False)
# #####


df_jobcard_pkg_streaching['Item Name according to me'] = df_jobcard_pkg_streaching['Item Name'].apply(
    get_item_for_me_for_streaching_or_winding_or_pending)
df_jobcard_pkg_streaching = df_jobcard_pkg_streaching.groupby(['Item Name according to me', 'Shade']).agg(
    total_Batch_Qnty=('Batch Qnty', 'sum'),
    Batches=('Batch', ', '.join)
).reset_index()
print(df_jobcard_pkg_streaching.columns)
# print(df_jobcard_pkg_streaching.head)
df_jobcard_pkg_streaching.to_excel('tempo_output_streaching.xlsx', index=False)
##########


df_jobcard_pkg_winding['Item Name according to me'] = df_jobcard_pkg_winding['Item Name'].apply(
    get_item_for_me_for_streaching_or_winding_or_pending)
df_jobcard_pkg_winding = df_jobcard_pkg_winding.groupby(['Item Name according to me', 'Shade']).agg(
    total_Batch_Qnty=('Batch Qnty', 'sum'),
    Batches=('Batch', ', '.join)
).reset_index()
print(df_jobcard_pkg_winding.columns)
# print(df_jobcard_pkg_winding.head)
df_jobcard_pkg_winding.to_excel('tempo_output_winding.xlsx', index=False)
##########


df_rptstapprovalpending['Item Name according to me'] = df_rptstapprovalpending['Finish Item'].apply(
    get_item_for_me_for_streaching_or_winding_or_pending)
df_rptstapprovalpending = df_rptstapprovalpending.groupby(['Item Name according to me', 'Shade.']).agg(
    total_Batch_Qnty=('Quantity', 'sum'),
    JobCard_nos=('JobCard No.', ', '.join)
).reset_index()
print(df_rptstapprovalpending.columns)
# print(df_rptstapprovalpending.head)
df_rptstapprovalpending.to_excel('tempo_output_pendding.xlsx', index=False)
##########
##########
##########
##########


len_df_ProductionPlanning = len(df_ProductionPlanning)
for i in range(len_df_ProductionPlanning):

    item_accor_to_me = df_ProductionPlanning.iloc[i]['Item Mapped']
    # print(item_accor_to_me)
    shade = str(df_ProductionPlanning.iloc[i]['Shade'])
    shade=re.sub(r"\s+","", shade)
    shade=shade.upper()


    prod_req_in_KG = df_ProductionPlanning.iloc[i]['prod req in KG']


###################
    df_jobcard_pkg_streaching['Shade'] = df_jobcard_pkg_streaching['Shade'].str.upper()
    match = df_jobcard_pkg_streaching[(df_jobcard_pkg_streaching['Item Name according to me'].astype(str) == str(item_accor_to_me)) & (
                df_jobcard_pkg_streaching['Shade'].astype(str) == str(shade))]
    stock_in_streaching = 0
    if not match.empty:
        stock_in_streaching = match.iloc[0]['total_Batch_Qnty']
        # print(stock_in_streaching)



##############################
    df_jobcard_pkg_winding['Shade'] = df_jobcard_pkg_winding['Shade'].str.upper()
    match = df_jobcard_pkg_winding[(df_jobcard_pkg_winding['Item Name according to me'].astype(str) == str(item_accor_to_me)) & (
                df_jobcard_pkg_winding['Shade'].astype(str) == str(shade))]
    stock_in_winding = 0
    if not match.empty:
        stock_in_winding = match.iloc[0]['total_Batch_Qnty']
        # print(stock_in_winding)



##############################
    df_rptstapprovalpending['Shade.'] = df_rptstapprovalpending['Shade.'].str.upper()
    match = df_rptstapprovalpending[(df_rptstapprovalpending['Item Name according to me'].astype(str) == str(item_accor_to_me)) & (
                df_rptstapprovalpending['Shade.'].astype(str) == str(shade))]
    stock_in_pending_approval = 0
    if not match.empty:
        stock_in_pending_approval = match.iloc[0]['total_Batch_Qnty']
        # print(stock_in_pending_approval)

    ##############
    ##############
    total_in_stock=stock_in_pending_approval+stock_in_winding+stock_in_streaching
    df_ProductionPlanning.loc[i, 'Stock_in_winding_streaching_and_PA'] =total_in_stock
    remaining_to_be_planned_for_dyeing=prod_req_in_KG-total_in_stock

    if remaining_to_be_planned_for_dyeing>0:
        df_ProductionPlanning.loc[i, 'Need_to_plan_for_dyeing_KG'] =remaining_to_be_planned_for_dyeing
    else:
        df_ProductionPlanning.loc[i, 'Need_to_plan_for_dyeing_KG'] = None

df_ProductionPlanning['S.No.'] = df_ProductionPlanning.index
len_nearly_same_group_count=len(nearly_same_group_count)
df_ProductionPlanning['Item Mapped image']=df_ProductionPlanning['Item Mapped']
for i in range(len_nearly_same_group_count):
    group_list=nearly_same_group_count[i]
    item_name_image=group_list[0]
    for j in range(len(group_list)):
        for k in range(len_df_ProductionPlanning):
            if df_ProductionPlanning.loc[k, 'Item Mapped image']==group_list[j]:
                df_ProductionPlanning.loc[k, 'Item Mapped image'] =item_name_image






####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before
# date_str = datetime.now().strftime("%d-%m-%Y")
# file_name_to_save_all_data="temporary_output_Winding_lead_time_of_"+date_str+".xlsx"
file_name_to_save_all_data='OUTPUT_order_plan_for_dyeing.xlsx'
print(file_name_to_save_all_data)
if os.path.isfile(file_name_to_save_all_data):
    os.remove(file_name_to_save_all_data)
# df_result_2.to_excel(file_name_to_save_all_data, index=True)

df_ProductionPlanning.to_excel(file_name_to_save_all_data, index=False)

####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before

df_ProductionPlanning_for_further = df_ProductionPlanning.dropna(subset=['Need_to_plan_for_dyeing_KG'])
df_ProductionPlanning_for_further = df_ProductionPlanning_for_further.groupby(['Item Mapped image', 'Shade']).agg(
    total_Need_to_plan_for_dyeing_KG=('Need_to_plan_for_dyeing_KG', 'sum'),
    serial_nos=('S.No.',lambda x: ', '.join(x.astype(str))),
    count_Need_to_plan_for_dyeing_KG=('Need_to_plan_for_dyeing_KG', 'count')

).reset_index()


df_ProductionPlanning_for_further.to_excel('OUTPUT_ProductionPlanning_for_furt.xlsx', index=False)


