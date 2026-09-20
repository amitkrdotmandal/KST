import pandas as pd
import matplotlib.pyplot as plt
import re
# from search_by_raw_material_name_function import search_by_raw_material_name_function
import math
import numpy as np
import os

Article_no_to_plan = 3302
Shade_to_plan = 'N00904'
Qty_to_plan_in_no_for_cones_or_box = 100
Order_already_punched_in_PDMS=0 ### If already punched '1', if not already punched '0'



file_name_jobcard_pkg_streaching = "input_files/jobcard_pkg_stretching.xlsx"
file_name_jobcard_pkg_winding = "input_files/jobcard_pkg_winding.xlsx"
file_name_rptstapprovalpending = "input_files/rptstapprovalpending.xlsx"
file_name_daily_dyeing_order_to_despatch_report = "input_files/ordertodspsummary.xlsx"
file_name_ProductionPlanning = "input_files/ProductionPlanning.xlsx"

file_name_DB_of_converted_item_name = "To_convert_after_dyeing_ITEM_NAME_to_match_our_ITEM_NAME/OUTPUT_converted_item_name_of_dyed_soft_pkg_DB.xlsx"  ###should not be changed
file_name_combined_dry_weight_DB = "combined_dry_weight_DB.xlsx"





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


#### jobcard_pkg_streaching, winding, pending and dyeing
#### jobcard_pkg_streaching, winding, pending and dyeing
#### jobcard_pkg_streaching, winding, pending and dyeing
#### jobcard_pkg_streaching, winding, pending and dyeing
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
print('Dyeing')
df_daily_dyeing_order_to_despatch_report = pd.read_excel(file_name_daily_dyeing_order_to_despatch_report)
print(df_daily_dyeing_order_to_despatch_report.columns)
# print(df_daily_dyeing_order_to_despatch_report.head())
#### jobcard_pkg_streaching, winding, pending and dyeing
#### jobcard_pkg_streaching, winding, pending and dyeing
#### jobcard_pkg_streaching, winding, pending and dyeing
#### jobcard_pkg_streaching, winding, pending and dyeing
# print(missing_codes.head())


##################INITIAL DATA PROCESSING
##################INITIAL DATA PROCESSING
##########
##########
##########
##########

##for order plan
Shade_to_plan = Shade_to_plan.upper()
match = df_combined_dry_weight_DB[df_combined_dry_weight_DB['Article No.'].astype(str) == str(Article_no_to_plan)]
if not match.empty:
    Item_Name_according_to_me_for_order_plan = (match.iloc[0]['Item Mapped'])
    Dry_weight_of_cone_or_box=(match.iloc[0]['Dry weight(gram)'])
    total_weight_for_planned_odr_in_kg = Qty_to_plan_in_no_for_cones_or_box * (match.iloc[0]['Dry weight(gram)']) / 1000

else:
    Item_Name_according_to_me_for_order_plan = None
    total_weight_for_planned_odr_in_kg = None
##for order plan
# #####


############ for ProductionPlanning
df_ProductionPlanning['Item Name according to me'] = df_ProductionPlanning['Article'].apply(
    get_item_mapped_for_ProdPlan)
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
        df_ProductionPlanning['Dry weight(gram)'] * df_ProductionPlanning['Prod. Req. Cone'] / 1000,
        df_ProductionPlanning['Dry weight(gram)'] * df_ProductionPlanning['Prod. Req. Box'] / 1000
    ],
    default=None
)
df_ProductionPlanning['Shade'] = df_ProductionPlanning['Shade'].str.upper()
df_ProductionPlanning['Shade'] = df_ProductionPlanning['Shade'].str.replace(r"\s+", "", regex=True)

####
match = df_ProductionPlanning[
        (df_ProductionPlanning['Article'].astype(str) == str(Article_no_to_plan)) & (
                df_ProductionPlanning['Shade'].astype(str) == str(Shade_to_plan))]
no_of_cone_or_boxes_in_finish = 0     #### If no_of_cone_or_boxes_in_finish<0, then will be stock or production is required
if not match.empty:
    no_of_cone_or_boxes_in_finish=np.select(
    [
        match.iloc[0]['StockOn'] == 'Cones',
        match.iloc[0]['StockOn'] == 'Boxes'
    ],
    [
        match.iloc[0]['Prod. Req. Cone'],
        match.iloc[0]['Prod. Req. Box']
    ],
    default=None
)



if Order_already_punched_in_PDMS==1:
    Adjusted_no_of_cone_or_boxes_in_finish = no_of_cone_or_boxes_in_finish-Qty_to_plan_in_no_for_cones_or_box
else:
    Adjusted_no_of_cone_or_boxes_in_finish = no_of_cone_or_boxes_in_finish

if Adjusted_no_of_cone_or_boxes_in_finish<0:
    Qty_to_plan_in_no_for_cones_or_box_After_Finish_CheckUp=Qty_to_plan_in_no_for_cones_or_box-(-1)*Adjusted_no_of_cone_or_boxes_in_finish
else:
    Qty_to_plan_in_no_for_cones_or_box_After_Finish_CheckUp = Qty_to_plan_in_no_for_cones_or_box

Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp=Qty_to_plan_in_no_for_cones_or_box_After_Finish_CheckUp*Dry_weight_of_cone_or_box/1000
####




df_ProductionPlanning_aggregated = df_ProductionPlanning.groupby(
    ['Item Name according to me', 'Shade']
).agg(
    total_prod_req_in_KG=('prod req in KG', 'sum'),
    Article_nos=('Article', lambda x: ', '.join(x.astype(str)))
).reset_index()

print(df_ProductionPlanning.columns)
df_ProductionPlanning.to_excel('tempo_output_prodd_plan_extended.xlsx', index=False)
df_ProductionPlanning_aggregated.to_excel('tempo_output_prodd_plan_aggregated.xlsx', index=False)
# #####
############ for ProductionPlanning ENDS


############ for streaching
df_jobcard_pkg_streaching['Shade'] = df_jobcard_pkg_streaching['Shade'].str.upper()
df_jobcard_pkg_streaching['Item Name according to me'] = df_jobcard_pkg_streaching['Item Name'].apply(
    get_item_for_me_for_streaching_or_winding_or_pending)
df_jobcard_pkg_streaching.loc[df_jobcard_pkg_streaching['Balance'] < 0, 'Balance'] = 0
df_jobcard_pkg_streaching = df_jobcard_pkg_streaching.groupby(['Item Name according to me', 'Shade']).agg(
    total_Batch_Qnty=('Balance', 'sum'),
    Batches=('Batch', ', '.join)
).reset_index()


print(df_jobcard_pkg_streaching.columns)
# print(df_jobcard_pkg_streaching.head)
df_jobcard_pkg_streaching.to_excel('tempo_output_streaching.xlsx', index=False)
##########
############ for streaching ENDS


############ for winding
df_jobcard_pkg_winding['Shade'] = df_jobcard_pkg_winding['Shade'].str.upper()
df_jobcard_pkg_winding['Item Name according to me'] = df_jobcard_pkg_winding['Item Name'].apply(
    get_item_for_me_for_streaching_or_winding_or_pending)
df_jobcard_pkg_winding.loc[df_jobcard_pkg_winding['Balance'] < 0, 'Balance'] = 0
df_jobcard_pkg_winding = df_jobcard_pkg_winding.groupby(['Item Name according to me', 'Shade']).agg(
    total_Batch_Qnty=('Balance', 'sum'),
    Batches=('Batch', ', '.join)
).reset_index()

print(df_jobcard_pkg_winding.columns)
# print(df_jobcard_pkg_winding.head)
df_jobcard_pkg_winding.to_excel('tempo_output_winding.xlsx', index=False)
##########
############ for winding ENDS



############ for pending
df_rptstapprovalpending['Shade.'] = df_rptstapprovalpending['Shade.'].str.upper()
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
############ for pending ENDS



############ for dyeing
##FILTER 1 Dont allow "Status" as "Close"
df_daily_dyeing_order_to_despatch_report = df_daily_dyeing_order_to_despatch_report[
    df_daily_dyeing_order_to_despatch_report['Status'] != 'Close']
##FILTER 2 allow Despatch date NaN
df_daily_dyeing_order_to_despatch_report = df_daily_dyeing_order_to_despatch_report[
    df_daily_dyeing_order_to_despatch_report['Dsp. Dt'].isna()]
df_daily_dyeing_order_to_despatch_report['Item Name according to me'] = df_daily_dyeing_order_to_despatch_report[
    'Item Description'].apply(
    get_item_for_me_for_streaching_or_winding_or_pending)
df_daily_dyeing_order_to_despatch_report = df_daily_dyeing_order_to_despatch_report.groupby(
    ['Item Name according to me', 'Shade']).agg(
    total_order_Qnty=('D.O. Qty', 'sum'),
    order_nos=('Order No.', ', '.join)
).reset_index()
print(df_daily_dyeing_order_to_despatch_report.columns)
# print(df_rptstapprovalpending.head)
df_daily_dyeing_order_to_despatch_report.to_excel('tempo_output_dyeing.xlsx', index=False)
############ for dyeing ENDS
##########
##########
##########
##########
##################INITIAL DATA PROCESSING ENDS
##################INITIAL DATA PROCESSING ENDS





################# COMPARISON AND ALL CALCULATIONS TOGETHER
################# COMPARISON AND ALL CALCULATIONS TOGETHER
################# COMPARISON AND ALL CALCULATIONS TOGETHER
################# COMPARISON AND ALL CALCULATIONS TOGETHER
Shade_to_plan
Item_Name_according_to_me_for_order_plan
total_weight_for_planned_odr_in_kg
Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp
total_weight_for_planned_odr_in_kg
Qty_to_plan_in_no_for_cones_or_box_After_Finish_CheckUp

# finish_stock_weight_accdg_to_item_name



################### for productionplanning
df_ProductionPlanning_aggregated['Shade'] = df_ProductionPlanning_aggregated['Shade'].str.upper()
finish_stock_weight_accdg_to_item_name=0
match = df_ProductionPlanning_aggregated[
    (df_ProductionPlanning_aggregated['Item Name according to me'].astype(str) == str(Item_Name_according_to_me_for_order_plan)) & (
            df_ProductionPlanning_aggregated['Shade'].astype(str) == str(Shade_to_plan))]
if Order_already_punched_in_PDMS==1:
    finish_stock_weight_accdg_to_item_name = -total_weight_for_planned_odr_in_kg
else:
    finish_stock_weight_accdg_to_item_name = 0
if not match.empty:
    finish_stock_weight_accdg_to_item_name =finish_stock_weight_accdg_to_item_name+ match.iloc[0]['total_prod_req_in_KG']


    # print(stock_in_streaching)





################### for streaching
df_jobcard_pkg_streaching['Shade'] = df_jobcard_pkg_streaching['Shade'].str.upper()
match = df_jobcard_pkg_streaching[
    (df_jobcard_pkg_streaching['Item Name according to me'].astype(str) == str(Item_Name_according_to_me_for_order_plan)) & (
            df_jobcard_pkg_streaching['Shade'].astype(str) == str(Shade_to_plan))]
running_in_streaching = 0
if not match.empty:
    running_in_streaching = running_in_streaching+match.iloc[0]['total_Batch_Qnty']
stock_in_streaching=running_in_streaching-0
    # print(stock_in_streaching)



############################## for winding
df_jobcard_pkg_winding['Shade'] = df_jobcard_pkg_winding['Shade'].str.upper()
match = df_jobcard_pkg_winding[
        (df_jobcard_pkg_winding['Item Name according to me'].astype(str) == str(Item_Name_according_to_me_for_order_plan)) & (
                df_jobcard_pkg_winding['Shade'].astype(str) == str(Shade_to_plan))]
running_in_winding = 0
if not match.empty:
    running_in_winding =running_in_winding+ match.iloc[0]['total_Batch_Qnty']
stock_in_winding=running_in_winding-0
    # print(stock_in_winding)


############################## for pending
df_rptstapprovalpending['Shade.'] = df_rptstapprovalpending['Shade.'].str.upper()
match = df_rptstapprovalpending[
        (df_rptstapprovalpending['Item Name according to me'].astype(str) == str(Item_Name_according_to_me_for_order_plan)) & (
                df_rptstapprovalpending['Shade.'].astype(str) == str(Shade_to_plan))]
running_in_pending_approval = 0
if not match.empty:
    running_in_pending_approval = running_in_pending_approval+match.iloc[0]['total_Batch_Qnty']
stock_in_pending_approval=running_in_pending_approval-0
    # print(stock_in_pending_approval)

############################## for dyeing
df_daily_dyeing_order_to_despatch_report['Shade'] = df_daily_dyeing_order_to_despatch_report['Shade'].str.upper()
match = df_daily_dyeing_order_to_despatch_report[
        (df_daily_dyeing_order_to_despatch_report['Item Name according to me'].astype(str) == str(Item_Name_according_to_me_for_order_plan)) & (
                df_daily_dyeing_order_to_despatch_report['Shade'].astype(str) == str(Shade_to_plan))]
running_in_dyeing = 0
if not match.empty:
    running_in_dyeing = running_in_dyeing+match.iloc[0]['total_order_Qnty']
stock_in_dyeing=running_in_dyeing-0


######CLUBBED
excess_in_dying_w_s_p=stock_in_streaching+ stock_in_winding+ stock_in_pending_approval + stock_in_dyeing - finish_stock_weight_accdg_to_item_name
if excess_in_dying_w_s_p > 0:
    Order_need_to_plan_for_dyeing_KG=Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp-excess_in_dying_w_s_p
else:
    Order_need_to_plan_for_dyeing_KG = Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp

######CLUBBED

################# COMPARISON AND ALL CALCULATIONS TOGETHER
################# COMPARISON AND ALL CALCULATIONS TOGETHER
################# COMPARISON AND ALL CALCULATIONS TOGETHER
################# COMPARISON AND ALL CALCULATIONS TOGETHER



print('order Qty in Kg:',total_weight_for_planned_odr_in_kg)
print('Prod_Req_Qty in cone or boxes for article:',no_of_cone_or_boxes_in_finish)
print('Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp', Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp)
print('Prod_Req_weight_accdg_to_item_name',finish_stock_weight_accdg_to_item_name)
print('dyeing:',stock_in_dyeing)
print('winding:',stock_in_winding)
print('stretching:',stock_in_streaching)
print('pending:',stock_in_pending_approval)
print('excess_in_dying_w_s_p:',excess_in_dying_w_s_p)
print('order need to plan for dyeing :',Order_need_to_plan_for_dyeing_KG)
print('For Article ',Article_no_to_plan,' & shade ',Shade_to_plan,' with order qTY of ',Qty_to_plan_in_no_for_cones_or_box,', the Qty to plan for dyeing in KG is :',Order_need_to_plan_for_dyeing_KG)
dict_to_return={

    'order_Qty_in_Kg':total_weight_for_planned_odr_in_kg,
    'Prod_Req_Qty_in_cone_or_boxes_for_article':no_of_cone_or_boxes_in_finish,
    'Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp':Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp,
    'Prod_Req_weight_accdg_to_item_name':finish_stock_weight_accdg_to_item_name,
    'dyeing':stock_in_dyeing,
    'winding':stock_in_winding,
    'stretching':stock_in_streaching,
    'pending':stock_in_pending_approval,
    'excess_in_dying_w_s_p':excess_in_dying_w_s_p,
    'Order_need_to_plan_for_dyeing_KG':Order_need_to_plan_for_dyeing_KG


}
print(dict_to_return)
