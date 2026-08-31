import pandas as pd
from order_planning_calculation_upgrade_Upgrade_func import order_planning_calculation_upgrade_Upgrade_func
import matplotlib.pyplot as plt
import re
# from search_by_raw_material_name_function import search_by_raw_material_name_function

import math
import numpy as np
import os

file_name_input_series_article_shade_qty = "input_files/input_series_article_shade_qty.xlsx"
df_input_series_article_shade_qty = pd.read_excel(file_name_input_series_article_shade_qty)
# print(df_input_series_article_shade_qty.columns)
# print(len(df_input_series_article_shade_qty))
df_input_series_article_shade_qty['order_Qty_in_Kg']=None
df_input_series_article_shade_qty['Prod_Req_Qty_in_cone_or_boxes_for_article']=None
df_input_series_article_shade_qty['Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp']=None
df_input_series_article_shade_qty['Prod_Req_weight_accdg_to_item_name']=None
df_input_series_article_shade_qty['dyeing']=None
df_input_series_article_shade_qty['winding']=None
df_input_series_article_shade_qty['stretching']=None
df_input_series_article_shade_qty['pending']=None
df_input_series_article_shade_qty['excess_in_dying_w_s_p']=None
df_input_series_article_shade_qty['Order_need_to_plan_for_dyeing_KG']=None



for i in range (len(df_input_series_article_shade_qty)):
    Article_no_to_plan = df_input_series_article_shade_qty['Article'].iloc[i]
    Shade_to_plan=df_input_series_article_shade_qty['Shade'].iloc[i]
    Qty_to_plan_in_no_for_cones_or_box=df_input_series_article_shade_qty['Qty'].iloc[i]
    Order_already_punched_in_PDMS=df_input_series_article_shade_qty['Already_punched_or_not'].iloc[i]
    result=order_planning_calculation_upgrade_Upgrade_func(Article_no_to_plan,Shade_to_plan,Qty_to_plan_in_no_for_cones_or_box,Order_already_punched_in_PDMS)
    print(result)
    df_input_series_article_shade_qty.loc[i, 'order_Qty_in_Kg']=result['order_Qty_in_Kg']
    df_input_series_article_shade_qty.loc[i, 'Prod_Req_Qty_in_cone_or_boxes_for_article'] = result['Prod_Req_Qty_in_cone_or_boxes_for_article']
    df_input_series_article_shade_qty.loc[i, 'Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp'] = result['Qty_to_plan_in_KG_for_cones_or_box_After_Finish_CheckUp']
    df_input_series_article_shade_qty.loc[i, 'Prod_Req_weight_accdg_to_item_name'] = result['Prod_Req_weight_accdg_to_item_name']
    df_input_series_article_shade_qty.loc[i, 'dyeing'] = result['dyeing']
    df_input_series_article_shade_qty.loc[i, 'winding'] = result['winding']
    df_input_series_article_shade_qty.loc[i, 'stretching'] = result['stretching']
    df_input_series_article_shade_qty.loc[i, 'pending'] = result['pending']
    df_input_series_article_shade_qty.loc[i, 'excess_in_dying_w_s_p'] = result['excess_in_dying_w_s_p']
    df_input_series_article_shade_qty.loc[i, 'Order_need_to_plan_for_dyeing_KG'] = result['Order_need_to_plan_for_dyeing_KG']



df_input_series_article_shade_qty.to_excel('Output_files/OUTPUT_input_series_article_shade_qty_Upgraded.xlsx', index=False)
# print(df_jobcard_pkg_streaching.head())