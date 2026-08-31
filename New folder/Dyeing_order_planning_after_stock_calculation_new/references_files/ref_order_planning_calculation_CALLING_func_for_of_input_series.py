import pandas as pd
from order_planning_calculation_func import order_planning_calculation_func
import matplotlib.pyplot as plt
import re
# from search_by_raw_material_name_function import search_by_raw_material_name_function

import math
import numpy as np
import os

file_name_input_series_article_shade_qty = "input_series_article_shade_qty.xlsx"
df_input_series_article_shade_qty = pd.read_excel(file_name_input_series_article_shade_qty)
# print(df_input_series_article_shade_qty.columns)
# print(len(df_input_series_article_shade_qty))
df_input_series_article_shade_qty['Dyeing_plan_Qty_in_KG']=None

for i in range (len(df_input_series_article_shade_qty)):
    Article_no_to_plan = df_input_series_article_shade_qty['Article'].iloc[i]
    Shade_to_plan=df_input_series_article_shade_qty['Shade'].iloc[i]
    Qty_to_plan_in_no_for_cones_or_box=df_input_series_article_shade_qty['Qty'].iloc[i]
    Order_already_punched_in_PDMS=df_input_series_article_shade_qty['Already_punched_or_not'].iloc[i]
    result=order_planning_calculation_func(Article_no_to_plan,Shade_to_plan,Qty_to_plan_in_no_for_cones_or_box,Order_already_punched_in_PDMS)
    df_input_series_article_shade_qty.loc[i, 'Dyeing_plan_Qty_in_KG']=result
    print(result)


df_input_series_article_shade_qty.to_excel('OUTPUT_input_series_article_shade_qty.xlsx', index=False)
# print(df_jobcard_pkg_streaching.head())