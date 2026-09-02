import pandas as pd
import matplotlib.pyplot as plt
import re
from search_by_raw_material_name_function_Considering_ply import search_by_raw_material_name_function_Considering_ply

file_name_1 = "OrderHistory_all.xlsx"
file_dry_weight = "combined_dry_weight.xlsx"

df_dry_weight = pd.read_excel(file_dry_weight)
len_df_dry_weight=len(df_dry_weight)
ddata=[]
for i in range(len_df_dry_weight):
    item_name = df_dry_weight['Item Mapped'][i]
    input_article_number = df_dry_weight['Article No.'][i]  #####This has to be changed while for CONTINOUS FILAMENT and spun
    a = search_by_raw_material_name_function_Considering_ply(file_name_1, file_dry_weight, item_name, input_article_number)

    ddata.append({'item mapped':a[0],'total_order_for_item_per_month_in_kg':a[1],'total_order_for_item_for_BW_per_month_in_kg':a[2],'total_order_for_item_WITHOUT_BW_per_month_in_kg':a[3]})








df_data=pd.DataFrame(ddata)
df_cleaned_data=df_data.drop_duplicates()
df_cleaned_data.to_excel("temporary_output_Considering_ply.xlsx")
