import pandas as pd
import numpy as np
import math
import matplotlib.pyplot as plt
import re
import os
from datetime import datetime

file_name_ProductionPlanning = "ProductionPlanning.xlsx"
file_name_capacity="4 August -Expanded capacity calculation_DB.xlsx"
SSP_TXP_PPC_sheet_name="New-Cycle time SSP-TXP-PPC"
filament_sheet_name="New Cycle time-Filament"
tube_sheet_name="New-Cycle time TUBE"





df_ProductionPlanning = pd.read_excel(file_name_ProductionPlanning)
print(df_ProductionPlanning.columns)
df_ProductionPlanning = df_ProductionPlanning[df_ProductionPlanning['Prod. Req. Box'] >= 0]
df_ProductionPlanning = df_ProductionPlanning[df_ProductionPlanning['Prod. Req. Cone'] >= 0]
df2_df_ProductionPlanning = df_ProductionPlanning.groupby(['Article','StockOn']).agg(
    Prod_Req_box_total=('Prod. Req. Box', 'sum'),
    Prod_Req_Cone_total=('Prod. Req. Cone', 'sum'),
).reset_index()
del df_ProductionPlanning
df2_df_ProductionPlanning['Article'] = pd.to_numeric(df2_df_ProductionPlanning['Article'], errors='coerce')
# df2_df_ProductionPlanning['is_cone'] = np.where(df2_df_ProductionPlanning['StockOn'] == 'Cones', 1, 0)
# df2_df_ProductionPlanning['is_box'] = np.where(df2_df_ProductionPlanning['StockOn'] == 'Boxes', 1, 0)


print(df2_df_ProductionPlanning.columns)
print(df2_df_ProductionPlanning)




df_SSP_TXP_PPC = pd.read_excel(file_name_capacity, sheet_name=SSP_TXP_PPC_sheet_name)
df_SSP_TXP_PPC=df_SSP_TXP_PPC.iloc[:, [1,5,9,19]]
df_filament = pd.read_excel(file_name_capacity, sheet_name=filament_sheet_name)
df_filament=df_filament.iloc[:, [1,5,9,19]]
df_tube = pd.read_excel(file_name_capacity, sheet_name=tube_sheet_name)
df_tube=df_tube.iloc[:, [1,5,9,19]]
# print(df_SSP_TXP_PPC.columns)
# print(df_SSP_TXP_PPC.head())
# print(df_filament.columns)
# print(df_filament.head())
# print(df_tube.columns)
# print(df_tube.head())


df_new_all_type_RM = pd.concat([df_SSP_TXP_PPC, df_filament, df_tube], ignore_index=True)
# print(df_new_all_type_RM.columns)
# print(df_new_all_type_RM)
# print(len(df_new_all_type_RM))
del df_SSP_TXP_PPC, df_filament, df_tube
df_new_all_type_RM = df_new_all_type_RM.rename(columns={'FG\ncode': 'Article'})
# print(df_new_all_type_RM.columns)
# print(df2_df_ProductionPlanning.columns)
df_result = df_new_all_type_RM.merge(
    df2_df_ProductionPlanning,
    on="Article",
    how="left"
)
print(df_result.columns)
print(df_result.head())
df_result = df_result.dropna(subset=['StockOn'])
df_result['No. of days required'] = np.where(
    df_result['StockOn'] == 'Cones',
    df_result['Prod_Req_Cone_total'] / df_result['Total capacity-actual'],
    df_result['Prod_Req_box_total'] / df_result['Total capacity-actual boxes']
)

df_result.loc[df_result['StockOn'] == 'Boxes', 'Make up'] = (
    df_result.loc[df_result['StockOn'] == 'Boxes', 'Make up']
      .str.replace(r'\s*\(.*?\)', '', regex=True)
)

print(df_result.columns)
# df_result.to_excel("temporary_output_new.xlsx", index=True)

df_result_2=df_result.groupby(['Make up']).agg(
    Lead_time_in_days=('No. of days required', 'sum'),
    No_of_cones_required_to_produce=('Prod_Req_Cone_total', 'sum'),
    No_of_boxes_required_to_produce=('Prod_Req_box_total', 'sum')

).reset_index()
df_result_2['Lead_time_in_days'] = df_result_2['Lead_time_in_days'].round(2)
# print(df_result_2)



####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before
date_str = datetime.now().strftime("%d-%m-%Y")
# file_name_to_save_all_data="temporary_output_Winding_lead_time_of_"+date_str+".xlsx"
file_name_to_save_all_data="Output_Winding_lead_time_of"+".xlsx"
print(file_name_to_save_all_data)
if os.path.isfile(file_name_to_save_all_data):
    os.remove(file_name_to_save_all_data)
# df_result_2.to_excel(file_name_to_save_all_data, index=True)

with pd.ExcelWriter(file_name_to_save_all_data, engine="openpyxl") as writer:
    df_result_2.to_excel(writer, sheet_name="Sheet1", startrow=2, index=False)
    worksheet = writer.sheets["Sheet1"]
    worksheet["A1"] = "Winding_lead_time_of_"+date_str

####File_to_save_delete_before
####File_to_save_delete_before
####File_to_save_delete_before





