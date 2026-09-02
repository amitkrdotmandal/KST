
import pandas as pd
from datetime import timedelta

file_name_DispatchReport = "DispatchReport.xlsx"
file_name_StockReportArticleShadeWise = "StockReportArticleShadeWise.xlsx"


no_of_days_to_check_in_stock = int(input("Enter name no_of_days_to_check_in_stock: "))
# no_of_days_to_check_in_stock = 3
df_DispatchReport = pd.read_excel(file_name_DispatchReport)
print(df_DispatchReport.columns)
df_StockReportArticleShadeWise= pd.read_excel(file_name_StockReportArticleShadeWise)
df_StockReportArticleShadeWise.rename(columns={'Article No.': 'Article'}, inplace=True)
print(df_StockReportArticleShadeWise.columns)

df_DispatchReport['Challan Date'] = pd.to_datetime(df_DispatchReport['Challan Date'])
latest_date = df_DispatchReport['Challan Date'].max()
Beginning_date = latest_date - timedelta(days=no_of_days_to_check_in_stock)
print(latest_date)
print(Beginning_date)

df_DispatchReport_filtered = df_DispatchReport[
    (df_DispatchReport['Challan Date'] >= Beginning_date) &
    (df_DispatchReport['Challan Date'] <= latest_date)
]

df_DispatchReport_filtered.to_excel('temp_out.xlsx', index=False)

merged_df = pd.merge(
    df_StockReportArticleShadeWise,
    df_DispatchReport_filtered,
    on=['Article', 'Shade'],
    how='left'
)
merged_df = merged_df[merged_df['Challan No.'].isna()]

merged_df.to_excel('temp_out2.xlsx', index=False)

Output_merged_df=merged_df[['Article','Shade','Stock Managed on','Closing Boxes','Closing Cones']]
Output_merged_df.to_excel('Output_NonMovingFinish_Check_in_last_few_days.xlsx', index=False)





