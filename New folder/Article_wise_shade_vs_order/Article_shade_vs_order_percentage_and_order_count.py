import pandas as pd
import matplotlib.pyplot as plt

file_name_1="OrderHistory_all.xlsx"
# [1201,2303,3302,3301,2311,3310,2315,2316]
article_list=[1201]
number_of_shade_toplot=20
# file_name_1="OrderHistory_all.xlsx"
df = pd.read_excel(file_name_1)
# print(df.head())

print(df.columns)
# df.drop('Party')
df2=df[['Order Date','Article', 'Color']].copy()
df2['Actual order']=df['Order Qty']-df['Cancel Qty']
del df
# # print(df2.head())
# # df3=df2.pivot_table(index=['Article', 'Color'], values=['Actual order','Order Date'], aggfunc=['sum','count'])
# df3=df2.pivot_table(index=['Article', 'Color'], values=['Actual order'], aggfunc=['sum'])
# print(df3)
#
# print(df3.columns)


# df.groupby(['Department', 'Gender']).agg(
#     Count=('Salary', 'count'),
#     Mean=('Salary', 'mean'),
#     Sum=('Salary', 'sum'),
#     Max=('Salary', 'max'),
#     Min=('Salary', 'min')
# ).reset_index()



df3=df2.groupby(['Article', 'Color']).agg(

    Sum_order=('Actual order', 'sum'),
    Count_order=('Order Date', 'count')
).reset_index()
del df2

df4=df3[df3['Article']==article_list[0]]
total_order=df4['Sum_order'].sum()
df4['Order_percentage']=df4['Sum_order']/total_order*100
sorted_df4=df4.sort_values(by='Order_percentage', ascending=False)
del df4
length_of_data=len(sorted_df4)
print(sorted_df4)
size_plot=number_of_shade_toplot
if length_of_data<number_of_shade_toplot:
    size_plot=length_of_data

# plt.plot(sorted_df4['Color'][0:size_plot], sorted_df4['Order_percentage'][0:size_plot], marker='o')
# plt.plot(sorted_df4['Color'][0:size_plot], sorted_df4['Count_order'][0:size_plot], marker='o')
# plt.xlabel('Shade_number')
# plt.ylabel('percentage')
# plt.title('Shade_number vs percentage')
# plt.grid(True)
# plt.show()


fig, ax1 = plt.subplots(figsize=(8,5))

# Left Y-axis (Blue)
ax1.plot(sorted_df4['Color'][0:size_plot], sorted_df4['Order_percentage'][0:size_plot], color='blue', marker='o', linewidth=2, label='Order_percentage')
ax1.set_xlabel('Shade')
ax1.set_ylabel('Order_percentage', color='blue')
ax1.tick_params(axis='y', colors='blue')          # Tick labels
ax1.spines['left'].set_color('blue')              # Left axis line

# Right Y-axis (Red)
ax2 = ax1.twinx()
ax2.plot(sorted_df4['Color'][0:size_plot], sorted_df4['Count_order'][0:size_plot], color='red', marker='s', linestyle='--',
         linewidth=2, label='Count_order')
ax2.set_ylabel('Count_order', color='red')
ax2.tick_params(axis='y', colors='red')           # Tick labels
ax2.spines['right'].set_color('red')              # Right axis line

# Combined legend
lines = ax1.get_lines() + ax2.get_lines()
labels = [line.get_label() for line in lines]
ax1.legend(lines, labels, loc='upper left')
strinn="Shade vs order % and order count of "+str(article_list[0]) + " article"
plt.title(strinn)
plt.grid(True)
plt.show()