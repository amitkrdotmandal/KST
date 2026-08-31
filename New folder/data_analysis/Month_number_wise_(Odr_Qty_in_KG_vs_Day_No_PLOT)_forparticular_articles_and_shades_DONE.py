import pandas as pd
import matplotlib.pyplot as plt

month_no=7
article_list=[1201,1206] ### when list is blank it takes whole articles
shade_no_list=[] ### when list is blank it takes whole shades
seve_figure=1# 1 for yes 0 for no



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





df_OrderHistory = pd.read_excel(file_name_OrderHistory)
print(df_OrderHistory.columns)
# print(df_OrderHistory.head())
df_combined_dry_weight_DB = pd.read_excel(file_name_combined_dry_weight_DB)
print(df_combined_dry_weight_DB.columns)
# print(df_combined_dry_weight_DB.head())











article_list = [str(x) for x in article_list]
df_OrderHistory['Article'] = df_OrderHistory['Article'].astype(str)
df_OrderHistory['Actual order']=df_OrderHistory['Order Qty']-df_OrderHistory['Cancel Qty']
df_OrderHistory['Order Date'] = pd.to_datetime(df_OrderHistory['Order Date'], format='%d/%m/%Y')
df_OrderHistory['Item Mapped'] = df_OrderHistory['Article'].apply(get_item_mapped_for_ProdPlan)
df_OrderHistory['Dry weight(gram)'] = df_OrderHistory['Article'].apply(get_dry_wt_for_ProdPlan)
df_OrderHistory['Total order in KG']=df_OrderHistory['Actual order']*df_OrderHistory['Dry weight(gram)']/1000
# print(df_OrderHistory['Order Date'].dtype)
print(df_OrderHistory.columns)
print(df_OrderHistory.head())



# given_date = pd.to_datetime('30/04/2026', format='%d/%m/%Y')
#####date filter
df_OrderHistory_t=df_OrderHistory[df_OrderHistory['Order Date'].dt.month == month_no]
#####Article filter
if len(article_list)>0:
    df_OrderHistory_t=df_OrderHistory_t[df_OrderHistory_t['Article'].isin(article_list)]
#####Shade filter
if len(shade_no_list)>0:
    df_OrderHistory_t=df_OrderHistory_t[df_OrderHistory_t['Color'].isin(shade_no_list)]



dday=[]
sum_of_actual_order_for=[]
for i in range(30):
    dday.append(i)
    sum_of_a_day=df_OrderHistory_t[df_OrderHistory_t['Order Date'].dt.day == i][['Total order in KG']].sum()
    sum_of_actual_order_for.append(sum_of_a_day)

plt.plot(dday, sum_of_actual_order_for,color='red', marker='s', linestyle='--',
         linewidth=2,)


plt.xlabel('day of month')
plt.ylabel('Total order in KG')
strinn="total order for month "+str(month_no)
# if len(article_list)>0:
#     strinn=strinn+ " of  Articles: "+str((article_list))
# if len(shade_no_list)>0:
#     strinn=strinn+ " of  shades: "+str((shade_no_list))

plt.title(strinn)
plt.grid(True)
if seve_figure==1:
    to_save_figure_name = 'Save figure/' + strinn + '.png'
    plt.savefig(to_save_figure_name, dpi=300, bbox_inches='tight')
plt.show()

df_OrderHistory_t.to_excel('temp_out.xlsx', index=False)
