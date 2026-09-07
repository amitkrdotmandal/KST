import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime
import re


start_day='01/04/2026'
parties_to_be_analysed=['jak','T R THREADS','hillson','SI INTERPACK PRIVATE LIMITED','K.K. HOSIERY MATERIAL STORE', 'CHIRAG CREATIONS','FANTAIL'] #if list is blank, it takes all parties
# parties_to_be_analysed=[] #if list is blank, it takes all parties








file_name_OrderHistory="OrderHistory.xlsx"
df_OrderHistory = pd.read_excel(file_name_OrderHistory)
print(df_OrderHistory.columns)
file_name_combined_dry_weight_DB = "combined_dry_weight_DB.xlsx"
# print(df_OrderHistory.head())
df_combined_dry_weight_DB = pd.read_excel(file_name_combined_dry_weight_DB)
print(df_combined_dry_weight_DB.columns)
# print(df_combined_dry_weight_DB.head())






#============================================
#date filter starts
#============================================
df_combined_dry_weight_DB.rename(columns={'Article No.': 'Article'}, inplace=True)
df_OrderHistory['Actual order']=df_OrderHistory['Order Qty']-df_OrderHistory['Cancel Qty']
df_OrderHistory['Order Date'] = pd.to_datetime(df_OrderHistory['Order Date'], format='%d/%m/%Y')
df_OrderHistory = pd.merge(df_OrderHistory, df_combined_dry_weight_DB, on='Article', how='left')
df_OrderHistory['actual_tot_wt_in_KG']=df_OrderHistory['Actual order']*df_OrderHistory['Dry weight(gram)']/1000
start_day = datetime.strptime(start_day, '%d/%m/%Y')
max_date = df_OrderHistory['Order Date'].max()
df_OrderHistory_date_filtered = df_OrderHistory[
    (df_OrderHistory['Order Date'] >= start_day) &
    (df_OrderHistory['Order Date'] <= max_date)
]
# df_OrderHistory_date_filtered.to_excel('tempo_out1.xlsx', index=False)
#============================================
#date filter ends
#============================================


#============================================
#party filter starts
#============================================
df_OrderHistory_date_and_party_filtered=df_OrderHistory_date_filtered
parties_to_be_analysed = [party.lower() for party in parties_to_be_analysed]
len_parties_to_be_analysed= len(parties_to_be_analysed)
if len_parties_to_be_analysed>1:
    df_OrderHistory_date_and_party_filtered = df_OrderHistory_date_and_party_filtered[
        df_OrderHistory_date_and_party_filtered['Party'].str.lower().str.contains(
            '|'.join(parties_to_be_analysed),
            na=False
        )
    ]

    ##### Rename

    df_OrderHistory_date_and_party_filtered['Party'] = (          ##### May be deleted if PARTY name need not to be renamed
        df_OrderHistory_date_and_party_filtered['Party']
        .apply(
            lambda x: next(
                (party.upper() for party in parties_to_be_analysed
                 if party.lower() in str(x).lower()),
                x
            )
        )
    )


# df_OrderHistory_date_and_party_filtered.to_excel('tempo_out2.xlsx', index=False)
#============================================
#party filter ends
#============================================









#============================================
#Article_shade vs Party starts
#============================================
df_OrderHistory_grouped1 = df_OrderHistory_date_and_party_filtered.groupby(
    ['Article', 'Color','Party'],
    as_index=False
).agg(
    Last_date_of_Odr=('Order Date', 'max'),
    actual_tot_in_Qty=('Actual order', 'sum'),
    actual_tot_wt_in_KG=('actual_tot_wt_in_KG', 'sum'),
    Actual_order_Qty=('Actual order', 'sum'),
    Frequency=('actual_tot_wt_in_KG', 'count')
)

df_OrderHistory_grouped2 = df_OrderHistory_grouped1.groupby(
    ['Article', 'Color'],
    as_index=False
).agg(
    Last_date_of_Odr=('Last_date_of_Odr', 'max'),
    actual_tot_in_Qty=('actual_tot_in_Qty', 'sum'),
    actual_tot_wt_in_KG=('actual_tot_wt_in_KG', 'sum'),
    Total_Frequency=('Frequency', 'sum'),
    No_of_parties=('Party', 'count'),
    Parties=('Party', lambda x: ', '.join(x.astype(str))),
)
df_OrderHistory_grouped2.to_excel('Output_files/Output_all_Art_Shade_vs_Party.xlsx', index=False)


first_digits = df_OrderHistory_grouped2['Article'].astype(str).str[0].astype(int).unique().tolist()
with pd.ExcelWriter('Output_files/Output_Art_wise_sheet_shade_vs_party.xlsx') as Writter:
    for first_digit in first_digits:
        filtered_df = df_OrderHistory_grouped2[
            df_OrderHistory_grouped2['Article'].astype(str).str.startswith(str(first_digit))]
        sheet_name = 'Art_starts_with_' + str(first_digit)
        filtered_df.to_excel(Writter, sheet_name=sheet_name, index=False)
#============================================
#Article_shade vs Party ends
#============================================




#============================================
# Party vs Article and Shade Starts
#============================================

df_OrderHistory_grouped_party = df_OrderHistory_date_and_party_filtered.groupby(
    ['Party','Article','Color'],
    as_index=False
).agg(
    Last_date_of_Odr=('Order Date', 'max'),
    actual_tot_in_Qty=('Actual order', 'sum'),
    actual_tot_wt_in_KG=('actual_tot_wt_in_KG', 'sum'),
    Actual_order_Qty=('Actual order', 'sum'),
    Frequency=('actual_tot_wt_in_KG', 'count')
)
df_OrderHistory_grouped_party.to_excel('tempo1.xlsx', index=False)


party_list = df_OrderHistory_grouped_party['Party'].astype(str).unique().tolist()
print(party_list)
with pd.ExcelWriter('Output_files/Output_Party_vs_Article_Shade.xlsx') as Writter:
    for party in party_list:
        filtered_df = df_OrderHistory_grouped_party[
            df_OrderHistory_grouped_party['Party']==party]
        sheet_name = party
        filtered_df.to_excel(Writter, sheet_name=sheet_name, index=False)

#============================================
# Party vs Article and Shade Ends
#==========================================










