import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime
import re


file_name_party_list="Input_files/Input_party_list.xlsx"
df_party_list = pd.read_excel(file_name_party_list)
print(df_party_list.columns)
parties_to_be_analysed=df_party_list['Party'].tolist()
# parties_to_be_analysed=['jak','T R THREADS','hillson','SI INTERPACK PRIVATE LIMITED','K.K. HOSIERY MATERIAL STORE', 'CHIRAG CREATIONS','krishna'] #if list is blank, it takes all parties
# parties_to_be_analysed=[] #if list is blank, it takes all parties








file_name_PendingOrderReport="PendingOrderReport.xlsx"
df_PendingOrderReport = pd.read_excel(file_name_PendingOrderReport)
print(df_PendingOrderReport.columns)
print(df_PendingOrderReport.head())

df_PendingOrderReport['Order Date'] = pd.to_datetime(df_PendingOrderReport['Order Date'], format='%d/%m/%Y')
df_PendingOrderReport['Days'] = pd.Timestamp.today().normalize() - df_PendingOrderReport['Order Date']
# df_PendingOrderReport.to_excel('tempo_out.xlsx', index=False)


#============================================
#party filter starts
#============================================

parties_to_be_analysed = [party.lower() for party in parties_to_be_analysed]
len_parties_to_be_analysed= len(parties_to_be_analysed)
if len_parties_to_be_analysed>1:
    df_PendingOrderReport = df_PendingOrderReport[
        df_PendingOrderReport['Party'].str.lower().str.contains(
            '|'.join(parties_to_be_analysed),
            na=False
        )
    ]

    ##### Rename

    df_PendingOrderReport['Party'] = (  ##### May be deleted if PARTY name need not to be renamed
        df_PendingOrderReport['Party']
        .apply(
            lambda x: next(
                (party.upper() for party in parties_to_be_analysed
                 if party.lower() in str(x).lower()),
                x
            )
        )
    )
# df_PendingOrderReport.to_excel('tempo_out2.xlsx', index=False)
#============================================
#party filter ends
#============================================




party_list = df_PendingOrderReport['Party'].astype(str).unique().tolist()
print(party_list)
with pd.ExcelWriter('Output_files/Output_Party_vs_Article_Shade_Pending_odr_days.xlsx') as Writter:
    for party in party_list:
        filtered_df = df_PendingOrderReport[
            df_PendingOrderReport['Party']==party]
        filtered_df = filtered_df.sort_values(by='Days', ascending=False)
        filtered_df = filtered_df.drop(columns=['Dispatch Qty', 'Cancel Qty', 'Rate'])
        sheet_name = party
        filtered_df.to_excel(Writter, sheet_name=sheet_name, index=False)