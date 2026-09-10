import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime
import re



def customer_analysis_from_order_history_func(start_day,parties_to_be_analysed):
    # start_day = '01/04/2026'

    # This party list is being for shortening the party name and make the group common under one name
    # parties_to_be_analysed = ['zigma', 'safira', 'jak', 'T R THREADS', 'hillson', 'SI INTERPACK',
    #                           'K.K. HOSIERY MATERIAL STORE', 'CHIRAG CREATIONS', 'FANTAIL']
    # This party list is being for shortening the party name and make the group common under one name

    file_name_OrderHistory = "OrderHistory.xlsx"
    df_OrderHistory = pd.read_excel(file_name_OrderHistory)
    print(df_OrderHistory.columns)
    file_name_combined_dry_weight_DB = "combined_dry_weight_DB.xlsx"
    # print(df_OrderHistory.head())
    df_combined_dry_weight_DB = pd.read_excel(file_name_combined_dry_weight_DB)
    print(df_combined_dry_weight_DB.columns)
    # print(df_combined_dry_weight_DB.head())

    # ============================================
    # date filter starts
    # ============================================
    df_combined_dry_weight_DB.rename(columns={'Article No.': 'Article'}, inplace=True)
    df_OrderHistory['Actual order'] = df_OrderHistory['Order Qty'] - df_OrderHistory['Cancel Qty']
    df_OrderHistory['Order Date'] = pd.to_datetime(df_OrderHistory['Order Date'], format='%d/%m/%Y')
    df_OrderHistory = pd.merge(df_OrderHistory, df_combined_dry_weight_DB, on='Article', how='left')
    df_OrderHistory['actual_tot_wt_in_KG'] = df_OrderHistory['Actual order'] * df_OrderHistory[
        'Dry weight(gram)'] / 1000
    start_day = datetime.strptime(start_day, '%d/%m/%Y')
    max_date = df_OrderHistory['Order Date'].max()
    df_OrderHistory_date_filtered = df_OrderHistory[
        (df_OrderHistory['Order Date'] >= start_day) &
        (df_OrderHistory['Order Date'] <= max_date)
        ]
    # df_OrderHistory_date_filtered.to_excel('tempo_out1.xlsx', index=False)
    # ============================================
    # date filter ends
    # ============================================

    # ============================================
    # party filter starts
    # ============================================
    df_OrderHistory_date_and_party_filtered = df_OrderHistory_date_filtered
    parties_to_be_analysed = [party.lower() for party in parties_to_be_analysed]
    len_parties_to_be_analysed = len(parties_to_be_analysed)

    #     ##### Rename
    df_OrderHistory_date_and_party_filtered['Party'] = (
    ##### May be deleted if PARTY name need not to be renamed or grouping is not required
        df_OrderHistory_date_and_party_filtered['Party']
        .apply(
            lambda x: next(
                (party.upper() for party in parties_to_be_analysed
                 if party.lower() in str(x).lower()),
                x
            )
        )
    )

    # if len_parties_to_be_analysed>1:
    #     df_OrderHistory_date_and_party_filtered = df_OrderHistory_date_and_party_filtered[
    #         df_OrderHistory_date_and_party_filtered['Party'].str.lower().str.contains(
    #             '|'.join(parties_to_be_analysed),
    #             na=False
    #         )
    #     ]
    #

    # df_OrderHistory_date_and_party_filtered.to_excel('tempo_out2.xlsx', index=False)
    # ============================================
    # party filter ends
    # ============================================

    # ============================================
    # Renaming black to KS- Starts
    # ============================================
    # df.loc[df['Party'] == 'ABC', 'Party'] = 'XYZ'

    # ============================================
    # Renaming black to OTHER BLACK STARTS
    # ============================================
    df_OrderHistory_date_and_party_filtered.loc[
        (
            df_OrderHistory_date_and_party_filtered['Article']
            .astype(str)
            .str.startswith('1')
        )
        & (
                df_OrderHistory_date_and_party_filtered['Color'] == 'BLACK'
        ),
        'Color'
    ] = 'KS-BLACK'
    df_OrderHistory_date_and_party_filtered.loc[
        (
            df_OrderHistory_date_and_party_filtered['Article']
            .astype(str)
            .str.startswith('4')
        )
        & (
                df_OrderHistory_date_and_party_filtered['Color'] == 'BLACK'
        ),
        'Color'
    ] = 'KS-BLACK'
    df_OrderHistory_date_and_party_filtered.loc[
        (
            df_OrderHistory_date_and_party_filtered['Article']
            .astype(str)
            .str.startswith('3')
        )
        & (
                df_OrderHistory_date_and_party_filtered['Color'] == 'BLACK'
        ),
        'Color'
    ] = 'N-BLACK'
    df_OrderHistory_date_and_party_filtered.loc[
        (
            df_OrderHistory_date_and_party_filtered['Article']
            .astype(str)
            .str.startswith('6')
        )
        & (
                df_OrderHistory_date_and_party_filtered['Color'] == 'BLACK'
        ),
        'Color'
    ] = 'N-BLACK'
    df_OrderHistory_date_and_party_filtered.loc[
        (
            df_OrderHistory_date_and_party_filtered['Article']
            .astype(str)
            .str.startswith('2')
        )
        & (
                df_OrderHistory_date_and_party_filtered['Color'] == 'BLACK'
        ),
        'Color'
    ] = 'P-BLACK'
    df_OrderHistory_date_and_party_filtered.loc[
        (
            df_OrderHistory_date_and_party_filtered['Article']
            .astype(str)
            .str.startswith('5')
        )
        & (
                df_OrderHistory_date_and_party_filtered['Color'] == 'BLACK'
        ),
        'Color'
    ] = 'P-BLACK'
    # ============================================
    # Renaming black to OTHER BLACK ENDS
    # ============================================

    # ============================================
    # Item Mapped_shade vs Party starts
    # ============================================

    df_OrderHistory_grouped1 = df_OrderHistory_date_and_party_filtered.groupby(
        ['Item Mapped', 'Color', 'Party'],
        as_index=False
    ).agg(
        Last_date_of_Odr=('Order Date', 'max'),
        actual_tot_in_Qty=('Actual order', 'sum'),
        actual_tot_wt_in_KG=('actual_tot_wt_in_KG', 'sum'),
        Actual_order_Qty=('Actual order', 'sum'),
        Frequency=('actual_tot_wt_in_KG', 'count')
    )

    df_OrderHistory_grouped2 = df_OrderHistory_grouped1.groupby(
        ['Item Mapped', 'Color'],
        as_index=False
    ).agg(
        Last_date_of_Odr=('Last_date_of_Odr', 'max'),
        actual_tot_in_Qty=('actual_tot_in_Qty', 'sum'),
        actual_tot_wt_in_KG=('actual_tot_wt_in_KG', 'sum'),
        Total_Frequency=('Frequency', 'sum'),
        No_of_parties=('Party', 'count'),
        Parties=('Party', lambda x: ', '.join(x.astype(str))),
    )
    df_OrderHistory_grouped2.to_excel('Output_files/Output_all_Item_Shade_vs_Party_Odr_Hist.xlsx', index=False)



    return df_OrderHistory_grouped2,start_day, max_date

    # Item_list = df_OrderHistory_grouped2['Item Mapped'].astype(str).unique().tolist()
    # with pd.ExcelWriter('Output_files/Output_Item_wise_sheet_shade_vs_party_Odr_Hist.xlsx') as Writter:
    #     for item in Item_list:
    #         filtered_df = df_OrderHistory_grouped2[
    #             df_OrderHistory_grouped2['Item Mapped']==item]
    #         item = item.replace('/', ' by ')
    #         item = item.replace(' ', '_')
    #         sheet_name = item
    #         filtered_df.to_excel(Writter, sheet_name=sheet_name, index=False)

    # ============================================
    # Item Mapped_shade vs Party ends
    # ============================================













