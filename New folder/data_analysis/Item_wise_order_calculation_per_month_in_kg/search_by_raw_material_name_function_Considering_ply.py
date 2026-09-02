import pandas as pd
import matplotlib.pyplot as plt
import re





# file_name_1 = "OrderHistory_all.xlsx"
# file_dry_weight = "combined_dry_weight.xlsx"
# item_name = '210 HT CONTINOUS POLYESTER FILAMENT'
# input_article_number = 3200  #####This has to be changed while for CONTINOUS FILAMENT and spun
# search_by_raw_material_name_function(file_name_1,file_dry_weight,item_name,input_article_number)





def search_by_raw_material_name_function_Considering_ply(file_name_1,file_dry_weight,item_name,input_article_number):

    ####INPUT VARIABLES
    ####INPUT VARIABLES
    # file_name_1 = "C:/Users/amitk/Desktop/ooffice/New folder/OrderHistory_all.xlsx"
    # file_dry_weight = "C:/Users/amitk/Desktop/ooffice/New folder/combined_dry_weight.xlsx"
    # item_name = '210 HT CONTINOUS POLYESTER FILAMENT'
    # input_article_number = 3200  #####This has to be changed while for CONTINOUS FILAMENT and spun
    ####INPUT VARIABLES
    ####INPUT VARIABLES

    # ####   FOR regex
    # ####   FOR regex
    # ####   FOR regex
    # if str(input_article_number).startswith(('2', '3', '5', '6')):
    #     item_name = re.sub(r'/\d', '', item_name)
    # print(item_name)
    # ####   FOR regex
    # ####   FOR regex
    # ####   FOR regex

    df = pd.read_excel(file_name_1)
    print(df.columns)
    number_of_shade_toplot = 20
    df2 = df[['Order Date', 'Article', 'Color']].copy()
    # df2['Actual order'] = df['Order Qty'] - df['Cancel Qty']
    df2['Actual order'] = df['Order Qty']
    df['Date'] = pd.to_datetime(df['Order Date'], format="%d/%m/%Y")
    no_of_days = df['Date'].max() - df['Date'].min()
    print('No. of days:')
    print(no_of_days)
    del df
    df3 = df2.groupby(['Article', 'Color']).agg(
        Sum_order=('Actual order', 'sum'),
        Count_order=('Order Date', 'count')
    ).reset_index()
    del df2

    df_dry_weight = pd.read_excel(file_dry_weight)
    print(df_dry_weight.columns)
    # print(df_dry_weight.head())

    # ####   FOR regex
    # ####   FOR regex
    # ####   FOR regex
    #
    # df_dry_weight['Item Mapped'] = df_dry_weight.apply(
    #     lambda row: re.sub(r'/\d', '', row['Item Mapped'])
    #     if str(row['Article No.']).startswith(('2', '3', '5', '6'))
    #     else row['Item Mapped'],
    #     axis=1
    # )
    # # df_dry_weight.to_excel("C:/Users/amitk/Desktop/ooffice/New folder/temporary_output_regx.xlsx", index=True)
    # # print(df_dry_weight['Item Mapped'])
    # ####   FOR regex
    # ####   FOR regex
    # ####   FOR regex

    df_dry_weight_2 = df_dry_weight.groupby(['Item Mapped', 'Article No.']).agg(
        Weight=('Dry weight(gram)', 'max'),
    ).reset_index()

    print(df_dry_weight_2[df_dry_weight_2['Item Mapped'] == item_name])
    df_dry_weight_3 = df_dry_weight_2[df_dry_weight_2['Item Mapped'] == item_name]
    len_of_df_dry_weight_3 = len(df_dry_weight_3)

    total_order_for_item_WITHOUT_BW = 0
    total_order_for_item_for_BW = 0
    total_order_for_item = 0
    for i in range(len_of_df_dry_weight_3):
        article_number = df_dry_weight_3.iloc[i]['Article No.']
        weight_of_cone = df_dry_weight_3.iloc[i]['Weight']
        print('Article number:')
        print(article_number)
        print('weight_of_cone_in_gm:')
        print(weight_of_cone)

        article_list = [article_number]

        df4 = df3[df3['Article'] == article_list[0]]
        sorted_df4 = df4.sort_values(by='Sum_order', ascending=False)
        print(sorted_df4)
        # if article_list[0]==1201:
        #     sorted_df4.to_excel("C:/Users/amitk/Desktop/ooffice/New folder/temporary_output2.xlsx", index=True)
        del df4
        length_of_data = len(sorted_df4)
        # print(sorted_df4)
        size_plot = length_of_data

        color_list = sorted_df4['Color'].tolist()
        sum_order_list = sorted_df4['Sum_order'].tolist()
        print(color_list)
        print(sum_order_list)
        print(len(color_list))

        total_order_for_article = sum(sum_order_list)
        print('total_order_for_article')
        print(total_order_for_article)
        # print(sorted_df4['Sum_order'].count())
        total_order_for_item = total_order_for_item + total_order_for_article * weight_of_cone / 1000

        ###### DON'T CHANGE UNTILL THIS
        ###### DON'T CHANGE UNTILL THIS
        ###### DON'T CHANGE UNTILL THIS

        total_order_for_article_for_BW = 0
        total_order_for_article_WITHOUT_BW = 0

        for j in range(len(color_list)):
            if color_list[j] == 'BW' or color_list[j] == 'KS-BLEACH':
                total_order_for_article_for_BW = total_order_for_article_for_BW + sum_order_list[j]
            else:
                total_order_for_article_WITHOUT_BW = total_order_for_article_WITHOUT_BW + sum_order_list[j]

        print('total_order_for_article_for_BW:')
        print(total_order_for_article_for_BW)
        print('total_order_for_article_WITHOUT_BW:')
        print(total_order_for_article_WITHOUT_BW)
        total_order_for_item_for_BW = total_order_for_item_for_BW + total_order_for_article_for_BW * weight_of_cone / 1000
        total_order_for_item_WITHOUT_BW = total_order_for_item_WITHOUT_BW + total_order_for_article_WITHOUT_BW * weight_of_cone / 1000

    total_order_for_item_per_month_in_kg = total_order_for_item / no_of_days.days * 30
    print('Total order in KG for ' + item_name + ' per month :')
    print(total_order_for_item_per_month_in_kg)

    total_order_for_item_for_BW_per_month_in_kg = total_order_for_item_for_BW / no_of_days.days * 30
    print('Total order in KG for ' + item_name + ' per month for BW:')
    print(total_order_for_item_for_BW_per_month_in_kg)

    total_order_for_item_WITHOUT_BW_per_month_in_kg = total_order_for_item_WITHOUT_BW / no_of_days.days * 30
    print('Total order  in KG for ' + item_name + ' per month WITHOUT BW:')
    print(total_order_for_item_WITHOUT_BW_per_month_in_kg)

    # df_dry_weight_2.to_excel("C:/Users/amitk/Desktop/ooffice/New folder/temporary_output.xlsx", index=True)
    # print(df_dry_weight_2)



    return [item_name,total_order_for_item_per_month_in_kg,total_order_for_item_for_BW_per_month_in_kg,total_order_for_item_WITHOUT_BW_per_month_in_kg]




# ####INPUT VARIABLES
# ####INPUT VARIABLES
# file_name_1 = "OrderHistory_all.xlsx"
# file_dry_weight = "combined_dry_weight.xlsx"
# item_name = '2/60 SPUN POLYESTER YARN'
# input_article_number = 1200  #####This has to be changed while for CONTINOUS FILAMENT and spun
# ####INPUT VARIABLES
# ####INPUT VARIABLES
# search_by_raw_material_name_function_Considering_ply(file_name_1,file_dry_weight,item_name,input_article_number)