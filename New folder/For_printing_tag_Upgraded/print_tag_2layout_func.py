from cmath import nan

from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import pandas as pd

def print_tag_func(dict_to_call_func):
    # image_name = "2.jpeg"
    # party_name = "ZIGMA"
    # article_num = 1201
    # llength = 1500
    # shade = "S1005"
    # brand_name = 'ROX gold'
    # batch_NO = 'DS00012'


    image_name = dict_to_call_func['image_name']
    party_name = dict_to_call_func['party_name']
    article_num = dict_to_call_func['article_num']
    llength = dict_to_call_func['llength']
    shade = dict_to_call_func['shade']
    brand_name = dict_to_call_func['brand_name']
    batch_NO = dict_to_call_func['batch_NO']
    date_to_show=dict_to_call_func['date_to_show']

    file_name_Spacing_position_DB = "DB_files/Spacing_position_DB.xlsx"
    df_Spacing_position_DB = pd.read_excel(file_name_Spacing_position_DB)
    print(df_Spacing_position_DB.columns)
    # print(df_Spacing_position_DB.head())




    draw_party_logo_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'draw_party_logo'].iloc[0].to_dict()
    party_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'party_name'].iloc[0].to_dict()
    brand_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'brand_name'].iloc[0].to_dict()
    ART_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'ART_name'].iloc[0].to_dict()
    shade_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'shade'].iloc[0].to_dict()
    batch_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'batch_name'].iloc[0].to_dict()
    llength_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'llength_name'].iloc[0].to_dict()
    date_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'date_name'].iloc[0].to_dict()

    layout_x_times_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'layout_x_times'].iloc[0].to_dict()
    layout_y_times_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'layout_y_times'].iloc[
        0].to_dict()
    layout_x_spacing_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'layout_x_spacing'].iloc[
        0].to_dict()
    layout_y_spacing_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'layout_y_spacing'].iloc[
        0].to_dict()



    # ART_name = "ART: " + str(article_num)
    # llength_name = str(llength) + ' MTR'
    # batch_name = "Batch No.: " + batch_NO
    # image = ImageReader("logos/" + image_name)

    ART_name = ("" if pd.isna(ART_name_dict['Prefix']) else str(ART_name_dict['Prefix'])) + str(article_num)+ ("" if pd.isna(ART_name_dict['Suffix']) else str(ART_name_dict['Suffix']))
    llength_name = ("" if pd.isna(llength_name_dict['Prefix']) else str(llength_name_dict['Prefix'])) + str(llength) +("" if pd.isna(llength_name_dict['Suffix']) else str(llength_name_dict['Suffix']))
    batch_name = ("" if pd.isna(batch_name_dict['Prefix']) else str(batch_name_dict['Prefix'])) + batch_NO + ("" if pd.isna(batch_name_dict['Suffix']) else str(batch_name_dict['Suffix']))


    image = ImageReader("logos/" + image_name)



    # layout_x_times=2
    # layout_y_times = 1
    # layout_x_spacing = 1.77 * 72  # 30*45mm # 1 inch=72 points
    # layout_y_spacing = 1.18 * 72


    layout_x_times = layout_x_times_dict['Value']
    layout_y_times = layout_y_times_dict['Value']
    layout_x_spacing = layout_x_spacing_dict['Value']  # 30*45mm
    layout_y_spacing = layout_y_spacing_dict['Value']
    width = layout_x_spacing * layout_x_times
    height = layout_y_spacing * layout_y_times
    file_name_to_save = 'Output_files/'+str(article_num) + '_' + str(shade) + '_' + party_name + '_tag.pdf'





    pdf = canvas.Canvas(file_name_to_save, pagesize=(width, height))










    for j in range(int(layout_y_times)):
        for i in range(int(layout_x_times)):
            x_pacing_add =i*layout_x_spacing
            y_pacing_add =j*layout_y_spacing
            print('spacing:'+str(x_pacing_add))

            if draw_party_logo_dict['Active'] == 1:
                pdf.drawImage(
                    image,
                    2 + x_pacing_add,  # x position
                    layout_y_spacing-int(draw_party_logo_dict['Size'])-2 + y_pacing_add,  # y position
                    width=int(draw_party_logo_dict['Size']),  # image width
                    height=int(draw_party_logo_dict['Size'])  # image height
                )

            if party_name_dict['Active'] == 1:
                font_name = "Helvetica-Bold"
                font_size = party_name_dict['Size']

                pdf.setFont(font_name, font_size)

                # Calculate width of party name
                party_name_width = pdf.stringWidth(
                    str(party_name),
                    font_name,
                    font_size
                )

                # Automatically center party name horizontally
                party_x_position = (layout_x_spacing - party_name_width) / 2

                pdf.drawString(
                    party_x_position + party_name_dict['Off_set_x_must_be_0'] + x_pacing_add,
                    party_name_dict['y_position'] + y_pacing_add,
                    str(party_name)
                )

            if brand_name_dict['Active'] == 1:
                font_name = "Helvetica-Bold"
                font_size = brand_name_dict['Size']

                pdf.setFont(font_name, font_size)

                # Calculate width of brand name
                brand_name_width = pdf.stringWidth(
                    str(brand_name),
                    font_name,
                    font_size
                )

                # Automatically center brand name horizontally
                brand_x_position = (layout_x_spacing - brand_name_width) / 2

                pdf.drawString(
                    brand_x_position+brand_name_dict['Off_set_x_must_be_0']  + x_pacing_add,
                    brand_name_dict['y_position'] + y_pacing_add,
                    str(brand_name)
                )

            if ART_name_dict['Active'] == 1:
                font_name = "Helvetica-Bold"
                font_size = ART_name_dict['Size']

                pdf.setFont(font_name, font_size)

                # Calculate width of ART name
                ART_name_width = pdf.stringWidth(
                    str(ART_name),
                    font_name,
                    font_size
                )

                # Automatically center ART name horizontally
                ART_x_position = (layout_x_spacing - ART_name_width) / 2

                pdf.drawString(
                    ART_x_position+ART_name_dict['Off_set_x_must_be_0'] + x_pacing_add,
                    ART_name_dict['y_position'] + y_pacing_add,
                    str(ART_name)
                )

            if shade_dict['Active'] == 1:
                font_name = "Helvetica-Bold"
                font_size = shade_dict['Size']

                pdf.setFont(font_name, font_size)

                # Calculate width of shade text
                shade_width = pdf.stringWidth(
                    str(shade),
                    font_name,
                    font_size
                )

                # Center the shade inside the tag
                shade_x_position = (layout_x_spacing - shade_width) / 2

                pdf.drawString(
                    shade_x_position+shade_dict['Off_set_x_must_be_0'] + x_pacing_add,
                    shade_dict['y_position'] + y_pacing_add,
                    str(shade)
                )

            if batch_name_dict['Active'] == 1:
                font_name = "Helvetica-Bold"
                font_size = batch_name_dict['Size']

                pdf.setFont(font_name, font_size)

                # Calculate width of batch name
                batch_width = pdf.stringWidth(
                    str(batch_name),
                    font_name,
                    font_size
                )

                # Automatically center batch_name horizontally
                batch_x_position = (layout_x_spacing - batch_width) / 2

                pdf.drawString(
                    batch_x_position+batch_name_dict['Off_set_x_must_be_0'] + x_pacing_add,
                    batch_name_dict['y_position'] + y_pacing_add,
                    str(batch_name)
                )

            if llength_name_dict['Active'] == 1:
                font_name = "Helvetica-Bold"
                font_size = llength_name_dict['Size']

                pdf.setFont(font_name, font_size)

                # Calculate width of llength text
                llength_width = pdf.stringWidth(
                    str(llength_name),
                    font_name,
                    font_size
                )

                # Automatically center llength_name horizontally
                llength_x_position = (layout_x_spacing - llength_width) / 2

                pdf.drawString(
                    llength_x_position+llength_name_dict['Off_set_x_must_be_0'] + x_pacing_add,
                    llength_name_dict['y_position'] + y_pacing_add,
                    str(llength_name)
                )
            #########
            if date_name_dict['Active'] == 1:
                font_name = "Helvetica-Bold"
                font_size = date_name_dict['Size']

                pdf.setFont(font_name, font_size)

                # Calculate width of llength text
                date_width = pdf.stringWidth(
                    str(date_to_show),
                    font_name,
                    font_size
                )

                # Automatically center llength_name horizontally
                date_x_position = (layout_x_spacing - date_width) / 2

                pdf.drawString(
                    date_x_position+date_name_dict['Off_set_x_must_be_0'] + x_pacing_add,
                    date_name_dict['y_position'] + y_pacing_add,
                    str(date_to_show)
                )











    # pdf.setFont("Helvetica-Bold", 15)
    # pdf.drawString(45, 70, party_name)
    # pdf.setFont("Helvetica-Bold", 12)
    # pdf.drawString(40, 58, brand_name)
    # pdf.setFont("Helvetica-Bold", 8)
    # pdf.drawString(42, 48, ART_name)
    # pdf.setFont("Helvetica-Bold", 15)
    # pdf.drawString(30, 30, shade)
    # pdf.setFont("Helvetica-Bold", 8)
    # pdf.drawString(30, 20, batch_name)
    # pdf.setFont("Helvetica-Bold", 8)
    # pdf.drawString(45, 5, llength_name)
    pdf.save()
    return file_name_to_save


# image_name = "3.jpeg"
# party_name = "ZIGMA"
# article_num = 1205
# llength = 1500
# shade = "S1005"
# brand_name = 'ROX gold'
# batch_NO = 'DS00012'
# print_tag_func(image_name,party_name,article_num,llength,shade,brand_name,batch_NO)