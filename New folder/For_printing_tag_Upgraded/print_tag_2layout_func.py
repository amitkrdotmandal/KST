from cmath import nan

from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import pandas as pd

def print_tag_func(image_name,party_name,article_num,llength,shade,brand_name,batch_NO):
    # image_name = "2.jpeg"
    # party_name = "ZIGMA"
    # article_num = 1201
    # llength = 1500
    # shade = "S1005"
    # brand_name = 'ROX gold'
    # batch_NO = 'DS00012'

    file_name_Spacing_position_DB = "DB_files/Spacing_position_DB.xlsx"
    df_Spacing_position_DB = pd.read_excel(file_name_Spacing_position_DB)
    print(df_Spacing_position_DB.columns)
    # print(df_Spacing_position_DB.head())

    ART_name = "ART: " + str(article_num)
    llength_name = str(llength) + ' MTR'
    batch_name = "Batch No.: " + batch_NO
    image = ImageReader("logos/" + image_name)

    party_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'party_name'].iloc[0].to_dict()
    brand_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'brand_name'].iloc[0].to_dict()
    ART_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'ART_name'].iloc[0].to_dict()
    shade_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'shade'].iloc[0].to_dict()
    batch_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'batch_name'].iloc[0].to_dict()
    llength_name_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'llength_name'].iloc[0].to_dict()

    layout_x_times_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'layout_x_times'].iloc[0].to_dict()
    layout_y_times_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'layout_y_times'].iloc[
        0].to_dict()
    layout_x_spacing_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'layout_x_spacing'].iloc[
        0].to_dict()
    layout_y_spacing_dict = df_Spacing_position_DB[df_Spacing_position_DB['Field_name'] == 'layout_y_spacing'].iloc[
        0].to_dict()

    # layout_x_times=2
    # layout_y_times = 1
    # layout_x_spacing = 1.77 * 72  # 30*45mm
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
            print('spaing:'+str(x_pacing_add))

            pdf.drawImage(
                image,
                1+x_pacing_add,  # x position
                55+y_pacing_add,  # y position
                width=30,  # image width
                height=30  # image height
            )

            pdf.setFont("Helvetica-Bold", party_name_dict['Size'])
            pdf.drawString(party_name_dict['x_position']+x_pacing_add, party_name_dict['y_position']+y_pacing_add, party_name)
            pdf.setFont("Helvetica-Bold", brand_name_dict['Size'])
            pdf.drawString(brand_name_dict['x_position']+x_pacing_add, brand_name_dict['y_position']+y_pacing_add, brand_name)
            pdf.setFont("Helvetica-Bold", ART_name_dict['Size'])
            pdf.drawString(ART_name_dict['x_position']+x_pacing_add, ART_name_dict['y_position']+y_pacing_add, ART_name)
            pdf.setFont("Helvetica-Bold", shade_dict['Size'])
            pdf.drawString(shade_dict['x_position']+x_pacing_add, shade_dict['y_position']+y_pacing_add, shade)
            pdf.setFont("Helvetica-Bold", batch_name_dict['Size'])
            pdf.drawString(batch_name_dict['x_position']+x_pacing_add, batch_name_dict['y_position']+y_pacing_add, batch_name)
            pdf.setFont("Helvetica-Bold", llength_name_dict['Size'])
            pdf.drawString(llength_name_dict['x_position']+x_pacing_add, llength_name_dict['y_position']+y_pacing_add, llength_name)





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