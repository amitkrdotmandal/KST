import pdfplumber
import pandas as pd
import re

# pdf_path = "ST01166-26 - Challan.pdf"
pdf_path = "ST01097-26 - Challan.pdf"

file_name_party_list_DB = "party_list_DB.xlsx"
df_file_name_party_list_DB = pd.read_excel(file_name_party_list_DB)
print(df_file_name_party_list_DB.columns)
# print(df_file_name_party_list_DB.head())








with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
    text = page.extract_text()

buyer_details = re.search(
    r"BUYER DETAILS(.*?)ARTICLE DETAILS",
    text,
    re.DOTALL
).group(1).strip()
# print(buyer_details)
# print(text)
buyer_details_lists = re.split(r'\n', buyer_details)
probable_party_name = buyer_details_lists[1]
with open("output.txt", "w", encoding="utf-8") as f:
    f.write(probable_party_name)
for i in range(len(df_file_name_party_list_DB)):
    party_name = df_file_name_party_list_DB['party_list'].iloc[i]
    if re.search(re.escape(party_name), probable_party_name):
        actual_party_name = party_name
        break
    else:
        actual_party_name = None



###ITEM list extraction starts
###ITEM list extraction starts
###ITEM list extraction starts
###ITEM list extraction starts
ITEMS = re.search(
    r"ITEMS(.*?)Totals",
    text,
    re.DOTALL
).group(1).strip()
# print(ITEMS)
df_items = pd.DataFrame(columns=["Sr No", "Article","Shade"])

item_lists = re.split(r'\n', ITEMS)
for i in range(1,len(item_lists)):
    print(item_lists[i])
    words_in_item= item_lists[i].split(" ")
    df_items.loc[i, "Sr No"] = int(words_in_item[0])
    print(words_in_item)
    if len(words_in_item)==7:
        start_pt=2
    else:
        start_pt = 1
    df_items.loc[i, "Party_name"] = actual_party_name
    df_items.loc[i, "Article"] = words_in_item[start_pt]
    df_items.loc[i, "Shade"] = words_in_item[start_pt+1]

    if words_in_item[start_pt+3]!="-":
        df_items.loc[i, "is_cone"] = 0
        df_items.loc[i, "is_box"] = 1
        qty_no = words_in_item[start_pt + 3].replace(",", "")
        df_items.loc[i, "Quantity_of_box_or_cone"] = int(qty_no)
    if words_in_item[start_pt+4]!="-":
        df_items.loc[i, "is_cone"] = 1
        df_items.loc[i, "is_box"] = 0
        qty_no=words_in_item[start_pt+4].replace(",", "")
        df_items.loc[i, "Quantity_of_box_or_cone"] = int(qty_no)


df_items.to_excel('tempo_output_ITEMS.xlsx', index=False)

###ITEM list extraction ENDS
###ITEM list extraction ENDS
###ITEM list extraction ENDS
###ITEM list extraction ENDS

