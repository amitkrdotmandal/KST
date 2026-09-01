import tkinter as tk
from tkinter import ttk
import pandas as pd
from print_tag_func import print_tag_func



file_name_combined_dry_weight_DB="DB_files/combined_dry_weight_DB.xlsx"
file_name_Party_logo_name_DB="DB_files/Party_logo_name_DB.xlsx"
file_name_shade_library_DB="DB_files/shade_library_DB.xlsx"
df_combined_dry_weight_DB= pd.read_excel(file_name_combined_dry_weight_DB)
print(df_combined_dry_weight_DB.columns)
# print(df_combined_dry_weight_DB.head())
df_Party_logo_name_DB= pd.read_excel(file_name_Party_logo_name_DB)
print(df_Party_logo_name_DB.columns)
# print(df_Party_logo_name_DB.head())
df_shade_library_DB= pd.read_excel(file_name_shade_library_DB)
print(df_shade_library_DB.columns)
# print(df_shade_library_DB.head())

party_list =df_Party_logo_name_DB['Party'].tolist()
article_list = df_combined_dry_weight_DB['Article No.'].astype(str).tolist()
shade_list =df_shade_library_DB['Shade no'].tolist()




window = tk.Tk()
window.title("Tag_PDF_maker")
window.geometry("450x450")

#
# party_list = [
#     "Select Party",
#     "ABC Industries",
#     "Kiran Industries",
#     "Kiran Textiles",
#     "XYZ Industries"
# ]
#
# article_list = [
#     "Select article",
#     "Cotton",
#     "Polyester",
#     "Viscose",
#     "Spun Polyester Yarn"
# ]
#
# shade_list = [
#     "Select Color",
#     "Red",
#     "Blue",
#     "Black",
#     "Royal Blue"
# ]


# -------------------------
# Autocomplete function
# -------------------------

def autocomplete(event, dropdown, original_list):

    typed = dropdown.get().lower()

    filtered = [
        item for item in original_list
        if typed in item.lower()
    ]

    dropdown["values"] = filtered

    if filtered:
        dropdown.event_generate("<Down>")


# -------------------------
# Party
# -------------------------

tk.Label(window, text="Party Name").pack(pady=(20, 5))

party_dropdown = ttk.Combobox(
    window,
    values=party_list,
    width=40
)

party_dropdown.pack()
party_dropdown.set(party_list[0])

party_dropdown.bind(
    "<KeyRelease>",
    lambda event: autocomplete(
        event,
        party_dropdown,
        party_list
    )
)


# -------------------------
# Article
# -------------------------

tk.Label(window, text="Article No.").pack(pady=(20, 5))

article_dropdown = ttk.Combobox(
    window,
    values=article_list,
    width=40
)

article_dropdown.pack()
article_dropdown.set(article_list[0])

article_dropdown.bind(
    "<KeyRelease>",
    lambda event: autocomplete(
        event,
        article_dropdown,
        article_list
    )
)


# -------------------------
# Shade
# -------------------------

tk.Label(window, text="Shade No.").pack(pady=(20, 5))

Shade_dropdown = ttk.Combobox(
    window,
    values=shade_list,
    width=40
)

Shade_dropdown.pack()
Shade_dropdown.set(shade_list[0])

Shade_dropdown.bind(
    "<KeyRelease>",
    lambda event: autocomplete(
        event,
        Shade_dropdown,
        shade_list
    )
)

# -------------------------
# Batch no
# -------------------------

tk.Label(window, text="Enter Batch No. :").pack(pady=(20, 5))

batch_no_entry = tk.Entry(window, width=42)
batch_no_entry.pack()


#######



# -------------------------
# Submit function
# -------------------------

def submit():
    party = party_dropdown.get()
    article = article_dropdown.get()
    shade = Shade_dropdown.get()
    batch_no = batch_no_entry.get()


    errorr=1
    if party in party_list and article in article_list and shade in shade_list and batch_no != '':
        errorr=0

    if errorr == 0:

        image_name = df_Party_logo_name_DB.loc[df_Party_logo_name_DB['Party'] == party, 'logo'].iloc[0]
        party_name = df_Party_logo_name_DB.loc[df_Party_logo_name_DB['Party'] == party, 'name'].iloc[0]
        if pd.isna(party_name) or party_name == "nan":
            party_name=""
        # print(party_name)
        article_num = article
        llength = df_combined_dry_weight_DB.loc[df_combined_dry_weight_DB['Article No.'].astype(str) == article, 'Length'].iloc[0]
        shade =shade
        brand_name = df_combined_dry_weight_DB.loc[df_combined_dry_weight_DB['Article No.'].astype(str) == article, 'Brand'].iloc[0]
        batch_NO = batch_no







        msg_from_func=print_tag_func(image_name,party_name,article_num,llength,shade,brand_name,batch_NO)






        result_label.config(
            text=f"Saved as: {msg_from_func}\n"


        )
    else:
        result_label.config(
            text=f"Enter data correctly.\n"


        )









# -------------------------
# Submit button
# -------------------------

submit_button = tk.Button(
    window,
    text="Submit",
    command=submit
)

submit_button.pack(pady=20)


# -------------------------
# Result Label
# -------------------------

result_label = tk.Label(
    window,
    text="",
    font=("Arial", 12),
    justify="left"
)

result_label.pack(pady=10)


window.mainloop()