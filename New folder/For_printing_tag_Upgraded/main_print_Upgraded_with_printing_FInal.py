
import tkinter as tk
import pandas as pd
from PIL import Image, ImageTk, ImageWin
import pypdfium2 as pdfium

import win32print
import win32ui
from win32con import HORZRES, VERTRES

# from print_tag_func import print_tag_func
from print_tag_2layout_func import print_tag_func
from print_to_printer_func import print_to_printer_func


# ============================================================
# LOAD DATABASE FILES
# ============================================================

file_name_combined_dry_weight_DB = (
    "DB_files/combined_dry_weight_DB.xlsx"
)

file_name_Party_logo_name_DB = (
    "DB_files/Party_logo_name_DB.xlsx"
)

file_name_shade_library_DB = (
    "DB_files/shade_library_DB.xlsx"
)


# ============================================================
# READ EXCEL FILES
# ============================================================

df_combined_dry_weight_DB = pd.read_excel(
    file_name_combined_dry_weight_DB
)

df_Party_logo_name_DB = pd.read_excel(
    file_name_Party_logo_name_DB
)

df_shade_library_DB = pd.read_excel(
    file_name_shade_library_DB
)


# ============================================================
# CREATE LISTS
# ============================================================

party_list = (
    df_Party_logo_name_DB["Party"]
    .dropna()
    .astype(str)
    .tolist()
)

article_list = (
    df_combined_dry_weight_DB["Article No."]
    .dropna()
    .astype(str)
    .tolist()
)

shade_list = (
    df_shade_library_DB["Shade no"]
    .dropna()
    .astype(str)
    .tolist()
)


# ============================================================
# MAIN WINDOW
# ============================================================

window = tk.Tk()

window.title("Sticker_PDF_maker")

window.geometry("1000x700")

submit_already_pressed = 0
pdf_=None
image_ = None
photo_ = None
print_path_to_pass=None
# ============================================================
# MAIN FRAME
# ============================================================

main_frame = tk.Frame(
    window
)

main_frame.pack(
    fill="both",
    expand=True
)


# ============================================================
# LEFT FRAME
# ============================================================

left_frame = tk.Frame(
    main_frame
)

left_frame.pack(
    side="left",
    fill="y",
    padx=20,
    pady=20
)


# ============================================================
# RIGHT FRAME
# ============================================================

right_frame = tk.Frame(
    main_frame,
    bd=2,
    relief="groove",

)

right_frame.pack(
    side="right",
    fill="both",
    expand=True,
    padx=20,
    pady=20
)


# ============================================================
# RIGHT CANVAS
# ============================================================

right_canvas = tk.Canvas(
    right_frame,
    bg="lightgray"
)

right_canvas.pack(
    fill="both",
    expand=True
)


# ============================================================
# SEARCH BOX + FILTERED LIST
# ============================================================

def create_search_box(parent, values, width=42, list_height=6):

    # --------------------------------------------------------
    # Frame
    # --------------------------------------------------------

    frame = tk.Frame(parent)

    # --------------------------------------------------------
    # Entry / Textbox
    # --------------------------------------------------------

    entry = tk.Entry(
        frame,
        width=width
    )

    entry.pack(
        anchor="w"
    )

    # --------------------------------------------------------
    # Listbox
    # --------------------------------------------------------

    listbox = tk.Listbox(
        frame,
        width=width,
        height=list_height
    )

    listbox.pack(
        anchor="w"
    )

    # --------------------------------------------------------
    # Initially show all values
    # --------------------------------------------------------

    for item in values:

        listbox.insert(
            tk.END,
            item
        )

    # --------------------------------------------------------
    # FILTER FUNCTION
    # --------------------------------------------------------

    def filter_list(event=None):

        typed = entry.get().lower()

        # Delete previous results
        listbox.delete(
            0,
            tk.END
        )

        # Filter values
        filtered = [
            item
            for item in values
            if typed in str(item).lower()
        ]

        # Insert filtered values
        for item in filtered:

            listbox.insert(
                tk.END,
                item
            )

        # Select first result
        if filtered:

            listbox.selection_clear(
                0,
                tk.END
            )

            listbox.selection_set(
                0
            )

            listbox.activate(
                0
            )

        # Keep cursor in textbox
        entry.focus_set()

        entry.icursor(
            tk.END
        )

    # --------------------------------------------------------
    # SELECT ITEM
    # --------------------------------------------------------

    def select_item(event=None):

        selection = listbox.curselection()

        if not selection:
            return

        selected_value = listbox.get(
            selection[0]
        )

        # Put selected value into textbox
        entry.delete(
            0,
            tk.END
        )

        entry.insert(
            0,
            selected_value
        )

        # Keep focus in textbox
        entry.focus_set()

        entry.icursor(
            tk.END
        )

    # --------------------------------------------------------
    # KEYBOARD NAVIGATION
    # --------------------------------------------------------

    def keyboard_navigation(event):

        current_selection = listbox.curselection()

        # -----------------------------------------
        # DOWN
        # -----------------------------------------

        if event.keysym == "Down":

            if listbox.size() == 0:
                return "break"

            if not current_selection:

                index = 0

            else:

                index = current_selection[0] + 1

            if index >= listbox.size():

                index = listbox.size() - 1

            listbox.selection_clear(
                0,
                tk.END
            )

            listbox.selection_set(
                index
            )

            listbox.activate(
                index
            )

            listbox.see(
                index
            )

            return "break"

        # -----------------------------------------
        # UP
        # -----------------------------------------

        elif event.keysym == "Up":

            if listbox.size() == 0:
                return "break"

            if not current_selection:

                index = 0

            else:

                index = current_selection[0] - 1

            if index < 0:

                index = 0

            listbox.selection_clear(
                0,
                tk.END
            )

            listbox.selection_set(
                index
            )

            listbox.activate(
                index
            )

            listbox.see(
                index
            )

            return "break"

        # -----------------------------------------
        # ENTER
        # -----------------------------------------

        elif event.keysym == "Return":

            select_item()

            return "break"

    # --------------------------------------------------------
    # EVENTS
    # --------------------------------------------------------

    entry.bind(
        "<KeyRelease>",
        filter_list
    )

    entry.bind(
        "<Down>",
        keyboard_navigation
    )

    entry.bind(
        "<Up>",
        keyboard_navigation
    )

    entry.bind(
        "<Return>",
        keyboard_navigation
    )

    listbox.bind(
        "<ButtonRelease-1>",
        select_item
    )

    # --------------------------------------------------------
    # RETURN ENTRY
    # --------------------------------------------------------

    return entry


# ============================================================
# FORM FRAME
# ============================================================

form_frame = tk.Frame(
    left_frame
)

form_frame.pack(
    padx=0,
    pady=0,
    anchor="w"
)


# ============================================================
# PARTY NAME
# ============================================================

party_label = tk.Label(
    form_frame,
    text="Party Name:",
    width=15,
    anchor="w"
)

party_label.grid(
    row=0,
    column=0,
    padx=(0, 15),
    pady=10,
    sticky="nw"
)


party_dropdown = create_search_box(
    form_frame,
    party_list
)

party_dropdown.master.grid(
    row=0,
    column=1,
    padx=0,
    pady=10,
    sticky="nw"
)


# ============================================================
# ARTICLE NUMBER
# ============================================================

article_label = tk.Label(
    form_frame,
    text="Article No.:",
    width=15,
    anchor="w"
)

article_label.grid(
    row=1,
    column=0,
    padx=(0, 15),
    pady=10,
    sticky="nw"
)


article_dropdown = create_search_box(
    form_frame,
    article_list
)

article_dropdown.master.grid(
    row=1,
    column=1,
    padx=0,
    pady=10,
    sticky="nw"
)


# ============================================================
# SHADE NUMBER
# ============================================================

shade_label = tk.Label(
    form_frame,
    text="Shade No.:",
    width=15,
    anchor="w"
)

shade_label.grid(
    row=2,
    column=0,
    padx=(0, 15),
    pady=10,
    sticky="nw"
)


Shade_dropdown = create_search_box(
    form_frame,
    shade_list
)

Shade_dropdown.master.grid(
    row=2,
    column=1,
    padx=0,
    pady=10,
    sticky="nw"
)


# ============================================================
# BATCH NUMBER
# ============================================================

batch_label = tk.Label(
    form_frame,
    text="Batch No.:",
    width=15,
    anchor="w"
)

batch_label.grid(
    row=3,
    column=0,
    padx=(0, 15),
    pady=10,
    sticky="w"
)


batch_no_entry = tk.Entry(
    form_frame,
    width=42
)

batch_no_entry.grid(
    row=3,
    column=1,
    padx=0,
    pady=10,
    sticky="w"
)


# ==================================================
# OPEN PDF
# ==================================================
def open_pdf(file_path):
    global pdf_
    # print('IN open_pdf')
    file_path = file_path
    # print(file_path)

    if not file_path:
        return

    try:

        pdf_ = pdfium.PdfDocument(file_path)
        print(pdf_)



        show_page()

    except Exception as e:

        print("Error:", e)



# ==================================================
# SHOW PAGE
# ==================================================
def show_page():
    global pdf_
    global image_
    global photo_
    global right_canvas

    if pdf_ is None:
        print('IN show_page none')
        return

    # Get current PDF page
    page = pdf_[0]

    # Render PDF page
    bitmap = page.render(
        scale=1
    )

    # Convert to PIL image
    image_ = bitmap.to_pil()

    # Convert PIL image to Tkinter image
    photo_ = ImageTk.PhotoImage(
        image_
    )


        # Clear canvas
    right_canvas.delete("all")

    # Display image
    right_canvas.create_image(
        0,
        0,
        image=photo_,
        anchor="nw"
    )








# ============================================================
# print FUNCTION
# ============================================================

def print_cmd():
    global print_path_to_pass
    global window
    print_to_printer_func(window,print_path_to_pass)
    # pass

# ============================================================
# SUBMIT FUNCTION
# ============================================================

def submit():
    global print_path_to_pass

    # --------------------------------------------------------
    # GET VALUES
    # --------------------------------------------------------

    party = party_dropdown.get()

    article = article_dropdown.get()

    shade = Shade_dropdown.get()

    batch_no = batch_no_entry.get()


    # --------------------------------------------------------
    # VALIDATION
    # --------------------------------------------------------

    errorr = 1

    if (
        party in party_list
        and article in article_list
        and shade in shade_list
        and batch_no != ""
    ):

        errorr = 0


    # --------------------------------------------------------
    # VALID DATA
    # --------------------------------------------------------

    if errorr == 0:

        # ----------------------------------------------------
        # GET PARTY LOGO
        # ----------------------------------------------------

        image_name = (
            df_Party_logo_name_DB
            .loc[
                df_Party_logo_name_DB["Party"] == party,
                "logo"
            ]
            .iloc[0]
        )


        # ----------------------------------------------------
        # GET PARTY DISPLAY NAME
        # ----------------------------------------------------

        party_name = (
            df_Party_logo_name_DB
            .loc[
                df_Party_logo_name_DB["Party"] == party,
                "name"
            ]
            .iloc[0]
        )


        if pd.isna(party_name) or party_name == "nan":

            party_name = ""


        # ----------------------------------------------------
        # ARTICLE NUMBER
        # ----------------------------------------------------

        article_num = article


        # ----------------------------------------------------
        # GET LENGTH
        # ----------------------------------------------------

        llength = (
            df_combined_dry_weight_DB
            .loc[
                df_combined_dry_weight_DB[
                    "Article No."
                ].astype(str) == article,
                "Length"
            ]
            .iloc[0]
        )


        # ----------------------------------------------------
        # GET BRAND
        # ----------------------------------------------------

        brand_name = (
            df_combined_dry_weight_DB
            .loc[
                df_combined_dry_weight_DB[
                    "Article No."
                ].astype(str) == article,
                "Brand"
            ]
            .iloc[0]
        )


        # ----------------------------------------------------
        # BATCH NUMBER
        # ----------------------------------------------------

        batch_NO = batch_no


        # ----------------------------------------------------
        # CALL PRINT FUNCTION
        # ----------------------------------------------------

        msg_from_func = print_tag_func(
            image_name,
            party_name,
            article_num,
            llength,
            shade,
            brand_name,
            batch_NO
        )


        # ----------------------------------------------------
        # SHOW RESULT
        # ----------------------------------------------------

        result_label.config(
            text=f"Saved as: {msg_from_func}"
        )
        submit_already_pressed=1
        print_button.config(
            state="normal"
        )
        open_pdf(msg_from_func)
        print_path_to_pass=msg_from_func


    else:

        result_label.config(
            text="Enter data correctly."
        )



# ============================================================
# SUBMIT BUTTON
# ============================================================

submit_button = tk.Button(
    left_frame,
    text="Submit",
    width=15,
    command=submit
)

submit_button.pack(
    pady=15
)
# ============================================================
# print BUTTON
# ============================================================

print_button = tk.Button(
    left_frame,
    text="Print",
    width=15,
    command=print_cmd,
    state="disabled"
)

print_button.pack(
    pady=30
)


# ============================================================
# RESULT LABEL
# ============================================================

result_label = tk.Label(
    left_frame,
    text="",
    font=("Arial", 12),
    justify="left"
)

result_label.pack(
    pady=10
)


# ============================================================
# START APPLICATION
# ============================================================

window.mainloop()

