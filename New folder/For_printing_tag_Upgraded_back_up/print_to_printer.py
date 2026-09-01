import tkinter as tk
from tkinter import filedialog

from PIL import Image, ImageTk, ImageWin

import pypdfium2 as pdfium

import win32print
import win32ui

from win32con import (
    HORZRES,
    VERTRES,
    PHYSICALWIDTH,
    PHYSICALHEIGHT,
    PHYSICALOFFSETX,
    PHYSICALOFFSETY
)


class PDFViewer:

    def __init__(self, root):

        self.root = root

        self.root.title("PDF Viewer")

        self.root.geometry("1000x700")


        # ==================================================
        # VARIABLES
        # ==================================================

        self.pdf = None

        self.page_number = 0

        # Current PIL image
        self.image = None

        # Current Tkinter image
        self.photo = None

        # Selected printer
        self.selected_printer = (
            win32print.GetDefaultPrinter()
        )

        # Printer paper size
        self.paper_width = 1000
        self.paper_height = 1400

        # Printable area
        self.printable_width = 0
        self.printable_height = 0

        # Physical offset
        self.offset_x = 0
        self.offset_y = 0


        # ==================================================
        # GET PRINTER PAPER SIZE
        # ==================================================

        self.update_printer_information(
            self.selected_printer
        )


        # ==================================================
        # TOP FRAME
        # ==================================================

        top_frame = tk.Frame(
            root
        )

        top_frame.pack(
            fill="x"
        )


        # ==================================================
        # OPEN PDF
        # ==================================================

        # tk.Button(
        #     top_frame,
        #     text="Open PDF",
        #     command=self.open_pdf
        # ).pack(
        #     side="left",
        #     padx=5,
        #     pady=5
        # )


        # ==================================================
        # PREVIOUS
        # ==================================================

        tk.Button(
            top_frame,
            text="Previous",
            command=self.previous_page
        ).pack(
            side="left",
            padx=5
        )


        # ==================================================
        # NEXT
        # ==================================================

        tk.Button(
            top_frame,
            text="Next",
            command=self.next_page
        ).pack(
            side="left",
            padx=5
        )


        # ==================================================
        # SELECT PRINTER
        # ==================================================

        tk.Button(
            top_frame,
            text="Select Printer",
            command=self.choose_printer
        ).pack(
            side="left",
            padx=10
        )


        # ==================================================
        # COPIES LABEL
        # ==================================================

        tk.Label(
            top_frame,
            text="Copies:"
        ).pack(
            side="left",
            padx=(10, 2)
        )


        # ==================================================
        # COPIES ENTRY
        # ==================================================

        self.copies_entry = tk.Entry(
            top_frame,
            width=5
        )

        self.copies_entry.pack(
            side="left"
        )

        self.copies_entry.insert(
            0,
            "1"
        )


        # ==================================================
        # PRINT BUTTON
        # ==================================================

        tk.Button(
            top_frame,
            text="Print",
            command=self.print_page
        ).pack(
            side="left",
            padx=10
        )


        # ==================================================
        # PAGE LABEL
        # ==================================================

        self.page_label = tk.Label(
            top_frame,
            text="No PDF"
        )

        self.page_label.pack(
            side="left",
            padx=10
        )


        # ==================================================
        # PRINTER LABEL
        # ==================================================

        self.printer_label = tk.Label(
            top_frame,
            text=f"Printer: {self.selected_printer}"
        )

        self.printer_label.pack(
            side="left",
            padx=10
        )


        # ==================================================
        # DISPLAY FRAME
        # ==================================================

        display_frame = tk.Frame(
            root
        )

        display_frame.pack(
            fill="both",
            expand=True
        )


        # ==================================================
        # LEFT EMPTY AREA
        # ==================================================

        left_space = tk.Frame(
            display_frame
        )

        left_space.pack(
            side="left",
            fill="both",
            expand=True
        )


        # ==================================================
        # RIGHT PAPER FRAME
        # ==================================================

        paper_frame = tk.Frame(
            display_frame
        )

        paper_frame.pack(
            side="right",
            fill="both",
            expand=True
        )


        # ==================================================
        # CANVAS
        # ==================================================

        self.canvas = tk.Canvas(
            paper_frame,
            bg="gray"
        )

        self.canvas.pack(
            side="left",
            fill="both",
            expand=True
        )


        # ==================================================
        # VERTICAL SCROLLBAR
        # ==================================================

        scrollbar = tk.Scrollbar(
            paper_frame,
            orient="vertical",
            command=self.canvas.yview
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )


        self.canvas.configure(
            yscrollcommand=scrollbar.set
        )


        # ==================================================
        # HORIZONTAL SCROLLBAR
        # ==================================================

        horizontal_scrollbar = tk.Scrollbar(
            paper_frame,
            orient="horizontal",
            command=self.canvas.xview
        )

        horizontal_scrollbar.pack(
            side="bottom",
            fill="x"
        )


        self.canvas.configure(
            xscrollcommand=horizontal_scrollbar.set
        )


        # ==================================================
        # CANVAS RESIZE
        # ==================================================

        self.canvas.bind(
            "<Configure>",
            self.canvas_resized
        )




    # ==================================================
    # UPDATE PRINTER INFORMATION
    # ==================================================

    def update_printer_information(
        self,
        printer_name
    ):

        try:

            printer = win32ui.CreateDC()

            printer.CreatePrinterDC(
                printer_name
            )


            # ==========================================
            # PHYSICAL PAPER SIZE
            # ==========================================

            self.paper_width = (
                printer.GetDeviceCaps(
                    PHYSICALWIDTH
                )
            )

            self.paper_height = (
                printer.GetDeviceCaps(
                    PHYSICALHEIGHT
                )
            )


            # ==========================================
            # PRINTABLE AREA
            # ==========================================

            self.printable_width = (
                printer.GetDeviceCaps(
                    HORZRES
                )
            )

            self.printable_height = (
                printer.GetDeviceCaps(
                    VERTRES
                )
            )


            # ==========================================
            # PHYSICAL OFFSET
            # ==========================================

            self.offset_x = (
                printer.GetDeviceCaps(
                    PHYSICALOFFSETX
                )
            )

            self.offset_y = (
                printer.GetDeviceCaps(
                    PHYSICALOFFSETY
                )
            )


            printer.DeleteDC()


            print(
                "================================"
            )

            print(
                "Printer:",
                printer_name
            )

            print(
                "Paper width:",
                self.paper_width
            )

            print(
                "Paper height:",
                self.paper_height
            )

            print(
                "Printable width:",
                self.printable_width
            )

            print(
                "Printable height:",
                self.printable_height
            )

            print(
                "Offset X:",
                self.offset_x
            )

            print(
                "Offset Y:",
                self.offset_y
            )

            print(
                "================================"
            )


        except Exception as e:

            print(
                "Printer information error:",
                e
            )


    # ==================================================
    # OPEN PDF
    # ==================================================

    def open_pdf(self):

        # file_path = filedialog.askopenfilename(
        #     filetypes=[
        #         ("PDF Files", "*.pdf")
        #     ]
        # )

        file_path ="Output_files/1201_D30019(PG0722)_KIRAN_tag.pdf"


        print(
            file_path
        )


        if not file_path:

            return


        try:

            self.pdf = pdfium.PdfDocument(
                file_path
            )

            self.page_number = 0

            self.show_page()


        except Exception as e:

            print(
                "Error:",
                e
            )


    # ==================================================
    # SHOW PAGE
    # ==================================================

    def show_page(self):

        if self.pdf is None:

            return


        # ==============================================
        # GET PDF PAGE
        # ==============================================

        page = self.pdf[
            self.page_number
        ]


        # ==============================================
        # RENDER PDF
        # ==============================================

        bitmap = page.render(
            scale=1
        )


        # ==============================================
        # CONVERT TO PIL
        # ==============================================

        self.image = bitmap.to_pil()


        # ==============================================
        # SHOW PAPER
        # ==============================================

        self.show_paper()


        # ==============================================
        # PAGE NUMBER
        # ==============================================

        self.page_label.config(
            text=(
                f"Page "
                f"{self.page_number + 1}"
                f" / "
                f"{len(self.pdf)}"
            )
        )


    # ==================================================
    # SHOW PAPER
    # ==================================================

    def show_paper(self):

        if self.image is None:

            return


        # ==================================================
        # CANVAS SIZE
        # ==================================================

        canvas_width = (
            self.canvas.winfo_width()
        )

        canvas_height = (
            self.canvas.winfo_height()
        )


        if canvas_width <= 1:

            canvas_width = 500


        if canvas_height <= 1:

            canvas_height = 600


        # ==================================================
        # CREATE PAPER
        # ==================================================

        paper = Image.new(
            "RGB",
            (
                self.paper_width,
                self.paper_height
            ),
            "white"
        )


        # ==================================================
        # PDF SIZE
        # ==================================================

        pdf_width, pdf_height = (
            self.image.size
        )


        # ==================================================
        # MARGIN
        # ==================================================

        margin = 20


        available_width = (
            self.paper_width
            -
            (2 * margin)
        )


        available_height = (
            self.paper_height
            -
            (2 * margin)
        )


        # ==================================================
        # SCALE PDF
        # ==================================================

        scale_x = (
            available_width
            /
            pdf_width
        )


        scale_y = (
            available_height
            /
            pdf_height
        )


        scale = min(
            scale_x,
            scale_y
        )


        # ==================================================
        # NEW PDF SIZE
        # ==================================================

        new_width = int(
            pdf_width * scale
        )


        new_height = int(
            pdf_height * scale
        )


        # ==================================================
        # RESIZE PDF
        # ==================================================

        pdf_image = self.image.resize(
            (
                new_width,
                new_height
            ),
            Image.Resampling.LANCZOS
        )


        # ==================================================
        # CENTER PDF ON PAPER
        # ==================================================

        x = int(
            (
                self.paper_width
                -
                new_width
            )
            /
            2
        )


        y = int(
            (
                self.paper_height
                -
                new_height
            )
            /
            2
        )


        # ==================================================
        # PUT PDF ON PAPER
        # ==================================================

        paper.paste(
            pdf_image,
            (
                x,
                y
            )
        )


        # ==================================================
        # SCALE PAPER FOR CANVAS
        # ==================================================

        scale_x = (
            canvas_width
            /
            self.paper_width
        )


        scale_y = (
            canvas_height
            /
            self.paper_height
        )


        display_scale = min(
            scale_x,
            scale_y,
            1
        )


        display_width = int(
            self.paper_width
            *
            display_scale
        )


        display_height = int(
            self.paper_height
            *
            display_scale
        )


        # ==================================================
        # RESIZE PAPER
        # ==================================================

        display_paper = paper.resize(
            (
                display_width,
                display_height
            ),
            Image.Resampling.LANCZOS
        )


        # ==================================================
        # CONVERT TO TKINTER IMAGE
        # ==================================================

        self.photo = ImageTk.PhotoImage(
            display_paper
        )


        # ==================================================
        # CLEAR CANVAS
        # ==================================================

        self.canvas.delete(
            "all"
        )


        # ==================================================
        # CENTER PAPER
        # ==================================================

        canvas_x = int(
            (
                canvas_width
                -
                display_width
            )
            /
            2
        )


        canvas_y = int(
            (
                canvas_height
                -
                display_height
            )
            /
            2
        )


        if canvas_x < 0:

            canvas_x = 0


        if canvas_y < 0:

            canvas_y = 0


        # ==================================================
        # DISPLAY PAPER
        # ==================================================

        self.canvas.create_image(
            canvas_x,
            canvas_y,
            image=self.photo,
            anchor="nw"
        )


        # ==================================================
        # SCROLL REGION
        # ==================================================

        self.canvas.configure(
            scrollregion=(
                0,
                0,
                max(
                    canvas_width,
                    display_width
                ),
                max(
                    canvas_height,
                    display_height
                )
            )
        )


    # ==================================================
    # CANVAS RESIZED
    # ==================================================

    def canvas_resized(
        self,
        event=None
    ):

        if self.image is not None:

            self.show_paper()


    # ==================================================
    # NEXT PAGE
    # ==================================================

    def next_page(self):

        if self.pdf is None:

            return


        if (
            self.page_number
            <
            len(self.pdf) - 1
        ):

            self.page_number += 1

            self.show_page()


    # ==================================================
    # PREVIOUS PAGE
    # ==================================================

    def previous_page(self):

        if self.pdf is None:

            return


        if self.page_number > 0:

            self.page_number -= 1

            self.show_page()


    # ==================================================
    # CHOOSE PRINTER
    # ==================================================

    def choose_printer(self):

        try:

            # ==========================================
            # GET PRINTERS
            # ==========================================

            printers = win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL
                |
                win32print.PRINTER_ENUM_CONNECTIONS
            )


            printer_names = [
                printer[2]
                for printer in printers
            ]


            if not printer_names:

                print(
                    "No printers found."
                )

                return


            # ==========================================
            # PRINTER WINDOW
            # ==========================================

            printer_window = tk.Toplevel(
                self.root
            )


            printer_window.title(
                "Select Printer"
            )


            printer_window.geometry(
                "500x180"
            )


            printer_window.resizable(
                False,
                False
            )


            # ==========================================
            # LABEL
            # ==========================================

            tk.Label(
                printer_window,
                text="Select Printer:",
                font=(
                    "Arial",
                    11
                )
            ).pack(
                pady=(15, 5)
            )


            # ==========================================
            # VARIABLE
            # ==========================================

            selected_printer = (
                tk.StringVar()
            )


            # ==========================================
            # DROPDOWN
            # ==========================================

            printer_dropdown = tk.OptionMenu(
                printer_window,
                selected_printer,
                *printer_names
            )


            printer_dropdown.config(
                width=45
            )


            printer_dropdown.pack(
                padx=20
            )


            # ==========================================
            # DEFAULT PRINTER
            # ==========================================

            default_printer = (
                win32print.GetDefaultPrinter()
            )


            if default_printer in printer_names:

                selected_printer.set(
                    default_printer
                )

            else:

                selected_printer.set(
                    printer_names[0]
                )


            # ==========================================
            # CONFIRM
            # ==========================================

            def confirm_printer():

                self.selected_printer = (
                    selected_printer.get()
                )


                # --------------------------------------
                # GET NEW PAPER SIZE
                # --------------------------------------

                self.update_printer_information(
                    self.selected_printer
                )


                # --------------------------------------
                # UPDATE LABEL
                # --------------------------------------

                self.printer_label.config(
                    text=(
                        f"Printer: "
                        f"{self.selected_printer}"
                    )
                )


                # --------------------------------------
                # REFRESH PAPER PREVIEW
                # --------------------------------------

                if self.image is not None:

                    self.show_paper()


                print(
                    "Selected printer:",
                    self.selected_printer
                )


                printer_window.destroy()


            # ==========================================
            # SELECT BUTTON
            # ==========================================

            tk.Button(
                printer_window,
                text="Select",
                width=12,
                command=confirm_printer
            ).pack(
                pady=15
            )


            # ==========================================
            # MODAL
            # ==========================================

            printer_window.transient(
                self.root
            )


            printer_window.grab_set()


            self.root.wait_window(
                printer_window
            )


        except Exception as e:

            print(
                "Printer selection error:",
                e
            )


    # ==================================================
    # PRINT CURRENT PAGE
    # ==================================================

    def print_page(self):

        if self.image is None:

            print(
                "No page to print."
            )

            return


        # ==================================================
        # GET COPIES
        # ==================================================

        try:

            copies = int(
                self.copies_entry.get()
            )


            if copies <= 0:

                print(
                    "Copies must be greater than 0."
                )

                return


        except ValueError:

            print(
                "Please enter a valid number of copies."
            )

            return


        try:

            # ==========================================
            # SELECTED PRINTER
            # ==========================================

            printer_name = (
                self.selected_printer
            )


            print(
                "Printer:",
                printer_name
            )


            print(
                "Copies:",
                copies
            )


            # ==========================================
            # CREATE PRINTER DC
            # ==========================================

            printer = win32ui.CreateDC()


            printer.CreatePrinterDC(
                printer_name
            )


            # ==========================================
            # GET ACTUAL PAPER SIZE
            # ==========================================

            paper_width = (
                printer.GetDeviceCaps(
                    PHYSICALWIDTH
                )
            )


            paper_height = (
                printer.GetDeviceCaps(
                    PHYSICALHEIGHT
                )
            )


            # ==========================================
            # GET PRINTABLE AREA
            # ==========================================

            printable_width = (
                printer.GetDeviceCaps(
                    HORZRES
                )
            )


            printable_height = (
                printer.GetDeviceCaps(
                    VERTRES
                )
            )


            # ==========================================
            # GET OFFSET
            # ==========================================

            offset_x = (
                printer.GetDeviceCaps(
                    PHYSICALOFFSETX
                )
            )


            offset_y = (
                printer.GetDeviceCaps(
                    PHYSICALOFFSETY
                )
            )


            print(
                "Paper:",
                paper_width,
                "x",
                paper_height
            )


            print(
                "Printable:",
                printable_width,
                "x",
                printable_height
            )


            # ==========================================
            # PREPARE IMAGE
            # ==========================================

            image = self.image.copy()


            image_width, image_height = (
                image.size
            )


            # ==========================================
            # FIT PDF TO PRINTABLE AREA
            # ==========================================

            scale_x = (
                printable_width
                /
                image_width
            )


            scale_y = (
                printable_height
                /
                image_height
            )


            scale = min(
                scale_x,
                scale_y
            )


            new_width = int(
                image_width * scale
            )


            new_height = int(
                image_height * scale
            )


            # ==========================================
            # RESIZE
            # ==========================================

            image = image.resize(
                (
                    new_width,
                    new_height
                ),
                Image.Resampling.LANCZOS
            )


            # ==========================================
            # CENTER IN PRINTABLE AREA
            # ==========================================

            x = (
                offset_x
                +
                int(
                    (
                        printable_width
                        -
                        new_width
                    )
                    /
                    2
                )
            )


            y = (
                offset_y
                +
                int(
                    (
                        printable_height
                        -
                        new_height
                    )
                    /
                    2
                )
            )


            # ==========================================
            # CREATE DIB
            # ==========================================

            dib = ImageWin.Dib(
                image
            )


            # ==========================================
            # PRINT COPIES
            # ==========================================

            for copy_number in range(
                copies
            ):

                print(
                    f"Printing copy "
                    f"{copy_number + 1}"
                    f" of "
                    f"{copies}"
                )


                # --------------------------------------
                # START DOCUMENT
                # --------------------------------------

                printer.StartDoc(
                    "PDF Page"
                )


                printer.StartPage()


                # --------------------------------------
                # DRAW IMAGE
                # --------------------------------------

                dib.draw(
                    printer.GetHandleOutput(),
                    (
                        x,
                        y,
                        x + new_width,
                        y + new_height
                    )
                )


                # --------------------------------------
                # END PAGE
                # --------------------------------------

                printer.EndPage()


                # --------------------------------------
                # END DOCUMENT
                # --------------------------------------

                printer.EndDoc()


            # ==========================================
            # RELEASE PRINTER
            # ==========================================

            printer.DeleteDC()


            print(
                "Printed successfully."
            )


        except Exception as e:

            print(
                "Printing error:",
                e
            )


# ======================================================
# MAIN
# ======================================================

root = tk.Tk()

app = PDFViewer(
    root
)
app.open_pdf()

root.mainloop()