Running sequence:
1. Only run "main_print_Upgraded_with_printing_FInal".




DB files:
1. "DB_files/combined_dry_weight_DB.xlsx" :->connected to "main_print_Upgraded_with_printing_FInal"
2. "DB_files/Party_logo_name_DB.xlsx"  :->connected to "main_print_Upgraded_with_printing_FInal"
3. "DB_files/shade_library_DB.xlsx"  :->connected to "main_print_Upgraded_with_printing_FInal"
4. "DB_files/Spacing_position_DB.xlsx"  :->connected to "print_tag_2layout_func"


File purposes:
1.MAIN for GUI data Entry:-><"main_print_Upgraded_with_printing_FInal" (Activated, All components are activated)> or
	<"main_print_Upgraded_with_textbox_and_list" (NOT Activated, with text box and list, NOT printing to Printer) : searching is smooth > or
	<"main_print" (NOT Activated, with dropdown box, NOT printing to Printer) : searching is NOT smooth >

1. To make PDF: -> "print_tag_func" (NOT In use, to make 2 layout, can be connected to
	 <"main_print_Upgraded_with_printing_FInal" or "main_print" or "main_print_Upgraded_with_textbox_and_list"> )

	and "print_tag_2layout_func" (In use, to make 2 layout, can be connected to
	 <"main_print_Upgraded_with_printing_FInal" or "main_print" or "main_print_Upgraded_with_textbox_and_list" >)
2. To print to printer: -> "print_to_printer_func" (In use,  can be connected to
	 <"main_print_Upgraded_with_printing_FInal">)
