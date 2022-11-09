"""--------------------------------------------------------------------------------------------------------------------
SNT_GUI:

Created by:     Cameron Jupp
Date Started:   Aug 24, 2022
--------------------------------------------------------------------------------------------------------------------"""

# -------------------------------------------------------------------------------------------------------------------- #
# -------------------------------------  /^\  / /  / __  /  /__ __/  / ___/ ------------------------------------------ #
# ------------------------------------  / /\\/ /  / /_/ /    / /    / __/  ------------------------------------------- #
# -----------------------------------  /_/  \_/  /_____/    /_/    /____/ -------------------------------------------- #
# -------------------------------------------------------------------------------------------------------------------- #
# Last left off at:
#
# - Add serial number generate button and display
#
# - Fix mismatch for generate_sn and add_flyway with hex/decimal values
#
# -------------------------------------------------------------------------------------------------------------------- #
import sys

print(sys.path)
import tkinter as tk
import tkinter.ttk as ttk
import SNT_Module as snt

# -------------------------------------------------------------------------------------------------------------------- #
# ----------------------  / ___/  / / //  /^\  / /  / ___/ /__  __/  /_  _/  / __  /  /^\  / / ----------------------- #
# ---------------------  / __/   / /_//  / /\\/ /  / /__     / /      / /   / /_/ /  / /\\/ / ------------------------ #
# --------------------  /_/     /____/  /_/  \_/  /____/    /_/    /____/  /_____/  /_/  \_/ ------------------------- #
# -------------------------------------------------------------------------------------------------------------------- #

"""------------------------------------------------------------------------------------
gui_init: 
                -----------------------------------------------
Arguments:
 -
                -----------------------------------------------
Returns: nothing
                -----------------------------------------------
Created by: Cameron Jupp
Date:       Sep 14, 2022
Edited:
------------------------------------------------------------------------------------"""
def gui_init():
    return snt.init_sheets()

"""------------------------------------------------------------------------------------
generate_sn:
                -----------------------------------------------
Arguments:
 - 
                -----------------------------------------------
Returns: nothing
                -----------------------------------------------
Created by: Cameron Jupp
Date:       , 2022
Edited:
------------------------------------------------------------------------------------"""
def generate_sn(worksheet, style, component, display, feedback_text):
    feedback_text.configure(text = "Generating S/N...", bg = primary_bg)
    new_sn = hex(snt.get_highest_sn(worksheet, snt.fw_sn_col, style, component) + 1).upper(); new_sn = new_sn[2:]

    display.configure(text = new_sn)
    return 0

"""------------------------------------------------------------------------------------
gui_add_flyway: 
                -----------------------------------------------
Arguments:
 -
                -----------------------------------------------
Returns: nothing
                -----------------------------------------------
Created by: Cameron Jupp
Date:       Sep 14, 2022
Edited:
------------------------------------------------------------------------------------"""
def gui_add_flyway(sheet, body_id_entry, amp_sn_entry, cont_sn_entry, style, feedback_text):
    body_id = body_id_entry.get()
    amp_sn = amp_sn_entry.get()
    cont_sn = cont_sn_entry.get()

    feedback_text.configure(text = "Checking spreadsheet...", bg = primary_bg)

    return_var = snt.add_flyway(sheet, body_id, amp_sn, cont_sn, style)
    if return_var == "":
        feedback_text.configure(text = "Failed to add flyway: Body ID already exists!", bg = "light red")
    else:
        feedback_text.configure(text="Flyway successfully added!", bg = "light green")

"""------------------------------------------------------------------------------------
clear_flyway_entry: 
                -----------------------------------------------
Arguments:
 -
                -----------------------------------------------
Returns: nothing
                -----------------------------------------------
Created by: Cameron Jupp
Date:       Sep 14, 2022
Edited:
------------------------------------------------------------------------------------"""
def clear_fw_entry(body_id_entry, amp_sn_entry, cont_sn_entry):
    body_id_entry.delete(0, tk.END)
    amp_sn_entry.delete(0, tk.END)
    cont_sn_entry.delete(0, tk.END)
    return 0

"""------------------------------------------------------------------------------------
gui_add_controller: 
                -----------------------------------------------
Arguments:
 -
                -----------------------------------------------
Returns: nothing
                -----------------------------------------------
Created by: Cameron Jupp
Date:       Sep 14, 2022
Edited:
------------------------------------------------------------------------------------"""

"""------------------------------------------------------------------------------------
gui_add_body: 
                -----------------------------------------------
Arguments:
 -
                -----------------------------------------------
Returns: nothing
                -----------------------------------------------
Created by: Cameron Jupp
Date:       Sep 14, 2022
Edited:
------------------------------------------------------------------------------------"""

"""------------------------------------------------------------------------------------
gui_add_sensor: 
                -----------------------------------------------
Arguments:
 -
                -----------------------------------------------
Returns: nothing
                -----------------------------------------------
Created by: Cameron Jupp
Date:       Sep 14, 2022
Edited:
------------------------------------------------------------------------------------"""



#--------------------------------------------------------------------------------------------------------------------- #
# -------------------------------------  /^\/^\      /^^\    |_ _|    /^\  / /  -------------------------------------- #
# ------------------------------------  / /\/\ \    / /_\\    | |    / /\\/ /  --------------------------------------- #
# -----------------------------------  /_/    \_\  /_/   \\  |___|  /_/  \_/  ---------------------------------------- #
#----------------------------------------------------------------------------------------------------------------------#
primary_bg =    "light gray"
secondary_bg =  "gray"
entry_bg =      "white"

universal_font = "Calibri"
entry_font_size = 18
title_font_size = 20
radiobutton_font_size = 18
button_font_size = 20
tab_font_size = 20

worksheets_dict = gui_init()






Main = tk.Tk()

Main.title("Serial Number Tool")
Main.geometry("1200x600")

main_notebook = ttk.Notebook(Main)
main_notebook.pack(expand = "yes", fill = "both")
#----------------------------------------------------------------------------------------------------------------------
#------------------------------------------------ Notebook Tabs -------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

ttk.Style().configure('TNotebook.Tab',
                      font=(universal_font, tab_font_size),
                      background="dark grey",
                      foreground='black',
                      borderwidth=1)
ttk.Style().configure('TNotebook',
                      font=(universal_font, tab_font_size),
                      background='black',
                      foreground='black',
                      borderwidth=0)

flyway_assembly_tab = ttk.Frame(main_notebook)
flyway_assembly_tab.pack(expand="yes", fill="both")
main_notebook.add(flyway_assembly_tab, text='Flyway Assembly')

controller_programming_tab = ttk.Frame(main_notebook)
controller_programming_tab.pack(expand="yes", fill="both")
main_notebook.add(controller_programming_tab, text='Controller Programming')

mover_tab = ttk.Frame(main_notebook)
mover_tab.pack(expand="yes", fill="both")
main_notebook.add(mover_tab, text='Mover Assembly')

#----------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------- Flyway Tab --------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

# Border Frame
fw_deco_frame = tk.LabelFrame(flyway_assembly_tab, bg = primary_bg)
fw_deco_frame.pack(expand = "yes", fill = "both", padx = 10, pady = 10)

#----------------------------------------------------------------------------------------------------------------------
body_id_var =   tk.StringVar()
amp_sn_var =    tk.StringVar()
cont_sn_var =   tk.StringVar()

body_id_frame = tk.LabelFrame(fw_deco_frame,
                              text="Body ID",
                              font = (universal_font, title_font_size),
                              bg = primary_bg)
body_id_frame.grid(row = 1, rowspan = 1, column = 1, columnspan = 1, padx = 30, pady = 30, sticky = tk.NSEW)

body_id_entry = tk.Entry(body_id_frame,
                         font = (universal_font,
                                 entry_font_size),
                         bg = entry_bg,
                         textvariable = body_id_var)
body_id_entry.pack(side = tk.LEFT, expand = "yes", fill = "both")
#------------------------------
amp_sn_frame = tk.LabelFrame(fw_deco_frame,
                             text="Amp S/N",
                             font = (universal_font,
                                     title_font_size),
                             bg = primary_bg)
amp_sn_frame.grid(row = 1, rowspan = 1, column = 2, columnspan = 2, padx = 30, pady = 30, sticky = tk.NSEW)

amp_sn_entry = tk.Entry(amp_sn_frame,
                        font = (universal_font, entry_font_size),
                        bg = entry_bg,
                        textvariable = amp_sn_var)
amp_sn_entry.pack(side = tk.LEFT, expand = "yes", fill = "both")
#------------------------------
controller_sn_frame = tk.LabelFrame(fw_deco_frame,
                                    text="Controller S/N",
                                    font = (universal_font, title_font_size),
                                    bg = primary_bg)
controller_sn_frame.grid(row = 1, rowspan = 1, column = 4, columnspan = 1, padx = 30, pady = 30, sticky = tk.NSEW)

controller_sn_entry = tk.Entry(controller_sn_frame,
                               font = (universal_font,
                                       entry_font_size),
                               bg = entry_bg,
                               textvariable = cont_sn_var)
controller_sn_entry.pack(side = tk.LEFT, expand = "yes", fill = "both")
#----------------------------------------------------------------------------------------------------------------------
customer_var =  tk.StringVar()
customer_var.initialize("1")

choice_frame = tk.LabelFrame(fw_deco_frame, text="Style", font = (universal_font, title_font_size), bg = primary_bg)
choice_frame.grid(row = 2, rowspan = 3, column = 4, columnspan = 1, padx = 30, pady = 30, sticky = tk.NSEW)


bandr_choice = tk.Radiobutton(choice_frame,
                              value = "1",
                              text="B & R",
                              variable = customer_var,
                              font = (universal_font, radiobutton_font_size),
                              bg = primary_bg)
bandr_choice.grid(row = 2, rowspan = 1, column = 1, columnspan = 1, padx = 20, pady = 0, sticky = tk.W)

other_choice = tk.Radiobutton(choice_frame,
                              value = "0",
                              text="PMI",
                              variable = customer_var,
                              font = (universal_font, radiobutton_font_size),
                              bg = primary_bg)
other_choice.grid(row = 3, rowspan = 1, column = 1, columnspan = 1, padx = 20, pady = 0, sticky = tk.W)

#----------------------------------------------------------------------------------------------------------------------

sn_display_frame = tk.LabelFrame(fw_deco_frame,
                                 text = "S/N",
                                 font = (universal_font, entry_font_size),
                                 bg = primary_bg)
sn_display_frame.grid(row = 3, rowspan = 1, column = 2, columnspan = 2, padx = 30, pady = 30, sticky = tk.NSEW)

gen_sn_display = tk.Label(sn_display_frame,
                          text = "",
                          font = (universal_font, entry_font_size),
                          bg = entry_bg)
gen_sn_display.pack(expand = "yes", fill = "both")




generate_sn_button = tk.Button(fw_deco_frame,
                              text = "Generate S/N",
                              font = (universal_font, entry_font_size),
                              bg = "orange",
                              height = 1,
                              command = lambda:
                              generate_sn(worksheets_dict["Production Log"], customer_var.get(), "flyway", gen_sn_display, fw_feedback_text))
generate_sn_button.grid(row = 3, rowspan = 1, column = 1, columnspan = 1, padx = 30, pady = 30, sticky = tk.NSEW)

#----------------------------------------------------------------------------------------------------------------------

fw_submit_button = tk.Button(fw_deco_frame,
                             text="Submit",
                             font = (universal_font, button_font_size),
                             bg = "blue",
                             command=lambda:
                             gui_add_flyway(worksheets_dict["Production Log"], body_id_entry, amp_sn_entry, controller_sn_entry, customer_var.get(), fw_feedback_text))
fw_submit_button.grid(row = 6, rowspan = 1, column = 1, columnspan = 1, padx = 70, pady = 30, sticky = tk.NSEW)

fw_clear_button = tk.Button(fw_deco_frame,
                            text="Clear",
                            font = (universal_font, button_font_size),
                            bg = "dark gray",
                            command=lambda:
                            clear_fw_entry(body_id_entry, amp_sn_entry, controller_sn_entry))
fw_clear_button.grid(row = 6, rowspan = 1, column = 4, columnspan = 1, padx = 70, pady = 30, sticky = tk.NSEW)

fw_feedback_text = tk.Label(fw_deco_frame,
                            text="Awaiting input...",
                            font = (universal_font, entry_font_size),
                            bg = "light grey")
fw_feedback_text.grid(row = 5, rowspan = 1, column = 1, columnspan = 4, padx = 20, pady = 30, sticky = tk.NSEW)

#----------------------------------------------------------------------------------------------------------------------
#----------------------------------------------- Controller Tab -------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

Main.mainloop()