import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title('MTO Check Transmittal')

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

################## Main Frame ##################
'''
    outer_frame.pack() makes the app responsive.
    Since this is the main widget, adjusting the size of the UI/app will keep
    the components centered
'''
outer_frame = ttk.Frame(root) # parent widget
outer_frame.pack()


################## Widgets Frame ##################
''' 
    The user-input form area.
    All other ttk widgets should have "widgets_frame" as their parent
    since they are supposed to be "grouped-together"... except from, of course, the
    widgets_frame for its root will be the outer_frame (see above).
'''

### col0, row0 of the root_frame : Enclosure Widget for all other input widgets
widgets_frame = ttk.LabelFrame(outer_frame, text="Widgets Frame")
widgets_frame.grid(row=0, column=0, padx=20, pady=10) # padding on x & y axis


w00 = ttk.Label(widgets_frame, text="Check Date")
w00.grid(row=0, column=0)

w01 = ttk.Label(widgets_frame, text="Check #")
w01.grid(row=0, column=1)

w02 = ttk.Label(widgets_frame, text="DV #")
w02.grid(row=0, column=2)

w03 = ttk.Label(widgets_frame, text="Particulars")
w03.grid(row=0, column=3)

w04 = ttk.Label(widgets_frame, text="Amount")
w04.grid(row=0, column=4)

w05 = ttk.Label(widgets_frame, text="Status")
w05.grid(row=0, column=5)


check_date = ttk.Entry(widgets_frame, width=10)
# check_date.insert(0, "help text here")
# check_date.bind("<FocusIn>", lambda e: check_date.delete('0', 'end'))
check_date.grid(row=1, column=0,padx=5, pady=(0,5), sticky="ew")

check_number = ttk.Entry(widgets_frame,width=10)
check_number.grid(row=1, column=1,padx=5, pady=(0,5), sticky="ew")

dv_number = ttk.Entry(widgets_frame, width=10)
dv_number.grid(row=1, column=2,padx=5, pady=(0,5), sticky="ew")

check_particulars = ttk.Entry(widgets_frame, width=50)
check_particulars.grid(row=1, column=3,padx=5, pady=(0,5), sticky="ew")

check_amount = ttk.Entry(widgets_frame, width=20)
check_amount.grid(row=1, column=4,padx=5, pady=(0,5), sticky="ew")

status_list = ["Paid", "Cancelled", "Stale"]
check_status_dropdown = ttk.Combobox(widgets_frame, values=status_list)
check_status_dropdown.current(0) # default selected value
check_status_dropdown.grid(row=1, column=5,padx=5, pady=(0,5), sticky="ew")

### separator line###
separator = ttk.Separator(widgets_frame, )
separator.grid(row=2, columnspan=10, padx=20, pady=10, sticky="ew")

### switch (dark/light)
def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style="Switch",
    command=toggle_mode) # this triggers the toggle_mode function above
mode_switch.grid(row=3, column=0, padx=5, pady=10, sticky="nsew")

root.mainloop()
