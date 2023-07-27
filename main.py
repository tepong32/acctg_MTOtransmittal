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


def insert_row():
    '''
     This function is the one where user input will be added to the excel file
     and displayed on the UI.
     It is composed of three parts:
    '''
    # retrieving data and assigning variables to them
    checkDate = check_date.get()
    checkNumber = check_number.get()
    dvNumber = dv_number.get()
    checkParticulars = check_particulars.get()
    checkAmount = check_amount.get()
    checkStatus = check_status_dropdown.get()
    print(checkDate, checkNumber, dvNumber, checkParticulars, checkAmount, checkStatus)

    # inserting the row to the excel file
    path = r"C:\Users\Administrator\Desktop\Github\acctg_MTOtransmittal\data.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [checkDate, checkNumber, dvNumber, checkParticulars, checkAmount, checkStatus]
    sheet.append(row_values)
    workbook.save(path)

    # displaying the inserted row on the UI (treeView)
    treeView.insert('', tk.END, values=row_values)

    # clear the values after inserting the new row
    # and then resetting the values to default
    check_date.delete(0, "end")
    # name_entry.insert(0, "Name") # setting "help text" on entry widget if needed

    check_number.delete(0, "end")
    dv_number.delete(0, "end")
    check_particulars.delete(0, "end")
    check_amount.delete(0, "end")
    check_status_dropdown.set(status_list[0])
    # returns the focus to the check_date widget after inserting the new row
    check_date.focus_set()





# Defining button hover events
### OPTIONS ARE NOT WORKING ON TTK ###################################
def on_enter(e):
   button.config(background='OrangeRed3', foreground= "white")

def on_leave(e):
   button.config(background= 'SystemButtonFace', foreground= 'black')

# Button that calls the insert_row() above
button = ttk.Button(widgets_frame, text="Insert", command=insert_row, takefocus=1) #############
button.grid(row=4, column=3, sticky="nsew")

# Binding the hover events to the Button
button.bind('<Enter>', on_enter)
button.bind('<Leave>', on_leave)

### separator line (removed) ###############################################
# separator = ttk.Separator(widgets_frame, )
# separator.grid(row=2, columnspan=10, padx=20, pady=10, sticky="ew")


################## TreeView / Excel LabelFrame ####################################
### This is where the preview of the excel file's data will be displayed

### Display Frame
treeFrame = ttk.Frame(outer_frame, takefocus=0)
treeFrame.grid(row=5, column=0, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y") # this sets the scrollbar to the right side of the frame,
                                        # covering its whole height
cols = ("Check Date", "Check #", "DV #","Particulars", "Amount", "Status") # column of the preview related to the excel file
treeView = ttk.Treeview(treeFrame, show="headings", 
                        yscrollcommand=treeScroll.set, columns=cols, height=15)
# these set the width of the columns specifically
# COLUMN NAMES SHOULD MATCH THOSE INDICATED ON THE EXCEL FILE
treeView.column("Check Date", width=100)
treeView.column("Check #", width=100)
treeView.column("DV #", width=100)
treeView.column("Particulars", width=200)
treeView.column("Amount", width=100)
treeView.column("Status", width=80)
treeView.pack()
treeScroll.config(command=treeView.yview) # this line attaches the treeScroll widget to the treeView, scrolling vertically


# Event Listener function highlighting selected items on the treeView list
def selected():
    print(listbox.get(listbox.curselection()[0]))
    
treeView.bind("<<ListboxSelect>>", lambda x: selected())   


### attaching the excel file to the UI starts here:
import openpyxl

def load_data():
    # used prefix (r) to avoid unicodeescape error
    # see https://stackoverflow.com/questions/1347791/unicode-error-unicodeescape-codec-cant-decode-bytes-cannot-open-text-file
    path = r"C:\Users\Administrator\Desktop\Github\acctg_MTOtransmittal\data.xlsx" # windows
    # path = r"./data.xlsx" # linux
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    list_values = list(sheet.values)
    # print(list_values) # to see the data inside the active sheet of the excel file
    for col_name in list_values[0]:
        # this loop gets the first "values" on the excel sheet (ie: headings of the columns)
        # those will then be set as the headings on the tkinter UI
        treeView.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        # starting from [1] onwards, are the data (lists) we need loaded into the treeView
        treeView.insert('', tk.END, values=value_tuple)

load_data()
################## /TreeView / Excel LabelFrame ####################################

### switch (dark/light)
def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

mode_switch = ttk.Checkbutton(outer_frame, text="Mode", style="Switch",
    command=toggle_mode, takefocus=0) # this triggers the toggle_mode function above
mode_switch.grid(row=4, column=1, padx=5, pady=10, sticky="nsew")






# This sets the cursor to automatically be on the Check Date field by default whenever
# the program runs.
check_date.focus_set()

root.mainloop()
