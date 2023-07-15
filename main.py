import tkinter as tk
from tkinter import ttk

root = tk.Tk()

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")


frame = ttk.Frame(root) # parent widget
frame.pack()    # this function call makes the app resposive. Since this is
                # the root widget, adjusting the size of the UI/app will keep
                # the components centered

### col#1, row1 of the frame/root widget
widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
widgets_frame.grid(row=0, column=0, padx=20, pady=10) # padding on x & y axis


################## "Insert Row" LabelFrame ##################
### All ttk widgets here should have "widgets_frame" as their parent
### since they are supposed to be "grouped-together".

### col1, row1 ### Entry Field
ctrl_number = ttk.Entry(widgets_frame)
ctrl_number.insert(0, "Control #")    # Placeholder: insert str("Name") at index 0
ctrl_number.bind("<FocusIn>", lambda e: ctrl_number.delete('0', 'end')) # clears the text of the placeholder from index 0 to end
ctrl_number.grid(row=0, column=0,padx=5, pady=(0,5), sticky="ew") # "ew" means stretch from east-west

### col1, row2 ### Entry Field
dv_number = ttk.Entry(widgets_frame)
dv_number.insert(0, "DV #")
dv_number.bind("<FocusIn>", lambda e: dv_number.delete('0', 'end'))
dv_number.grid(row=1, column=0,padx=5, pady=(0,5), sticky="ew")

### col1, row3 ### Entry Field
check_date = ttk.Entry(widgets_frame)
check_date.insert(0, "Check Date")
check_date.bind("<FocusIn>", lambda e: check_date.delete('0', 'end'))
check_date.grid(row=2, column=0,padx=5, pady=(0,5), sticky="ew")

### col1, row4 ### Entry Field
check_number = ttk.Entry(widgets_frame)
check_number.insert(0, "Check #")    
check_number.bind("<FocusIn>", lambda e: check_number.delete('0', 'end')) 
check_number.grid(row=3, column=0,padx=5, pady=(0,5), sticky="ew")

### col1, row5 ### Entry Field
check_payee = ttk.Entry(widgets_frame)
check_payee.insert(0, "Check Amount")    
check_payee.bind("<FocusIn>", lambda e: check_payee.delete('0', 'end')) 
check_payee.grid(row=4, column=0,padx=5, pady=(0,5), sticky="ew")

### col1, row6 ### Entry Field
check_amount = ttk.Entry(widgets_frame)
check_amount.insert(0, "Check Amount")    
check_amount.bind("<FocusIn>", lambda e: check_amount.delete('0', 'end')) 
check_amount.grid(row=5, column=0,padx=5, pady=(0,5), sticky="ew")

### col1, row7 ### Dropdown
# set values for the options using a variable
combo_list = ["Paid", "Cancelled", "Stale"]
# set the widget
check_status = ttk.Combobox(widgets_frame, values=combo_list)
check_status.current(0) # default value selected from combo_list var
check_status.grid(row=6, column=0, padx=5, pady=(0,5), sticky="ew")


### col1, row5 ### Insert Row button
def insert_row():
    '''
     This function is the one where user input will be added to the excel file
     and displayed on the UI.
     It is composed of three parts:
    '''
    # retrieving data and assigning variables to them
    ctrlNumber = ctrl_number.get()
    dvNumber = dv_number.get()
    checkDate = check_date.get()
    checkNumber = check_number.get()
    checkPayee = check_payee.get()
    checkAmount = check_amount.get()
    checkStatus = check_status.get()
    # testing line for the above variable assignments
    print(ctrlNumber, dvNumber, checkDate, checkNumber,checkPayee, checkAmount, checkStatus)

    # inserting the row to the excel file
    # path = r"C:\Users\Administrator\Desktop\Github\xl_tkinter\people.xlsx"
    path = r"./mto_transmittal.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [ctrlNumber, dvNumber, checkDate, checkNumber, checkPayee, checkAmount, checkStatus]
    sheet.append(row_values)
    workbook.save(path)

    # displaying the inserted row on the UI (treeView)
    treeView.insert('', tk.END, values=row_values)

    # clear the values after inserting the new row
    # and then resetting the values to default
    ctrl_number.delete(0, "end")
    ctrl_number.insert(0, "Control #")
    dv_number.delete(0, "end")
    dv_number.insert(0, "DV #")
    check_date.delete(0, "end")
    check_date.insert(0, "Check Date")
    check_number.delete(0, "end")
    check_number.insert(0, "Check #")
    check_payee.delete(0, "end")
    check_payee.insert(0, "Payee")
    check_amount.delete(0, "end")
    check_amount.insert(0, "Amount")
    check_status.delete(0, "end")
    check_status.set(combo_list[0])



button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=7, column=0, sticky="nsew")

### separator ###
separator = ttk.Separator(widgets_frame)
separator.grid(row=8, column=0, padx=20, pady=10, sticky="ew")

### switch (dark/light)
def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style="Switch",
    command=toggle_mode) # this triggers the toggle_mode function above
mode_switch.grid(row=9, column=0, padx=5, pady=10, sticky="nsew")

################## /"Insert Row" LabelFrame ##################



################## TreeView / Excel LabelFrame ####################################
### This is where the preview of the excel file's data will be displayed

### Outer Frame
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Control #", "DV", "Check Date", "Check #", "Payee", "Amount", "Status") # column of the preview related to the excel file
treeView = ttk.Treeview(treeFrame, show="headings", 
                        yscrollcommand=treeScroll.set, columns=cols, height=15)
# these set the width of the columns specifically
treeView.column("Control #", width=50)
treeView.column("DV", width=50)
treeView.column("Check Date", width=80)
treeView.column("Check #", width=100)
treeView.column("Payee", width=300)
treeView.column("Amount", width=200)
treeView.column("Status", width=100)
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
    # path = r"C:\Users\Administrator\Desktop\Github\xl_tkinter\people.xlsx" # windows
    path = r"./mto_transmittal.xlsx" # linux
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    list_values = list(sheet.values)
    print(list_values) # to see the data inside the active sheet of the excel file
    for col_name in list_values[0]:
        # this loop gets the first "values" on the excel sheet (ie: headings of the columns)
        # those will then be set as the headings on the tkinter UI
        treeView.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        # starting from [1] onwards, are the data we need loaded into the treeView
        treeView.insert('', tk.END, values=value_tuple)

load_data()
################## /TreeView / Excel LabelFrame ####################################


root.mainloop()
