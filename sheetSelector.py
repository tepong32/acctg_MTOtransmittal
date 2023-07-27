import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import openpyxl

def load_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)
        load_data(file_path)

def load_data(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    list_values = list(sheet.values)
    for col_name in list_values[0]:
        treeView.heading(col_name, text=col_name)

    treeView.delete(*treeView.get_children())
    for value_tuple in list_values[1:]:
        treeView.insert('', tk.END, values=value_tuple)

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

# Create the main window
root = tk.Tk()
root.title('MTO Check Transmittal')

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

# Outer Frame
outer_frame = ttk.Frame(root)
outer_frame.pack()

# Widgets Frame
widgets_frame = ttk.LabelFrame(outer_frame, text="Widgets Frame")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

name_label = ttk.Label(widgets_frame, text="Name")  # Add this line
name_label.grid(row=2, column=0)  # Add this line

name_entry = ttk.Entry(widgets_frame, width=20)  # Add this line
name_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=(0, 5), sticky="ew")  # Add this line

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

# Select File Button
btn_load = ttk.Button(widgets_frame, text="Load Excel File", command=load_excel_file)
btn_load.grid(row=0, column=6, padx=5, pady=(0, 5), sticky="ew")

# Select Sheet Button (Inside Widgets Frame)
def select_sheet():
    selected_item = treeView.focus()
    if selected_item:
        item_values = treeView.item(selected_item, "values")
        selected_sheet_name = item_values[0]  # Assuming the first value is the "Check Date" field
        print(f"Selected sheet: {selected_sheet_name}")

btn_select_sheet = ttk.Button(widgets_frame, text="Select Sheet", command=select_sheet, state=tk.DISABLED)
btn_select_sheet.grid(row=1, column=6, padx=5, pady=(0, 5), sticky="ew")

# TreeView / Excel LabelFrame
treeFrame = ttk.Frame(outer_frame)
treeFrame.grid(row=5, column=0, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Check Date", "Check #", "DV #", "Particulars", "Amount", "Status")
treeView = ttk.Treeview(treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=15)
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


# Insert button (Inside Widgets Frame)
button = ttk.Button(widgets_frame, text="Insert", command=insert_row, takefocus=1)
button.grid(row=4, column=3, sticky="nsew")

# Switch (Dark/Light)
def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

mode_switch = ttk.Checkbutton(outer_frame, text="Mode", style="Switch", command=toggle_mode, takefocus=0)
mode_switch.grid(row=4, column=1, padx=5, pady=10, sticky="nsew")

# Set initial focus to the Check Date field
check_date.focus_set()

root.mainloop()
