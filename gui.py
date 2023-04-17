### Import the required libraries
from tkinter import *
from tkinter import filedialog
import pandas as pd
from functions import *

### set frames and base grid
# create instance
root = Tk()
# set base dimensions and features
root.geometry("475x400")
root.title('Bill_Bot v1')

### set dimensions for frames
task_frame = LabelFrame(root, text="Tasks", padx=2, pady=2)
task_frame.grid(rowspan=8, row=0, column=0)

menu_frame = LabelFrame(root, padx=5, pady=5)
menu_frame.grid(rowspan=1, row=1, column=1)

### title
label = Label(root, text="Welcome!")
label.grid(row=0, column=1)

### select files
# function to select files
def select_docs():
    menu_frame.filename = filedialog.askopenfilename(\
        initialdir="/",\
        title='Select A Folder')
# button to intiate function to select files
select_button = Button(\
    menu_frame,\
    text="Select Documents",\
    padx=50,\
    command=select_docs)\
    .grid(row=1,column=10)

### Create dropdown Menus
listbox_daily = Listbox(task_frame, width=40, height=20, selectmode=MULTIPLE)
# Inserting the listbox items
listbox_daily.insert(1, "d-Posted/Unposted")
listbox_daily.insert(2, "d-Focus File")
listbox_daily.insert(3, "d-Overdue Invoices")
listbox_daily.insert(4, "d-All Daily")


### Function to process selected requests
tasks = []
def submit_requests():
    # Traverse the tuple returned by
    # curselection method and print
    # corresponding value(s) in the listbox
    for i in listbox_daily.curselection():
        tasks.append(listbox_daily.get(i))
    label = Label(root, text="Your requests are being processed...")
    label.grid(row=0, column=1)
    all_tasks(tasks)



### Feedback input
# dimension
feedback = Entry(menu_frame, width=30)
# location
feedback.grid(row=6, column=10)
# function to submit feedback
def submit_feedback():
    label = Label(root, text="Your feedback will be reviewed soon...")
    label.grid(row=0, column=1)
# button to send feedback
feedback_button = Button(\
    menu_frame,\
    text="Send Feedback",\
    padx=50,\
    command=submit_feedback)\
    .grid(row=5,column=10)
# button to process reports
submit_button = Button(\
    menu_frame,\
    text="Prepare Reports",\
    padx=50,\
    command=submit_requests)\
    .grid(row=4,column=10)

listbox_daily.pack()
### needed for app
root.mainloop()
