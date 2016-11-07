from tkinter import *
import pandas as pd
from tkinter.filedialog import askopenfilename

root = Tk()

global first_file
first_file = ""
global second_file
second_file = ""

first_excel_name = StringVar()
second_excel_name = StringVar()

def browse_first_excel(var):
    Tk().withdraw()
    global first_file
    first_file = askopenfilename()
    var.set(first_file)

def browse_second_excel(var):
    Tk().withdraw()
    global second_file
    second_file = askopenfilename()
    var.set(second_file)

# Define the diff function to show the changes in each field
def report_diff(x):
    return x[0] if x[0] == x[1] else '{} ---> {}'.format(*x)

# We want to be able to easily tell which rows have changes
def has_change(row):
    if "--->" in row.to_string():
        return "Y"
    else:
        return "N"
def excel_differcence(file_name1,file_name2,col_file1,col_file2):
    # Read in both excel files
    df1 = pd.read_excel(file_name1, '{}'.format(col_file1), na_values=['NA'])
    df2 = pd.read_excel(file_name2, '{}'.format(col_file2), na_values=['NA'])

    # Make sure we order by account number so the comparisons work
    df1.sort_values(by="Stupac prvi")
    df1=df1.reindex()
    df2.sort_values(by="Stupac prvi")
    df2=df2.reindex()

    # Create a panel of the two dataframes
    diff_panel = pd.Panel(dict(df1=df1,df2=df2))

    #Apply the diff function
    diff_output = diff_panel.apply(report_diff, axis=0)

    # Flag all the changes
    diff_output['has_change'] = diff_output.apply(has_change, axis=1)

    #Save the changes to excel but only include the columns we care about
    diff_output[(diff_output.has_change == 'Y')].to_excel('my-diff-1.xlsx',index=False)

    Label(root, text="File Saved, you can quit now").grid(row=8, column=3)

def quit_app():
    quit()

excel_file_first_label = Label(root, text="File").grid(row=1, column=0)
excel_file_first_bar = Entry(root, textvariable = first_excel_name).grid(row=1, column=1)

excel_file_first_bbutton = Button(root, text="Browse", command=lambda: browse_first_excel(first_excel_name))
excel_file_first_bbutton.grid(row=1, column=3)

excel_file_second_label = Label(root, text="File").grid(row=2, column=0)
excel_file_second_bar = Entry(root, textvariable = second_excel_name).grid(row=2, column=1)

excel_file_second_bbutton = Button(root, text="Browse", command=lambda: browse_second_excel(second_excel_name))
excel_file_second_bbutton.grid(row=2, column=3)


compbutton = Button(root, text="Compare", command = lambda: excel_differcence(first_file,second_file,col_file1="Sheet1",col_file2="Sheet1"))
compbutton.grid(row=9, column=3)
cbutton = Button(root, text="Close", command = quit_app)
cbutton.grid(row=10, column=3)


root.mainloop()