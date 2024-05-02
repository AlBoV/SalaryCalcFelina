import pandas
import tkinter
import tkcalendar
import datetime

from tkinter import messagebox
from tkinter import *

from tkinter import filedialog
from tkinter import ttk

def last_day_of_month(any_day):
    # The day 28 exists in every month. 4 days later, it's always next month
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
    # subtracting the number of the current day brings us back one month
    return next_month - datetime.timedelta(days=next_month.day)
def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?", parent=user_data_frame):
        for widget in user_data_frame.winfo_children():
            try:
                t = widget.get()
            except:
                continue
            cell = cells.get(widget)
            if cell is not None:
                arrayxlsx[cell[1]][cell[2]]=t;
        df = pandas.DataFrame(arrayxlsx, columns=tablexls.columns)
        df.to_excel('data/nexkell.xlsx', index=False)
        root.destroy()
def click_button():
    global worktable
    global workingtime
    global Categories
    global EmployeesWithCategory
    global workingmonth

    # Reading the company's employee list file
    try:
        headers = ['TAJCode', 'Name', 'StartDate', 'EndDate', 'UnitWork', 'QuitDate', 'UnitName', 'UnitCode']
        worktable = pandas.read_csv('data/dolgadatnap.csv',
                                    header=None,
                                    names=headers,
                                    dtype={0: 'str', 1: 'str', 2: 'str', 3: 'str', 4: 'str', 5: 'str', 6: 'str', 7: 'str'},
                                    sep=';',
                                    encoding='latin-1')
        worktable['UnitWork'] = worktable['UnitWork'].fillna(worktable['UnitName'])
        worktable.loc[:, 'UnitName'] = 'x'
        worktable['TAJCode'] = worktable['TAJCode'].astype(str)
    except:
        messagebox.showerror(title="Error",
                         message="File not found 'data/dolgadatnap.csv'!\nPlace a file with this name in the specified directory.")
        return

    # Reading the employee list file with categories
    try:
        headersCOE = ['ADO', 'TAJCode', 'Name', 'Category']
        EmployeesWithCategory = pandas.read_csv('data/nbesorolas.csv',
                                    header=None,
                                    names=headersCOE,
                                    dtype={0: 'str', 1: 'str', 2: 'str', 3: 'str'},
                                    sep=';',
                                    encoding='latin-1')
        EmployeesWithCategory['TAJCode'] = EmployeesWithCategory['TAJCode'].astype(str)
    except:
        messagebox.showerror(title="Error",
                         message="File not found 'data/nbesorolas.csv'!\nPlace a file with this name in the specified directory.")
        return

    # Reading categories file
    try:
        Categories = pandas.read_excel('data/kategoriak.xlsx')
    except:
        messagebox.showerror(title="Error",
                             message="File not found 'data/kategoriak.xlsx'!\nPlace a file with this name in the specified directory.")
        return

    # Getting the list of completed work from the LOGIN database
    filename_csv = filedialog.askopenfilename(title="Please select a working time file", defaultextension="csv")
    if filename_csv != "":
        # read the working time data file
        headersWT = ['TAJCode', 'Name', 'Unit', 'SiteName', 'Datum', 'WorkHours', 'OtherHours', 'OverHours', 'AbsenceHours', 'AbsenceType', 'NormaMinutes', 'ChangeBy']
        workingtime = pandas.read_csv(filename_csv,
                                      header=0,
                                      names=headersWT,
                                      dtype={0: 'str', 1: 'str', 2: 'str', 3: 'str', 4: 'str', 5: 'str', 6: 'str', 7: 'str', 8: 'str', 9: 'str', 10: 'str', 11: 'str'},
                                      sep=';',
                                      encoding='latin-1')
        # set the DATUM column to the Date type
        DatumCol = workingtime.columns[4]
        workingtime[DatumCol] = pandas.to_datetime(workingtime[DatumCol])
        workingtime['TAJCode'] = workingtime['TAJCode'].astype(str)

        # Get the date from the form and determine the start date of the month and end date of the month
        currentdate = cal.get_date()
        firstdayofmonth = currentdate.replace(day=1)
        lastdayofmonth = last_day_of_month(currentdate)

        workingmonth = workingtime.loc[(workingtime[DatumCol].dt.date >= firstdayofmonth) & (workingtime[DatumCol].dt.date <= lastdayofmonth)]

        if workingmonth.empty:
            messagebox.showerror(title="Error",
                                 message="The selected file does not contain data for the specified period, "
                                         "select another file or select another date.")
            return
    else:
        messagebox.showerror(title="Error",
                             message="You have not selected a file!\nSelect the file again.")
        return

    # Processing the file



# Main dict initialisation
cells = {}
worktable = pandas.DataFrame()              # LLOGIN FILE
workingtime = pandas.DataFrame()            # VARRSOR FILE
Categories = pandas.DataFrame()             # KATEGORIAK FILE
EmployeesWithCategory = pandas.DataFrame()  # nbesorolas file
workingmonth = pandas.DataFrame()           # LLOGIN FILE FILTERED

root = tkinter.Tk()
root.title('From Login to Nexon crating a file import of bonuses')
root.geometry('600x400')

# Interface creation
mainframe = Frame(root)
mainframe.pack()

# Uploading and saving user information about pathes
main_data_frame = LabelFrame(mainframe, text="Main data")
main_data_frame.grid(sticky="NEWS", padx=10, pady=10)

lable_1 = Label(main_data_frame, text="Enter month of salary calculation ")
lable_1.grid(row=0, column=1, sticky="W")

cal = tkcalendar.DateEntry(main_data_frame, width=12, borderwidth=2, date_pattern='MM/dd/yyyy')
cal.grid(row=0, column=3, sticky="E")


for widget in main_data_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

user_data_frame = LabelFrame(mainframe, text="User entered data")
user_data_frame.grid(sticky="NEWS", padx=10, pady=10)

# Reading the settings file for application of additional bonuses by subdivisions
try:
    tablexls = pandas.read_excel('data/nexkell.xlsx')
except:
    messagebox.showerror(title="Error",
                         message="File not found 'data/nexkell.xlsx'!\nPlace a file with this name in the specified directory.")
    root.destroy()

arrayxlsx = tablexls.to_numpy()

total_columns = tablexls.columns.__len__()
total_rows = tablexls.__len__()

j=0
for header in tablexls.columns:
        l = tkinter.Label(user_data_frame, text=header.upper(), relief=tkinter.FLAT, font=('Arial', 10, 'bold'))
        l.grid(row=0, column=j, sticky="NSEW")
        j=j+1

for i in range(total_rows):
    for j in range(total_columns):
        if j==0:
            e = Entry(user_data_frame, width=20,
                            font=('Arial', 10, 'bold'))
            e.config(state='normal')
        else:
            e = Entry(user_data_frame, width=10,
                      font=('Arial', 10))
        id = f'{i+1}{j}'
        e.grid(row=i+1, column=j)
        e.insert(END, arrayxlsx[i][j])

        cells[e]=[arrayxlsx[i][j], i, j]

# Button for reading all auxiliary data along the set paths
btnRead = Button(mainframe, text="Read all data", command=click_button)
btnRead.grid(sticky="NEWS", padx=10, pady=10)

# Button to create a file for import into NEXON
def click_button_export():
    worktable['TAJCode'] = worktable['TAJCode'].astype(str)
    workingmonth['TAJCode'] = workingmonth['TAJCode'].astype(str)
    EmployeesWithCategory['TAJCode'] = EmployeesWithCategory['TAJCode'].astype(str)

    worktableresult = pandas.merge(worktable, workingmonth, how='inner', on='TAJCode', suffixes=('', '_work'))
    worktableresult = pandas.merge(worktableresult, EmployeesWithCategory, how='inner', on='TAJCode', suffixes=('', '_emp'))

btnExport = Button(mainframe, text="NEXON end of month import LOGIN", command=click_button_export)
btnExport.grid(sticky="NEWS", padx=10, pady=10)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()