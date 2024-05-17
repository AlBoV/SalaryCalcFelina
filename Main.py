import pandas
import tkinter
import tkcalendar
import datetime
import numpy as np
import os

from xhtml2pdf import pisa

from tkinter import messagebox
from tkinter import *

from tkinter import filedialog


# Service procedure to get end of month by date
def last_day_of_month(any_day):
    # The day 28 exists in every month. 4 days later, it's always next month
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
    # subtracting the number of the current day brings us back one month
    return next_month - datetime.timedelta(days=next_month.day)


# Procedure when the window close button is pressed
def on_closing():
    # Opening a warning about closing the window, if the answer is yes, then execute, otherwise return to the main window without action
    if messagebox.askokcancel("Quit", "Do you want to quit?", parent=user_data_frame):
        # Loop through all subordinate elements of the main program window
        for widget in user_data_frame.winfo_children():
            # Try to get the data of subitems, if it is an element of the table of values on accruals
            try:
                t = widget.get()
            except:
                continue
            # Then read the value from the element
            cell = cells.get(widget)
            # If the value is filled, write it to the array cell by index
            if cell is not None:
                arrayxlsx[cell[1]][cell[2]] = t
        # Write an array of values to the dataframe to save it to a file
        df = pandas.DataFrame(arrayxlsx, columns=tablexls.columns)
        # Save the array data to a file
        df.to_excel('data/nexkell.xlsx', index=False)
        # Closing the main window
        root.destroy()


# Procedure for processing a press of the button to read data
def click_button_read_all_data():
    # Pass global variables to the procedure
    global worktable
    global workingtime
    global Categories
    global EmployeesWithCategory
    global workingmonth

    # Reading the company's employee list from file
    try:
        headers = ['TAJCode', 'Name', 'StartDate', 'EndDate', 'UnitWork', 'QuitDate', 'UnitName', 'UnitCode']
        worktable = pandas.read_csv('data/dolgadatnap.csv',
                                    header=None,
                                    names=headers,
                                    dtype={0: 'str', 1: 'str', 2: 'str', 3: 'str', 4: 'str', 5: 'str', 6: 'str', 7: 'str'},
                                    sep=';',
                                    encoding='latin_1')
        worktable['UnitWork'] = worktable['UnitWork'].fillna(worktable['UnitName'])
        worktable.loc[:, 'UnitName'] = 'x'
        worktable['TAJCode'] = worktable['TAJCode'].astype(str)

        worktable['StartDate'] = pandas.to_datetime(worktable['StartDate'])
        worktable['EndDate'] = pandas.to_datetime(worktable['EndDate'])
        worktable['QuitDate'] = pandas.to_datetime(worktable['QuitDate'])

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
                                    encoding='latin_1')
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
    filename_csv = filedialog.askopenfilename(title="Please select a working time file",
                                              defaultextension="csv",
                                              filetypes=[("Csv files", ".csv")])
    if filename_csv != "":
        # read the working time data file
        headersWT = ['TAJCode', 'Name', 'Unit', 'SiteName', 'Datum', 'WorkHours', 'OtherHours', 'OverHours', 'AbsenceHours', 'AbsenceType', 'NormaMinutes', 'ChangeBy']
        workingtime = pandas.read_csv(filename_csv,
                                      header=0,
                                      names=headersWT,
                                      dtype={0: 'str', 1: 'str', 2: 'str', 3: 'str', 4: 'str', 5: 'str', 6: 'str', 7: 'str', 8: 'str', 9: 'str', 10: 'str', 11: 'str'},
                                      sep=';',
                                      encoding='latin_1')
        # set the DATUM column to the Date type
        DatumCol = workingtime.columns[4]
        workingtime[DatumCol] = pandas.to_datetime(workingtime[DatumCol])
        workingtime['TAJCode'] = workingtime['TAJCode'].astype(str)

        # Get the date from the form and determine the start date of the month and end date of the month
        currentdate = cal.get_date()
        firstdayofmonth = currentdate.replace(day=1)
        lastdayofmonth = last_day_of_month(currentdate)

        # Filter the working timetable by dates within the selected month
        workingmonth = workingtime.loc[(workingtime[DatumCol].dt.date >= firstdayofmonth) & (workingtime[DatumCol].dt.date <= lastdayofmonth)]

        # If nothing is found, we inform you that the selected file does not contain the searched data
        if workingmonth.empty:
            messagebox.showerror(title="Error",
                                 message="The selected file does not contain data for the specified period, "
                                         "select another file or select another date.")
            return
    else:
        messagebox.showerror(title="Error",
                             message="You have not selected a file!\nSelect the file again.")
        return

# Main dict initialisation
cells = {}
worktable = pandas.DataFrame()              # LLOGIN FILE -
workingtime = pandas.DataFrame()            # VARRSOR FILE
Categories = pandas.DataFrame()             # KATEGORIAK FILE
EmployeesWithCategory = pandas.DataFrame()  # nbesorolas file
workingmonth = pandas.DataFrame()           # LLOGIN FILE FILTERED
workingtableforreports = pandas.DataFrame()
workingtableforreportssum = pandas.DataFrame()

# Create the main program window
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
tablexls = pandas.DataFrame()

try:
    tablexls = pandas.read_excel('data/nexkell.xlsx')
except:
    messagebox.showerror(title="Error",
                         message="File not found 'data/nexkell.xlsx'!\nPlace a file with this name in the specified directory.")
    root.destroy()

# Create an array of values based on the read data of the table
if not tablexls.empty:
    arrayxlsx = tablexls.to_numpy()
    total_columns = tablexls.columns.__len__()
    total_rows = tablexls.__len__()

    # Draw a table on the main form
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
            e.insert(END, str(arrayxlsx[i][j]))

            cells[e] = [arrayxlsx[i][j], i, j]

# Button for reading all auxiliary data along the set paths
btnRead = Button(mainframe, text="Read all data", command=click_button_read_all_data)
btnRead.grid(sticky="NEWS", padx=10, pady=10)


# Button to create a file for import into NEXON
def click_button_export():
    global workingtableforreportssum
    global workingtableforreports

    # Read the selected date from the form
    currentdate = cal.get_date()
    curentdatestr = currentdate.strftime("%Y%m")
    firstdayofmonth = currentdate.replace(day=1)
    lastdayofmonth = last_day_of_month(currentdate)


    # Read user-defined accrual data
    UnitFactorUsage = pandas.DataFrame(arrayxlsx, columns=['UnitWork', 'useage', 'mp', 'mpfel', 'kap', 'kapfel'])

    # Translate the TAJCode column to the string value for further use in connections
    worktable['TAJCode'] = worktable['TAJCode'].astype(str)

    # Filter the main worksheet by dates within the selected month
    filtredworktable = worktable[(worktable.StartDate.dt.date <= firstdayofmonth) & (((worktable.EndDate.dt.date >= firstdayofmonth) & (worktable.EndDate.dt.date <= lastdayofmonth)) | (worktable.EndDate.isnull())) & ((worktable.QuitDate.dt.date >= firstdayofmonth) | (worktable.QuitDate.isnull()))]

    # Create a copy of the main table for the working month in order to modify it, but the main table remains the same
    workingmonthcopy = workingmonth.copy()
    # Translate the TAJCode column to the string value for further use in connections
    workingmonthcopy['TAJCode'] = workingmonthcopy['TAJCode'].astype(str)
    # Translate the WorkHours,OtherHours,OverHours,AbsenceHours  column to the string value for further useage
    workingmonthcopy['WorkHours'] = workingmonthcopy['WorkHours'].str.replace(',', '.')
    workingmonthcopy['OtherHours'] = workingmonthcopy['OtherHours'].str.replace(',', '.')
    workingmonthcopy['OverHours'] = workingmonthcopy['OverHours'].str.replace(',', '.')
    workingmonthcopy['AbsenceHours'] = workingmonthcopy['AbsenceHours'].str.replace(',', '.')
    workingmonthcopy['AbsenceType'] = workingmonthcopy['AbsenceType'].fillna("")
    workingmonthcopy['AbsenceDays'] = 0

    # Convert column values to a float
    workingmonthcopy[['WorkHours', 'OtherHours', 'OverHours', 'AbsenceHours', 'NormaMinutes']] = workingmonthcopy[['WorkHours', 'OtherHours', 'OverHours', 'AbsenceHours', 'NormaMinutes']].astype(float)
    workingmonthcopy[['WorkHours', 'OtherHours', 'OverHours', 'AbsenceHours', 'NormaMinutes']] = workingmonthcopy[['WorkHours', 'OtherHours', 'OverHours', 'AbsenceHours', 'NormaMinutes']].fillna(0)
    # Preparation of an error table on the reflection of working time
    worktablewithmistakes = workingmonthcopy[((workingmonthcopy['Unit'].str.contains('default', case=False)) | (workingmonthcopy['SiteName'].str.contains('default', case=False)))]

    # IT1. Group the table by major columns and summarize the hours, in the first iteration NormaMinutes should remain on the left.
    # That way we get rid of their repeats.
    workingmonthcopy.loc[((workingmonthcopy.AbsenceHours>0)
                          & ((workingmonthcopy['Unit'].str.contains('default', case=False))
                             |(workingmonthcopy['SiteName'].str.contains('default', case=False)))), 'AbsenceHours'] = 0

    summaryworkingmonth = workingmonthcopy.groupby(by=['TAJCode', 'Name', 'Datum', 'NormaMinutes', 'AbsenceType'], as_index=False).agg({'WorkHours': 'sum', 'OtherHours': 'sum', 'OverHours': 'sum', 'AbsenceHours': 'sum'})
    summaryworkingmonth = summaryworkingmonth.sort_values(['Name', 'Datum'])
    # Populate the WorkHours column with 8 hours if there are hours in the OverHours column
    # summaryworkingmonth.loc[((summaryworkingmonth.WorkHours==0) & (summaryworkingmonth.OverHours>0)), 'WorkHours'] = 8
    # summaryworkingmonth.loc[((summaryworkingmonth.AbsenceType.str.contains('szabadság', case=False))), 'WorkHours'] = 8
    summaryworkingmonth.loc[((summaryworkingmonth.AbsenceType.str.contains('Táppénz', case=False))), 'AbsenceDays'] = 1
    summaryworkingmonth.loc[((summaryworkingmonth.AbsenceType.str.contains('nem fizetett', case=False))), 'AbsenceDays'] = 1
    summaryworkingmonth['AbsenceDays'] = summaryworkingmonth['AbsenceDays'].fillna(0)
    summaryworkingmonth.loc[(((summaryworkingmonth.AbsenceHours>0) & (summaryworkingmonth.OverHours>0))), 'AbsenceHours'] = 0

    workingtableforreports = summaryworkingmonth.copy()

    # Filter the table by those rows where there are filled hours, we don't need other rows
    # summaryworkingmonth = summaryworkingmonth[(summaryworkingmonth.WorkHours > 0) | (summaryworkingmonth.OtherHours > 0) | (summaryworkingmonth.OverHours > 0) | (summaryworkingmonth.AbsenceHours > 0) | (summaryworkingmonth.AbsenceType.isnull() != True)]
    # IT2. Group the table by major columns and summarize the hours, in the second iteration NormaMinutes goes on the right.
    summaryworkingmonth = summaryworkingmonth.groupby(by=['Name', 'TAJCode'], as_index=False).agg({'NormaMinutes': 'sum', 'WorkHours': 'sum', 'OtherHours': 'sum', 'OverHours': 'sum', 'AbsenceHours': 'sum', 'AbsenceDays': 'sum'})

    # Translate the TAJCode column to the string value for further use in connections
    EmployeesWithCategory['TAJCode'] = EmployeesWithCategory['TAJCode'].astype(str)

    # Merge the tables of working time and the main table of work by employees by TAJCode column
    worktableresult = pandas.merge(filtredworktable, summaryworkingmonth, how='inner', on='TAJCode', suffixes=('', '_work'))
    # Merge the tables of working result and employees with categories by TAJCode column
    worktableresult = pandas.merge(worktableresult, EmployeesWithCategory, how='inner', on='TAJCode', suffixes=('', '_emp'))
    # IT3. Group the table by major columns and summarize the hours, in the second iteration NormaMinutes goes on the right.
    worktableresult = worktableresult.groupby(by=['Name', 'UnitWork', 'TAJCode', 'Category'], as_index=False).agg({'NormaMinutes': 'sum', 'WorkHours': 'sum', 'OtherHours': 'sum', 'OverHours': 'sum', 'AbsenceHours': 'sum', 'AbsenceDays': 'sum'})

    # Calculate the working hours in the WorkHours column using the formula
    worktableresult['WorkHours'] = worktableresult['WorkHours'] - worktableresult['OtherHours']  + worktableresult['OverHours'] - worktableresult['AbsenceHours']
    # ?alculate the value of efficiency coefficient by the formula (the formula is correct)
    worktableresult['AvgEfficiencyFactor'] = round(worktableresult['NormaMinutes']/(worktableresult['WorkHours'])/10, 2)
    worktableresult['AvgEfficiencyFactor'] = worktableresult['AvgEfficiencyFactor'].fillna(0)

    # Merge the tables of working result and user-defined usage parameters by UnitWork column
    worktableresult = pandas.merge(worktableresult, UnitFactorUsage, how='inner', on='UnitWork', suffixes=('', '_param'))
    # Calculate the value of overtime accrual by time of absence
    worktableresult['AccrueOverhead'] = np.where(worktableresult['AbsenceDays'] > 2, 0, 1)
    # Add new columns for calculations
    worktableresult['MP'] = 0
    worktableresult['JP'] = 0
    worktableresult['KAP'] = 0

    # Convert the columns of the Categories table to the string type, for use in a string search
    Categories['kat'] = Categories['kat'].astype(str)
    Categories['ervhotol'] = Categories['ervhotol'].astype(str)
    Categories['ervhoig'] = Categories['ervhoig'].astype(str)
    worktableresult['Category'] = worktableresult['Category'].astype(str)

    # For each row of the main result table we fill the columns according to the following conditions
    for row in worktableresult.iterrows():
        LineForMP = Categories.loc[((Categories['kat'] == row[1]['Category'])
                               & (Categories['pot'] == 'MP')
                               & (Categories['ervhotol'] <= curentdatestr)
                               & (Categories['ervhoig'] >= curentdatestr)
                               & (Categories['szazmin'] <= row[1]['AvgEfficiencyFactor'])
                               & (Categories['szazmax'] > row[1]['AvgEfficiencyFactor'])
                               )]
        try:
            LineForMPSum = LineForMP.iloc[0]['osszeg']
        except:
            LineForMPSum = 0
        # column MP
        if ((row[1]['mp'] != 0)
                & (row[1]['AccrueOverhead'] == 1)):
            worktableresult.at[row[0], 'MP'] = LineForMPSum
        if ((row[1]['mpfel'] != 0)
                & (row[1]['AccrueOverhead'] == 1)):
            worktableresult.at[row[0], 'MP'] = LineForMPSum/2

        LineForJP = Categories.loc[((Categories['kat'] == row[1]['Category'])
                               & (Categories['pot'] == 'JP')
                               & (Categories['ervhotol'] <= curentdatestr)
                               & (Categories['ervhoig'] >= curentdatestr)
                               & (Categories['szazmin'] <= row[1]['AvgEfficiencyFactor'])
                               & (Categories['szazmax'] > row[1]['AvgEfficiencyFactor'])
                               )]
        try:
            LineForJPSum = LineForJP.iloc[0]['osszeg']
        except:
            LineForJPSum = 0
        # column JP
        if row[1]['AccrueOverhead'] == 1:
            worktableresult.at[row[0], 'JP'] = LineForJPSum

        LineForKAP = Categories.loc[((Categories['kat'] == row[1]['Category'])
                                    & (Categories['pot'] == 'KAP')
                                    & (Categories['ervhotol'] <= curentdatestr)
                                    & (Categories['ervhoig'] >= curentdatestr)
                                    & (Categories['szazmin'] <= row[1]['AvgEfficiencyFactor'])
                                    & (Categories['szazmax'] > row[1]['AvgEfficiencyFactor'])
                                    )]
        try:
            LineForKAPSum = LineForKAP.iloc[0]['osszeg']
        except:
            LineForKAPSum = 0
        # column KAP
        if ((row[1]['kap'] != 0)
                & (row[1]['AccrueOverhead'] == 1)
                & (row[1]['AvgEfficiencyFactor'] >= float(80))):
            worktableresult.at[row[0], 'KAP'] = LineForKAPSum
        if ((row[1]['kapfel'] != 0)
                & (row[1]['AccrueOverhead'] == 1)
                & (row[1]['AvgEfficiencyFactor'] >= float(80))):
            worktableresult.at[row[0], 'KAP'] = LineForKAPSum / 2

    workingtableforreportssum = worktableresult.copy()

    # Write the obtained result of the main result table into an Excel file for checking
    try:
        worktableresult = worktableresult.sort_values(['Name', 'UnitWork', 'Category'])
        worktableresult.to_excel('data/monthly_supplements_'+curentdatestr+'.xlsx',
                             index=False,
                             columns=['Name', 'UnitWork', 'TAJCode', 'Category', 'NormaMinutes', 'WorkHours',
                                      'OtherHours', 'OverHours', 'AbsenceHours', 'AbsenceDays', 'AvgEfficiencyFactor', 'MP', 'JP',
                                      'KAP', 'AccrueOverhead'])

        messagebox.showinfo(title="Success",
                            message="Files was been successfully saved in 'data/monthly_supplements_"+curentdatestr+"'.xlsx'.")
    except:
        messagebox.showerror(title="Error",
                             message="File can't be saved in 'data/monthly_supplements_"+curentdatestr+"'.xlsx'!.\nThe file is probably already in use.")

    # Write the obtained result of the main result table into an Excel file where mistakes appear
    if worktablewithmistakes.__len__()>0:
        try:
            worktablewithmistakes = worktablewithmistakes.sort_values(['Name', 'Unit', 'SiteName'])
            worktablewithmistakes.to_excel('data/Mistakes_'+curentdatestr+'.xlsx',
                                 index=False,
                                 columns=['Name', 'Unit', 'SiteName', 'TAJCode', 'Datum', 'NormaMinutes', 'WorkHours',
                                          'OtherHours', 'OverHours', 'AbsenceHours'])

            messagebox.showwarning(title="Warning of incorrect data",
                                    message="Files was been successfully saved in 'data/Mistakes_"+curentdatestr+"'.xlsx'!")
        except:
            messagebox.showerror(title="Error",
                                 message="File can't be saved in 'data/Mistakes_"+curentdatestr+"'.xlsx'!.\nThe file is probably already in use.")

    # Write the obtained result of the main result table into an CSV file for importing in NEXON
    # try:
    #     worktableresultforNBKifiz = worktableresult[['TAJCode', 'MP', 'JP', 'KAP']].copy()
    #     worktableresultforNBKifiz = worktableresultforNBKifiz.melt(id_vars=['TAJCode'], var_name='Code',
    #                                                                value_name='Sum')
    #     worktableresultforNBKifiz['Active'] = 0
    #     worktableresultforNBKifiz['Percentage'] = 0
    #     worktableresultforNBKifiz['Time'] = 0
    #     worktableresultforNBKifiz['StartFrom'] = firstdayofmonth.strftime("%Y.%m.%d.")
    #     worktableresultforNBKifiz['StartTill'] = lastdayofmonth.strftime("%Y.%m.%d.")
    #
    #     worktableresultforNBKifiz.to_csv('//10.3.1.1/bér/import/NBkifiz.csv',
    #                                      index=False,
    #                                      header=False,
    #                                      sep=';',
    #                                      columns=['TAJCode', 'Active', 'Code', 'Sum', 'Percentage',
    #                                               'Time', 'StartFrom', 'StartTill'],
    #                                      date_format='%Y.%m.%d.',
    #                                      decimal=',',
    #                                      float_format='%.2f')
    #     messagebox.showinfo(title="Success",
    #                         message="Files was been successfully saved in '//10.3.1.1/bér/import/NBkifiz.csv'!")
    # except:
    #     messagebox.showerror(title="Error",
    #                          message="File can't be saved in '//10.3.1.1/bér/import/NBkifiz.csv'!.\nThe file is probably already in use or no access to this catalog.")


def convert_html_to_pdf(html_string, pdf_path):
    with open(pdf_path, "wb") as pdf_file:
        pisa_status = pisa.CreatePDF(html_string, dest=pdf_file)

    return not pisa_status.err


def filling_page_header(pagedata, textdata: str):

    textdata = textdata.replace('{{Name}}', pagedata[1]['Name'])
    textdata = textdata.replace('{{TAJCode}}', pagedata[1]['TAJCode'])
    textdata = textdata.replace('{{Unit}}', pagedata[1]['UnitWork'])
    return textdata


def filling_page_footer(pagedata, textdata: str, currentmonth: str):

    textdata = textdata.replace('{{TotalMonth}}', currentmonth)
    textdata = textdata.replace('{{TotalHours}}', str(pagedata[1]['WorkHours']+pagedata[1]['OtherHours']+pagedata[1]['OverHours']))
    textdata = textdata.replace('{{TotalOtherHours}}', str(pagedata[1]['OtherHours']))
    textdata = textdata.replace('{{TotalNormaMinutes}}', str(pagedata[1]['NormaMinutes']))
    textdata = textdata.replace('{{TotalPerfHours}}', str(pagedata[1]['WorkHours']))
    textdata = textdata.replace('{{TotalProcent}}', str(pagedata[1]['AvgEfficiencyFactor']))

    textdata = textdata.replace('{{Category}}', pagedata[1]['Category'])
    CategorySum = 0
    if pagedata[1]['Category'] == 'A':
        CategorySum = 266800
    if pagedata[1]['Category'] == 'B':
        CategorySum = 281175
    if pagedata[1]['Category'] == 'C':
        CategorySum = 295550

    textdata = textdata.replace('{{CategorySum}}', str(CategorySum))
    textdata = textdata.replace('{{AbsenceDays}}', str(pagedata[1]['AbsenceDays']))
    textdata = textdata.replace('{{AbsenceSum}}', str(pagedata[1]['JP']))

    UnitWork = pagedata[1]['UnitWork']
    if pagedata[1]['AccrueOverhead'] == 0:
        UnitWork = 'sok a hiányzás.'

    textdata = textdata.replace('{{Unit}}', UnitWork)
    textdata = textdata.replace('{{UnitSum}}', str(pagedata[1]['MP']))
    textdata = textdata.replace('{{ProductivitySum}}', str(pagedata[1]['KAP']))

    return textdata


def filling_line(linedata, textdata: str):

    textdata = textdata.replace('{{Date}}', linedata[1]['Datum'].strftime("%Y.%m.%d."))

    if linedata[1]['AbsenceType'] != '':
        textdata = textdata.replace('{{Hours}}', linedata[1]['AbsenceType'])
        textdata = textdata.replace('{{PerfHours}}', '')
        Procent = ''
    else:
        textdata = textdata.replace('{{Hours}}', str('8'))
        textdata = textdata.replace('{{PerfHours}}', str(linedata[1]['WorkHours']-linedata[1]['OtherHours']))
        if linedata[1]['WorkHours'] - linedata[1]['OtherHours']!=0:
            Procent = round(linedata[1]['NormaMinutes'] / (linedata[1]['WorkHours'] - linedata[1]['OtherHours']) / 10, 2)
        else:
            Procent = ''

    textdata = textdata.replace('{{OtherHours}}', str(linedata[1]['OtherHours']))
    textdata = textdata.replace('{{NormaMinutes}}', str(linedata[1]['NormaMinutes']))
    textdata = textdata.replace('{{Procent}}', str(Procent))

    return textdata


# Button to Create a button for employee report
def click_button_reports():
    global workingtableforreportssum
    global workingtableforreports

    currentdate = cal.get_date()
    currentmonth = currentdate.strftime("%Y.%m")

    # Specify HTML string
    html = open('data/template.html', 'r', encoding='latin_1').read()

    html = html.replace('{{PicturePath}}', os.path.abspath(os.curdir)+"\\data\\SalaryScale.bmp")

    FirstSplit = html.rsplit(sep='<!--page-->')

    NewHTML = ""

    Header = FirstSplit[0]
    PageText = FirstSplit[1]
    Footer = FirstSplit[2]

    SecondSplit = PageText.rsplit(sep='<!--tableline-->')

    PageHeader = SecondSplit[0]
    TableLine = SecondSplit[1]
    PageFooter = SecondSplit[2]

    NewHTML = Header
    
    for page in workingtableforreportssum.iterrows():

        NewHTML = NewHTML+filling_page_header(page, PageHeader)
        PageText = ''

        for line in workingtableforreports[workingtableforreports['TAJCode'] == page[1]['TAJCode']].iterrows():

            PageText = PageText + filling_line(line, TableLine)

        NewHTML = NewHTML + PageText + filling_page_footer(page, PageFooter, currentmonth)

    NewHTML = NewHTML + Footer

    with open('data/HtmlTable.html', 'w', encoding='Latin_1') as f:
         f.write(NewHTML)

    os.startfile(os.path.abspath(os.curdir)+"\\data\\HtmlTable.html")

    #convert_html_to_pdf(NewHTML, 'data/HtmlTable.pdf')


# Create a button for uploading data to NEXON
btnExport = Button(mainframe, text="NEXON end of month import LOGIN", command=click_button_export)
btnExport.grid(sticky="NEWS", padx=10, pady=10)

# Create a button for employee report
btnExport = Button(mainframe, text="Salary scales sheets for employee", command=click_button_reports)
btnExport.grid(sticky="NEWS", padx=10, pady=10)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()