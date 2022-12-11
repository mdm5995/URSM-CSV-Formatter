import csv
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkinter.messagebox import showerror
import traceback

class WorkDay:
    def __init__(
            self, date, employeeNumber, 
            jobNumber, costCode, wageCT, 
            wageNY, wageOT, wageDBL, hoursReg, hoursOT, 
            hoursDBL, travel
            ):
        self.date = date
        self.employeeNumber = employeeNumber
        self.jobNumber = jobNumber
        self.costCode = costCode
        self.wageCT = wageCT
        self.wageNY = wageNY
        self.wageOT = wageOT
        self.wageDBL = wageDBL
        self.hoursReg = float(hoursReg)
        self.hoursOT = float(hoursOT)
        self.hoursDBL = float(hoursDBL)
        self.travel = float(travel)
        self.payRate = ''

def format_csv(input_data, output_data):
    # add encoding 'utf-8-sig' to remove BOM at beginning of csvs from excel
    with open(input_data, 'rt', newline='', encoding='utf-8-sig') as csv_file: 
        csv_reader = csv.DictReader(csv_file, dialect='excel')
        with open(output_data, 'wt', newline='') as new_file: 
            fieldnames = [
                    'Date', 
                    'Employee Number', 
                    'Job Number', 
                    'Extra',
                    'Cost Code', 
                    'Union ID',
                    'Union Local',
                    'Union Class',
                    'State',
                    'WC State',
                    'WC Code',
                    'Certified',
                    'Certified Class',
                    'Pay ID',
                    'Units',
                    'Rate',
                    'Amount'
                    ]
            csv_writer = csv.DictWriter(new_file, fieldnames = fieldnames)
    
            # loop through csv_reader lines
            for line in csv_reader:

                # Handle Null Employee # gracefully
                if (line['E#'] == None):
                        continue
                
                # translate 'shop' to proper sage job number
                if (line['JOB #'].lower() == 'shop' or line['JOB #'].lower() == 'pto'):
                    line['JOB #'] = '010111'
                else:
                    # return everything before '-'
                    line['JOB #'] = line['JOB #'].split('-')[0]
    
                # sage requires 4 digit employee number
                if (len(line['E#']) < 4):
                    zeroes = '0' * (4 - len(line['E#']))
                    line['E#'] = zeroes + line['E#']
    
                # change zero values to empty strings for sage
                if (float(line['OT WAGE']) == 0):
                    line['OT WAGE'] = ''
                if (float(line['DBL WAGE']) == 0):
                    line['DBL WAGE'] = ''

                ## if ct wage or ny wage != null, make Ot wage 1.5x that, and DBL wage 2x that
                # make WorkDay object for each line
                tempWorkDay = WorkDay(
                        line['DATE'],
                        line['E#'],
                        line['JOB #'],
                        line['COSTCODE'],
                        line['CTPWAGE'],
                        line['NYPWAGE'],
                        line['OT WAGE'],
                        line['DBL WAGE'],
                        line['REG HRS'],
                        line['OT HRS'],
                        line['DBL HRS'],
                        line['Travel'],
                        )
    
                # if CT wage or NY wage is not null, set payRate attribute 
                if (float(tempWorkDay.wageNY) != 0):
                    tempWorkDay.payRate = float(tempWorkDay.wageNY)
                if (float(tempWorkDay.wageCT) != 0):
                    tempWorkDay.payRate = float(tempWorkDay.wageCT)

                if (tempWorkDay.payRate != ''):
                    if(tempWorkDay.wageOT == ''):
                        tempWorkDay.wageOT = tempWorkDay.payRate * 1.5
                    if(tempWorkDay.wageDBL == ''):
                        tempWorkDay.wageDBL = tempWorkDay.payRate * 2
    
                # Regular hours
                if (tempWorkDay.hoursReg != 0):
                    csv_writer.writerow({
                        'Date': tempWorkDay.date, 
                        'Employee Number': tempWorkDay.employeeNumber, 
                        'Job Number': tempWorkDay.jobNumber, 
                        'Extra': '',
                        'Cost Code': tempWorkDay.costCode, 
                        'Union ID': '',
                        'Union Local': '',
                        'Union Class': '',
                        'WC Code': '',
                        'WC State': '',
                        'State': '',
                        'Certified': '',
                        'Certified Class': '',
                        'Pay ID': 'REG',
                        'Units': tempWorkDay.hoursReg,
                        'Rate': tempWorkDay.payRate,
                        'Amount': ''
                        })
    
                # OverTime Hours
                if (tempWorkDay.hoursOT != 0):
                    csv_writer.writerow({
                        'Date': tempWorkDay.date, 
                        'Employee Number': tempWorkDay.employeeNumber, 
                        'Job Number': tempWorkDay.jobNumber, 
                        'Extra': '',
                        'Cost Code': tempWorkDay.costCode, 
                        'Union ID': '',
                        'Union Local': '',
                        'Union Class': '',
                        'WC Code': '',
                        'WC State': '',
                        'State': '',
                        'Certified': '',
                        'Certified Class': '',
                        'Pay ID': 'OT',
                        'Units': tempWorkDay.hoursOT,
                        'Rate': tempWorkDay.wageOT,
                        'Amount': ''
                        })
    
                # Double Time Hours
                if (tempWorkDay.hoursDBL != 0):
                    csv_writer.writerow({
                        'Date': tempWorkDay.date, 
                        'Employee Number': tempWorkDay.employeeNumber, 
                        'Job Number': tempWorkDay.jobNumber, 
                        'Extra': '',
                        'Cost Code': tempWorkDay.costCode, 
                        'Union ID': '',
                        'Union Local': '',
                        'Union Class': '',
                        'WC Code': '',
                        'WC State': '',
                        'State': '',
                        'Certified': '',
                        'Certified Class': '',
                        'Pay ID': 'DBL',
                        'Units': tempWorkDay.hoursDBL,
                        'Rate': tempWorkDay.wageDBL,
                        'Amount': ''
                        })
    
                # Travel pay
                if (float(tempWorkDay.travel) != 0):
                    csv_writer.writerow({
                        'Date': tempWorkDay.date, 
                        'Employee Number': tempWorkDay.employeeNumber, 
                        'Job Number': tempWorkDay.jobNumber, 
                        'Extra': '',
                        'Cost Code': '01800', 
                        'Union ID': '',
                        'Union Local': '',
                        'Union Class': '',
                        'WC Code': '',
                        'WC State': '',
                        'State': '',
                        'Certified': '',
                        'Certified Class': '',
                        'Pay ID': 'TRAVEL',
                        'Units': '',
                        'Rate': '',
                        'Amount': tempWorkDay.travel
                        })
    # messagebox here
    messagebox.showinfo('Success', 'Formatting Completed Successfully')

# Start GUI



## init gui
root = tk.Tk()
root.title('URSM Payroll Formatter')
#root.resizable(False,False)
root.geometry('550x150')
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)
root.rowconfigure(2, weight=1)

## Add frame

frame = tk.Frame(root)
frame.pack(expand=True, padx=10, pady=10)

## override callback exception handling
def report_callback_exception(self, *args):
        showerror("An Error Occurred", message = traceback.format_exc(limit=-1))

root.report_callback_exception = report_callback_exception


## define globals
csv_filename = tk.StringVar()
csv_filename.set('No File Chosen')

## Choose CSV inputData

### modifies global variable 'csv_filename'
def change_csv_filename():
    global csv_filename 
    filename = filedialog.askopenfilename(
            title='Choose input CSV', 
            filetypes=(('Comma Separated Values', '*.csv'),)
            )
    csv_filename.set(filename)

open_button = tk.Button(
        frame,
        text='Select input CSV',
        command=change_csv_filename
        )

csv_filename_label = tk.Label(frame, textvariable = csv_filename)
file_chosen_label = tk.Label(frame, text='File Chosen: ')

## run format_csv
format_button = tk.Button(
        frame,
        text='Format CSV',
        command=lambda: format_csv(csv_filename.get(), csv_filename.get().split('.')[0] + '_FORMATTED' + '.txt')
        )

## Place items on screen
open_button.grid(row=0, column=0, sticky=tk.E+tk.W, padx=10, pady=10)
format_button.grid(row=0, column=1, sticky=tk.E+tk.W, padx=10, pady=10)
file_chosen_label.grid(row=1, column=0, columnspan=2, padx=10, pady=10)
csv_filename_label.grid(row=2, column=0, columnspan=2)


root.mainloop()

# End GUI
