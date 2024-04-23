import os
import sys
import openpyxl
import pandas as pd
import xlrd
import prettytable
from prettytable import PrettyTable

print("Program launched successfully")
os.system('pause')

print("Welcome to ExcelReader")
os.system('pause')

print("Command line arguments", sys.argv) # prints every parameter in sys.argv
os.system('pause')

print("Name of program", sys.argv[0]) # prints specific index of parameter in sys.argv
os.system('pause')

for parm in sys.argv: # prints every single parameter of sys.argv in separate array
    print("parm = ", parm)
    os.system('pause')

total_cmd_parms = len(sys.argv) # Calculates the amount of parameters in sys.argv including the name of the program.

num_parms = total_cmd_parms - 1 # Calculates the amount of parameters in sys.argv minus the name of the program.

fileName = sys.argv[1] # Shows the index of where the file name is
fileExists = os.path.exists(fileName) # Checks whether file exists
myExcel = PrettyTable() # Creates table that displays sheets in Excel file
myExcel.title = [fileName] # Creates a header that shows the file name
myExcel.field_names = ["Name", "Visibility"] # Names of columns in table

if fileExists: # Function for if the file is found
    print("file: ", fileName, "exists")
    dataFrame0 = pd.read_excel(fileName)
    df = pd.ExcelFile(fileName) # Function that reads excel file
    sheetNames= df.sheet_names # variable that stores names of sheets
    print(sheetNames)
    numSheets = len(sheetNames) # calculates number of sheets in excel file
    print("There are ", numSheets, " sheets in ", fileName, ".")
    sheets = df.book.worksheets # prints name of sheets and status as visible or hidden
    numHiddenSheets = 0
    numVisibleSheets = 0
    print(dataFrame0)
    for sheet in sheets: # function that checks whether sheet is hidden or visible
        myExcel.add_row([sheet.title, sheet.sheet_state]) # function that creates a table which displays name and visibility of sheets
        if sheet.sheet_state == "hidden":
            numHiddenSheets += 1
        else:
            numVisibleSheets += 1
    print(myExcel)
else:
    print("Error: file ", fileName ," does not exist") # function for if file is not found



