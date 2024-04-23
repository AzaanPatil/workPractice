import os
import sys
import openpyxl
import xlsxwriter
import pandas as pd

DM = {'Name': [], 'Age': [], 'Salary': []}
VCA = {'Name': [], 'Age': [], 'Salary': []}
ipAddresses = {'Computer': [], 'Phone': [], 'Laptop': []}
Cover = {'': [], '': [], 'Salary': []}
vSP2K = {'': [], '': [], '': []}
GMS = {'': [], '': [], 'Salary': []}

fileName = "test2.xlsx"
with pd.ExcelWriter(fileName) as writer:
    df = pd.DataFrame(DM)
    df2 = pd.DataFrame(VCA)
    df3 = pd.DataFrame(ipAddresses)
    df4 = pd.DataFrame(Cover)
    df5 = pd.DataFrame(vSP2K)
    df6 = pd.DataFrame(GMS)
    df.to_excel(writer, sheet_name='DM')
    df2.to_excel(writer, sheet_name='VCA')
    df3.to_excel(writer, sheet_name='IP Addresses')
    df4.to_excel(writer, sheet_name='Cover')
    df5.to_excel(writer, sheet_name='vSP2K')
    df6.to_excel(writer, sheet_name='GMS')
    worksheet0 = writer.sheets['vSP2K']
    worksheet0.hide()
    worksheet1 = writer.sheets['GMS']
    worksheet1.hide()

fileExists = os.path.exists(fileName)

if fileExists:
    print(fileName," created successfully")
else:
    print("Error, ", fileName ," not created")