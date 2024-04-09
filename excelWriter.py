import os
import sys
import openpyxl
import xlsxwriter
import pandas as pd


fileName = "test2.xlsx"
with pd.ExcelWriter(fileName) as writer:
    df = pd.DataFrame()
    df2 = pd.DataFrame()
    df3 = pd.DataFrame()
    df4 = pd.DataFrame()
    df5 = pd.DataFrame()
    df6 = pd.DataFrame()
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