import os
import sys
import pandas as pd

with pd.ExcelWriter("test2.xlsx") as writer:
    df = pd.DataFrame()
    df2 = pd.DataFrame()
    df3 = pd.DataFrame()
    df4 = pd.DataFrame()
    df.to_excel(writer, sheet_name='DM')
    df2.to_excel(writer, sheet_name='VCA')
    df3.to_excel(writer, sheet_name='IP Addresses')
    df4.to_excel(writer, sheet_name='Cover')