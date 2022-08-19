#!/usr/bin/env python3

# Python 3.9.5

# 02_read_with_multiple_worksheets.py

# Dependencies
import os
import pandas

path = 'C:\\Users\\user\\spreadsheets' # Spreadsheets are stored within this directory
os.chdir(path)

# ============================================================================================
# Access WorkBook, make multiple DataFrame objects by using the with-method:
for dirname, subfolders, filenames in os.walk(path):
    for filename in filenames:
        if filename.find('xls') != -1:
            with pd.ExcelFile("workbook.xlsx") as f:
                
                df1 = pd.read_excel(f, "Sheet1", skiprows=1, usecols="B:F", nrows=2)
                df2 = pd.read_excel(f, "Sheet2", skiprows=1, usecols="B:F", nrows=2)
                
                # Parameters:
                # skiprows: Skip over the indicated number of rows.
                # usecols:  Indicate the columns, which hold desired data.
                # nrows:    Number of rows, which are going to be read.
                
print(df1)
print(df2)
