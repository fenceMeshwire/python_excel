#!/usr/bin/env python3

# Python 3.9.5

# 01_read_spreadsheet.py

# Dependencies
import os
import pandas

path = 'C:\\Users\\user\\spreadsheets' # Spreadsheets are stored within this directory
os.chdir(path)

for dirname, subfolders, filenames in os.walk(path):
    for filename in filenames:
        # Find and read a single spreadsheet into DataFrame:
        if filename.find('xls') != -1:
            print(f'READING:\t{filename}')
            df = pd.read_excel(filename, sheet_name="Sheet1", usecols="A:F")
            
            # Parameters:
            # sheet_name:   Provides the name of the WorkSheet
            # usecols:      Indicate the columns, which hold desired data. 

df.info() # Show DataFrame object info
