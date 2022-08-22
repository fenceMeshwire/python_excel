#!/usr/bin/env python3

# Python 3.9.5

# 05_read_xlsb_with_pyxlsb.py

# Dependencies

import os
import pandas as pd
import pyxlsb

# Change the current working directory:
path = 'C:\\Users\\user\\...'
os.chdir(path)

# Walk through the current working directory:
for dirname, subfolders, filenames in os.walk(path):
    # Check each filename for xlsb:
    for filename in filenames:
        if filename.find('xlsb') != -1:
            result = filename
            with pyxlsb.open_workbook(filename) as workbook:
                for worksheet in (workbook.sheets):
                    print("WorkSheet.Name =", worksheet)
                    with workbook.get_sheet(worksheet) as WorkSheet:
                        worksheet_usedrange = WorkSheet.dimension
                        print("Worksheet.UsedRange =", worksheet_usedrange)
                        print()
                        
# Import each WorkSheet into a DataFrame object.
# Count number of WorkSheets previously.
with pyxlsb.open_workbook(result) as workbook:
    df = pd.read_excel(workbook, 0, engine='pyxlsb')

# Check the imported data:
df.head()
              
