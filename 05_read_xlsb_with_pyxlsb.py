#!/usr/bin/env python3

# Python 3.9.5

# 05_read_xlsb_with_pyxlsb.py

# Dependencies

import os
import pandas as pd
import pyxlsb

def get_xlsb_workbooks(path):
    result = []
    # Walk through the current working directory:
    for dirname, subfolders, filenames in os.walk(path):
        # Check each filename for xlsb:
        for filename in filenames:
            if filename.find('xlsb') != -1:
                result.append(filename)
                with pyxlsb.open_workbook(filename) as workbook:
                    for worksheet in (workbook.sheets):
                        print("WorkSheet.Name =", worksheet)
                        with workbook.get_sheet(worksheet) as WorkSheet:
                            worksheet_usedrange = WorkSheet.dimension
                            print("Worksheet.UsedRange =", worksheet_usedrange)
                            print()
    return result
                        
if __name__ == '__main__':
    # Change the current working directory:
    path = 'C:\\Users\\user\\...}'
    os.chdir(path)
    # Obtain a list of all xlsb workbooks
    workbooks = get_xlsb_workbooks(path)
    # Open the first WorkSheet of the first WorkBook
    with pyxlsb.open_workbook(workbooks[0]) as workbook:
        df = pd.read_excel(workbook, 0, engine='pyxlsb')

    # Check the imported data:
    df.head()
