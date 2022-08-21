#!/usr/bin/env python3

# Python 3.9.5

# 00_introductions.py

# Dependencies
import sys

xlsx = {"read": "OpenPyXL", "write": "OpenPyXL, XlsxWriter", "edit": "OpenPyXL"}
xlsm = {"read": "OpenPyXL", "write": "OpenPyXL, XlsxWriter", "edit": "OpenPyXL"}
xltx = {"read": "OpenPyXL", "write": "OpenPyXL", "edit": "OpenPyXL"}
xltm = xltx
xlsb = {"read": "pyxlsb", "write": "---", "edit": "---"}
xls = {"read": "xlrd", "write": "xlwt", "edit": "xlutils"}
xlt = xls

file_formats = {'xlsx': xlsx, 'xlsm': xlsm, 'xltx': xltx, 'xltm': xltm, 'xlsb': xlsb, 'xls': xls, 'xlt': xlt}

try:
    while True:
        print('Please choose the following options (q for exit):')
        print('======================================')
        print('xlsx, xlsm, xltx, xltm, xlsb, xls, xlt')
        print('======================================')
        file_format_input = input('Please enter the Excel File Format:')
        file_format_input = file_format_input.lower()
        for file_format in file_formats:
            if file_format_input == file_format:
                print()
                print('Read:\t', file_formats[file_format]["read"])
                print('Write:\t', file_formats[file_format]["write"])
                print('Edit:\t', file_formats[file_format]["edit"])
                print()
            if file_format_input == 'q':
                print('End of program.')
                sys.exit()

except KeyboardInterrupt:
    print('Program exit by CTRL + C')
