#!/usr/bin/env python3

# Python 3.9.5

# 06_read_pandas_openpyxl.py

# Dependencies
import pandas as pd
import os

path = 'C:\\Users\\user\\...'
os.chdir(path)
os.listdir('.')

with pd.ExcelFile("sample.xlsx", engine='openpyxl') as input_file:
  df = pd.read_excel(input_file, sheet_name='Sheet1')

df.head()
# ...
