#!/usr/bin/env python3

# Python 3.9.5

# 07_write_df_to_excel.py

# Dependencies
import os
import pandas as pd
from pathlib import Path

os.chdir(Path.home()) # Change current working directory to home (storing output).

data = [[1, 2, 3, 4], [5, 6, 7, 8]]

df = pd.DataFrame(data, columns=["First", "Second", "Third", "Fourth"])
df.head() # Check dataframe

# Write to Excel with index:
with pd.ExcelWriter("Pandas_Output.xlsx") as writer:
    df.to_excel(writer, sheet_name="Sheet1", startrow=0, startcol=0)

# Write to Excel with NO index:
with pd.ExcelWriter("Pandas_Output.xlsx") as writer:
    df.to_excel(writer, sheet_name="Sheet1", startrow=0, startcol=0, index=False)
