import pandas as pd
import re

path = r"Extract Text-211029.xls"
excel = pd.ExcelFile(path)
sheets = excel.sheet_names
print (sheets)

df = pd.read_excel(excel,sheet_name="Summary")
print(df.describe())