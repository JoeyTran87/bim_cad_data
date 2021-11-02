import pandas as pd

path = r"F:\_NGHIEN CUU\_Github\Python\py_autocad\dwg\Extract Text-211029.xls"
excel = pd.ExcelFile(path)
sheets = excel.sheet_names
print (sheets)

df = pd.read_excel(excel,sheet_name="Summary")
print(df.describe())