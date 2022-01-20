import openpyxl
import pandas as pd

#Excel読み込み
df = pd.read_excel("test.xlsx", sheet_name = 0)
print(len(df))