import pandas as pd

file_path = "C:/Users/user/Desktop/지역아동센터 new 명단.xlsx"
# 파일 불러오기
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names

print(sheet_names)