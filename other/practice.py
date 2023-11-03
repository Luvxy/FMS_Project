import openpyxl
from openpyxl import Workbook
import random
import pandas as pd
from tkinter import filedialog as fd
from datetime import datetime

load_path = ""
save_path = "out.xlsx"

filetypes = (
    ('excel files', '*.xlsx'),
    ('All files', '*.*')
)

filename = fd.askopenfilename(
    title='Open a file',
    initialdir='/',
    filetypes=filetypes)

path = filename

current_date = datetime.now()
formatted_date = current_date.strftime("%m월%d일")

# 엑셀 파일을 데이터프레임으로 읽어옵니다.
df = pd.read_excel(path, sheet_name=formatted_date)

# 데이터프레임의 행 수를 얻습니다.
row_count = len(df)

# Print the value at cell (3, 3)
value = df.iloc[2,10]  # 2 corresponds to the third row and third column (0-based indexing)

print(value)
