import json
import pandas as pd
from tkinter import filedialog as fd
import pyautogui as pa
import pyperclip as pc
import time
from tkinter import *

print("파일 불러오기")
sh_name = input('시트 이름: ')
try:
    df = pd.read_excel("H:/.shortcut-targets-by-id/1A0TIuPAsmbBdRF01K7yT89PSL4LTaONB/익산행복나눔마켓뱅크/3. 양식 및 도구 (정리 필요 !!!)/아동센터 목록(+배분실적).xlsx", sheet_name=sh_name)
    print(df)
except:
    print("시트 이름이 잘못되었습니다.")
    exit()




while(True):
    num = input('기관 수: ')
    sum = input("시작 행: ")
    
    print("빵 등록")
    
    row_num = int(sum)
    data = df.iloc[row_num]
    print("기관 명:"+ str(data.iloc[0]))
    
    pa.hotkey('alt', 'tab')
    
    for i in range(int(num)):
        print(i)
        data = df.iloc[row_num+i]
        # print(data)
        print("기관명: "+str(data.iloc[1]))
        print("만약 20000원 이면 케이크로 등록")
        if data.iloc[3] == 20000:
            print("케이크 등록")
        else:
            print("빵 등록")

    print("빵 제공")
    
    end = input("종료하시겠습니까? (y/n)")
    if end == 'y':
        break
    else:
        continue