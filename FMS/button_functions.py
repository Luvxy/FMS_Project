# button_functions.py
import json
import pandas as pd
import ui_FMS as ui
from tkinter import filedialog as fd
from ui_FMS import *
import pyautogui as pag
import pyperclip as pc
import time


################################################################################
# 이용자등록 tab
################################################################################
def read_config():
    try:
        with open("config.json", 'r', encoding="utf-8") as config_file:
            config = json.load(config_file)
        return config
    except FileNotFoundError:
        print("Configuration file not found. Creating a new one.")
        return {}

def write_config(config):
    with open('config.json', 'w', encoding='utf-8') as config_file:
        json.dump(config, config_file, indent=4, ensure_ascii=False)

def on_button1_clicked(): # 다음
    print("Button 1 clicked")

def on_button2_clicked(): # 이전
    print("Button 2 clicked")

def on_button3_clicked(): # 파일선택
    print("Button 3 clicked")
    config_r = read_config()
    path = ''

    filetypes = (
        ('excel files', '*.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    path = filename
    config_r['user_path'] = path #void return {= 짱구 ☆  ^0^ ☆ 짱구 =}
    
    name = path.split('/')[len(path.split('/'))-1]
    print(name)
    
    print(config_r['user_path'])
    write_config(config_r)

def on_button4_clicked():
    print("Button 4 clicked")
    
################################################################################
# 제공등록 tab
################################################################################
def on_button5_clicked(): # 파일선택
    print("Button 5 clicked")
    config_r = read_config()
    path = ''

    filetypes = (
        ('excel files', '*.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    path = filename
    config_r['item_path'] = path
    
    name = path.split('/')[len(path.split('/'))-1]
    print(name)
    
    print(config_r['item_path'])
    write_config(config_r)

def on_button6_clicked(): # 시작
    print("Button 6 clicked")
    # load config
    config_r = read_config()
    # access excel file
    path = config_r['item_path']
    df = pd.read_excel(path)
    print(df.lioc[0])
    # start macro
    # 1. 날짜변경
    # 2. 물품검색
    # 3. 물품선택
    # 4. 이용자 검색
    # 5. 수량변경
    # 6. 저장
    
def on_button7_clicked(): # 다음
    print("Button 7 clicked")
    
def on_button8_clicked(): # 이전
    print("Button 8 clicked")

################################################################################
# 접수등록 tab
################################################################################
def on_button12_clicked(): # 파일선택
    print("Button 3 clicked")
    config_r = read_config()
    path = ''

    filetypes = (
        ('excel files', '*.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    path = filename
    config_r['item_path'] = path #void return {= 짱구 ☆  ^0^ ☆ 짱구 =}
    
    name = path.split('/')[len(path.split('/'))-1]
    print(name)
    
    print(config_r['item_path'])
    write_config(config_r)

def on_button9_clicked(date): # 시작
    print("파일 불러오기")
    config_ = read_config()
    # date form change (ex. 2021-01-01 -> 0101)
    date = date.split('-')
    date = date[1]+date[2]
    print(date)
    
    try:
        df = pd.read_excel(config_["bread_path"], sheet_name=date)
        print(df)
    except:
        print("시트 이름이 잘못되었습니다.")
        exit()

    # num = input('기관 수: ')
    num = 1
    # sum = input("시작 행: ")
    
    print("빵 등록")
    
    row_num = int(sum)
    data_fr = df.iloc[row_num]
    print("기관 명:"+ str(data_fr.iloc[0]))
    print(row_num)
    
    pag.hotkey('alt', 'tab')
    time.sleep(0.5)
    for i in range(int(num)):
        print(i)
        data = df.iloc[row_num+i]
        # print(data)
        print("기관명: "+str(data.iloc[1]))
        print("만약 20000원 이면 케이크로 등록")
        if data.iloc[3] == 20000:
            print("케이크 등록")
            # pag.click(x=454, y=199)
            # time.sleep(0.5)
            # pag.click(x=1434, y=284)
            # time.sleep(0.5)
            # pag.click(x=709, y=357)
            # time.sleep(0.5)
            # pag.click(x=1168, y=362)
            # pc.copy(str(data.iloc[1]))
            # pag.hotkey('ctrl', 'v')
            # time.sleep(0.5)
            # pag.press('f2')
            # time.sleep(0.5)
            # pag.click(x=953, y=467, clicks=2, button='left')
            # time.sleep(0.2)
            # pag.click(x=1872, y=431)
            # time.sleep(0.5)
            # pag.click(x=955, y=455)
            # time.sleep(0.5)
            # pc.copy('기타케익')
            # pag.hotkey('ctrl', 'v')
            # time.sleep(0.5)
            # pag.click(x=1101, y=453)
            # time.sleep(0.5)
            # pag.press('f2')
            # time.sleep(0.5)
            # pag.click(x=1069, y=526, clicks=2, button='left')
            # time.sleep(0.2)
            # pag.click(x=950, y=767)
            # time.sleep(0.2)
            # pag.click(x=736, y=504)
            # pc.copy(str(data.iloc[2]))
            # time.sleep(0.2)
            # pag.hotkey('ctrl', 'v')
            # pag.press('tab')
            # pc.copy(str(int(data.iloc[2])))
            # pag.hotkey('ctrl', 'v')
            # pag.press('tab')
            # pag.press('tab') 
            # pag.press('tab')
            # pag.press('tab')
            # pc.copy(str(data.iloc[4]))
            # pag.hotkey('ctrl', 'v')
            # time.sleep(0.2)
            # pag.click(x=1867, y=281)
            # pag.press('enter')
            # time.sleep(0.2)
            # pag.press('enter')
            # time.sleep(0.2)
            # pag.press('enter')
            # time.sleep(0.2)
        else:
            print("빵 등록")
            # pag.click(x=454, y=199)
            # time.sleep(0.5)
            # pag.click(x=1434, y=284)
            # time.sleep(0.5)
            # pag.click(x=709, y=357)
            # time.sleep(0.5)
            # pag.click(x=1168, y=362)
            # pc.copy(str(data.iloc[1]))
            # pag.hotkey('ctrl', 'v')
            # time.sleep(0.5)
            # pag.press('f2')
            # time.sleep(0.5)
            # pag.click(x=953, y=467, clicks=2, button='left')
            # time.sleep(0.2)
            # pag.click(x=1872, y=431)
            # time.sleep(0.5)
            # pag.click(x=955, y=455)
            # time.sleep(0.5)
            # pc.copy('기타빵')
            # pag.hotkey('ctrl', 'v')
            # time.sleep(0.5)
            # pag.click(x=1101, y=453)
            # time.sleep(0.5)
            # pag.press('f2')
            # time.sleep(0.5)
            # pag.click(x=1069, y=526, clicks=2, button='left')
            # time.sleep(0.2)
            # pag.click(x=950, y=767)
            # time.sleep(0.2)
            # pag.click(x=736, y=504)
            # pc.copy(str(data.iloc[2]))
            # time.sleep(0.2)
            # pag.hotkey('ctrl', 'v')
            # time.sleep(0.2)
            # pag.press('tab')
            # time.sleep(0.2)
            # pc.copy(str(int(int(data.iloc[2])/5)))
            # pag.hotkey('ctrl', 'v')
            # print(int(int(data.iloc[2])/5))
            # pag.press('tab')
            # pag.press('tab')
            # pag.press('tab')
            # pag.press('tab')
            # pc.copy(str(data.iloc[4]))
            # pag.hotkey('ctrl', 'v')
            # time.sleep(0.2)
            # pag.click(x=1867, y=281)
            # pag.press('enter')
            # time.sleep(0.2)
            # pag.press('enter')
            # time.sleep(0.2)
            # pag.press('enter')
            # time.sleep(0.2)

    # pag.hotkey('alt', 'tab')
    print("빵 제공")