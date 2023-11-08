# button_functions.py
import json
import pandas as pd
from tkinter import filedialog as fd
from ui_FMS import *


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
def on_button9_clicked(date): # 시작
    print("Button 9 clicked")
    config_r = read_config()
    data = ''
    print(date)
    # 날짜 변경
    # 기부자 검색
    # 물품선택
    # 수량, 무게, 금액 변경
    # 저장
    
    