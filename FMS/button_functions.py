# button_functions.py
import json
import pandas as pd
from tkinter import filedialog as fd
from ui_FMS import *
import pyautogui as pag
import pyperclip as pc


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
    print("Button 9 clicked")
    config_r = read_config()
    df = pd.read_excel(config_r["item_path"])
    print(df)
    len_df = len(df)
    active_num = 1
    zero_num = 0
    count = 0
    
    # active_num = 선택된 기관, zero_num = 기관 수, count = 빵 등록 수
    for i in range(len_df):
        if df.iloc[active_num+i][0] == '' or df.iloc[active_num+i][0] == None:
            count += 1
        else:
            zero_num += 1
            break
    # 1. 빵 등록
    # 오늘 날짜 불러오기
    # 제공등록 클릭 (mouse point x=349, y=212)
    pag.click(349, 212)
    
    # 엑셀의 row 수 만큼 반복
    for i in range(count):
        # 등록 클릭 (mouse point x=1471, y=300)
        pag.click(1471, 300)
        # 기부자 검색 (x=713, y=368), (x=977, y=470)
        pag.click(713, 368)
        pc.copy(df.iloc[active_num+i][1])
        pag.hotkey('ctrl', 'v')
        pag.press('f2')
        pag.click(clicks=2, button='left',x=977, y=470)
        # 물품 검색 및 선택 (x=1873, y=445), (x=920, y=462), (x=1101, y=455), (x=918, y=532), (x=946, y=778)
        pag.click(1873, 445)
        pag.click(920, 462)
        pc.copy('기타빵')
        pag.hotkey('ctrl', 'v')
        pag.click(1101, 455)
        pag.press('f2')
        pag.click(x=918, y=532, clicks=2, button='left')
        pag.click(946, 778)
        # 수량 입력 (1, 1, 4)
        pc.copy(df.iloc[active_num+i][2])
        # 무게 입력 
        pag.press('tab')
        pc.copy(int(df.iloc[active_num+i][2])/5)
        pag.hotkey('ctrl', 'v')
        # 금액 입력
        pag.press('tab')
        pc.copy(df.iloc[active_num+i][4])
        pag.hotkey('ctrl', 'v')
        # 저장 (x=1862, y=293)
        pag.click(1862, 293)
        pag.press('esc')
    
    # 2. 빵 제공등록
    # 제공등록 클릭 (x=469, y=206)
    pag.click(469, 206)
    # 물품 검색 (x=889, y=369)
    pag.click(889, 369)
    pc.copy('기타빵')
    pag.hotkey('ctrl', 'v')
    pag.press('f2')
    # 물품 선택 (x=272, y=446)
    pag.click(272, 446)
    # plus (x=1870, y=419)
    pag.click(1870, 419)
    # 이용기관 선택 (x=681, y=398)
    pag.click(681, 398)
    # 기관 검색 (x=784, y=336)
    pag.click(784, 336)
    try:
        pc.copy(df.iloc[active_num][0])
    except:
        print("기관이 없습니다.")
        return
    pag.hotkey('ctrl', 'v')
    pag.press('f2')
    # 기관 선택 (x=962, y=473)
    pag.click(x=962, y=473, clicks=2, button='left')
    # 확인 (x=941, y=915)
    pag.click(941, 915)
    # 수량 수정 (x=1825, y=416)
    pag.click(1825, 416)
    pc.copy(df.iloc[active_num][2])
    pc.hotkey('ctrl', 'v')
    pag.press('enter')
    # 저장 (x=1860, y=291)
    pag.click(1860, 291)
    
    