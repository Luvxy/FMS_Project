import json
import pandas as pd
from tkinter import filedialog as fd
import pyautogui as pag
import pyperclip as pc
import time

active_num = 0 # 선택된 기관
zero_num = 1 # 기관 수
count = 1 # 빵집 수
is_cake = False # 빵집 여부

def read_config():
    try:
        with open("config.json", 'r', encoding="utf-8") as config_file:
            config = json.load(config_file)
        return config
    except FileNotFoundError:
        print("Configuration file not found. Creating a new one.")
        return {}

def on_button9_clicked(): # 시작
    global active_num, zero_num, count, is_cake
    print("Button 9 clicked")
    config_r = read_config()
    sh_name = input("시트 이름을 입력하세요: ")
    df = pd.read_excel(config_r["item_path"], sheet_name=sh_name)
    df = df.fillna('')
    print(df)
    len_df = len(df)
    
    print(len_df)
    
    # active_num = 선택된 기관, zero_num = 기관 수, count = 빵 등록 수
    for i in range(5):
        data = df.iloc[active_num+i]
        print("검사중..."+str(data.iloc[1]))
        if str(data.iloc[1]).find('c') != -1:
            cake = active_num+i
            is_cake = True
            print("중지1")
            break
        elif data.iloc[0] != '':
            count += 1
            print("중지2..."+str(count))
            print(data.iloc[0])
            break
        else:
            count += 1
    # 1. 빵 등록
    # 오늘 날짜 불러오기
    # 제공등록 클릭 (mouse point x=349, y=212)
    pag.click(349, 212)
    
    if is_cake:
        data = df.iloc[cake]
        print("케이크를 등록합니다... "+str(data.iloc[1]))
        time.sleep(3)
        print("케이크 등록 완료")
        # # 등록 클릭 (mouse point x=1471, y=300)
        # pag.click(1471, 300)
        # # 기부자 검색 (x=713, y=368), (x=977, y=470)
        # pag.click(713, 368)
        # pc.copy(str(data.iloc[1]))
        # pag.hotkey('ctrl', 'v')
        # pag.press('f2')
        # pag.click(clicks=2, button='left',x=977, y=470)
        # # 물품 검색 및 선택 (x=1873, y=445), (x=920, y=462), (x=1101, y=455), (x=918, y=532), (x=946, y=778)
        # pag.click(1873, 445)
        # pag.click(920, 462)
        # pc.copy('기타')
        # pag.hotkey('ctrl', 'v')
        # pag.click(1101, 455)
        # pag.press('f2')
        # pag.click(x=918, y=532, clicks=2, button='left')
        # pag.click(946, 778)
        # # 수량 입력 (1, 1, 4)
        # pc.copy(str(data.iloc[2]))
        # # 무게 입력 
        # pag.press('tab')
        # pc.copy(int(data.iloc[2])/5)
        # pag.hotkey('ctrl', 'v')
        # # 금액 입력
        # pag.press('tab')
        # pc.copy(str(data.iloc[4]))
        # pag.hotkey('ctrl', 'v')
        # # 저장 (x=1862, y=293)
        # pag.click(1862, 293)
        # pag.press('esc')
    
    # 엑셀의 row 수 만큼 반복
    for i in range(count):
        data = df.iloc[active_num+i]
        print("빵 접수등록 중... "+str(data.iloc[1]))
        time.sleep(1)
        # # 등록 클릭 (mouse point x=1471, y=300)
        # pag.click(1471, 300)
        # # 기부자 검색 (x=713, y=368), (x=977, y=470)
        # pag.click(713, 368)
        # pc.copy(str(data.iloc[1]))
        # pag.hotkey('ctrl', 'v')
        # pag.press('f2')
        # pag.click(clicks=2, button='left',x=977, y=470)
        # # 물품 검색 및 선택 (x=1873, y=445), (x=920, y=462), (x=1101, y=455), (x=918, y=532), (x=946, y=778)
        # pag.click(1873, 445)
        # pag.click(920, 462)
        # pc.copy('기타빵')
        # pag.hotkey('ctrl', 'v')
        # pag.click(1101, 455)
        # pag.press('f2')
        # pag.click(x=918, y=532, clicks=2, button='left')
        # pag.click(946, 778)
        # # 수량 입력 (1, 1, 4)
        # pc.copy(str(data.iloc[2]))
        # # 무게 입력 
        # pag.press('tab')
        # pc.copy(int(data.iloc[2])/5)
        # pag.hotkey('ctrl', 'v')
        # # 금액 입력
        # pag.press('tab')
        # pc.copy(str(data.iloc[4]))
        # pag.hotkey('ctrl', 'v')
        # # 저장 (x=1862, y=293)
        # pag.click(1862, 293)
        # pag.press('esc')
    
    print("빵 제공등록 시작")
    print("기관명: "+str(data.iloc[0]))
    # 2. 빵 제공등록
    # 제공등록 클릭 (x=469, y=206)
    # pag.click(469, 206)
    # # 물품 검색 (x=889, y=369)
    # pag.click(889, 369)
    # pc.copy('기타빵')
    # pag.hotkey('ctrl', 'v')
    # pag.press('f2')
    # # 물품 선택 (x=272, y=446)
    # pag.click(272, 446)
    # # plus (x=1870, y=419)
    # pag.click(1870, 419)
    # # 이용기관 선택 (x=681, y=398)
    # pag.click(681, 398)
    # # 기관 검색 (x=784, y=336)
    # pag.click(784, 336)
    # data = df.iloc[active_num]
    # try:
    #     pc.copy(str(data.iloc[0]))
    # except:
    #     print("기관이 없습니다.")
    #     return
    # pag.hotkey('ctrl', 'v')
    # pag.press('f2')
    # # 기관 선택 (x=962, y=473)
    # pag.click(x=962, y=473, clicks=2, button='left')
    # # 확인 (x=941, y=915)
    # pag.click(941, 915)
    # # 수량 수정 (x=1825, y=416)
    # pag.click(1825, 416)
    # pc.copy(str(data.iloc[2]))
    # pag.hotkey('ctrl', 'v')
    # pag.press('enter')
    # # 저장 (x=1860, y=291)
    # pag.click(1860, 291)
    
on_button9_clicked()