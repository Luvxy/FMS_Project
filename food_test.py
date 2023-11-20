import json
import pandas as pd
from tkinter import filedialog as fd
import pyautogui as pag
import pyperclip as pc
import time

active_num = -1 # 선택된 기관
zero_num = 1 # 기관 수
count = 0 # 빵집 수
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
    food_list = ['오레시피', '왕언니', '방교', '오유민']
    cost_list = [3000, 10000, 6000, 9000]
    count_list = [0,5,5,5]
    
    # 1. 빵 등록
    # 오늘 날짜 불러오기
    # 제공등록 클릭 (mouse point x=349, y=212)
    pag.click(349, 212)
    
    # 엑셀의 row 수 만큼 반복
    for i in food_list:
        count = count_list[food_list.index(i)]
        if i == '오레시피':
            num = input("오레시피의 수량을 입력하세요: ")
            # time.sleep(1)
            print('오레시피 수량: '+num)
            # 등록 클릭 (mouse point x=1471, y=300)
            print('등록 클릭')
            pag.click(1471, 300)
            # 기부자 검색 (x=713, y=368), (x=977, y=470)
            print('기부자 검색')
            pag.click(713, 368)
            pc.copy(i)
            print(i)
            pag.hotkey('ctrl', 'v')
            pag.press('f2')
            pag.click(clicks=2, button='left',x=977, y=470)
            # # 물품 검색 및 선택 (x=1873, y=445), (x=920, y=462), (x=1101, y=455), (x=918, y=532), (x=946, y=778)
            print("오레시피 검색")
            pag.click(1873, 445)
            pag.click(920, 462)
            pc.copy(i)
            print(i)
            pag.hotkey('ctrl', 'v')
            pag.click(1101, 455)
            pag.press('f2')
            pag.click(x=918, y=532, clicks=2, button='left')
            pag.click(946, 778)
            # 수량 입력 ()
            pc.copy(num)
            print("수량 입력: "+num)
            # 무게 입력 
            pag.press('tab')
            pc.copy("1")
            print("무게: 1")
            pag.hotkey('ctrl', 'v')
            print(str(int(num)/5))
            # # 금액 입력
            pag.press('tab')
            pc.copy(str(int(num)*cost_list[0]))
            print("금액: "+str(int(num)*cost_list[0]))
            pag.hotkey('ctrl', 'v')
            # 저장 (x=1862, y=293)
            pag.click(1862, 293)
            pag.press('esc')
            time.sleep(1)
        else:
            # 등록 클릭 (mouse point x=1471, y=300)
            pag.click(1471, 300)
            # 기부자 검색 (x=713, y=368), (x=977, y=470)
            pag.click(713, 368)
            pc.copy(i)
            pag.hotkey('ctrl', 'v')
            pag.press('f2')
            pag.click(clicks=2, button='left',x=977, y=470)
            # 물품 검색 및 선택 (x=1873, y=445), (x=920, y=462), (x=1101, y=455), (x=918, y=532), (x=946, y=778)
            print('물품 검색')
            pag.click(1873, 445)
            pag.click(920, 462)
            pc.copy(i)
            print(i)
            pag.hotkey('ctrl', 'v')
            pag.click(1101, 455)
            pag.press('f2')
            pag.click(x=918, y=532, clicks=2, button='left')
            pag.click(946, 778)
            # 수량 입력 ()
            print('수량 입력: '+str(count))
            pc.copy(str(count))
            # 무게 입력 
            print('무게 입력'+str(int(count)))
            pag.press('tab')
            pc.copy(str(int(count)))
            pag.hotkey('ctrl', 'v')
            # # 금액 입력
            pag.press('tab')
            pc.copy(str(cost_list[food_list.index(i)]*count))
            pag.hotkey('ctrl', 'v')
            print("금액: " + str(cost_list[food_list.index(i)]*count))
            # # 저장 (x=1862, y=293)
            pag.click(1862, 293)
            pag.press('esc')
    
on_button9_clicked()