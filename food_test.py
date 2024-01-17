import json
import pandas as pd
from tkinter import filedialog as fd
import pyautogui as pag
import pyperclip as pc
import time

def on_button9_clicked(): # 시작
    food_list = ['오레시피', '방교', '오유민'] # 이름
    cost_list = [3000, 6000, 10000] # 가격
    count_list = [0,5,5] # 수량
    weight_list = [0,1,2] # 무게
    num = input("오레시피의 수량을 입력하세요: ")
    count_list[0] = int(num)
    if num == '0':
        print('오레시피는 0개입니다.')
        food_list = ['방교', '오유민'] # 이름
        cost_list = [6000, 10000]
        count_list = [5,5]
        weight_list = [1,2]
        print(cost_list)
        print(count_list)
        print(weight_list)
    time_delay = 0.5
    pag.hotkey('alt', 'tab')
    
    
    # 1. 빵 등록
    # 오늘 날짜 불러오기
    # 제공등록 클릭 (mouse point x=349, y=212)
    # pag.click(349, 212)
    
    # 엑셀의 row 수 만큼 반복
    for i in food_list:
        num_count = food_list.index(i) # 인덱스
        cost = cost_list[num_count] # 가격
        count = count_list[num_count] # 수량
        weight = weight_list[num_count] # 무게
        
        print(i, cost, count, weight)
        
        pag.click(x=454, y=199)
        time.sleep(time_delay)
        pag.click(x=1434, y=284)
        time.sleep(time_delay)
        pag.click(x=709, y=357)
        time.sleep(time_delay)  
        pag.click(x=1168, y=362)
        pc.copy(i)
        pag.hotkey('ctrl', 'v')
        time.sleep(time_delay)
        pag.press('f2')
        time.sleep(time_delay)
        pag.click(x=953, y=467, clicks=2, button='left')
        time.sleep(time_delay)
        pag.click(x=1870, y=447)
        time.sleep(time_delay)
        pag.click(x=955, y=455)
        time.sleep(time_delay)
        pc.copy(i)
        pag.hotkey('ctrl', 'v')
        time.sleep(time_delay)
        # pag.click(x=1101, y=453)
        # time.sleep(time_delay)
        pag.press('f2')
        time.sleep(time_delay)
        pag.click(x=1069, y=526, clicks=2, button='left')
        time.sleep(time_delay)
        pag.click(x=950, y=767)
        time.sleep(time_delay)
        pag.click(x=749, y=519)
        # 수량
        pc.copy(str(count))
        time.sleep(time_delay)
        pag.hotkey('ctrl', 'v')
        time.sleep(time_delay)
        pag.press('tab')
        time.sleep(time_delay)
        # 무게
        if i == '오레시피':
            pc.copy('1')
        else:
            pc.copy(str(weight*count))
        pag.hotkey('ctrl', 'v')
        pag.press('tab')
        pag.press('tab')
        pag.press('tab')
        pag.press('tab')
        # 가격
        pc.copy(str(cost*count))
        pag.hotkey('ctrl', 'v')
        time.sleep(time_delay)
        # 저장
        pag.click(x=1869, y=291)
        pag.press('enter')
        time.sleep(time_delay)
        pag.press('enter')
        time.sleep(time_delay)
        pag.press('enter')
        time.sleep(time_delay)
    
on_button9_clicked()
