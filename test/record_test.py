import numpy as np
import time
import pandas as pd
import threading
import tkinter as tk
import tkinter.ttk as ttk
import pyautogui
import pyperclip
import cv2
import pytesseract as tr
import re
from tkinter import filedialog as fd
from tkinter import messagebox
import datetime
from tkinter import simpledialog
from PIL import ImageGrab
import traceback

active_num = 2
path = "out.xlsx"
do_list=["이용자 등록","접수등록","제공등록","접수목록 찾기",
            "제공목록 찾기","접수현황 수정","제공현황 수정"]
date_to = datetime.date.today()
date = "2024.2.2"
data = ''
type_num = 1

def data_print(_data, num, type_num):
    global active_num
    active_num += int(num)
    print(active_num)

#엑셀 path select
def select_file():
    global path
    global data
    global active_num
    global date
    global type_num
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
    data = get_data_from_excel(path, date)
    #log(str(data.iloc[int(active_num)]))
    data_print(data, 0, type_num)

#엑셀 불러오기
def get_data_from_excel(_path, date):
    global active_num
    
    # 엑셀 파일을 데이터프레임으로 읽어옵니다.
    try:
        df = pd.read_excel(_path, sheet_name=str(date))
    except:
        messagebox.showinfo("error","날짜를 확인해주세요. 엑셀의 시트 이름과 날짜를 일치시켜주세요.")

    # data return
    return df

# restart
def restart_sign_new_user():
    global data
    global active_num
    num = int(active_num)
    user_name = str(data.iloc[active_num,4])
    a = str(data.iloc[active_num,3])
    
    user_age = a.split('.')
    
    x1, y1, x2, y2 = 700, 480, 775, 510
    y_plus = 28
    is_data = False
    i = 1
    
    pyperclip.copy(user_name)
    pyautogui.click(x=907, y=339, clicks=1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.click(x=356, y=388)
    
    # 생년월일 입력
    a = str(data.iloc[num, 5])
    print(a)
    age = a.split('-')
    print(age)
    sex_num = age[1]
    print(sex_num[0])
    if(sex_num[0] == '1' or sex_num[0] == '2'):
        print("19")
        time.sleep(0.1)
        pyautogui.press('1')
        time.sleep(0.1)
        pyautogui.press('9')
    elif(sex_num[0] == '3' or sex_num[0] == '4'):
        print("20")
        time.sleep(0.1)
        pyautogui.press('2')
        time.sleep(0.1)
        pyautogui.press('0')
    else:
        print("age error")
        return
    time.sleep(0.1)
    pyperclip.copy(str(age[0]))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)
    pyautogui.press("enter")
    time.sleep(0.1)
    pyautogui.press("f2")
    time.sleep(0.5)
    
    # 가장 위에 클릭
    #Point(x=1144, y=212)
    start_time = time.time()
    green_color = (254, 230, 200)  # (R, G, B) values for pure green
    target_pixel = (282, 515)
    box_size = 10
    time.sleep(2)
    while True:
        box = (
            target_pixel[0] - box_size,
            target_pixel[1] - box_size,
            target_pixel[0] + box_size,
            target_pixel[1] + box_size
        )
        # Capture only the region around the target pixel
        screenshot = ImageGrab.grab(bbox=box)

        # Get the pixel color at the center of the bounding 신경희
        center_pixel_color = screenshot.getpixel((box_size, box_size))

        if center_pixel_color == green_color:
            print("Green color found at ({}, {})".format(*target_pixel))
            is_data = True
            break
        else:
            print("not found")
        
        # Check if 3 seconds have passed
        if time.time() - start_time >= 4:
            break
    
    if is_data:
        pyautogui.click(x=1001, y=507, clicks=2, button='left')
        time.sleep(1)
        
        pyautogui.press("esc")
        time.sleep(0.5)
        pyautogui.press("esc")
        time.sleep(0.5)
        pyautogui.click(x=1040, y=655)
        time.sleep(0.2)
        pyautogui.click(x=1099, y=670)
        time.sleep(0.2)
        pyautogui.click(x=1756, y=286)
        time.sleep(0.5) 
        pyautogui.press("enter")
        time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(5)
        pyautogui.press('enter')
        
select_file()
pyautogui.hotkey('alt','tab')
        
while True:
    restart_sign_new_user()
    active_num += 1