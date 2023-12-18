import numpy as np
import time
import pandas as pd
import threading
import tkinter as tk
import tkinter.ttk as ttk
import pyautogui as pa
import pyperclip as pc
import cv2
import pytesseract as tr
import re
from tkinter import filedialog as fd
from tkinter import messagebox
import datetime
from tkinter import simpledialog
from PIL import ImageGrab
import traceback

# load excel file
# 엑셀 파일 불러오기
sheet_name = input("엑셀 시트명을 입력하세요(ex.12월15일): ")
fd = fd.askopenfilename(title="파일을 선택하세요.", 
                        filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
df = pd.read_excel(fd, sheet_name=sheet_name)

data = df.values
data = data.tolist()
dt = data[2]
print(dt[4])
count = 2


while True:
    pa.hotkey('alt', 'tab')
    dt = data[count]
    # ID 자동생성
    pa.click(x=458, y=342)
    time.sleep(0.5)
    # 이름 4 13
    pa.click(x=451, y=431)
    time.sleep(0.5)
    try:
        sum = str(dt[13]).split('-')
        pc.copy("＊"+str(dt[4]))
    except:
        pc.copy(str(dt[4]))
    pa.hotkey('ctrl', 'v')
    pa.press('tab')
    # 주민번호 5
    try:
        time.sleep(0.5)
        num = str(dt[5]).split('-')
        pc.copy(num[0])
        pa.hotkey('ctrl', 'v')
        pc.copy(num[1])
        pa.hotkey('ctrl', 'v')
        time.sleep(0.5)
    except:
        print("주민번호 오류")
    
    # 예외처리(이미 등록된 이용자)
    is_match = False
    #Point(x=1144, y=212)
    start_time = time.time()
    green_color = (27, 119, 228)  # (R, G, B) values for pure green
    target_pixel = (1159, 196)
    box_size = 10
    while True:
        box = (
            target_pixel[0] - box_size,
            target_pixel[1] - box_size,
            target_pixel[0] + box_size,
            target_pixel[1] + box_size
        )
        # Capture only the region around the target pixel
        screenshot = ImageGrab.grab(bbox=box)

        # Get the pixel color at the center of the bounding box
        center_pixel_color = screenshot.getpixel((box_size, box_size))

        if center_pixel_color == green_color:
            print("Green color found at ({}, {})".format(*target_pixel))
            is_match = True
            break
        else:
            print("not found")
        
        # Check if 3 seconds have passed
        if time.time() - start_time >= 4:
            break

    if is_match:
        time.sleep(1)
        pa.press('esc')
        continue
    
    # 이용자구분 9
    pa.click(x=485, y=463)
    time.sleep(0.5)
    # 이용자발굴지 9
    pa.click(x=779, y=463)
    pa.press('down')
    pa.press('down')
    pa.press('down')
    time.sleep(0.5)
    # 이용자분류 10
    pa.click(x=1077, y=463)
    time.sleep(0.5)
    # 주소 6
    pa.click(x=425, y=525)
    time.sleep(3)
    pa.click(x=937, y=436)
    time.sleep(0.5)
    pc.copy(str(dt[6]))
    pa.hotkey('ctrl', 'v')
    pa.press('f2')
    time.sleep(3)
    pa.click(x=998, y=520, clicks=2, button='left')
    
    # 주소 검색 오류 예외처리
    
    # 전화번호 7
    try:
        num = str(dt[7]).split('-')
        pa.click(x=385, y=557)
        time.sleep(0.5)
        pa.press('tab')
        pc.copy(num[1]) 
        pa.hotkey('ctrl', 'v')
        pa.press('tab')
        pc.copy(num[2])
        pa.hotkey('ctrl', 'v')
    except:
        print("전화번호 오류")
        
    # 지원기간
    pa.click(x=1018, y=588)
    # 특이사항 13
    pa.click(x=830, y=748)
    pc.copy("특이사항")
    pa.hotkey('ctrl', 'v')
    
    count += 1
    
    pa.hotkey('alt', 'tab')
    
    end = input("Do you want to end the program? (y/n)")
    if end == "y":
        break
    else:
        continue