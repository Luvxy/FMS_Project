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

path = 'H:/.shortcut-targets-by-id/1A0TIuPAsmbBdRF01K7yT89PSL4LTaONB/익산행복나눔마켓뱅크/2. 마켓, 곳간/1. 이용자 명단/1. 일일 명단 (시청)/\'23.11.03. 일일신청명단(3200~3203.xlsx'
date = '11월2일'
data = pd.read_excel(path, sheet_name=str(date))
active_num = 3
num = int(active_num)
user_name = str(data.iloc[active_num,4])
x1, y1, x2, y2 = 700, 500, 780, 515
y_plus = 28
is_data = True
i = 1
all_number = 1

while is_data:
    img = ImageGrab.grab((x1, y1, x2, y2))
    img.save("test"+str(i)+".png")

    image = cv2.imread("test"+str(i)+".png")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    text = tr.image_to_string(gray, lang='eng')

    # Use a regular expression to extract numbers from the text
    numbers = re.findall(r'\d+', text)
    print(numbers)
    
    # If you want to access individual numbers, you can access them by index
    if len(numbers) >= 2:
        number1 = numbers[0][2:4]
        number2 = numbers[1]
        number3 = numbers[2]
        all_number = number1+number2+number3
        print(all_number)
    else:
        break
        
    print(numbers)
    user_num = str(data.iloc[num, 5]).split('-')
    if user_num[0] == str(all_number):
        pyautogui.click(x=x1,y=y1,clicks=2,button="left")
        is_data = True
        break
                
    y1 += y_plus
    y2 += y_plus    

    i += 1
    print(i)
