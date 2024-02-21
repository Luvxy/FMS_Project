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

#File I/O path load
active_num = 0
path = "out.xlsx"
do_list=["이용자 등록","접수등록","제공등록","접수목록 찾기",
            "제공목록 찾기","접수현황 수정","제공현황 수정"]
date_to = datetime.date.today()
type_num = 1
name = active_num*5+1
id_number = active_num*5+2
address = active_num*5+3
phone_number = active_num*5+4
dataF = {'이름': [], '주민등록번호': [], '휴대전화': [], '주소': []}

def add_data(name, id_number, phone_number, address, existing_data_path='demo.xlsx'):
    try:
        # Read existing data from Excel file
        df_initial = pd.read_excel(existing_data_path)
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame
        df_initial = pd.DataFrame()

    # Add more data to the DataFrame
    new_data = {'이름': [name], '주민등록번호': [id_number], '휴대전화': [phone_number], '주소': [address]}
    df = pd.concat([df_initial, pd.DataFrame(new_data)], ignore_index=True)

    # Create a Pandas Excel writer object using XlsxWriter as the engine
    writer = pd.ExcelWriter(existing_data_path, engine='xlsxwriter')

    # Write data to the Excel sheet
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Close the file
    writer.close()

def data_print(_data, num, type_num):
    global active_num
    active_num += int(num)
    print(active_num)

#엑셀 path select
def select_file(sheet_name):
    global path
    global data
    global active_num
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
    data = get_data_from_excel(path, sheet_name)
    data_print(data, 0, type_num)
    
#엑셀 불러오기
def get_data_from_excel(_path, date):
    global active_num
    
    # 엑셀 파일을 데이터프레임으로 읽어옵니다.
    try:
        df = pd.read_excel(_path, sheet_name=str(date))
    except:
        messagebox.showinfo("error","날짜를 확인해주세요. 엑셀의 시트 이름과 날짜를 일치시켜주세요.")
        return

    # data return
    return df

# no address
def no_address(name, id_number, phone_number, Address):
    return

# restart
def restart_sign_new_user(sheet_name):
    global data
    global active_num
    global name
    global id_number
    global address
    global phone_number
    user_name = str(data.loc[name].iloc[0])
    a = str(data.loc[id_number].iloc[0])
    
    user_age = a.split('.')
    
    x1, y1, x2, y2 = 700, 480, 775, 510
    y_plus = 28
    is_data = False
    i = 1
    
    pyautogui.click(x=1736, y=295)
    time.sleep(1)
    pyautogui.click(x=969, y=340)
    pyperclip.copy(user_name)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.click(x=356, y=388)
    
    # 생년월일 입력
    a = str(data.loc[id_number].iloc[0])
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

        # Get the pixel color at the center of the bounding box
        center_pixel_color = screenshot.getpixel((box_size, box_size))

        if center_pixel_color == green_color:
            print("Green color found at ({}, {})".format(*target_pixel))
            is_data = True
            break
        else:
            print("not found")
        
        # Check if 3 seconds have passed
        if time.time() - start_time >= 6:
            break
    
    if is_data:
        pyautogui.click(x=1001, y=507, clicks=2, button='left')
        time.sleep(1)
        pyautogui.press("esc")
        time.sleep(0.5)
        pyautogui.press("esc")

        # 이용자명 입력(mouse_pos = 408, 432)
        pyautogui.click(x=489, y=431, clicks=1, button='left')
        user_name = str(data.loc[name].iloc[0])
        pyperclip.copy("○"+user_name)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        #user_type1 

        
        # 기관 입력
        pyautogui.click(x=604, y=494)
        time.sleep(0.5)
        pyautogui.click(x=795, y=395)
        time.sleep(0.5)
        pyperclip.copy(sheet_name)
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("f2")
        time.sleep(2)
        pyautogui.click(x=944, y=486, clicks=2, button='left') # x=955, y=484
        time.sleep(1)
        
        # 이용자특성은 전부 기초생활 and 독거노인으로 통일
        pyautogui.click(x=509, y=465)
        time.sleep(0.5)
        pyautogui.click(x=484, y=511)
        
        # 주소 입력(mouse_pos1 = 427, 523)(mouse_pos2 = 861, 430)(mouse_pos3 = 1132, 442)(mouse_pos4 = 948, 521)
        pyautogui.click(x=425, y=522, clicks=1, button='left')
        time.sleep(1)
        pyautogui.click(x=873, y=434, clicks=1, button='left')
        pyperclip.copy(data.loc[address].iloc[0])
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.click(x=1132, y=441, clicks=1, button='left')
        time.sleep(1)
        # 가장 위에 클릭
        #Point(x=1144, y=212)
        start_time = time.time()
        green_color = (254, 230, 200)  # (R, G, B) values for pure green
        target_pixel = (1028, 513)
        box_size = 10
        is_adress = False
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

            # Get the pixel color at the center of the bounding box
            center_pixel_color = screenshot.getpixel((box_size, box_size))

            if center_pixel_color == green_color:
                print("Green color found at ({}, {})".format(*target_pixel))
                is_adress = True
                break
            else:
                print("not found")
            
            # Check if 3 seconds have passed
            if time.time() - start_time >= 3:
                break
            
        if is_adress:
            pyautogui.click(x=948, y=525, clicks=2, button='left')
            time.sleep(1)
        else:
            # close the windows
            pyautogui.click(x=1001, y=776)
            # save the name to json or excel
            
            user_names = data.loc[name].iloc[0] + "(이미 등록됨)"
            user_id_number = data.loc[id_number].iloc[0]
            user_phone_number = data.loc[phone_number].iloc[0]
            user_address = data.loc[address].iloc[0]
            add_data(user_names, user_id_number, user_phone_number, user_address)

        # 번호 입력(mouse_pos1 = 395, 561)(mouse_pos2 = 384, 603)
        user_num = str(data.loc[phone_number].iloc[0]).split('-')
        if(user_num[0] == "010"):
            pyautogui.click(x=400, y=558, clicks=1, button='left')
            pyautogui.click(x=384, y=601, clicks=1, button='left')
            pyautogui.press("tab")
            pyperclip.copy(user_num[1])
            pyautogui.hotkey('ctrl', 'v')
            pyperclip.copy(user_num[2])
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
        else:
            pyautogui.click(x=400, y=558, clicks=1, button='left')
            for i in range(23):
                pyautogui.press("down")
            pyautogui.press("enter")
            pyautogui.press("tab")
            pyperclip.copy(user_num[1])
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press("tab")
            pyperclip.copy(user_num[2])
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
        
        # 이머전시 해제
        pyautogui.click(x=1040, y=655)
        time.sleep(0.2)
        pyautogui.click(x=1099, y=670)
        time.sleep(0.2)

        # 지원기간 선택(mouse_pos1 = 1035, 587)
        time.sleep(1)
        pyautogui.click(x=1065, y=615, clicks=1, button='left')
        for i in range(2):
            pyautogui.press('up')
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.click(x=1035, y=589, clicks=1, button='left')
        for i in range(13):
            pyautogui.press('up')
        for i in range(12):
            pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(0.5)

        # 비고
        pyautogui.click(x=630, y=824)
        time.sleep(0.5)
        pyautogui.press('right')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyperclip.copy(str(sheet_name)+' 연계(2024)')
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        
        global type_num
        # Your code to execute when 'Yes' is clicked
        pyautogui.click(x=1735, y=294)
        pyautogui.press('enter')
        time.sleep(6)
        pyautogui.press('enter')
        pyautogui.click(x=1791, y=292)
        data_print(data,1, type_num)
    else:
        user_names = data.loc[name].iloc[0]
        user_id_number = data.loc[id_number].iloc[0]
        user_phone_number = data.loc[phone_number].iloc[0]
        user_address = data.loc[address].iloc[0]
        add_data(user_names, user_id_number, user_phone_number, user_address)
        pyautogui.click(x=1805, y=291)
        
# Function to handle 'Yes' button click
def on_click_yes():
    global type_num
    time.sleep(1)
    # Your code to execute when 'Yes' is clicked
    # messagebox.showinfo("안내","이용자 정보를 저장합니다.'")
    pyautogui.click(x=1804, y=280)
    pyautogui.press('enter')
    time.sleep(6)
    pyautogui.press('enter')
    pyautogui.click(x=1804, y=280)
    data_print(data,1, type_num)

# Function to handle 'No' button click
def on_click_no():
    # Your code to execute when 'No' is clicked
    messagebox.showinfo("Info", "You clicked 'No'")

# Function to display the message box
def message_box():
    # response = messagebox.askquestion("확인","모든 정보가 정확합니까?")   
    # if response == "yes":
    on_click_yes()
    # else:
    #     on_click_no()
        
# 이용자 등록
def sign_new_user(user_data, date, sheet_name):
    global active_num
    global name
    global id_number
    global address
    global phone_number

    # ID자동생성 체크박스 클릭(mouse_pos = 456, 340)
    pyautogui.moveTo(x=457, y=342)
    pyautogui.click(clicks=1, button='left')
    time.sleep(0.5)

    # 이용자명 입력(mouse_pos = 408, 432)
    pyautogui.click(x=461, y=431, clicks=1, button='left')
    is_spe = True
    user_name = str(user_data.loc[name].iloc[0])
    pyperclip.copy("○"+user_name)
    pyautogui.hotkey('ctrl', 'v')
    
    time.sleep(0.5)
    #주민등록 번호 입력(mouse_pos1 = 665, 434)
    pyautogui.click(x=665, y=431, clicks=1, button='left')
    user_num = str(user_data.loc[id_number].iloc[0]).split('-')
    pyperclip.copy(user_num[0])
    pyautogui.hotkey('ctrl', 'v')
    pyperclip.copy(user_num[1])
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    is_match = False
    #Point(x=1144, y=212)
    start_time = time.time()
    green_color = (250, 253, 244)  # (R, G, B) values for pure green
    target_pixel = (1091, 210)
    box_size = 15
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
        pyautogui.press('esc')
        restart_sign_new_user(sheet_name)
        return
    
    #user_type1
    pyautogui.click(x=471, y=465)
    time.sleep(0.5)
    pyautogui.click(x=484, y=511)
    time.sleep(0.5)
    pyautogui.click(x=1051, y=464)
    time.sleep(0.5)
    pyautogui.press("down")
    pyautogui.press("down")
    time.sleep(0.2)
    pyautogui.press('enter')
    
    # 기관 입력
    time.sleep(1)
    pyautogui.click(x=604, y=494)
    time.sleep(1)
    pyautogui.click(x=795, y=395)
    time.sleep(0.5)
    pyperclip.copy(str(sheet_name))
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press("f2")
    time.sleep(2)
    pyautogui.click(x=944, y=486, clicks=2, button='left') # x=955, y=484
    time.sleep(1)
    
    # 주소 입력
    pyautogui.click(x=425, y=522, clicks=1, button='left')
    time.sleep(1)
    pyautogui.click(x=873, y=434, clicks=1, button='left')
    pyperclip.copy(data.loc[address].iloc[0])
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.click(x=1132, y=441, clicks=1, button='left')
    time.sleep(1)
    # 가장 위에 클릭
    #Point(x=1144, y=212)
    start_time = time.time()
    green_color = (254, 230, 200)  # (R, G, B) values for pure green
    target_pixel = (1028, 513)
    box_size = 10
    is_adress = False
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

        # Get the pixel color at the center of the bounding box
        center_pixel_color = screenshot.getpixel((box_size, box_size))

        if center_pixel_color == green_color:
            print("Green color found at ({}, {})".format(*target_pixel))
            is_adress = True
            break
        else:
            print("not found")
        
        # Check if 3 seconds have passed
        if time.time() - start_time >= 3:
            break
        
    if is_adress:
        pyautogui.click(x=948, y=525, clicks=2, button='left')
        time.sleep(1)
    else:
        # close the windows
        pyautogui.click(x=1001, y=776)
        # save the name to array
        
        user_names = data.loc[name].iloc[0]
        user_id_number = data.loc[id_number].iloc[0]
        user_phone_number = data.loc[phone_number].iloc[0]
        user_address = data.loc[address].iloc[0]
        add_data(user_names, user_id_number, user_phone_number, user_address)
        
        # restart
        pyautogui.click(x=1735, y=290)
        time.sleep(1)
        pyautogui.click(x=1806, y=300)
        return

    # 번호 입력(mouse_pos1 = 395, 561)(mouse_pos2 = 384, 603) 23
    user_num = str(data.loc[phone_number].iloc[0]).split('-')
    if(user_num[0] == "010"):
        pyautogui.click(x=400, y=558, clicks=1, button='left')
        pyautogui.click(x=384, y=601, clicks=1, button='left')
        pyautogui.press("tab")
        pyperclip.copy(user_num[1])
        pyautogui.hotkey('ctrl', 'v')
        pyperclip.copy(user_num[2])
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
    else:
        pyautogui.click(x=400, y=558, clicks=1, button='left')
        for i in range(23):
            pyautogui.press("down")
        pyautogui.press("enter")
        pyautogui.press("tab")
        pyperclip.copy(user_num[1])
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("tab")
        pyperclip.copy(user_num[2])
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

    # 지원기간 선택(mouse_pos1 = 1035, 587)
    pyautogui.click(x=1051, y=586, clicks=1, button='left')
    for i in range(12):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(0.5)
    
    # 비고
    pyautogui.click(x=630, y=824)
    time.sleep(0.5)
    pyautogui.press('right')
    time.sleep(0.5)
    pyperclip.copy(str(sheet_name)+' 연계(2024)')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    message_box()
    
## main
counting = int(input("몇번 반복: "))
# date_input = input("날짜(ex.2024.1.1): ")
date_input = "123"
sheet_name = "북익산노인복지센터" #input("시트 이름: ")익산시 오산면 영성길 18
select_file(sheet_name)
pyautogui.hotkey('alt','tab')

for i in range(counting):
    name = active_num*5+1
    id_number = active_num*5+2
    address = active_num*5+3
    phone_number = active_num*5+4
    
    print(str(data.loc[name].iloc[0]))
    print(str(data.loc[id_number].iloc[0]))
    print(str(data.loc[address].iloc[0]))
    print(str(data.loc[phone_number].iloc[0]))
    time.sleep(1)
    sign_new_user(data, date_input, sheet_name)
    active_num += 1
    time.sleep(1)
    