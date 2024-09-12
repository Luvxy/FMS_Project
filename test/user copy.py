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
date = str(date_to.month)+"월"+str(date_to.day)+"일"
data = ''
type_num = 1
bin1 = ["긴급 위기상황 발생자", "읍면동장추천", "긴급지원대상자(9개월)", "기타읍면동장책임제 추천자(9개월)"]
bin2 = ["차상위계층","복지급여 탈락", "실직 및 휴폐업","주거,교육수급자(6개월)", "차상위계층(6개월)", "복지급여탈락자중지자(6개월)"]
bin3 = ["실직 및 휴폐업", "소득감소자", "생계,의료급여수급자(3개월)", "위기상황발생자(3개월)", "실직 및 휴폐업(3개월)", "소득감소자(3개월)"]
type_list = ["결식아동","다문화가정","독거어르신","소년소녀가장",
                    "외국인노동자","재가장애인","저소득가정","조손가정",
                    "한부모가정","기타","청장년 1인가구","미혼모부가구",
                    "부부중심가구","노인부부가구","새터민가구","공통체가구"]
    
user_type_list = ["긴급 위기상황 발생자","수급자","차상위계층",
                    "소득감소자","복지급여 탈락", "실직 및 휴폐업", "읍면동장추천"]

date_to = datetime.date.today()
type_num = 1
name = 0
id_number = 1
address = 2
phone_number = 3
spe1 = 0
spe1 = 0
dataF = {'이름': [], '주민등록번호': [], '휴대전화': [], '주소': []}

def add_data(name, id_number, phone_number, address, sheet_name, existing_data_path='demo.xlsx'):
    try:
        # Read existing data from Excel file
        df_initial = pd.read_excel(existing_data_path)
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame
        df_initial = pd.DataFrame()

    # Add more data to the DataFrame
    new_data = {'이름': [name], '주민등록번호': [id_number], '휴대전화': [phone_number], '주소': [address], '시설':[sheet_name]}
    df = pd.concat([df_initial, pd.DataFrame(new_data)], ignore_index=True)

    # Create a Pandas Excel writer object using XlsxWriter as the engine
    writer = pd.ExcelWriter(existing_data_path, engine='xlsxwriter')

    # Write data to the Excel sheet
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Close the file
    writer.close()

def gick(name):
    pyautogui.click(x=608, y=493)
    time.sleep(0.5)
    pyautogui.click(x=786, y=391)
    pyperclip.copy(name)
    pyautogui.hotkey("ctrl",'v')
    time.sleep(0.2)
    pyautogui.press("f2")
    time.sleep(1)
    pyautogui.click(x=802, y=461, clicks=2, button='left')

def click_user_spe(data, t):
    num_list = [['1','6','7','9'],['3','5'],['2'],['8'],['4']]
    time.sleep(0.2)
    pyautogui.click(x=719, y=460)
    time.sleep(0.2)
    pyautogui.click(x=752, y=529)

def click_user_spe2(data, t):
    global type_list
    j = 0
    time.sleep(0.5)
    pyautogui.click(x=1001, y=466, clicks=1, button='left')
    time.sleep(0.5)
    if t == 1:
        for i in range(16):
            pyautogui.press('up')
    for k in range(9):
        pyautogui.press("down")
    pyautogui.press("enter")

    time.sleep(0.5) # 15

def edit_date(date, num):
    tem = date.split('.')
    a = int(tem[1])+num
    tem[1] = "0"+str(a)
    text = tem[0]+'.'+tem[1]+'.'+tem[2]
    return text

def write_special_note(data, date):
    global bin1
    pyautogui.click(x=999, y=777, clicks=2, button='left')
    time.sleep(0.5)
    
    is_bin = True
    bin = ""
    add = 0
    
    for i in bin1:
        if i == data:
            bin = "1순위"
            is_bin = False
            add = edit_date(date,9)
    
    for i in bin2:
        if i == data:
            bin = "2순위"
            is_bin = False
            add = edit_date(date,6)
            
    for i in bin3:
        if i == data:
            bin = "3순위"
            is_bin = False
            add = edit_date(date,3)

    if is_bin:
        bin = "3순위"

    text = "("+str(date)+" ~ "+add+") 마켓 "+bin+" 이용자"
    pyperclip.copy(text)

    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

def write_address(a, data):
    global name
    global id_number
    global address
    global phone_number
    pyautogui.click(x=424, y=524)
    time.sleep(1)
    pyautogui.click(x=861, y=437, clicks=1, button='left')
    pyperclip.copy(a)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.click(x=1132, y=437, clicks=1, button='left')
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
        time.sleep(0.5)
        return 1
    else:
        pyautogui.click(x=1000, y=773)
        user_names = data.iloc[active_num, name] + "(이미 등록됨)"
        user_id_number = data.iloc[active_num, id_number]
        user_phone_number = data.iloc[active_num, phone_number]
        user_address = data.iloc[active_num, address]
        add_data(user_names, user_id_number, user_phone_number, user_address, sheet_name)
        return -1


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

# restart
def restart_sign_new_user(sheet_name):
    global data
    global active_num
    global bin1
    global bin2
    global bin3
    global name
    global id_number
    global address
    global phone_number
    num = int(active_num)
    user_name = str(data.iloc[active_num, name])
    
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
    # sex_num1, sex_num2 = str(data.loc[num].iloc[id_number]).split('-')
    sex_num1, sex_num2 = str(data.iloc[active_num, id_number]).split("-")
    if(sex_num2[0] == '1' or sex_num2[0] == '2'):
        print("19")
        time.sleep(0.1)
        pyautogui.press('1')
        time.sleep(0.1)
        pyautogui.press('9')
    elif(sex_num2[0] == '3' or sex_num2[0] == '4'):
        print("20")
        time.sleep(0.1)
        pyautogui.press('2')
        time.sleep(0.1)
        pyautogui.press('0')
    else:
        print("age error")
        return
    time.sleep(0.1)
    pyperclip.copy(str(sex_num1))
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
        if time.time() - start_time >= 4:
            break
    
    if is_data:
        pyautogui.click(x=1001, y=507, clicks=2, button='left')
        time.sleep(1)
        pyautogui.press("esc")
        time.sleep(0.5)
        pyautogui.press("esc")

        # 이용자명 입력(mouse_pos = 408, 432)
        pyautogui.click(x=489, y=431, clicks=1, button='left')
        user_name = str(data.iloc[active_num, name])
        pyperclip.copy(user_name)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        # 이용자 구분, 이용자발굴지 선택(mouse_pos1 = 690, 434), 이용자분류 선택(mouse_pos2 = 1050, 458) 2*16
        pyautogui.click(x=690, y=448, clicks=1, button='left')
        global type_list

        time.sleep(0.5)
        
        # 기관 입력
        gick(sheet_name)
        
        # 주소 입력(mouse_pos1 = 427, 523)(mouse_pos2 = 861, 430)(mouse_pos3 = 1132, 442)(mouse_pos4 = 948, 521)
        ad = write_address(data.iloc[active_num, address], data)

        # 번호 입력(mouse_pos1 = 395, 561)(mouse_pos2 = 384, 603)
        user_num = str(data.iloc[active_num, phone_number]).split('-')
        if(user_num[0] == "010"):
            pyautogui.click(x=400, y=558, clicks=1, button='left')
            pyautogui.click(x=384, y=601, clicks=1, button='left')
        else:
            pyautogui.click(x=400, y=558, clicks=1, button='left')
            for i in range(23):
                pyautogui.press('up')
            for i in range(23):
                pyautogui.press("down")
            pyautogui.press("enter")
        pyautogui.press("tab")
        pyperclip.copy(user_num[1])
        pyautogui.hotkey('ctrl', 'v')
        pyperclip.copy(user_num[2])
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)        

        # 지원기간 선택(mouse_pos1 = 1035, 587)
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
        
        # # 비고
        # pyautogui.click(x=630, y=824)
        # time.sleep(0.5)
        # pyperclip.copy('2024 설맞이 꾸러미 수령')
        # pyautogui.hotkey('ctrl', 'v')
        # time.sleep(0.5)

        global type_num
        # Your code to execute when 'Yes' is clicked
        pyautogui.click(x=1735, y=294)
        pyautogui.press('enter')
        time.sleep(6)
        pyautogui.press('enter')
        pyautogui.click(x=1791, y=292)
    else:
        user_names = data.iloc[active_num, name]
        user_id_number = data.iloc[active_num, id_number]
        user_phone_number = data.iloc[active_num, phone_number]
        user_address = data.iloc[active_num, address]
        add_data(user_names, user_id_number, user_phone_number, user_address, sheet_name)
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
    global bin1
    global bin2
    global bin3
    global name
    global id_number
    global address
    global phone_number
    
    num = int(active_num)

    # ID자동생성 체크박스 클릭(mouse_pos = 456, 340)
    pyautogui.moveTo(x=457, y=342)
    pyautogui.click(clicks=1, button='left')
    time.sleep(0.5)

    # 이용자명 입력(mouse_pos = 408, 432)
    pyautogui.click(x=461, y=431, clicks=1, button='left')
    user_name = str(user_data.iloc[active_num, name])
    pyperclip.copy(user_name)

    pyautogui.hotkey('ctrl', 'v')
    
    time.sleep(0.5)
    #주민등록 번호 입력(mouse_pos1 = 665, 434)
    pyautogui.click(x=665, y=431, clicks=1, button='left')
    # user_num1, user_num2 = str(user_data.iloc[num, id_number]).split('-')
    user_num1, user_num2 = str(user_data.iloc[active_num, id_number]).split("-")
    
    pyperclip.copy(user_num1)
    pyautogui.hotkey('ctrl', 'v')
    pyperclip.copy(user_num2)
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
    
    gick(sheet_name)

    # 이용자 구분, 이용자발굴지 선택(mouse_pos1 = 690, 434), 이용자분류 선택(mouse_pos2 = 1050, 458) 2*16
    global type_list
    #user_type1 
    click_user_spe('hi', 0)

    #user_type3
    click_user_spe2('hi', 0)
    
    # 주소 입력(mouse_pos1 = 427, 523)(mouse_pos2 = 861, 430)(mouse_pos3 = 1132, 442)(mouse_pos4 = 948, 521)
    ad = write_address(user_data.iloc[active_num, address], user_data)

    if ad == -1:
        pyautogui.click(x=1729, y=291)
        pyautogui.click(x=1806, y=295)
        return
    
    # 번호 입력(mouse_pos1 = 395, 561)(mouse_pos2 = 384, 603)
    user_num = str(user_data.iloc[active_num, phone_number]).split('-')
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
            pyautogui.press('up')
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

    # # 비고
    # pyautogui.click(x=630, y=824)
    # time.sleep(0.5)
    # pyautogui.press('right')
    # time.sleep(0.5)
    # pyperclip.copy('2024 설맞이 꾸러미 수령')
    # pyautogui.hotkey('ctrl', 'v')
    # time.sleep(0.5)

    message_box()

## mai박금순n
counting = 19
# da이름te_input = input("날짜(ex.2024.1.1): ")
date_input = "123"
sheet_name = "예담" #input("시트 이름: ")익산시 오산면 영성길 18
select_file(sheet_name)
pyautogui.hotkey('alt','tab')

for i in range(counting):
    time.sleep(1)
    sign_new_user(data, date_input, sheet_name)
    active_num += 1
    time.sleep(1)