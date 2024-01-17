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

# 매크로 내용
# 1. 받아온 데이터를 웹사이트에 입력
# 2. 웹사이트의 저장 버튼을 누름
# 3. 다음 데이터를 입력하기 전에 3초간 대기
# 4. 반복
# 강신순

# background exe program

# 1. 엑셀 파일에서 데이터를 받아옴
# 엑셀 파일에서 데이터를 받아오는 함수
# sort data columns
# 품     명
# 수량
# 거래
# 생년월일
# 각 Columns별 정렬 가능
# 받아온 sorted data 출력

#File I/O path load
active_num = 2
path = "out.xlsx"
do_list=["이용자 등록","접수등록","제공등록","접수목록 찾기",
            "제공목록 찾기","접수현황 수정","제공현황 수정"]
date_to = datetime.date.today()
date = str(date_to.month)+"월"+str(date_to.day)+"일"
data = ''
type_num = 1


#log
def log_user(message):
    log_text.delete(1.0, tk.END)
    num_range = ['4','5','6','7','9','10','13']
    print_data = message.iloc[active_num]
    #4,5,6,7,9,10,11,13
    log_text.insert(tk.END,'이       름: '+str(message.iloc[active_num,4])+'\n\n')
    log_text.insert(tk.END,'주 민 번 호: '+str(message.iloc[active_num,5])+'\n\n')
    log_text.insert(tk.END,'주       소: '+str(message.iloc[active_num,6])+'\n\n')
    log_text.insert(tk.END,'전 화 번 호: '+str(message.iloc[active_num,7])+'\n\n')
    log_text.insert(tk.END,'이용자 구분: '+str(message.iloc[active_num,9])+'\n\n')
    log_text.insert(tk.END,'이용자 특성: '+str(message.iloc[active_num,10])+'\n\n')
    log_text.insert(tk.END,'신 청 구 분: '+str(message.iloc[active_num,13])+'\n\n')
        

#log
def log(message):
    log_text.delete(1.0, tk.END)
    log_text.insert(tk.END, message)
    #4,5,6,7,9,10,11,13

#엑셀 데이터 1행 출력
def data_print(_data, num, type_num):
    global active_num
    active_num += int(num)
    print(active_num)

    try:
        if type_num == 1:
            log_user(_data)
        else:
            log(str(_data.iloc[active_num,type_num]))
    except:
        log("ERROR: 파일의 데이터 형식이 알맞지 않습니다."+"\n"+
            "다음 내용을 확인해주세요."+"\n"+
            "1. 이름에 공백 혹은 다른 문자가 들어있지 않은지 확인해주세요."+"\n"+
            "2. 주민등록번호가 -로 구분되어 있는지 확인해주세요."+"\n"+
            "3. 주소가 정확한지 확인해주세요."+"\n"+
            "4. 전화번호가 -로 구분되어 있는지 확인해주세요."+"\n"+
            "5. 이용자 구분이 정확한지 확인해주세요."+"\n"+
            "6. 이용자 특성이 정확한지 확인해주세요."+"\n"+
            "7. 데이터가 없습니다.")
        traceback.print_exc()   
    
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

# 2. 원하는 작업 선택
# 작업 종류: 이용자 등록, 접수등록, 제공등록, 접수목록 찾기, 제공목록 찾기
#           접수현황 수정, 제공현황 수정
# 작업방식 선택
# a) 매크로 (빠르게 구현가능 하고 쉽게 이해가능, 정확도 down)
# b) 소스코딩 (정확도 up, 이해하기 어렵고 구현난이도 높음)
# 완전 자동화 x, 사용자가 넘어가기 버튼 클릭 시 다음 작업 실행
# 23.10.20 작업방식 매크로로 전체 통일 추후 소스코딩으로 변경

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
    
    pyautogui.click(x=1736, y=295)
    time.sleep(1)
    pyautogui.click(x=969, y=340)
    pyperclip.copy(user_name)
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
        is_spe = True
        user_name = str(data.iloc[num, 4])
        if str(data.iloc[num, 13]).find("+") != -1:
            pyperclip.copy("＊"+user_name)
        else:
            is_spe = False
            pyperclip.copy(user_name)

        pyautogui.hotkey('ctrl', 'v')

        time.sleep(0.5)
        if is_spe:
            pyperclip.copy("("+str(date)+") 나눔곳간 1차 이용")
        else:
            pyperclip.copy("("+str(date)+") 행복나눔마켓 1차 이용")

        # 이용자 구분, 이용자발굴지 선택(mouse_pos1 = 690, 434), 이용자분류 선택(mouse_pos2 = 1050, 458) 2*16
        user_type1 = str(data.iloc[num, 9]).split('.')
        user_type3 = str(data.iloc[num, 10]).split('.')
    
        pyautogui.click(x=690, y=448, clicks=1, button='left')
        type_list = ["결식아동","다문화가정","독거어르신","소년소녀가장",
                    "외국인노동자","재가장애인","저소득가정","조손가정",
                    "한부모가정","기타","청장년 1인가구","미혼모부가구",
                    "부부중심가구","노인부부가구","새터민가구","공통체가구"]
    
        user_type_list = ["긴급 위기상황 발생자","수급자","차상위계층",
                            "소득감소자","복지급여 탈락", "실직 및 휴폐업", "읍면동장추천"]
        #user_type1 
        j = 0
        pyautogui.click(x=459, y=463, clicks=1, button='left')
        time.sleep(0.5)
        for i in user_type_list:
            print("lkasjf       "+i)
            if i == user_type1[1]:
                pyautogui.press("up")
                pyautogui.press("up")
                pyautogui.press("up")
                pyautogui.press("up")
                if j <= 0:
                    break
                for k in range(j):
                    pyautogui.press("down")
                pyautogui.press("enter")
                break
            else:
                j+=1
                if j > 4:
                    pyautogui.press("up")
                    pyautogui.press("enter")
                break
        time.sleep(0.5)
        #user_type2
        pyautogui.click(x=722, y=463, clicks=1, button='left')
        time.sleep(0.5)
        print(j)
        if j <= 0:
            pyautogui.press("down")
            pyautogui.press("down")
            pyautogui.press("down")
            pyautogui.press("enter") 
        

        #user_type3
        j = 0
        time.sleep(0.5)
        pyautogui.click(x=1050, y=463, clicks=1, button='left')
        time.sleep(0.5)
        for i in type_list:
            if i == user_type3[1]:
                print(user_type3)
                for k in range(15):
                    pyautogui.press("up")
                for k in range(j):
                    pyautogui.press("down")
                pyautogui.press("enter")
                break
            j += 1


        time.sleep(0.5)
        
        # 주소 입력(mouse_pos1 = 427, 523)(mouse_pos2 = 861, 430)(mouse_pos3 = 1132, 442)(mouse_pos4 = 948, 521)
        pyautogui.click(x=425, y=522, clicks=1, button='left')
        time.sleep(1)
        pyautogui.click(x=873, y=434, clicks=1, button='left')
        pyperclip.copy(data.iloc[num, 6])
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.click(x=1132, y=441, clicks=1, button='left')
        time.sleep(1)
        pyautogui.click(x=948, y=525, clicks=2, button='left')
        time.sleep(1)

        # 번호 입력(mouse_pos1 = 395, 561)(mouse_pos2 = 384, 603)
        pyautogui.click(x=400, y=558, clicks=1, button='left')
        pyautogui.click(x=384, y=601, clicks=1, button='left')
        user_num = str(data.iloc[num, 7]).split('-')
        pyautogui.press("tab")
        pyperclip.copy(user_num[1])
        pyautogui.hotkey('ctrl', 'v')
        pyperclip.copy(user_num[2])
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        
        # 가구 인원수
        

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

        # 신청구분 특이사항 입력(date + 이용기관 + '1차 이용')(mouse_pos1 = 435, 718)
        pyautogui.click(x=925, y=755, clicks=2, button='left')
        time.sleep(0.5)
        pyautogui.press("right")
        pyautogui.press("enter")
        time.sleep(0.5)

        if is_spe:
            pyperclip.copy("("+str(date)+") 나눔곳간 1차 이용")
        else:
            pyperclip.copy("("+str(date)+") 행복나눔마켓 1차 이용")

        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.hotkey('alt','tab')

        global type_num
        # Your code to execute when 'Yes' is clicked
        response = messagebox.askquestion("확인","모든 정보가 정확합니까?")
        if response == "yes":
            messagebox.showinfo("안내","이용자 정보를 저장합니다.'")
            pyautogui.click(x=1735, y=294)
            pyautogui.press('enter')
            time.sleep(3)
            pyautogui.press('enter')
            pyautogui.click(x=1791, y=292)
            data_print(data,1, type_num)
            pyautogui.hotkey('alt','tab')
        else:
            on_click_no()
    else:
        messagebox.showinfo("error","?")


# Function to handle 'Yes' button click
def on_click_yes():
    global type_num
    # Your code to execute when 'Yes' is clicked
    messagebox.showinfo("안내","이용자 정보를 저장합니다.'")
    pyautogui.click(x=1804, y=280)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('enter')
    pyautogui.click(x=1804, y=280)
    data_print(data,1, type_num)
    pyautogui.hotkey('alt','tab')

# Function to handle 'No' button click
def on_click_no():
    # Your code to execute when 'No' is clicked
    messagebox.showinfo("Info", "You clicked 'No'")

# Function to display the message box
def message_box():
    response = messagebox.askquestion("확인","모든 정보가 정확합니까?")
    if response == "yes":
        on_click_yes()
    else:
        on_click_no()

# 이용자 등록
def sign_new_user(user_data, date):
    global active_num
    num = int(active_num)

    # ID자동생성 체크박스 클릭(mouse_pos = 456, 340)
    pyautogui.hotkey('alt','tab')
    pyautogui.moveTo(x=457, y=342)
    pyautogui.click(clicks=1, button='left')
    time.sleep(0.5)

    # 이용자명 입력(mouse_pos = 408, 432)
    pyautogui.click(x=461, y=431, clicks=1, button='left')
    is_spe = True
    user_name = str(user_data.iloc[num, 4])
    try:
        sum1 = str(user_data.iloc[num, 13]).split('+')
        if sum1[1] == "곳간":
            is_spe = True
            pyperclip.copy("＊"+user_name)
        else:
            is_spe = False
            pyperclip.copy(user_name)
    except:
        is_spe = False
        pyperclip.copy(user_name)

    pyautogui.hotkey('ctrl', 'v')
    
    time.sleep(0.5)
    #주민등록 번호 입력(mouse_pos1 = 665, 434)
    pyautogui.click(x=665, y=431, clicks=1, button='left')
    user_num = str(user_data.iloc[num, 5]).split('-')
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
        restart_sign_new_user()
        return
    

    # 이용자 구분, 이용자발굴지 선택(mouse_pos1 = 690, 434), 이용자분류 선택(mouse_pos2 = 1050, 458) 2*16
    user_type1 = str(user_data.iloc[num, 9]).split('.')
    user_type3 = str(user_data.iloc[num, 10]).split('.')
    
    type_list = ["결식아동","다문화가정","독거어르신","소년소녀가장",
                    "외국인노동자","재가장애인","저소득가정","조손가정",
                    "한부모가정","기타","청장년 1인가구","미혼모부가구",
                    "부부중심가구","노인부부가구","새터민가구","공통체가구"]
    
    user_type_list = ["긴급 위기상황 발생자","수급자","차상위계층",
                          "소득감소자","복지급여 탈락", "실직 및 휴폐업", "읍면동장추천"]
    #user_type1
    j = 0
    pyautogui.click(x=459, y=463, clicks=1, button='left')
    time.sleep(0.5)
    for i in user_type_list:
        if i == user_type1[1]:
            if j <= 0 or j > 4:
                pyautogui.press("enter")
                break
            pyautogui.press("up")
            for k in range(j):
                pyautogui.press("down")
            pyautogui.press("enter")
            break
        elif j > 4:
            pyautogui.press("enter")
            break
        else:
            j+=1
    time.sleep(0.5)
    #user_type2
    pyautogui.click(x=722, y=463, clicks=1, button='left')
    time.sleep(0.5)
    if j <= 0:
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("enter") 
    

    #user_type3
    j = 0
    time.sleep(0.5)
    pyautogui.click(x=1050, y=463, clicks=1, button='left')
    time.sleep(0.5)
    for i in type_list:
        if i == user_type3[1]:
            print(user_type3)
            for k in range(15):
                pyautogui.press("up")
            for k in range(j):
                pyautogui.press("down")
            pyautogui.press("enter")
            break
        j += 1

    time.sleep(0.5)
    
    # 주소 입력(mouse_pos1 = 427, 523)(mouse_pos2 = 861, 430)(mouse_pos3 = 1132, 442)(mouse_pos4 = 948, 521)
    pyautogui.click(x=427, y=525, clicks=1, button='left')
    time.sleep(1)
    pyautogui.click(x=861, y=437, clicks=1, button='left')
    pyperclip.copy(user_data.iloc[num, 6])
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.click(x=1132, y=437, clicks=1, button='left')
    time.sleep(1)
    pyautogui.click(x=948, y=516, clicks=2, button='left')
    time.sleep(1)

    # 번호 입력(mouse_pos1 = 395, 561)(mouse_pos2 = 384, 603)
    pyautogui.click(x=400, y=557, clicks=1, button='left')
    pyautogui.click(x=384, y=603, clicks=1, button='left')
    user_num = str(user_data.iloc[num, 7]).split('-')
    pyautogui.click(x=437, y=557, clicks=1, button='left')
    pyperclip.copy(user_num[1])
    pyautogui.hotkey('ctrl', 'v')
    pyperclip.copy(user_num[2])
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    
    # 가구 인원수
    pyautogui.click(x=445, y=711, clicks=2, button='left')
    fm_num = str(user_data.iloc[num, 8])
    pyperclip.copy(fm_num)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('tab')
    pyperclip.copy('0')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # 지원기간 선택(mouse_pos1 = 1035, 587)
    pyautogui.click(x=1051, y=586, clicks=1, button='left')
    for i in range(12):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(0.5)

    # 신청구분 특이사항 입력(date + 이용기관 + '1차 이용')(mouse_pos1 = 435, 718)
    pyautogui.click(x=532, y=748, clicks=2, button='left')
    time.sleep(0.5)

    if is_spe:
        pyperclip.copy("("+str(date)+") 나눔곳간 1차 이용")
    else:
        pyperclip.copy("("+str(date)+") 행복나눔마켓 1차 이용")

    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    pyautogui.hotkey('alt', 'tab')
    message_box()

# 접수등록
def receipt_user():
    1+1

# 제공등록
def supply_user():
    1+1

# 접수목록 찾기
def search_receipt_user():
    1+1

# 제공목록 찾기
def search_supply_user():
    1+1

# 접수현황 수정
def edit_receipt_user():
    1+1

# 제공현황 수정
def edit_supply_user():
    1+1

# start
def start_bt(combo):
    global do_list
    global data
    global date
    global tpye_num
    print("시작")
    
    if combo == do_list[0]:
        thread = threading.Thread(target=sign_new_user(data, date))
        thread.start()
        tpye_num = 1
    else:
        print("error")


# Function to handle the entered date
def get_date():
    global date
    # Get the date input from the GUI
    input_date = simpledialog.askstring("Date Input", "Enter a date (X월X일):")

    # Parse the date and format it as desired
    try:
        result_label.config(text=f"선택된 날짜: {input_date}")
        date = input_date
    except ValueError:
        result_label.config(text="날짜가 올바르지 않습니다.")
        

#~#
window = tk.Tk()
window.title("FMS")
window.geometry("400x350")  # 창 크기 고정
print(active_num)

#python gui
button_frame = tk.Frame(window)
button_frame.pack()
#log_text = tk.Text(window, height=10, width=50)

#combobox
combobox = ttk.Combobox(button_frame)
combobox.config(height=7)
combobox.config(values=do_list)
combobox.config(state="readonly")
combobox.set("작업선택")

#button
path_selector = tk.Button(button_frame, text='파일 선택', command=lambda: select_file())
next_button = tk.Button(button_frame, text="다음", command=lambda: data_print(data,1, type_num))
back_button = tk.Button(button_frame, text="이전", command=lambda: data_print(data,-1, type_num))
start_button = tk.Button(button_frame, text="시작", command=lambda: start_bt(combobox.get()))
#exit_excel_button = tk.Button(button_frame, text="엑셀 종료", command=lambda: exit_excel())

# 설정 값을 입력받는 프레임 추가
settings_frame = tk.Frame(window)
settings_frame.pack()

#날짜
input_button = tk.Button(button_frame, text="날짜 입력", command=get_date)
input_button.pack(pady=10)

# 결과 텍스트를 표시할 레이블 생성
result_label = tk.Label(window, text="", font=("Helvetica", 12))
result_label.pack()

back_button.pack(side="left")
next_button.pack(side="left")
path_selector.pack(side="right")
combobox.pack(side="right")
start_button.pack(side="right")
#exit_excel_button.pack(side="right")

# 로그 창 (Text 위젯) 추가
log_text = tk.Text(window, height=18, width=45)
log_text.pack()

# GUI 실행
window.mainloop()

