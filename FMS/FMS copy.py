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
# 안길동

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
class user_sign(object):
    #File I/O path load
    active_num = 2
    path = "out.xlsx"
    do_list=["이용자 등록","접수등록","제공등록","접수목록 찾기",
                "제공목록 찾기","접수현황 수정","제공현황 수정"]
    date = datetime.date.today()
    data = ''
    type_num = 1


    #log
    # def log_user(self, message):
    #     log_text.delete(1.0, tk.END)
    #     num_range = ['4','5','6','7','9','10','13']
    #     print_data = message.iloc[active_num]
    #     #4,5,6,7,9,10,11,13
    #     log_text.insert(tk.END,'이       름: '+message.iloc[active_num,4]+'\n\n')
    #     log_text.insert(tk.END,'주 민 번 호: '+message.iloc[active_num,5]+'\n\n')
    #     log_text.insert(tk.END,'주       소: '+message.iloc[active_num,6]+'\n\n')
    #     log_text.insert(tk.END,'전 화 번 호: '+message.iloc[active_num,7]+'\n\n')
    #     log_text.insert(tk.END,'이용자 구분: '+message.iloc[active_num,9]+'\n\n')
    #     log_text.insert(tk.END,'이용자 특성: '+message.iloc[active_num,10]+'\n\n')
    #     log_text.insert(tk.END,'신 청 구 분: '+message.iloc[active_num,13]+'\n\n')
            

    #log
    # def log(self, message, log_text):
    #     log_text.delete(1.0, tk.END)
    #     log_text.insert(tk.END, message)
        #4,5,6,7,9,10,11,13

    #엑셀 데이터 1행 출력
    def data_print(self, _data, num, type_num):
        global active_num
        active_num += int(num)
        print(active_num)

    #엑셀 불러오기
    def get_data_from_excel(self, _path, date):
        global active_num
        
        # 엑셀 파일을 데이터프레임으로 읽어옵니다.
        df = pd.read_excel(_path, sheet_name=str(date))

        # data return
        return df

    #엑셀 path select
    def select_file(self):
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

    # 2. 원하는 작업 선택
    # 작업 종류: 이용자 등록, 접수등록, 제공등록, 접수목록 찾기, 제공목록 찾기
    #           접수현황 수정, 제공현황 수정
    # 작업방식 선택
    # a) 매크로 (빠르게 구현가능 하고 쉽게 이해가능, 정확도 down)
    # b) 소스코딩 (정확도 up, 이해하기 어렵고 구현난이도 높음)
    # 완전 자동화 x, 사용자가 넘어가기 버튼 클릭 시 다음 작업 실행
    # 23.10.20 작업방식 매크로로 전체 통일 추후 소스코딩으로 변경

    # restart
    def restart_sign_new_user(self):
        global data
        global active_num
        num = int(active_num)
        user_name = str(data.iloc[active_num,4])
        x1, y1, x2, y2 = 700, 500, 780, 515
        y_plus = 28
        is_data = True
        i = 1
        is_match = False
        
        pyautogui.click(1720, 295)
        time.sleep(1)
        pyautogui.click(951,340)
        pyperclip.copy(user_name)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press("f2")
        time.sleep(0.5)

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
                is_data = False
                return
                
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

        x1, y1, x2, y2 = 700, 500, 780, 515
        i = 1

        if is_data:
            time.sleep(0.5)
            start_time = time.time()
            green_color = (51, 96, 52)  # (R, G, B) values for pure green
            while True:
                # Take a screenshot of the entire screen
                screenshot = pyautogui.screenshot()
                green_found = False  # Flag to check if green color is found
                # Search for the green color in the screenshot
                for x in range(screenshot.width):
                    for y in range(screenshot.height):
                        pixel_color = screenshot.getpixel((1109, 218))
                        if pixel_color == green_color:
                            # Green color found, click a button (you can modify this action)
                            pyautogui.click(1111, 327)
                            print("Green color found at ({}, {})".format(1111, 327))
                            green_found = True
                            is_match = True
                            break  # Exit the inner loop

                    if green_found:
                        break  # Exit the outer loop

                # Check if 3 seconds have passed
                if time.time() - start_time >= 2:
                    break

            if is_match:
                time.sleep(0.5)
                pyautogui.press("enter")
            
            # ID자동생성 체크박스 클릭(mouse_pos = 456, 340)
            pyautogui.moveTo(456,340)
            pyautogui.click(clicks=1, button='left')
            time.sleep(0.5)

            # 이용자명 입력(mouse_pos = 408, 432)
            pyautogui.click(x=408, y=432, clicks=1, button='left')
            is_spe = True
            user_name = str(data.iloc[num, 4])
            try:
                sum1 = str(data.iloc[num, 13]).split('+')
                pyperclip.copy("＊"+user_name)
            except:
                is_spe = False
                pyperclip.copy(user_name)

            pyautogui.hotkey('ctrl', 'v')

            time.sleep(0.5)
            if is_spe:
                pyperclip.copy("("+str(date)+") 나눔곳간 1차 이용")
            else:
                pyperclip.copy("("+str(date)+") 행복나눔마켓 1차 이용")

            # 이용자 구분, 이용자발굴지 선택(mouse_pos1 = 690, 434), 이용자분류 선택(mouse_pos2 = 1050, 458) 2*16
            pyautogui.click(x=690, y=434, clicks=1, button='left')
            user_type1 = str(data.iloc[num, 9]).split('.')
            user_type3 = str(data.iloc[num, 10]).split('.')
            
            type_list = ["결식아동","다문화가정","독거어르신","소년소녀가장",
                            "외국인노동자","재가장애인","저소득가정","조손가정",
                            "한부모가정","기타","청장년 1인가구","미혼모부가구",
                            "부부중심가구","노인부부가구","새터민가구","공통체가구"]
            
            user_type_list = ["긴급 위기상황 발생자","수급자","차상위계층","소득감소자","복지급여 탈락"]

            #user_type1
            j = 0
            for i in user_type_list:
                if i == user_type1[1]:
                    pyautogui.click(x=459, y=460, clicks=1, button='left')
                    for z in range(3):
                        pyautogui.press("up")
                    for k in range(j):
                        pyautogui.press("down")
                    pyautogui.press("enter")
                    break
                else:
                    j+=1

                

            #user_type2
            if j == 0:
                pyautogui.click(x=722, y=459, clicks=1, button='left')
                pyautogui.press("down")
                pyautogui.press("down")
                pyautogui.press("down")
                pyautogui.press("enter")
            

            #user_type3
            j = 0
            for i in type_list:
                if i == user_type3[1]:
                    pyautogui.click(x=1050, y=458, clicks=1, button='left')
                    for k in range(15):
                        pyautogui.press("up")
                    for k in range(j):
                        pyautogui.press("down")
                    pyautogui.press("enter")
                    break
                j += 1

            time.sleep(0.5)
            
            # 주소 입력(mouse_pos1 = 427, 523)(mouse_pos2 = 861, 430)(mouse_pos3 = 1132, 442)(mouse_pos4 = 948, 521)
            pyautogui.click(x=427, y=523, clicks=1, button='left')
            time.sleep(1)
            pyautogui.click(x=861, y=430, clicks=1, button='left')
            pyperclip.copy(data.iloc[num, 6])
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.click(x=1132, y=442, clicks=1, button='left')
            time.sleep(1)
            pyautogui.click(x=948, y=521, clicks=2, button='left')
            time.sleep(1)

            # 번호 입력(mouse_pos1 = 395, 561)(mouse_pos2 = 384, 603)
            pyautogui.click(x=400, y=554, clicks=1, button='left')
            pyautogui.click(x=384, y=603, clicks=1, button='left')
            user_num = str(data.iloc[num, 7]).split('-')
            pyautogui.click(x=437, y=556, clicks=1, button='left')
            pyperclip.copy(user_num[1])
            pyautogui.hotkey('ctrl', 'v')
            pyperclip.copy(user_num[2])
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)

            # 지원기간 선택(mouse_pos1 = 1035, 587)
            pyautogui.click(x=1008, y=590, clicks=1, button='left')
            for i in range(2):
                pyautogui.press('up')
            pyautogui.click(x=1035, y=587, clicks=1, button='left')
            for i in range(13):
                pyautogui.press('up')
            for i in range(12):
                pyautogui.press('down')
            pyautogui.press('enter')
            time.sleep(0.5)

            # 신청구분 특이사항 입력(date + 이용기관 + '1차 이용')(mouse_pos1 = 435, 718)
            pyautogui.click(x=532, y=728, clicks=2, button='left')
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
                pyautogui.click(1745, 294)
                pyautogui.press('enter')
                time.sleep(2)
                pyautogui.press('enter')
                pyautogui.click(1808,289)
                data_print(data,1, type_num)
                pyautogui.hotkey('alt','tab')
            else:
                on_click_no()
        else:
            messagebox.showinfo("error","?")


    # Function to handle 'Yes' button click
    def on_click_yes(self):
        global type_num
        # Your code to execute when 'Yes' is clicked
        messagebox.showinfo("안내","이용자 정보를 저장합니다.'")
        pyautogui.click(1786, 294)
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('enter')
        pyautogui.click(1808,289)
        data_print(data,1, type_num)
        pyautogui.hotkey('alt','tab')

    # Function to handle 'No' button click
    def on_click_no(self):
        # Your code to execute when 'No' is clicked
        messagebox.showinfo("Info", "You clicked 'No'")

    # Function to display the message box
    def message_box(self):
        response = messagebox.askquestion("확인","모든 정보가 정확합니까?")
        if response == "yes":
            on_click_yes()
        else:
            on_click_no()

    # 이용자 등록
    def sign_new_user(self, user_data, date):
        global active_num
        num = int(active_num)

        # ID자동생성 체크박스 클릭(mouse_pos = 456, 340)
        pyautogui.hotkey('alt','tab')
        pyautogui.moveTo(456,340)
        pyautogui.click(clicks=1, button='left')
        time.sleep(0.5)

        # 이용자명 입력(mouse_pos = 408, 432)
        pyautogui.click(x=408, y=432, clicks=1, button='left')
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
        pyautogui.click(x=665, y=434, clicks=1, button='left')
        user_num = str(user_data.iloc[num, 5]).split('-')
        pyperclip.copy(user_num[0])
        pyautogui.hotkey('ctrl', 'v')
        pyperclip.copy(user_num[1])
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        is_match = False
        #Point(x=1144, y=212)
        start_time = time.time()
        green_color = (51, 96, 52)  # (R, G, B) values for pure green
        while True:
            # Take a screenshot of the entire screen
            screenshot = pyautogui.screenshot()

            green_found = False  # Flag to check if green color is found

            # Search for the green color in the screenshot
            for x in range(screenshot.width):
                for y in range(screenshot.height):
                    pixel_color = screenshot.getpixel((1109, 218))
                    if pixel_color == green_color:
                        # Green color found, click a button (you can modify this action)
                        pyautogui.click(1109, 218)
                        print("Green color found at ({}, {})".format(1109, 218))
                        green_found = True
                        is_match = True
                        break  # Exit the inner loop

                if green_found:
                    break  # Exit the outer loop

            # Check if 3 seconds have passed
            if time.time() - start_time >= 2:
                break

        if is_match:
            time.sleep(1)
            pyautogui.press('enter')
            restart_sign_new_user()
            return
        

        # 이용자 구분, 이용자발굴지 선택(mouse_pos1 = 690, 434), 이용자분류 선택(mouse_pos2 = 1050, 458) 2*16
        pyautogui.click(x=690, y=434, clicks=1, button='left')
        user_type1 = str(user_data.iloc[num, 9]).split('.')
        user_type3 = str(user_data.iloc[num, 10]).split('.')
        
        type_list = ["결식아동","다문화가정","독거어르신","소년소녀가장",
                        "외국인노동자","재가장애인","저소득가정","조손가정",
                        "한부모가정","기타","청장년 1인가구","미혼모부가구",
                        "부부중심가구","노인부부가구","새터민가구","공통체가구"]
        
        user_type_list = ["긴급 위기상황 발생자","수급자","차상위계층","소득감소자","복지급여 탈락"]

        #user_type1
        j = 0
        for i in user_type_list:
            if i == user_type1[1]:
                pyautogui.click(x=459, y=460, clicks=1, button='left')
                for k in range(j):
                    pyautogui.press("down")
                pyautogui.press("enter")
                break
            else:
                j+=1
            
            if j >= 5:
                messagebox.showinfo("error","엑셀파일의 이용자특성 및 구분 확인")
                break
                
        #user_type2
        if j == 0:
            pyautogui.click(x=722, y=459, clicks=1, button='left')
            pyautogui.press("down")
            pyautogui.press("down")
            pyautogui.press("down")
            pyautogui.press("enter")
        else:
            pyautogui.click(x=722, y=459, clicks=1, button='left')

        #user_type3
        j = 0
        for i in type_list:
            if i == user_type3[1]:
                pyautogui.click(x=1050, y=458, clicks=1, button='left')
                for k in range(j):
                    pyautogui.press("down")
                pyautogui.press("enter")
                break
            j += 1

        time.sleep(0.5)
        
        # 주소 입력(mouse_pos1 = 427, 523)(mouse_pos2 = 861, 430)(mouse_pos3 = 1132, 442)(mouse_pos4 = 948, 521)
        pyautogui.click(x=427, y=523, clicks=1, button='left')
        time.sleep(1)
        pyautogui.click(x=861, y=430, clicks=1, button='left')
        pyperclip.copy(user_data.iloc[num, 6])
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.click(x=1132, y=442, clicks=1, button='left')
        time.sleep(1)
        pyautogui.click(x=948, y=521, clicks=2, button='left')
        time.sleep(1)

        # 번호 입력(mouse_pos1 = 395, 561)(mouse_pos2 = 384, 603)
        pyautogui.click(x=400, y=554, clicks=1, button='left')
        pyautogui.click(x=384, y=603, clicks=1, button='left')
        user_num = str(user_data.iloc[num, 7]).split('-')
        pyautogui.click(x=437, y=556, clicks=1, button='left')
        pyperclip.copy(user_num[1])
        pyautogui.hotkey('ctrl', 'v')
        pyperclip.copy(user_num[2])
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        # 지원기간 선택(mouse_pos1 = 1035, 587)
        pyautogui.click(x=1035, y=587, clicks=1, button='left')
        for i in range(12):
            pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(0.5)

        # 신청구분 특이사항 입력(date + 이용기관 + '1차 이용')(mouse_pos1 = 435, 718)
        pyautogui.click(x=532, y=728, clicks=2, button='left')
        time.sleep(0.5)

        if is_spe:
            pyperclip.copy("("+str(date)+") 나눔곳간 1차 이용")
        else:
            pyperclip.copy("("+str(date)+") 행복나눔마켓 1차 이용")

        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        pyautogui.hotkey('alt', 'tab')
        message_box()

    # start
    def start_bt(self, combo):
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
    def get_date(self):
        global date
        # Get the date input from the GUI
        input_date = simpledialog.askstring("Date Input", "Enter a date (X월X일):")
        date = input_date
