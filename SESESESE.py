import pytesseract as tr
import cv2
from PIL import ImageGrab
import re
import pyautogui as pa
import time
import keyboard
import pandas as pd

left = 365
top = 273
right = 416
bottom = 295

y_p = 22

print("파일 불러오기")
sh_name = input('시트 이름: ')
active_num = 0


try:
    df = pd.read_excel("test.xlsx", sheet_name=sh_name)
    print(df)
except:
    print("시트 이름이 잘못되었습니다.")
    exit()
    
start = input("시작하시겠습니까? (y/n): ")
if start == 'n':
    exit()
else:
    print("시작합니다.")
    time.sleep(1)

count = 0
pa.hotkey('alt', 'tab')

while True:
    # 물품 선택, 이용자 검색


    img = ImageGrab.grab((left, top, right, bottom))
    img.save("screenshot.png")

    image = cv2.imread("screenshot.png")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    text = tr.image_to_string(gray, lang='eng')

    # Use a regular expression to extract numbers from the text
    numbers = re.findall(r'\d+', text)
    print(numbers)

    try:
        num = str(numbers[0])
    except:
        print('over')
    
    # Print the extracted numbers
    print(num)
    
    if num == '' or num == None:
        break

    #Point(x=1492, y=460) 1425,465, 1515,488
    #Point(x=1577, y=485)
    # 460 + 25 y축
    is_data = True
    x1, x2, y1, y2 = 1425, 1515, 465, 488
    #when click the esc, stop this while looop
    if keyboard.is_pressed('esc'):
        print("Exiting program.")
        break
    while is_data:
        time.sleep(0.5)
        y_plus = 23

        img = ImageGrab.grab((x1, y1, x2, y2))
        img.save("test1.png")

        image = cv2.imread("test1.png")
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        text = tr.image_to_string(gray, lang='eng')

        # 이미지 이진화 (Binary Thresholding)
        ret, binary_image = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)

        # 노이즈 제거
        denoised_image = cv2.fastNlMeansDenoising(binary_image, None, h=10)

        text = tr.image_to_string(denoised_image, lang='eng', config='--psm 6')
        
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
        try:
            if num == all_number:
                print(num, all_number)
                on_num = True
                is_data = False
        except:
            print('no user')
            on_num = False
            break
        
        y1 += y_plus
        y2 += y_plus
        
        #when click the esc, stop this while looop
        if keyboard.is_pressed('esc'):
            print("Exiting program.")
            on_num = False
            break
        
    #when click the esc, stop this while looop
    if keyboard.is_pressed('esc'):
        print("Exiting program.")
        on_num = False
        break
    
    if on_num:
        time.sleep(0.2)
        pa.click(x=1310, y=y1-15, clicks=2, button='left')
        time.sleep(0.2)
        pa.click(x=1174, y=924)
        time.sleep(0.2)
        pa.click(18, top+12)
        pa.click(18, top+12)
        time.sleep(0.2)
        pa.hotkey('ctrl', 'c')
        pa.click(x=1331, y=474)
        time.sleep(0.2)
        pa.hotkey('ctrl', 'v')
        time.sleep(0.2)
        pa.click(18, top+12)
        pa.click(18, top+12)
        time.sleep(0.2)
        pa.press('right')
        pa.press('right')
        pa.hotkey('ctrl', 'c')
        time.sleep(0.2)
        pa.click(x=1685, y=473)
        time.sleep(0.2)
        pa.press('delete')
        pa.press('delete')
        pa.press('delete')
        time.sleep(0.2)
        pa.hotkey('ctrl', 'v')
        pa.click(1869, 290)
        time.sleep(0.2)
        pa.press('enter')
        time.sleep(0.2)
        pa.press('enter')
        time.sleep(0.2)
        pa.press('enter')
        time.sleep(0.2)
        pa.press('enter')
        time.sleep(0.2)
        pa.click(13, top+12)
        time.sleep(0.2)
        pa.click(13, top+12)
        time.sleep(0.2)
        pa.click(13, top+12, clicks=1, button='right')
        time.sleep(0.2)
        pa.press('d')
        
        #when click the esc, stop this while looop
        if keyboard.is_pressed('esc'):
            print("Exiting program.")
            break
        
        
    else:
        print("no")
        top = top + y_p
        bottom = bottom + y_p
        pa.click(x=1215, y=918)
        time.sleep(0.2)
        pa.click(x=744, y=497)
        time.sleep(0.2)
        continue