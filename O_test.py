import pytesseract as tr
import cv2
from PIL import ImageGrab
import re
import pyautogui
import time
import keyboard

# 오뚜기 테스트 

pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.keyUp('alt')

time.sleep(0.2)

pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.keyUp('alt')

left = 365
top = 273
right = 416
bottom = 295

y_p = 22

while True:
    on_num = False
    
    time.sleep(0.2)
    pyautogui.click(x=746, y=498)
    time.sleep(0.2)
    pyautogui.click(1868, 419)
    time.sleep(0.2)
    pyautogui.click(18, top+12)
    time.sleep(0.5)
    pyautogui.click(18, top+12)
    time.sleep(0.5)
    pyautogui.press('right')
    pyautogui.press('right')
    pyautogui.press('right')
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.5)
    pyautogui.click(1024, 329)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('f2')

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
        pyautogui.click(x=1310, y=y1-15, clicks=2, button='left')
        time.sleep(0.2)
        pyautogui.click(x=1174, y=924)
        time.sleep(0.2)
        pyautogui.click(18, top+12)
        pyautogui.click(18, top+12)
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'c')
        pyautogui.click(x=1331, y=474)
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.2)
        pyautogui.click(18, top+12)
        pyautogui.click(18, top+12)
        time.sleep(0.2)
        pyautogui.press('right')
        pyautogui.press('right')
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.2)
        pyautogui.click(x=1685, y=473)
        time.sleep(0.2)
        pyautogui.press('delete')
        pyautogui.press('delete')
        pyautogui.press('delete')
        time.sleep(0.2)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.click(1869, 290)
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(0.2)
        pyautogui.click(13, top+12)
        time.sleep(0.2)
        pyautogui.click(13, top+12)
        time.sleep(0.2)
        pyautogui.click(13, top+12, clicks=1, button='right')
        time.sleep(0.2)
        pyautogui.press('d')
        
        #when click the esc, stop this while looop
        if keyboard.is_pressed('esc'):
            print("Exiting program.")
            break
        
        
    else:
        print("no")
        top = top + y_p
        bottom = bottom + y_p
        pyautogui.click(x=1215, y=918)
        time.sleep(0.2)
        pyautogui.click(x=744, y=497)
        time.sleep(0.2)
        continue