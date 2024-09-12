import time
import pyautogui as pa
import pandas as pd
import pyperclip as pc
import cv2
import re
import pytesseract as tr
from PIL import ImageGrab

def add_data(name, id_number, sheet_name, existing_data_path='시설미등록이용자.xlsx'):
    try:
        # Read existing data from Excel file
        df_initial = pd.read_excel(existing_data_path)
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame
        df_initial = pd.DataFrame()

    # Add more data to the DataFrame
    new_data = {'이름': [name], '주민등록번호': [id_number], '시설':[sheet_name]}
    df = pd.concat([df_initial, pd.DataFrame(new_data)], ignore_index=True)

    # Create a Pandas Excel writer object using XlsxWriter as the engine
    writer = pd.ExcelWriter(existing_data_path, engine='xlsxwriter')

    # Write data to the Excel sheet
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Close the file
    writer.close()

def find_obj(x, y):
    start_time = time.time()
    green_color = (254, 230, 200)  # (R, G, B) values for pure green
    target_pixel = (x, y)
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
        return True
    else:
        return False

def find_user(num):
    y_plus = 23
    x1, x2, y1, y2 = 1188, 1276, 484, 507
    
    for i in range(10):
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

        try:
            if str(num) == str(all_number):
                print(num, all_number)
                print("찾았다")
                pa.click(x=912, y=y1+10, clicks=2, button='left')
                time.sleep(1)
                return True
            else:
                print("!!이용자 검색 실패!!")
                print(num, all_number)
                print("못찾음")
                y1 += y_plus
                y2 += y_plus
        except:
            print('no user', num)
            return False
    
    pa.click(x=987, y=919)
    return False    

file_path = "C:/Users/user/Desktop/지역아동센터 new 명단.xlsx"
# 파일 불러오기
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names
is_cols = False
delay = 0.5

pa.hotkey('alt', 'tab')

# 시트 개수 만큼 반복
for sheet in sheet_names:
    active_num = 0
    df = pd.read_excel(file_path, sheet_name=sheet)
    
    print(sheet)

    if sheet != "예담":
        continue
    
    # 기관 검색
    pc.copy(sheet)
    pa.click(x=804, y=345)
    pa.hotkey("ctrl", "v")
    pa.press("f2")
    time.sleep(delay)
    # 검색 안되면 그룹 추가
    response = find_obj(x=281, y=438)
    if response == False:
        pa.click(x=623, y=392)
        time.sleep(1)
        pa.click(x=371, y=448, clicks=2)
        # 기관명 입력
        time.sleep(1)
        pc.copy(sheet)
        pa.hotkey('ctrl', 'v')
        time.sleep(1)
        # 설명 24.01.01 ~ 24.12.31
        pa.click(x=518, y=446, clicks=2)
        time.sleep(1)
        pc.copy("24.04.29 ~ 25.04.29")
        pa.hotkey('ctrl', 'v')
        time.sleep(1)
        pa.click(x=1785, y=299)
        time.sleep(0.5)
        pa.press('enter')
        time.sleep(5)
        pa.press('enter')
        
    time.sleep(delay)
    # 전체 선택 및 삭제
    pa.click(x=680, y=418)
    time.sleep(delay)
    pa.click(x=1887, y=396)
    time.sleep(delay)
    pa.press("enter")

    try:
        name = df.loc[active_num].iloc[0]
        age = df.loc[active_num].iloc[1]
        is_cols = True
        if age == None:
            is_cols = False
    except:
        name_num = active_num
        age_num = active_num + 1
        age = df.loc[age_num]
        name = df.loc[name_num]
        is_cols = False

    # 이용자 추가
        
    if is_cols: # 정상적인 형태    
        for i in range(len(df)):
            name = df.loc[active_num].iloc[0]
            age = df.loc[active_num].iloc[1]
            
            print(name, age)
            
            pa.click(x=1866, y=390)
            time.sleep(2)
            pa.click(x=737, y=337)
            pc.copy(str(name))
            pa.hotkey('ctrl', 'v')
            pa.press('f2')
            time.sleep(1)
            # 이용자 년도 확인
            try:
                age_1, age_2 = str(age).split('-')
            except:
                add_data(name, age, sheet)
            response = find_user(age_1)
            
            if response:
                pa.click(x=937, y=923)
            else:
                pa.click(x=979, y=910)
                add_data(name, age, sheet)
            
            time.sleep(delay)
            active_num += 1

    else: # 비정상적인 형태
        for i in range(int(len(df)/4)):
            name_num = active_num*4
            age_num = active_num*4+1
            
            age = df.loc[age_num].iloc[0]
            name = df.loc[name_num].iloc[0]
            
            print(name, age)
            
            pa.click(x=1866, y=390)
            time.sleep(2)
            pa.click(x=737, y=337)
            pc.copy(str(name))
            pa.hotkey('ctrl', 'v')
            pa.press('f2')
            time.sleep(1)
            # 이용자 년도 확인
            try:
                age_1, age_2 = str(age).split('-')
            except:
                add_data(name, age, sheet)
            response = find_user(age_1)

            if response:
                pa.click(x=937, y=923)
            else:
                pa.click(x=979, y=910)
                add_data(name, age, sheet)
            
            time.sleep(delay)
            active_num += 1
            
    pa.click(x=1785, y=299)
    time.sleep(0.5)
    pa.press('enter')
    time.sleep(5)
    pa.press('enter')