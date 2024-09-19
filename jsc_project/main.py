import jaego as jg
import json
import pyautogui as pa
import pyperclip as pc
import time
from PIL import ImageGrab
import pandas as pd
import cv2
import re
import pytesseract as tr

#뭔데요
x1, x2, y1, y2 = 1188, 1276, 484, 507
json_file_path = 'config.json'

def find_index_by_name(data_list, target_name):
    for index, data in enumerate(data_list):
        if data["물품명"] == target_name:
            return index
    # 이름을 찾지 못한 경우 처리
    print(f"Name '{target_name}' not found in the list.")
    return None

def add_data_in_excel(name, year, item, donation_date, num,  existing_data_path='찌꺼기.xlsx'):
    try:
        # Read existing data from Excel file
        df_initial = pd.read_excel(existing_data_path)
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame
        df_initial = pd.DataFrame()

    # Add more data to the DataFrame
    new_data = {'이름': [name], '생년월일': [year], '물품명': [item], '수량': [num], '일자': [donation_date]}
    df = pd.concat([df_initial, pd.DataFrame(new_data)], ignore_index=True)

    # Create a Pandas Excel writer object using XlsxWriter as the engine
    writer = pd.ExcelWriter(existing_data_path, engine='xlsxwriter')

    # Write data to the Excel sheet
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Close the file
    writer.close()

with open(json_file_path, 'r', encoding='utf-8-sig') as json_file:
    config = json.load(json_file)

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
        try:
            if len(numbers) >= 2:
                number1 = numbers[0][2:4]
                number2 = numbers[1]
                number3 = numbers[2]
                all_number = number1+number2+number3
                int(all_number)
                print(all_number)
        except:
            return False
    
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
    
def play_t(df):
    delay = 0.5
    
    # 제품 검색
    pa.click(x=912, y=363, clicks=3, button='left')
    pc.copy(str(df['물품명']))
    pa.hotkey('ctrl','v')    
    pa.press('f2')
    time.sleep(2)
    if find_obj(285, 486) != True:
        return
    # 제공 물품 선택
    pa.click(x=274, y=493)
    time.sleep(delay)
    # '+'버튼 클릭
    pa.click(x=1866, y=417)
    time.sleep(delay)
    # 이용자ID 입력
    # pa.click(x=783, y=377)
    pa.click(x=696, y=414)
    time.sleep(delay)
    pa.click(x=793, y=331)
    time.sleep(delay)
    pc.copy(df['이용자ID'])
    pa.hotkey('ctrl','v')
    time.sleep(delay)
    pa.press("f2")
    time.sleep(delay)
    # 검색된 이용자 선택
    if find_obj(1207, 491):
        time.sleep(delay)
    else:
        print('이용자 검색 실패')
        pa.click(x=982, y=916)
        add_data_in_excel(df['물품명'], df['생년월일'], df['일자'], df['수량'])
        return
    # 확인
    pa.click(x=931, y=919)
    time.sleep(delay)
    # 수량 변경
    pa.click(x=1208, y=469, clicks=2, button='left')
    time.sleep(delay)
    pc.copy(str(df['수량']))
    pa.hotkey('ctrl','v')
    time.sleep(delay)
    pa.press('enter')
    time.sleep(delay)
    pa.press('enter')
    # 제공일자 변경
    pa.click(x=861, y=475)
    time.sleep(delay)
    pc.copy(str(df['일자']))
    pa.hotkey('ctrl','v')
    time.sleep(delay)
    # 저장
    pa.click(x=1869, y=289)
    time.sleep(delay)
    pa.press('enter')
    time.sleep(3)
    pa.press('enter')

def user_t(df):
    global y1
    delay = 0.8
    
    # # 제품 검색
    # pa.click(x=912, y=363, clicks=3, button='left')
    # pc.copy(str(df['물품명']))
    # pa.hotkey('ctrl','v')    
    # pa.press('f2')
    # time.sleep(2)
    if find_obj(285, 486) != True:
        return
    # 제공 물품 선택
    pa.click(x=274, y=493)
    time.sleep(delay)
    # '+'버튼 클릭
    pa.click(x=1866, y=417)
    time.sleep(delay)
    # 이용자 입력
    # pa.click(x=783, y=377)
    pa.click(x=801, y=329)
    time.sleep(delay)
    pc.copy(df['이름'])
    pa.hotkey('ctrl','v')
    time.sleep(delay)
    pa.press("f2")
    time.sleep(delay)
    # 검색된 이용자 선택
    if find_user(int(df['생년월일'])):
        ...
    else:
        print("!!이용자 검색 실패!!")
        print(df)
        pa.click(x=987, y=919)
        add_data_in_excel(df['물품명'], df['생년월일'], df['이름'], df['일자'], df['수량'])
        pa.click(x=274, y=493)
        return
    
    # 확인
    pa.click(x=931, y=919)
    time.sleep(delay)
    # 수량 변경
    pa.click(x=1208, y=469, clicks=2, button='left')
    time.sleep(delay)
    pc.copy(str(df['수량']))
    pa.hotkey('ctrl','v')
    time.sleep(delay)
    pa.press('enter')
    time.sleep(delay)
    pa.press('enter')
    # 제공일자 변경
    pa.click(x=861, y=475)
    time.sleep(delay)
    pc.copy(str(df['일자']))
    pa.hotkey('ctrl','v')
    time.sleep(delay)
    # 저장
    pa.click(x=1869, y=289)
    time.sleep(delay)
    pa.press('enter')
    time.sleep(3)
    pa.press('enter')
    
def play_s(df):
    delay = 1
    
    # 제품 검색
    pa.click(x=912, y=363, clicks=3, button='left')
    pc.copy(str(df['물품명']))
    pa.hotkey('ctrl','v')    
    pa.press('f2')
    time.sleep(2)
    if find_obj(285, 486) != True:
        return
    # 제공 물품 선택
    pa.click(x=274, y=493)
    time.sleep(delay)
    # '+'버튼 클릭
    pa.click(x=1866, y=417)
    time.sleep(delay)
    # 이용자ID 입력
    pa.click(x=783, y=377)
    pc.copy(df['이용자ID'])
    pa.hotkey('ctrl','v')
    time.sleep(delay)
    pa.press("f2")
    time.sleep(delay)
    # 확인
    if find_obj(x=648, y=497) == False:
        pa.click(x=985, y=917)
        time.sleep(delay)
        pa.click(x=461, y=207)
        time.sleep(delay)
        pa.click(x=1243, y=344, clicks=2)
        time.sleep(delay)
        pc.copy(str(df['이용자ID']))
        pa.hotkey('ctrl', 'v')
        time.sleep(delay)
        pa.press('f2')
        time.sleep(delay)
        pa.click(x=992, y=506, clicks=2)
        time.sleep(delay)
        pa.press("enter")
        time.sleep(delay)
        pa.click(x=1018, y=618)
        time.sleep(delay)
        pa.click(x=1046, y=639)
        time.sleep(delay)
        pa.click(x=1057, y=587)
        time.sleep(delay)
        pa.click(x=1045, y=633)
        time.sleep(delay)
        pa.click(x=1739, y=290)
        time.sleep(1)
        pa.press('enter')
        time.sleep(4)
        pa.press('enter')
        time.sleep(delay)
        pa.click(x=346, y=208)
        pa.click(x=912, y=363, clicks=3, button='left')
        pc.copy(str(df['물품명']))
        pa.hotkey('ctrl','v')    
        pa.press('f2')
        time.sleep(2)
        if find_obj(285, 486) != True:
            return
        # 제공 물품 선택
        pa.click(x=274, y=493)
        time.sleep(delay)
        # '+'버튼 클릭
        pa.click(x=1866, y=417)
        time.sleep(delay)
        # 이용자ID 입력
        pa.click(x=783, y=377)
        pc.copy(df['이용자ID'])
        pa.hotkey('ctrl','v')
        time.sleep(delay)
        pa.press("f2")
        time.sleep(delay)
    
    pa.click(x=960, y=495, clicks=2)
    time.sleep(delay)
    pa.click(x=931, y=919)
    time.sleep(delay)
    # 수량 변경
    pa.click(x=1208, y=469, clicks=2, button='left')
    time.sleep(delay)
    pc.copy(str(df['수량']))
    pa.hotkey('ctrl','v')
    time.sleep(delay)
    pa.press('enter')
    time.sleep(delay)
    pa.press('enter')
    pa.click(x=861, y=475)
    time.sleep(delay)
    pc.copy(str(df['일자']))
    pa.hotkey('ctrl','v')
    time.sleep(delay)
    # 저장
    pa.click(x=1869, y=289)
    time.sleep(delay)
    pa.press('enter')
    time.sleep(3)
    pa.press('enter')

print("----------------------------------------------------------")
# selected_columns = ['일자', '이용자ID', '수량', '물품명']
selected_columns = ['일자', '이름', '수량', '물품명', '생년월일']
data = jg.load_data(selected_columns)
data_length = len(data)

pa.hotkey('alt', 'tab')
for i in range(data_length):
    print("----------------------------------------------------------")
    data.loc[data.index[i], '물품명'] = 'CJ'
    # data.loc[data.index[i], '수량'] = 2
    # play_t(data.iloc[i]) # 시설
    user_t(data.iloc[i]) # 이용자(생년월일)
    # play_s(data.iloc[i]) # 이용자(ID)
    print(data.iloc[i])
    time.sleep(2)
print("----------------------------------------------------------")
