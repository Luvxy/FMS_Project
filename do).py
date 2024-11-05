import pyautogui as pa
import time
import pandas as pd
import json
from tkinter import filedialog as fd
import pyperclip as pc
from PIL import ImageGrab

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

def select_excel_file():
    filetypes = (
            ('excel files', '*.xlsx'),
            ('All files', '*.*')
        )
    try:
        filename = fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes)
        return filename
    except:
        print("파일을 선택하지 않았습니다.")
        exit()

path = select_excel_file()
df = pd.read_excel(path, sheet_name="Sheet2")
df = df['후원 기록 x']

pa.hotkey('alt', 'tab')

for i in range(421):
    delay = 0.5
    pa.click(x=837, y=369)
    time.sleep(delay)
    pc.copy(str(df[i]))
    pa.hotkey('ctrl', 'v')
    time.sleep(delay)
    pa.press('f2')
    time.sleep(delay)
    if find_obj(x=295, y=472) == False:
        pa.click(x=1714, y=296)
        continue
    
    pa.click(x=1014, y=474, clicks=2, button='left')
    time.sleep(delay)
    pa.click(x=1094, y=426)
    time.sleep(0.2)
    pa.click(x=1053, y=450)
    time.sleep(0.2)
    pa.click(x=391, y=490)
    time.sleep(0.2)
    for i in range(3):
        pa.press('down')
    pa.press('enter')
    pa.press('tab')
    for i in range(8):
        pa.press('0')
    pa.click(x=1774, y=297)
    pa.press('enter')
    time.sleep(delay+1)
    pa.press('enter')
    time.sleep(1)
    pa.click(x=1714, y=296)