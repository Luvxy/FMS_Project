import pyautogui as pa
import keyboard
import time
import PIL
import pyscreeze

def click_sequence(x1, y1, x2, y2, sleep_time):
    pa.click(x=x1, y=y1)
    time.sleep(sleep_time)
    pa.click(x=x2, y=y2)

def find_and_click(image_path, clicks=1, interval=0.25, duration=0.1):
    """지정된 이미지를 화면에서 찾아 클릭합니다.

    Args:
        image_path (str): 찾을 이미지 파일의 경로.
        clicks (int): 클릭 횟수.
        interval (float): 클릭 사이의 간격 (초).
        duration (float): 클릭 동작의 지속 시간 (초).
    """
    try:
        # 화면에서 이미지 위치 찾기
        location = pa.locateOnScreen(image_path)
        if location:
            # 이미지 중앙에 마우스 포인터 이동
            center = pa.center(location)
            # 이미지 위치 클릭
            pa.click(center.x, center.y, clicks=clicks, interval=interval, duration=duration)
            print(f"Clicked on {image_path}")
        else:
            print(f"Image not found on the screen: {image_path}")
    except pa.ImageNotFoundException:
        print("Provided image not found on the screen.")
    except Exception as e:
        print(f"An error occurred: {e}")

def comand_1():
    click_sequence(283, 389, 702, 415, 0.1)

def comand_2():
    click_sequence(555, 379, 932, 461, 0.1)

def comand_3():
    click_sequence(1865, 300, 923, 728, 0.7)
    
def comand_4():
    find_and_click("C:/Users/user/Desktop/coding/OK.png")

def setup_hotkeys():
    keyboard.add_hotkey('alt+1', comand_1)
    keyboard.add_hotkey('alt+2', comand_2)
    keyboard.add_hotkey('alt+3', comand_3)
    keyboard.add_hotkey('insert', comand_4)
    print("Hotkeys are set up. Press 'end' to exit.")
    keyboard.wait('end')

if __name__ == '__main__':
    setup_hotkeys()