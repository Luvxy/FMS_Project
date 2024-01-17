import time
from pynput import mouse, keyboard
import pyautogui as pa

def edit_pos():
    return pa.position()

def on_click(x, y, button, pressed):
    if pressed:
        print('Mouse clicked at ({0}, {1}) with {2}'.format(x, y, button))
        return x, y

sh_name = input('시트 이름: ')

# 마우스 위치 수정 및 저장
rows = 21
cols = 2
edit_count = 0  # 진행중인 순서
mouse_pos = [[0 for j in range(cols)] for i in range(rows)]  # 마우스 위치 배열

if sh_name == '수정':  # 시트 이름을 수정으로 입력 시 마우스 위치 수정
    print("수정")
    i = 0
    with mouse.Listener(on_click=on_click) as listener:
        listener.join()
        
    while True:
        # esc 키 입력 시 마우스 위치 저장 종료
        if pa.keyDown('esc'):
            break

        # 마우스 버튼 클릭 시 마우스 위치 저장
        click_result = on_click(None, None, None, True)
        if click_result:
            mouse_pos[i] = click_result
            edit_count += 1
            print(edit_count)

        # 전체 다 저장 후 종료
        if edit_count == rows:
            break
        i += 1

print(mouse_pos)
