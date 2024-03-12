import json
import pandas as pd
from tkinter import filedialog as fd
import pyautogui as pa
import pyperclip as pc
import time
import threading as th
from tkinter import *
import keyboard
import _thread

is_true = False

def edit_pos():
    return pa.position()
    
def background_task():
    while True:
        # 주기적으로 실행되어야 하는 작업 시뮬레이션
        if keyboard.is_pressed('esc'):
            _thread.interrupt_main()
            break


print("파일 불러오기")
sh_name = input('시트 이름: ')
active_num = 0

# 마우스 위치 수정 및 저장
rows = 21
cols = 2
edit_count = 0 # 진행중인 순서
mouse_pos = [[0 for j in range(cols)] for i in range(rows)] # 마우스 위치 배열
if(sh_name == '수정'): # 시트이릅을 수정으로 입력 시 마우스 위치 수정
    print("수정")
    while(True):
        # 마우스 버튼 클릭 시 마우스 위치 저장 후 다음 배열로 이동
        if pa.mouseUp() == True:        
            mouse_pos.append(edit_pos())
            print(mouse_pos[edit_count])
            edit_count = edit_count + 1
        elif pa.keyUp('esc') == True: # esc 키 입력 시 마우스 위치 저장 종료
            break
        # 전체 다 저장후 종료
        if(edit_count == 2):
            break

try:
    df = pd.read_excel("test.xlsx", sheet_name=sh_name)
    print(df)
except:
    print("시트 이름이 잘못되었습니다.")
    exit()

num_in = []
while(True):
    a = input("순서: ")
    if a == '.':
        break
    else:
        num_in.append(int(a))
num = 0
sum = 0
count = 0
aoj = 0

for i in num_in:
    if i == '.':
        continue
    try:
        print(df.iloc[aoj+i][0])
        aoj += i
    except:
        print("몰라")
    
start = input("시작하시겠습니까? (y/n): ")
if start == 'n':
    exit()
else:
    print("시작합니다.")
    time.sleep(1)

count = 0
pa.hotkey('alt', 'tab')

# 백그라운드 작업을 담당할 스레드 생성
background_thread = th.Thread(target=background_task)
# 백그라운드 스레드 시작
background_thread.start()

while(True):
    # 한 기관에 등록되는 빵집 수를 리스트로 입력 받음
    # 입력반은 리스트는 num_in에 저장
    # 만약'.'을 입력하면 입력을 종료하고 빵 등록 시작
    # 빵 등록이 끝나면 다음 기관으로 넘어감
    # 기관 시작 위치 = sum
    # 기관당 빵집 수 = num
    # 기관당 빵집 수를 더한 값 = active_num
    # acrive_num을 통해 다음 기관을 엑셀 데이터에서 지정
    # active_num을 num에 저장
    # 문제가 생기면 정지

    active_num = sum + num
    sum = sum + num
    num = int(num_in[count])

    print("빵 등록")
    
    row_num = sum
    data_fr = df.iloc[row_num]
    print("기관 명:"+ str(data_fr.iloc[0]))
    print(row_num)
    time.sleep(0.5)
    for i in range(num):
        print(i)
        data = df.iloc[row_num+i]
        # print(data)
        print("기관명: "+str(data.iloc[1]))
        print("만약 20000원 이면 케이크로 등록")
        if data.iloc[3] == 20000:
            print("케이크 등록")
            pa.click(x=454, y=199)
            time.sleep(0.5)
            pa.click(x=1434, y=284)
            time.sleep(1)
            pa.click(x=709, y=366)
            time.sleep(0.5)
            pa.click(x=1168, y=362)
            pc.copy(str(data.iloc[1]))
            pa.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pa.press('f2')
            time.sleep(0.5)
            pa.click(x=953, y=467, clicks=2, button='left')
            time.sleep(0.2)
            pa.click(x=1870, y=448)
            time.sleep(0.5)
            pa.click(x=955, y=455)
            time.sleep(0.5)
            pc.copy('기타케익')
            pa.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pa.click(x=1101, y=453)
            time.sleep(0.5)
            pa.press('f2')
            time.sleep(0.5)
            pa.click(x=1069, y=526, clicks=2, button='left')
            time.sleep(0.2)
            pa.click(x=950, y=767)
            time.sleep(0.2)
            pa.click(x=748, y=519)
            pc.copy(str(data.iloc[2]))
            time.sleep(0.2)
            pa.hotkey('ctrl', 'v')
            pa.press('tab')
            pc.copy(str(int(data.iloc[2])))
            pa.hotkey('ctrl', 'v')
            pa.press('tab')
            pa.press('tab') 
            pa.press('tab')
            pa.press('tab') 
            pc.copy(str(data.iloc[4]))
            pa.hotkey('ctrl', 'v')
            time.sleep(0.2)
            pa.click(x=1867, y=281)
            pa.press('enter')
            time.sleep(0.2)
            pa.press('enter')
            time.sleep(0.2)
            pa.press('enter')
            time.sleep(0.2)
        else:
            print("빵 등록")
            pa.click(x=454, y=199)
            time.sleep(0.5)
            pa.click(x=1434, y=284)
            time.sleep(1)
            pa.click(x=709, y=366)
            time.sleep(0.5)
            pa.click(x=1168, y=362)
            pc.copy(str(data.iloc[1]))
            pa.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pa.press('f2')
            time.sleep(0.5)
            pa.click(x=953, y=467, clicks=2, button='left')
            time.sleep(0.5)
            pa.click(x=1870, y=448)
            time.sleep(0.5)
            pa.click(x=955, y=455)
            time.sleep(0.5)
            pc.copy('기타빵')
            pa.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pa.click(x=1101, y=453)
            time.sleep(0.5)
            pa.press('f2')
            time.sleep(1)
            pa.click(x=1069, y=526, clicks=2, button='left')
            time.sleep(0.5)
            pa.click(x=950, y=767)
            time.sleep(0.5)
            pa.click(x=748, y=519)
            time.sleep(0.5)
            pc.copy(str(data.iloc[2]))
            time.sleep(0.5)
            pa.hotkey('ctrl', 'v')
            time.sleep(0.2)
            pa.press('tab')
            time.sleep(0.2)
            pc.copy(str(int(int(data.iloc[2])/5)))
            if int(int(data.iloc[2])/5) == 0:
                pc.copy(str('1'))
            pa.hotkey('ctrl', 'v')
            print(int(int(data.iloc[2])/5))
            pa.press('tab')
            pa.press('tab')
            pa.press('tab')
            pa.press('tab')
            pc.copy(str(data.iloc[4]))
            pa.hotkey('ctrl', 'v')
            time.sleep(0.2)
            pa.click(x=1867, y=281)
            pa.press('enter')
            time.sleep(0.5)
            pa.press('enter')
            time.sleep(0.5)
            pa.press('enter')
            time.sleep(0.5)
    
    print("빵 제공")
    time.sleep(0.5)
    # Point(x=338, y=198)당고
    pa.click(x=338, y=198)
    time.sleep(0.5)
    # Point(x=1794, y=278)
    pa.click(x=1804, y=291)
    time.sleep(1) 
    # Point(x=267, y=431) 재고
    pa.click(x=269, y=445)
    time.sleep(0.5)
    # Point(x=1867, y=407) 플러스18
    pa.click(x=1866, y=418)
    time.sleep(0.5)
    # Point(x=684, y=392)
    pa.click(x=680, y=429)
    time.sleep(0.5)
    # Point(x=767, y=325)
    pa.click(x=793, y=333)
    time.sleep(0.5)
    pc.copy(str(data_fr.iloc[0]))
    pa.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pa.press('f2')
    time.sleep(1)
    # Point(x=806, y=472)
    pa.click(x=956, y=494, clicks=2, button='left')
    time.sleep(0.5)
    # Point(x=937, y=915)
    pa.click(x=937, y=920)
    time.sleep(0.5)
    # Point(x=1866, y=278)
    pa.click(x=1847, y=291)
    time.sleep(0.5)
    pa.press('enter')
    time.sleep(0.5)
    pa.press('enter')
    
    count = count + 1
    time.sleep(1)
    
    # con = input("계속 하시겠습니까? (y/n): ")
    # if con == 'n':
    #     break
    # else:
    #     print("기관 명:"+ str(data_fr.iloc[0])) 
    #     pa.hotkey('alt','tab')
    #     if count == len(num_in):
    #         break

# 백그라운드 스레드 종료
background_thread.join()
print("Main thread exits")