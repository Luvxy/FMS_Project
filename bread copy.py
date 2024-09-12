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
import keyboard
import pandas as pd

bread_name_list = [
    '파리어양', '파리이편한', '던킨영등', '파리고래등',
    '쿱스영등', '초코', '파리고래등', '파리이편한', '그랜드',
    '파리배산', '파리동산', '쿱스어양', '뚜레원대', '파리제일',
    '크리영등', '파리부송', '뚜레동산', '파리동산', '파리동산',
    '브레드팜', '파리원대', '홍윤', '하나인화', '온스', '파리동산',
    '안영순', '파리동부', '쿱스모현', '니코니코', '뚜레제일', '파리하나',
    '하나어양', '던킨익산역', '크리모현', '눈들재', '당고', '오소','뚜레영등','서울치즈'
]

name_list = [
    '파리바게뜨 익산어양', '파리바게뜨 익산이편한', '던킨도너츠(영등점)', '파리바게뜨 익산고래등',
    '쿱스토어전북 영등점', '초코', '파리바게뜨 익산고래등', '파리바게뜨 익산이편한', '그랜드제과',
    '파리바게뜨 배산', '파리바게뜨 동산', '쿱스토어전북 어양점', '뚜레쥬르 익산원광', '파리바게뜨 부송제일',
    '크리스피크림도넛(영등점)', '파리바게뜨 익산부송', '뚜레쥬르 익산동산', '파리바게뜨 익산동산', '파리바게뜨 동산',
    '브레드팜', '파리바게뜨 익산원광', '홍윤', '하나로 베이커리 인화점', '온스베이커리', '파리바게뜨 익산동산',
    '안영순', '파리바게뜨 익산동부', '쿱스토어전북 모현점', '니코니코 과자점', '뚜레쥬르 익산제일', '파리바게뜨 부송하나',
    '하나로 베이커리(어양점)', '던킨도너츠 익산역점', '크리스피크림도넛 익산모현', '눈들재', '당고', '오소베이커리',
    '뚜레쥬르 익산영등점','익산군산서울치즈대리점'
]

change_list = []

sh_name = input('시트 이름: ')

path = "H:/.shortcut-targets-by-id/1A0TIuPAsmbBdRF01K7yT89PSL4LTaONB/익산행복나눔마켓뱅크/3. 양식 및 도구/빵 관련/아동센터 목록(+배분실적)24.2.16ver..xlsx"

df = pd.read_excel(path, sheet_name=sh_name)
try:
    print(len(df['빵집']))
except:
    print("column 이름이 잘못되었습니다.")

# 각 행을 반복하면서 빵 이름을 확인하고 일치하는 이름을 change_list에 추가
for i, row in df.iterrows():
    bread_name = row['빵집']
    if bread_name in bread_name_list:
        index = bread_name_list.index(bread_name)
        change_list.append(name_list[index])

print(change_list)
df = df.reindex(range(len(change_list)))
df['빵집'] = change_list

# pandas를 사용하여 엑셀 파일 쓰기
df.to_excel('test.xlsx', sheet_name=sh_name, index=False)

is_true = False

def edit_pos():
    return pa.position()
    
print("파일 불러오기")
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

# num_in = []
# while(True):
#     a = input("순서: ")
#     if a == '.':
#         break
#     else:
#         num_in.append(int(a))
num = 0
sum = 0
count = 0
aoj = 0

# for i in num_in:
#     if i == '.':
#         continue
#     try:
#         print(df.iloc[aoj+i][0])
#         aoj += i
#     except:
#         print("몰라")
    
# start = input("시작하시겠습니까? (y/n): ")
# if start == 'n':
#     exit()
# else:
#     print("시작합니다.")
#     time.sleep(1)

count = 0
pa.hotkey('alt', 'tab')

# # 백그라운드 작업을 담당할 스레드 생성
# background_thread = th.Thread(target=background_task)
# # 백그라운드 스레드 시작
# background_thread.start()

row_num = 0

def bread():
    global sum
    global num
    global row_num
    # global num_in

    print("빵 등록")

    data_fr = df.iloc[row_num]
    print("기관 명:"+ str(data_fr.iloc[0]))
    print(row_num)
    time.sleep(0.5)
    while True:
        print(i)
        data = df.iloc[row_num]
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
            
        row_num = row_num + 1
        if df.iloc[row_num, 0] != None:
            break
            
    
def setup_hotkeys():
    keyboard.add_hotkey('alt+1', bread)
    print("Hotkeys are set up. Press 'end' to exit.")
    keyboard.wait('insert')

if __name__ == '__main__':
    setup_hotkeys()


# while(True):
#     # 한 기관에 등록되는 빵집 수를 리스트로 입력 받음
#     # 입력반은 리스트는 num_in에 저장
#     # 만약'.'을 입력하면 입력을 종료하고 빵 등록 시작
#     # 빵 등록이 끝나면 다음 기관으로 넘어감
#     # 기관 시작 위치 = sum
#     # 기관당 빵집 수 = num
#     # 기관당 빵집 수를 더한 값 = active_num
#     # acrive_num을 통해 다음 기관을 엑셀 데이터에서 지정
#     # active_num을 num에 저장
#     # 문제가 생기면 정지

#     active_num = sum + num
#     sum = sum + num
#     num = int(num_in[count])

#     print("빵 등록")
    
#     row_num = sum
#     data_fr = df.iloc[row_num]
#     print("기관 명:"+ str(data_fr.iloc[0]))
#     print(row_num)
#     time.sleep(0.5)
#     for i in range(num):
#         print(i)
#         data = df.iloc[row_num+i]
#         # print(data)
#         print("기관명: "+str(data.iloc[1]))
#         print("만약 20000원 이면 케이크로 등록")
#         if data.iloc[3] == 20000:
#             print("케이크 등록")
#             pa.click(x=454, y=199)
#             time.sleep(0.5)
#             pa.click(x=1434, y=284)
#             time.sleep(1)
#             pa.click(x=709, y=366)
#             time.sleep(0.5)
#             pa.click(x=1168, y=362)
#             pc.copy(str(data.iloc[1]))
#             pa.hotkey('ctrl', 'v')
#             time.sleep(0.5)
#             pa.press('f2')
#             time.sleep(0.5)
#             pa.click(x=953, y=467, clicks=2, button='left')
#             time.sleep(0.2)
#             pa.click(x=1870, y=448)
#             time.sleep(0.5)
#             pa.click(x=955, y=455)
#             time.sleep(0.5)
#             pc.copy('기타케익')
#             pa.hotkey('ctrl', 'v')
#             time.sleep(0.5)
#             pa.click(x=1101, y=453)
#             time.sleep(0.5)
#             pa.press('f2')
#             time.sleep(0.5)
#             pa.click(x=1069, y=526, clicks=2, button='left')
#             time.sleep(0.2)
#             pa.click(x=950, y=767)
#             time.sleep(0.2)
#             pa.click(x=748, y=519)
#             pc.copy(str(data.iloc[2]))
#             time.sleep(0.2)
#             pa.hotkey('ctrl', 'v')
#             pa.press('tab')
#             pc.copy(str(int(data.iloc[2])))
#             pa.hotkey('ctrl', 'v')
#             pa.press('tab')
#             pa.press('tab') 
#             pa.press('tab')
#             pa.press('tab') 
#             pc.copy(str(data.iloc[4]))
#             pa.hotkey('ctrl', 'v')
#             time.sleep(0.2)
#             pa.click(x=1867, y=281)
#             pa.press('enter')
#             time.sleep(0.2)
#             pa.press('enter')
#             time.sleep(0.2)
#             pa.press('enter')
#             time.sleep(0.2)
#         else:
#             print("빵 등록")
#             pa.click(x=454, y=199)
#             time.sleep(0.5)
#             pa.click(x=1434, y=284)
#             time.sleep(1)
#             pa.click(x=709, y=366)
#             time.sleep(0.5)
#             pa.click(x=1168, y=362)
#             pc.copy(str(data.iloc[1]))
#             pa.hotkey('ctrl', 'v')
#             time.sleep(0.5)
#             pa.press('f2')
#             time.sleep(0.5)
#             pa.click(x=953, y=467, clicks=2, button='left')
#             time.sleep(0.5)
#             pa.click(x=1870, y=448)
#             time.sleep(0.5)
#             pa.click(x=955, y=455)
#             time.sleep(0.5)
#             pc.copy('기타빵')
#             pa.hotkey('ctrl', 'v')
#             time.sleep(0.5)
#             pa.click(x=1101, y=453)
#             time.sleep(0.5)
#             pa.press('f2')
#             time.sleep(1)
#             pa.click(x=1069, y=526, clicks=2, button='left')
#             time.sleep(0.5)
#             pa.click(x=950, y=767)
#             time.sleep(0.5)
#             pa.click(x=748, y=519)
#             time.sleep(0.5)
#             pc.copy(str(data.iloc[2]))
#             time.sleep(0.5)
#             pa.hotkey('ctrl', 'v')
#             time.sleep(0.2)
#             pa.press('tab')
#             time.sleep(0.2)
#             pc.copy(str(int(int(data.iloc[2])/5)))
#             if int(int(data.iloc[2])/5) == 0:
#                 pc.copy(str('1'))
#             pa.hotkey('ctrl', 'v')
#             print(int(int(data.iloc[2])/5))
#             pa.press('tab')
#             pa.press('tab')
#             pa.press('tab')
#             pa.press('tab')
#             pc.copy(str(data.iloc[4]))
#             pa.hotkey('ctrl', 'v')
#             time.sleep(0.2)
#             pa.click(x=1867, y=281)
#             pa.press('enter')
#             time.sleep(0.5)
#             pa.press('enter')
#             time.sleep(0.5)
#             pa.press('enter')
#             time.sleep(0.5)
    
#     print("빵 제공")
#     time.sleep(0.5)
#     # Point(x=338, y=198)당고
#     pa.click(x=338, y=198)
#     time.sleep(0.5)
#     # Point(x=1794, y=278)
#     pa.click(x=1804, y=291)
#     time.sleep(1) 
#     # Point(x=267, y=431) 재고
#     pa.click(x=269, y=445)
#     time.sleep(0.5)
#     # Point(x=1867, y=407) 플러스18
#     pa.click(x=1866, y=418)
#     time.sleep(0.5)
#     # Point(x=684, y=392)
#     pa.click(x=680, y=429)
#     time.sleep(0.5)
#     # Point(x=767, y=325)
#     pa.click(x=793, y=333)
#     time.sleep(0.5)
#     pc.copy(str(data_fr.iloc[0]))
#     pa.hotkey('ctrl', 'v')
#     time.sleep(0.5)
#     pa.press('f2')
#     time.sleep(1)
#     # Point(x=806, y=472)
#     pa.click(x=956, y=494, clicks=2, button='left')
#     time.sleep(0.5)
#     # Point(x=937, y=915)
#     pa.click(x=937, y=920)
#     time.sleep(0.5)
#     # Point(x=1866, y=278)
#     pa.click(x=1847, y=291)
#     time.sleep(0.5)
#     pa.press('enter')
#     time.sleep(0.5)
#     pa.press('enter')
    
    # con = input("계속 하시겠습니까? (y/n): ")
    # if con == 'n':
    #     break
    # else:
    #     print("기관 명:"+ str(data_fr.iloc[0])) 
    #     pa.hotkey('alt','tab')
    #     if count == len(num_in):
    #         break

# # 백그라운드 스레드 종료
# background_thread.join()
# print("Main thread exits")