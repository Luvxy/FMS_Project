# 설문조사 통계 프로그램
import pandas as pd
import random as rd

lo_name = []
choice_list1 = []
choice_list2 = []
choice_list3 = []
ran_list = [5,5,5,5,5,5,5,5,5,4,4,3]
ran_list2 = [1,1,1,1,4,2,3,2,6,3,3,3,4,5,6]
ran_list3 = [1,1,1,1,1,1,1,1,2,3,4,5,6]
ran_list4 = [1,1,1,1,1,1,1,1,1,1,1,1,2]
# 남녀 성비 3:7
ran_list5 = [1,1,1,2,2,2,2,2,2,2]
# 0:1:1:3:3:4:5:4 연령대
ran_list6 = [2,3,4,4,4,5,5,5,6,6,6,6,7,7,7,7,7,8,8,8,8]
# 수급형태 7:6:7
ran_list7 = [1,1,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3,3,3]
# 주거현황 7:1:7:1:1:3
ran_list8 = [1,1,1,1,1,1,1,2,3,3,3,3,3,3,3,4,5,6,6,6]
# 주거 형태 7:10:1:1:1:1
ran_list9 = [1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,3,4,5,6]
# 방문경로 1,8,8,1,1,1,1
ran_list10 = [1,2,2,2,3,3,3,4,4,4,4,4,4,4,4,4,4,4,5,6]

user_num = 180

path = 'C:/Users/user/Desktop/설문통계.xlsx'
df = pd.read_excel(path, sheet_name='Sheet1')
print(df.iloc[1]) # 첫 번째 줄
    
print(choice_list1)
# 성별
for i in range(user_num):
    df.iloc[i+1, 2] = rd.choice(ran_list5)

# 연령대
for i in range(user_num):
    df.iloc[i+1, 3] = rd.choice(ran_list6)

# 수급형태
for i in range(user_num):
    df.iloc[i+1, 4] = rd.choice(ran_list7)

# 주거현황
for i in range(user_num):
    df.iloc[i+1, 5] = rd.choice(ran_list8)

# 주거형태
for i in range(user_num):
    df.iloc[i+1, 6] = rd.choice(ran_list9)

# 방문경로
for i in range(user_num):
    df.iloc[i+1, 7] = rd.choice(ran_list10)

# 만족도
for j in range(7):
    for i in range(user_num):
        df.iloc[i+1, j+8] = rd.choice(ran_list)

# 욕구조사
for i in range(user_num):
    dd = rd.sample(ran_list2, 2)
    df.iloc[i+1, 15] = dd[0]
    df.iloc[i+1, 16] = dd[1]

for i in range(user_num):
    dd = rd.sample(ran_list2, 2)
    df.iloc[i+1, 17] = dd[0]
    df.iloc[i+1, 18] = dd[1]
    
for i in range(user_num):
    df.iloc[i+1, 19] = rd.choice(ran_list2)

for i in range(user_num):
    df.at[i+1, 20] = rd.choice(ran_list4)

save_path = 'C:/Users/user/Desktop/설문통계_save.xlsx'

# 저장
df.to_excel(save_path, sheet_name='Sheet1', index=False)

