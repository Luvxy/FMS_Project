import pandas as pd

bread_name_list = [
    '파리어양', '파리이편한', '던킨영등', '파리고래등',
    '쿱스영등', '초코', '파리고래등', '파리이편한', '그랜드',
    '파리배산', '파리동산', '쿱스어양', '뚜레원광', '파리제일',
    '크리영등', '파리부송', '뚜레동산', '파리동산', '파리동산',
    '브레드팜', '파리원광', '홍윤', '하나인화', '온스', '파리동산',
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
