import pandas as pd
import os

# 엑셀 파일들이 있는 디렉토리 경로
directory = 'C:/Users/user/Desktop/완'  # 엑셀 파일들이 있는 폴더 경로로 변경하세요

# 지정한 단어
keyword = 'CJ'  # 찾고자 하는 단어로 변경하세요

# 결과를 저장할 빈 데이터프레임 생성
combined_df = pd.DataFrame()

# 디렉토리 내의 모든 엑셀 파일 읽기
for filename in os.listdir(directory):
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        file_path = os.path.join(directory, filename)
        df = pd.read_excel(file_path)

        # 지정한 단어가 포함된 행 필터링
        filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(keyword).any(), axis=1)]

        # 필터링된 데이터를 결과 데이터프레임에 추가
        combined_df = pd.concat([combined_df, filtered_df], ignore_index=True)

# 결과를 새로운 엑셀 파일로 저장
combined_df.to_excel('filtered_results.xlsx', index=False)

print("작업이 완료되었습니다. 결과는 'filtered_results.xlsx'에 저장되었습니다.")
