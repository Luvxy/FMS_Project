import pandas as pd
import json

# Excel 파일 경로
excel_file_path = 'H:/.shortcut-targets-by-id/1A0TIuPAsmbBdRF01K7yT89PSL4LTaONB/익산행복나눔마켓뱅크/3. 양식 및 도구 (정리 필요 !!!)/아동센터 목록(+배분실적).xlsx'

# 읽어올 열(column)의 이름
column_name = 1

# Excel 파일 읽기
try:
    df = pd.read_excel(excel_file_path, sheet_name='1110', usecols=[column_name])
except FileNotFoundError:
    print(f"Error: The file '{excel_file_path}' not found.")
    exit()


print(df)
# 열의 데이터를 리스트로 변환
column_data = df.values.tolist()

config_data = {
    'column_data': str(column_data)
}

# config.json 파일에 저장
config_file_path = 'config2.json'
with open(config_file_path, 'w', encoding='utf-8') as config_file:
        json.dump(config_data, config_file, indent=4, ensure_ascii=False)

print(f"Data from column '{column_name}' has been successfully saved to '{config_file_path}'.")
