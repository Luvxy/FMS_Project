import os
import win32com.client as win32

# 엑셀 파일들이 있는 디렉토리 경로
directory = os.path.normpath('C:/Users/user/Desktop/완')  # Ensure path consistency

count = 0

for filename in os.listdir(directory):
    if filename.endswith(".XLS"):
        file_path = os.path.join(directory, filename)
    try:
        # Excel 애플리케이션 시작
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        
        # .xls 파일 열기
        wb = excel.Workbooks.Open(file_path)
        
        # 변환할 .xlsx 파일 경로 (확실히 고유한 이름 사용)
        new_filename = os.path.join(directory, f'converted_{count}.xlsx')
        count += 1
        
        # .xlsx로 저장
        wb.SaveAs(new_filename, FileFormat=51)  # 51은 xlsx 포맷
        
        # 파일 닫기
        wb.Close()

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

    finally:
        # Excel 애플리케이션 종료
        excel.Application.Quit()

print("모든 파일이 변환되었습니다.")
