import pandas as pd
import json
from tkinter import filedialog as fd

# Step 1: Read Excel Data
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

def set_excel_data_to_json(excel_file_path, json_file_path, selected_columns):
    # Step 1: Read Excel Data
    df = pd.read_excel(excel_file_path)

    # Step 2: Select Relevant Columns
    selected_data = df[selected_columns]

    # Step 3: Convert Data to JSON
    json_data = selected_data.to_json(orient='records', force_ascii=False, indent=4)
    selected_data['일자'] = selected_data['일자'].dt.strftime('%Y-%m-%d')

    # Step 4: Save to config.json
    with open(json_file_path, 'w', encoding='utf-8-sig') as json_file:
        json_file.write(json_data)

def add_excel_data_to_json(excel_file_path, json_file_path, selected_columns):
    # Step 1: Read Excel Data
    df = pd.read_excel(excel_file_path)

    # Optional: Normalize column names by stripping trailing spaces
    df.columns = df.columns.str.strip()

    # Step 2: Select Relevant Columns (make sure these match exactly with your Excel columns)
    selected_columns = ['이용자ID', '물품명', '일자', "수량"]
    selected_data = df[selected_columns]
    
    print (selected_data)

    # Step 3: Convert datetime columns to strings with a specific format
    selected_data['일자'] = selected_data['일자'].dt.strftime('%Y-%m-%d')

    # Step 4: Convert Data to JSON
    json_data = selected_data.to_dict(orient='records')

    # Step 5: Load existing JSON data (if any) and handle BOM manually
    try:
        with open(json_file_path, 'r', encoding='utf-8') as existing_json_file:
            file_content = existing_json_file.read()

            # Remove BOM if present
            if file_content.startswith('\ufeff'):
                file_content = file_content[1:]

            existing_data = json.loads(file_content)
    except FileNotFoundError:
        existing_data = []

    # Step 6: Append new data to existing data
    existing_data.extend(json_data)
    
    # if same "물품명" and "기부일자" exists, remove it and edit "현 재고"
    existing_data = [i for n, i in enumerate(existing_data) if i not in existing_data[n + 1:]]

    # Step 7: Save updated data to the JSON file
    with open(json_file_path, 'w', encoding='utf-8-sig') as json_file:
        json.dump(existing_data, json_file, default=str, indent=4, ensure_ascii=False)
        
def load_data(selected_columns, sheet_name="Sheet1"):
    path = select_excel_file()
    df = pd.read_excel(path, sheet_name=sheet_name)
    
    # Optional: Normalize column names by stripping trailing spaces
    df.columns = df.columns.str.strip()

    # Step 2: Select Relevant Columns (make sure these match exactly with your Excel columns)
    selected_data = df[selected_columns]
    
    return selected_data