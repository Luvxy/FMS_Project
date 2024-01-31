import jaego as jg
import json

path = jg.select_excel_file()
json_file_path = 'config.json'

def find_index_by_name(data_list, target_name):
    for index, data in enumerate(data_list):
        if data["물품명"] == target_name:
            return index
    # 이름을 찾지 못한 경우 처리
    print(f"Name '{target_name}' not found in the list.")
    return None

def add_data_in_excel(path, name, num, donation_date, expiration_date):
    data_string = """
        "물품명": "{name}",
        "현 재고": {num},
        "기부일자": "{donation_date}",
        "유통기한(소비기한)": "{expiration_date}"
    """

jg.add_excel_data_to_json(path, json_file_path)

with open(json_file_path, 'r', encoding='utf-8-sig') as json_file:
    config = json.load(json_file)

find_it = find_index_by_name(config, "담꽃)통통단팥죽 250g")

print("----------------------------------------------------------")
print(find_it)
print("----------------------------------------------------------")
print(config[find_it])
print("----------------------------------------------------------")
name = input("name: ")
num = input("num: ")
donation_date = input("donation_date: ")
expiration_date = input("expiration_date: ")
print("----------------------------------------------------------")

add_data_in_excel(config, name, num, donation_date, expiration_date)