import pytesseract as tr
import cv2
from PIL import ImageGrab
import re
import pyautogui as pa
import time
import keyboard
import pandas as pd
import jsc_project.jaego as jg

def remove_specific_words(text):
    st_list = ["월", "일", " "]
    result = ""
    tem = ""
    for i in range(len(text)):
        if text[i] in st_list:
            continue
        else:
            result += text[i]
        
    return result

def sort_data_data(df):
    # Assuming we have a DataFrame `df` with a '날짜' column
    # Let's create a sample DataFrame to demonstrate

    # Convert the '날짜' column to datetime
    df['최근'] = pd.to_datetime(df['최근'])

    # Sort the DataFrame by the '날짜' column
    df_sorted = df.sort_values(by='최근')

    return df_sorted

def select_group(data, num):
    """데이터(data)를 받아서 주어진 빵 전체 개수에 맞게 배분가능한 기관을 선택

    Args:
        data (list): data frame from pandas excel
        num (int): number of bread
    """
    length = len(data)
    data = sort_data_data(data)
    total = 0
    result = pd.DataFrame(columns=['시설명', '빵수', '최근'])
    for i in range(length):
        if total < num:
            tem = int(data.iloc[i]["빵수"])
            if tem == 0:
                continue
            total += tem
            result = pd.concat([result, data.iloc[[i]]], ignore_index=True)
        else:
            break
    
    return result

def distribute_numbers(left_numbers, right_numbers):
    from itertools import combinations

    # Initialize result list
    result = []

    # Iterate over each left number and try to find a matching combination from right numbers
    for left in left_numbers:
        found = False
        for i in range(1, len(right_numbers) + 1):
            for comb in combinations(right_numbers, i):
                if sum(comb) == left:
                    result.append((left, list(comb)))
                    for num in comb:
                        right_numbers.remove(num)
                    found = True
                    break
            if found:
                break
        if not found:
            result.append((left, []))  # If no combination is found

    return result

selected_columns = ['시설명', '빵수', '최근']
group_data = jg.load_data(selected_columns, sheet_name="빵배분시설 목록(24.2.16~)")

group_list = select_group(group_data, 100)

print(group_list)

left_numbers = group_list["빵수"]
right_numbers = group_list["수량"]

