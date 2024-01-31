import pandas as pd
import json
import pytesseract as tr
import cv2
from PIL import ImageGrab
import re
import pyautogui
import time
import keyboard

y1, y2 = 250, 285
y_plus = 34
count = 40

# Function to extract text from a given region in the image
def extract_text_from_region(image_path, coordinates, lang='eng'):
    img = cv2.imread(image_path)
    region = img[coordinates[1]:coordinates[3], coordinates[0]:coordinates[2]]
    cv2.imwrite("temp.png", region)
    image = cv2.imread("temp.png")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    text = tr.image_to_string(gray, lang=lang, config='--psm 6')
    return text.strip()

for i in range(count):
    # Define image path
    original_image_path = "list_image.jpg"

    # Define coordinates for each region
    name_coords = (107, y1, 221, y2)
    id_number_coords = (240, y1, 430, y2)
    address_coords = (450, y1, 960, y2)
    phone_number_coords = (980, y1, 1180, y2)

    # Extract information from each region
    name = extract_text_from_region(original_image_path, name_coords, lang='kor')
    id_number = extract_text_from_region(original_image_path, id_number_coords, lang='eng')
    address = extract_text_from_region(original_image_path, address_coords, lang='kor')
    phone_number = extract_text_from_region(original_image_path, phone_number_coords, lang='eng')

    # Print the extracted information
    print("Name:", name)
    print("ID Number:", id_number)
    print("Address:", address)
    print("Phone Number:", phone_number)

    # Create a DataFrame
    data = {'Name': [name], 'ID Number': [id_number], 'Address': [address], 'Phone Number': [phone_number]}
    df = pd.DataFrame(data)

    # Export to Excel
    df.to_excel("output.xlsx", index=False)
    y1 += y_plus
    y2 += y_plus
    
    print(count)