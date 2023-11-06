# button_functions.py
import json
import pandas as pd
import FMS
from easydict import EasyDict
from tkinter import filedialog as fd

def read_config():
    try:
        with open("config.json", 'r', encoding="utf-8") as config_file:
            config = json.load(config_file)
        return config
    except FileNotFoundError:
        print("Configuration file not found. Creating a new one.")
        return {}

def write_config(config):
    with open('config.json', 'w', encoding='utf-8') as config_file:
        json.dump(config, config_file, indent=4, ensure_ascii=False)

def on_button1_clicked(): # 다음
    print("Button 1 clicked")

def on_button2_clicked(): # 이전
    print("Button 2 clicked")

def on_button3_clicked(): # 파일선택
    print("Button 3 clicked")
    config_r = read_config()
    path = ''

    filetypes = (
        ('excel files', '*.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    path = filename
    config_r['path'] = path
    
    print(config_r['path'])
    
    write_config(config_r)
    return path

# Add more functions for other buttons if needed