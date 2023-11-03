# button_functions.py
import json
import pandas as pd
from tkinter import filedialog as fd

def read_config():
    # Read the configuration from the JSON file
    with open('config.json', 'r') as config_file:
        config = json.load(config_file)
    return config

def write_config(config):
    with open('config.json', 'w') as config_file:
        json.dump(config, 'config.json', indent=4)

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
    
    write_config(config_r)

# Add more functions for other buttons if needed
