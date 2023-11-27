import pyautogui
import numpy as np
import time
import pandas as pd
import threading
import tkinter as tk
import tkinter.ttk as ttk
import pyautogui
import pyperclip
import cv2
import pytesseract as tr
import re
from tkinter import filedialog as fd
from tkinter import messagebox
import datetime
from tkinter import simpledialog
from PIL import ImageGrab
import traceback


is_match = False
#Point(x=1144, y=212)
start_time = time.time()
green_color = (27, 119, 228)  # (R, G, B) values for pure green
target_pixel = (1159, 196)
box_size = 10
while True:
    box = (
        target_pixel[0] - box_size,
        target_pixel[1] - box_size,
        target_pixel[0] + box_size,
        target_pixel[1] + box_size
    )
    # Capture only the region around the target pixel
    screenshot = ImageGrab.grab(bbox=box)

    # Get the pixel color at the center of the bounding box
    center_pixel_color = screenshot.getpixel((box_size, box_size))

    if center_pixel_color == green_color:
        print("Green color found at ({}, {})".format(*target_pixel))
        is_match = True
        break

    # # Search for the green color in the screenshot
    # for x in range():
    #     for y in range(center_pixel_color.height):
    #         pixel_color = center_pixel_color.getpixel(target_pixel)
    #         if pixel_color == green_color:
    #             # Green color found, click a button (you can modify this action)
    #             print("Green color found at ({}, {})".format(*target_pixel))
    #             green_found = True
    #             is_match = True
    #             break  # Exit the inner loop

    #     if green_found:
    #         break  # Exit the outer loop

    # if green_found:
    #         break  # Exit the outer loop

    # Check if 3 seconds have passed
    if time.time() - start_time >= 3:
        break

if is_match:
    time.sleep(1)
    print("Match found!")

else:
    print("Match not found!")