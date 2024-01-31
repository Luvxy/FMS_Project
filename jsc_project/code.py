# name changer
forward = """
    // Find the input element by its ID
    var myInput = document.getElementById("
"""

behind = """
");

    // Change the value of the input
    myInput.value = 
"""

value = input("value: ")
id_sum = input("id: ")

text = forward + id_sum + behind + value + ";"

print(text)

import cv2
import numpy as np
import pyautogui
import time

def find_target_image(target_image_path, screenshot_path):
    try:
        # Load the target image and the screenshot
        target_image = cv2.imread(target_image_path)
        screenshot = cv2.imread(screenshot_path)

        if target_image is not None and screenshot is not None:
            # Convert both images to grayscale
            target_image_gray = cv2.cvtColor(target_image, cv2.COLOR_BGR2GRAY)
            screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

            # Use cv2.matchTemplate
            result = cv2.matchTemplate(screenshot_gray, target_image_gray, cv2.TM_CCOEFF_NORMED)

            # Get the location of the best match
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            top_left = max_loc

            # Return the top-left coordinates
            return top_left
        else:
            print(f"Failed to read images. Target or screenshot image is None.")
            return None

    except Exception as e:
        print(f"Error: {e}")
        return None

# Specify the path to the target image file and the screenshot
target_image_path = r'target_image.png'
screenshot_path = r'Screenshot.png'

# Switch to the target window
pyautogui.hotkey('alt', 'tab')
time.sleep(1)

# Save a screenshot
screenshot = pyautogui.screenshot()
screenshot_np = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
cv2.imwrite(screenshot_path, screenshot_np)

# Find the target image and get its position
image_position = find_target_image(target_image_path, screenshot_path)

if image_position:
    print(f"Target image found at position: {image_position}")
else:
    print("Target image not found.")

pyautogui.moveTo(image_position)