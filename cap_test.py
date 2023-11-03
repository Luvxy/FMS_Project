from PIL import ImageGrab
import cv2
import pytesseract as tr
import re

x1, y1, x2, y2 = 700, 500, 780, 515

img = ImageGrab.grab((x1, y1 + 28, x2, y2 + 28))
img.save("test.png")

image = cv2.imread("test.png")
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
text = tr.image_to_string(gray, lang='eng')

# Use a regular expression to extract numbers from the text
numbers = re.findall(r'\d+', text)

# If you want to access individual numbers, you can access them by index
if len(numbers) >= 2:
    number1 = numbers[0][2:4]
    number2 = numbers[1]
    number3 = numbers[2]
    all_number = number1+number2+number3
    print(all_number)
else:
    print("Not enough numbers found in the text.")
