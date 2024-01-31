import tkinter as tk
from tkinter import simpledialog

# Create the main window
root = tk.Tk()

# Function to handle button click
def get_input():
    user_input = simpledialog.askstring("Input", "Enter something:")
    if user_input:
        result_label.config(text="You entered: " + user_input)

# Create a button
button = tk.Button(root, text="Get Input", command=get_input)
button.pack(pady=10)

# Create a label to display the result
result_label = tk.Label(root, text="")
result_label.pack()

# Run the Tkinter event loop
root.mainloop()
