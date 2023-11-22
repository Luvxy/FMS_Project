import tkinter as tk
import O_test

def button_click():
    print("Button clicked!")

window = tk.Tk()

button = tk.Button(window, text="Click Me", command=button_click)
button.pack()

window.mainloop()
