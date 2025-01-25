import tkinter as tk
from tkinter import messagebox

def show_popup():
    messagebox.showinfo("Message", "Hello World")

# Create main window
root = tk.Tk()
root.title("Simple GUI")
root.geometry("300x300")

# Create a button
button = tk.Button(root, text="Click Me!", command=show_popup)
button.pack(pady=20)

# Run the application
root.mainloop()