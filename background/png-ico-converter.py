import tkinter as tk
from tkinter import filedialog
from PIL import Image
import os

def create_ico():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(title="Select PNG Image", filetypes=[("PNG files", "*.png")])

    if not file_path:
        print("No file selected.")
        return 
    
    try:
        img = Image.open(file_path)

        output_path = os.path.splitext(file_path)[0] + ".ico"

        icon_sized = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

        img.save(output_path, format="ICO", sizes=icon_sized)
        print(f"ICO file created at: {output_path}")
    
    except Exception as e:
        print(f"Error creating ICO: {e}")

if __name__ == "__main__":
    create_ico()