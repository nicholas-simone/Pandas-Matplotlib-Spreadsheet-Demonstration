import base64
import io
import os
import tkinter as tk
from tkinter import PhotoImage, Canvas
from tkinter import filedialog, messagebox

import matplotlib.pyplot as plt
import pandas as pd


def load_excel_file():
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file:

        filename_only = os.path.basename(file)
        root.title(f"Loaded File: {filename_only}") # Set the title bar

        df = pd.read_excel(file)
        if 'Date' in df.columns and 'Close' in df.columns:
            plt.figure(figsize=(8, 4)) # Adjust the figure size
            plt.plot(df['Date'], df['Close'], marker='o')
            plt.title(filename_only)
            plt.xlabel('Date')
            plt.ylabel('Close Price')
            plt.grid(True)
            plt.tight_layout()
            # Convert plot to PNG Image
            buf = io.BytesIO()
            plt.savefig(buf, format='png')
            buf.seek(0)
            # Convert PNG Image to Tkinter Canvas widget
            fig_photo = PhotoImage(data=base64.b64encode(buf.read()))
            canvas.create_image(0, 0, anchor='nw', image=fig_photo)
            canvas.image = fig_photo # Keep a reference to the image

        else:
            messagebox.showerror('Error', 'Excel file must have Date and Close columns!')


# Create the main window
root = tk.Tk()
root.title('Nicholas C. Simone - Python Lab 14 (Pandas)')
root.geometry('820x480')

# Create and place the canvas
canvas = Canvas(root, width=800, height=400)
canvas.pack()

# Create and place the buttons
load_button = tk.Button(root, text='Load Excel File', command=load_excel_file)
load_button.pack()
exit_button = tk.Button(root, text='Exit', command=root.quit)
exit_button.pack()
# Run the Tkinter event loop
root.mainloop()