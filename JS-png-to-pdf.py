import os
import tkinter as tk
from tkinter import filedialog
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import easygui

def close_application():
    window.destroy()
    
def browse_files():
    file_paths = easygui.fileopenbox(title="Select PNG files", filetypes=["*.png"], multiple=True)
    if file_paths:
        convert_to_pdf(file_paths)
    else:
        print("No files selected.")

def convert_to_pdf(file_paths):
    pdf_file = "combined.pdf"

    pdf_canvas = canvas.Canvas(pdf_file, pagesize=letter)
    for file_path in file_paths:
        try:
            image = Image.open(file_path)
            image_width, image_height = image.size
            pdf_canvas.setPageSize((image_width, image_height))
            pdf_canvas.drawImage(file_path, 0, 0, width=image_width, height=image_height)
            pdf_canvas.showPage()
        except Exception as e:
            print(f"Error processing {file_path}: {e}")

    pdf_canvas.save()
    print("PDF file 'combined.pdf' successfully created.")
    
# Create the main window
window = tk.Tk()

# Create a label with a text message
message_label = tk.Label(window, text="Please click the 'Browse' button to select the PNG files.")
message_label.pack(pady=10)

# Create a button to browse and select the Excel file
browse_button = tk.Button(window, text="Browse", command=browse_files, padx=10, pady=5)
browse_button.pack()

# Create a close button to exit the application
close_button = tk.Button(window, text="Close", command=close_application, padx=10, pady=5)
close_button.pack()

# Start the application's main event loop
window.mainloop()
