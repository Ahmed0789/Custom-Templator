from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import openpyxl

def close_application():
    window.destroy()

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        df = openpyxl.load_workbook(file_path)
        voucher_codes = df.active
        
        data = []
        for obj in voucher_codes['A']:
            data.append({
                "name": "Connect to wireless network",
                "password": "Enter Password: \n",
                "instructions": "Open browser to redirect to a login page.",
                "code": f"{str(obj.value)}" + "\n",
                "amount": "Please keep your voucher safe, you will need it to reconnect.\nJazakAllah If you have any issues please contact - ",
                "email": "IT@jalsasalana.org.uk"
            })

        items_per_page = 10
        cards_per_row = 2
        card_width = 630
        card_height = 280
        card_margin = 10
        font_size = 18
        font_bold_size = 20
        font_header_size = 24
        font_color = (0, 0, 0)
        font_path = "arial.ttf"
        italic_font_path = "arialbi.ttf"
        bold_font_path = "arialbd.ttf"  # Path to the bold variant of the font

        card_total_width = card_width * cards_per_row + card_margin * (cards_per_row + 1)
        card_total_height = card_height * (items_per_page // cards_per_row) + card_margin * (items_per_page // cards_per_row + 1)

        num_pages = len(data) // items_per_page + 1

        for page_num in range(num_pages):
            page = Image.new("RGB", (card_total_width, card_total_height), "white")
            draw = ImageDraw.Draw(page)
            font = ImageFont.truetype(font_path, font_size)
            bold_font = ImageFont.truetype(bold_font_path, font_bold_size)
            italic_font = ImageFont.truetype(italic_font_path, font_header_size)  # Bold font

            start_index = page_num * items_per_page
            end_index = min((page_num + 1) * items_per_page, len(data))

            for i, index in enumerate(range(start_index, end_index)):
                row = i // cards_per_row
                col = i % cards_per_row

                card_x = col * (card_width + card_margin) + card_margin
                card_y = row * (card_height + card_margin) + card_margin

                card = Image.new("RGB", (card_width, card_height), "#f2f2f2")
                card_draw = ImageDraw.Draw(card)

                name = data[index]["name"]
                password = data[index]["password"]
                instructions = data[index]["instructions"]
                code = data[index]["code"]
                amount = data[index]["amount"]
                email = data[index]["email"]

                name_x = card_margin
                name_y = card_margin
                card_draw.text((name_x + 80, name_y), "Jalsa Salana UK Guest WIFI access", font=italic_font, fill="black")
                card_draw.text((name_x + 90, name_y + 40), name, font=font, fill=font_color)
                card_draw.text((name_x + 330, name_y + 40), "'Guest_internet'", font=bold_font, fill="blue")

                password_x = card_margin + 130
                password_y = card_margin + 90
                card_draw.text((password_x, password_y), password, font=font, fill=font_color)
                card_draw.text((password_x + 150, password_y), "Jalsa2023", font=bold_font, fill="blue")

                instructions_x = card_margin + 130
                instructions_y = card_margin + 120
                card_draw.text((instructions_x, instructions_y), f"{instructions}", font=font, fill=font_color)

                code_x = card_margin + 130
                code_y = card_margin + 150
                card_draw.text((code_x, code_y), "Enter the voucher code:", font=font, fill=font_color)
                card_draw.text((code_x + 210, code_y), code, font=bold_font, fill="blue")

                amount_x = card_margin + 10
                amount_y = card_margin + 200
                card_draw.text((amount_x, amount_y), amount, font=font, fill=font_color)

                email_x = card_margin + 420
                email_y = card_margin + 220
                card_draw.text((email_x, email_y), f"{email}", font=font, fill="blue")

                page.paste(card, (card_x, card_y))

            page.save(f"page_{page_num + 1}.png")

# Create the main window
window = tk.Tk()

# Create a label with a text message
message_label = tk.Label(window, text="Welcome to the Voucher Application")
message_label.pack(pady=10)

# Create a button to browse and select the Excel file
browse_button = tk.Button(window, text="Browse", command=browse_file, padx=10, pady=5)
browse_button.pack()

# Create a close button to exit the application
close_button = tk.Button(window, text="Close", command=close_application, padx=10, pady=5)
close_button.pack()

# Start the application's main event loop
window.mainloop()
