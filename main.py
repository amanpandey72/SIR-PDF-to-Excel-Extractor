import tkinter as tk
from tkinter import filedialog, messagebox
import pytesseract
import cv2
import re
import pandas as pd
from pdf2image import convert_from_path
import os

# ==============================
# 🔥 SET YOUR POPPLER PATH
# ==============================
POPPLER_PATH = r"C:\Users\amanp\Downloads\Release-25.12.0-0\poppler-25.12.0\Library\bin"

# ==============================
# 🔥 SET YOUR TESSERACT PATH
# ==============================
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def process_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not pdf_path:
        return

    try:
        status_label.config(text="Processing PDF... Please wait...")
        root.update()

        pages = convert_from_path(
            pdf_path,
            dpi=300,
            poppler_path=POPPLER_PATH
        )

        all_data = []
        serial_no = 1

        for page_number, page in enumerate(pages):

            image_path = f"temp_page_{page_number}.jpg"
            page.save(image_path, "JPEG")

            img = cv2.imread(image_path)
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

            # Improve OCR Quality
            gray = cv2.GaussianBlur(gray, (5, 5), 0)
            gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]

            # 🔥 Better for multi-column layout
            text = pytesseract.image_to_string(
                gray,
                config='--oem 3 --psm 4',
                lang="eng+guj"
            )

            # ==============================
            # 🔥 CUSTOM EXTRACTION LOGIC
            # ==============================

            # Split using EPIC number pattern
            voter_blocks = re.split(r'(?=[A-Z]{3}\d{7})', text)

            for block in voter_blocks:

                epic_match = re.search(r'[A-Z]{3}\d{7}', block)

                if epic_match:

                    name_match = re.search(r'નામ[:\s]*([^\n]+)', block)
                    father_match = re.search(r'(પિતા|પતિ)[^\n]*[:\s]*([^\n]+)', block)
                    house_match = re.search(r'ઘર[^\n]*[:\s]*([^\n]+)', block)
                    age_match = re.search(r'ઉંમર[:\s]*(\d+)', block)
                    gender_match = re.search(r'(પુરુષ|સ્ત્રી)', block)

                    data = {
                        "SerialNo": serial_no,
                        "NAME": name_match.group(1).strip() if name_match else "",
                        "FATHER_NAME": father_match.group(2).strip() if father_match else "",
                        "EPNO": epic_match.group(0),
                        "ADDRESS": house_match.group(1).strip() if house_match else "",
                        "AGE": age_match.group(1).strip() if age_match else "",
                        "GENDER": gender_match.group(1).strip() if gender_match else ""
                    }

                    all_data.append(data)
                    serial_no += 1

            os.remove(image_path)

        if not all_data:
            messagebox.showwarning("Warning", "No data extracted. Check PDF quality.")
            status_label.config(text="No data extracted.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if output_path:
            df = pd.DataFrame(all_data)
            df.to_excel(output_path, index=False)

            messagebox.showinfo("Success", f"{len(all_data)} records extracted successfully!")
            status_label.config(text="Extraction Completed Successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        status_label.config(text="Error occurred.")


# ==============================
# GUI DESIGN
# ==============================

root = tk.Tk()
root.title("SIR PDF to Excel Extractor (Final Version)")
root.geometry("500x240")
root.resizable(False, False)

title_label = tk.Label(
    root,
    text="SIR Final Report Extractor\n(Gujarati + English | 3-Column Layout Supported)",
    font=("Arial", 12, "bold")
)
title_label.pack(pady=15)

btn = tk.Button(
    root,
    text="Select SIR PDF & Extract",
    command=process_pdf,
    width=32,
    height=2
)
btn.pack(pady=10)

status_label = tk.Label(root, text="Ready", fg="blue")
status_label.pack(pady=10)

root.mainloop()