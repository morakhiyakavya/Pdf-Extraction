import cv2
import fitz  # PyMuPDF

# CONFIG
pdf_path = "DDx Tabelle.pdf"
img_path = r"C:\Users\kavya\Documents\My_programming\upwork\page13_embedded_images\page13_img1.png"
page_number = 29  # Page 30 in 0-based



# Step 2: Map image coordinates back to PDF space
doc = fitz.open(pdf_path)
print(f"Page {page_number + 1} loaded.")
for page_num in range(20, 40):
    print(f"Processing page {page_num + 1}...")
    page = doc.load_page(page_num)
    blocks = page.get_text("dict")["blocks"]
    for block in blocks:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"].strip()
                print(f"Text: '{text}' | font size: {span['size']} | font flags: {span['flags']} | font: {span['font']}")
                if "Malnutrition (selten in den entwickelten Ländern; weltweit verbreitet)" in text:
                    print(f"Page {page_num + 1}: '{text}' | font size: {span['size']} | font flags: {span['flags']} | font: {span['font']}")