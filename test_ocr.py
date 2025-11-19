from pathlib import Path
import pdfplumber
import pytesseract

# Tell pytesseract where Tesseract is installed
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pdf_path = Path(r"C:\Users\Thoma\Downloads\IAS102_seca.pdf")

with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
    pil_image = page.to_image(resolution=300).original
    ocr_text = pytesseract.image_to_string(pil_image)

print(ocr_text)