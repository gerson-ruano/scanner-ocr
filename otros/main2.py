import os
import re
import PyPDF2
import pytesseract
from pdf2image import convert_from_path

def extract_invoice_info(pdf_file_path):
    text = ""

    with open(pdf_file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            page_text = page.extract_text()
            if page_text:
                text += page_text
            else:
                # Fallback a OCR
                
                print(f"⚠️ Página {page_num + 1} no tiene texto embebido. Usando OCR...")
                images = convert_from_path(pdf_file_path, first_page=page_num + 1, last_page=page_num + 1)
                text += pytesseract.image_to_string(images[0], lang='eng')

    print("📝 Texto completo extraído:")
    print(text)

# Ejecutar
extract_invoice_info("invoices/invoice3.pdf")

        