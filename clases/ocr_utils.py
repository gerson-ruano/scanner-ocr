import pytesseract
from PIL import Image
from pdf2image import convert_from_path

def extract_text_with_tesseract(file_path):
    """Extrae texto de imágenes usando Tesseract."""
    try:
        image = Image.open(file_path)
        return pytesseract.image_to_string(image, lang='spa')
    except Exception as e:
        print(f"❌ Error procesando imagen {file_path}: {e}")
        return ""

def extract_text_from_pdf_tesseract(pdf_path):
    """Convierte PDF en imágenes y extrae texto con Tesseract."""
    try:
        pages = convert_from_path(pdf_path, poppler_path=r"C:\poppler\Library\bin")
        text = ""
        for page in pages:
            text += pytesseract.image_to_string(page, lang='spa')
        return text
    except Exception as e:
        print(f"❌ Error procesando PDF {pdf_path}: {e}")
        return ""

