import os
import pytesseract

# Carpeta base del proyecto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Carpetas
CARPETA_PROCESSED = os.path.join(BASE_DIR, "processed")
CARPETA_ARCHIVOS = os.path.join(BASE_DIR, "invoices")

# Archivo Excel en carpeta processed
RUTA_EXCEL = os.path.join(CARPETA_PROCESSED, "datos.xlsx")

# Crear carpeta processed si no existe
os.makedirs(CARPETA_PROCESSED, exist_ok=True)

# Configuración Tesseract (ajusta si está en otra ruta)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
print("Tesseract path:", pytesseract.pytesseract.tesseract_cmd)

