import os
from config import CARPETA_ARCHIVOS, RUTA_EXCEL
from ocr_utils import extract_text_from_pdf_tesseract, extract_text_with_tesseract
from extractors import extract_fields
from excel_utils import guardar_en_excel

def main():
    if not os.path.exists(CARPETA_ARCHIVOS):
        print(f"❌ Carpeta no encontrada: {CARPETA_ARCHIVOS}")
        return

    for nombre_archivo in os.listdir(CARPETA_ARCHIVOS):
        ruta_completa = os.path.join(CARPETA_ARCHIVOS, nombre_archivo)

        if not os.path.isfile(ruta_completa):
            continue

        if nombre_archivo.lower().endswith('.pdf'):
            texto = extract_text_from_pdf_tesseract(ruta_completa)
        elif nombre_archivo.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff')):
            texto = extract_text_with_tesseract(ruta_completa)
        else:
            continue

        print(f"📄 Procesando: {nombre_archivo}")

        if texto.strip():
            campos = extract_fields(texto, nombre_archivo)
            guardar_en_excel(campos, RUTA_EXCEL)
        else:
            print(f"⚠️ No se pudo extraer texto de {nombre_archivo}")

    print("✅ Proceso completado. Revisa el archivo Excel.")

if __name__ == "__main__":
    main()
