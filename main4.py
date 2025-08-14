import os
import re
import requests
import pandas as pd

# ---------- FUNCIONES DE EXTRACCIÓN ----------

def extract_text_with_ocr_space(file_path, api_key='K81896350488957'):
    url = 'https://api.ocr.space/parse/image'
    
    with open(file_path, 'rb') as f:
        response = requests.post(
            url,
            files={'file': f},
            data={
                'apikey': api_key,
                'language': 'spa',
                'isOverlayRequired': False,
                'OCREngine': 2
            }
        )

    try:
        result = response.json()
    except ValueError:
        print("❌ Error: La respuesta no es JSON. Contenido recibido:")
        print(response.text)
        return ""

    if result.get('IsErroredOnProcessing'):
        print("❌ Error de OCR:", result.get('ErrorMessage'))
        return ""

    parsed_results = result.get('ParsedResults')
    if parsed_results:
        return parsed_results[0].get('ParsedText', '')
    return ""

def extract_unidad_emisora(text):
    lines = text.splitlines()
    unidad_emisora = []

    for i, line in enumerate(lines):
        if re.search(r'Ministerio\s+', line, re.IGNORECASE):
            for j in range(1, 4):
                if i + j < len(lines):
                    siguiente_linea = lines[i + j].strip()
                    if re.search(r'OFICIO|Guatemala', siguiente_linea, re.IGNORECASE):
                        break
                    if siguiente_linea:
                        linea_limpia = re.sub(r'^Ministerio\s+de\s+Gobernación\s*', '', siguiente_linea, flags=re.IGNORECASE)
                        unidad_emisora.append(linea_limpia)
            break

    return ' '.join(unidad_emisora).strip() if unidad_emisora else None

def extract_asunto(text):
    match = re.search(r'ASUNTO\s*:?[\s]*([^\n]+)(?:\n([^\n]+))?', text, re.IGNORECASE)
    if match:
        asunto = match.group(1).strip()
        if match.group(2):
            asunto += " " + match.group(2).strip()
        return asunto
    return None

def extract_fields(text, archivo):
    fields = {}

    # Número de oficio
    match_oficio = re.search(r'OFICIO\s+No\.\s+([0-9\-]+)', text, re.IGNORECASE)
    fields['oficio'] = match_oficio.group(1) if match_oficio else None

    # Referencia: REF. / Ref. / Ref./
    match_ref = re.search(r'\bREF[./]?\s*[:\-]?\s*([A-Z0-9\-\/]+)', text, re.IGNORECASE)
    fields['referencia'] = match_ref.group(1) if match_ref else None

    # Asunto
    fields['asunto'] = extract_asunto(text)

    # Unidad emisora
    fields['unidad_emisora'] = extract_unidad_emisora(text)

    # Firmante:
    ultimas_lineas = [l.strip() for l in text.strip().split('\n') if l.strip() != ''][-20:]
    firmante = None
    for i in range(len(ultimas_lineas) - 1, 1, -1):
        if re.search(r'(LIC\.|OFICIAL|JEFE|DIRECTOR|COORDINADOR|ENCARGADO)', ultimas_lineas[i], re.IGNORECASE):
            linea1 = ultimas_lineas[i - 1]
            linea2 = ultimas_lineas[i - 2] if i >= 2 else ''
            firmante = f"{linea2} / {linea1} / {ultimas_lineas[i]}".strip(' /')
            break
    fields['firmante'] = firmante

    # Fecha/hora de entrega (opcional)
    match_fecha = re.search(r'el día .*?(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})\s+a\s+las\s+(\d{1,2}:\d{2})', text, re.IGNORECASE)
    if match_fecha:
        fields['fecha_entrega'] = match_fecha.group(1)
        fields['hora_entrega'] = match_fecha.group(2)
    else:
        fields['fecha_entrega'] = None
        fields['hora_entrega'] = None

    # Archivo original
    fields['archivo'] = archivo

    return fields

# ---------- FUNCIÓN PARA GUARDAR EN EXCEL ----------

def guardar_en_excel(nuevos_datos, archivo="datos.xlsx"):
    if not nuevos_datos:
        print("❌ No hay datos para guardar.")
        return

    if os.path.exists(archivo):
        df_existente = pd.read_excel(archivo)
        df_nuevo = pd.DataFrame([nuevos_datos])
        df = pd.concat([df_existente, df_nuevo], ignore_index=True)
        df = df.drop_duplicates(subset=['archivo'], keep='last')
    else:
        df = pd.DataFrame([nuevos_datos])

    try:
        df.to_excel(archivo, index=False)
    except Exception as e:
        print(f"❌ Error guardando archivo Excel: {e}")


# ---------- PROCESAR TODOS LOS PDF ----------

def procesar_pdfs(carpeta_pdf):
    for nombre in os.listdir(carpeta_pdf):
        if nombre.lower().endswith(".pdf"):
            ruta_pdf = os.path.join(carpeta_pdf, nombre)
            print(f"Procesando {nombre}...")
            
            texto = extract_text_with_ocr_space(ruta_pdf)
            if texto:
                datos = extract_fields(texto, nombre)
                guardar_en_excel(datos)
            else:
                print(f"No se pudo extraer texto de {nombre}")

# ---------- EJECUCIÓN ----------

if __name__ == "__main__":
    carpeta = r"C:\Users\RAP2\Desktop\Proyectos varios\Proyectos Python\Lector_pdf\invoices"
    procesar_pdfs(carpeta)
    print("Proceso completado. Revisa datos.xlsx")




        