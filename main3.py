import requests
import re
import os
import pandas as pd

RUTA_EXCEL = r"C:\Users\RAP2\Desktop\Proyectos varios\Proyectos Python\Lector_pdf\processed\datos.xlsx" 
CARPETA_ARCHIVOS = r"C:\Users\RAP2\Desktop\Proyectos varios\Proyectos Python\Lector_pdf\invoices"

def extract_text_with_ocr_space(file_path, api_key='K81896350488957'):
    url = 'https://api.ocr.space/parse/image'
    
    with open(file_path, 'rb') as f:
        response = requests.post(
            url,
            files={'file': f},
            data={
                'apikey': api_key,
                'language': 'spa',  # cambia a español si el documento lo es
                'isOverlayRequired': False,
                'OCREngine': 2
            }
        )

    try:
        result = response.json()  # intenta convertir a JSON
    except ValueError:
        print("❌ Error: La respuesta no es JSON. Contenido recibido:")
        print(response.text)  # imprime el contenido crudo
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
                        linea_limpia = re.sub(r'^(Ministerio\s+de\s+Gobernación|Gobernación)\s*', '', siguiente_linea, flags=re.IGNORECASE)
                        unidad_emisora.append(linea_limpia)
            break

    return ' '.join(unidad_emisora).strip() if unidad_emisora else None

def extract_fields(text):
    fields = {}

    # Número de oficio
    match = re.search(r'OFICIO\s+No\.\s+([0-9\-]+)', text, re.IGNORECASE)
    fields['oficio'] = match.group(1) if match else None

    # Referencia: REF. / Ref. / Ref./
    match = re.search(r'\bREF[./]?\s*[:\-]?\s*([A-Z0-9\-\/]+)', text, re.IGNORECASE)
    fields['referencia'] = match.group(1) if match else None

    # Asunto
    match = re.search(r'ASUNTO\s*:?[\s]*([^\n]+)(?:\n([^\n]+))?', text, re.IGNORECASE)
    if match:
        asunto = match.group(1).strip()
        if match.group(2):
            asunto += " " + match.group(2).strip()
        fields['asunto'] = asunto
    else:
        fields['asunto'] = None

    # Emision
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
    
    # Fecha/hora de recepcion
    match = re.search(r'el día .*?(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})\s+a\s+las\s+(\d{1,2}:\d{2})', text, re.IGNORECASE)
    if match:
        fields['fecha_entrega'] = match.group(1)
        fields['hora_entrega'] = match.group(2)
    else:
        fields['fecha_entrega'] = None
        fields['hora_entrega'] = None

    return fields

def guardar_en_excel(fields, ruta_excel=RUTA_EXCEL):
    if os.path.exists(ruta_excel):
        df = pd.read_excel(ruta_excel)
    else:
        df = pd.DataFrame(columns=["archivo", "oficio", "referencia", "asunto", "unidad_emisora", "firmante", "fecha_entrega", "hora_entrega"])

    # Evitar duplicados por archivo
    if fields.get("archivo") in df["archivo"].values:
        # Actualizar fila existente
        mask = df["archivo"] == fields["archivo"]
        df.loc[mask, fields.keys()] = pd.DataFrame([fields]).values
    else:
        # Agregar nuevo registro
        df = pd.concat([df, pd.DataFrame([fields])], ignore_index=True)


    df.to_excel(ruta_excel, index=False)
    print(f"✅ Datos guardados/actualizados en {ruta_excel}")

def main():
    for nombre_archivo in os.listdir(CARPETA_ARCHIVOS):
        ruta_completa = os.path.join(CARPETA_ARCHIVOS, nombre_archivo)
        
        if not os.path.isfile(ruta_completa):
            continue  # saltar si no es archivo
        
        # Opcional: filtrar solo PDFs o imágenes
        if not nombre_archivo.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png', '.tiff')):
            continue
        
        print(f"Procesando: {nombre_archivo}")
        texto = extract_text_with_ocr_space(ruta_completa)
        if texto:
            campos = extract_fields(texto)
            campos["archivo"] = nombre_archivo
            guardar_en_excel(campos)
        else:
            print(f"⚠️ No se pudo extraer texto de {nombre_archivo}")

if __name__ == "__main__":
    main()



        