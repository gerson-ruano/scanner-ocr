import re

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
                        linea_limpia = re.sub(
                            r'^(Ministerio\s+de\s+Gobernación|Gobernación)\s*',
                            '',
                            siguiente_linea,
                            flags=re.IGNORECASE
                        )
                        unidad_emisora.append(linea_limpia)
            break

    return ' '.join(unidad_emisora).strip() if unidad_emisora else None


def extract_fields(text, archivo):
    fields = {}

    # Número de oficio
    match = re.search(r'OFICIO\s+No\.\s+([0-9\-]+)', text, re.IGNORECASE)
    fields['oficio'] = match.group(1) if match else None

    # Referencia
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

    # Unidad emisora
    fields['unidad_emisora'] = extract_unidad_emisora(text)

    # Firmante
    ultimas_lineas = [l.strip() for l in text.strip().split('\n') if l.strip() != ''][-20:]
    firmante = None
    for i in range(len(ultimas_lineas) - 1, 1, -1):
        if re.search(r'(LIC\.|OFICIAL|JEFE|DIRECTOR|COORDINADOR|ENCARGADO)', ultimas_lineas[i], re.IGNORECASE):
            linea1 = ultimas_lineas[i - 1]
            linea2 = ultimas_lineas[i - 2] if i >= 2 else ''
            firmante = f"{linea2} / {linea1} / {ultimas_lineas[i]}".strip(' /')
            break
        fields['firmante'] = firmante

    # Fecha del documento
    match = re.search(r'Guatemala,\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})', text)
    fields['fecha_documento'] = match.group(1) if match else None

    # Nombre de archivo
    fields['archivo'] = archivo

    return fields

