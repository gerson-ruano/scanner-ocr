import re

# UNIDAD EMISORA
def extract_unidad_emisora(text):
    lines = text.splitlines()
    
    # Patrones que indican el inicio de la unidad emisora
    patrones_inicio = [
        r'Ministerio\s+de\s+Gobernación',
        r'^DE:\s*$',
        r'Unidad\s+Emisora',
        r'Remitente:',
        r'De:\s+'
    ]
    
    unidad_emisora = []
    capturando = False
    
    for line in lines:
        linea_limpia = line.strip()
        
        # Verificar si encontramos un patrón de inicio
        if not capturando:
            for patron in patrones_inicio:
                if re.search(patron, linea_limpia, re.IGNORECASE):
                    capturando = True
                    # Extraer texto después del patrón si existe
                    resto = re.sub(patron, '', linea_limpia, flags=re.IGNORECASE).strip()
                    if resto:
                        unidad_emisora.append(resto)
                    break
        else:
            # Condiciones para dejar de capturar
            if re.search(r'^(OFICIO|CIRCULAR|No\.?|\d{2,}|Guatemala,|Fecha|ASUNTO|Destinatario)', 
                        linea_limpia, re.IGNORECASE):
                break
            
            if linea_limpia and not re.search(r'^\s*[-*]\s*$', linea_limpia):
                unidad_emisora.append(linea_limpia)
    
    return ' '.join(unidad_emisora).strip() if unidad_emisora else None


# Tipo y Número de oficio
def extract_oficio_info(text):
    tipos_oficio = [
        'OFICIO',
        'OFICIO CIRCULAR', 
        'OFICIO EJECUTIVO',
        'CIRCULAR',
        'MEMORANDUM'
    ]
    
    # Crear patrón dinámicamente
    tipos_pattern = '|'.join(tipos_oficio)
    pattern = rf'({tipos_pattern})\s+No\.?\s*([0-9\-\/]+)'
    
    match = re.search(pattern, text, re.IGNORECASE)
    
    if match:
        return {
            'tipo': match.group(1).title(),
            'numero': match.group(2),
            #'texto_completo': match.group(0)
        }
    return None

# Firmante
def extract_firmante(text):
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    # Buscar en las últimas 20 líneas
    seccion_final = lines[-20:] if len(lines) > 20 else lines
    
    firmante_info = []
    encontrado_patron = False
    
    # Recorrer de abajo hacia arriba
    for i in range(len(seccion_final) - 1, -1, -1):
        linea = seccion_final[i]
        
        # Patrones que indican información del firmante
        patrones_firma = [
            r'^[A-Z][a-z]+(?:\s+[A-Z][a-z]+)+\s*$',  # Nombres propios
            r'\b(LIC\.|ING\.|ARC\.|DR\.|ABOG\.)\s+[A-Z]',
            r'(DIRECTOR|JEFE|COORDINADOR|ENCARGADO|MINISTRO).*$',
            r'^[A-Z\s]{10,}$'  # Texto en mayúsculas (posible cargo)
        ]
        
        for patron in patrones_firma:
            if re.search(patron, linea, re.IGNORECASE):
                if not encontrado_patron:
                    firmante_info.insert(0, linea)
                    encontrado_patron = True
                else:
                    # Si ya encontramos un patrón, esta podría ser línea anterior
                    firmante_info.insert(0, linea)
                break
        
        # Limitar a 3 líneas máximo
        if len(firmante_info) >= 3:
            break
    
    if firmante_info:
        # Unir y limpiar
        resultado = " / ".join(firmante_info).strip()
        # Remover caracteres especiales al inicio/final
        resultado = re.sub(r'^[/\\|\s]+|[/\\|\s]+$', '', resultado)
        return resultado if resultado else None
    
    return None


def extract_fields(text, archivo):
    fields = {}

    # Tipo y Número de oficio
    resultado = extract_oficio_info(text)
    if resultado:
         fields.update(resultado)

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
    fields['firmante'] = extract_firmante(text)

    # Fecha del documento
    match = re.search(r'Guatemala,\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})', text)
    fields['fecha_documento'] = match.group(1) if match else None

    # Nombre de archivo
    fields['archivo'] = archivo

    return fields

