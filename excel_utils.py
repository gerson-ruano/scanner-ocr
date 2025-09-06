import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

def guardar_en_excel(fields, ruta_excel):
    # Si existe, leerlo, si no crear DataFrame vacío
    if os.path.exists(ruta_excel):
        df = pd.read_excel(ruta_excel)
    else:
        df = pd.DataFrame(columns=[
            "archivo", "oficio", "referencia", "asunto",
            "unidad_emisora", "firmante", "fecha_documento"
        ])

    # Actualizar o agregar
    if fields["archivo"] in df["archivo"].values:
        mask = df["archivo"] == fields["archivo"]
        for col in fields:
            df.loc[mask, col] = fields[col]
    else:
        df = pd.concat([df, pd.DataFrame([fields])], ignore_index=True)

    # Guardar Excel con pandas
    df.to_excel(ruta_excel, index=False)

    # --- Ajustar tamaños de columna y wrap text ---
    try:
        wb = load_workbook(ruta_excel)
        ws = wb.active

        col_widths = {
            "A": 30,  # archivo
            "B": 15,  # oficio
            "C": 20,  # referencia
            "D": 50,  # asunto
            "E": 55,  # unidad_emisora
            "F": 40,  # firmante
            "G": 20,  # fecha_documento
        }

        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                if cell.column_letter in ["D", "E", "F"]:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

        wb.save(ruta_excel)
        wb.close()
    except Exception as e:
        print(f"⚠️ No se pudo ajustar el formato de columnas: {e}")

    print(f"✅ Datos guardados/actualizados en {ruta_excel}")

