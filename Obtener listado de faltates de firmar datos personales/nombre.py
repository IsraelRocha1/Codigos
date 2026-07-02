import pandas as pd
from pathlib import Path
import unicodedata
import os 
from openpyxl import load_workbook

# -----------------------
# CONFIGURACIÓN
# -----------------------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

CSV_PATH = os.path.join(BASE_DIR, "Archivos descargados desde hourglass", "hourglass-contactlist.csv")

PDF_DIR  = Path(r"C:\Users\israe\OneDrive\Congregacion\Hermanos\Tratamiento de Datos Personales")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT = os.path.join(SCRIPT_DIR, "personas_sin_pdf.xlsx")

# -----------------------
# NORMALIZACIÓN
# -----------------------
def normalizar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().lower()
    texto = " ".join(texto.split())
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('utf-8')
    return texto

# -----------------------
# CARGAR CSV
# -----------------------
df = pd.read_csv(CSV_PATH)
df["fullname_norm"] = df["fullname"].apply(normalizar)

# -----------------------
# PDF EXISTENTES
# -----------------------
pdf_files = {normalizar(p.stem) for p in PDF_DIR.glob("*.pdf")}

# -----------------------
# PDFs QUE SOBRAN (NO ESTÁN EN CSV)
# -----------------------
csv_nombres = set(df["fullname_norm"])

pdf_sobrantes = pdf_files - csv_nombres
if pdf_sobrantes:
    print("\n📌 PDFs que SOBRAN (no están en CSV):")
    for i, pdf in enumerate(sorted(pdf_sobrantes), 1):
        print(f"{i}. {pdf}")
    
# -----------------------
# VERIFICAR
# -----------------------
df["tiene_pdf"] = df["fullname_norm"].isin(pdf_files)

faltantes = df[df["tiene_pdf"] == False][["fullname", "group_overseer"]]

# -----------------------
# EXPORTAR EXCEL
# -----------------------
faltantes.to_excel(OUTPUT, index=False)

# -----------------------
# AJUSTE DE COLUMNAS
# -----------------------
wb = load_workbook(OUTPUT)
ws = wb.active

# Ajustar ancho de columnas automáticamente
for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter  # letra de la columna
    
    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass

    # factor de ajuste (puedes subirlo si quieres más ancho)
    ws.column_dimensions[col_letter].width = max_length + 2

wb.save(OUTPUT)

print(f"✔ Listo. Generado: {OUTPUT}")
print(f"Total faltantes: {len(faltantes)}")