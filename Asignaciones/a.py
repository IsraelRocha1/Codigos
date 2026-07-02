from bs4 import BeautifulSoup
import pandas as pd

# =========================
# CONFIGURACIÓN
# =========================

ruta_html = r"C:\Users\israe\Downloads\2.html"

salida_excel = r"C:\Users\israe\Documents\Codigos\Asignaciones\spans_extraidos.xlsx"

clases = [
    "core_row mm_part",
    "tgw_row mm_part",
    "fm_row mm_part",
    "lac_row mm_part"
]

# =========================
# LEER HTML
# =========================

with open(ruta_html, "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, "html.parser")

# =========================
# EXTRAER TEXTO VISIBLE
# =========================

datos = []

for clase in clases:
    
    elementos = soup.find_all(class_=clase)

    for i, elemento in enumerate(elementos, start=1):

        # Extrae TODO el texto visible interno
        texto = elemento.get_text(separator=" ", strip=True)

        datos.append({
            "Clase": clase,
            "Indice": i,
            "Texto visible": texto
        })

# =========================
# EXPORTAR A EXCEL
# =========================

df = pd.DataFrame(datos)

with pd.ExcelWriter(salida_excel, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Texto")

    ws = writer.sheets["Texto"]

    # Ajustar ancho automático
    for columna in ws.columns:
        max_len = 0
        letra = columna[0].column_letter

        for celda in columna:
            try:
                if len(str(celda.value)) > max_len:
                    max_len = len(str(celda.value))
            except:
                pass

        ws.column_dimensions[letra].width = min(max_len + 5, 100)

print(f"Excel generado en:\n{salida_excel}")