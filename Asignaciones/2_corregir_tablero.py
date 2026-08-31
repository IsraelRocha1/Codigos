import pymupdf
import os
import glob
import re
from itertools import zip_longest
from win32com.client import Dispatch

# ========= RUTAS =========
base_proyecto = os.path.dirname(os.path.abspath(__file__))
carpeta_Tablero = os.path.join(base_proyecto, "Tablero")

# ========= BUSCAR Y ORDENAR ARCHIVOS EXCEL =========
patron_busqueda = os.path.join(carpeta_Tablero, "temp_html*.xlsx")
todos_los_excels = glob.glob(patron_busqueda)

def extraer_numero(ruta):
    nombre = os.path.basename(ruta)
    numeros = re.findall(r'\d+', nombre)
    return int(numeros[0]) if numeros else 0

todos_los_excels.sort(key=extraer_numero)

pdf_superiores = []
pdf_inferiores = []

# ========= CONVERTIR EXCEL A PDF =========
# Modificamos la función para que reciba la instancia de Excel abierta
def excel_a_pdf(excel_app, ruta_excel, ruta_pdf):
    wb = excel_app.Workbooks.Open(os.path.abspath(ruta_excel))
    wb.Worksheets(1).ExportAsFixedFormat(0, os.path.abspath(ruta_pdf))
    wb.Close(False)

# ABRIMOS EXCEL UNA SOLA VEZ AQUÍ
print("Iniciando Excel...")
excel = Dispatch("Excel.Application")
excel.DisplayAlerts = False
excel.Visible = False

try:
    for i, ruta_excel in enumerate(todos_los_excels):
        nombre = os.path.basename(ruta_excel) 
        num = nombre.replace("temp_html", "").replace(".xlsx", "") 
        ruta_pdf = os.path.join(carpeta_Tablero, f"html{num}.pdf")
        
        # Le pasamos la instancia "excel" que ya está abierta
        excel_a_pdf(excel, ruta_excel, ruta_pdf)
        
        if i % 2 == 0:
            pdf_superiores.append(ruta_pdf)
        else:
            pdf_inferiores.append(ruta_pdf)
            
finally:
    # CERRAMOS EXCEL UNA SOLA VEZ AL FINAL
    # El bloque try/finally asegura que Excel se cierre incluso si hay un error
    excel.Quit()
    print("Conversión completada. Excel cerrado.")

# ========= COMBINAR PARES (USANDO PYMUPDF) =========
def combinar_pares(pdf_superiores, pdf_inferiores, salida_path,
                   texto="S-140-S 11/23",
                   offset_y=50,
                   x_texto=20,
                   y_margen_inferior=30,
                   font_size=8):
                   
    nuevo_pdf = pymupdf.open()
    ancho, alto = pymupdf.paper_size("a4")

    for sup_path, inf_path in zip_longest(pdf_superiores, pdf_inferiores):
        pdf_sup = pymupdf.open(sup_path)

        pagina = nuevo_pdf.new_page(width=ancho, height=alto)
        pagina.show_pdf_page(pymupdf.Rect(0, 0, ancho, alto), pdf_sup, 0)

        if inf_path:
            pdf_inf = pymupdf.open(inf_path)
            offset = alto / 2 - offset_y
            pagina.show_pdf_page(pymupdf.Rect(0, offset, ancho, offset + alto), pdf_inf, 0)
            pdf_inf.close()

        y_texto = alto - y_margen_inferior
        pagina.insert_text(
            pymupdf.Point(x_texto, y_texto),
            texto,
            fontsize=font_size,
            fontname="helv",
            color=(0, 0, 0),
        )

        pdf_sup.close()

    nuevo_pdf.save(salida_path)
    nuevo_pdf.close()
    print(f"\n✅ PDF final generado: {salida_path}")

ruta_programa = os.path.join(base_proyecto, "programa.pdf")
combinar_pares(pdf_superiores, pdf_inferiores, ruta_programa)