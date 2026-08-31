import fitz  # PyMuPDF
import os
import glob
import re
from itertools import zip_longest
from win32com.client import Dispatch

# ========= RUTAS =========
base_proyecto = os.path.dirname(os.path.abspath(__file__))
carpeta_Tablero = os.path.join(base_proyecto, "Tablero")

# ========= BUSCAR Y ORDENAR ARCHIVOS EXCEL =========
# Buscar todos los archivos que coincidan con temp_html*.xlsx
patron_busqueda = os.path.join(carpeta_Tablero, "temp_html*.xlsx")
todos_los_excels = glob.glob(patron_busqueda)

# Función para extraer el número del archivo y ordenar correctamente (1, 2, 3... 10, 11)
def extraer_numero(ruta):
    nombre = os.path.basename(ruta)
    numeros = re.findall(r'\d+', nombre)
    return int(numeros[0]) if numeros else 0

todos_los_excels.sort(key=extraer_numero)

pdf_superiores = []
pdf_inferiores = []

# ========= CONVERTIR EXCEL A PDF =========
def excel_a_pdf(ruta_excel, ruta_pdf):
    excel = Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(ruta_excel))
    wb.Worksheets(1).ExportAsFixedFormat(0, os.path.abspath(ruta_pdf))
    wb.Close(False)
    excel.Quit()

# Convertir todos los excels encontrados y repartirlos dinámicamente
for i, ruta_excel in enumerate(todos_los_excels):
    nombre = os.path.basename(ruta_excel) 
    num = nombre.replace("temp_html", "").replace(".xlsx", "") 
    ruta_pdf = os.path.join(carpeta_Tablero, f"html{num}.pdf")
    
    excel_a_pdf(ruta_excel, ruta_pdf)
    
    # Si el índice es par (0, 2, 4...) es el 1º, 3º, 5º archivo -> Superior
    if i % 2 == 0:
        pdf_superiores.append(ruta_pdf)
    # Si el índice es impar (1, 3, 5...) es el 2º, 4º, 6º archivo -> Inferior
    else:
        pdf_inferiores.append(ruta_pdf)

# ========= COMBINAR PARES =========
def combinar_pares(pdf_superiores, pdf_inferiores, salida_path,
                   texto="S-140-S 11/23",
                   offset_y=50,
                   x_texto=20,
                   y_margen_inferior=30,
                   font_size=8):
    nuevo_pdf = fitz.open()
    ancho, alto = fitz.paper_size("a4")

    # zip_longest empareja hasta la lista más larga, rellenando con None si falta uno (impar)
    for sup_path, inf_path in zip_longest(pdf_superiores, pdf_inferiores):
        pdf_sup = fitz.open(sup_path)

        pagina = nuevo_pdf.new_page(width=ancho, height=alto)
        pagina.show_pdf_page(fitz.Rect(0, 0, ancho, alto), pdf_sup, 0)

        # Solo procesar e insertar el PDF inferior si existe
        if inf_path:
            pdf_inf = fitz.open(inf_path)
            offset = alto / 2 - offset_y
            pagina.show_pdf_page(fitz.Rect(0, offset, ancho, offset + alto), pdf_inf, 0)
            pdf_inf.close()

        y_texto = alto - y_margen_inferior
        pagina.insert_text(
            (x_texto, y_texto),
            texto,
            fontsize=font_size,
            fontname="helv",
            fill=(0, 0, 0),
        )

        pdf_sup.close()

    nuevo_pdf.save(salida_path)
    nuevo_pdf.close()
    print(f"\n✅ PDF final generado: {salida_path}")

# programa.pdf en la raíz del proyecto
ruta_programa = os.path.join(base_proyecto, "programa.pdf")
combinar_pares(pdf_superiores, pdf_inferiores, ruta_programa)