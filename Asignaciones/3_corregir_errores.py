import fitz  # PyMuPDF
import os
from win32com.client import Dispatch

# ========= RUTAS =========
base_proyecto = os.path.dirname(os.path.abspath(__file__))
carpeta_Tablero = os.path.join(base_proyecto, "Tablero")

# Temp Excel (en Tablero)
excels_superiores = [
    os.path.join(carpeta_Tablero, "temp_html1.xlsx"),
    os.path.join(carpeta_Tablero, "temp_html3.xlsx"),
]
excels_inferiores = [
    os.path.join(carpeta_Tablero, "temp_html2.xlsx"),
    os.path.join(carpeta_Tablero, "temp_html4.xlsx"),
]

# Nombres de PDF (quieres html1..4.pdf)
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

# Convertir superiores (temp_html1 -> html1.pdf, temp_html3 -> html3.pdf)
for ruta_excel in excels_superiores:
    nombre = os.path.basename(ruta_excel)  # temp_html1.xlsx
    num = nombre.replace("temp_html", "").replace(".xlsx", "")  # 1
    ruta_pdf = os.path.join(carpeta_Tablero, f"html{num}.pdf")
    excel_a_pdf(ruta_excel, ruta_pdf)
    pdf_superiores.append(ruta_pdf)

# Convertir inferiores (temp_html2 -> html2.pdf, temp_html4 -> html4.pdf)
for ruta_excel in excels_inferiores:
    nombre = os.path.basename(ruta_excel)
    num = nombre.replace("temp_html", "").replace(".xlsx", "")
    ruta_pdf = os.path.join(carpeta_Tablero, f"html{num}.pdf")
    excel_a_pdf(ruta_excel, ruta_pdf)
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

    for sup_path, inf_path in zip(pdf_superiores, pdf_inferiores):
        pdf_sup = fitz.open(sup_path)
        pdf_inf = fitz.open(inf_path)

        pagina = nuevo_pdf.new_page(width=ancho, height=alto)
        pagina.show_pdf_page(fitz.Rect(0, 0, ancho, alto), pdf_sup, 0)

        offset = alto / 2 - offset_y
        pagina.show_pdf_page(fitz.Rect(0, offset, ancho, offset + alto), pdf_inf, 0)

        y_texto = alto - y_margen_inferior
        pagina.insert_text(
            (x_texto, y_texto),
            texto,
            fontsize=font_size,
            fontname="helv",
            fill=(0, 0, 0),
        )

        pdf_sup.close()
        pdf_inf.close()

    nuevo_pdf.save(salida_path)
    nuevo_pdf.close()
    print(f"\n✅ PDF final generado: {salida_path}")

# programa.pdf en la raíz del proyecto
ruta_programa = os.path.join(base_proyecto, "programa.pdf")
combinar_pares(pdf_superiores, pdf_inferiores, ruta_programa)
