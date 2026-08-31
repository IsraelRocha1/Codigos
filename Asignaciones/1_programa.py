import os
import re
import subprocess
import sys
import time

import pymupdf  # antes "fitz" (nombre nuevo del mismo paquete, sin warning de deprecación)
import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup
from win32com.client import Dispatch

# ========= CONFIGURA AQUÍ CADA SEMANA =========
# 1) Los HTML del programa (asignaciones/participantes) YA LOS TIENES
#    descargados a mano en esta carpeta, nombrados 1.html, 2.html, 3.html...
#    (el mismo orden en el que quieres que aparezcan en el programa final).
#    El script SOLO LOS LEE, nunca los descarga.
CARPETA_HTML_DESCARGADOS = r"C:\Users\israe\Downloads"  # <-- ajusta a tu carpeta real

# 2) ENLACES es una fuente aparte, solo para complementar con las CANCIONES
#    de cada sala (se visitan con Selenium). Deben ir en el MISMO ORDEN que
#    tus archivos 1.html, 2.html, 3.html... Si no necesitas canciones para
#    alguna sala, puedes dejar ese elemento como None.
ENLACES = [
    "https://wol.jw.org/es/wol/d/r4/lp-s/202026256",
    "https://wol.jw.org/es/wol/d/r4/lp-s/202026257",
    "https://wol.jw.org/es/wol/d/r4/lp-s/202026254",
    # agrega o quita enlaces libremente: 1, 2, 3, 6... el pipeline se adapta solo
]

# ========= RUTAS DEL PROYECTO =========
BASE = os.path.dirname(os.path.abspath(__file__))
CARPETA_RECURSOS = os.path.join(BASE, "Recursos")    # modelo1-4.xlsx, canticos_lista.xlsx
CARPETA_TABLERO = os.path.join(BASE, "Tablero")      # salidas intermedias y PDFs
os.makedirs(CARPETA_TABLERO, exist_ok=True)

ARCHIVO_CANCIONES = os.path.join(CARPETA_RECURSOS, "canticos_lista.xlsx")
RUTA_PROGRAMA_FINAL = os.path.join(BASE, "programa.pdf")

# Script adicional a correr automáticamente al terminar (déjalo en None para
# desactivarlo). Ajusta la ruta a tu 4_generar_hojas.py u otro que uses.
SCRIPT_ADICIONAL = os.path.join(BASE, "3_corregir_asignaciones.py")

TEXTO_PIE = "S-140-S 11/23"

TEMPLATES = {
    1: {"fm_rows": [11, 13, 15], "lac_rows": [19], "local_row": 18, "estudio_row": 20, "cierre_row": 23},
    2: {"fm_rows": [11, 13, 15], "lac_rows": [19, 20], "local_row": 18, "estudio_row": 21, "cierre_row": 24},
    3: {"fm_rows": [11, 13, 15, 17], "lac_rows": [21], "local_row": 20, "estudio_row": 22, "cierre_row": 25},
    4: {"fm_rows": [11, 13, 15, 17], "lac_rows": [21, 22], "local_row": 20, "estudio_row": 23, "cierre_row": 26},
}


# =====================================================================
# PASO 1a — Localizar los HTML ya descargados a mano (no se descargan aquí)
# =====================================================================
def listar_html_locales(carpeta, cantidad_enlaces):
    """Busca 1.html, 2.html, 3.html... en la carpeta, en ese orden.
    Si tienes menos HTML que enlaces (o quieres nombrarlos distinto),
    ajusta esta función o pasa la lista de rutas directamente."""
    rutas = []
    for i in range(1, cantidad_enlaces + 1):
        ruta = os.path.join(carpeta, f"{i}.html")
        if not os.path.exists(ruta):
            raise FileNotFoundError(
                f"No encontré '{ruta}'. Verifica CARPETA_HTML_DESCARGADOS y que "
                f"ya hayas descargado manualmente ese archivo."
            )
        rutas.append(ruta)
    return rutas


# =====================================================================
# PASO 1b — Complementar con las canciones (petición HTTP simple, sin
# navegador — evita cualquier bloqueo de política de aplicaciones)
# =====================================================================
def obtener_canciones(enlaces):
    """Visita cada URL solo para sacar la fecha, el tema y las 3 canciones
    de esa semana/sala (apertura, intermedia y de cierre). No tiene nada
    que ver con el HTML del programa en sí (asignaciones/participantes)."""
    df_titulos = pd.read_excel(ARCHIVO_CANCIONES)
    mapa_titulos = dict(zip(df_titulos["Canción #"], df_titulos["Título"]))
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    resultado = []
    for url in enlaces:
        if not url:
            resultado.append({"fecha": None, "tema": None, "apertura": None, "intermedia": None, "cierre": None})
            continue

        resp = requests.get(url, headers=headers, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        texto = soup.get_text("\n")

        # Fecha ("5-11 DE OCTUBRE", "31 DE AGOSTO A 6 DE SEPTIEMBRE"...) y
        # tema (libro/capítulos) aparecen siempre en ese orden, en líneas
        # consecutivas justo antes de la 1ª "Canción N". Se capturan juntos
        # en una sola expresión para no confundirlos entre sí (por separado
        # se enganchaba por error un encabezado posterior como
        # "NUESTRA VIDA CRISTIANA", o la fecha y el tema quedaban pegados).
        fecha = tema = None
        m_fecha_tema = re.search(
            r"\n\s*(\d{1,2}(?:-\d{1,2})?\s+DE\s+[A-ZÁÉÍÓÚÑ]+(?:\s+A\s+\d{1,2}\s+DE\s+[A-ZÁÉÍÓÚÑ]+)?)"
            r"\s*\n\s*([A-ZÁÉÍÓÚÑ0-9][A-ZÁÉÍÓÚÑ0-9,\-\s]{2,60}?)\s*\n\s*Canción\s+\d+",
            texto,
        )
        if m_fecha_tema:
            fecha = " ".join(m_fecha_tema.group(1).split())
            tema = " ".join(m_fecha_tema.group(2).split())
        else:
            print(f"   ⚠️  No pude reconocer la fecha/tema en {url}")
            print("       Primeras líneas de la página (para diagnosticar):")
            for linea in [l for l in texto.split("\n") if l.strip()][:12]:
                print("       |", linea)

        # Las 3 canciones de la reunión, en orden de aparición:
        # 1ª = apertura, 2ª = intermedia (inicio de Nuestra Vida Cristiana),
        # 3ª = cierre.
        canciones, vistos = [], set()
        for m in re.finditer(r"Canción\s+(\d+)", texto):
            numero = f"Canción {m.group(1)}"
            if numero in vistos:
                continue
            vistos.add(numero)
            titulo = mapa_titulos.get(numero, "")
            canciones.append(f"{numero}: {titulo}" if titulo else numero)

        apertura = canciones[0] if len(canciones) > 0 else None
        intermedia = canciones[1] if len(canciones) > 1 else None
        cierre = canciones[2] if len(canciones) > 2 else None

        resultado.append({
            "fecha": fecha, "tema": tema,
            "apertura": apertura, "intermedia": intermedia, "cierre": cierre,
        })

    return resultado


# =====================================================================
# PASO 2 — Extraer cada HTML por ROL (reemplaza toda la lógica de IDs)
# =====================================================================
def _limpiar(txt):
    return " ".join(txt.split()).strip() if txt else ""


def extraer_anio(path):
    """Saca el año del selector de fecha del HTML, ej.
    <input class="date-picker-input ..." value="2026-10-05"> -> 2026"""
    with open(path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")
    campo = soup.select_one("input.date-picker-input")
    if campo and campo.get("value"):
        m = re.match(r"(\d{4})-\d{2}-\d{2}", campo["value"])
        if m:
            return int(m.group(1))
    return None


def parse_meeting_html(path):
    with open(path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    core, tgw, fm, lac = [], [], [], []
    tiempo_pendiente = None

    for elem in soup.find_all(["div", "p"]):
        clases = elem.get("class")
        if not clases:
            continue
        clase_str = " ".join(clases)

        if elem.name == "p" and set(clases).issuperset(
            {"part-time", "text-muted", "text-end", "pe-2", "mt-1", "pt-3"}
        ):
            tiempo_pendiente = _limpiar(elem.get_text())
            continue

        if clase_str not in (
            "core_row mm_part", "tgw_row mm_part", "fm_row mm_part", "lac_row mm_part",
        ):
            continue

        title_div = elem.select_one(".part-title")
        if title_div:
            # el título puede venir en 2 líneas: el nombre de la parte y,
            # debajo, un subtítulo (método de predicación, referencia de
            # lección, etc.). Guardamos ambas cosas por separado: "titulo"
            # (solo el nombre de la parte, para imprimir) y "titulo_completo"
            # (con el subtítulo, por si se necesita para otro uso).
            titulo_lineas = [_limpiar(s) for s in title_div.get_text(separator="\n").split("\n") if _limpiar(s)]
            titulo = titulo_lineas[0] if titulo_lineas else ""
            titulo_completo = " — ".join(titulo_lineas)
            resto_div = next((d for d in elem.find_all("div", recursive=False) if d is not title_div), None)
            participantes = (
                [_limpiar(s) for s in resto_div.get_text(separator="\n").split("\n") if _limpiar(s)]
                if resto_div else []
            )
        else:
            lineas = [_limpiar(l) for l in elem.get_text(separator="\n").split("\n") if _limpiar(l)]
            if not lineas:
                continue
            titulo, participantes = lineas[0], lineas[1:]
            titulo_completo = titulo

        item = {"titulo": titulo, "titulo_completo": titulo_completo,
                "participantes": participantes, "duracion": tiempo_pendiente}
        tiempo_pendiente = None

        {"core_row mm_part": core, "tgw_row mm_part": tgw,
         "fm_row mm_part": fm, "lac_row mm_part": lac}[clase_str].append(item)

    # ---- CORE ----
    presidente = consejero = oracion_inicial = oracion_final = None
    for it in core:
        t = it["titulo"].lower()
        if "presidente" in t:
            nombres = it["participantes"]
            if nombres:
                presidente = nombres[0]
            if len(nombres) > 1:
                consejero = nombres[-1]
        elif "oraci" in t:
            nombre = it["participantes"][0] if it["participantes"] else None
            if oracion_inicial is None:
                oracion_inicial = nombre
            else:
                oracion_final = nombre

    # ---- TGW: siempre 3 partes fijas ----
    discurso = tgw[0] if len(tgw) > 0 else None
    perlas = tgw[1] if len(tgw) > 1 else None
    lectura = tgw[2] if len(tgw) > 2 else None

    lectura_biblia = {"principal": None, "auxiliar": None}
    if lectura and lectura["participantes"]:
        nombres = lectura["participantes"]
        lectura_biblia["principal"] = nombres[0]
        if len(nombres) > 1:
            lectura_biblia["auxiliar"] = nombres[-1]

    # ---- FM: "Seamos mejores maestros" (variable, 3 o 4) ----
    fm_items = []
    for it in fm:
        nombres = it["participantes"]
        principal_final, aux_final = [], []
        vistos_aux = False
        for n in nombres:
            if "sala auxiliar" in n.lower():
                vistos_aux = True
                continue
            (aux_final if vistos_aux else principal_final).append(n)
        fm_items.append({
            "titulo": it["titulo"], "duracion": it["duracion"],
            "principal": principal_final, "auxiliar": aux_final,
        })

    # ---- LAC: "Nuestra vida cristiana" (variable, 1 o 2 + estudio bíblico) ----
    lac_items = []
    estudio_biblico = {"conductor": None, "lector": None}
    for it in lac:
        t = it["titulo"].lower()
        if "estudio bíblico" in t or "estudio biblico" in t:
            if it["participantes"]:
                estudio_biblico["conductor"] = it["participantes"][0]
        elif it["titulo"].lower().startswith("lector"):
            if it["participantes"]:
                estudio_biblico["lector"] = it["participantes"][0]
        else:
            lac_items.append({
                "titulo": it["titulo_completo"], "duracion": it["duracion"],
                "nombre": it["participantes"][0] if it["participantes"] else None,
            })

    return {
        "presidente": presidente, "consejero": consejero,
        "oracion_inicial": oracion_inicial, "oracion_final": oracion_final,
        "discurso_tesoros": discurso,
        "perlas_escondidas_nombre": perlas["participantes"][0] if perlas and perlas["participantes"] else None,
        "lectura_biblia": lectura_biblia,
        "fm_items": fm_items, "lac_items": lac_items,
        "estudio_biblico": estudio_biblico,
    }


# =====================================================================
# PASO 3 — Elegir plantilla y llenarla por rol
# =====================================================================
def elegir_condicion(datos):
    n_fm, n_lac = len(datos["fm_items"]), len(datos["lac_items"])
    if n_fm > 4 or n_lac > 2:
        raise ValueError(
            f"Esta semana tiene {n_fm} partes en 'Seamos mejores maestros' y "
            f"{n_lac} en 'Nuestra vida cristiana': no hay plantilla para ese "
            f"tamaño (máximo actual: 4 y 2). Hace falta crear un modelo5.xlsx."
        )
    tabla = {(3, 1): 1, (3, 2): 2, (4, 1): 3, (4, 2): 4}
    return tabla[(3 if n_fm <= 3 else 4, 1 if n_lac <= 1 else 2)]


def llenar_plantilla(ws, datos, condicion, info_semana, anio=None):
    cfg = TEMPLATES[condicion]

    ws["I2"] = info_semana.get("fecha")
    ws["J2"] = info_semana.get("tema")
    ws["I4"] = info_semana.get("apertura")
    if anio is not None:
        ws["N2"] = anio

    ws["K2"] = datos["presidente"]
    ws["K3"] = datos["consejero"]
    ws["K4"] = datos["oracion_inicial"]

    disc = datos["discurso_tesoros"] or {}
    ws["I7"] = disc.get("titulo")
    ws["J7"] = disc.get("duracion")
    ws["K7"] = (disc.get("participantes") or [None])[0]

    ws["K8"] = datos["perlas_escondidas_nombre"]
    ws["K9"] = datos["lectura_biblia"]["principal"]
    ws["Q9"] = datos["lectura_biblia"]["auxiliar"]

    for row, item in zip(cfg["fm_rows"], datos["fm_items"]):
        ws[f"L{row}"] = item["titulo"]
        ws[f"M{row}"] = item["duracion"]
        principal, auxiliar = item["principal"] or [], item["auxiliar"] or []
        if len(principal) > 0:
            ws[f"K{row}"] = principal[0]
        if len(principal) > 1:
            ws[f"K{row + 1}"] = principal[1]
        if len(auxiliar) > 0:
            ws[f"Q{row}"] = auxiliar[0]
        if len(auxiliar) > 1:
            ws[f"Q{row + 1}"] = auxiliar[1]

    ws[f"I{cfg['local_row']}"] = info_semana.get("intermedia")

    for row, item in zip(cfg["lac_rows"], datos["lac_items"]):
        ws[f"I{row}"] = item["titulo"]
        ws[f"J{row}"] = item["duracion"]
        ws[f"K{row}"] = item["nombre"]

    est = datos["estudio_biblico"]
    ws[f"K{cfg['estudio_row']}"] = est.get("conductor")
    ws[f"K{cfg['estudio_row'] + 1}"] = est.get("lector")
    ws[f"I{cfg['cierre_row']}"] = info_semana.get("cierre")
    ws[f"K{cfg['cierre_row']}"] = datos["oracion_final"]


def generar_pdf_individual(datos, condicion, info_semana, anio, nombre_salida):
    modelo_path = os.path.join(CARPETA_RECURSOS, f"modelo{condicion}.xlsx")
    wb = openpyxl.load_workbook(modelo_path)
    ws = wb["plantilla"]
    llenar_plantilla(ws, datos, condicion, info_semana, anio)

    temporal_xlsx = os.path.join(CARPETA_TABLERO, f"temp_{nombre_salida}.xlsx")
    wb.save(temporal_xlsx)
    time.sleep(1)

    pdf_path = os.path.join(CARPETA_TABLERO, f"{nombre_salida}.pdf")
    excel = Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    wb_excel = excel.Workbooks.Open(os.path.abspath(temporal_xlsx))
    wb_excel.Worksheets("plantilla").ExportAsFixedFormat(0, os.path.abspath(pdf_path))
    wb_excel.Close(False)
    excel.Quit()
    return pdf_path


# =====================================================================
# PASO 4 — Combinar todas las páginas de a 2 por hoja A4 para imprimir
# =====================================================================
def combinar_en_pares(pdfs, salida_path, texto=TEXTO_PIE, offset_y=50, x_texto=20, y_margen_inferior=30, font_size=8):
    nuevo_pdf = pymupdf.open()
    ancho, alto = pymupdf.paper_size("a4")

    for k in range(0, len(pdfs), 2):
        par = pdfs[k:k + 2]
        pagina = nuevo_pdf.new_page(width=ancho, height=alto)

        pdf_sup = pymupdf.open(par[0])
        pagina.show_pdf_page(pymupdf.Rect(0, 0, ancho, alto), pdf_sup, 0)
        pdf_sup.close()

        if len(par) > 1:
            pdf_inf = pymupdf.open(par[1])
            offset = alto / 2 - offset_y
            pagina.show_pdf_page(pymupdf.Rect(0, offset, ancho, offset + alto), pdf_inf, 0)
            pdf_inf.close()

        pagina.insert_text((x_texto, alto - y_margen_inferior), texto, fontsize=font_size, fontname="helv", fill=(0, 0, 0))

    nuevo_pdf.save(salida_path)
    nuevo_pdf.close()
    print(f"✅ Programa final generado: {salida_path}")


def limpiar_carpeta_tablero():
    """Borra los .xlsx/.pdf de una corrida anterior para evitar que se
    mezclen con los de esta semana (ej. si esta semana hay menos HTML
    que la anterior, no queda un temp_html4.pdf viejo colándose)."""
    for nombre in os.listdir(CARPETA_TABLERO):
        ruta = os.path.join(CARPETA_TABLERO, nombre)
        if os.path.isfile(ruta) and nombre.lower().endswith((".xlsx", ".pdf")):
            os.remove(ruta)


# =====================================================================
# EJECUCIÓN
# =====================================================================
def main():
    limpiar_carpeta_tablero()

    print(f"📂 Leyendo {len(ENLACES)} HTML ya descargados en {CARPETA_HTML_DESCARGADOS}...")
    rutas_html = listar_html_locales(CARPETA_HTML_DESCARGADOS, len(ENLACES))

    print("🎵 Obteniendo fecha, tema y canciones desde los enlaces...")
    info_por_html = obtener_canciones(ENLACES)

    pdfs_generados = []
    for i, (ruta_html, info_semana) in enumerate(zip(rutas_html, info_por_html), start=1):
        print(f"\n📄 Procesando html{i}...")
        datos = parse_meeting_html(ruta_html)
        anio = extraer_anio(ruta_html)
        condicion = elegir_condicion(datos)
        print(f"   → {len(datos['fm_items'])} partes en Seamos mejores maestros, "
              f"{len(datos['lac_items'])} en Nuestra vida cristiana → condición {condicion}")
        pdf_path = generar_pdf_individual(datos, condicion, info_semana, anio, f"html{i}")
        pdfs_generados.append(pdf_path)

    print(f"\n🔗 Combinando {len(pdfs_generados)} páginas en pares para imprimir...")
    combinar_en_pares(pdfs_generados, RUTA_PROGRAMA_FINAL)

    if SCRIPT_ADICIONAL:
        print(f"\n▶️  Ejecutando {SCRIPT_ADICIONAL}...")
        subprocess.run([sys.executable, SCRIPT_ADICIONAL], check=True)


if __name__ == "__main__":
    main()