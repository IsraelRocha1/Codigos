import fitz  # PyMuPDF
import os
import re
import unicodedata
import csv
import json
import shutil

# =========================
# CONFIGURACIÓN
# =========================

import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
#            ↑ sube de scripts/ → tu_proyecto/

CSV_CONTACTOS = os.path.join(BASE_DIR, "Archivos descargados desde hourglass", "hourglass-contactlist.csv")
JSON_GRUPOS = os.path.join(BASE_DIR, "Archivos descargados desde hourglass", "grupos.json")
PDF_UNICO = os.path.join(BASE_DIR, "Archivos descargados desde hourglass", "S-21.pdf")
PDF_TOTALES = os.path.join(BASE_DIR, "Archivos descargados desde hourglass", "S-21_2026.pdf")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "Registros_Publicadores")


# =========================
# NORMALIZACIÓN
# =========================

def normalizar(texto):
    texto = texto.lower()
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    return texto.strip()

# =========================
# LIMPIAR SOLO PARA ARCHIVO
# =========================

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|]', '', nombre).strip()

# =========================
# PALABRAS A IGNORAR
# =========================

PALABRAS_IGNORAR = {"de", "del", "la", "los"}

def limpiar_palabras(texto):
    return [p for p in texto.split() if p not in PALABRAS_IGNORAR]

# =========================
# EXTRAER NOMBRE PDF
# =========================

def formatear_nombre(nombre):
    if "," in nombre:
        partes = nombre.split(",")
        apellido = partes[0].strip()
        resto = partes[1].strip()
        return f"{resto} {apellido}"
    return nombre

def extraer_nombre_pagina(texto):
    for l in texto.split("\n"):
        l = l.strip()
        if "nombre:" in l.lower():
            nombre = l.split(":", 1)[1].strip()
            return formatear_nombre(nombre)
    return "Desconocido"

# =========================
# CARGAR CONTACTOS
# =========================

def cargar_contactos(csv_path):
    contactos = {}

    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)

        for row in reader:
            firstname = row.get('firstname', '').strip()
            middlename = row.get('middlename', '').strip()
            lastname = row.get('lastname', '').strip()

            clave = normalizar(
                " ".join(filter(None, [lastname, firstname, middlename]))
            )

            contactos[clave] = {
                "fullname": row.get("fullname", "").strip(),
                "overseer": row.get("group_overseer", "").strip(),
                "status": row.get("status", "").strip(),
                "inactive": row.get("inactive", "").strip()
            }

    return contactos

# =========================
# CARGAR GRUPOS
# =========================

def cargar_grupos(json_path):
    with open(json_path, encoding='utf-8') as f:
        data = json.load(f)
    
    # SOLO ancianos
    ancianos = data.get("ancianos", {})

    return {normalizar(k): v for k, v in ancianos.items()}

# =========================
# MATCHING INTELIGENTE
# =========================

def encontrar_contacto(nombre_pdf, contactos):
    nombre_pdf_norm = normalizar(nombre_pdf)
    palabras_pdf = set(limpiar_palabras(nombre_pdf_norm))

    mejor_match = None
    mejor_score = 0

    for clave, data in contactos.items():
        palabras_csv = set(limpiar_palabras(clave))

        score = len(palabras_pdf.intersection(palabras_csv))

        if score > mejor_score:
            mejor_score = score
            mejor_match = data

    if mejor_score >= 2:
        return mejor_match

    return None

# =========================
# CREAR CARPETAS
# =========================

def crear_estructura():
    #  Borrar carpeta completa si ya existe
    if os.path.exists(OUTPUT_DIR):
        print("🧹 Eliminando carpeta anterior...")
        shutil.rmtree(OUTPUT_DIR)

    #  Crear estructura limpia
    os.makedirs(f"{OUTPUT_DIR}/Activos/Precursores", exist_ok=True)
    os.makedirs(f"{OUTPUT_DIR}/Activos/Publicadores", exist_ok=True)
    os.makedirs(f"{OUTPUT_DIR}/Inactivos", exist_ok=True)

# =========================
# GUARDAR PDF
# =========================
def guardar_persona(nombre_pdf, paginas, contactos, mapa_grupos):
    contacto = encontrar_contacto(nombre_pdf, contactos)

    if contacto:
        fullname = contacto["fullname"]
        overseer = normalizar(contacto["overseer"])
        grupo = mapa_grupos.get(overseer, "Sin_Grupo")

        status = contacto["status"]
        inactive = contacto["inactive"]

        if inactive == "1":
            tipo = "inactivos"
        elif status == "Regular Pioneer":
            tipo = "precursores"
        else:
            tipo = "activos"

    else:
        fullname = nombre_pdf
        grupo = "Sin_Grupo"
        tipo = "activos"

    nombre_archivo = limpiar_nombre_archivo(fullname)

    if tipo == "precursores":
        carpeta = f"{OUTPUT_DIR}/Activos/Precursores"

    elif tipo == "activos":
        carpeta = f"{OUTPUT_DIR}/Activos/Publicadores/{grupo}"
        os.makedirs(carpeta, exist_ok=True)

    else:
        carpeta = f"{OUTPUT_DIR}/Inactivos"

    ruta = os.path.join(carpeta, f"{nombre_archivo}.pdf")

    nuevo_pdf = fitz.open()
    for p in paginas:
        nuevo_pdf.insert_pdf(p.parent, from_page=p.number, to_page=p.number)

    nuevo_pdf.save(ruta)
    nuevo_pdf.close()

    print(f"Guardado: {ruta}")
# =========================
# PROCESAR POR RANGO
# =========================

def procesar_rango(doc, inicio, fin, contactos, mapa_grupos):
    #  Convertir de base 1 → base 0
    inicio -= 1
    fin -= 1

    persona_actual = None
    paginas_persona = []

    for i in range(inicio, fin + 1):
        if i >= len(doc):
            break

        page = doc[i]
        texto = page.get_text()
        nombre = extraer_nombre_pagina(texto)

        if persona_actual is None:
            persona_actual = nombre

        if nombre != persona_actual:
            guardar_persona(persona_actual, paginas_persona, contactos, mapa_grupos)
            paginas_persona = []
            persona_actual = nombre

        paginas_persona.append(page)

    if paginas_persona:
        guardar_persona(persona_actual, paginas_persona, contactos, mapa_grupos)

# =========================
# PROCESAR TOTALES
# =========================

def detectar_tipo_totales(texto):
    t = texto.lower()

    if "publicadores" in t:
        return "Publicadores"
    elif "precursores auxiliares" in t:
        return "Precursores Auxiliares"
    elif "Precursores regulares y especiales y misioneros en el campo" in t or "precursor especial" in t:
        return "Precursores Regulares"
    elif "misionero" in t:
        return "Misioneros"
    
    return "Otros"

def procesar_totales(pdf_path):
    print("Procesando TOTALES...")

    carpeta_totales = os.path.join(OUTPUT_DIR, "Totales")
    os.makedirs(carpeta_totales, exist_ok=True)

    doc = fitz.open(pdf_path)

    grupos = {}

    for page in doc:
        texto = page.get_text()
        tipo = detectar_tipo_totales(texto)

        if tipo not in grupos:
            grupos[tipo] = []

        grupos[tipo].append(page)

    #  Guardar cada tipo en su PDF
    for tipo, paginas in grupos.items():
        ruta = os.path.join(carpeta_totales, f"{tipo}.pdf")

        nuevo_pdf = fitz.open()
        for p in paginas:
            nuevo_pdf.insert_pdf(p.parent, from_page=p.number, to_page=p.number)

        nuevo_pdf.save(ruta)
        nuevo_pdf.close()

        print(f"Guardado Totales: {ruta}")

# =========================
# MAIN
# =========================

def main():
    crear_estructura()

    print("Cargando contactos...")
    contactos = cargar_contactos(CSV_CONTACTOS)

    print("Cargando grupos...")
    mapa_grupos = cargar_grupos(JSON_GRUPOS)

    print("Abriendo PDF único...")
    doc = fitz.open(PDF_UNICO)

    print("Procesando registros...")
    procesar_rango(doc, 1, len(doc), contactos, mapa_grupos)

    procesar_totales(PDF_TOTALES)
    print("✅ Todo listo")

if __name__ == "__main__":
    main()