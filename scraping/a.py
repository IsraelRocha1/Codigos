"""
Extrae nombre, teléfono, correo y dato adicional desde varios archivos HTML
y los guarda en un solo Excel, cada uno en su propia pestaña.

Requisitos (instalar una sola vez):
    pip install beautifulsoup4 pandas openpyxl

Uso:
    python scrape_contactos.py
"""

from bs4 import BeautifulSoup
import pandas as pd

# --- Configuración ---
# Clave = ruta del archivo HTML, Valor = nombre de la pestaña en el Excel
ARCHIVOS = {
    r"C:\Users\israe\Downloads\1mañana.html": "Viernes - Mañana",
    r"C:\Users\israe\Downloads\1tarde.html":  "Viernes - Tarde",
    r"C:\Users\israe\Downloads\2mañana.html": "Sabado - Mañana",
    r"C:\Users\israe\Downloads\2tarde.html":  "Sabado - Tarde",
    r"C:\Users\israe\Downloads\3mañana.html": "Domingo - Mañana",
    r"C:\Users\israe\Downloads\3tarde.html":  "Domingo - Tarde",
}

RUTA_SALIDA = r"C:\Users\israe\Downloads\contactos.xlsx"


def extraer_datos(ruta_html):
    with open(ruta_html, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    tarjetas = soup.find_all("article", class_="card")
    print(f"  -> {len(tarjetas)} tarjetas encontradas en {ruta_html}")

    datos = []
    for tarjeta in tarjetas:
        header = tarjeta.find("ptrn-card-header")
        if not header:
            continue

        nombre_el = header.find("span", class_="transliteration__vernacular")
        nombre = nombre_el.get_text(strip=True) if nombre_el else ""

        if not nombre:
            continue

        telefono_el = tarjeta.find("bdi", class_="text-phone")
        telefono = telefono_el.get_text(strip=True) if telefono_el else ""

        email_el = tarjeta.find(class_="email")
        email = email_el.get_text(strip=True) if email_el else ""

        todos_els = tarjeta.find_all("span", class_="transliteration__vernacular")
        extras = []
        for el in todos_els:
            if el is nombre_el:
                continue
            texto = el.get_text(strip=True)
            if texto:
                extras.append(texto)
        congregacion = "; ".join(extras)

        datos.append({
            "Nombre": nombre,
            "Telefono": telefono,
            "Correo": email,
            "Congregacion": congregacion,
        })

    return datos


def main():
    with pd.ExcelWriter(RUTA_SALIDA, engine="openpyxl") as writer:
        for ruta_html, nombre_hoja in ARCHIVOS.items():
            print(f"Procesando: {ruta_html}")
            try:
                datos = extraer_datos(ruta_html)
            except FileNotFoundError:
                print(f"  !! No se encontró el archivo, se omite: {ruta_html}")
                continue

            df = pd.DataFrame(datos, columns=["Nombre", "Telefono", "Correo", "Congregacion"])
            # Excel no permite nombres de hoja > 31 caracteres
            hoja = nombre_hoja[:31]
            df.to_excel(writer, sheet_name=hoja, index=False)
            print(f"  -> {len(datos)} contactos guardados en pestaña '{hoja}'")

    print(f"\nListo. Excel guardado en: {RUTA_SALIDA}")


if __name__ == "__main__":
    main()