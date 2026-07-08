"""
Extrae nombre, teléfono y correo desde un archivo HTML y los guarda en un Excel.

Requisitos (instalar una sola vez):
    pip install beautifulsoup4 pandas openpyxl

Uso:
    python scrape_contactos.py
"""

from bs4 import BeautifulSoup
import pandas as pd

# --- Configuración ---
RUTA_HTML = r"C:\Users\irocha\Downloads\3.html"
RUTA_SALIDA = r"C:\Users\irocha\Downloads\contactos3.xlsx"


def main():
    with open(RUTA_HTML, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    # Cada persona está en un <article class="card ..."> con id="ptrn-card-N"
    tarjetas = soup.find_all("article", class_="card")
    print(f"Se encontraron {len(tarjetas)} tarjetas.")

    datos = []
    for tarjeta in tarjetas:
        # Nombre: dentro de <ptrn-card-header> hay un <span class="transliteration__vernacular ...">
        header = tarjeta.find("ptrn-card-header")
        if not header:
            continue  # tarjetas sin encabezado de persona (p.ej. "Información del turno")

        nombre_el = header.find("span", class_="transliteration__vernacular")
        nombre = nombre_el.get_text(strip=True) if nombre_el else ""

        if not nombre:
            continue  # ignorar tarjetas que no son de personas

        # Teléfono: <bdi class="text-phone ..."> en cualquier parte de la tarjeta
        telefono_el = tarjeta.find("bdi", class_="text-phone")
        telefono = telefono_el.get_text(strip=True) if telefono_el else ""

        # Correo: cualquier elemento con class="email" (sin importar el tag)
        email_el = tarjeta.find(class_="email")
        email = email_el.get_text(strip=True) if email_el else ""

        datos.append({"Nombre": nombre, "Telefono": telefono, "Correo": email})

    df = pd.DataFrame(datos, columns=["Nombre", "Telefono", "Correo"])
    df.to_excel(RUTA_SALIDA, index=False)
    print(f"Listo. Excel guardado en: {RUTA_SALIDA} ({len(datos)} contactos)")


if __name__ == "__main__":
    main()