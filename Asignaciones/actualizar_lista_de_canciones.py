import re
import os
import pandas as pd
import requests
from bs4 import BeautifulSoup

url = "https://wol.jw.org/es/wol/publication/r4/lp-s/sjj/164"

# ===== RUTAS =====
base_proyecto = os.path.dirname(os.path.abspath(__file__))
carpeta_recursos = os.path.join(base_proyecto, "Recursos")
ruta_salida = os.path.join(carpeta_recursos, "canticos_lista.xlsx")

# ===== DESCARGAR LA PÁGINA (sin navegador) =====
# La página trae el número y el título de cada canción directamente en el
# HTML, así que no hace falta Selenium/Chrome para esto — evita el mismo
# bloqueo de política de aplicaciones que dio problemas antes.
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
resp = requests.get(url, headers=headers, timeout=20)
resp.raise_for_status()
soup = BeautifulSoup(resp.text, "html.parser")

# ===== EXTRAER "CANCIÓN N Título" DE CADA ENLACE =====
numeros, titulos = [], []
vistos = set()
for enlace in soup.find_all("a"):
    texto = " ".join(enlace.get_text().split())
    m = re.match(r"^CANCI[ÓO]N\s+(\d+)\s+(.+)$", texto, re.IGNORECASE)
    if not m:
        continue
    numero = f"Canción {m.group(1)}"
    if numero in vistos:
        continue
    vistos.add(numero)
    numeros.append(numero)
    titulos.append(m.group(2).strip())

if not numeros:
    raise RuntimeError(
        "No encontré ninguna canción en la página. Puede que wol.jw.org haya "
        "cambiado de formato — revisa manualmente la URL."
    )

df = pd.DataFrame({"Canción #": numeros, "Título": titulos})
df.to_excel(ruta_salida, index=False)

print(f"✅ Archivo creado en: {ruta_salida} ({len(df)} canciones)")