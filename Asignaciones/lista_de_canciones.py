from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import re
import time
import os

url = "https://wol.jw.org/es/wol/publication/r4/lp-s/sjj/329"

# ===== RUTAS =====
base_proyecto = os.path.dirname(os.path.abspath(__file__))
carpeta_recursos = os.path.join(base_proyecto, "Recursos")
ruta_salida = os.path.join(carpeta_recursos, "canticos_lista.xlsx")

# ===== SELENIUM (sin chromedriver descargado) =====
chrome_options = Options()
# chrome_options.add_argument("--headless=new")  # opcional

driver = webdriver.Chrome(options=chrome_options)

try:
    driver.get(url)

    WebDriverWait(driver, 15).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, "cardTitleBlock"))
    )

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)

    elementos = driver.find_elements(By.CLASS_NAME, "cardTitleBlock")
    bloques = [e.text.strip() for e in elementos if e.text.strip()]

finally:
    driver.quit()

# ===== PROCESAR BLOQUES =====
numeros, titulos = [], []

for texto in bloques:
    lineas = [l.strip() for l in str(texto).split("\n") if l.strip()]
    numero, titulo = "", ""

    for i, linea in enumerate(lineas):
        if re.match(r"^CANCIÓN\s+\d+", linea.upper()):
            m = re.search(r"\d+", linea)
            if m:
                numero = f"Canción {m.group()}"
            if i + 1 < len(lineas):
                titulo = " ".join(lineas[i + 1:])
            break

    numeros.append(numero)
    titulos.append(titulo)

df = pd.DataFrame({"Canción #": numeros, "Título": titulos})

# Guarda en Recursos
df.to_excel(ruta_salida, index=False)

print(f"✅ Archivo creado en: {ruta_salida}")