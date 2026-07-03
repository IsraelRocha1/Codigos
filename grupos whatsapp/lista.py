"""
Crea un grupo de WhatsApp y agrega contactos automáticamente vía WhatsApp Web.

REQUISITOS:
- pip install selenium webdriver-manager
- Chrome instalado
- Tener WhatsApp Web vinculado (vas a escanear el QR la primera vez)

IMPORTANTE:
- Esto usa WhatsApp Web de forma automatizada, lo cual NO está oficialmente
  soportado por WhatsApp y puede infringir sus Términos de Servicio.
  Úsalo bajo tu propio riesgo (posible bloqueo temporal o permanente del número).
- Solo puedes agregar directo al grupo a contactos que te tengan guardado
  (o según su config de privacidad); al resto WhatsApp les enviará invitación.
- Usa pausas largas entre cada persona para reducir el riesgo de bloqueo.
"""

import time
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

# --------- CONFIGURACIÓN ---------
NOMBRE_GRUPO = "Nombre del grupo"
ARCHIVO_CONTACTOS = r"C:\Users\israe\Documents\Codigos\grupos whatsapp\hourglass-contactlist.csv"
COLUMNA_NOMBRE = "fullname"         # nombre de la columna que contiene el nombre del contacto
PAUSA_ENTRE_CONTACTOS = 15            # segundos, sube esto si puedes (30+ es más seguro)
# ----------------------------------


def iniciar_driver():
    import os
    perfil_path = os.path.join(os.getcwd(), "whatsapp_session")
    os.makedirs(perfil_path, exist_ok=True)

    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={perfil_path}")  # guarda sesión, no vuelves a escanear QR
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--remote-debugging-port=9222")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://web.whatsapp.com")
    print("Escanea el QR si es la primera vez... esperando a que cargue WhatsApp Web")
    WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Lista de chats"]'))
    )
    print("WhatsApp Web listo.")
    return driver


def cerrar_modales(driver):
    """Cierra cualquier popup/modal que tape la interfaz (avisos, tips, etc.)"""
    posibles_cierres = [
        '//div[@aria-label="Cerrar" or @aria-label="Close"]',
        '//button[@aria-label="Cerrar" or @aria-label="Close"]',
        '//div[@data-testid="x-viewer-close"]',
    ]
    for xpath in posibles_cierres:
        try:
            elem = driver.find_element(By.XPATH, xpath)
            if elem.is_displayed():
                elem.click()
                time.sleep(1)
        except Exception:
            pass
    # Como último recurso, ESC cierra la mayoría de los overlays de WhatsApp Web
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        time.sleep(0.5)
    except Exception:
        pass


def click_seguro(driver, elemento):
    """Intenta click con mouse real (ActionChains); si falla, click normal; si falla, click vía JS."""
    try:
        ActionChains(driver).move_to_element(elemento).pause(0.3).click(elemento).perform()
        return
    except Exception:
        pass
    try:
        elemento.click()
        return
    except Exception:
        pass
    driver.execute_script("arguments[0].click();", elemento)


def buscar_elemento_robusto(driver, xpaths, timeout=20, descripcion="elemento"):
    """Intenta varios XPaths distintos (la UI de WhatsApp Web cambia seguido)."""
    fin = time.time() + timeout
    ultimo_error = None
    while time.time() < fin:
        for xpath in xpaths:
            try:
                elem = driver.find_element(By.XPATH, xpath)
                if elem.is_displayed():
                    return elem
            except Exception as e:
                ultimo_error = e
        time.sleep(0.5)
    # Si no se encontró nada, guarda captura para depurar
    try:
        ruta_captura = f"error_{descripcion.replace(' ', '_')}.png"
        driver.save_screenshot(ruta_captura)
        print(f"⚠ No se encontró '{descripcion}'. Captura guardada en: {ruta_captura}")
    except Exception:
        pass
    raise TimeoutException(f"No se encontró '{descripcion}' con ninguno de los selectores probados.") from ultimo_error


def crear_grupo(driver, nombre_grupo):
    cerrar_modales(driver)

    xpaths_nuevo_chat = [
        '//div[@title="Nuevo chat" or @title="New chat"]',
        '//div[@aria-label="Nuevo chat" or @aria-label="New chat"]',
        '//button[@aria-label="Nuevo chat" or @aria-label="New chat"]',
        '//span[@data-icon="new-chat-outline"]',
        '//span[@data-icon="new-chat"]',
        '//div[@data-icon="new-chat-outline"]',
    ]
    nuevo_chat = buscar_elemento_robusto(driver, xpaths_nuevo_chat, descripcion="boton nuevo chat")
    click_seguro(driver, nuevo_chat)
    time.sleep(1)

    xpaths_nuevo_grupo = [
        '//div[contains(text(),"Nuevo grupo") or contains(text(),"New group")]',
        '//span[contains(text(),"Nuevo grupo") or contains(text(),"New group")]',
        '//li[contains(.,"Nuevo grupo") or contains(.,"New group")]',
    ]

    nuevo_grupo = None
    for intento in range(4):
        try:
            nuevo_grupo = buscar_elemento_robusto(
                driver, xpaths_nuevo_grupo, timeout=4, descripcion="opcion nuevo grupo"
            )
            break
        except TimeoutException:
            print(f"  Menú no apareció, reintentando clic en 'Nuevo chat' ({intento + 1}/4)...")
            nuevo_chat = buscar_elemento_robusto(driver, xpaths_nuevo_chat, descripcion="boton nuevo chat")
            click_seguro(driver, nuevo_chat)
            time.sleep(1)

    if nuevo_grupo is None:
        raise TimeoutException("No se pudo abrir el menú de 'Nuevo grupo' tras varios intentos.")

    click_seguro(driver, nuevo_grupo)
    time.sleep(1)
    print("Grupo iniciado. Ahora agregando contactos...")


def agregar_contacto(driver, nombre_contacto):
    try:
        xpaths_buscador = [
            '//div[@contenteditable="true"][@data-tab]',
            '//div[@contenteditable="true"][@role="textbox"]',
            '//p[@contenteditable="true"]',
            '//div[@title="Buscar nombre o número" or @title="Search name or number"]',
            '//input[@type="text"]',
        ]
        buscador = buscar_elemento_robusto(driver, xpaths_buscador, timeout=10, descripcion="buscador de contactos")

        # Limpiar ANTES de escribir, para no arrastrar texto de la búsqueda anterior
        buscador.click()
        buscador.send_keys(Keys.CONTROL + "a")
        buscador.send_keys(Keys.DELETE)
        time.sleep(0.3)

        buscador.send_keys(nombre_contacto)
        time.sleep(2)

        xpaths_resultado = [
            f'//span[@title="{nombre_contacto}"]',
        ]
        resultado = buscar_elemento_robusto(
            driver, xpaths_resultado, timeout=8, descripcion=f"resultado busqueda {nombre_contacto}"
        )
        click_seguro(driver, resultado)

        # Limpiar después de seleccionar, para dejar listo el campo para el siguiente contacto
        buscador.send_keys(Keys.CONTROL + "a")
        buscador.send_keys(Keys.DELETE)
        print(f"  ✓ Agregado: {nombre_contacto}")
        return True
    except Exception as e:
        print(f"  ✗ No se pudo agregar a {nombre_contacto}: {type(e).__name__}: {e}")
        # Intentar limpiar el buscador igual, para no arrastrar el error al siguiente contacto
        try:
            buscador.send_keys(Keys.CONTROL + "a")
            buscador.send_keys(Keys.DELETE)
        except Exception:
            pass
        return False


def finalizar_creacion_grupo(driver, nombre_grupo):
    # Flecha para continuar tras seleccionar contactos
    xpaths_siguiente = [
        '//div[@aria-label="Siguiente" or @aria-label="Next"]',
        '//span[@data-icon="arrow-forward" or @data-icon="forward-icon"]',
        '//button[@aria-label="Siguiente" or @aria-label="Next"]',
    ]
    siguiente = buscar_elemento_robusto(driver, xpaths_siguiente, descripcion="boton siguiente")
    click_seguro(driver, siguiente)
    time.sleep(2)

    xpaths_nombre_grupo = [
        '//div[@contenteditable="true"][@data-tab]',
        '//div[@contenteditable="true"][@role="textbox"]',
        '//p[@contenteditable="true"]',
        '//div[@title="Escribe el nombre del grupo" or @title="Type group subject"]',
    ]
    nombre_input = buscar_elemento_robusto(driver, xpaths_nombre_grupo, descripcion="campo nombre del grupo")
    nombre_input.click()
    nombre_input.send_keys(nombre_grupo)
    time.sleep(1)

    xpaths_crear = [
        '//div[@aria-label="Crear grupo" or @aria-label="Create group"]',
        '//span[@data-icon="checkmark" or @data-icon="checkmark-medium"]',
        '//button[@aria-label="Crear grupo" or @aria-label="Create group"]',
    ]
    crear = buscar_elemento_robusto(driver, xpaths_crear, descripcion="boton crear grupo")
    click_seguro(driver, crear)
    print(f"Grupo '{nombre_grupo}' creado.")


def leer_contactos(archivo, columna=COLUMNA_NOMBRE):
    with open(archivo, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if columna not in reader.fieldnames:
            raise ValueError(
                f"La columna '{columna}' no existe en el CSV. "
                f"Columnas disponibles: {reader.fieldnames}"
            )
        return [row[columna].strip() for row in reader if row.get(columna, "").strip()]


def main():
    contactos = leer_contactos(ARCHIVO_CONTACTOS)
    print(f"{len(contactos)} contactos cargados desde {ARCHIVO_CONTACTOS}")

    driver = iniciar_driver()
    crear_grupo(driver, NOMBRE_GRUPO)

    agregados = 0
    for i, contacto in enumerate(contactos, start=1):
        print(f"[{i}/{len(contactos)}] Agregando {contacto}...")
        if agregar_contacto(driver, contacto):
            agregados += 1
        time.sleep(PAUSA_ENTRE_CONTACTOS)

    finalizar_creacion_grupo(driver, NOMBRE_GRUPO)
    print(f"\nProceso terminado. {agregados}/{len(contactos)} contactos agregados.")
    print("Nota: los que no se pudieron agregar directo recibirán invitación por link.")

    input("Presiona Enter para cerrar el navegador...")
    driver.quit()


if __name__ == "__main__":
    main()