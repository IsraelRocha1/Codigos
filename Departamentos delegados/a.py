from PIL import Image, ImageDraw, ImageFont
import json
import os

# ==========================
# CONFIGURACIÓN
# ==========================

JSON_FILE = r"C:\Users\israe\Documents\Codigos\Departamentos delegados\departamentos.json"

ANCHO = 1600
ALTO = 2400

AZUL = "#0B2E75"
GRIS = "#F8F8F8"
BLANCO = "#FFFFFF"

# ==========================
# FUENTES
# ==========================

try:
    fuente_titulo = ImageFont.truetype("arialbd.ttf", 72)
    fuente_subtitulo = ImageFont.truetype("arialbd.ttf", 40)
    fuente_texto = ImageFont.truetype("arial.ttf", 34)
except:
    fuente_titulo = ImageFont.load_default()
    fuente_subtitulo = ImageFont.load_default()
    fuente_texto = ImageFont.load_default()

# ==========================
# FUNCIONES
# ==========================

def caja(draw, x1, y1, x2, y2, titulo, contenido):

    draw.rounded_rectangle(
        [x1, y1, x2, y2],
        radius=25,
        outline=AZUL,
        width=3,
        fill=BLANCO
    )

    draw.text(
        ((x1+x2)//2, y1+25),
        titulo,
        fill=AZUL,
        font=fuente_subtitulo,
        anchor="ma"
    )

    area_texto_y = y1 + 75
    area_texto_h = (y2 - y1) - 75

    bbox = draw.multiline_textbbox(
        (0, 0),
        contenido,
        font=fuente_texto,
        spacing=6
    )

    alto_texto = bbox[3] - bbox[1]

    texto_y = area_texto_y + (area_texto_h - alto_texto)//2

    draw.multiline_text(
        ((x1+x2)//2, texto_y),
        contenido,
        fill="black",
        font=fuente_texto,
        anchor="ma",
        align="center",
        spacing=6
    )

def linea(draw, x1, y1, x2, y2):
    draw.line((x1, y1, x2, y2), fill=AZUL, width=5)

def calcular_altura(texto, ancho_max=600, font=fuente_texto):
    bbox = draw.multiline_textbbox(
        (0, 0),
        texto,
        font=font,
        spacing=6
    )
    alto_texto = bbox[3] - bbox[1]
    return alto_texto + 120  # margen interno

# ==========================
# LEER JSON
# ==========================

with open(JSON_FILE, encoding="utf-8") as f:
    datos = json.load(f)

cocina = datos["Alimentación"]["Cocina"]

# ==========================
# EXTRAER DATOS
# ==========================

asistente_alimentacion = "\n".join(cocina["Asistente de Alimentación"])
encargado = cocina["Encargado de Cocina"][0]
asistente = cocina["Asistente de encargado de Cocina"][0]

areas = [
    (
        "Cocina Caliente",
        cocina["Responsable de Cocina Caliente"][0],
        cocina["Ayudantes de cocina caliente"]
    ),
    (
        "Cocina Fría",
        cocina["Responsable de Cocina Fria"][0],
        cocina["Ayudantes de cocina fria"]
    ),
    (
        "Repostería",
        cocina["Responsable de Reposteria"][0],
        cocina["Ayudantes de Reposteria"]
    ),
    (
        "Limpieza",
        cocina["Responsable de Limpieza"][0],
        cocina["Ayudantes de Limpieza"]
    )
]

backup = cocina["Backup"]

# ==========================
# CREAR LIENZO
# ==========================

img = Image.new("RGB", (ANCHO, ALTO), GRIS)
draw = ImageDraw.Draw(img)

# Marco

draw.rounded_rectangle(
    [20, 20, ANCHO-20, ALTO-20],
    radius=40,
    outline=AZUL,
    width=8
)
# ==========================
# LOGO + TITULO CENTRADO
# ==========================

LOGO = r"C:\Users\israe\Documents\Codigos\Departamentos delegados\logo.png"

titulo_departamento = "COCINA"

if os.path.exists(LOGO):

    logo = Image.open(LOGO).convert("RGBA")

    ancho_logo = 340

    proporcion = ancho_logo / logo.width
    alto_logo = int(logo.height * proporcion)

    logo = logo.resize(
        (ancho_logo, alto_logo),
        Image.LANCZOS
    )

    bbox_titulo = draw.textbbox(
        (0, 0),
        titulo_departamento,
        font=fuente_titulo
    )

    ancho_titulo = bbox_titulo[2] - bbox_titulo[0]

    separacion = 40

    ancho_total = ancho_logo + separacion + ancho_titulo

    x_inicio = (ANCHO - ancho_total) // 2

    x_logo = x_inicio
    y_logo = 40

    img.paste(
        logo,
        (x_logo, y_logo),
        logo
    )

    x_titulo = x_logo + ancho_logo + separacion

    draw.text(
        (
            x_titulo,
            y_logo + alto_logo // 2
        ),
        titulo_departamento,
        fill=AZUL,
        font=fuente_titulo,
        anchor="lm"
    )

else:

    draw.text(
        (ANCHO//2, 140),
        titulo_departamento,
        fill=AZUL,
        font=fuente_titulo,
        anchor="mm"
    )

# Fin real del encabezado
fin_encabezado = y_logo + alto_logo

# Espacio debajo del encabezado
y = fin_encabezado + 50

altura = calcular_altura(asistente_alimentacion)

caja(
    draw,
    450, y,
    1150, y + altura,
    "Asistente de Alimentación",
    asistente_alimentacion
)

y += altura + 40


linea(draw, 800, y-40, 800, y)

caja(
    draw,
    450, y,
    1150, y+140,
    "Encargado de Cocina",
    encargado
)

y += 180

linea(draw, 800, y-40, 800, y)

caja(
    draw,
    450, y,
    1150, y+140,
    "Asistente Encargado",
    asistente
)

# ==========================
# AREAS
# ==========================

y += 220

for nombre, responsable, ayudantes in areas:

    linea(draw, 800, y-80, 800, y)

    # Responsable

    caja(
        draw,
        250, y,
        850, y+170,
        nombre,
        responsable
    )

    # Ayudantes

    texto_ayudantes = "\n".join(ayudantes)

    bbox = draw.multiline_textbbox(
        (0, 0),
        texto_ayudantes,
        font=fuente_texto,
        spacing=8
    )

    alto_texto = bbox[3] - bbox[1]

    alto_caja = max(
        170,
        alto_texto + 130
    )

    draw.rounded_rectangle(
        [980, y, 1400, y + alto_caja],
        radius=25,
        outline=AZUL,
        width=3,
        fill=BLANCO
)

    draw.text(
        (1190, y+20),
        f"Ayudantes ({len(ayudantes)})",
        fill=AZUL,
        font=fuente_subtitulo,
        anchor="ma"
    )

    draw.multiline_text(
        (1190, y+80),
        texto_ayudantes,
        fill="black",
        font=fuente_texto,
        anchor="ma",
        align="center"
    )

    centro_izq = y + 170 // 2
    centro_der = y + alto_caja // 2

    centro = (centro_izq + centro_der) // 2

    linea(draw, 850, centro, 980, centro)

    y += max(220, alto_caja + 40)

# ==========================
# BACKUP
# ==========================

texto_backup = "\n".join(backup)

caja(
    draw,
    450,
    y,
    1150,
    y+180,
    "Backup",
    texto_backup
)

# Altura real del contenido
y += 220

# ==========================
# LOGO INFERIOR DERECHO
# ==========================

LOGO2 = r"C:\Users\israe\Documents\Codigos\Departamentos delegados\logo2.png"

if os.path.exists(LOGO2):

    logo2 = Image.open(LOGO2).convert("RGBA")

    ancho_logo2 = 350

    proporcion = ancho_logo2 / logo2.width
    alto_logo2 = int(logo2.height * proporcion)

    logo2 = logo2.resize(
        (ancho_logo2, alto_logo2),
        Image.LANCZOS
    )

    # Transparencia tipo marca de agua
    logo2.putalpha(100)

    x_logo2 = ANCHO - ancho_logo2 - 60
    y_logo2 = ALTO - alto_logo2 - 60

    img.paste(
        logo2,
        (x_logo2, y_logo2),
        logo2
    )
# ==========================
# GUARDAR
# ==========================

salida = r"C:\Users\israe\Documents\Codigos\Departamentos delegados\Cocina_Estilo.png"

img.save(salida)

print("Imagen creada:")
print(salida)