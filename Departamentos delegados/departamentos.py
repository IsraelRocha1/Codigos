import pandas as pd
import json
import unicodedata
from collections import defaultdict

# ---------------------------
# Función de normalización
# ---------------------------
def normalizar(texto):
    if pd.isna(texto):
        return ""

    texto = str(texto).strip().lower()

    if texto == "":
        return ""

    # Proteger la ñ y Ñ temporalmente
    texto = texto.replace("ñ", "__enie__")
    texto = texto.replace("Ñ", "__ENIE__")

    # Eliminar tildes
    texto = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )

    # Restaurar la ñ
    texto = texto.replace("__enie__", "ñ")
    texto = texto.replace("__ENIE__", "Ñ")

    return texto


# ---------------------------
# Extracción recursiva de nombres
# ---------------------------
def extraer_nombres(nodo, departamento, nombres_por_departamento):
    """
    Recorre recursivamente el JSON (sin importar cuántos niveles de
    subdepartamentos/roles tenga) y agrega cada nombre encontrado
    a la lista del departamento de nivel superior correspondiente.
    """
    if isinstance(nodo, list):
        for nombre in nodo:
            # Ignorar valores vacíos, None, o strings en blanco
            if nombre is None:
                continue
            if isinstance(nombre, str) and nombre.strip() == "":
                continue
            nombres_por_departamento[departamento].append(nombre)

    elif isinstance(nodo, dict):
        for valor in nodo.values():
            extraer_nombres(valor, departamento, nombres_por_departamento)

    # Si no es ni lista ni dict (caso inesperado), se ignora


# ---------------------------
# Leer CSV
# ---------------------------
df = pd.read_csv(
    r"C:\Users\israe\Documents\Codigos\Archivos descargados desde hourglass\hourglass-contactlist.csv",
    dtype=str
)

# ---------------------------
# Leer JSON
# ---------------------------
with open(
    r"C:\Users\israe\Documents\Codigos\Departamentos delegados\departamentos.json",
    "r",
    encoding="utf-8"
) as f:
    departamentos = json.load(f)

# ---------------------------
# Extraer todos los nombres, agrupados por departamento de nivel superior
# ---------------------------
nombres_por_departamento = defaultdict(list)

for departamento, contenido in departamentos.items():
    extraer_nombres(contenido, departamento, nombres_por_departamento)

# =====================================================
# PARTE 1: PERSONAS NO ASIGNADAS
# =====================================================

nombres_departamentos = set()

for lista_nombres in nombres_por_departamento.values():
    for nombre in lista_nombres:
        nombres_departamentos.add(normalizar(nombre))

# Quitar cadena vacía por seguridad (no debería quedar ninguna, pero por si acaso)
nombres_departamentos.discard("")

df["fullname_normalizado"] = df["fullname"].apply(normalizar)

df_no_asignados = df[
    ~df["fullname_normalizado"].isin(nombres_departamentos)
].copy()

df_no_asignados.drop(
    columns=["fullname_normalizado"],
    inplace=True
)

archivo_no_asignados = "personas_no_asignadas.xlsx"

df_no_asignados.to_excel(
    archivo_no_asignados,
    index=False
)

# =====================================================
# PARTE 2: PERSONAS EN VARIOS DEPARTAMENTOS
# =====================================================

personas_departamentos = defaultdict(list)

for departamento, lista_nombres in nombres_por_departamento.items():
    for nombre in lista_nombres:

        nombre_norm = normalizar(nombre)
        if nombre_norm == "":
            continue

        personas_departamentos[nombre_norm].append({
            "Nombre": nombre,
            "Departamento": departamento
        })

filas_repetidos = []

for nombre_norm, datos in personas_departamentos.items():

    departamentos_unicos = sorted(
        set(d["Departamento"] for d in datos)
    )

    if len(departamentos_unicos) > 1:

        filas_repetidos.append({
            "Nombre": datos[0]["Nombre"],
            "Cantidad_Departamentos": len(departamentos_unicos),
            "Departamentos": ", ".join(departamentos_unicos)
        })

df_repetidos = pd.DataFrame(filas_repetidos)

if not df_repetidos.empty:
    df_repetidos = df_repetidos.sort_values(
        by="Cantidad_Departamentos",
        ascending=False
    )

archivo_repetidos = "personas_en_varios_departamentos.xlsx"

df_repetidos.to_excel(
    archivo_repetidos,
    index=False
)

# =====================================================
# RESUMEN
# =====================================================

print("\n===================================")
print("RESUMEN")
print("===================================")

print(f"\nPersonas no asignadas: {len(df_no_asignados)}")
print(f"Archivo generado: {archivo_no_asignados}")

print(f"\nPersonas en varios departamentos: {len(df_repetidos)}")
print(f"Archivo generado: {archivo_repetidos}")

if not df_repetidos.empty:

    print("\nPERSONAS REPETIDAS:\n")

    for _, fila in df_repetidos.iterrows():
        print(
            f"{fila['Nombre']} "
            f"({fila['Cantidad_Departamentos']} departamentos) -> "
            f"{fila['Departamentos']}"
        )