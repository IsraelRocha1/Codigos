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
# Leer CSV
# ---------------------------
df = pd.read_csv(
    r"C:\Users\israe\Documents\Codigos\Departamentos delegados\hourglass-contactlist.csv",
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

# =====================================================
# PARTE 1: PERSONAS NO ASIGNADAS
# =====================================================

nombres_departamentos = set()

for lista in departamentos.values():
    for nombre in lista:
        nombres_departamentos.add(normalizar(nombre))

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

for departamento, lista_personas in departamentos.items():
    for nombre in lista_personas:

        nombre_norm = normalizar(nombre)

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