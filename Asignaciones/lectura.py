import pandas as pd
import re
import unicodedata
from rapidfuzz import fuzz, process

# =========================
# CONFIG
# =========================
EXCEL_IN   = "telefonos.xlsx"
SHEET_NAME = 0
COL_DIR    = "C"                  # dirección en columna D
COL_TEL    = "telefono"           # ajusta si tu columna teléfono se llama distinto
CUADRAS_CSV = "cuadras_detalladas.csv"
EXCEL_OUT  = "telefonos_ordenados_por_cuadra.xlsx"

MIN_SCORE_CALLE = 60
MIN_SCORE_CRUCE = 70

# =========================
# HELPERS
# =========================
STOPWORDS = {
    "ecuador","quito","pichincha","barrio","sector","urb","urbanizacion",
    "conjunto","edificio","torre","departamento","depto","casa","lote","mz","manzana",
    "frente","junto","diagonal","esquina","esq","entre","y","sn","s/n","nro","no","numero"
}

def tokens_clean(s: str):
    s = fold(s)
    # deja solo letras/numeros/espacios
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    toks = [t for t in s.split() if t and t not in STOPWORDS]
    return toks

def fallback_calle_from_full_address(dir_raw: str, calles_unicas_f, min_score=55):
    """
    Busca la mejor calle OSM 'dentro' del texto completo de la dirección.
    Estrategia:
      - genera n-gramas (2 a 6 palabras) desde la dirección
      - hace fuzzy match de cada n-grama contra el set de calles
      - se queda con el mejor score
    """
    toks = tokens_clean(dir_raw)
    if len(toks) < 2:
        return ("", 0)

    best = ("", 0)
    max_n = min(6, len(toks))

    for n in range(2, max_n + 1):
        for i in range(0, len(toks) - n + 1):
            phrase = " ".join(toks[i:i+n])
            m = process.extractOne(phrase, calles_unicas_f, scorer=fuzz.token_set_ratio)
            if m:
                choice, score, _ = m
                if score > best[1]:
                    best = (choice, score)

    return best if best[1] >= min_score else ("", best[1])

def fold(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = " ".join(s.split())
    s = s.replace("av.", "avenida").replace("av ", "avenida ")
    return s

def best_match(query, choices, min_score=0):
    if not query:
        return ("", 0)
    m = process.extractOne(query, choices, scorer=fuzz.token_set_ratio)
    if not m:
        return ("", 0)
    choice, score, _ = m
    return (choice, score) if score >= min_score else ("", score)

def parse_direccion(dir_raw: str):
    """
    Intenta extraer:
      - calle_principal
      - cruce1
      - cruce2 (opcional)
    Acepta formatos tipo:
      "Quichuas y Sigchos"
      "Quichuas esq Sigchos"
      "Quichuas esquina Sigchos"
      "Quichuas entre Sigchos y Paquisha"
      "Av. Mariscal Sucre y Ambuqui"
    """
    s = fold(dir_raw)

    # limpiar cosas típicas que estorban
    s = re.sub(r"\b(nro|no|num|numero|#)\b.*$", "", s).strip()  # quita "Nro 123..."
    s = re.sub(r"[;,]", " ", s)
    s = " ".join(s.split())

    # entre A y B
    m = re.search(r"^(.*?)\s+entre\s+(.*?)\s+y\s+(.*)$", s)
    if m:
        calle = m.group(1).strip()
        c1 = m.group(2).strip()
        c2 = m.group(3).strip()
        return calle, c1, c2

    # esq / esquina
    m = re.search(r"^(.*?)\s+(esq|esquina)\s+(.*)$", s)
    if m:
        calle = m.group(1).strip()
        c1 = m.group(3).strip()
        return calle, c1, ""

    # A y B
    m = re.search(r"^(.*?)\s+y\s+(.*)$", s)
    if m:
        calle = m.group(1).strip()
        c1 = m.group(2).strip()
        return calle, c1, ""

    # si no pudo, devuelve todo como calle (sin cruce)
    return s, "", ""

# =========================
# 1) Cargar datos
# =========================
df_x = pd.read_excel(EXCEL_IN, sheet_name=SHEET_NAME)

# leer columna D por posición (4ta columna)
col_d_name = df_x.columns[2]   # D = índice 3
df_x["direccion_raw"] = df_x[col_d_name].astype(str)

# =========================
# 2) Parsear dirección
# =========================
parsed = df_x["direccion_raw"].apply(parse_direccion)
df_x["calle_in"] = parsed.apply(lambda t: t[0])
df_x["cruce1_in"] = parsed.apply(lambda t: t[1])
df_x["cruce2_in"] = parsed.apply(lambda t: t[2])

# =========================
# 3) Cargar cuadras y preparar
# =========================
df_c = pd.read_csv(CUADRAS_CSV, encoding="utf-8-sig")
df_c["calle_f"] = df_c["calle"].map(fold)
df_c["desde_f"] = df_c["desde_cruce"].map(fold)
df_c["hasta_f"] = df_c["hasta_cruce"].map(fold)

calles_unicas_f = sorted(df_c["calle_f"].unique())

# =========================
# 4) Asignación de cuadra
# =========================
def asignar(row):
    calle_in_f = fold(row["calle_in"])
    cr1_f = fold(row["cruce1_in"])
    cr2_f = fold(row["cruce2_in"])

    # 1) intento normal (por calle_in)
    calle_match_f, score_calle = best_match(calle_in_f, calles_unicas_f, MIN_SCORE_CALLE)

    # 2) si no calza, fallback usando direccion completa
    origen = "CALLE_IN"
    if not calle_match_f:
        calle_match_f, score_fb = fallback_calle_from_full_address(row["direccion_raw"], calles_unicas_f, min_score=55)
        if calle_match_f:
            score_calle = score_fb
            origen = "FALLBACK_DIR"

    if not calle_match_f:
        return pd.Series({
            "calle_osm": None, "desde_osm": None, "hasta_osm": None,
            "score_calle": score_calle, "score_cruce": 0,
            "estado": "NO_MATCH_CALLE",
            "origen_calle": origen
        })

    sub = df_c[df_c["calle_f"] == calle_match_f].copy()

    # (igual que tu lógica de cruces...)
    if cr1_f and cr2_f:
        def score_tramo(r):
            s1 = fuzz.token_set_ratio(cr1_f, r["desde_f"]) + fuzz.token_set_ratio(cr2_f, r["hasta_f"])
            s2 = fuzz.token_set_ratio(cr1_f, r["hasta_f"]) + fuzz.token_set_ratio(cr2_f, r["desde_f"])
            return max(s1, s2) / 2.0
        sub["score_cruce"] = sub.apply(score_tramo, axis=1)

    elif cr1_f:
        def score_tramo_1(r):
            return max(fuzz.token_set_ratio(cr1_f, r["desde_f"]), fuzz.token_set_ratio(cr1_f, r["hasta_f"]))
        sub["score_cruce"] = sub.apply(score_tramo_1, axis=1)

    else:
        r0 = sub.iloc[0]
        return pd.Series({
            "calle_osm": r0["calle"], "desde_osm": None, "hasta_osm": None,
            "score_calle": score_calle, "score_cruce": 0,
            "estado": "MATCH_SOLO_CALLE",
            "origen_calle": origen
        })

    best = sub.sort_values("score_cruce", ascending=False).head(1).iloc[0]
    estado = "OK" if float(best["score_cruce"]) >= MIN_SCORE_CRUCE else "MATCH_CALLE_CRUCE_DUDOSO"

    return pd.Series({
        "calle_osm": best["calle"],
        "desde_osm": best["desde_cruce"],
        "hasta_osm": best["hasta_cruce"],
        "score_calle": score_calle,
        "score_cruce": float(best["score_cruce"]),
        "estado": estado,
        "origen_calle": origen
    })

asig = df_x.apply(asignar, axis=1)
df_out = pd.concat([df_x, asig], axis=1)

# ID de cuadra para ordenar / agrupar
df_out["cuadra_id"] = df_out.apply(
    lambda r: None if pd.isna(r["calle_osm"]) else (
        f"{r['calle_osm']} | {r['desde_osm']} -> {r['hasta_osm']}"
        if pd.notna(r["desde_osm"]) else f"{r['calle_osm']}"
    ),
    axis=1
)

# Ordenar
tel_col = COL_TEL if COL_TEL in df_out.columns else None
sort_cols = ["calle_osm", "cuadra_id"] + ([tel_col] if tel_col else [])
df_out = df_out.sort_values(sort_cols, na_position="last")

df_out.to_excel(EXCEL_OUT, index=False)
print("LISTO ->", EXCEL_OUT)
print(df_out["estado"].value_counts(dropna=False))