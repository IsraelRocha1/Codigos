import osmnx as ox
import geopandas as gpd
from shapely.geometry import Polygon, Point
from shapely.ops import unary_union, nearest_points
import pandas as pd
import unicodedata

# ============================================================
# CONFIG
# ============================================================
CITY = "Quito, Ecuador"
NETWORK_TYPE = "drive"

# Punto dentro del sector (lat, lon) para elegir la intersección correcta si hay varias
CENTER_LATLON = (-0.2550, -78.5350)  # <-- CAMBIA ESTO (punto dentro del sector)

# Lista de vértices del borde en ORDEN (aquí metes la "muesca")
# Cada vértice es una intersección (CalleA, CalleB)
VERTICES = [
    ("Quichuas", "Sigchos"),
    ("Sigchos", "Paquisha"),
    ("Paquisha", "S14"),
    ("S14", "El Pangui"),
    ("El Pangui", "Lucas de Queva"),
    ("Lucas de Queva", "Avenida Cardenal Carlos de la Torre"),
    ("Avenida Cardenal Carlos de la Torre", "Ambuqui"),
    ("Ambuqui", "Avenida Mariscal Sucre"),
    ("Avenida Mariscal Sucre", "Quichuas"),
    # agrega más pares si tu borde tiene más quiebres (muesca)
]

# Si un vértice NO cruza y el punto más cercano queda muy lejos,
# es señal de que esa pareja de calles está mal.
MAX_FALLBACK_DIST_M = 80.0
# ============================================================


def normalize_name(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    if isinstance(x, list):
        return str(x[0]).strip() if x else None
    return str(x).strip()

def fold(s: str) -> str:
    s = s.strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = " ".join(s.split())
    return s

def street_geom_from_edges(edges_gdf, street_name):
    target = fold(street_name)
    names_raw = edges_gdf["name"].apply(normalize_name)
    names_fold = names_raw.fillna("").map(fold)

    sel = edges_gdf.loc[names_fold == target]
    if sel.empty:
        sel = edges_gdf.loc[names_fold.str.contains(target, na=False)]

    if sel.empty:
        raise ValueError(f"No encontré la calle '{street_name}' en OSM dentro del área descargada.")

    return unary_union(sel.geometry.values)

def intersection_point_near_center(edges_gdf, street_a, street_b, center_point_proj):
    ga = street_geom_from_edges(edges_gdf, street_a)
    gb = street_geom_from_edges(edges_gdf, street_b)

    inter = ga.intersection(gb)

    if inter.is_empty:
        p1, p2 = nearest_points(ga, gb)
        return p1, float(p1.distance(p2)), True

    if inter.geom_type == "Point":
        return inter, 0.0, False

    if inter.geom_type == "MultiPoint":
        pts = list(inter.geoms)
        pts_sorted = sorted(pts, key=lambda p: p.distance(center_point_proj))
        return pts_sorted[0], 0.0, False

    rp = inter.representative_point()
    return rp, 0.0, False

def build_node_to_streets(edges_gdf):
    d = {}
    tmp = edges_gdf.copy()
    tmp["street_name"] = tmp["name"].apply(normalize_name)

    for _, r in tmp.iterrows():
        s = r["street_name"]
        if not s:
            continue
        u = r["u"]; v = r["v"]
        d.setdefault(u, set()).add(s)
        d.setdefault(v, set()).add(s)
    return d

def cross_label(node_to_streets, node_id, current_street):
    st = node_to_streets.get(node_id, set())
    crosses = sorted([x for x in st if x != current_street])
    return " / ".join(crosses) if crosses else "SIN_CRUCE_NOMBRADO"


# ============================================================
# 1) Descargar red vial "grande" para poder encontrar calles
# ============================================================
print("Descargando red vial de la ciudad... (puede tardar)")
ox.settings.use_cache = True
ox.settings.log_console = False

G = ox.graph_from_place(CITY, network_type=NETWORK_TYPE, simplify=True)
nodes, edges = ox.graph_to_gdfs(G, nodes=True, edges=True)

edges_proj = ox.projection.project_gdf(edges)

center_ll = gpd.GeoSeries([Point(CENTER_LATLON[1], CENTER_LATLON[0])], crs="EPSG:4326")
center_proj = center_ll.to_crs(edges_proj.crs).iloc[0]

# ============================================================
# 2) Calcular vertices del polígono (muesca incluida)
# ============================================================
print("Calculando vertices del sector...")

corner_points = []
for i, (a, b) in enumerate(VERTICES, start=1):
    p, dist_m, fallback = intersection_point_near_center(edges_proj, a, b, center_proj)
    if fallback and dist_m > MAX_FALLBACK_DIST_M:
        raise ValueError(
            f"VERTICE {i}: '{a}' y '{b}' no se cruzan (fallback {dist_m:.1f} m). "
            f"Esa pareja no es una esquina real del borde. Corrígela."
        )
    if fallback:
        print(f"AVISO VERTICE {i}: '{a}' & '{b}' no cruzan en OSM, usando punto más cercano (dist {dist_m:.1f} m)")
    corner_points.append(p)

poly_proj = Polygon([(p.x, p.y) for p in corner_points] + [(corner_points[0].x, corner_points[0].y)])

poly_gdf = gpd.GeoDataFrame(geometry=[poly_proj], crs=edges_proj.crs)
poly_latlon = poly_gdf.to_crs("EPSG:4326").geometry.iloc[0]

print("Vertices calculados (lat, lon):")
for p in gpd.GeoSeries(corner_points, crs=edges_proj.crs).to_crs("EPSG:4326"):
    print((round(p.y, 6), round(p.x, 6)))

# ============================================================
# 3) Descargar red dentro del polígono (sector final)
# ============================================================
print("Descargando red dentro del sector...")
G2 = ox.graph_from_polygon(poly_latlon, network_type=NETWORK_TYPE, simplify=True)
nodes2, edges2 = ox.graph_to_gdfs(G2, nodes=True, edges=True)

# FIX u/v en columnas
if "u" not in edges2.columns or "v" not in edges2.columns:
    edges2 = edges2.reset_index()

edges2.to_file("sector_tramos.geojson", driver="GeoJSON")
print("LISTO -> sector_tramos.geojson")

# ============================================================
# 4) Generar cuadras (tramos entre intersecciones)
# ============================================================
print("Generando cuadras (tramo entre intersecciones)...")

edges2 = edges2.copy()
edges2["street_name"] = edges2["name"].apply(normalize_name)
edges2 = edges2[edges2["street_name"].notna()]

node_to_streets = build_node_to_streets(edges2)

rows = []
for _, r in edges2.iterrows():
    street = r["street_name"]
    u = r["u"]; v = r["v"]
    length_m = float(r.get("length", 0.0))

    rows.append({
        "calle": street,
        "desde_cruce": cross_label(node_to_streets, u, street),
        "hasta_cruce": cross_label(node_to_streets, v, street),
        "longitud_m": round(length_m, 2),
        "u": u, "v": v
    })

df = pd.DataFrame(rows)

# quitar duplicados ida/vuelta
df["uv_key"] = df.apply(lambda r: f"{r['calle']}|{min(r['u'], r['v'])}|{max(r['u'], r['v'])}", axis=1)
df = df.drop_duplicates("uv_key").drop(columns=["uv_key"])

df = df.sort_values(["calle", "desde_cruce", "hasta_cruce", "longitud_m"],
                    ascending=[True, True, True, False])

df.to_csv("cuadras_detalladas.csv", index=False, encoding="utf-8-sig")
print("LISTO -> cuadras_detalladas.csv")
print(df.head(25).to_string(index=False))