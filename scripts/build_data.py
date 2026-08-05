"""
build_data.py
Genera src/precios.json a partir de:
  - sources/Precios_de_Alquiler.xls   → precios por distrito/municipio
  - sources/municipios_ine.csv         → listado oficial INE de municipios de Madrid

Uso:
  pip install xlrd openpyxl pandas
  python scripts/build_data.py
"""

import csv, json, re, unicodedata, sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Instala dependencias: pip install xlrd pandas")

ROOT = Path(__file__).parent.parent
XLS  = ROOT / "sources" / "Precios_de_Alquiler.xls"
MUNI = ROOT / "sources" / "municipios_ine.csv"
OUT  = ROOT / "src" / "precios.json"

# ── helpers ──────────────────────────────────────────────────────────────────

def ascii_low(s):
    return unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode().lower()

def deartic(name):
    """'Rozas de Madrid (Las)' → 'Las Rozas de Madrid'"""
    m = re.match(r"^(.*?),\s*(El|La|Los|Las|L\u2019|L\'|Els|Les)$", name)
    return (m.group(2) + " " + m.group(1)).strip() if m else name

def slugify(name):
    d = deartic(name)
    d = re.sub(r"^l['\u2019]\s*", "", d, flags=re.I)
    b = ascii_low(d)
    return re.sub(r"[^a-z0-9]+", "-", b).strip("-")

def base_norm(name):
    d = deartic(name)
    d = re.sub(r"\([^)]*\)", "", d)
    return re.sub(r"[^a-z0-9]+", "", ascii_low(d))

def num(x):
    try:
        return round(float(x), 3)
    except (TypeError, ValueError):
        return None

def pct(x):
    """Yields: 4 decimales. Con round(...,3) el 7,25% se guardaba como 0.072 y
    el 7,75% como 0.077, perdiendo el cuarto de punto."""
    try:
        return round(float(x), 4)
    except (TypeError, ValueError):
        return None

# distritos de Madrid y Barcelona (slug manual, más fiable que el automático)
MAD_DIST = {
    "Arganzuela": "arganzuela", "Barajas": "barajas", "Carabanchel": "carabanchel",
    "Centro": "centro", "Chamartín": "chamartin", "Chamberí": "chamberi",
    "Ciudad Lineal": "ciudad-lineal", "Fuencarral": "fuencarral-el-pardo",
    "Hortaleza": "hortaleza", "Latina": "latina", "Moncloa": "moncloa-aravaca",
    "Mortalaz": "moratalaz", "Puente Vallecas": "puente-de-vallecas",
    "Retiro": "retiro", "Salamanca": "salamanca", "San Blas": "san-blas-canillejas",
    "Tetuán": "tetuan", "Usera": "usera", "Vicálvaro": "vicalvaro",
    "Villa de Vallecas": "villa-de-vallecas", "Villaverde": "villaverde",
}
BCN_DIST = {
    "Ciutat Vella": "ciutat-vella", "Eixample": "eixample", "Grácia": "gracia",
    "Horta Guinardó": "horta-guinardo", "Les Corts": "les-corts",
    "Nou Barris": "nou-barris", "Sant Andreu": "sant-andreu",
    "Sant Martí": "sant-marti", "Sants-Montjuic": "sants-montjuic",
    "Sarriá-Sant Gervassi": "sarria-sant-gervasi",
}
PROV_OVERRIDE = {
    "Balears (Illes)": "balears-illes", "Coruña (A)": "a-coruna",
    "Rioja (La)": "la-rioja", "Palmas (Las)": "las-palmas",
    "Alicante/Alacant": "alicante", "Castellón/Castelló": "castellon",
    "Valencia/València": "valencia",
}
BASE_URL = "https://www.idealista.com/alquiler-locales/"

# ── 1. Leer XLS ──────────────────────────────────────────────────────────────

print("Leyendo", XLS)
sheets = pd.read_excel(XLS, sheet_name=None, header=None)
data = []

for sheet_name, df in sheets.items():
    mode = "municipio"
    for i in range(df.shape[0]):
        c1  = df.iloc[i, 1]
        loc = df.iloc[i, 2]
        yld = df.iloc[i, 3]
        m6  = df.iloc[i, 6]
        if pd.notna(c1) and str(c1).strip() == "Población":
            mode = "distrito"
            continue
        if pd.isna(loc) or pd.isna(yld) or not isinstance(yld, (int, float)):
            continue
        if pd.isna(m6) or not isinstance(m6, (int, float)):
            continue

        loc_str = str(loc).strip()
        clean   = re.sub(r"\s*-\s*A\d$", "", loc_str)
        prov    = ("Madrid" if sheet_name == "Madrid"
                   else "Barcelona" if sheet_name == "Barcelona"
                   else str(c1).strip())

        # URL Idealista
        if sheet_name == "Madrid" and mode == "distrito":
            url = BASE_URL + "madrid/" + MAD_DIST.get(clean, slugify(clean)) + "/"
        elif sheet_name == "Barcelona" and mode == "distrito":
            url = BASE_URL + "barcelona/" + BCN_DIST.get(clean, slugify(clean)) + "/"
        elif sheet_name == "Madrid":
            url = BASE_URL + slugify(clean) + "-madrid/"
        elif sheet_name == "Barcelona":
            url = BASE_URL + slugify(clean) + "-barcelona/"
        else:
            ps = PROV_OVERRIDE.get(prov, slugify(prov))
            url = BASE_URL + slugify(clean) + "-" + ps + "/"

        def rng(a, b, med):
            va = num(df.iloc[i, a]); vb = num(df.iloc[i, b]); vm = num(df.iloc[i, med])
            if va is None or vb is None:
                return None
            return {"min": min(va, vb), "max": max(va, vb), "medio": vm}

        data.append({
            "provincia":   prov,
            "localizacion": loc_str,
            "tipo":        mode,
            "yield":       pct(yld),
            "idealista":   url,
            "hasData":     True,
            "r1": rng(4,  5,  6),
            "r2": rng(7,  8,  9),
            "r3": rng(10, 11, 12),
            "r4": rng(13, 14, 15),
        })

print(f"  {len(data)} registros del XLS")

# ── 2. Municipios de Madrid (INE) ────────────────────────────────────────────

print("Leyendo", MUNI)
existing = {base_norm(d["localizacion"]): d for d in data
            if d["provincia"] == "Madrid" and d["tipo"] == "municipio"}

added = 0
with open(MUNI, newline="", encoding="utf-8") as f:
    for row in csv.DictReader(f):
        if not row["municipio_id"].startswith("28"):
            continue
        nm   = row["nombre"]
        key  = base_norm(nm)
        disp = deartic(nm)
        url  = BASE_URL + slugify(nm) + "-madrid/"
        if key in existing:
            existing[key]["idealista"] = url   # actualiza slug
            continue
        data.append({
            "provincia":    "Madrid",
            "localizacion": disp,
            "tipo":         "municipio",
            "yield":        None,
            "idealista":    url,
            "hasData":      False,
            "r1": None, "r2": None, "r3": None, "r4": None,
        })
        added += 1

print(f"  {added} municipios de Madrid añadidos sin precio")

# ── 3. Guardar ────────────────────────────────────────────────────────────────

OUT.parent.mkdir(exist_ok=True)
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

print(f"✓ {OUT}  ({len(data)} registros, {OUT.stat().st_size // 1024} KB)")
