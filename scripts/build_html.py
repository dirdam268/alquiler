"""
build_html.py
Genera dist/tasador_alquiler.html a partir de:
  src/precios.json   — dataset generado por build_data.py
  src/styles.css     — estilos
  src/app.js         — lógica
  src/index.html     — plantilla (etiquetas <!-- INJECT:* --> como marcadores)

Uso:
  python scripts/build_html.py
"""

import json, re
from pathlib import Path

ROOT = Path(__file__).parent.parent
SRC  = ROOT / "src"
DIST = ROOT / "dist"
DIST.mkdir(exist_ok=True)

css    = (SRC / "styles.css").read_text(encoding="utf-8")
js     = (SRC / "app.js").read_text(encoding="utf-8")
data   = json.loads((SRC / "precios.json").read_bytes())
tmpl   = (SRC / "index.html").read_text(encoding="utf-8")
js_data = "window.DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";"

n_mad = sum(1 for d in data if d["provincia"] == "Madrid" and d["tipo"] == "municipio")
footer = (
    f"Referencia de mercado: Precios_de_Alquiler.xls (2009) · {len(data)} ubicaciones, "
    f"incl. los {n_mad} municipios de la Comunidad de Madrid (15 con precio; "
    "el resto enlazan a Idealista sin precio inventado). "
    "Tasa de esfuerzo según Tabla_calculo_renta_locales.xlsx: ventas netas = ventas/1,104; "
    "venta media diaria = netas/(350 si abre domingos, 308 si no); "
    "tasa = renta mensual / venta media diaria; >80% Riesgo, 50–80% Viable, <50% Óptimo. "
    "El precio medio actual se consulta en Idealista. Orientativo, no es una tasación."
)

html = (tmpl
    .replace("<!-- INJECT:CSS -->",    f"<style>\n{css}\n</style>")
    .replace("<!-- INJECT:DATA -->",   f"<script>\n{js_data}\n</script>")
    .replace("<!-- INJECT:JS -->",     f"<script>\n{js}\n</script>")
    .replace("<!-- INJECT:FOOTER -->", footer)
)

out = DIST / "tasador_alquiler.html"
out.write_text(html, encoding="utf-8")
print(f"✓ {out}  ({out.stat().st_size // 1024} KB)")
