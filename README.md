# Alquiler

Aplicación (PWA) para móvil y ordenador que compara una renta de local comercial con el precio de mercado por zona y calcula la **tasa de esfuerzo** según la metodología propia. Funciona 100% offline una vez instalada.

---

## Cómo usarla / instalarla

- **Ordenador (sin instalar):** abrir `app\index.html` con doble clic. Todo va embebido en ese único archivo.
- **Ordenador (como app):** servirla en local (`powershell -ExecutionPolicy Bypass -File scripts\serve.ps1` → http://localhost:8765) o publicada en una URL, y en Chrome/Edge pulsar el icono **Instalar** de la barra de direcciones. Se crea acceso directo y ventana propia.
- **Móvil:** necesita estar publicada en una URL https (por ejemplo GitHub Pages, subiendo la carpeta `app\`). Después, en Android: menú de Chrome → **Añadir a pantalla de inicio**; en iPhone: Safari → Compartir → **Añadir a pantalla de inicio**. Queda como una app más, con su icono, y funciona sin conexión.

---

## Estructura del proyecto

```
ALQUILER/
├── sources/                      # Ficheros fuente originales (no modificar)
│   ├── Precios_de_Alquiler.xls        # Precios de referencia por distrito/municipio (2009)
│   ├── Tabla_calculo_renta_locales.xlsx  # Metodología tasa de esfuerzo
│   └── municipios_ine.csv             # Municipios de España (INE, vía CodeForSpain)
│
├── src/                          # Código fuente separado
│   ├── index.html                     # Plantilla con marcadores <!-- INJECT:* -->
│   ├── styles.css                     # Estilos (dark theme, variables CSS)
│   ├── app.js                         # Lógica: autocomplete, cálculo, renders
│   └── precios.json                   # Dataset generado por build_data.py (no editar a mano)
│
├── pwa/                          # Activos de la app instalable
│   ├── manifest.webmanifest           # Nombre, colores, iconos de la app
│   ├── sw.js                          # Service worker (funcionamiento offline)
│   └── icon-*.png                     # Iconos (generados por make_icons.ps1)
│
├── scripts/                      # Scripts de build
│   ├── build.ps1                      # Genera app\ (PowerShell, sin dependencias)
│   ├── make_icons.ps1                 # Regenera los iconos PNG
│   ├── serve.ps1                      # Servidor local de prueba (puerto 8765)
│   ├── build_data.py                  # Genera src/precios.json (requiere Python)
│   └── build_html.py                  # Build original en Python (alternativa a build.ps1)
│
└── app/                          # Salida de producción — carpeta lista para publicar
    ├── index.html                     # HTML único autocontenido (CSS+JS+datos embebidos)
    ├── manifest.webmanifest / sw.js / icon-*.png
```

---

## Flujo de build

```
sources/*.xls + municipios_ine.csv
        ↓  build_data.py (Python, solo si cambian los datos)
    src/precios.json
        ↓  build.ps1  (+ src/styles.css + src/app.js + src/index.html + pwa/*)
    app/   ← carpeta final para usar/publicar
```

### Comandos

```powershell
# Compilar la app (cada vez que cambies algo en src\ o pwa\)
powershell -ExecutionPolicy Bypass -File scripts\build.ps1

# Regenerar iconos (solo si cambias el diseño del icono)
powershell -ExecutionPolicy Bypass -File scripts\make_icons.ps1

# Probar en local
powershell -ExecutionPolicy Bypass -File scripts\serve.ps1
# → http://localhost:8765

# (Opcional, requiere Python) Regenerar dataset si cambian los XLS/CSV
pip install xlrd pandas openpyxl
python scripts/build_data.py
```

---

## Metodología

### Precio de referencia de mercado

Fuente: `Precios_de_Alquiler.xls` (datos de 2009). Cuatro rangos de superficie:

| Rango | m² |
|-------|----|
| r1 | 300 – 600 |
| r2 | 1.000 – 1.400 |
| r3 | 1.400 – 1.800 |
| r4 | 1.800 – 2.200 |

Cada rango tiene `min`, `max` y `medio` en €/m²/mes. El veredicto compara la renta introducida con estos umbrales:

- **Por debajo de mercado**: `unit < min`
- **Buen precio**: `min ≤ unit ≤ medio`
- **En mercado**: `medio < unit ≤ max`
- **Por encima de mercado**: `unit > max`

### Ajuste por eje comercial (coeficiente de calle)

El fichero llega a nivel de distrito/municipio. Se aplica un multiplicador sobre la referencia del distrito:

| Tipo | Coeficiente |
|------|-------------|
| Eje prime / primera línea | ×2,2 |
| Calle comercial principal  | ×1,4 |
| Calle media (referencia)   | ×1,0 |
| Calle secundaria            | ×0,65 |

Las calles prime conocidas (Montera, Preciados, Serrano, Portal de l'Àngel, Passeig de Gràcia…) se autodetectan y precargan el coeficiente.

> ⚠️ El coeficiente escala la referencia del distrito. **No es un precio medido por calle.** Para el precio real y actual, usar el enlace a Idealista.

### Tasa de esfuerzo

Basada en `Tabla_calculo_renta_locales.xlsx`:

```
CVN (venta neta) = CVB (ventas brutas) / 1,104
Venta media diaria = CVN / días_año
  - Abre domingos:   350 días
  - Cierra domingos: 308 días

Tasa de esfuerzo = Renta mensual / Venta media diaria

Valoración:
  > 80%         → Riesgo
  50% – 80%     → Viable
  < 50%         → Óptimo
```

### Municipios de Madrid

El dataset base del XLS incluye 15 municipios de Madrid con precio. Se completan con el listado oficial INE (179 municipios totales de la provincia, vía [CodeForSpain](https://github.com/codeforspain/ds-organizacion-administrativa)). Los municipios sin precio en el XLS **no reciben un precio inventado**: solo se muestra el enlace a Idealista.

---

## Ideas de desarrollo (roadmap)

### Datos
- [ ] Actualizar precios de referencia con fuente más reciente (TINSA, Savills, datos abiertos)
- [ ] Añadir precios por m² de venta además de alquiler
- [ ] Incorporar fuente de precios de local por calle real (si se consigue)
- [ ] Ampliar municipios con precio (actualmente solo 15 de 180 en Madrid)

### Funcionalidad
- [ ] Proyección plurianual (1–20 años) con IPC + crecimiento extra, igual que el Excel
- [ ] Exportar resultado a PDF / compartir URL con parámetros
- [ ] Comparar dos locales en paralelo
- [ ] Cálculo inverso: dado un objetivo de tasa de esfuerzo, ¿cuánto debo vender?
- [ ] Historial de consultas (localStorage)
- [ ] Modo "¿cuánto debería pagar?" dado un CVB y una tasa objetivo

### UX
- [ ] Buscar por código postal además de por nombre de municipio
- [ ] Ampliar lista de calles prime conocidas (especialmente fuera de Madrid/Barcelona)
- [ ] Añadir tooltip con definición de tasa de esfuerzo en la UI

---

## Dependencias

Solo Python estándar + pandas/xlrd para el build. El HTML final es **100% offline**, sin CDN ni dependencias externas en runtime.

```
Python >= 3.9
pandas
xlrd       (para .xls legacy)
openpyxl   (para .xlsx)
```
