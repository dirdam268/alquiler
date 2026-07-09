# build.ps1 — genera la app final en app\ a partir de src\ y pwa\
# (equivalente en PowerShell de build_html.py, sin necesidad de Python)
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\build.ps1

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$src  = Join-Path $root "src"
$pwa  = Join-Path $root "pwa"
$app  = Join-Path $root "app"
New-Item -ItemType Directory -Force $app | Out-Null

$css  = [IO.File]::ReadAllText((Join-Path $src "styles.css"))
$js   = [IO.File]::ReadAllText((Join-Path $src "app.js"))
$data = [IO.File]::ReadAllText((Join-Path $src "precios.json")).Trim()
$tmpl = [IO.File]::ReadAllText((Join-Path $src "index.html"))

# Recuento para el pie (municipios de Madrid en el dataset)
$json  = $data | ConvertFrom-Json
$nMad  = @($json | Where-Object { $_.provincia -eq "Madrid" -and $_.tipo -eq "municipio" }).Count
$nTot  = @($json).Count

$footer = "Referencia de mercado: Precios_de_Alquiler.xls (2009) - $nTot ubicaciones, " +
    "incl. los $nMad municipios de la Comunidad de Madrid (15 con precio; " +
    "el resto enlazan a Idealista sin precio inventado). " +
    "Tasa de esfuerzo seg&#250;n Tabla_calculo_renta_locales.xlsx: ventas netas = ventas/1,104; " +
    "venta media diaria = netas/(350 si abre domingos, 308 si no); " +
    "tasa = renta mensual / venta media diaria; &gt;80% Riesgo, 50-80% Viable, &lt;50% &#211;ptimo. " +
    "El precio medio actual se consulta en Idealista. Orientativo, no es una tasaci&#243;n."

$html = $tmpl.
    Replace("<!-- INJECT:CSS -->",    "<style>`n$css`n</style>").
    Replace("<!-- INJECT:DATA -->",   "<script>`nwindow.DATA = $data;`n</script>").
    Replace("<!-- INJECT:JS -->",     "<script>`n$js`n</script>").
    Replace("<!-- INJECT:FOOTER -->", $footer)

$utf8 = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $app "index.html"), $html, $utf8)

# Copiar activos PWA (manifest, service worker, iconos)
Copy-Item (Join-Path $pwa "*") $app -Force

$kb = [int]((Get-Item (Join-Path $app "index.html")).Length / 1KB)
Write-Host "OK  app\index.html ($kb KB) + activos PWA copiados"
