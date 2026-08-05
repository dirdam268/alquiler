# build.ps1 — genera la app EN CLARO en build\index.html a partir de src\ y pwa\
# (equivalente en PowerShell de build_html.py, sin necesidad de Python)
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\build.ps1
#
# IMPORTANTE: build\index.html lleva los alquileres reales de las tiendas y NO
# se publica: esta en .gitignore. Para publicar hay que cifrarlo despues con
#   powershell -ExecutionPolicy Bypass -File scripts\build-secure.ps1 -Password '<clave>'
# que es quien escribe el app\index.html cifrado que sí va a GitHub.

$ErrorActionPreference = "Stop"
$root  = Split-Path $PSScriptRoot -Parent
$src   = Join-Path $root "src"
$pwa   = Join-Path $root "pwa"
$app   = Join-Path $root "app"
$build = Join-Path $root "build"
New-Item -ItemType Directory -Force $app   | Out-Null
New-Item -ItemType Directory -Force $build | Out-Null

$css  = [IO.File]::ReadAllText((Join-Path $src "styles.css"))
$js   = [IO.File]::ReadAllText((Join-Path $src "app.js"))
$data = [IO.File]::ReadAllText((Join-Path $src "precios.json")).Trim()
$tien = [IO.File]::ReadAllText((Join-Path $src "tiendas.json")).Trim()
$tmpl = [IO.File]::ReadAllText((Join-Path $src "index.html"))

# Recuento para el pie (municipios de Madrid en el dataset)
$json  = $data | ConvertFrom-Json
$nMad  = @($json | Where-Object { $_.provincia -eq "Madrid" -and $_.tipo -eq "municipio" }).Count
$nTot  = @($json).Count

# Recuento de alquileres reales (fichero de tiendas Express)
$jsonT  = $tien | ConvertFrom-Json
$nTien  = @($jsonT).Count

$footer = "Referencia de mercado: Precios_de_Alquiler.xls (2009) - $nTot ubicaciones, " +
    "incl. los $nMad municipios de la Comunidad de Madrid (15 con precio; " +
    "el resto enlazan a Idealista sin precio inventado). " +
    "Tasa de esfuerzo seg&#250;n Tabla_calculo_renta_locales.xlsx: ventas netas = ventas/1,104; " +
    "venta media diaria = netas/(350 si abre domingos, 308 si no); " +
    "tasa = renta mensual / venta media diaria; &gt;80% Riesgo, 50-80% Viable, &lt;50% &#211;ptimo. " +
    "Alquileres reales: Datos economicos tiendas Express.xlsx (09/2016) - $nTien inmuebles, " +
    "renta + gastos repercutidos; el fichero no trae superficie, por eso no hay &#8364;/m&#178;. " +
    "El precio medio actual se consulta en Idealista. Orientativo, no es una tasaci&#243;n."

$html = $tmpl.
    Replace("<!-- INJECT:CSS -->",     "<style>`n$css`n</style>").
    Replace("<!-- INJECT:DATA -->",    "<script>`nwindow.DATA = $data;`n</script>").
    Replace("<!-- INJECT:TIENDAS -->", "<script>`nwindow.TIENDAS = $tien;`n</script>").
    Replace("<!-- INJECT:JS -->",      "<script>`n$js`n</script>").
    Replace("<!-- INJECT:FOOTER -->",  $footer)

$utf8 = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $build "index.html"), $html, $utf8)

# Copiar activos PWA (manifest, service worker, iconos): no son sensibles y van
# tal cual junto al index.html cifrado.
Copy-Item (Join-Path $pwa "*") $app -Force

$kb = [int]((Get-Item (Join-Path $build "index.html")).Length / 1KB)
Write-Host "OK  build\index.html ($kb KB, EN CLARO, no publicar) + activos PWA en app\"
Write-Host "    Siguiente paso: scripts\build-secure.ps1 -Password '<clave>'  ->  app\index.html cifrado"
