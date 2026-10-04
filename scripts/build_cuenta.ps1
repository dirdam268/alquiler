# build_cuenta.ps1 - extrae el modelo de cuenta de explotacion (xlsx interno) y
# genera src\cuenta.json, que build.ps1 inyecta en la app como window.CUENTA.
#
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\build_cuenta.ps1 [ruta_xlsx]
#
# Del Excel solo se leen las HIPOTESIS (columnas L:O de la hoja de cuentas y los
# crecimientos de la hoja de 10 anos). Las formulas se reproducen en app.js; asi
# la app recalcula la cuenta con la venta y el alquiler que ponga el usuario.
#
# src\cuenta.json es dato interno: esta en .gitignore y solo viaja cifrado.
# Este script es ASCII puro a proposito (no depende de la codificacion).

param(
  [string]$Xlsx = ""
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent

# Sin ruta: el modelo mas reciente de Descargas que case con el patron. Se busca
# por patron (y no por nombre fijo) para no dejar el nombre del fichero en el repo.
if (-not $Xlsx) {
  $f = Get-ChildItem (Join-Path $env:USERPROFILE "Downloads") -Filter "Modelo cuenta*franquicia*.xlsx" -ErrorAction SilentlyContinue |
       Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($f) { $Xlsx = $f.FullName }
}
if (-not $Xlsx -or -not (Test-Path $Xlsx)) { throw "No encuentro el modelo de cuenta en Descargas (Modelo cuenta*franquicia*.xlsx). Pasa la ruta como parametro." }

# Fila del Excel -> clave de la linea de la cuenta (la etiqueta se lee del fichero)
$FILAS = [ordered]@{
  6='ventas'; 8='margen_bruto'; 9='merma'; 10='margen_comercial'; 11='total_gastos';
  13='gastos_comerciales'; 15='personal'; 16='consumibles'; 17='reparacion'; 18='prl';
  19='energia'; 20='seguridad'; 21='comunicaciones'; 22='alquiler'; 23='renting';
  24='tributos'; 25='tarjetas'; 26='rappel'; 27='ensena'; 28='admin'; 29='resto';
  31='cash_flow'; 32='amortizacion'; 33='res_explotacion'; 34='gastos_fin';
  35='bai'; 36='impuesto'; 37='beneficio'
}

$inv = [Globalization.CultureInfo]::InvariantCulture
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($Xlsx, 0, $true)
  $cm = $wb.Worksheets.Item(1)   # cuentas del Ano 1
  $md = $wb.Worksheets.Item(2)   # media 10 anos

  # Ojo: en un [ordered] con claves numericas, $FILAS[6] es la POSICION 6, no la
  # clave 6. Por eso se recorren los pares en vez de indexar.
  $lineas = [ordered]@{}
  foreach ($e in $FILAS.GetEnumerator()) {
    $lineas[$e.Value] = ("" + $cm.Cells.Item([int]$e.Key, 1).Text).Trim()
  }

  # Escenarios: columnas M, N, O (13..15) de hipotesis; bloques A/E/I de titulo;
  # crecimientos en la hoja de 10 anos: M5:V5, X5:AG5, AI5:AR5
  $esc = @()
  $defs = @(
    @{ id='bajo';  col=13; tit=1; g0=13 },
    @{ id='medio'; col=14; tit=5; g0=24 },
    @{ id='alto';  col=15; tit=9; g0=35 }
  )
  foreach ($d in $defs) {
    $p = [ordered]@{}
    for ($r = 2; $r -le 33; $r++) {
      $k = $cm.Cells.Item($r, 12).Value2
      if (-not $k) { continue }
      $v = $cm.Cells.Item($r, $d.col).Value2
      if ($v -is [double]) { $p["$k"] = [math]::Round($v, 6) }
    }
    $g = @()
    for ($c = $d.g0; $c -lt $d.g0 + 10; $c++) { $g += [math]::Round([double]$md.Cells.Item(5, $c).Value2, 6) }
    $nombre = ("" + $cm.Cells.Item(3, $d.tit).Text) -replace '(?i)^\s*TIENDA\s+', ''
    $esc += [ordered]@{
      id = $d.id
      nombre = (Get-Culture).TextInfo.ToTitleCase($nombre.Trim().ToLower())
      params = $p
      crecimiento = $g
    }
  }

  $out = [ordered]@{
    titulo    = ("" + $cm.Cells.Item(1, 1).Text).Trim()
    fuente    = (Split-Path $Xlsx -Leaf)
    lineas    = $lineas
    escenarios = $esc
  }
  $wb.Close($false)
} finally {
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

$json = $out | ConvertTo-Json -Depth 6 -Compress
# ConvertTo-Json escribe los numeros con la cultura invariante: no hay comas decimales
[IO.File]::WriteAllText((Join-Path $root "src\cuenta.json"), $json, (New-Object Text.UTF8Encoding($false)))
$ids = ($esc | ForEach-Object { "$($_.id) $($_.params.sqm) m2" }) -join ", "
Write-Host "OK  src\cuenta.json - escenarios: $ids"
