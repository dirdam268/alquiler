# build_cuenta.ps1 - extrae el modelo de cuenta de explotacion (xlsx interno) y
# genera src\cuenta.json, que build.ps1 inyecta en la app como window.CUENTA.
#
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\build_cuenta.ps1 [ruta_xlsx]
#
# Lee del Excel TODO lo editable y sus textos: la hoja de entrada de datos (5
# datos por talla), la de hipotesis (con su unidad y su criterio) y la guia
# (pasos, cautelas, controles, historial). Las formulas se reproducen en app.js;
# asi la app permite editar cada cifra y recalcula la cuenta entera.
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

# Fila de la hoja -> clave de parametro que usa el motor de app.js
$MAP_ENTRADA = [ordered]@{ 7='sqm'; 8='rent_monthly'; 9='equipment'; 10='works'; 11='gross_sales' }
$MAP_HIPOTESIS = [ordered]@{
  7='vat'; 8='gross_margin'; 9='shrink'; 10='personnel_pct'; 11='personnel_reduction';
  15='consumables'; 16='repairs'; 17='prl'; 18='energy_ref'; 19='energy_ref_sqm';
  20='security'; 21='communications'; 22='other_taxes'; 23='cards'; 24='rappel';
  25='brand'; 26='admin'; 27='other';
  31='renting_rate'; 32='renting_months'; 33='bank_share'; 34='bank_rate';
  35='bank_months'; 36='grace_months'; 37='depreciation_years'; 38='tax_rate';
  42='ipc'; 43='g2'; 44='g3'; 45='g4'
}
# Titulo de cada bloque de la hoja de hipotesis -> filas que contiene
$MAP_BLOQUES = @( @{ t=4; a=7; b=11 }, @{ t=13; a=15; b=27 }, @{ t=29; a=31; b=38 }, @{ t=40; a=42; b=45 } )
# Fila de la hoja de cuentas -> clave de linea
$MAP_FILAS = [ordered]@{
  6='ventas'; 8='margen_bruto'; 9='merma'; 10='margen_comercial'; 11='total_gastos';
  13='gastos_comerciales'; 15='personal'; 16='consumibles'; 17='reparacion'; 18='prl';
  19='energia'; 20='seguridad'; 21='comunicaciones'; 22='alquiler'; 23='renting';
  24='tributos'; 25='tarjetas'; 26='rappel'; 27='ensena'; 28='admin'; 29='resto';
  31='cash_flow'; 32='amortizacion'; 33='res_explotacion'; 34='gastos_fin';
  35='bai'; 36='impuesto'; 37='beneficio'
}
$MAP_TRAMOS = @( @{ id='bajo'; col=3 }, @{ id='medio'; col=4 }, @{ id='alto'; col=5 } )

function Tipo([string]$unidad) {
  $u = ("" + $unidad).Trim()
  if ($u.StartsWith("%")) { return "pct" }
  if ($u.Contains([string][char]0x20AC)) { return "eur" }
  return "num"
}
# En un [ordered] con claves numericas, $h[7] es la POSICION 7, no la clave 7.
function MapKey($map, [int]$k) {
  foreach ($e in $map.GetEnumerator()) { if ([int]$e.Key -eq $k) { return $e.Value } }
  return $null
}
function Txt($ws, $r, $c) { return ("" + $ws.Cells.Item([int]$r, [int]$c).Text).Trim() }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($Xlsx, 0, $true)
  if ($wb.Worksheets.Count -lt 6 -or -not ($wb.Worksheets.Item(1).Name -like "Entrada*")) {
    throw "Este no es el modelo con hoja 'Entrada datos' (6 hojas). Pasa el modelo de ventas editables."
  }
  $en = $wb.Worksheets.Item(1)   # Entrada datos
  $cm = $wb.Worksheets.Item(2)   # Cuentas (Ano 1)
  $hp = $wb.Worksheets.Item(4)   # Hipotesis
  $gu = $wb.Worksheets.Item(6)   # Guia y controles

  $lineas = [ordered]@{}
  foreach ($e in $MAP_FILAS.GetEnumerator()) { $lineas[$e.Value] = Txt $cm ([int]$e.Key) 1 }

  # --- entrada de datos (las 15 celdas amarillas) ---
  $entradaFilas = @()
  foreach ($e in $MAP_ENTRADA.GetEnumerator()) {
    $r = [int]$e.Key
    $vals = [ordered]@{}
    foreach ($t in $MAP_TRAMOS) { $vals[$t.id] = [double]$en.Cells.Item($r, $t.col).Value2 }
    $u = Txt $en $r 2
    $entradaFilas += [ordered]@{ key=$e.Value; etiqueta=(Txt $en $r 1); unidad=$u; tipo=(Tipo $u); criterio=(Txt $en $r 6); valores=$vals }
  }
  # --- indicadores calculados (no editables) ---
  $indicadores = @()
  $MAP_IND = [ordered]@{ 12='inversion_total'; 25='fondos_propios'; 26='prestamo'; 27='inversion_m2'; 28='ventas_m2' }
  foreach ($e in $MAP_IND.GetEnumerator()) {
    $r = [int]$e.Key
    $indicadores += [ordered]@{ key=$e.Value; etiqueta=(Txt $en $r 1); unidad=(Txt $en $r 2); criterio=(Txt $en $r 6) }
  }

  # --- hipotesis ---
  $hipotesis = @()
  foreach ($b in $MAP_BLOQUES) {
    $filas = @()
    for ($r = $b.a; $r -le $b.b; $r++) {
      if (-not $MAP_HIPOTESIS.Contains($r)) { continue }
      $vals = [ordered]@{}
      foreach ($t in $MAP_TRAMOS) { $vals[$t.id] = [double]$hp.Cells.Item($r, $t.col).Value2 }
      $u = Txt $hp $r 2
      $filas += [ordered]@{ key=(MapKey $MAP_HIPOTESIS $r); etiqueta=(Txt $hp $r 1); unidad=$u; tipo=(Tipo $u); criterio=(Txt $hp $r 6); valores=$vals }
    }
    $hipotesis += [ordered]@{ titulo=(Txt $hp $b.t 1); filas=$filas }
  }

  # --- guia ---
  $pasos = @(); foreach ($r in 3..5) { $pasos += (Txt $gu $r 1) }
  $origen = @(); foreach ($r in 8..9) { $origen += (Txt $gu $r 1) }
  $cautelas = @(); foreach ($r in 37..42) { $cautelas += (Txt $gu $r 1) }
  $controles = @(); foreach ($r in 28..34) { $controles += [ordered]@{ etiqueta=(Txt $gu $r 1); criterio=(Txt $gu $r 6) } }
  $historial = @()
  foreach ($r in 12..24) {
    $historial += [ordered]@{ concepto=(Txt $gu $r 1); origen=(Txt $gu $r 2); bajo=(Txt $gu $r 3); medio=(Txt $gu $r 4); alto=(Txt $gu $r 5); nota=(Txt $gu $r 6) }
  }

  $esc = @()
  foreach ($t in $MAP_TRAMOS) {
    $esc += [ordered]@{ id=$t.id; nombre=(Get-Culture).TextInfo.ToTitleCase((Txt $en 6 $t.col).ToLower()) }
  }

  $out = [ordered]@{
    titulo     = Txt $cm 1 1
    fuente     = (Split-Path $Xlsx -Leaf)
    lineas     = $lineas
    tramos     = $esc
    entrada    = [ordered]@{ titulo=(Txt $en 5 1); intro=(Txt $en 3 1); nota=(Txt $en 13 1); cashflow=(Txt $en 30 1); filas=$entradaFilas }
    indicadores = $indicadores
    hipotesis  = $hipotesis
    guia       = [ordered]@{ titulo=(Txt $gu 1 1); pasos=$pasos; origen=$origen; cautelas=$cautelas; controles=$controles; historial=$historial }
  }
  $wb.Close($false)
} finally {
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

$json = $out | ConvertTo-Json -Depth 8 -Compress
[IO.File]::WriteAllText((Join-Path $root "src\cuenta.json"), $json, (New-Object Text.UTF8Encoding($false)))
$nH = 0; foreach ($b in $hipotesis) { $nH += $b.filas.Count }
Write-Host "OK  src\cuenta.json - $($entradaFilas.Count) datos de entrada, $nH hipotesis, $($controles.Count) controles, $($cautelas.Count) cautelas"
