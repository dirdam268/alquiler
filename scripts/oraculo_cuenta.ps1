# oraculo_cuenta.ps1 - prueba del motor de la cuenta contra el propio Excel.
#
# Abre el modelo en SOLO LECTURA, le cambia datos de entrada e hipotesis EN
# MEMORIA (no guarda nada), deja que Excel recalcule y vuelca lo que calcula a
# un json. Despues la app ejecuta su motor con los mismos parametros y se
# comparan todas las cifras. Es la comprobacion de que la cuenta de la app es
# la del Excel, tambien con datos distintos a los de partida.
#
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\oraculo_cuenta.ps1 -Out <ruta.json> [-Xlsx <ruta>]

param(
  [Parameter(Mandatory=$true)][string]$Out,
  [string]$Xlsx = ""
)
$ErrorActionPreference = "Stop"
if (-not $Xlsx) {
  $f = Get-ChildItem (Join-Path $env:USERPROFILE "Downloads") -Filter "Modelo cuenta*editables*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $f) { throw "No encuentro el modelo de ventas editables en Descargas" }
  $Xlsx = $f.FullName
}

$ENT = @{ sqm=7; rent_monthly=8; equipment=9; works=10; gross_sales=11 }
$HIP = @{ vat=7; gross_margin=8; shrink=9; personnel_pct=10; personnel_reduction=11;
  consumables=15; repairs=16; prl=17; energy_ref=18; energy_ref_sqm=19; security=20;
  communications=21; other_taxes=22; cards=23; rappel=24; brand=25; admin=26; other=27;
  renting_rate=31; renting_months=32; bank_share=33; bank_rate=34; bank_months=35;
  grace_months=36; depreciation_years=37; tax_rate=38; ipc=42; g2=43; g3=44; g4=45 }
$LIN = @('ventas','margen_bruto','merma','margen_comercial','total_gastos','gastos_comerciales','personal','consumibles','reparacion','prl','energia','seguridad','comunicaciones','alquiler','renting','tributos','tarjetas','rappel','ensena','admin','resto','cash_flow','amortizacion','res_explotacion','gastos_fin','bai','impuesto','beneficio')
# fila de cada linea en la hoja de cuentas / media (A1) y en la proyeccion (bloque bajo; +40 medio, +80 alto)
$F1 = @(6,8,9,10,11,13,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,31,32,33,34,35,36,37)
$FP = @(6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33)
$COLS = @{ bajo=@{c=3; a1=2; off=0}; medio=@{c=4; a1=6; off=40}; alto=@{c=5; a1=10; off=80} }

# Escenarios: [hoja]:[clave] = valor por talla (bajo, medio, alto). Fila 'Entrada' o 'Hip'.
$ESCENARIOS = @(
  @{ nombre='base'; cambios=@() },
  @{ nombre='editado'; cambios=@(
      @('E','gross_sales',1400000,700000,2000000), @('E','rent_monthly',6000,1500,9000),
      @('E','sqm',90,240,310), @('E','equipment',150000,40000,210000), @('E','works',60000,300000,20000),
      @('H','vat',0.21,0.10,0.07), @('H','gross_margin',0.24,0.31,0.28), @('H','shrink',0.03,0.015,0.02),
      @('H','personnel_pct',0.15,0.11,0.125), @('H','personnel_reduction',0,12000,3000),
      @('H','renting_months',60,120,84), @('H','renting_rate',0.06,0.03,0.045),
      @('H','bank_months',120,60,96), @('H','grace_months',12,0,3), @('H','bank_share',0.5,0.9,0.7),
      @('H','bank_rate',0.07,0.04,0.065), @('H','tax_rate',0.30,0.20,0.25), @('H','depreciation_years',10,25,15),
      @('H','ipc',0.03,0.015,0.025), @('H','g2',0.2,0.0,0.05), @('H','g3',0.0,0.1,0.07), @('H','g4',0.05,-0.02,0.04),
      @('H','energy_ref_sqm',100,200,150), @('H','energy_ref',18000,30000,22800)
  ) },
  @{ nombre='limites'; cambios=@(
      @('E','works',0,200000,80000), @('H','bank_share',0.7,0,1.0),
      @('H','renting_months',200,12,119), @('H','bank_months',36,24,84), @('H','grace_months',0,6,12),
      @('H','depreciation_years',5,3,40), @('E','gross_sales',300000,3000000,1100000),
      @('H','personnel_reduction',90000,0,5000)
  ) }
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$todo = [ordered]@{}
try {
  $wb = $xl.Workbooks.Open($Xlsx, 0, $true)
  $en = $wb.Worksheets.Item(1); $cm = $wb.Worksheets.Item(2); $md = $wb.Worksheets.Item(3)
  $hp = $wb.Worksheets.Item(4); $pr = $wb.Worksheets.Item(5)

  # valores de partida, para volver a ellos entre escenarios
  $orig = @{}
  foreach ($k in $ENT.Keys) { foreach ($t in $COLS.Keys) { $orig["E|$k|$t"] = $en.Cells.Item($ENT[$k], $COLS[$t].c).Value2 } }
  foreach ($k in $HIP.Keys) { foreach ($t in $COLS.Keys) { if ($k -ne 'gross_sales') { $orig["H|$k|$t"] = $hp.Cells.Item($HIP[$k], $COLS[$t].c).Value2 } } }

  foreach ($esc in $ESCENARIOS) {
    # restaurar
    foreach ($k in $orig.Keys) {
      $p = $k.Split('|'); $c = $COLS[$p[2]].c
      if ($p[0] -eq 'E') { $en.Cells.Item($ENT[$p[1]], $c).Value2 = $orig[$k] } else { $hp.Cells.Item($HIP[$p[1]], $c).Value2 = $orig[$k] }
    }
    # aplicar cambios
    foreach ($ch in $esc.cambios) {
      $hoja = $ch[0]; $k = $ch[1]; $i = 2
      foreach ($t in @('bajo','medio','alto')) {
        $c = $COLS[$t].c
        if ($hoja -eq 'E') { $en.Cells.Item($ENT[$k], $c).Value2 = [double]$ch[$i] } else { $hp.Cells.Item($HIP[$k], $c).Value2 = [double]$ch[$i] }
        $i++
      }
    }
    $xl.Calculate(); $xl.CalculateFull()

    $res = [ordered]@{}
    foreach ($t in @('bajo','medio','alto')) {
      $c = $COLS[$t].c
      $params = [ordered]@{}
      foreach ($k in $ENT.Keys) { $params[$k] = [double]$en.Cells.Item($ENT[$k], $c).Value2 }
      foreach ($k in $HIP.Keys) { $params[$k] = [double]$hp.Cells.Item($HIP[$k], $c).Value2 }
      $a1 = [ordered]@{}; $me = [ordered]@{}
      for ($i = 0; $i -lt $LIN.Count; $i++) {
        $a1[$LIN[$i]] = [double]$cm.Cells.Item($F1[$i], $COLS[$t].a1).Value2
        $me[$LIN[$i]] = [double]$md.Cells.Item($F1[$i], $COLS[$t].a1).Value2
      }
      $anios = @()
      for ($y = 2; $y -le 11; $y++) {
        $fila = [ordered]@{}
        for ($i = 0; $i -lt $LIN.Count; $i++) { $fila[$LIN[$i]] = [double]$pr.Cells.Item($FP[$i] + $COLS[$t].off, $y).Value2 }
        $anios += $fila
      }
      $res[$t] = [ordered]@{ params=$params; anio1=$a1; media=$me; anios=$anios }
    }
    $todo[$esc.nombre] = $res
  }
  $wb.Close($false)   # sin guardar: el fichero original no se toca
} finally {
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
[IO.File]::WriteAllText($Out, ($todo | ConvertTo-Json -Depth 8 -Compress), (New-Object Text.UTF8Encoding($false)))
Write-Host ("OK  oraculo escrito: {0} escenarios -> {1}" -f $todo.Count, $Out)
