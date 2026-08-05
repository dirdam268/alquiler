# build_tiendas.ps1 - extrae los alquileres reales pagados desde el fichero de
# datos economicos de tiendas Express (export SAP) y genera src\tiendas.json
#
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\build_tiendas.ps1 [ruta_xlsx]
#
# Notas sobre el origen (importantes para entender el parseo):
#  - El texto del asiento viene truncado a 40 caracteres por SAP.
#  - Un mismo local puede aparecer en varios asientos con el texto ligeramente
#    distinto (dobles espacios, coma y numero de portal) porque el inmueble esta
#    repartido entre varios propietarios. Por eso se agrupa por
#    codigo + tipo + calle normalizada, no por el texto literal.
#  - Un mismo codigo de tienda puede tener local + almacen + espacio de vending;
#    se guardan por separado y etiquetados, porque un almacen no es comparable
#    con un local a pie de calle.
#  - Se descartan las periodificaciones (GASTOS ANTICIPADOS): no son cuota mensual.
#  - El fichero NO trae superficie, asi que no es posible calcular EUR/m2.

param(
  [string]$Xlsx = "C:\Users\Usuario\Downloads\Datos economicos tiendas Express.xlsx"
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent

# -- 1. Descomprimir el xlsx (es un zip) -------------------------------------
$work = Join-Path $env:TEMP "claude\tiendas_build"
if (Test-Path $work) { Remove-Item -Recurse -Force $work }
New-Item -ItemType Directory -Force $work | Out-Null
$tmpZip = Join-Path $work "_book.zip"
Copy-Item $Xlsx $tmpZip -Force
Expand-Archive -Path $tmpZip -DestinationPath $work -Force

# -- 2. Cadenas compartidas --------------------------------------------------
[xml]$ss = Get-Content (Join-Path $work "xl\sharedStrings.xml") -Encoding UTF8
$strings = @()
foreach ($si in $ss.sst.si) {
  $txt = if ($si.t -is [string]) { $si.t }
         elseif ($si.t.'#text')  { $si.t.'#text' }
         else { ($si.r | ForEach-Object { if ($_.t -is [string]) { $_.t } else { $_.t.'#text' } }) -join '' }
  $strings += $txt
}

# -- 3. Filas ----------------------------------------------------------------
[xml]$sh = Get-Content (Join-Path $work "xl\worksheets\sheet1.xml") -Encoding UTF8
$rows = @()
foreach ($r in $sh.worksheet.sheetData.row) {
  $cells = @{}
  foreach ($c in $r.c) {
    $v = $c.v
    if ($c.t -eq 's' -and $null -ne $v) { $v = $strings[[int]$v] }
    $cells[($c.r -replace '\d','')] = $v
  }
  $rows += [PSCustomObject]@{
    codigo   = $cells['E']
    importe  = $cells['I']
    concepto = $cells['K']
    texto    = $cells['L']
  }
}

# Solo asientos con concepto, texto y codigo
$datos = $rows | Where-Object {
  $_.concepto -and $_.concepto -ne 'Concepto' -and $_.texto -and $_.codigo
}
# Periodificaciones: no son cuota mensual de un local
$datos = $datos | Where-Object { $_.texto -notmatch 'GASTOS ANTICIPADOS' }
# Asientos sin direccion: no sirven como referencia por calle
$datos = $datos | Where-Object { $_.texto -notmatch '^\s*CUOTA MES' }

# -- 4. Utilidades de texto --------------------------------------------------

# Quita acentos sin escribir caracteres no ASCII en este script
function Remove-Accents([string]$s) {
  if (-not $s) { return '' }
  $d = $s.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object Text.StringBuilder
  foreach ($ch in $d.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
        [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
  }
  return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

# Ciudades tal y como aparecen en los textos. Se comparan sin acentos.
# 'ALICAN' esta truncado en origen por el limite de 40 caracteres.
$CIUDADES = @{
  'MADRID'                 = 'Madrid'
  'BARCELONA'              = 'Barcelona'
  'VALENCIA'               = 'Valencia'
  'BILBAO'                 = 'Bilbao'
  'SALAMANCA'              = 'Salamanca'
  'ALICANTE'               = 'Alicante'
  'ALICAN'                 = 'Alicante'
  'BENIDORM'               = 'Benidorm'
  'GANDIA'                 = 'Gandia'
  'GETXO'                  = 'Getxo'
  'LEON'                   = 'Leon'
  'DERIO'                  = 'Derio'
  'ANDOAIN'                = 'Andoain'
  'CASTRO URDIALES'        = 'Castro Urdiales'
  'NUEVO BAZTAN'           = 'Nuevo Baztan'
  'JARAIZ DE LA VERA'      = 'Jaraiz de la Vera'
  'SANTA MARIA DEL PARAMO' = 'Santa Maria del Paramo'
  'LAS PEDRONERAS'         = 'Las Pedroneras'
  'PTO.STA.MARIA'          = 'El Puerto de Santa Maria'
}

# Ciudad por codigo de tienda cuando el texto del asiento no la dice (la mayoria
# de los asientos solo llevan la calle). Solo se rellena donde hay evidencia:
#   - CIF: las dos cifras tras la letra son la provincia de registro del
#     arrendador (A28/B78-B88 Madrid, 20 Gipuzkoa, 48 Bizkaia, 29 Malaga,
#     11 Cadiz, 07 Baleares, 46 Valencia, 24 Leon...).
#   - Nombre inequivoco: calles o municipios que solo existen en una ciudad.
# Las tiendas que no aparecen aqui se quedan sin ciudad a proposito: no hay
# dato suficiente y es preferible que no salgan en una busqueda por distrito
# antes que atribuirlas a la ciudad equivocada.
$CIUDAD_MANUAL = @{
  # --- Madrid capital (CIF de Madrid y/o calle inequivoca de Madrid) ---
  '2810' = 'Madrid'  # Andres Mellado      B80 Madrid
  '2811' = 'Madrid'  # General Peron       A28 Madrid
  '2812' = 'Madrid'  # Magallanes          A28 Madrid
  '2817' = 'Madrid'  # Mayor 46            B78 Madrid
  '2820' = 'Madrid'  # Campezo (San Blas)
  '2824' = 'Madrid'  # Bailen
  '2826' = 'Madrid'  # Manuel Becerra
  '2827' = 'Madrid'  # Montera             B78 Madrid
  '2830' = 'Madrid'  # Toledo              B79 Madrid
  '2834' = 'Madrid'  # Ortega y Gasset     B83 Madrid
  '2838' = 'Madrid'  # Hermosilla          A28 Madrid
  '2988' = 'Madrid'  # Jose del Hierro
  '3104' = 'Madrid'  # Calle Leon (Letras)
  '3134' = 'Madrid'  # Castello            B83 Madrid
  '3135' = 'Madrid'  # Cea Bermudez        B83 Madrid
  '3136' = 'Madrid'  # Espiritu Santo      A78 Madrid
  '3137' = 'Madrid'  # Fuencarral          E82 Madrid
  '3138' = 'Madrid'  # Pez                 A79 Madrid
  '3139' = 'Madrid'  # San Bernardo
  '3157' = 'Madrid'  # Jardines (Centro)
  '3166' = 'Madrid'  # Cartagena
  '3201' = 'Madrid'  # Antonio Lopez       A78 Madrid
  '3202' = 'Madrid'  # Arturo Soria        A79 Madrid
  '3230' = 'Madrid'  # Jose Abascal        B80 Madrid
  '3231' = 'Madrid'  # Martinez Campos     B82 Madrid
  '3234' = 'Madrid'  # Nunez de Balboa (barrio de Salamanca)
  '3256' = 'Madrid'  # Gravina (Chueca)
  '3264' = 'Madrid'  # Hortaleza           V28 Madrid
  '3277' = 'Madrid'  # Nunez Morgado
  '3285' = 'Madrid'  # San Bernardino      E87 Madrid
  '3286' = 'Madrid'  # Santa Isabel        B78 Madrid
  # --- Fuera de Madrid (el CIF y/o el nombre lo desmienten) ---
  '2821' = 'Bilbao'                 # estacion de Abando (local, almacen y vending)
  '2845' = 'Nuevo Baztan'           # B82 Madrid, pero es municipio propio de la
                                    # C. de Madrid: no debe mezclarse con la capital
  '2822' = 'Getxo'                  # A48 Bizkaia
  '2836' = 'Palma de Mallorca'      # A07 Baleares (Gremi Tintorers)
  '2837' = 'Palma de Mallorca'      # A07 Baleares (Carrer Arago)
  '2842' = 'Valencia'               # E46 Valencia (Fernando el Catolico)
  '3170' = 'Leon'                   # E24 Leon
  '3180' = 'Jaraiz de la Vera'      # municipio (Caceres)
  '3189' = 'Santa Maria del Paramo' # municipio (Leon)
  '3191' = 'Andoain'                # B20 Gipuzkoa
  '3247' = 'Castro Urdiales'        # municipio (Cantabria)
  '3251' = 'Cadiz'                  # B11 Cadiz (Ana de Viya)
  '3282' = 'Las Pedroneras'         # municipio (Cuenca)
  '3284' = 'Malaga'                 # A29 Malaga (Alameda Principal)
  # Sin asignar por falta de evidencia: 3080 Feria, 3161 Gabriel y Galan,
  # 3239 Padre Huerfanos (arrendador con DNI, nombre presente en varias ciudades).
}

function Get-Tipo([string]$t) {
  $n = Remove-Accents $t
  if ($n -match '(?i)ALMACEN')                    { return 'almacen' }
  if ($n -match '(?i)ESPACIO\s+MAQUINAS\s+VEND')  { return 'vending' }
  return 'local'
}

# Devuelve @{ calle; ciudad } a partir del texto del asiento.
# $codigo se usa para reconocer restos del codigo de tienda cortado por la
# truncacion a 40 caracteres (p.ej. "... (DERIO) 32" con codigo 3205).
function Parse-Direccion([string]$t, [string]$codigo) {
  $truncado = ($t.Length -ge 40)
  $s = $t
  $s = $s -replace '^\s*\d{2}\.\d{4}\s*', ''                  # 09.2016
  # Corta en el codigo de tienda y descarta lo que venga detras: "(3256) JOSE J"
  $s = $s -replace '\(\s*\d{3,4}\s*\).*$', ''
  $s = $s -replace '\(\s*\d{0,4}\s*$', ''                     # "(283" o "(" truncado
  # Palabras de formato/tipo
  $s = $s -replace '(?i)\bC\.\s*EXPRESS?\b', ''
  $s = $s -replace '(?i)\bEXPRESS?\b', ''
  $s = $s -replace '(?i)\bALMAC[E\u00C9]N\b', ''
  $s = $s -replace '(?i)\bESPACIO\s+MAQUINAS\s+VENDING\b', ''
  $s = $s -replace '(?i)\bESTACI[O\u00D3]N\b', ''
  $s = $s.Trim(' ', ',', '-', '.')

  $ciudad = ''

  # Ciudad entre parentesis: "MIKEL DEUNA (DERIO) 32"
  if ($s -match '\(([^)]+)\)') {
    $cand = (Remove-Accents $Matches[1]).Trim().ToUpper()
    if ($CIUDADES.ContainsKey($cand)) {
      $ciudad = $CIUDADES[$cand]
      $s = ($s -replace '\([^)]+\)', ' ')
    }
  }

  # Ciudad delante del prefijo de via: "GANDIA C/GABRIEL MIRO"
  if ($s -match '(?i)^(.+?)\s+(C\/|CALLE|PZA\.?|PLAZA|AVDA\.?|AVENIDA)\s*(.+)$') {
    $cand = (Remove-Accents $Matches[1]).Trim().ToUpper()
    if ($CIUDADES.ContainsKey($cand)) {
      $ciudad = $CIUDADES[$cand]
      $s = $Matches[3]
    }
  }

  # Ciudad al final: "RUZAFA, VALENCIA" / "AGASTIA -MADRID"
  # Se busca sobre el texto sin acentos pero se corta el original: quitar una
  # marca diacritica no cambia la longitud, asi que los indices coinciden.
  foreach ($k in ($CIUDADES.Keys | Sort-Object { $_.Length } -Descending)) {
    $pat = '(?i)[\s,\-]+' + [regex]::Escape($k) + '\s*[,\.]?\s*$'
    $m = [regex]::Match((Remove-Accents $s), $pat)
    if ($m.Success) {
      if (-not $ciudad) { $ciudad = $CIUDADES[$k] }
      $s = $s.Substring(0, $m.Index)
      break
    }
  }

  # Ciudad al principio: "BILBAO-ABANDO" (la estacion de Abando esta en Bilbao)
  foreach ($k in ($CIUDADES.Keys | Sort-Object { $_.Length } -Descending)) {
    $pat = '(?i)^' + [regex]::Escape($k) + '[\s,\-]+'
    $m = [regex]::Match((Remove-Accents $s), $pat)
    if ($m.Success -and $s.Length -gt $m.Length) {
      if (-not $ciudad) { $ciudad = $CIUDADES[$k] }
      $s = $s.Substring($m.Length)
      break
    }
  }

  # Prefijo de via al principio
  $s = $s -replace '(?i)^\s*(C\/|CALLE|PZA\.?|PLAZA|AVDA\.?|AVENIDA)\s*', ''
  $s = $s.Trim(' ', ',', '-', '.', '(', ')')

  # Resto del codigo de tienda cortado por la truncacion: "MIKEL DEUNA ... 32"
  # con codigo 3205. Solo se aplica si el texto venia truncado y las cifras
  # finales son el principio del codigo, para no tocar numeros de portal reales.
  if ($truncado -and $codigo -and $s -match '\s(\d{1,4})$') {
    if ($codigo.StartsWith($Matches[1])) {
      $s = $s -replace '\s\d{1,4}$', ''
    }
  }

  $s = $s.Trim(' ', ',', '-', '.', '(', ')')
  $s = ($s -replace '\s+', ' ').Trim()

  return @{ calle = $s; ciudad = $ciudad }
}

# Clave para agrupar: sin acentos, sin puntuacion y sin numero de portal,
# de forma que "GRAVINA" y "GRAVINA, 13" sean el mismo local.
function Get-CalleKey([string]$calle) {
  $k = (Remove-Accents $calle).ToUpper()
  $k = $k -replace '[^A-Z0-9 ]', ' '
  $k = $k -replace '\s*\b\d+\b\s*$', ''
  return ($k -replace '\s+', ' ').Trim()
}

# -- 5. Agrupar por local real ----------------------------------------------
$parsed = @()
foreach ($d in $datos) {
  $p = Parse-Direccion $d.texto $d.codigo
  # Sin calle pero con ciudad es valido: hay tiendas referenciadas solo por
  # municipio (p.ej. Santa Maria del Paramo), utiles como referencia de zona.
  if (-not $p.calle -and -not $p.ciudad) { continue }
  $parsed += [PSCustomObject]@{
    codigo   = $d.codigo
    tipo     = (Get-Tipo $d.texto)
    calle    = $p.calle
    ciudad   = $p.ciudad
    key      = "$($d.codigo)|$(Get-Tipo $d.texto)|$(Get-CalleKey $p.calle)"
    importe  = [double]$d.importe
    concepto = $d.concepto
  }
}

$locales = @()
foreach ($g in ($parsed | Group-Object key)) {
  $first  = $g.Group | Select-Object -First 1
  # Nombre de calle mas completo de las variantes encontradas
  $calle  = ($g.Group | Sort-Object { $_.calle.Length } -Descending | Select-Object -First 1).calle
  # Ciudad: la del texto del asiento; si no la trae, la tabla manual
  $ciudad = ($g.Group | Where-Object { $_.ciudad } | Select-Object -First 1).ciudad
  if (-not $ciudad -and $CIUDAD_MANUAL.ContainsKey($first.codigo)) {
    $ciudad = $CIUDAD_MANUAL[$first.codigo]
  }
  if (-not $ciudad) { $ciudad = '' }

  $renta  = ($g.Group | Where-Object { $_.concepto -eq 'RENTA ALQUILER' } |
               Measure-Object -Property importe -Sum).Sum
  if (-not $renta) { $renta = 0 }

  $desglose = [ordered]@{}
  foreach ($c in ($g.Group | Where-Object { $_.concepto -ne 'RENTA ALQUILER' } |
                    Group-Object concepto | Sort-Object Name)) {
    $etiqueta = ($c.Name -replace '(?i)^ALQUILER INMUEBLES_', '')
    $desglose[$etiqueta] = [math]::Round((($c.Group | Measure-Object -Property importe -Sum).Sum), 2)
  }
  $extras = 0
  foreach ($v in $desglose.Values) { $extras += $v }

  $locales += [PSCustomObject]@{
    codigo   = $first.codigo
    tipo     = $first.tipo
    calle    = $calle
    ciudad   = $ciudad
    renta    = [math]::Round($renta, 2)
    extras   = [math]::Round($extras, 2)
    total    = [math]::Round($renta + $extras, 2)
    desglose = $desglose
  }
}

$locales = $locales | Sort-Object calle, codigo

# -- 6. Escribir JSON --------------------------------------------------------
$outFile = Join-Path $root "src\tiendas.json"
$json = $locales | ConvertTo-Json -Depth 5
$utf8 = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText($outFile, $json, $utf8)

$nLocal   = @($locales | Where-Object { $_.tipo -eq 'local' }).Count
$nCiudad  = @($locales | Where-Object { $_.ciudad }).Count
$sumTotal = ($locales | Measure-Object -Property total -Sum).Sum
Write-Host ("OK  src\tiendas.json - {0} inmuebles ({1} locales, {2} con ciudad en el fichero), total {3:N2} EUR/mes" -f `
  $locales.Count, $nLocal, $nCiudad, $sumTotal)
