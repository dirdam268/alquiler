# publicar.ps1 - compila la app y la deja cifrada y lista para publicar.
#
# Uso: doble clic derecho > "Ejecutar con PowerShell", o desde una terminal:
#      powershell -ExecutionPolicy Bypass -File publicar.ps1
#
# Pide la contrasena por teclado (no se ve al escribirla y no queda guardada
# en ningun fichero). Es la misma que usas en las otras apps.

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

Write-Host ""
Write-Host "=== Publicar Alquiler ===" -ForegroundColor Cyan
Write-Host ""

# 1) Alquileres reales de las tiendas (solo si esta el xlsx original)
$xlsx = "C:\Users\Usuario\Downloads\Datos economicos tiendas Express.xlsx"
if (Test-Path $xlsx) {
  Write-Host "[1/3] Extrayendo alquileres reales del fichero de tiendas..."
  & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "scripts\build_tiendas.ps1")
} else {
  if (Test-Path (Join-Path $root "src\tiendas.json")) {
    Write-Host "[1/3] Uso el src\tiendas.json ya generado (no encuentro el xlsx original)." -ForegroundColor Yellow
  } else {
    throw "No encuentro ni el xlsx ni src\tiendas.json. Deja el fichero en $xlsx"
  }
}

# 2) App en claro (queda en build\, no se publica)
Write-Host "[2/3] Compilando la app..."
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "scripts\build.ps1")

# 3) Cifrar
Write-Host ""
Write-Host "[3/3] Contrasena para abrir la app (la misma de tus otras apps)."
Write-Host "      No se vera mientras la escribes." -ForegroundColor DarkGray
$sec = Read-Host "      Contrasena" -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
try {
  $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
} finally {
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
}

if ([string]::IsNullOrWhiteSpace($plain)) { throw "Contrasena vacia, no se ha cifrado nada." }

& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "scripts\build-secure.ps1") -Password $plain
$plain = $null
[GC]::Collect()

Write-Host ""
Write-Host "LISTO. app\index.html ya esta cifrado." -ForegroundColor Green
Write-Host "Ahora dile a Claude 'hecho' para que lo suba a GitHub."
Write-Host ""
