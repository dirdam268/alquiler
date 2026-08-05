# build-secure.ps1 - cifra la app y genera el app\index.html que se publica.
#
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\build-secure.ps1 -Password '<clave>'
#
# Flujo completo:
#   1) scripts\build_tiendas.ps1   -> src\tiendas.json   (datos internos, gitignored)
#   2) scripts\build.ps1           -> build\index.html   (app en claro, gitignored)
#   3) scripts\build-secure.ps1    -> app\index.html     (cifrado, este si se publica)
#
# Mismo esquema que las otras apps del usuario (AES-256-CBC + PBKDF2-SHA256,
# descifrado en el navegador con Web Crypto). La contrasena NO se guarda en
# ningun fichero: se pasa por parametro en cada build.

param(
  [Parameter(Mandatory=$true)][string]$Password,
  [string]$Src = "build\index.html",
  [string]$Out = "app\index.html"
)
$ErrorActionPreference = 'Stop'
$root    = Split-Path $PSScriptRoot -Parent
$srcPath = Join-Path $root $Src
$outPath = Join-Path $root $Out

if (-not (Test-Path $srcPath)) {
  throw "No existe $Src. Ejecuta antes scripts\build.ps1"
}

# 1) Leer la app en claro y anteponer un sello para verificar el descifrado
$plainBytes = [System.IO.File]::ReadAllBytes($srcPath)
$sentinel = [System.Text.Encoding]::UTF8.GetBytes("ALQ_OK|")
$data = New-Object byte[] ($sentinel.Length + $plainBytes.Length)
[Array]::Copy($sentinel, 0, $data, 0, $sentinel.Length)
[Array]::Copy($plainBytes, 0, $data, $sentinel.Length, $plainBytes.Length)

# 2) Derivar clave (PBKDF2-SHA256) y cifrar (AES-256-CBC)
$rng  = [System.Security.Cryptography.RandomNumberGenerator]::Create()
$salt = New-Object byte[] 16; $rng.GetBytes($salt)
$iv   = New-Object byte[] 16; $rng.GetBytes($iv)
$iter = 100000
$kdf  = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $salt, $iter, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
$key  = $kdf.GetBytes(32)
$aes  = [System.Security.Cryptography.Aes]::Create()
$aes.KeySize = 256; $aes.Mode = 'CBC'; $aes.Padding = 'PKCS7'; $aes.Key = $key; $aes.IV = $iv
$ct   = $aes.CreateEncryptor().TransformFinalBlock($data, 0, $data.Length)

$payload = @{
  v = 1; iter = $iter
  salt = [Convert]::ToBase64String($salt)
  iv   = [Convert]::ToBase64String($iv)
  ct   = [Convert]::ToBase64String($ct)
} | ConvertTo-Json -Compress

# 3) Pantalla de acceso (descifra en el navegador con Web Crypto)
$gate = @'
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Alquiler</title>
<meta name="description" content="Compara la renta de un local comercial con el mercado y calcula la tasa de esfuerzo."/>
<meta name="theme-color" content="#eef1f6"/>
<link rel="manifest" href="manifest.webmanifest"/>
<link rel="icon" type="image/png" href="icon-192.png"/>
<link rel="apple-touch-icon" href="icon-180.png"/>
<meta name="apple-mobile-web-app-capable" content="yes"/>
<meta name="apple-mobile-web-app-status-bar-style" content="default"/>
<meta name="apple-mobile-web-app-title" content="Alquiler"/>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: "Inter", system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
    background: radial-gradient(1200px 600px at 80% -10%, #dfe8f2 0%, transparent 60%), #eef1f6;
    color: #182028; min-height: 100vh;
    display: flex; align-items: center; justify-content: center; padding: 20px;
    -webkit-font-smoothing: antialiased;
  }
  .gate {
    background: #fff; border: 1px solid #d8e0e9; border-radius: 18px;
    box-shadow: 0 10px 40px rgba(20,30,40,.10);
    padding: 32px 28px; width: 100%; max-width: 380px; text-align: center;
  }
  .lock { font-size: 40px; margin-bottom: 14px; }
  .logo { font-weight: 800; font-size: 24px; letter-spacing: -.5px; margin-bottom: 4px; }
  .logo span { color: #9c6a07; }
  .sub { font-size: 13px; color: #64707c; margin-bottom: 22px; }
  input {
    width: 100%; border: 1.5px solid #d8e0e9; border-radius: 10px;
    padding: 12px 14px; font-size: 16px; outline: none;
    color: #182028; background: #eef2f7; margin-bottom: 12px; font-family: inherit;
  }
  input:focus { border-color: #f2ac2c; background: #fff; }
  button {
    width: 100%; background: #f2ac2c; border: none; color: #3b2a06;
    border-radius: 10px; padding: 13px; font-weight: 740; font-size: 15px;
    cursor: pointer; font-family: inherit;
  }
  button:disabled { opacity: .6; cursor: default; }
  .err { color: #d6455a; font-size: 13px; font-weight: 600; margin-top: 12px; min-height: 18px; }
  .recover { margin-top: 18px; font-size: 12px; color: #64707c; }
  .recover a { color: #9c6a07; font-weight: 600; text-decoration: none; }
</style>
</head>
<body>
<div class="gate">
  <div class="lock">&#128274;</div>
  <div class="logo">Alquiler<span>.</span></div>
  <div class="sub">Acceso privado &middot; datos internos</div>
  <input id="pw" type="password" placeholder="Contrase&ntilde;a" autocomplete="current-password" autofocus>
  <button id="go">Entrar</button>
  <div class="err" id="err"></div>
  <div class="recover">&iquest;Olvidaste la contrase&ntilde;a?<br><a href="mailto:bordetass@gmail.com?subject=Acceso%20Alquiler">Escribe a bordetass@gmail.com</a></div>
</div>
<script>
const PAYLOAD = __PAYLOAD__;
const b64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));

async function decryptApp(password) {
  const salt = b64(PAYLOAD.salt), iv = b64(PAYLOAD.iv), ct = b64(PAYLOAD.ct);
  const km = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt, iterations: PAYLOAD.iter, hash: 'SHA-256' },
    km, { name: 'AES-CBC', length: 256 }, false, ['decrypt']
  );
  const buf = await crypto.subtle.decrypt({ name: 'AES-CBC', iv }, key, ct);
  const text = new TextDecoder().decode(buf);
  if (!text.startsWith('ALQ_OK|')) throw new Error('sello');
  return text.slice(7);
}

async function enter() {
  const btn = document.getElementById('go');
  const err = document.getElementById('err');
  const pw  = document.getElementById('pw').value;
  // Textos con escapes \uXXXX a proposito: asi este .ps1 es ASCII puro y no
  // depende de que PowerShell lo lea como UTF-8 (sin BOM lo leeria como ANSI
  // y saldria "ContraseAa incorrecta").
  if (!pw) { err.textContent = 'Introduce la contrase\u00f1a.'; return; }
  btn.disabled = true; err.textContent = 'Descifrando\u2026';
  try {
    const html = await decryptApp(pw);
    // Recordar en este dispositivo: solo se pide la primera vez
    try { localStorage.setItem('alq_pw', pw); } catch(_) {}
    document.open(); document.write(html); document.close();
  } catch (e) {
    try { localStorage.removeItem('alq_pw'); } catch(_) {}
    err.textContent = 'Contrase\u00f1a incorrecta.';
    btn.disabled = false;
  }
}

document.getElementById('go').addEventListener('click', enter);
document.getElementById('pw').addEventListener('keydown', e => { if (e.key === 'Enter') enter(); });

// Si ya se valido antes en este dispositivo, entrar directo sin preguntar
const saved = (() => { try { return localStorage.getItem('alq_pw'); } catch(_) { return null; } })();
if (saved) {
  document.getElementById('pw').value = saved;
  enter();
}

if ('serviceWorker' in navigator &&
    (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1')) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(()=>{}));
}
</script>
</body>
</html>
'@

$gate = $gate.Replace('__PAYLOAD__', $payload)
[System.IO.File]::WriteAllText($outPath, $gate, (New-Object System.Text.UTF8Encoding($false)))
Write-Host ("OK  {0} generado y cifrado ({1} KB). Este es el fichero que se publica." -f $Out, [math]::Round($ct.Length/1024))
