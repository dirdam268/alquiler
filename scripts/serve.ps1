# serve.ps1 — servidor local para probar la app (http://localhost:8765)
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\serve.ps1
param([int]$Port = 8765)

$root = Join-Path (Split-Path $PSScriptRoot -Parent) "app"
$mime = @{
    ".html" = "text/html; charset=utf-8"
    ".js"   = "text/javascript; charset=utf-8"
    ".json" = "application/json; charset=utf-8"
    ".webmanifest" = "application/manifest+json; charset=utf-8"
    ".png"  = "image/png"
    ".css"  = "text/css; charset=utf-8"
    ".ico"  = "image/x-icon"
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Sirviendo $root en http://localhost:$Port/  (Ctrl+C para parar)"

while ($listener.IsListening) {
    $ctx  = $listener.GetContext()
    $path = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath.TrimStart("/"))
    if ([string]::IsNullOrEmpty($path)) { $path = "index.html" }
    $file = Join-Path $root $path
    try {
        if ((Test-Path $file -PathType Leaf) -and ([IO.Path]::GetFullPath($file)).StartsWith([IO.Path]::GetFullPath($root))) {
            $bytes = [IO.File]::ReadAllBytes($file)
            $ext = [IO.Path]::GetExtension($file).ToLower()
            if ($mime.ContainsKey($ext)) { $ctx.Response.ContentType = $mime[$ext] }
            $ctx.Response.ContentLength64 = $bytes.Length
            $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $ctx.Response.StatusCode = 404
        }
    } catch { $ctx.Response.StatusCode = 500 }
    $ctx.Response.OutputStream.Close()
}
