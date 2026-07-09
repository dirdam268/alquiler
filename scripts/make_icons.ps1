# make_icons.ps1 — genera los iconos PNG de la PWA en pwa\
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\make_icons.ps1

Add-Type -AssemblyName System.Drawing

$root = Split-Path $PSScriptRoot -Parent
$out  = Join-Path $root "pwa"

function New-Icon([int]$size, [string]$file, [double]$glyphScale) {
    $bmp = New-Object System.Drawing.Bitmap($size, $size)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias

    # Fondo degradado oscuro (mismo tono que la app)
    $rect  = New-Object System.Drawing.Rectangle(0, 0, $size, $size)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect,
        [System.Drawing.Color]::FromArgb(29, 39, 49),
        [System.Drawing.Color]::FromArgb(16, 22, 28),
        [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal)
    $g.FillRectangle($brush, $rect)

    # Letra "A" en ámbar
    $fontSize = [float]($size * $glyphScale)
    $font  = New-Object System.Drawing.Font("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $amber = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(244, 184, 64))
    $fmt   = New-Object System.Drawing.StringFormat
    $fmt.Alignment     = [System.Drawing.StringAlignment]::Center
    $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rectF = New-Object System.Drawing.RectangleF(0, [float]($size * -0.02), $size, $size)
    $g.DrawString("A", $font, $amber, $rectF, $fmt)

    # Subrayado ámbar (barra de "precio")
    $barW = $size * 0.34; $barH = [Math]::Max(2, $size * 0.045)
    $g.FillRectangle($amber, [float](($size - $barW) / 2), [float]($size * 0.80), [float]$barW, [float]$barH)

    $g.Dispose()
    $bmp.Save((Join-Path $out $file), [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host "OK  $file  ($size x $size)"
}

New-Icon 512 "icon-512.png"          0.62
New-Icon 512 "icon-512-maskable.png" 0.44   # glifo más pequeño: zona segura maskable
New-Icon 192 "icon-192.png"          0.62
New-Icon 180 "icon-180.png"          0.62   # apple-touch-icon
