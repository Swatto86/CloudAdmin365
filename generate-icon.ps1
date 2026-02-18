<#
.SYNOPSIS
    Generates CloudAdmin365.ico — a cog/gear with "365" text on a blue background.

.DESCRIPTION
    Uses System.Drawing to produce the same icon design as IconGenerator.cs,
    then writes a proper .ico file (with 16, 32, and 48 px sizes) that can be
    embedded in the EXE via <ApplicationIcon> in the .csproj.

    Run this any time the icon design changes, then commit the resulting .ico.

.EXAMPLE
    .\generate-icon.ps1
#>

Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

# ── Colours ────────────────────────────────────────────────────────────────
$BgBlue    = [System.Drawing.Color]::FromArgb(0, 120, 212)
$CogWhite  = [System.Drawing.Color]::White
$TextBlue  = [System.Drawing.Color]::FromArgb(0, 120, 212)
$ShadowCol = [System.Drawing.Color]::FromArgb(40, 0, 0, 0)

function New-GearPolygon {
    param([float]$cx, [float]$cy, [float]$outerR, [float]$innerR, [int]$teeth)

    $count = $teeth * 2
    [System.Drawing.PointF[]]$pts = [System.Drawing.PointF[]]::new($count)
    $angleStep = [Math]::PI * 2.0 / $count

    for ($i = 0; $i -lt $count; $i++) {
        $angle = $i * $angleStep - [Math]::PI / 2.0
        $r = if ($i % 2 -eq 0) { $outerR } else { $innerR }
        $pts[$i] = [System.Drawing.PointF]::new(
            [float]($cx + [Math]::Cos($angle) * $r),
            [float]($cy + [Math]::Sin($angle) * $r))
    }
    return ,$pts
}

function New-RoundedRectPath {
    param([float]$x, [float]$y, [float]$w, [float]$h, [float]$r)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $path.AddArc($x,          $y,          $d, $d, 180, 90)
    $path.AddArc($x + $w - $d, $y,         $d, $d, 270, 90)
    $path.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0,   90)
    $path.AddArc($x,          $y + $h - $d, $d, $d, 90,  90)
    $path.CloseFigure()
    return $path
}

function New-IconBitmap {
    param([int]$size)

    $bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

    # Background
    $bgBrush = New-Object System.Drawing.SolidBrush $BgBlue
    $rrPath = New-RoundedRectPath 0 0 $size $size ($size * 0.22)
    $g.FillPath($bgBrush, $rrPath)

    $cx = $size / 2.0
    $cy = $size / 2.0
    $outerR = $size * 0.42
    $innerR = $size * 0.30
    $teeth = 8

    # Shadow
    $shadowBrush = New-Object System.Drawing.SolidBrush $ShadowCol
    $shadowPts = New-GearPolygon ($cx + 0.5) ($cy + 0.5) $outerR $innerR $teeth
    $shadowPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $shadowPath.AddPolygon($shadowPts)
    $g.FillPath($shadowBrush, $shadowPath)

    # Cog
    $cogBrush = New-Object System.Drawing.SolidBrush $CogWhite
    $cogPts = New-GearPolygon $cx $cy $outerR $innerR $teeth
    $cogPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $cogPath.AddPolygon($cogPts)
    $g.FillPath($cogBrush, $cogPath)

    # "365" text
    $textBrush = New-Object System.Drawing.SolidBrush $TextBlue
    $fontSize = [Math]::Max(5, $size * 0.27)
    $font = New-Object System.Drawing.Font "Segoe UI", $fontSize, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
    $textSize = $g.MeasureString("365", $font)
    $tx = $cx - $textSize.Width / 2
    $ty = $cy - $textSize.Height / 2
    $g.DrawString("365", $font, $textBrush, $tx, $ty)

    $g.Dispose()
    $bgBrush.Dispose()
    $shadowBrush.Dispose()
    $cogBrush.Dispose()
    $textBrush.Dispose()
    $font.Dispose()
    $rrPath.Dispose()
    $shadowPath.Dispose()
    $cogPath.Dispose()

    return $bmp
}

# ── Generate bitmaps at standard icon sizes ─────────────────────────────────
$sizes = @(16, 32, 48)
$bitmaps = @()
foreach ($sz in $sizes) {
    $bitmaps += New-IconBitmap $sz
    Write-Host "  Generated ${sz}x${sz} bitmap" -ForegroundColor Green
}

# ── Write .ico file ─────────────────────────────────────────────────────────
# ICO format: 6-byte header + (16-byte directory entry per image) + PNG data per image.
$icoPath = Join-Path $PSScriptRoot "CloudAdmin365.ico"

$ms = New-Object System.IO.MemoryStream
$bw = New-Object System.IO.BinaryWriter $ms

# Header: reserved (2) + type (2, 1=icon) + count (2)
$bw.Write([UInt16]0)
$bw.Write([UInt16]1)
$bw.Write([UInt16]$bitmaps.Count)

# Prepare PNG data for each bitmap
$pngDataList = @()
foreach ($bmp in $bitmaps) {
    $pngMs = New-Object System.IO.MemoryStream
    $bmp.Save($pngMs, [System.Drawing.Imaging.ImageFormat]::Png)
    $pngDataList += ,($pngMs.ToArray())
    $pngMs.Dispose()
}

# Calculate data start offset: header(6) + entries(16 * count)
$dataOffset = 6 + 16 * $bitmaps.Count

# Write directory entries
for ($i = 0; $i -lt $bitmaps.Count; $i++) {
    $sz  = $sizes[$i]
    $png = $pngDataList[$i]
    $w   = if ($sz -ge 256) { 0 } else { $sz }
    $h   = if ($sz -ge 256) { 0 } else { $sz }

    $bw.Write([byte]$w)          # width
    $bw.Write([byte]$h)          # height
    $bw.Write([byte]0)           # colour palette count
    $bw.Write([byte]0)           # reserved
    $bw.Write([UInt16]1)         # colour planes
    $bw.Write([UInt16]32)        # bits per pixel
    $bw.Write([UInt32]$png.Length) # image data size
    $bw.Write([UInt32]$dataOffset) # offset to image data

    $dataOffset += $png.Length
}

# Write image data
foreach ($png in $pngDataList) {
    $bw.Write($png)
}

$bw.Flush()
[System.IO.File]::WriteAllBytes($icoPath, $ms.ToArray())

$bw.Dispose()
$ms.Dispose()
foreach ($bmp in $bitmaps) { $bmp.Dispose() }

Write-Host ""
Write-Host "  Icon written to: $icoPath" -ForegroundColor Cyan
Write-Host "  Sizes: $($sizes -join ', ') px" -ForegroundColor Cyan
