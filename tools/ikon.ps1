# app.ico üretir: mavi zemin, beyaz zarf, yeşil aşağı ok. Bir kez çalıştırılır, çıktı depoda durur.
Add-Type -AssemblyName System.Drawing
$out = Join-Path (Split-Path -Parent $PSScriptRoot) 'app.ico'

function New-IconPng([int]$s) {
    $bmp = New-Object System.Drawing.Bitmap $s, $s
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $r = [single]($s * 0.18)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $w = [single]($s - 1)
    $path.AddArc(0, 0, $r * 2, $r * 2, 180, 90)
    $path.AddArc($w - $r * 2, 0, $r * 2, $r * 2, 270, 90)
    $path.AddArc($w - $r * 2, $w - $r * 2, $r * 2, $r * 2, 0, 90)
    $path.AddArc(0, $w - $r * 2, $r * 2, $r * 2, 90, 90)
    $path.CloseFigure()
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 28, 94, 168))), $path)
    # zarf
    $ex = $s * 0.16; $ey = $s * 0.22; $ew = $s * 0.68; $eh = $s * 0.44
    $white = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $g.FillRectangle($white, [single]$ex, [single]$ey, [single]$ew, [single]$eh)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 28, 94, 168)), ([single][Math]::Max(1, $s * 0.035))
    $g.DrawLines($pen, [System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF ([single]$ex), ([single]$ey)),
        (New-Object System.Drawing.PointF ([single]($ex + $ew / 2)), ([single]($ey + $eh * 0.58))),
        (New-Object System.Drawing.PointF ([single]($ex + $ew)), ([single]$ey))))
    # ok
    $green = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 46, 170, 90))
    $cx = $s * 0.5; $top = $s * 0.56; $bw = $s * 0.12
    $g.FillRectangle($green, [single]($cx - $bw / 2), [single]$top, [single]$bw, [single]($s * 0.16))
    $g.FillPolygon($green, [System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF ([single]($cx - $s * 0.17)), ([single]($s * 0.70))),
        (New-Object System.Drawing.PointF ([single]($cx + $s * 0.17)), ([single]($s * 0.70))),
        (New-Object System.Drawing.PointF ([single]$cx), ([single]($s * 0.90)))))
    $g.Dispose()
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    , $ms.ToArray()
}

$sizes = 256, 48, 32, 16
$pngs = foreach ($s in $sizes) { , (New-IconPng $s) }
$fs = [System.IO.File]::Create($out)
$bw = New-Object System.IO.BinaryWriter $fs
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $s = $sizes[$i]; $d = $pngs[$i]
    $bw.Write([byte]($(if ($s -ge 256) { 0 } else { $s })))
    $bw.Write([byte]($(if ($s -ge 256) { 0 } else { $s })))
    $bw.Write([byte]0); $bw.Write([byte]0)
    $bw.Write([uint16]1); $bw.Write([uint16]32)
    $bw.Write([uint32]$d.Length); $bw.Write([uint32]$offset)
    $offset += $d.Length
}
foreach ($d in $pngs) { $bw.Write($d) }
$bw.Dispose()
"Yazildi: $out"
