# Draws the Office Tabs Lite icon (a sheet with three tabs) and packs it into a
# multi-size .ico (16..256 px, PNG entries). Usage: make-icon.ps1 <out.ico> [preview.png]
param([Parameter(Mandatory)] [string] $OutIco, [string] $PreviewPng)

Add-Type -AssemblyName System.Drawing

function New-RoundRect([single]$x, [single]$y, [single]$w, [single]$h, [single]$r) {
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = 2 * $r
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

# Draws on a 256x256 design grid, scaled to $size.
function Draw-Icon([int]$size) {
    $bmp = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.PixelOffsetMode = 'HighQuality'
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.ScaleTransform($size / 256.0, $size / 256.0)

    $Rgb = { param($r, $gr, $b, $a = 255) [System.Drawing.Color]::FromArgb($a, $r, $gr, $b) }

    # glossy green disc: bright lime toward the top-left, deep green at the rim
    $disc = New-Object System.Drawing.Drawing2D.GraphicsPath
    $disc.AddEllipse(8, 8, 240, 240)
    $pg = New-Object System.Drawing.Drawing2D.PathGradientBrush $disc
    $pg.CenterPoint = New-Object System.Drawing.PointF 104, 84
    $pg.CenterColor = & $Rgb 170 235 60
    $pg.SurroundColors = @(& $Rgb 34 118 22)
    $g.FillPath($pg, $disc)
    $g.DrawEllipse((New-Object System.Drawing.Pen (& $Rgb 28 96 18), 4), 10, 10, 236, 236)

    # three file cards fanned behind (blue, green, orange) with a folded corner and text lines
    $cards = @(
        @{ X = 50;  Y = 66; A = -14; Top = (& $Rgb 70 140 240); Bottom = (& $Rgb 25 85 200) },
        @{ X = 93;  Y = 52; A = 0;   Top = (& $Rgb 60 190 100); Bottom = (& $Rgb 20 130 60) },
        @{ X = 136; Y = 66; A = 14;  Top = (& $Rgb 250 130 60); Bottom = (& $Rgb 215 75 30) }
    )
    foreach ($cd in $cards) {
        $state = $g.Save()
        $g.TranslateTransform($cd.X + 35, $cd.Y + 45)
        $g.RotateTransform($cd.A)
        $card = New-RoundRect -35 -45 70 90 9
        $fill = New-Object System.Drawing.Drawing2D.LinearGradientBrush (New-Object System.Drawing.PointF 0, -45), (New-Object System.Drawing.PointF 0, 45), $cd.Top, $cd.Bottom
        $g.FillPath($fill, $card)
        $fold = New-Object System.Drawing.Drawing2D.GraphicsPath
        $fold.AddPolygon(@((New-Object System.Drawing.PointF 15, -45), (New-Object System.Drawing.PointF 35, -25), (New-Object System.Drawing.PointF 15, -25)))
        $g.FillPath((New-Object System.Drawing.SolidBrush (& $Rgb 255 255 255 110)), $fold)
        foreach ($ly in -18, -6, 6) {
            $g.FillPath((New-Object System.Drawing.SolidBrush (& $Rgb 255 255 255 200)), (New-RoundRect -22 $ly 40 6 3))
        }
        $g.Restore($state)
    }

    # white tab in front: a tab on the top-left joined to a wide body
    $front = New-Object System.Drawing.Drawing2D.GraphicsPath
    $front.AddPath((New-RoundRect 40 112 84 40 10), $false)
    $front.AddPath((New-RoundRect 40 132 176 76 14), $false)
    $front.FillMode = 'Winding'
    $shadow = New-Object System.Drawing.SolidBrush (& $Rgb 0 50 0 60)
    $state = $g.Save(); $g.TranslateTransform(0, 5); $g.FillPath($shadow, $front); $g.Restore($state)
    $frontFill = New-Object System.Drawing.Drawing2D.LinearGradientBrush (New-Object System.Drawing.PointF 0, 112), (New-Object System.Drawing.PointF 0, 208), ([System.Drawing.Color]::White), (& $Rgb 226 232 240)
    $g.FillPath($frontFill, $front)

    # window dots on the tab and two text lines on the body
    $dots = @((& $Rgb 232 72 60), (& $Rgb 246 190 40), (& $Rgb 60 182 72))
    for ($i = 0; $i -lt 3; $i++) { $g.FillEllipse((New-Object System.Drawing.SolidBrush $dots[$i]), 54 + $i * 18, 121, 12, 12) }
    $text = New-Object System.Drawing.SolidBrush (& $Rgb 70 85 110)
    $g.FillPath($text, (New-RoundRect 62 158 132 12 6))
    $g.FillPath($text, (New-RoundRect 62 180 92 12 6))

    # gloss highlight across the top of the disc
    $gloss = New-Object System.Drawing.Drawing2D.GraphicsPath
    $gloss.AddEllipse(46, 18, 164, 70)
    $glossFill = New-Object System.Drawing.Drawing2D.LinearGradientBrush (New-Object System.Drawing.PointF 0, 18), (New-Object System.Drawing.PointF 0, 88), (& $Rgb 255 255 255 150), (& $Rgb 255 255 255 0)
    $g.FillPath($glossFill, $gloss)

    $g.Dispose()
    return $bmp
}

# Sizes below 256 are stored as classic 32-bit DIBs (PNG entries at small sizes are not
# understood by GDI+ and some older shell code); 256 is stored as PNG.
function Get-DibBytes([System.Drawing.Bitmap]$bmp) {
    $s = $bmp.Width
    $rect = New-Object System.Drawing.Rectangle 0, 0, $s, $s
    $data = $bmp.LockBits($rect, 'ReadOnly', [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $pixels = New-Object byte[] ($s * $s * 4)
    [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $pixels, 0, $pixels.Length)
    $bmp.UnlockBits($data)
    $maskStride = [int]([math]::Ceiling($s / 32.0) * 4)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter $ms
    # BITMAPINFOHEADER: height is doubled to cover the (unused, all-zero) AND mask
    $bw.Write([int]40); $bw.Write([int]$s); $bw.Write([int]($s * 2)); $bw.Write([uint16]1); $bw.Write([uint16]32)
    $bw.Write([int]0); $bw.Write([int]($s * $s * 4 + $maskStride * $s)); $bw.Write([int]0); $bw.Write([int]0)
    $bw.Write([int]0); $bw.Write([int]0)
    for ($row = $s - 1; $row -ge 0; $row--) { $bw.Write($pixels, $row * $s * 4, $s * 4) }   # bottom-up BGRA
    $bw.Write((New-Object byte[] ($maskStride * $s)))
    $bw.Flush()
    return , $ms.ToArray()
}

$sizes = 16, 24, 32, 48, 64, 128, 256
$pngs = @()
foreach ($s in $sizes) {
    $bmp = Draw-Icon $s
    if ($s -ge 256) {
        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        if ($PreviewPng) { $bmp.Save($PreviewPng, [System.Drawing.Imaging.ImageFormat]::Png) }
        $pngs += , $ms.ToArray()
    } else {
        $pngs += , (Get-DibBytes $bmp)
    }
    $bmp.Dispose()
}

# ICONDIR + ICONDIRENTRYs + image payloads
$out = New-Object System.IO.MemoryStream
$w = New-Object System.IO.BinaryWriter $out
$w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $dim = if ($sizes[$i] -ge 256) { 0 } else { $sizes[$i] }
    $w.Write([byte]$dim); $w.Write([byte]$dim); $w.Write([byte]0); $w.Write([byte]0)
    $w.Write([uint16]1); $w.Write([uint16]32)
    $w.Write([uint32]$pngs[$i].Length); $w.Write([uint32]$offset)
    $offset += $pngs[$i].Length
}
foreach ($p in $pngs) { $w.Write($p) }
$w.Flush()
[System.IO.File]::WriteAllBytes($OutIco, $out.ToArray())
"wrote $OutIco ($($out.Length) bytes, sizes: $($sizes -join ', '))"
