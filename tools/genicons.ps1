# Flat icon family for the Sky/Dense skins. Everything is drawn as vectors in a 256x256
# design space and rendered once per target pixel size, so 16px title-bar icons stay sharp
# instead of being a shrunk 256px picture. Output: icons/*.ico (16/20/24/32/40/48/64/128/256)
# plus app.ico, and two contact sheets under tools/restyle-shots for eyeball checking.
# Keep this file ASCII: Windows PowerShell 5.1 mis-decodes UTF-8 without a BOM.
param(
    [string]$ProjectDir = 'D:\work\wxd.files.winforms\wxd.files.winforms',
    [string]$SheetDir = 'D:\work\wxd.files.winforms\tools\restyle-shots'
)

Add-Type -AssemblyName System.Drawing

$Sizes = @(16, 20, 24, 32, 40, 48, 64, 128, 256)
$BadgeTop = [System.Drawing.ColorTranslator]::FromHtml('#3D9BEA')
$BadgeBottom = [System.Drawing.ColorTranslator]::FromHtml('#13589E')
$White = [System.Drawing.Color]::White
$Light = [System.Drawing.ColorTranslator]::FromHtml('#BFE0FF')

function New-Pen([System.Drawing.Color]$c, [single]$w) {
    $p = New-Object System.Drawing.Pen($c, $w)
    $p.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $p.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $p.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    return $p
}

function New-RR([single]$x, [single]$y, [single]$w, [single]$h, [single]$r) {
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $p.AddArc($x, $y, $r * 2, $r * 2, 180, 90)
    $p.AddArc($x + $w - $r * 2, $y, $r * 2, $r * 2, 270, 90)
    $p.AddArc($x + $w - $r * 2, $y + $h - $r * 2, $r * 2, $r * 2, 0, 90)
    $p.AddArc($x, $y + $h - $r * 2, $r * 2, $r * 2, 90, 90)
    $p.CloseFigure()
    return $p
}

function Add-RR([System.Drawing.Drawing2D.GraphicsPath]$target, [single]$x, [single]$y, [single]$w, [single]$h, [single]$r) {
    $p = New-RR $x $y $w $h $r
    $target.AddPath($p, $false)
    $p.Dispose()
}

function Fill-Path([System.Drawing.Graphics]$g, [System.Drawing.Drawing2D.GraphicsPath]$p, [System.Drawing.Color]$c) {
    $b = New-Object System.Drawing.SolidBrush($c)
    $g.FillPath($b, $p)
    $b.Dispose()
}

function Draw-Badge([System.Drawing.Graphics]$g) {
    $path = New-RR 16 16 224 224 54
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.RectangleF(16, 16, 224, 224)), $BadgeTop, $BadgeBottom, 90)
    $g.FillPath($brush, $path)
    $brush.Dispose()
    $path.Dispose()
}

# Shared table silhouette: header band, column dividers, optional highlighted cell.
function Draw-Table([System.Drawing.Graphics]$g, [single]$x, [single]$y, [single]$w, [single]$h, [bool]$hl) {
    $frame = New-RR $x $y $w $h 16
    $pen = New-Pen $White 14
    $g.DrawPath($pen, $frame)
    $pen.Dispose()

    $in = 7.0
    $headH = $h * 0.26
    $band = New-RR ($x + $in) ($y + $in) ($w - $in * 2) ($headH - $in) 9
    Fill-Path $g $band $White
    $band.Dispose()

    $bodyTop = $y + $headH
    $bodyBottom = $y + $h - $in
    $vpen = New-Pen $White 12
    $g.DrawLine($vpen, ($x + $w / 3.0), $bodyTop, ($x + $w / 3.0), $bodyBottom)
    $g.DrawLine($vpen, ($x + $w * 2.0 / 3.0), $bodyTop, ($x + $w * 2.0 / 3.0), $bodyBottom)
    $rowY = $bodyTop + ($bodyBottom - $bodyTop) / 2.0
    $g.DrawLine($vpen, ($x + $in), $rowY, ($x + $w - $in), $rowY)
    $vpen.Dispose()

    if ($hl) {
        $cell = New-RR ($x + $w / 3.0) ($bodyTop + 4) ($w / 3.0) ($rowY - $bodyTop - 4) 6
        Fill-Path $g $cell $Light
        $cell.Dispose()
    }
    $frame.Dispose()
}

# Folder silhouette: a light-blue back plate with the classic top-left tab, then a white
# front flap over its lower two thirds. Filled rather than outlined so the shape still reads
# at 16px title-bar size, where a two-tone flap keeps the tab from blurring into the badge.
function Draw-Folder([System.Drawing.Graphics]$g) {
    $tab = New-RR 48 72 78 46 12
    Fill-Path $g $tab $Light
    $tab.Dispose()
    $back = New-RR 48 88 160 110 14
    Fill-Path $g $back $Light
    $back.Dispose()
    $front = New-RR 42 120 172 78 14
    Fill-Path $g $front $White
    $front.Dispose()
}

function Draw-Glyph([System.Drawing.Graphics]$g, [string]$name) {
    switch ($name) {
        'app' {
            Draw-Folder $g
        }
        'reader' {
            Draw-Table $g 48 54 126 116 $false
            $pen = New-Pen $White 15
            $g.DrawEllipse($pen, 148, 140, 56, 56)
            $g.DrawLine($pen, 192, 184, 214, 206)
            $pen.Dispose()
        }
        'logs' {
            $pen = New-Pen $White 17
            $l1 = New-Pen $Light 17
            $b = New-Object System.Drawing.SolidBrush($White)
            $g.FillEllipse($b, 56, 82, 24, 24)
            $g.FillEllipse($b, 56, 122, 24, 24)
            $g.FillEllipse($b, 56, 162, 24, 24)
            $b.Dispose()
            $g.DrawLine($pen, 98, 94, 200, 94)
            $g.DrawLine($l1, 98, 134, 176, 134)
            $g.DrawLine($pen, 98, 174, 190, 174)
            $pen.Dispose()
            $l1.Dispose()
        }
        'notes' {
            # Page with a cut corner; the text lines are holes so they show the badge colour.
            $page = New-Object System.Drawing.Drawing2D.GraphicsPath
            $page.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate
            $pts = @(
                (New-Object System.Drawing.PointF(74, 58)),
                (New-Object System.Drawing.PointF(152, 58)),
                (New-Object System.Drawing.PointF(188, 94)),
                (New-Object System.Drawing.PointF(188, 198)),
                (New-Object System.Drawing.PointF(74, 198))
            )
            $page.AddPolygon($pts)
            $page.AddRectangle((New-Object System.Drawing.RectangleF(96, 112, 66, 15)))
            $page.AddRectangle((New-Object System.Drawing.RectangleF(96, 142, 66, 15)))
            $page.AddRectangle((New-Object System.Drawing.RectangleF(96, 172, 44, 15)))
            Fill-Path $g $page $White
            $page.Dispose()
            $fold = New-Object System.Drawing.Drawing2D.GraphicsPath
            $fold.AddPolygon(@(
                    (New-Object System.Drawing.PointF(152, 66)),
                    (New-Object System.Drawing.PointF(152, 100)),
                    (New-Object System.Drawing.PointF(180, 100))
                ))
            $fb = New-Object System.Drawing.SolidBrush($Light)
            $g.FillPath($fb, $fold)
            $fb.Dispose()
            $fold.Dispose()
        }
        'vcmp' {
            $pen = New-Pen $White 15
            $g.DrawLine($pen, 100, 91, 100, 169)
            $g.DrawBezier($pen, 100, 120, 130, 120, 158, 128, 158, 158)
            $pen.Dispose()
            $b = New-Object System.Drawing.SolidBrush($White)
            $lb = New-Object System.Drawing.SolidBrush($Light)
            $g.FillEllipse($b, 83, 57, 34, 34)
            $g.FillEllipse($b, 83, 169, 34, 34)
            $g.FillEllipse($lb, 141, 141, 34, 34)
            $b.Dispose()
            $lb.Dispose()
        }
        'cmdmgr' {
            $frame = New-RR 52 62 152 132 18
            $pen = New-Pen $White 14
            $g.DrawPath($pen, $frame)
            $pen.Dispose()
            $p2 = New-Pen $White 16
            $g.DrawLines($p2, @(
                    (New-Object System.Drawing.PointF(82, 106)),
                    (New-Object System.Drawing.PointF(110, 128)),
                    (New-Object System.Drawing.PointF(82, 150))
                ))
            $g.DrawLine($p2, 128, 150, 172, 150)
            $p2.Dispose()
        }
        'dup' {
            $back = New-RR 56 56 108 108 18
            $pen = New-Pen $White 14
            $g.DrawPath($pen, $back)
            $pen.Dispose()
            $front = New-RR 100 100 100 100 18
            Fill-Path $g $front $White
            $front.Dispose()
        }
        'colmgr' {
            $pen = New-Pen $White 14
            $g.DrawLine($pen, 58, 62, 198, 62)
            $pen.Dispose()
            $b = New-Object System.Drawing.SolidBrush($White)
            $lb = New-Object System.Drawing.SolidBrush($Light)
            $g.FillPath($b, (New-RR 62 86 34 104 8))
            $g.FillPath($lb, (New-RR 112 116 34 74 8))
            $g.FillPath($b, (New-RR 162 96 34 94 8))
            $b.Dispose()
            $lb.Dispose()
        }
        'celledit' {
            $cell = New-RR 50 62 84 84 14
            $pen = New-Pen $White 14
            $g.DrawPath($pen, $cell)
            $pen.Dispose()
            $cell.Dispose()
            # pencil: white shaft plus a light-blue wooden tip, kept clear of the cell
            $body = New-Object System.Drawing.Drawing2D.GraphicsPath
            $body.AddPolygon(@(
                    (New-Object System.Drawing.PointF(135.8, 177.0)),
                    (New-Object System.Drawing.PointF(157.0, 198.2)),
                    (New-Object System.Drawing.PointF(192.3, 162.9)),
                    (New-Object System.Drawing.PointF(171.1, 141.7))
                ))
            Fill-Path $g $body $White
            $body.Dispose()
            $tip = New-Object System.Drawing.Drawing2D.GraphicsPath
            $tip.AddPolygon(@(
                    (New-Object System.Drawing.PointF(128.0, 206.0)),
                    (New-Object System.Drawing.PointF(135.8, 177.0)),
                    (New-Object System.Drawing.PointF(157.0, 198.2))
                ))
            Fill-Path $g $tip $Light
            $tip.Dispose()
        }
        'link' {
            $state = $g.Save()
            $g.TranslateTransform([single]128, [single]128)
            $g.RotateTransform([single]-45)
            $pen = New-Pen $White 16
            $a = New-RR -84 -30 84 60 30
            $b = New-RR -4 -30 84 60 30
            $g.DrawPath($pen, $a)
            $g.DrawPath($pen, $b)
            $pen.Dispose()
            $a.Dispose()
            $b.Dispose()
            $g.Restore($state)
        }
        'jump' {
            $pen = New-Pen $White 16
            $g.DrawLine($pen, 128, 62, 128, 150)
            $g.DrawLines($pen, @(
                    (New-Object System.Drawing.PointF(98, 122)),
                    (New-Object System.Drawing.PointF(128, 154)),
                    (New-Object System.Drawing.PointF(158, 122))
                ))
            $pen.Dispose()
            $l = New-Pen $Light 18
            $g.DrawLine($l, 66, 192, 190, 192)
            $l.Dispose()
        }
        default { throw "unknown glyph $name" }
    }
}

function Render-Icon([string]$name, [int]$size) {
    $bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.Clear([System.Drawing.Color]::Transparent)
    $k = [single]($size / 256.0)
    $g.ScaleTransform($k, $k)
    Draw-Badge $g
    Draw-Glyph $g $name
    $g.Dispose()
    return $bmp
}

# BitConverter results must be pushed one byte at a time: PowerShell unrolls a byte[] into
# the pipeline, and BinaryWriter.Write(Object[]) silently binds to Write(byte) and writes 01.
function Write-I16($bw, [int]$v) {
    foreach ($x in [System.BitConverter]::GetBytes([int16]$v)) { $bw.Write([byte]$x) }
}

function Write-I32($bw, [int]$v) {
    foreach ($x in [System.BitConverter]::GetBytes($v)) { $bw.Write([byte]$x) }
}

# Entries up to 48px always stay uncompressed 32bpp DIB (older shells read them everywhere).
function New-BmpPayload($bmp) {
    $s = $bmp.Width
    $rect = New-Object System.Drawing.Rectangle(0, 0, $s, $s)
    $data = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
        [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $stride = [Math]::Abs($data.Stride)
    $src = New-Object byte[] ($stride * $s)
    [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $src, 0, $src.Length)
    $bmp.UnlockBits($data)

    $xor = New-Object byte[] ($stride * $s)
    for ($y = 0; $y -lt $s; $y++) {
        $from = ($s - 1 - $y) * $stride
        [Array]::Copy($src, $from, $xor, $y * $stride, $stride)
    }
    $maskRow = [int](([Math]::Ceiling($s / 32.0)) * 4)
    $mask = New-Object byte[] ($maskRow * $s)

    $out = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($out)
    Write-I32 $bw 40
    Write-I32 $bw $s
    Write-I32 $bw ($s * 2)
    Write-I16 $bw 1
    Write-I16 $bw 32
    Write-I32 $bw 0
    Write-I32 $bw $xor.Length
    Write-I32 $bw 0
    Write-I32 $bw 0
    Write-I32 $bw 0
    Write-I32 $bw 0
    $bw.Write($xor)
    $bw.Write($mask)
    $bw.Flush()
    $bytes = $out.ToArray()
    $bw.Dispose(); $out.Dispose()
    return ,$bytes
}

function New-PngPayload($bmp) {
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bytes = $ms.ToArray()
    $ms.Dispose()
    return ,$bytes
}

function Write-Ico([string]$path, $rendered, [bool]$AllowPng) {
    $payloads = @()
    foreach ($r in $rendered) {
        # Roslyn's ApplicationIcon reader only understands DIB entries; a PNG stream inside
        # app.ico fails the build with CS7065, so the exe icon stays uncompressed.
        if ($AllowPng -and $r.size -ge 64) { $payloads += ,(New-PngPayload $r.bmp) }
        else { $payloads += ,(New-BmpPayload $r.bmp) }
    }
    $fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Create)
    $bw = New-Object System.IO.BinaryWriter($fs)
    Write-I16 $bw 0
    Write-I16 $bw 1
    Write-I16 $bw $rendered.Count
    $offset = 6 + 16 * $rendered.Count
    for ($i = 0; $i -lt $rendered.Count; $i++) {
        $s = $rendered[$i].size
        $p = $payloads[$i]
        $dim = if ($s -ge 256) { [byte]0 } else { [byte]$s }
        $bw.Write($dim)
        $bw.Write($dim)
        $bw.Write([byte]0)
        $bw.Write([byte]0)
        Write-I16 $bw 1
        Write-I16 $bw 32
        Write-I32 $bw $p.Length
        Write-I32 $bw $offset
        $offset = $offset + $p.Length
    }
    foreach ($p in $payloads) { $bw.Write($p) }
    $bw.Flush(); $bw.Dispose(); $fs.Dispose()
}

function Write-Sheet256([string]$path, [string[]]$names) {
    $cell = 132
    $cols = 4
    $rows = [int][Math]::Ceiling($names.Count / $cols)
    $w = $cols * $cell + 8
    $h = $rows * $cell + 8
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.ColorTranslator]::FromHtml('#F1F7FC'))
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $font = New-Object System.Drawing.Font('Segoe UI', 9)
    $fg = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#143C64'))
    for ($i = 0; $i -lt $names.Count; $i++) {
        $cx = ($i % $cols) * $cell + 8
        $cy = [int][Math]::Floor($i / $cols) * $cell + 8
        $ly = $cy + 118
        $big = Render-Icon $names[$i] 256
        $rc = New-Object System.Drawing.Rectangle($cx, $cy, 120, 120)
        $g.DrawImage($big, $rc)
        $big.Dispose()
        $g.DrawString($names[$i], $font, $fg, $cx, $ly)
    }
    $font.Dispose(); $fg.Dispose(); $g.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

# Small sizes are where icons actually break, so show 16/32 at 4x zoom on both caption
# colours (light title bar and dark title bar).
function Write-SheetSmall([string]$path, [string[]]$names) {
    $zoom = 4
    $rowH = 32 * $zoom + 10
    $w = 960
    $h = $names.Count * $rowH + 12
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
    $font = New-Object System.Drawing.Font('Consolas', 10)
    $wb = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#143C64'))
    $db = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#E3EEF9'))
    $dark = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#2B2B2B'))
    for ($i = 0; $i -lt $names.Count; $i++) {
        $y = $i * $rowH + 6
        $ty = $y + 4
        $gy = $y + 26
        $g.FillRectangle([System.Drawing.Brushes]::White, 0, $y, 480, $rowH)
        $g.FillRectangle($dark, 480, $y, 480, $rowH)
        $x = 12
        foreach ($s in @(16, 32)) {
            $bmp2 = Render-Icon $names[$i] $s
            $rc = New-Object System.Drawing.Rectangle($x, $ty, ($s * $zoom), ($s * $zoom))
            $g.DrawImage($bmp2, $rc)
            $bmp2.Dispose()
            $x = $x + $s * $zoom + 14
        }
        $g.DrawString($names[$i], $font, $wb, $x, $gy)
        $x = 492
        foreach ($s in @(16, 32)) {
            $bmp2 = Render-Icon $names[$i] $s
            $rc = New-Object System.Drawing.Rectangle($x, $ty, ($s * $zoom), ($s * $zoom))
            $g.DrawImage($bmp2, $rc)
            $bmp2.Dispose()
            $x = $x + $s * $zoom + 14
        }
        $g.DrawString($names[$i], $font, $db, $x, $gy)
    }
    $font.Dispose(); $wb.Dispose(); $db.Dispose(); $dark.Dispose(); $g.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

$names = @('app', 'reader', 'logs', 'notes', 'vcmp', 'cmdmgr', 'dup', 'colmgr', 'celledit', 'link', 'jump')
$iconDir = Join-Path $ProjectDir 'icons'
if (-not (Test-Path $iconDir)) { New-Item -ItemType Directory -Path $iconDir | Out-Null }

foreach ($n in $names) {
    $rendered = @()
    foreach ($s in $Sizes) { $rendered += @{ size = $s; bmp = (Render-Icon $n $s) } }
    # 'app' is the executable icon, so it goes to the project root where ApplicationIcon reads it.
    $isApp = ($n -eq 'app')
    $target = if ($isApp) { Join-Path $ProjectDir 'app.ico' } else { Join-Path $iconDir ($n + '.ico') }
    $allowPng = if ($isApp) { $false } else { $true }
    Write-Ico $target $rendered $allowPng
    foreach ($r in $rendered) { $r.bmp.Dispose() }
    Write-Output ("ico " + $n + " -> " + (Get-Item $target).Length + " bytes")
}

Write-Sheet256 (Join-Path $SheetDir 'icon-sheet-256.png') $names
Write-SheetSmall (Join-Path $SheetDir 'icon-sheet-small.png') $names
Write-Output 'sheets written'
