param(
    [Parameter(Mandatory = $true)][string]$In,
    [Parameter(Mandatory = $true)][string]$Rgb,     # "31,113,188"
    [int]$Tolerance = 6,
    [int]$MinRun = 40                                # ignore isolated anti-alias pixels
)

# acccheck.ps1 tells you a frequent color is off-palette; this tells you WHERE it is,
# so you can map the pixels back to the control that painted them.
# It reports the bounding box plus a few sample rows, which is usually enough to name the widget.
Add-Type -AssemblyName System.Drawing
$parts = $Rgb.Split(",")
$tr = [int]$parts[0]; $tg = [int]$parts[1]; $tb = [int]$parts[2]

$bmp = New-Object System.Drawing.Bitmap($In)
$w = $bmp.Width; $h = $bmp.Height
$rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
$data = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
                     [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$rowBytes = New-Object byte[] ($data.Stride)
$buf = New-Object byte[] ($data.Stride * $h)
[System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $buf, 0, $buf.Length)
$bmp.UnlockBits($data)
$bmp.Dispose()

$x0 = -1; $y0 = -1; $x1 = -1; $y1 = -1; $n = 0
$tile = 200
$tiles = @{}
$rowHits = @{}
for ($y = 0; $y -lt $h; $y++) {
    $off = $y * $data.Stride
    $run = 0
    for ($x = 0; $x -lt $w; $x++) {
        $i = $off + $x * 4
        $dr = [Math]::Abs($buf[$i + 2] - $tr); $dg = [Math]::Abs($buf[$i + 1] - $tg); $db = [Math]::Abs($buf[$i] - $tb)
        if ($dr -le $Tolerance -and $dg -le $Tolerance -and $db -le $Tolerance) {
            $n++
            if ($x0 -lt 0) { $x0 = $x; $y0 = $y; $x1 = $x; $y1 = $y }
            if ($x -lt $x0) { $x0 = $x }; if ($x -gt $x1) { $x1 = $x }
            if ($y -lt $y0) { $y0 = $y }; if ($y -gt $y1) { $y1 = $y }
            $tk = ("{0},{1}" -f [Math]::Floor($x / $tile), [Math]::Floor($y / $tile))
            if (-not $tiles.ContainsKey($tk)) { $tiles[$tk] = 0 }
            $tiles[$tk] += 1
            $run++
        } else {
            if ($run -ge $MinRun) { if (-not $rowHits.ContainsKey($y)) { $rowHits[$y] = 0 }; $rowHits[$y] += $run }
            $run = 0
        }
    }
    if ($run -ge $MinRun) { if (-not $rowHits.ContainsKey($y)) { $rowHits[$y] = 0 }; $rowHits[$y] += $run }
}

Write-Output ("image={0} rgb={1} pixels={2} bbox={3},{4} - {5},{6} (w={7} h={8})" -f `
    (Split-Path -Leaf $In), $Rgb, $n, $x0, $y0, $x1, $y1, ($x1 - $x0 + 1), ($y1 - $y0 + 1))
$rows = $rowHits.Keys | Sort-Object
Write-Output ("longRuns rows={0} first={1} last={2}" -f $rows.Count, (($rows | Select-Object -First 3) -join "/"), (($rows | Select-Object -Last 3) -join "/"))
# 热点瓦片：像素是成片（底块）还是散在（字形边缘），看这里最直观
$hot = $tiles.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 6
foreach ($t in $hot) {
    Write-Output ("tile {0} x{1}-{2} y{3}-{4} hits={5}" -f $t.Name, `
        ([int]$t.Name.Split(",")[0] * $tile), ([int]$t.Name.Split(",")[0] * $tile + $tile - 1), `
        ([int]$t.Name.Split(",")[1] * $tile), ([int]$t.Name.Split(",")[1] * $tile + $tile - 1), $t.Value)
}
