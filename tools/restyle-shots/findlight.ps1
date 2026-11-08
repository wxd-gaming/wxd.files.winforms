param([string]$Path, [int]$Y0, [int]$Y1, [int]$X0, [int]$X1, [int]$Min = 200)
Add-Type -AssemblyName System.Drawing
$bmp = [System.Drawing.Bitmap]::new($Path)
for ($y = $Y0; $y -le $Y1; $y++) {
    $count = 0
    $first = -1
    $last = -1
    for ($x = $X0; $x -le $X1; $x++) {
        $c = $bmp.GetPixel($x, $y)
        if ($c.R -ge $Min -and $c.G -ge $Min -and $c.B -ge $Min) {
            $count++
            if ($first -lt 0) { $first = $x }
            $last = $x
        }
    }
    if ($count -gt 0) {
        Write-Output ('y={0} light={1} x={2}..{3}' -f $y, $count, $first, $last)
    }
}
$bmp.Dispose()
