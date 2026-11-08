param([string]$Path, [int]$Y0, [int]$Y1, [int]$X0, [int]$X1, [int]$Step = 60)
Add-Type -AssemblyName System.Drawing
$bmp = [System.Drawing.Bitmap]::new($Path)
for ($y = $Y0; $y -le $Y1; $y++) {
    $seen = [ordered]@{}
    for ($x = $X0; $x -le $X1; $x += $Step) {
        $c = $bmp.GetPixel($x, $y)
        $k = '#{0:X2}{1:X2}{2:X2}' -f $c.R, $c.G, $c.B
        if (-not $seen.Contains($k)) { $seen[$k] = $x }
    }
    $parts = $seen.Keys | ForEach-Object { '{0}@{1}' -f $_, $seen[$_] }
    Write-Output ('y={0}  {1}' -f $y, ($parts -join ' '))
}
$bmp.Dispose()
