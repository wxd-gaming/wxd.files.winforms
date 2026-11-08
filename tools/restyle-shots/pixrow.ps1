param([string]$Path, [int]$Y, [int]$X0, [int]$X1)
Add-Type -AssemblyName System.Drawing
$bmp = [System.Drawing.Bitmap]::new($Path)
$row = @()
for ($x = $X0; $x -le $X1; $x++) {
    $c = $bmp.GetPixel($x, $Y)
    $row += ('{0}:#{1:X2}{2:X2}{3:X2}' -f $x, $c.R, $c.G, $c.B)
}
Write-Output ($row -join ' ')
$bmp.Dispose()
