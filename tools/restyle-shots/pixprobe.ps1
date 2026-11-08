param(
    [Parameter(Mandatory = $true)][string]$In,
    [Parameter(Mandatory = $true)][string]$Region   # x,y,w,h
)

Add-Type -AssemblyName System.Drawing
$parts = $Region.Split(",")
$x = [int]$parts[0]; $y = [int]$parts[1]; $w = [int]$parts[2]; $h = [int]$parts[3]

$bmp = New-Object System.Drawing.Bitmap($In)
$hist = @{}
for ($j = 0; $j -lt $h; $j++) {
    for ($i = 0; $i -lt $w; $i++) {
        $c = $bmp.GetPixel($x + $i, $y + $j)
        $key = "{0},{1},{2}" -f $c.R, $c.G, $c.B
        if ($hist.ContainsKey($key)) { $hist[$key] = $hist[$key] + 1 } else { $hist[$key] = 1 }
    }
}
$sorted = $hist.GetEnumerator() | Sort-Object -Property Value -Descending
$bg = $sorted[0]
Write-Output ("bg=" + $bg.Key + " count=" + $bg.Value)
$bgLum = 0.299 * [int]$bg.Key.Split(",")[0] + 0.587 * [int]$bg.Key.Split(",")[1] + 0.114 * [int]$bg.Key.Split(",")[2]
$best = $null; $bestD = 0
foreach ($e in $sorted) {
    $p = $e.Key.Split(",")
    $lum = 0.299 * [int]$p[0] + 0.587 * [int]$p[1] + 0.114 * [int]$p[2]
    $d = [Math]::Abs($lum - $bgLum)
    if ($d -gt $bestD -and $e.Value -ge 3) { $bestD = $d; $best = $e }
}
Write-Output ("text=" + $best.Key + " count=" + $best.Value + " lumDiff=" + [Math]::Round($bestD, 1))
Write-Output ("top8=" + (($sorted | Select-Object -First 8 | ForEach-Object { $_.Key + ":" + $_.Value }) -join " "))
$bmp.Dispose()
