param([string]$In)
Add-Type -AssemblyName System.Drawing
$b = [System.Drawing.Bitmap]::FromFile($In)
$hits = @()
for ($y = 100; $y -lt 220 -and $y -lt $b.Height; $y++) {
  $run = 0; $cur = ""
  for ($x = 700; $x -lt 900 -and $x -lt $b.Width; $x++) {
    $p = $b.GetPixel($x, $y); $t = "$($p.R),$($p.G),$($p.B)"
    if ($t -eq $cur) { $run++ } else { $cur = $t; $run = 1 }
    if ($run -ge 14 -and ($t -eq "32,32,32" -or $t -eq "240,240,240")) { $hits += "y$y=$t"; break }
  }
}
$b.Dispose()
Write-Output ($hits -join " ")
