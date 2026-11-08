param([string]$In,[int]$Y,[int]$X0=600,[int]$X1=1000)
Add-Type -AssemblyName System.Drawing
$b = [System.Drawing.Bitmap]::FromFile($In)
$sb = New-Object System.Text.StringBuilder
$prev = ""
for ($x = $X0; $x -le $X1 -and $x -lt $b.Width; $x++) {
  $p = $b.GetPixel($x, $Y)
  $t = "$($p.R),$($p.G),$($p.B)"
  if ($t -ne $prev) { [void]$sb.Append("x$x=[$t] "); $prev = $t }
}
$b.Dispose()
Write-Output $sb.ToString()
