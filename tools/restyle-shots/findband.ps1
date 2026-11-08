param([string]$In,[int]$X=700,[int]$Y0=150,[int]$Y1=230)
Add-Type -AssemblyName System.Drawing
$b = [System.Drawing.Bitmap]::FromFile($In)
$sb = New-Object System.Text.StringBuilder
for ($y = $Y0; $y -le $Y1 -and $y -lt $b.Height; $y++) {
  $p = $b.GetPixel($X, $y); [void]$sb.Append("${y}=[$($p.R),$($p.G),$($p.B)] ")
}
$b.Dispose(); Write-Output $sb.ToString()
