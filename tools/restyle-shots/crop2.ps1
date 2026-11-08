param([string]$In,[string]$Out,[int]$X,[int]$Y,[int]$W,[int]$H,[int]$Scale=2)
Add-Type -AssemblyName System.Drawing
$src=[System.Drawing.Image]::FromFile($In)
$bw=[int]($W*$Scale)
$bh=[int]($H*$Scale)
$big=New-Object System.Drawing.Bitmap($bw,$bh)
$g2=[System.Drawing.Graphics]::FromImage($big)
$g2.InterpolationMode=[System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
$srcRect=New-Object System.Drawing.Rectangle($X,$Y,$W,$H)
$dstRect=New-Object System.Drawing.Rectangle(0,0,$bw,$bh)
[void]$g2.DrawImage($src,$dstRect,$srcRect,[System.Drawing.GraphicsUnit]::Pixel)
$g2.Dispose();$src.Dispose()
$big.Save($Out,[System.Drawing.Imaging.ImageFormat]::Png);$big.Dispose()
"saved=$Out"
