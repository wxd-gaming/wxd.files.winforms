param(
    [string]$MatchTitle = "*",
    [string]$ProcName = "*",
    [Parameter(Mandatory = $true)][string]$Out
)

Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies @('System.Drawing') -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
public static class PWC {
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    public static Bitmap Capture(IntPtr h, int w, int b) {
        var bmp = new Bitmap(w, b);
        using (var g = Graphics.FromImage(bmp)) {
            IntPtr dc = g.GetHdc();
            try { PrintWindow(h, dc, 2); } finally { g.ReleaseHdc(dc); }
        }
        return bmp;
    }
}
"@

$target = Get-Process | Where-Object { $_.ProcessName -like $ProcName -and $_.MainWindowTitle -like $MatchTitle } | Select-Object -First 1
if ($null -eq $target) { Write-Output "no window matching proc=$ProcName title=$MatchTitle"; exit 1 }

$h = $target.MainWindowHandle
$rc = New-Object PWC+RECT
[void][PWC]::GetWindowRect($h, [ref]$rc)
$w = $rc.Right - $rc.Left
$b = $rc.Bottom - $rc.Top
Write-Output ("title=" + $target.MainWindowTitle + " window=" + $w + "x" + $b)

$bmp = [PWC]::Capture($h, $w, $b)
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Output ("saved=" + $Out)
