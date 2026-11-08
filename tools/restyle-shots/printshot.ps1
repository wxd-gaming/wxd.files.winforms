param(
    [Parameter(Mandatory = $true)][string]$Exe,
    [string]$ArgList = "",
    [Parameter(Mandatory = $true)][string]$Out
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies @('System.Drawing') -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
public static class PW {
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
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

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $Exe
if ($ArgList) { $psi.Arguments = $ArgList }
$psi.UseShellExecute = $false
$p = [System.Diagnostics.Process]::Start($psi)

Start-Sleep -Milliseconds 4500
$p.Refresh()
$h = $p.MainWindowHandle
if ($h -eq [IntPtr]::Zero) { Write-Output "no window handle"; $p.Kill(); exit 1 }
[void][PW]::ShowWindow($h, 3)
Start-Sleep -Milliseconds 1200

$rc = New-Object PW+RECT
[void][PW]::GetWindowRect($h, [ref]$rc)
$w = $rc.Right - $rc.Left
$b = $rc.Bottom - $rc.Top
Write-Output ("window=" + $w + "x" + $b)

$bmp = [PW]::Capture($h, $w, $b)
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Output ("saved=" + $Out)
if ($env:THEMELAB_KEEP -ne "1") { try { $p.Kill() } catch {} }
