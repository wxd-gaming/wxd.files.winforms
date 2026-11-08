param([string]$TitleMatch,[int]$W,[int]$H,[int]$X=60,[int]$Y=60)
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public class WinMove {
  delegate bool EnumProc(IntPtr h, IntPtr lp);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lp);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, StringBuilder s, int max);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool repaint);
  static string _pattern;
  static IntPtr _hit;
  public static IntPtr Find(string pattern) {
    _pattern = pattern; _hit = IntPtr.Zero;
    EnumWindows((h, lp) => {
      if (!IsWindowVisible(h)) return true;
      var sb = new StringBuilder(512);
      GetWindowTextW(h, sb, 512);
      if (sb.ToString().IndexOf(_pattern, StringComparison.OrdinalIgnoreCase) >= 0) { _hit = h; return false; }
      return true;
    }, IntPtr.Zero);
    return _hit;
  }
}
'@
$h = [WinMove]::Find($TitleMatch)
Write-Host ("hwnd=" + $h)
if ($h -ne [IntPtr]::Zero) {
  [void][WinMove]::MoveWindow($h, $X, $Y, $W, $H, $true)
  Start-Sleep -Seconds 2
  Write-Host "moved"
}
