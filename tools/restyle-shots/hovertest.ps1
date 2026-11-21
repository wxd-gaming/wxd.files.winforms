param(
    [Parameter(Mandatory = $true)][string]$ProcName,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$X = 365,
    [int]$Y = 44,
    [int]$ProcId = 0,
    [switch]$NoHover,
    [switch]$Maximize,
    [switch]$Pressed
)

Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies @('System.Drawing') -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
public static class HV {
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    public static IntPtr Find(string name, int onlyPid) {
        var pids = new System.Collections.Generic.List<uint>();
        foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName(name)) pids.Add((uint)p.Id);
        IntPtr found = IntPtr.Zero;
        EnumWindows((h, l) => {
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (!pids.Contains(pid) || !IsWindowVisible(h)) return true;
            if (onlyPid != 0 && pid != (uint)onlyPid) return true;
            RECT r; GetWindowRect(h, out r);
            if (r.Right - r.Left < 200 || r.Bottom - r.Top < 200) return true;
            found = h; return false;
        }, IntPtr.Zero);
        return found;
    }
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

$h = [HV]::Find($ProcName, $ProcId)
if ($h -eq [IntPtr]::Zero) { Write-Output "no window"; exit 1 }
if ($Maximize) { [void][HV]::ShowWindow($h, 3); Start-Sleep -Milliseconds 900 }
$rc = New-Object HV+RECT
[void][HV]::GetWindowRect($h, [ref]$rc)
$w = $rc.Right - $rc.Left
$b = $rc.Bottom - $rc.Top
if (-not $NoHover) {
    [void][HV]::SetCursorPos($rc.Left + $X, $rc.Top + $Y)
    Start-Sleep -Milliseconds 400
}
if ($Pressed) {
    [HV]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 250
}
$bmp = [HV]::Capture($h, $w, $b)
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
if ($Pressed) { [HV]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero) }
Write-Output ("saved=" + $Out + " rect=" + $w + "x" + $b)
