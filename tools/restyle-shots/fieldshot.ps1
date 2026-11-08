param(
    [Parameter(Mandatory = $true)][string]$ProcName,
    [Parameter(Mandatory = $true)][string]$TitleMatch,
    [Parameter(Mandatory = $true)][string]$Out,
    [Parameter(Mandatory = $true)][int]$X,
    [Parameter(Mandatory = $true)][int]$Y,
    [string]$SetText = "",
    [switch]$Enter,
    [int]$WaitMs = 1500
)

Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies @('System.Drawing') -TypeDefinition @"
using System;
using System.Text;
using System.Drawing;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class FS {
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumWindowsProc cb, IntPtr l);
    public delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, string lp);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, IntPtr lp);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }

    public static IntPtr TopWindow(string procName, string titleMatch) {
        uint target = 0;
        foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName(procName))
            target = (uint)p.Id;
        IntPtr found = IntPtr.Zero;
        EnumWindows((h, l) => {
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (pid != target || !IsWindowVisible(h)) return true;
            var sb = new StringBuilder(512); GetWindowText(h, sb, 512);
            if (sb.ToString().IndexOf(titleMatch, StringComparison.OrdinalIgnoreCase) < 0) return true;
            found = h; return false;
        }, IntPtr.Zero);
        return found;
    }

    public static IntPtr EditAt(IntPtr top, int x, int y) {
        RECT tr; GetWindowRect(top, out tr);
        int sx = tr.Left + x, sy = tr.Top + y;
        IntPtr hit = IntPtr.Zero;
        EnumChildWindows(top, (h, l) => {
            var sb = new StringBuilder(256); GetClassName(h, sb, 256);
            if (sb.ToString().IndexOf("EDIT", StringComparison.OrdinalIgnoreCase) < 0) return true;
            RECT r; GetWindowRect(h, out r);
            if (sx >= r.Left && sx <= r.Right && sy >= r.Top && sy <= r.Bottom) { hit = h; return false; }
            return true;
        }, IntPtr.Zero);
        return hit;
    }

    public static void Type(IntPtr h, string text) {
        SendMessage(h, 0x000C, IntPtr.Zero, text);          // WM_SETTEXT
    }

    public static void PressEnter(IntPtr h) {
        IntPtr vk = new IntPtr(0x0D);                       // VK_RETURN
        SendMessage(h, 0x0100, vk, IntPtr.Zero);            // WM_KEYDOWN
        SendMessage(h, 0x0101, vk, IntPtr.Zero);            // WM_KEYUP
    }

    public static void DumpEdits(IntPtr top) {
        RECT tr; GetWindowRect(top, out tr);
        var lines = new List<string>();
        EnumChildWindows(top, (h, l) => {
            var sb = new StringBuilder(256); GetClassName(h, sb, 256);
            RECT r; GetWindowRect(h, out r);
            lines.Add(string.Format("hwnd={0:X} class={1} rel=({2},{3})-({4},{5}) vis={6}",
                (long)h, sb.ToString(), r.Left - tr.Left, r.Top - tr.Top, r.Right - tr.Left, r.Bottom - tr.Top,
                IsWindowVisible(h)));
            return true;
        }, IntPtr.Zero);
        foreach (var s in lines) Console.WriteLine(s);
    }

    public static Bitmap Capture(IntPtr h) {
        RECT r; GetWindowRect(h, out r);
        int w = r.Right - r.Left, b = r.Bottom - r.Top;
        var bmp = new Bitmap(w, b);
        using (var g = Graphics.FromImage(bmp)) {
            IntPtr dc = g.GetHdc();
            try { PrintWindow(h, dc, 2); } finally { g.ReleaseHdc(dc); }
        }
        return bmp;
    }
}
"@

$top = [FS]::TopWindow($ProcName, $TitleMatch)
if ($top -eq [IntPtr]::Zero) { Write-Output "no window"; exit 1 }
$edit = [FS]::EditAt($top, $X, $Y)
if ($edit -eq [IntPtr]::Zero) { Write-Output "no edit at $X,$Y"; [FS]::DumpEdits($top); exit 1 }

if ($SetText -ne "") { [FS]::Type($edit, $SetText) }
if ($Enter) { [FS]::PressEnter($edit) } else { [void][FS]::SetFocus($edit) }

Start-Sleep -Milliseconds $WaitMs
$bmp = [FS]::Capture($top)
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Output "saved=$Out"
