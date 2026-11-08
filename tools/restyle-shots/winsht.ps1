param(
    [Parameter(Mandatory = $true)][string]$ProcName,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$Index = 0
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies @('System.Drawing') -TypeDefinition @"
using System;
using System.Text;
using System.Drawing;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class WS {
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
    public delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }

    public static List<IntPtr> Tops(string procName) {
        var pids = new List<uint>();
        foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName(procName))
            pids.Add((uint)p.Id);
        var list = new List<IntPtr>();
        EnumWindows((h, l) => {
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (!pids.Contains(pid) || !IsWindowVisible(h)) return true;
            RECT r; GetWindowRect(h, out r);
            if (r.Right - r.Left < 80 || r.Bottom - r.Top < 80) return true;
            list.Add(h);
            return true;
        }, IntPtr.Zero);
        return list;
    }

    public static void Shot(IntPtr h, string outPath) {
        RECT r; GetWindowRect(h, out r);
        int w = r.Right - r.Left, b = r.Bottom - r.Top;
        var bmp = new Bitmap(w, b);
        using (var g = Graphics.FromImage(bmp)) {
            IntPtr dc = g.GetHdc();
            try { PrintWindow(h, dc, 2); } finally { g.ReleaseHdc(dc); }
        }
        bmp.Save(outPath, System.Drawing.Imaging.ImageFormat.Png);
        bmp.Dispose();
        Console.WriteLine("size=" + w + "x" + b);
    }
}
"@

$tops = [WS]::Tops($ProcName)
Write-Output ("count=" + $tops.Count)
if ($tops.Count -le $Index) { Write-Output "no window at index $Index"; exit 1 }
[WS]::Shot($tops[$Index], $Out)
Write-Output "saved=$Out"
