param(
    [Parameter(Mandatory = $true)][string]$ProcName,
    [Parameter(Mandatory = $true)][int]$Index,
    [Parameter(Mandatory = $true)][int]$X,
    [Parameter(Mandatory = $true)][int]$Y,
    [switch]$DoubleClick
)

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class CK {
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
    public delegate bool EnumWindowsProc(IntPtr h, IntPtr lp);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }

    public static IntPtr Pick(string procName, int index) {
        var pids = new System.Collections.Generic.List<uint>();
        foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName(procName))
            pids.Add((uint)p.Id);
        var list = new System.Collections.Generic.List<IntPtr>();
        EnumWindows((h, l) => {
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (!pids.Contains(pid) || !IsWindowVisible(h)) return true;
            RECT r; GetWindowRect(h, out r);
            if (r.Right - r.Left < 80 || r.Bottom - r.Top < 80) return true;
            list.Add(h);
            return true;
        }, IntPtr.Zero);
        return index < list.Count ? list[index] : IntPtr.Zero;
    }

    public static void ClickAt(IntPtr h, int x, int y, bool dbl) {
        RECT r; GetWindowRect(h, out r);
        POINT old; GetCursorPos(out old);
        SetCursorPos(r.Left + x, r.Top + y);
        System.Threading.Thread.Sleep(80);
        mouse_event(dbl ? 0x0008u : 0x0002u, 0u, 0u, 0u, UIntPtr.Zero);   // LEFTDOWN / LBUTTONDBLCLK
        mouse_event(0x0004u, 0u, 0u, 0u, UIntPtr.Zero);                   // LEFTUP
        if (dbl) { mouse_event(0x0008u, 0u, 0u, 0u, UIntPtr.Zero); mouse_event(0x0004u, 0u, 0u, 0u, UIntPtr.Zero); }
        SetCursorPos(old.X, old.Y);
    }
}
'@

$top = [CK]::Pick($ProcName, $Index)
if ($top -eq [IntPtr]::Zero) { Write-Output "no window at index $Index"; exit 1 }
[CK]::ClickAt($top, $X, $Y, [bool]$DoubleClick)
Write-Output "clicked"
