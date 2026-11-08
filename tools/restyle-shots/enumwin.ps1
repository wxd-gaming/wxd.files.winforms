param([string]$ProcName = '')
Add-Type -Namespace Win -Name Enum -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, System.Text.StringBuilder s, int n);
[DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr h);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
public delegate bool EnumProc(IntPtr h, IntPtr l);
public static System.Collections.Generic.List<string> Out = new System.Collections.Generic.List<string>();
public static bool Cb(IntPtr h, IntPtr l) {
  if (!IsWindowVisible(h)) return true;
  int len = GetWindowTextLength(h);
  if (len == 0) return true;
  var sb = new System.Text.StringBuilder(len + 1);
  GetWindowText(h, sb, sb.Capacity);
  uint pid; GetWindowThreadProcessId(h, out pid);
  RECT r; GetWindowRect(h, out r);
  Out.Add(pid + "|" + sb + "|" + (r.Right - r.Left) + "x" + (r.Bottom - r.Top));
  return true;
}
public static void Run() { Out.Clear(); EnumWindows(Cb, IntPtr.Zero); }
'@
[void][Win.Enum]::Run()
foreach ($line in [Win.Enum]::Out) {
  $parts = $line -split '\|', 3
  if ($ProcName) {
    $pname = (Get-Process -Id $parts[0] -ErrorAction SilentlyContinue).ProcessName
    if ($pname -ne $ProcName) { continue }
    "$pname $($parts[1]) $($parts[2])"
  } else {
    "$($parts[1]) $($parts[2])"
  }
}
