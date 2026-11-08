Add-Type -Namespace Win -Name Enum2 -MemberDefinition @'
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
  Out.Add(pid + "|" + sb.ToString() + "|" + (r.Right - r.Left) + "x" + (r.Bottom - r.Top) + "|" + r.Left + "," + r.Top);
  return true;
}
public static void Run() { Out.Clear(); EnumWindows(Cb, IntPtr.Zero); }
'@
[void][Win.Enum2]::Run()
$rows = foreach ($line in [Win.Enum2]::Out) {
  $p = $line -split '\|', 4
  $proc = Get-Process -Id $p[0] -ErrorAction SilentlyContinue
  [pscustomobject]@{ Pid = $p[0]; Proc = if ($proc) { $proc.ProcessName } else { '?' }; Title = $p[1]; Size = $p[2]; Pos = $p[3] }
}
$rows | Out-File -Encoding utf8 'D:/work/wxd.files.winforms/tools/restyle-shots/wins.txt'
