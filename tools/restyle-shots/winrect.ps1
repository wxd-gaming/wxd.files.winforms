param([string]$TitleMatch = '日志分析')
Add-Type -Namespace Win -Name Native -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
'@
foreach ($p in Get-Process | Where-Object { $_.MainWindowHandle -ne 0 }) {
  $t = $p.MainWindowTitle
  if ($t -notlike "*$TitleMatch*") { continue }
  $h = $p.MainWindowHandle
  $wr = New-Object 'Win.Native+RECT'
  $cr = New-Object 'Win.Native+RECT'
  [void][Win.Native]::GetWindowRect($h, [ref]$wr)
  [void][Win.Native]::GetClientRect($h, [ref]$cr)
  "pid=$($p.Id) hwnd=$h window=$($wr.Right-$wr.Left)x$($wr.Bottom-$wr.Top) client=$($cr.Right)x$($cr.Bottom) iconic=$([Win.Native]::IsIconic($h)) title=$t"
}
