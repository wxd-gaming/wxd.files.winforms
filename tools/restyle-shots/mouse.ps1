param(
  [Parameter(Mandatory=$true)][int]$X,
  [Parameter(Mandatory=$true)][int]$Y,
  [ValidateSet("rmb","lmb","move")][string]$Mode = "move"
)
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class MS {
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
  public const uint LD=0x0002, LU=0x0004, RD=0x0008, RU=0x0010;
  public static void Go(int x,int y){ SetCursorPos(x,y); }
  public static void Right(){ mouse_event(RD,0,0,0,IntPtr.Zero); System.Threading.Thread.Sleep(60); mouse_event(RU,0,0,0,IntPtr.Zero); }
  public static void Left(){ mouse_event(LD,0,0,0,IntPtr.Zero); System.Threading.Thread.Sleep(60); mouse_event(LU,0,0,0,IntPtr.Zero); }
}
'@
[MS]::Go($X,$Y)
Start-Sleep -Milliseconds 120
if ($Mode -eq "rmb") { [MS]::Right() }
if ($Mode -eq "lmb") { [MS]::Left() }
Write-Host "ok $Mode $X $Y"
