param([string]$TitleMatch,[int]$CX=12,[int]$CY=12)
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class TreeClick {
  delegate bool EnumProc(IntPtr h, IntPtr lp);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lp);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr lp);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, StringBuilder s, int max);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, StringBuilder s, int max);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr SendMessageW(IntPtr h, int msg, IntPtr wp, IntPtr lp);

  static string _pattern; static IntPtr _top;
  static bool TopMatch(IntPtr h, IntPtr lp) {
    if (!IsWindowVisible(h)) return true;
    var sb = new StringBuilder(512); GetWindowTextW(h, sb, 512);
    if (sb.ToString().IndexOf(_pattern, StringComparison.OrdinalIgnoreCase) >= 0) { _top = h; return false; }
    return true;
  }
  static List<IntPtr> _trees = new List<IntPtr>();
  static bool ChildMatch(IntPtr h, IntPtr lp) {
    var cn = new StringBuilder(256); GetClassNameW(h, cn, 256);
    if (cn.ToString().IndexOf("SysTreeView32", StringComparison.OrdinalIgnoreCase) >= 0) _trees.Add(h);
    return true;
  }
  public static IntPtr Top(string pattern) { _pattern = pattern; _top = IntPtr.Zero; EnumWindows(TopMatch, IntPtr.Zero); return _top; }
  public static IntPtr FirstTree(IntPtr top) { _trees.Clear(); EnumChildWindows(top, ChildMatch, IntPtr.Zero); return _trees.Count > 0 ? _trees[0] : IntPtr.Zero; }
  public static void Click(IntPtr tree, int x, int y) {
    IntPtr lp = (IntPtr)(y * 65536 + x);
    SendMessageW(tree, 0x0201, (IntPtr)0x0001, lp); // WM_LBUTTONDOWN, MK_LBUTTON
    SendMessageW(tree, 0x0202, IntPtr.Zero, lp);    // WM_LBUTTONUP
  }
}
'@
$top = [TreeClick]::Top($TitleMatch)
Write-Host ("top=" + $top)
$tree = [TreeClick]::FirstTree($top)
Write-Host ("tree=" + $tree)
if ($tree -ne [IntPtr]::Zero) { [TreeClick]::Click($tree, $CX, $CY); Write-Host "clicked" }
