param([string]$ProcName = "gridshot")
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @'
using System;using System.Text;using System.Collections.Generic;using System.Runtime.InteropServices;
public class RJ{
 delegate bool E(IntPtr h,IntPtr l);
 [DllImport("user32.dll")] static extern bool EnumWindows(E cb,IntPtr l);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint pid);
 [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h,out R r);
 [StructLayout(LayoutKind.Sequential)] public struct R{public int L,T,Rr,B;}
 public static List<string> Tops(string pn){
  var pids=new List<uint>();
  foreach(System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName(pn)) pids.Add((uint)p.Id);
  var outp=new List<string>();
  EnumWindows((h,l)=>{ uint pid; GetWindowThreadProcessId(h,out pid);
   if(pids.Contains(pid)&&IsWindowVisible(h)){ R r; GetWindowRect(h,out r);
    outp.Add(string.Format("{0},{1},{2},{3}",r.L,r.T,r.Rr,r.B)); }
   return true; },IntPtr.Zero);
  return outp; }
}
'@
foreach($s in [RJ]::Tops($ProcName)){ Write-Host "rect=$s" }
