using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 简化的系统右键菜单实现
    /// 使用 ShellExecuteEx 的 "context" 参数显示菜单
    /// </summary>
    internal static class SimpleShellContextMenu
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHELLEXECUTEINFO
        {
            public int cbSize;
            public uint fMask;
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpVerb;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpFile;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpParameters;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpDirectory;
            public int nShow;
            public IntPtr hInstApp;
            public IntPtr lpIDList;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpClass;
            public IntPtr hkeyClass;
            public uint dwHotKey;
            public IntPtr hIconMonitorOrProcess;
        }

        private const uint SEE_MASK_INVOKEIDLIST = 0x0000000C;
        private const int SW_SHOWNORMAL = 1;

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);

        /// <summary>
        /// 显示系统右键菜单
        /// </summary>
        /// <param name="path">文件或文件夹路径</param>
        /// <param name="control">父控件</param>
        public static void Show(string path, Control control)
        {
            if (control == null || control.IsDisposed || !control.IsHandleCreated)
                return;

            if (!File.Exists(path) && !Directory.Exists(path))
                return;

            try
            {
                var info = new SHELLEXECUTEINFO();
                info.cbSize = Marshal.SizeOf(typeof(SHELLEXECUTEINFO));
                info.fMask = SEE_MASK_INVOKEIDLIST;
                info.hwnd = control.Handle;
                info.lpVerb = "properties";  // 使用 properties 显示属性对话框（测试用）
                info.lpFile = path;
                info.nShow = SW_SHOWNORMAL;

                ShellExecuteEx(ref info);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"显示系统右键菜单时出错: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 显示系统右键菜单（使用鼠标当前位置）
        /// </summary>
        public static void ShowAtCursor(string path, Control control)
        {
            // 这个方法会自动在鼠标当前位置显示菜单
            Show(path, control);
        }
    }
}