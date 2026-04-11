using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 调用 Windows 系统原生右键菜单的工具类
    /// 基于 Windows Shell API 实现，不依赖第三方库
    /// </summary>
    internal static class WindowsShellContextMenu
    {
        // COM 接口 GUID
        private static readonly Guid IID_IShellFolder = new Guid("000214E6-0000-0000-C000-000000000046");
        private static readonly Guid IID_IContextMenu = new Guid("000214E4-0000-0000-C000-000000000046");
        private static readonly Guid IID_IContextMenu2 = new Guid("000214F4-0000-0000-C000-000000000046");
        private static readonly Guid IID_IContextMenu3 = new Guid("BCFCE0A0-EC17-11D0-8D10-00A0C90F2719");

        // 常量定义
        private const int CMF_NORMAL = 0x00000000;
        private const int TPM_RETURNCMD = 0x0100;
        private const int TPM_RIGHTBUTTON = 0x0002;
        private const int TPM_NONOTIFY = 0x0080;
        private const int WM_INITMENUPOPUP = 0x0117;
        private const int WM_DRAWITEM = 0x002B;
        private const int WM_MEASUREITEM = 0x002C;
        private const int WM_MENUCHAR = 0x0120;

        /// <summary>
        /// 显示系统右键菜单
        /// </summary>
        /// <param name="path">文件或文件夹路径</param>
        /// <param name="control">父控件</param>
        /// <param name="screenX">屏幕X坐标</param>
        /// <param name="screenY">屏幕Y坐标</param>
        public static void Show(string path, Control control, int screenX, int screenY)
        {
            if (control == null || control.IsDisposed || !control.IsHandleCreated)
                return;

            if (!File.Exists(path) && !Directory.Exists(path))
                return;

            if (!IsWindows())
                return;

            IntPtr pidl = IntPtr.Zero;
            IShellFolder desktopFolder = null;
            IShellFolder parentFolder = null;
            IContextMenu contextMenu = null;
            IContextMenu2 contextMenu2 = null;
            IContextMenu3 contextMenu3 = null;
            IntPtr hMenu = IntPtr.Zero;

            bool comInitialized = false;

            try
            {
                // 初始化 COM
                try
                {
                    int hr = Ole32.CoInitializeEx(IntPtr.Zero, Ole32.COINIT_APARTMENTTHREADED);
                    comInitialized = hr == 0 || hr == 1;
                    if (hr == Ole32.RPC_E_CHANGED_MODE)
                        comInitialized = false;
                }
                catch
                {
                    comInitialized = false;
                }

                // 获取桌面文件夹
                Guid shellFolderGuid = IID_IShellFolder;
                int hrDesktop = Shell32.SHGetDesktopFolder(out IntPtr desktopFolderPtr);
                System.Diagnostics.Debug.WriteLine($"SHGetDesktopFolder result: {hrDesktop}");
                if (hrDesktop != 0 || desktopFolderPtr == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine("SHGetDesktopFolder failed");
                    return;
                }

                desktopFolder = (IShellFolder)Marshal.GetObjectForIUnknown(desktopFolderPtr);

                // 解析路径为 PIDL
                uint pchEaten = 0;
                uint pdwAttributes = 0;
                desktopFolder.ParseDisplayName(control.Handle, IntPtr.Zero, path, out pchEaten, out pidl, ref pdwAttributes);
                System.Diagnostics.Debug.WriteLine($"ParseDisplayName PIDL: {pidl}");
                if (pidl == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine($"ParseDisplayName failed for path: {path}");
                    return;
                }

                // 获取父文件夹和相对PIDL
                // 分离路径获取父目录
                string parentPath = Path.GetDirectoryName(path);
                string fileName = Path.GetFileName(path);
                
                if (string.IsNullOrEmpty(parentPath))
                {
                    // 如果是根目录，直接使用桌面文件夹
                    parentFolder = desktopFolder;
                }
                else
                {
                    // 解析父目录路径
                    IntPtr parentPidl = IntPtr.Zero;
                    desktopFolder.ParseDisplayName(control.Handle, IntPtr.Zero, parentPath, out pchEaten, out parentPidl, ref pdwAttributes);
                    if (parentPidl == IntPtr.Zero)
                    {
                        System.Diagnostics.Debug.WriteLine($"ParseDisplayName failed for parent path: {parentPath}");
                        return;
                    }

                    try
                    {
                        // 绑定到父文件夹
                        Guid iidShellFolder = IID_IShellFolder;
                        desktopFolder.BindToObject(parentPidl, IntPtr.Zero, ref iidShellFolder, out IntPtr parentFolderPtr);
                        if (parentFolderPtr == IntPtr.Zero)
                        {
                            System.Diagnostics.Debug.WriteLine("BindToObject failed");
                            return;
                        }
                        parentFolder = (IShellFolder)Marshal.GetObjectForIUnknown(parentFolderPtr);
                    }
                    finally
                    {
                        Ole32.CoTaskMemFree(parentPidl);
                    }
                }

                // 在父文件夹中解析文件名
                IntPtr childPidl = IntPtr.Zero;
                parentFolder.ParseDisplayName(control.Handle, IntPtr.Zero, fileName, out pchEaten, out childPidl, ref pdwAttributes);
                System.Diagnostics.Debug.WriteLine($"Child PIDL: {childPidl}");
                if (childPidl == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine($"ParseDisplayName failed for file: {fileName}");
                    return;
                }

                try
                {
                    // 获取上下文菜单接口
                    // 分配内存存储 PIDL 数组
                    IntPtr apidlPtr = Marshal.AllocCoTaskMem(IntPtr.Size);
                    Marshal.WriteIntPtr(apidlPtr, childPidl);

                    Guid contextMenuGuid = IID_IContextMenu;
                    parentFolder.GetUIObjectOf(control.Handle, 1, apidlPtr, ref contextMenuGuid, IntPtr.Zero, out IntPtr contextMenuPtr);
                    
                    // 释放分配的内存（但不要释放 childPidl，它会在 finally 中释放）
                    Marshal.FreeCoTaskMem(apidlPtr);
                    
                    System.Diagnostics.Debug.WriteLine($"GetUIObjectOf result: {contextMenuPtr}");
                    if (contextMenuPtr == IntPtr.Zero)
                    {
                        System.Diagnostics.Debug.WriteLine("GetUIObjectOf returned null");
                        return;
                    }

                    contextMenu = (IContextMenu)Marshal.GetObjectForIUnknown(contextMenuPtr);

                    // 尝试获取 IContextMenu3，然后 IContextMenu2
                    IntPtr contextMenu3Ptr;
                    Guid contextMenu3Guid = IID_IContextMenu3;
                    if (contextMenu.QueryInterface(ref contextMenu3Guid, out contextMenu3Ptr) == 0)
                    {
                        contextMenu3 = (IContextMenu3)Marshal.GetObjectForIUnknown(contextMenu3Ptr);
                        contextMenu2 = contextMenu3;
                    }
                    else
                    {
                        IntPtr contextMenu2Ptr;
                        Guid contextMenu2Guid = IID_IContextMenu2;
                        if (contextMenu.QueryInterface(ref contextMenu2Guid, out contextMenu2Ptr) == 0)
                        {
                            contextMenu2 = (IContextMenu2)Marshal.GetObjectForIUnknown(contextMenu2Ptr);
                        }
                    }

                    // 创建弹出菜单
                    hMenu = User32.CreatePopupMenu();
                    if (hMenu == IntPtr.Zero)
                        return;

                    // 查询上下文菜单
                    int queryResult = contextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL);
                    System.Diagnostics.Debug.WriteLine($"QueryContextMenu result: {queryResult}");
                    if (queryResult < 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"QueryContextMenu failed with result: {queryResult}");
                        return;
                    }

                    // 安装消息处理程序
                    MenuHostWindow window = new MenuHostWindow(control.Handle, contextMenu2, contextMenu3);

                    // 显示菜单
                    System.Diagnostics.Debug.WriteLine($"显示菜单在位置: {screenX}, {screenY}");
                    uint commandId = User32.TrackPopupMenuEx(
                        hMenu,
                        TPM_RETURNCMD | TPM_RIGHTBUTTON,
                        screenX,
                        screenY,
                        control.Handle,
                        IntPtr.Zero);

                    System.Diagnostics.Debug.WriteLine($"TrackPopupMenuEx 返回命令ID: {commandId}");
                    window.ClearHandlers();

                    // 执行选中的命令
                    if (commandId > 0)
                    {
                        var invokeInfo = new CMINVOKECOMMANDINFOEX();
                        invokeInfo.cbSize = Marshal.SizeOf(typeof(CMINVOKECOMMANDINFOEX));
                        invokeInfo.hwnd = control.Handle;
                        invokeInfo.lpVerb = (IntPtr)(commandId - 1);
                        invokeInfo.nShow = 1;

                        contextMenu.InvokeCommand(ref invokeInfo);
                    }
                }
                finally
                {
                    if (childPidl != IntPtr.Zero)
                        Ole32.CoTaskMemFree(childPidl);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"显示系统右键菜单时出错: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"堆栈: {ex.StackTrace}");
            }
            finally
            {
                // 清理资源
                if (hMenu != IntPtr.Zero)
                    User32.DestroyMenu(hMenu);

                if (contextMenu3 != null)
                    Marshal.ReleaseComObject(contextMenu3);

                if (contextMenu2 != null && contextMenu2 != contextMenu3)
                    Marshal.ReleaseComObject(contextMenu2);

                if (contextMenu != null)
                    Marshal.ReleaseComObject(contextMenu);

                if (parentFolder != null && parentFolder != desktopFolder)
                    Marshal.ReleaseComObject(parentFolder);

                if (desktopFolder != null)
                    Marshal.ReleaseComObject(desktopFolder);

                if (pidl != IntPtr.Zero)
                    Ole32.CoTaskMemFree(pidl);

                if (comInitialized)
                {
                    try { Ole32.CoUninitialize(); }
                    catch { }
                }
            }
        }

        /// <summary>
        /// 显示系统右键菜单（使用控件当前位置）
        /// </summary>
        public static void Show(string path, Control control)
        {
            var screenPos = control.PointToScreen(new System.Drawing.Point(0, 0));
            Show(path, control, screenPos.X, screenPos.Y);
        }

        /// <summary>
        /// 显示系统右键菜单（使用鼠标当前位置）
        /// </summary>
        public static void ShowAtCursor(string path, Control control)
        {
            var cursorPos = System.Windows.Forms.Cursor.Position;
            Show(path, control, cursorPos.X, cursorPos.Y);
        }

        private static bool IsWindows()
        {
            return Path.DirectorySeparatorChar == '\\';
        }

        #region Win32 API 声明

        private static class Ole32
        {
            public const int COINIT_APARTMENTTHREADED = 0x2;
            public const int RPC_E_CHANGED_MODE = unchecked((int)0x80010106);

            [DllImport("ole32.dll", PreserveSig = false)]
            public static extern int CoInitializeEx(IntPtr pvReserved, int dwCoInit);

            [DllImport("ole32.dll")]
            public static extern void CoUninitialize();

            [DllImport("ole32.dll")]
            public static extern void CoTaskMemFree(IntPtr pv);
        }

        private static class Shell32
        {
            [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern int SHParseDisplayName(
                [MarshalAs(UnmanagedType.LPWStr)] string pszName,
                IntPtr pbc,
                out IntPtr ppidl,
                uint sfgaoIn,
                out uint psfgaoOut);

            [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern int SHBindToParent(
                IntPtr pidl,
                [MarshalAs(UnmanagedType.LPStruct)] ref Guid riid,
                out IntPtr ppv,
                out IntPtr ppidlLast);

            [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
            public static extern int SHGetDesktopFolder(out IntPtr ppshf);
        }

        private static class User32
        {
            [DllImport("user32.dll")]
            public static extern IntPtr CreatePopupMenu();

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool DestroyMenu(IntPtr hMenu);

            [DllImport("user32.dll")]
            public static extern uint TrackPopupMenuEx(
                IntPtr hMenu,
                uint uFlags,
                int x,
                int y,
                IntPtr hWnd,
                IntPtr lptpm);

            [DllImport("user32.dll")]
            public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        }

        #endregion

        #region COM 接口定义

        [ComImport]
        [Guid("000214E4-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IContextMenu
        {
            [PreserveSig]
            int QueryContextMenu(
                IntPtr hMenu,
                uint indexMenu,
                uint idCmdFirst,
                uint idCmdLast,
                uint uFlags);

            [PreserveSig]
            int InvokeCommand(ref CMINVOKECOMMANDINFOEX lpici);

            [PreserveSig]
            int GetCommandString(
                uint idCmd,
                uint uType,
                IntPtr pReserved,
                [MarshalAs(UnmanagedType.LPStr)] out string pszName,
                uint cchMax);

            [PreserveSig]
            int QueryInterface(ref Guid riid, out IntPtr ppvObject);

            [PreserveSig]
            int AddRef();

            [PreserveSig]
            int Release();
        }

        [ComImport]
        [Guid("000214F4-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IContextMenu2 : IContextMenu
        {
            [PreserveSig]
            new int QueryContextMenu(
                IntPtr hMenu,
                uint indexMenu,
                uint idCmdFirst,
                uint idCmdLast,
                uint uFlags);

            [PreserveSig]
            new int InvokeCommand(ref CMINVOKECOMMANDINFOEX lpici);

            [PreserveSig]
            new int GetCommandString(
                uint idCmd,
                uint uType,
                IntPtr pReserved,
                [MarshalAs(UnmanagedType.LPStr)] out string pszName,
                uint cchMax);

            [PreserveSig]
            new int QueryInterface(ref Guid riid, out IntPtr ppvObject);

            [PreserveSig]
            new int AddRef();

            [PreserveSig]
            new int Release();

            [PreserveSig]
            int HandleMenuMsg(uint uMsg, IntPtr wParam, IntPtr lParam);
        }

        [ComImport]
        [Guid("BCFCE0A0-EC17-11D0-8D10-00A0C90F2719")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IContextMenu3 : IContextMenu2
        {
            [PreserveSig]
            new int QueryContextMenu(
                IntPtr hMenu,
                uint indexMenu,
                uint idCmdFirst,
                uint idCmdLast,
                uint uFlags);

            [PreserveSig]
            new int InvokeCommand(ref CMINVOKECOMMANDINFOEX lpici);

            [PreserveSig]
            new int GetCommandString(
                uint idCmd,
                uint uType,
                IntPtr pReserved,
                [MarshalAs(UnmanagedType.LPStr)] out string pszName,
                uint cchMax);

            [PreserveSig]
            new int QueryInterface(ref Guid riid, out IntPtr ppvObject);

            [PreserveSig]
            new int AddRef();

            [PreserveSig]
            new int Release();

            [PreserveSig]
            new int HandleMenuMsg(uint uMsg, IntPtr wParam, IntPtr lParam);

            [PreserveSig]
            int HandleMenuMsg2(uint uMsg, IntPtr wParam, IntPtr lParam, ref IntPtr lpResult);
        }

        [ComImport]
        [Guid("000214E6-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellFolder
        {
            void ParseDisplayName(
                IntPtr hwnd,
                IntPtr pbc,
                [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName,
                out uint pchEaten,
                out IntPtr ppidl,
                ref uint pdwAttributes);

            void EnumObjects(
                IntPtr hwnd,
                uint grfFlags,
                out IntPtr enumIDList);

            void BindToObject(
                IntPtr pidl,
                IntPtr pbc,
                ref Guid riid,
                out IntPtr ppv);

            void BindToStorage(
                IntPtr pidl,
                IntPtr pbc,
                ref Guid riid,
                out IntPtr ppv);

            void CompareIDs(
                IntPtr lParam,
                IntPtr pidl1,
                IntPtr pidl2);

            void CreateViewObject(
                IntPtr hwndOwner,
                ref Guid riid,
                out IntPtr ppv);

            void GetAttributesOf(
                uint cidl,
                IntPtr apidl,
                ref uint rgfInOut);

            void GetUIObjectOf(
                IntPtr hwndOwner,
                uint cidl,
                IntPtr apidl,
                ref Guid riid,
                IntPtr rgfReserved,
                out IntPtr ppv);

            void GetDisplayNameOf(
                IntPtr pidl,
                uint uFlags,
                out IntPtr pName);

            void SetNameOf(
                IntPtr hwnd,
                IntPtr pidl,
                [MarshalAs(UnmanagedType.LPWStr)] string pszName,
                uint uFlags,
                out IntPtr ppidlOut);
        }

        #endregion

        #region 结构体定义

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct CMINVOKECOMMANDINFOEX
        {
            public int cbSize;
            public int fMask;
            public IntPtr hwnd;
            public IntPtr lpVerb;
            [MarshalAs(UnmanagedType.LPStr)] public string lpParameters;
            [MarshalAs(UnmanagedType.LPStr)] public string lpDirectory;
            public int nShow;
            public int dwHotKey;
            public IntPtr hIcon;
            [MarshalAs(UnmanagedType.LPStr)] public string lpTitle;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpVerbW;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpParametersW;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpDirectoryW;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpTitleW;
            public POINT ptInvoke;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        #endregion

        #region 菜单消息处理窗口

        private class MenuHostWindow : NativeWindow
        {
            private readonly IContextMenu2 _contextMenu2;
            private readonly IContextMenu3 _contextMenu3;

            public MenuHostWindow(IntPtr ownerHandle, IContextMenu2 contextMenu2, IContextMenu3 contextMenu3)
            {
                _contextMenu2 = contextMenu2;
                _contextMenu3 = contextMenu3;
                AssignHandle(ownerHandle);
            }

            public void ClearHandlers()
            {
                ReleaseHandle();
            }

            protected override void WndProc(ref Message m)
            {
                try
                {
                    bool handled = false;

                    if (_contextMenu3 != null)
                    {
                        if (m.Msg == WM_INITMENUPOPUP || m.Msg == WM_DRAWITEM || m.Msg == WM_MEASUREITEM || m.Msg == WM_MENUCHAR)
                        {
                            IntPtr result = IntPtr.Zero;
                            int hr = _contextMenu3.HandleMenuMsg2((uint)m.Msg, m.WParam, m.LParam, ref result);
                            if (hr == 0)
                            {
                                m.Result = result;
                                handled = true;
                            }
                        }
                    }
                    else if (_contextMenu2 != null)
                    {
                        if (m.Msg == WM_INITMENUPOPUP || m.Msg == WM_DRAWITEM || m.Msg == WM_MEASUREITEM || m.Msg == WM_MENUCHAR)
                        {
                            int hr = _contextMenu2.HandleMenuMsg((uint)m.Msg, m.WParam, m.LParam);
                            if (hr == 0)
                                handled = true;
                        }
                    }

                    if (!handled)
                    {
                        base.WndProc(ref m);
                    }
                }
                catch
                {
                    base.WndProc(ref m);
                }
            }
        }

        #endregion
    }
}