using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
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
        private const int CMF_EXPLORE = 0x00000004;
        private const int CMIC_MASK_UNICODE = 0x00004000;
        private const uint MF_BYPOSITION = 0x00000400;
        private const uint MF_POPUP = 0x00000010;
        private const uint ShellNewMenuIdFirst = 20000;
        private static readonly Guid CLSID_NewMenu = new("D969A300-E7FF-11D0-A93B-00A0C90F2719");
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
            System.Diagnostics.Debug.WriteLine($"WindowsShellContextMenu.Show: path='{path}', screenX={screenX}, screenY={screenY}");
            
            if (control == null || control.IsDisposed || !control.IsHandleCreated)
            {
                System.Diagnostics.Debug.WriteLine("控制无效或已释放");
                return;
            }

            bool isFile = File.Exists(path);
            bool isDir = Directory.Exists(path);
            System.Diagnostics.Debug.WriteLine($"路径存在检查: isFile={isFile}, isDir={isDir}");
            
            if (!isFile && !isDir)
            {
                System.Diagnostics.Debug.WriteLine($"路径不存在: {path}");
                return;
            }

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

                // 获取父文件夹和相对PIDL
                // 分离路径获取父目录
                string parentPath = Path.GetDirectoryName(path);
                string fileName = Path.GetFileName(path);
                
                System.Diagnostics.Debug.WriteLine($"parentPath='{parentPath}', fileName='{fileName}'");
                
                // 处理根目录情况，如 "C:\"
                if (string.IsNullOrEmpty(parentPath) && isDir)
                {
                    // 对于根目录，直接使用桌面文件夹，并传入驱动器名称
                    parentFolder = desktopFolder;
                    fileName = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    System.Diagnostics.Debug.WriteLine($"根目录处理: fileName='{fileName}'");
                }
                else if (string.IsNullOrEmpty(parentPath))
                {
                    // 如果是根目录，直接使用桌面文件夹
                    parentFolder = desktopFolder;
                }
                else
                {
                    // 解析父目录路径
                    IntPtr parentPidl = IntPtr.Zero;
                    uint pchEaten = 0;
                    uint pdwAttributes = 0;
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
                {
                    uint pchEaten = 0;
                    uint pdwAttributes = 0;
                    parentFolder.ParseDisplayName(control.Handle, IntPtr.Zero, fileName, out pchEaten, out childPidl, ref pdwAttributes);
                }
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

                    uint commandId;
                    try
                    {
                        // 显示菜单
                        System.Diagnostics.Debug.WriteLine($"显示菜单在位置: {screenX}, {screenY}");
                        commandId = User32.TrackPopupMenuEx(
                            hMenu,
                            TPM_RETURNCMD | TPM_RIGHTBUTTON,
                            screenX,
                            screenY,
                            control.Handle,
                            IntPtr.Zero);
                    }
                    finally
                    {
                        // 无论 TrackPopupMenuEx 是否异常，都必须解除子类化，
                        // 否则控件的窗口过程会一直挂着已失效的 IContextMenu 指针
                        window.ClearHandlers();
                    }

                    System.Diagnostics.Debug.WriteLine($"TrackPopupMenuEx 返回命令ID: {commandId}");

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
        /// 显示文件夹**背景**右键菜单（空白处 / 列头）。与选中文件、文件夹的菜单不同，
        /// 这里走 IShellFolder::CreateViewObject 拿 ShellView 的 IContextMenu，才会出现
        /// “新建 / 粘贴 / 查看 / 排序 / 刷新 / 属性”这些资源管理器背景菜单项。
        /// </summary>
        public static void ShowBackgroundAtCursor(string folderPath, Control control)
        {
            var cursor = Cursor.Position;
            ShowBackground(folderPath, control, cursor.X, cursor.Y);
        }

        public static void ShowBackground(string folderPath, Control control, int screenX, int screenY)
        {
            if (control == null || control.IsDisposed || !control.IsHandleCreated) return;
            if (!Directory.Exists(folderPath) || !IsWindows()) return;

            bool comInitialized = false;
            IntPtr pidl = IntPtr.Zero;
            IShellFolder? desktopFolder = null;
            IShellFolder? viewFolder = null;
            IContextMenu? contextMenu = null;
            IContextMenu2? contextMenu2 = null;
            IContextMenu3? contextMenu3 = null;
            ShellNewMenu? shellNewMenu = null;
            IntPtr hMenu = IntPtr.Zero;

            try
            {
                try
                {
                    int hr = Ole32.CoInitializeEx(IntPtr.Zero, Ole32.COINIT_APARTMENTTHREADED);
                    comInitialized = hr == 0 || hr == 1;
                }
                catch
                {
                    comInitialized = false;
                }

                if (Shell32.SHGetDesktopFolder(out IntPtr desktopPtr) != 0 || desktopPtr == IntPtr.Zero)
                    return;

                desktopFolder = (IShellFolder)Marshal.GetObjectForIUnknown(desktopPtr);
                Marshal.Release(desktopPtr);

                uint attributes = 0;
                if (Shell32.SHParseDisplayName(folderPath, IntPtr.Zero, out pidl, 0, out attributes) != 0
                    || pidl == IntPtr.Zero)
                    return;

                Guid shellFolderGuid = IID_IShellFolder;
                desktopFolder.BindToObject(pidl, IntPtr.Zero, ref shellFolderGuid, out IntPtr viewFolderPtr);
                if (viewFolderPtr == IntPtr.Zero) return;

                viewFolder = (IShellFolder)Marshal.GetObjectForIUnknown(viewFolderPtr);
                Marshal.Release(viewFolderPtr);

                Guid contextMenuGuid = IID_IContextMenu;
                viewFolder.CreateViewObject(control.Handle, ref contextMenuGuid, out IntPtr contextMenuPtr);
                if (contextMenuPtr == IntPtr.Zero) return;

                contextMenu = (IContextMenu)Marshal.GetObjectForIUnknown(contextMenuPtr);
                Marshal.Release(contextMenuPtr);

                IntPtr contextMenu3Ptr;
                Guid contextMenu3Guid = IID_IContextMenu3;
                if (contextMenu.QueryInterface(ref contextMenu3Guid, out contextMenu3Ptr) == 0)
                {
                    contextMenu3 = (IContextMenu3)Marshal.GetObjectForIUnknown(contextMenu3Ptr);
                    Marshal.Release(contextMenu3Ptr);
                    contextMenu2 = contextMenu3;
                }
                else
                {
                    IntPtr contextMenu2Ptr;
                    Guid contextMenu2Guid = IID_IContextMenu2;
                    if (contextMenu.QueryInterface(ref contextMenu2Guid, out contextMenu2Ptr) == 0)
                    {
                        contextMenu2 = (IContextMenu2)Marshal.GetObjectForIUnknown(contextMenu2Ptr);
                        Marshal.Release(contextMenu2Ptr);
                    }
                }

                hMenu = User32.CreatePopupMenu();
                if (hMenu == IntPtr.Zero) return;

                int queryResult = contextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF,
                    CMF_NORMAL | CMF_EXPLORE);
                OperationLog.Write("background-menu query=" + queryResult + " folder=" + folderPath);
                if (queryResult < 0) return;

                // 系统背景菜单里的“新建”子菜单在第三方宿主中为空。
                // 单独实例化 Shell 的 New 处理程序，把真正有内容的子菜单换进去。
                shellNewMenu = ShellNewMenu.Create(folderPath);
                if (shellNewMenu != null)
                {
                    int newIndex = FindMenuItemPosition(hMenu, "新建");
                    if (newIndex >= 0) User32.RemoveMenu(hMenu, (uint)newIndex, MF_BYPOSITION);

                    var newSubmenu = shellNewMenu.DetachMenu();
                    if (newIndex >= 0)
                        User32.InsertMenu(hMenu, (uint)newIndex, MF_BYPOSITION | MF_POPUP, newSubmenu, "新建");
                    else
                        User32.AppendMenu(hMenu, MF_POPUP, newSubmenu, "新建");
                    OperationLog.Write("background-menu new-items=" + shellNewMenu.Count);
                }

                // 背景菜单用独立隐藏窗口做宿主：ListView 自己的 WndProc 会吞掉
                // WM_INITMENUPOPUP，导致“新建”子菜单没人填充。
                var window = new MenuHostWindow(contextMenu2, contextMenu3);
                uint commandId;
                try
                {
                    commandId = User32.TrackPopupMenuEx(
                        hMenu,
                        TPM_RETURNCMD | TPM_RIGHTBUTTON,
                        screenX,
                        screenY,
                        window.Handle,
                        IntPtr.Zero);
                }
                finally
                {
                    window.ClearHandlers();
                }

                OperationLog.Write("background-menu command=" + commandId);
                if (commandId >= ShellNewMenuIdFirst && shellNewMenu != null)
                {
                    int invokeHr = shellNewMenu.Invoke(commandId, control.Handle, screenX, screenY, folderPath);
                    OperationLog.Write("background-menu new-invoke hr=" + invokeHr);
                }
                else if (commandId > 0)
                {
                    var invokeInfo = new CMINVOKECOMMANDINFOEX
                    {
                        cbSize = Marshal.SizeOf(typeof(CMINVOKECOMMANDINFOEX)),
                        fMask = CMIC_MASK_UNICODE,
                        hwnd = control.Handle,
                        lpVerbW = (IntPtr)(commandId - 1),
                        lpDirectoryW = folderPath,
                        nShow = 1,
                        ptInvoke = new POINT { x = screenX, y = screenY }
                    };
                    int invokeHr = contextMenu.InvokeCommand(ref invokeInfo);
                    OperationLog.Write("background-menu invoke hr=" + invokeHr);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("显示文件夹背景菜单失败: " + ex);
            }
            finally
            {
                if (hMenu != IntPtr.Zero)
                    User32.DestroyMenu(hMenu);

                if (contextMenu3 != null)
                    Marshal.ReleaseComObject(contextMenu3);

                if (contextMenu2 != null && contextMenu2 != contextMenu3)
                    Marshal.ReleaseComObject(contextMenu2);

                if (contextMenu != null)
                    Marshal.ReleaseComObject(contextMenu);

                shellNewMenu?.Dispose();

                if (viewFolder != null)
                    Marshal.ReleaseComObject(viewFolder);

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

        private static int FindMenuItemPosition(IntPtr hMenu, string prefix)
        {
            int count = User32.GetMenuItemCount(hMenu);
            for (uint i = 0; i < (uint)count; i++)
            {
                var buffer = new StringBuilder(128);
                User32.GetMenuString(hMenu, i, buffer, buffer.Capacity, MF_BYPOSITION);
                var text = buffer.ToString().Replace("&", "");
                if (text.StartsWith(prefix, StringComparison.Ordinal)) return (int)i;
            }
            return -1;
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

            [DllImport("user32.dll")]
            public static extern int GetMenuItemCount(IntPtr hMenu);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            public static extern int GetMenuString(IntPtr hMenu, uint uIDItem, StringBuilder lpString, int cchMax, uint uFlags);

            [DllImport("user32.dll")]
            public static extern bool RemoveMenu(IntPtr hMenu, uint uPosition, uint uFlags);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            public static extern bool AppendMenu(IntPtr hMenu, uint uFlags, IntPtr uIDNewItem, string lpNewItem);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            public static extern bool InsertMenu(IntPtr hMenu, uint uPosition, uint uFlags, IntPtr uIDNewItem, string lpNewItem);
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
        [Guid("000214E8-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellExtInit
        {
            [PreserveSig]
            int Initialize(IntPtr pidlFolder, IntPtr lpdobj, IntPtr hKeyProgID);
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
            public IntPtr lpVerbW;
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

        /// <summary>
        /// 系统“新建”子菜单。直接实例化 Shell 的 New 菜单处理程序（CLSID_NewMenu），
        /// 它会按 HKCR\ShellNew 填出“文件夹 / 文本文档 / Office 文档 / 压缩包”等模板。
        /// 这个处理程序原本是 Explorer 背景菜单里“新建”子菜单的实现。
        /// </summary>
        internal sealed class ShellNewMenu : IDisposable
        {
            private IContextMenu? _contextMenu;
            private IContextMenu2? _contextMenu2;
            private IContextMenu3? _contextMenu3;
            private IntPtr _pidlFolder;
            private bool _comInitialized;

            public IntPtr Menu { get; private set; }
            public int Count { get; private set; }

            public static ShellNewMenu? Create(string folderPath)
            {
                var host = new ShellNewMenu();
                try
                {
                    if (host.Initialize(folderPath)) return host;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("创建系统新建菜单失败: " + ex);
                }

                host.Dispose();
                return null;
            }

            private bool Initialize(string folderPath)
            {
                try
                {
                    int hr = Ole32.CoInitializeEx(IntPtr.Zero, Ole32.COINIT_APARTMENTTHREADED);
                    _comInitialized = hr == 0 || hr == 1;
                }
                catch
                {
                    _comInitialized = false;
                }

                if (Shell32.SHParseDisplayName(folderPath, IntPtr.Zero, out _pidlFolder, 0, out _) != 0
                    || _pidlFolder == IntPtr.Zero)
                    return false;

                var type = Type.GetTypeFromCLSID(CLSID_NewMenu);
                if (type == null) return false;

                object? instance = Activator.CreateInstance(type);
                if (instance is not IShellExtInit shellExtInit) return false;

                if (shellExtInit.Initialize(_pidlFolder, IntPtr.Zero, IntPtr.Zero) != 0)
                    return false;

                _contextMenu = instance as IContextMenu;
                if (_contextMenu == null) return false;

                IntPtr ptr;
                Guid contextMenu3Guid = IID_IContextMenu3;
                if (_contextMenu.QueryInterface(ref contextMenu3Guid, out ptr) == 0)
                {
                    _contextMenu3 = (IContextMenu3)Marshal.GetObjectForIUnknown(ptr);
                    Marshal.Release(ptr);
                    _contextMenu2 = _contextMenu3;
                }
                else
                {
                    Guid contextMenu2Guid = IID_IContextMenu2;
                    if (_contextMenu.QueryInterface(ref contextMenu2Guid, out ptr) == 0)
                    {
                        _contextMenu2 = (IContextMenu2)Marshal.GetObjectForIUnknown(ptr);
                        Marshal.Release(ptr);
                    }
                }

                Menu = User32.CreatePopupMenu();
                if (Menu == IntPtr.Zero) return false;

                Count = _contextMenu.QueryContextMenu(Menu, 0, ShellNewMenuIdFirst,
                    ShellNewMenuIdFirst + 0x7FFF, CMF_NORMAL);
                return Count > 0;
            }

            /// <summary>把子菜单句柄交给父菜单，之后由父菜单负责销毁。</summary>
            public IntPtr DetachMenu()
            {
                var menu = Menu;
                Menu = IntPtr.Zero;
                return menu;
            }

            public int Invoke(uint commandId, IntPtr hwnd, int screenX, int screenY, string folderPath)
            {
                if (_contextMenu == null || commandId < ShellNewMenuIdFirst) return -1;

                var invokeInfo = new CMINVOKECOMMANDINFOEX
                {
                    cbSize = Marshal.SizeOf(typeof(CMINVOKECOMMANDINFOEX)),
                    fMask = CMIC_MASK_UNICODE,
                    hwnd = hwnd,
                    lpVerbW = (IntPtr)(commandId - ShellNewMenuIdFirst),
                    lpDirectoryW = folderPath,
                    nShow = 1,
                    ptInvoke = new POINT { x = screenX, y = screenY }
                };
                return _contextMenu.InvokeCommand(ref invokeInfo);
            }

            public void Dispose()
            {
                if (Menu != IntPtr.Zero)
                {
                    User32.DestroyMenu(Menu);
                    Menu = IntPtr.Zero;
                }

                if (_contextMenu3 != null) Marshal.ReleaseComObject(_contextMenu3);
                if (_contextMenu2 != null && _contextMenu2 != _contextMenu3) Marshal.ReleaseComObject(_contextMenu2);
                if (_contextMenu != null) Marshal.ReleaseComObject(_contextMenu);

                if (_pidlFolder != IntPtr.Zero)
                {
                    Ole32.CoTaskMemFree(_pidlFolder);
                    _pidlFolder = IntPtr.Zero;
                }

                if (_comInitialized)
                {
                    try { Ole32.CoUninitialize(); } catch { }
                    _comInitialized = false;
                }
            }
        }

        #region 菜单消息处理窗口

        private class MenuHostWindow : NativeWindow
        {
            private readonly IContextMenu2? _contextMenu2;
            private readonly IContextMenu3? _contextMenu3;
            private readonly bool _ownWindow;

            public MenuHostWindow(IntPtr ownerHandle, IContextMenu2? contextMenu2, IContextMenu3? contextMenu3)
            {
                _contextMenu2 = contextMenu2;
                _contextMenu3 = contextMenu3;
                AssignHandle(ownerHandle);
            }

            /// <summary>独立隐藏窗口：背景菜单专用，确保 WM_INITMENUPOPUP 由这里处理。</summary>
            public MenuHostWindow(IContextMenu2? contextMenu2, IContextMenu3? contextMenu3)
            {
                _contextMenu2 = contextMenu2;
                _contextMenu3 = contextMenu3;
                _ownWindow = true;
                CreateHandle(new CreateParams
                {
                    Caption = "wxdMenuHost",
                    Style = unchecked((int)0x80000000),
                    ExStyle = 0x00000080,
                    X = 0,
                    Y = 0,
                    Width = 1,
                    Height = 1
                });
            }

            public void ClearHandlers()
            {
                if (_ownWindow) DestroyHandle();
                else ReleaseHandle();
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