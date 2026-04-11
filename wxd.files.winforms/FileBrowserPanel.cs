using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>单个窗格文件夹管理器控件</summary>
    public partial class FileBrowserPanel : UserControl
    {
        // ==================== 控件声明 ====================
        private ComboBox? _cboDrives;
        private Button?   _btnUp;
        private TextBox?  _txtPath;
        private ListView  _lstFiles = new();
        private TextBox?  _txtFilter;
        private Label?    _lblDrives;
        private ImageList _imgList = new();
        private Panel?    _pnlToolbar;
        private Panel?    _pnlBottom;
        private Button?   _btnOpenExplorer;

        // ==================== 状态字段 ====================
        private string          _currentDir     = "";
        private string          _filterText     = "";
        private List<FileInfo>  _allFiles       = new();
        private List<DirInfo>   _allDirs        = new();
        private bool            _suppressEvents = false;

        // ==================== Win32 API ====================
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr SHGetFileInfo(
            string pszPath, uint dwFileAttributes,
            out SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr ILCreateFromPath(string pszPath);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern void ILFree(IntPtr pidl);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr SHBrowseForFolder(ref BROWSEINFO lpbi);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int SHGetPathFromIDList(IntPtr pidl, StringBuilder pszPath);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHParseDisplayName(string pszName, IntPtr pbc, out IntPtr ppidl, uint sfgaoIn, out uint psfgaoOut);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SHBindToParent(IntPtr pidl, [MarshalAs(UnmanagedType.LPStruct)] ref Guid riid, out IntPtr ppv, out IntPtr pidlChild);

        [DllImport("ole32.dll")]
        private static extern int CoCreateInstance([MarshalAs(UnmanagedType.LPStruct)] ref Guid rclsid, IntPtr pUnkOuter, uint dwClsContext, [MarshalAs(UnmanagedType.LPStruct)] ref Guid riid, out IntPtr ppv);

        [DllImport("ole32.dll")]
        private static extern void CoTaskMemFree(IntPtr pv);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int    iIcon;
            public uint   dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHELLEXECUTEINFO
        {
            public int     cbSize;
            public uint    fMask;
            public IntPtr  hwnd;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpVerb;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpFile;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpParameters;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpDirectory;
            public int     nShow;
            public IntPtr  hInstApp;
            public IntPtr  lpIDList;
            [MarshalAs(UnmanagedType.LPTStr)] public string lpClass;
            public IntPtr  hkeyClass;
            public uint    dwHotKey;
            public IntPtr  hIconMonitorOrProcess;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct BROWSEINFO
        {
            public IntPtr hwndOwner;
            public IntPtr pidlRoot;
            public string pszDisplayName;
            public string lpszTitle;
            public uint ulFlags;
            public IntPtr lpfn;
            public IntPtr lParam;
            public int iImage;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CMINVOKECOMMANDINFO
        {
            public int cbSize;
            public int fMask;
            public IntPtr hwnd;
            public IntPtr lpVerb;
            public string lpParameters;
            public string lpDirectory;
            public int nShow;
            public int dwHotKey;
            public IntPtr hIcon;
        }



        // ==================== 常量 ====================
        private const uint SHGFI_ICON              = 0x100;
        private const uint SHGFI_SMALLICON         = 0x000000001;
        private const uint SHGFI_USEFILEATTRIBUTES = 0x10;
        private const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;
        private const uint SEE_MASK_INVOKEIDLIST   = 0xC;
        private const uint SEE_MASK_FLAG_NO_UI     = 0x400;

        // 上下文菜单常量
        private const int GCS_VERBA = 0x0000;
        private const int GCS_VERBW = 0x0001;
        private const int GCS_VALIDATEA = 0x0002;
        private const int GCS_VALIDATEW = 0x0003;
        private const int GCS_HELPTEXTA = 0x0004;
        private const int GCS_HELPTEXTW = 0x0005;
        private const int GCS_VERB = GCS_VERBW;
        private const int GCS_HELPTEXT = GCS_HELPTEXTW;
        private const int CMIC_MASK_HOTKEY = 0x20;
        private const int CMIC_MASK_ICON = 0x10;
        private const int CMIC_MASK_FLAG_NO_UI = 0x400;
        private const int CMIC_MASK_UNICODE = 0x4000;
        private const int CMIC_MASK_NO_CONSOLE = unchecked((int)0x8000);
        private const int CMF_NORMAL = 0x00000000;
        private const int CMF_DEFAULTONLY = 0x00000001;
        private const int CMF_VERBSONLY = 0x00000002;
        private const int CMF_EXPLORE = 0x00000004;
        private const int CMF_NOVERBS = 0x00000008;
        private const int CMF_CANRENAME = 0x00000010;
        private const int CMF_NODEFAULT = 0x00000020;
        private const int CMF_INCLUDESTATIC = 0x00000040;
        private const int CMF_ITEMMENU = 0x00000080;
        private const int CMF_EXTENDEDVERBS = 0x00000100;
        private const int CMF_DISABLEDVERBS = 0x00000200;
        private const int CMF_ASYNCVERBSTATE = 0x00000400;
        private const int CMF_OPTIMIZEFORINVOKE = 0x00000800;
        private const int CMF_SYNCCASCADEMENU = 0x00001000;
        private const int CMF_DONOTPICKDEFAULT = 0x00002000;
        private const int CMF_RESERVED = unchecked((int)0xffff0000);

        // 界面标识符
        private static readonly Guid IID_IContextMenu = new Guid("000214e4-0000-0000-c000-000000000046");
        private static readonly Guid IID_IShellFolder = new Guid("000214e6-0000-0000-c000-000000000046");
        private static readonly Guid CLSID_ShellFolder = new Guid("00021400-0000-0000-c000-000000000046");

        // ==================== 数据类 ====================
        private class DirInfo { public string Name = ""; public string FullName = ""; }
        private class FileItemTag { public string FullName = ""; public bool IsDir; }

        // ==================== 属性 ====================
        public string CurrentDirectory => _currentDir;
        public string FilterText => _filterText;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int PanelIndex { get; set; }

        /// <summary>状态持久化文件路径</summary>
        private static string StateFilePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "wxd.files.winforms", "state.json");

        // ==================== 构造 & 初始化 ====================
        public FileBrowserPanel() : this(-1) { }

        public FileBrowserPanel(int panelIndex)
        {
            PanelIndex = panelIndex;
            InitializeComponent();
            CreateControls();
            InitImageList();
            LoadDrives();
            // 恢复记忆的目录和过滤条件
            var state = LoadPanelState(panelIndex);
            var startPath = state?.Directory ?? "";
            var savedFilter = state?.Filter ?? "";
            if (!string.IsNullOrEmpty(startPath) && Directory.Exists(startPath))
                NavigateTo(startPath, restoreFilter: false);
            else
                NavigateTo(GetFirstDrive(), restoreFilter: false);
            // 恢复过滤条件
            if (!string.IsNullOrEmpty(savedFilter))
            {
                _filterText = savedFilter;
                if (_txtFilter != null)
                {
                    _txtFilter.Text = savedFilter;
                    _txtFilter.ForeColor = SystemColors.WindowText;
                }
                ApplyFilter();
            }
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            // 
            // FileBrowserPanel
            // 
            Name = "FileBrowserPanel";
            Size = new Size(896, 758);
            ResumeLayout(false);
        }

        /// <summary>实际创建和布局所有控件</summary>
        private void CreateControls()
        {
            SuspendLayout();

            BackColor    = Color.White;
            BorderStyle = BorderStyle.FixedSingle;
            Padding      = new Padding(1);

            // ===== 工具栏面板（顶部）=====
            _pnlToolbar = new Panel
            {
                Dock     = DockStyle.Top,
                Height   = 34,
                BackColor = Color.FromArgb(240, 243, 247),
                Padding  = new Padding(5, 2, 5, 2)
            };

            _lblDrives = new Label
            {
                Text      = "磁盘",
                Location  = new Point(6, 7),
                Size      = new Size(30, 20),
                AutoSize  = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Font      = new Font("Microsoft YaHei UI", 8.25f),
                ForeColor = Color.FromArgb(60, 60, 65)
            };

            _cboDrives = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location      = new Point(36, 4),
                Size          = new Size(82, 23),
                Font          = new Font("Microsoft YaHei UI", 8.25f)
            };
            _cboDrives.SelectedIndexChanged += OnDriveChanged;

            _btnUp = new Button
            {
                Text     = "↑ 上级",
                Location = new Point(124, 4),
                Size     = new Size(58, 24),
                Font     = new Font("Microsoft YaHei UI", 8.25f),
                Cursor    = Cursors.Hand
            };
            _btnUp.Click += (s, e) => GoUp();

            _txtPath = new TextBox
            {
                Location    = new Point(188, 4),
                Size        = new Size(180, 23),
                Font        = new Font("Microsoft YaHei UI", 8.25f),
                Anchor      = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _txtPath.KeyDown += OnPathKeyDown;

            _pnlToolbar.Controls.AddRange(new Control[] { _lblDrives, _cboDrives, _btnUp, _txtPath });

            // ===== 文件列表（中间区域）=====
            _lstFiles.Dock           = DockStyle.Fill;
            _lstFiles.View           = View.Details;
            _lstFiles.FullRowSelect  = true;
            _lstFiles.GridLines      = true;
            _lstFiles.HeaderStyle    = ColumnHeaderStyle.Clickable;
            _lstFiles.SmallImageList = _imgList;
            _lstFiles.Font           = new Font("Microsoft YaHei UI", 9f);
            _lstFiles.MultiSelect    = false;  // 禁止多选，避免右键多选行为

            _lstFiles.Columns.Add("名称", 220, HorizontalAlignment.Left);
            _lstFiles.Columns.Add("修改日期", 145, HorizontalAlignment.Left);
            _lstFiles.Columns.Add("类型", 110, HorizontalAlignment.Left);
            _lstFiles.Columns.Add("大小", 70, HorizontalAlignment.Right);

            _lstFiles.DoubleClick         += OnFileDoubleClick;
            _lstFiles.MouseClick          += OnListViewMouseClick;
            _lstFiles.SelectedIndexChanged += (s, e) =>
            {
                if (_lstFiles.SelectedItems.Count > 0) _lstFiles.Focus();
            };

            // ===== 底部面板（包含过滤框和打开资源管理器按钮）=====
            _pnlBottom = new Panel
            {
                Height = 28,
                Dock   = DockStyle.Bottom
            };

            // 打开资源管理器按钮
            _btnOpenExplorer = new Button
            {
                Width       = 80,
                Height      = 28,
                Dock        = DockStyle.Right,
                Text        = "打开目录",
                Font        = new Font("Microsoft YaHei UI", 8.25f),
                FlatStyle   = FlatStyle.Flat,
                BackColor   = Color.FromArgb(240, 240, 242),
                Cursor      = Cursors.Hand
            };
            _btnOpenExplorer.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            _btnOpenExplorer.Click += OnOpenExplorerClick;

            // 过滤框
            _txtFilter = new TextBox
            {
                Height      = 28,
                Dock        = DockStyle.Fill,
                Font        = new Font("Microsoft YaHei UI", 8.25f),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor   = Color.FromArgb(250, 250, 252)
            };

            SetPlaceholder();
            _txtFilter.TextChanged += OnFilterTextChanged;
            _txtFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    _txtFilter.Text = "";
                    _filterText = "";
                    ApplyFilter();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };

            _pnlBottom.Controls.Add(_txtFilter);
            _pnlBottom.Controls.Add(_btnOpenExplorer);

            Controls.Add(_lstFiles);
            Controls.Add(_pnlBottom);
            Controls.Add(_pnlToolbar);

            MinimumSize = new Size(280, 280);
            ResumeLayout(false);
        }

        private void SetPlaceholder()
        {
            if (_txtFilter == null) return;
            _txtFilter.Text = "输入关键词过滤 (支持 || 或 / && 与)";
            _txtFilter.ForeColor = SystemColors.GrayText;

            _txtFilter.GotFocus -= Filter_GotFocus;
            _txtFilter.GotFocus += Filter_GotFocus;
            _txtFilter.LostFocus -= Filter_LostFocus;
            _txtFilter.LostFocus += Filter_LostFocus;
        }

        private void Filter_GotFocus(object? sender, EventArgs e)
        {
            if (_txtFilter != null && _txtFilter.Text.StartsWith("输入关键词"))
            {
                _txtFilter.Text = "";
                _txtFilter.ForeColor = SystemColors.WindowText;
            }
        }

        private void Filter_LostFocus(object? sender, EventArgs e)
        {
            if (_txtFilter != null && string.IsNullOrWhiteSpace(_txtFilter.Text))
                SetPlaceholder();
        }

        // ==================== 图标列表 ====================
        private void InitImageList()
        {
            _imgList.ImageSize = new Size(16, 16);
            _imgList.ColorDepth = ColorDepth.Depth32Bit;

            var folderShfi = default(SHFILEINFO);
            SHGetFileInfo(@"C:\Windows", FILE_ATTRIBUTE_DIRECTORY, out folderShfi, (uint)Marshal.SizeOf(folderShfi),
                          SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
            _imgList.Images.Add("folder", Icon.FromHandle(folderShfi.hIcon));

            var fileShfi = default(SHFILEINFO);
            SHGetFileInfo(".txt", 0, out fileShfi, (uint)Marshal.SizeOf(fileShfi),
                          SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
            _imgList.Images.Add("file_default", Icon.FromHandle(fileShfi.hIcon));
        }

        // ==================== 磁盘操作 ====================
        private void LoadDrives()
        {
            _suppressEvents = true;
            _cboDrives!.Items.Clear();
            foreach (var drive in DriveInfo.GetDrives().Where(d => d.IsReady))
                _cboDrives.Items.Add(drive.RootDirectory.FullName);  // 保留 "C:\", "D:\"
            _suppressEvents = false;
        }

        private string GetFirstDrive() => _cboDrives!.Items.Count > 0 ? _cboDrives.Items[0].ToString()! : "C:\\";

        private void OnDriveChanged(object? sender, EventArgs e)
        {
            if (_suppressEvents || _cboDrives?.SelectedItem == null) return;
            var drivePath = _cboDrives.SelectedItem.ToString()!;
            if (Directory.Exists(drivePath))
                NavigateTo(drivePath);
        }

        // ==================== 导航操作 ====================
        public void NavigateTo(string path, bool restoreFilter = true)
        {
            try
            {
                if (!string.IsNullOrEmpty(path))
                    path = Path.GetFullPath(path);
                if (!Directory.Exists(path)) return;
                // 保留根目录的尾部反斜杠：C:\ 不能 Trim 成 C:
                var trimmed = path.TrimEnd(Path.DirectorySeparatorChar);
                if (trimmed.Length == 2 && trimmed[1] == ':')
                    _currentDir = trimmed + Path.DirectorySeparatorChar;
                else
                    _currentDir = trimmed;
                RefreshUI();
                SavePanelState();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法访问路径:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void GoUp()
        {
            try
            {
                var parent = Directory.GetParent(_currentDir);
                if (parent != null) NavigateTo(parent.FullName);
            }
            catch { }
        }

        private void OnPathKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                var input = _txtPath?.Text.Trim() ?? "";
                if (!string.IsNullOrEmpty(input))
                {
                    if (!Path.IsPathRooted(input) && !string.IsNullOrEmpty(_currentDir))
                        input = Path.Combine(_currentDir, input);
                    NavigateTo(input);
                }
            }
        }

        // ==================== 刷新 UI ====================
        private void RefreshUI()
        {
            _suppressEvents = true;

            if (_txtPath != null)
                _txtPath.Text = _currentDir;

            // 同步磁盘选择
            int targetIdx = -1;
            for (int i = 0; i < _cboDrives!.Items.Count; i++)
            {
                var d = _cboDrives.Items[i].ToString()!;
                if (_currentDir.Equals(d, StringComparison.OrdinalIgnoreCase) ||
                    (_currentDir.StartsWith(d, StringComparison.OrdinalIgnoreCase) && _currentDir.Length > d.Length))
                {
                    targetIdx = i;
                    break;
                }
            }
            if (targetIdx >= 0 && _cboDrives.SelectedIndex != targetIdx)
                _cboDrives.SelectedIndex = targetIdx;

            LoadFileList();
            ApplyFilter();
            _suppressEvents = false;
        }

        private void LoadFileList()
        {
            _allDirs.Clear();
            _allFiles.Clear();

            try
            {
                foreach (var dir in Directory.GetDirectories(_currentDir))
                {
                    try { _allDirs.Add(new DirInfo { Name = Path.GetFileName(dir), FullName = dir }); } catch { }
                }
                _allDirs.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));

                foreach (var file in Directory.GetFiles(_currentDir))
                {
                    try { _allFiles.Add(new FileInfo(file)); } catch { }
                }
                _allFiles.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
            }
            catch (UnauthorizedAccessException) { }
        }

        // ==================== 打开资源管理器 ====================
        private void OnOpenExplorerClick(object? sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
            {
                try
                {
                    Process.Start("explorer.exe", _currentDir);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法打开资源管理器: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // ==================== 过滤逻辑 ====================
        private void OnFilterTextChanged(object? sender, EventArgs e)
        {
            if (_suppressEvents || _txtFilter == null) return;
            var txt = _txtFilter.Text;
            if (txt.StartsWith("输入关键词"))
            {
                _filterText = "";
                ApplyFilter();
                SavePanelState();
                return;
            }
            _filterText = txt ?? "";
            ApplyFilter();
            SavePanelState();
        }

        private void ApplyFilter()
        {
            _lstFiles.BeginUpdate();
            _lstFiles.Items.Clear();

            foreach (var dir in FilterItems(_allDirs.Select(d => d.Name)))
            {
                var info = _allDirs.First(d => d.Name.Equals(dir, StringComparison.OrdinalIgnoreCase));
                AddListItem(info.Name, info.FullName, isDir: true);
            }

            foreach (var fn in FilterItems(_allFiles.Select(f => f.Name)))
            {
                var fi = _allFiles.First(f => f.Name.Equals(fn, StringComparison.OrdinalIgnoreCase));
                AddListItem(fi.Name, fi.FullName, isDir: false, fi.LastWriteTime, fi.Length);
            }

            AutoResizeColumns();
            _lstFiles.EndUpdate();
        }

        private IEnumerable<string> FilterItems(IEnumerable<string> names)
        {
            var filter = _filterText.Trim();
            if (string.IsNullOrEmpty(filter)) return names;

            filter = filter.Replace("||", " | ").Replace("&&", " & ");

            if (filter.Contains('|'))
            {
                var terms = filter.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                return names.Where(n => terms.Any(t => n.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0));
            }
            else if (filter.Contains('&'))
            {
                var terms = filter.Split('&', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                return names.Where(n => terms.All(t => n.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0));
            }
            else
            {
                return names.Where(n => n.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        // ==================== 列表项操作 ====================
        private void AddListItem(string name, string fullName, bool isDir,
                                DateTime lastWrite = default, long size = 0)
        {
            var item = new ListViewItem(name);
            item.Tag = new FileItemTag { FullName = fullName, IsDir = isDir };

            if (isDir)
            {
                item.ImageKey = "folder";
                item.SubItems.Add("");
                item.SubItems.Add("文件夹");
                item.SubItems.Add("");
            }
            else
            {
                item.ImageKey = GetFileIconKey(fullName);
                item.SubItems.Add(lastWrite == default ? "" : lastWrite.ToString("yyyy-MM-dd HH:mm"));
                item.SubItems.Add(GetFileType(fullName));
                item.SubItems.Add(size > 0 ? FormatFileSize(size) : "");
            }

            _lstFiles.Items.Add(item);
        }

        private string GetFileIconKey(string filePath)
        {
            try
            {
                var key = Path.GetExtension(filePath).ToLowerInvariant();
                if (!_imgList.Images.ContainsKey(key))
                {
                    var shfi = default(SHFILEINFO);
                    SHGetFileInfo(filePath, 0, out shfi, (uint)Marshal.SizeOf(shfi), SHGFI_ICON | SHGFI_SMALLICON);
                    if (shfi.hIcon != IntPtr.Zero)
                    {
                        _imgList.Images.Add(key, Icon.FromHandle(shfi.hIcon));
                        return key;
                    }
                }
                else return key;
            }
            catch { }
            return "file_default";
        }

        private string GetFileType(string filePath)
        {
            try
            {
                var shfi = default(SHFILEINFO);
                SHGetFileInfo(filePath, 0, out shfi, (uint)Marshal.SizeOf(shfi), 0x400);
                if (!string.IsNullOrEmpty(shfi.szTypeName)) return shfi.szTypeName;
            }
            catch { }
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            return string.IsNullOrEmpty(ext) ? "文件" : ext.ToUpperInvariant().TrimStart('.') + " 文件";
        }

        private static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("F1") + " KB";
            if (bytes < 1024 * 1024 * 1024) return (bytes / (1024.0 * 1024)).ToString("F1") + " MB";
            return (bytes / (1024.0 * 1024 * 1024)).ToString("F2") + " GB";
        }

        private void AutoResizeColumns()
        {
            for (int i = 0; i < _lstFiles.Columns.Count; i++)
                _lstFiles.Columns[i].Width = -2;
        }

        // ==================== 双击 & 右键菜单 ====================
        private void OnFileDoubleClick(object? sender, EventArgs e)
        {
            if (_lstFiles.SelectedItems.Count == 0) return;
            var tag = _lstFiles.SelectedItems[0].Tag as FileItemTag;
            if (tag == null) return;

            if (tag.IsDir)
                NavigateTo(tag.FullName);
            else
                OpenWithShell(tag.FullName, "open");
        }

        /// <summary>处理 ListView 鼠标点击事件（用于右键菜单）</summary>
        private void OnListViewMouseClick(object? sender, MouseEventArgs e)
        {
            // 检查控件状态，防止在已释放状态下处理事件
            if (this.IsDisposed || !this.IsHandleCreated) return;
            
            if (e.Button != MouseButtons.Right) return;

            // 添加调试信息
            Debug.WriteLine($"右键点击事件触发: X={e.X}, Y={e.Y}");

            // 检测列头区域 - 列名上右键点击显示系统右键菜单
            if (IsPointOnColumnHeader(e.Location))
            {
                Debug.WriteLine("点击列头区域");
                ShowSystemContextMenu(_currentDir);
                return;
            }

            var hit = _lstFiles.HitTest(e.Location);
            if (hit.Item != null)
            {
                Debug.WriteLine($"点击列表项: {hit.Item.Text}");
                hit.Item.Selected = true;
                ShowSystemContextMenu(hit.Item.Tag is FileItemTag t ? t.FullName : _currentDir);
            }
            else
            {
                Debug.WriteLine("点击空白区域");
                ShowSystemContextMenu(_currentDir);
            }
        }

        /// <summary>检测点击位置是否在列头上</summary>
        private bool IsPointOnColumnHeader(Point pt)
        {
            if (_lstFiles.HeaderStyle != ColumnHeaderStyle.None && _lstFiles.Columns.Count > 0)
            {
                int headerHeight = GetListViewHeaderHeight();
                return pt.Y <= headerHeight;
            }
            return false;
        }

        /// <summary>获取 ListView 列头的精确高度</summary>
        private int GetListViewHeaderHeight()
        {
            const int LVM_GETHEADER = 0x101F;
            IntPtr hHeader = SendMessage(_lstFiles.Handle, LVM_GETHEADER, IntPtr.Zero, IntPtr.Zero);
            if (hHeader == IntPtr.Zero) return 24;

            const int WM_GETFONT = 0x31;
            IntPtr hFont = SendMessage(hHeader, WM_GETFONT, IntPtr.Zero, IntPtr.Zero);
            using var font = Font.FromHfont(hFont);
            return font.Height + 8;
        }

        /// <summary>显示系统右键上下文菜单</summary>
        private void ShowSystemContextMenu(string itemPath)
        {
            // 检查控件状态，防止在已释放状态下处理事件
            if (this.IsDisposed || !this.IsHandleCreated) return;
            
            try
            {
                // 使用 Windows Shell API 显示系统右键菜单
                WindowsShellContextMenu.ShowAtCursor(itemPath, this);
            }
            catch (Exception ex)
            {
                // 如果失败，使用备用菜单
                Debug.WriteLine($"显示系统右键菜单失败: {ex.Message}");
                ShowFallbackContextMenu(itemPath);
            }
        }
        
        /// <summary>显示备用右键菜单</summary>
        private void ShowFallbackContextMenu(string itemPath)
        {
            try
            {
                if (!File.Exists(itemPath) && !Directory.Exists(itemPath)) return;
                
                var menu = new ContextMenuStrip();
                var isDir = Directory.Exists(itemPath);
                
                // 基本操作
                menu.Items.Add("打开", null, (s, e) => OpenWithShell(itemPath, "open"));
                if (isDir)
                    menu.Items.Add("在资源管理器中打开", null, (s, e) => OpenWithShell(itemPath, "explore"));
                
                menu.Items.Add("-");
                menu.Items.Add("属性", null, (s, e) => 
                {
                    var info = new SHELLEXECUTEINFO
                    {
                        cbSize = Marshal.SizeOf<SHELLEXECUTEINFO>(),
                        fMask = SEE_MASK_FLAG_NO_UI,
                        hwnd = this.Handle,
                        lpVerb = "properties",
                        lpFile = itemPath,
                        nShow = 1
                    };
                    ShellExecuteEx(ref info);
                });
                menu.Items.Add("复制路径", null, (s, e) => Clipboard.SetText(itemPath));
                
                // 显示菜单
                menu.Show(Cursor.Position);
            }
            catch
            {
                // 如果备用方法也失败，就什么也不做
            }
        }





        private void OpenWithShell(string filePath, string verb)
        {
            try
            {
                if (this.IsDisposed || !this.IsHandleCreated)
                {
                    return;
                }
                
                var info = new SHELLEXECUTEINFO
                {
                    cbSize = Marshal.SizeOf<SHELLEXECUTEINFO>(),
                    fMask  = SEE_MASK_FLAG_NO_UI,
                    hwnd   = this.Handle,
                    lpVerb = verb,
                    lpFile = filePath,
                    nShow  = 1
                };
                ShellExecuteEx(ref info);
            }
            catch (ObjectDisposedException)
            {
                // 忽略已释放的对象异常
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ==================== 公共接口 & 状态持久化 ====================
        public void RefreshCurrent() => RefreshUI();

        private class PanelState
        {
            public string Directory { get; set; } = "";
            public string Filter { get; set; } = "";
            public int PanelCount { get; set; } = 4;
            public int LayoutIndex { get; set; } = 2;
        }

        private static void EnsureStateDir()
        {
            var dir = Path.GetDirectoryName(StateFilePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }

        private void SavePanelState()
        {
            try
            {
                var allStates = LoadAllStates();
                var key = $"panel_{PanelIndex}";
                allStates[key] = new PanelState
                {
                    Directory = _currentDir,
                    Filter = _filterText.StartsWith("输入关键词") ? "" : _filterText
                };
                EnsureStateDir();
                File.WriteAllText(StateFilePath, JsonSerializer.Serialize(allStates, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { }
        }

        private static PanelState? LoadPanelState(int panelIndex)
        {
            try
            {
                var allStates = LoadAllStates();
                var key = $"panel_{panelIndex}";
                return allStates.ContainsKey(key) ? allStates[key] : null;
            }
            catch { return null; }
        }

        internal static void SaveGlobalState(int panelCount, int layoutIndex)
        {
            try
            {
                var allStates = LoadAllStates();
                allStates["global"] = new PanelState
                {
                    PanelCount = panelCount,
                    LayoutIndex = layoutIndex
                };
                EnsureStateDir();
                File.WriteAllText(StateFilePath, JsonSerializer.Serialize(allStates, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { }
        }

        internal static (int PanelCount, int LayoutIndex) LoadGlobalState()
        {
            try
            {
                var allStates = LoadAllStates();
                if (allStates.TryGetValue("global", out var g))
                    return (g.PanelCount, g.LayoutIndex);
            }
            catch { }
            return (4, 2);
        }

        private static Dictionary<string, PanelState> LoadAllStates()
        {
            try
            {
                if (File.Exists(StateFilePath))
                    return JsonSerializer.Deserialize<Dictionary<string, PanelState>>(File.ReadAllText(StateFilePath))
                           ?? new Dictionary<string, PanelState>();
            }
            catch { }
            return new Dictionary<string, PanelState>();
        }

        protected override void Dispose(bool disposing)
        {
            SavePanelState();
            if (disposing) 
            { 
                _imgList?.Dispose(); 
                _lstFiles?.Dispose(); 
            }
            base.Dispose(disposing);
        }
    }

    // Native 接口定义
    internal static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CMINVOKECOMMANDINFOEX
        {
            public int cbSize;
            public int fMask;
            public IntPtr hwnd;
            public IntPtr lpVerb;
            public string lpParameters;
            public string lpDirectory;
            public int nShow;
            public int dwHotKey;
            public IntPtr hIcon;
            public string lpTitle;
            public string lpVerbW;
            public string lpParametersW;
            public string lpDirectoryW;
            public string lpTitleW;
            public POINT ptInvoke;
        }

        [ComImport]
        [Guid("000214e4-0000-0000-c000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IContextMenu
        {
            [PreserveSig]
            int QueryContextMenu(
                IntPtr hMenu,
                uint indexMenu,
                uint idCmdFirst,
                uint idCmdLast,
                uint uFlags);

            [PreserveSig]
            int InvokeCommand(ref CMINVOKECOMMANDINFOEX pici);

            [PreserveSig]
            int GetCommandString(
                uint idCmd,
                uint uFlags,
                IntPtr pReserved,
                [MarshalAs(UnmanagedType.LPStr)] out string pszName,
                uint cchMax);
        }

        [ComImport]
        [Guid("000214e6-0000-0000-c000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IShellFolder
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
    }
}
