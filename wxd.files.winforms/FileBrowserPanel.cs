using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>单个窗格文件夹管理器控件</summary>
    public partial class FileBrowserPanel : UserControl
    {
        // ==================== 视图模式枚举 ====================
        public enum FileListViewMode
        {
            Details,    // 详细信息
            List,       // 列表
            SmallIcon,  // 小图标
            LargeIcon,  // 大图标
            Tile        // 平铺
        }

        // ==================== 控件声明 ====================
        private ComboBox? _cboPath;      // 路径选择下拉框（替代原先的 TextBox 路径 + 磁盘下拉框）
        private Button? _btnUp;
        private ListView _lstFiles = new();
        private TextBox? _txtFilter;
        private ImageList _imgList = new();
        private Panel? _pnlToolbar;
        private Panel? _pnlBottom;
        private Button? _btnOpenExplorer;
        private Button? _btnOpenCsvReader;
        private ComboBox? _cboViewMode;

        // ==================== 状态字段 ====================
        private string _currentDir = "";
        private string _filterText = "";
        private List<FileInfo> _allFiles = new();
        private List<DirInfo> _allDirs = new();
        private bool _suppressEvents = false;
        private bool _isLoading = false;
        private bool _isDark = false;
        private Label? _lblLoading;
        private Label? _lblFileCount;
        private FileListViewMode _viewMode = FileListViewMode.Details;
        private ImageList _imgListLarge = new();
        private ToolTip _pathToolTip = new();
        private System.Windows.Forms.Timer? _filterDebounceTimer;
        private System.Windows.Forms.Timer? _saveDebounceTimer;
        private bool _hasUnsavedChanges = false;
        private string _savedFilterForCompare = "";

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

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
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

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFILEOPSTRUCT
        {
            public IntPtr hwnd;
            public uint wFunc;
            [MarshalAs(UnmanagedType.LPWStr)] public string pFrom;
            [MarshalAs(UnmanagedType.LPWStr)] public string pTo;
            public uint fFlags;
            public int fAnyOperationsAborted;
            public IntPtr hNameMappings;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int SHFileOperation(ref SHFILEOPSTRUCT lpFileOp);



        // ==================== 常量 ====================
        private const uint SHGFI_ICON = 0x100;
        private const uint SHGFI_SMALLICON = 0x000000001;
        private const uint SHGFI_USEFILEATTRIBUTES = 0x10;
        private const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;
        private const uint SEE_MASK_INVOKEIDLIST = 0xC;
        private const uint SEE_MASK_FLAG_NO_UI = 0x400;
        private const uint LVM_ARRANGE = 0x1016;

        // 文件操作常量
        private const uint FO_DELETE = 3;
        private const uint FOF_ALLOWUNDO = 0x40;
        private const uint FOF_NOCONFIRMATION = 0x10;
        private const uint FOF_SIMPLEPROGRESS = 0x100;   // 显示进度对话框
        private const uint FOF_WANTNUKEWARNING = 0x4000; // 删除大文件时显示警告

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

        // ==================== 浅色天蓝配色方案 ====================
        private static readonly Color _colorLightToolbarBg = Color.FromArgb(235, 245, 255);
        private static readonly Color _colorLightBorder = Color.FromArgb(173, 209, 237);
        private static readonly Color _colorLightTextPrimary = Color.FromArgb(20, 60, 100);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(90, 130, 170);
        private static readonly Color _colorLightBackColor = Color.White;

        // ==================== 深色配色方案 (VS风格) ====================
        private static readonly Color _colorDarkToolbarBg = Color.FromArgb(51, 51, 55);       // VS工具栏
        private static readonly Color _colorDarkBorder = Color.FromArgb(62, 62, 66);       // VS边框
        private static readonly Color _colorDarkTextPrimary = Color.FromArgb(241, 241, 241);    // VS白色文字
        private static readonly Color _colorDarkTextSecondary = Color.FromArgb(180, 180, 180);    // VS灰色文字
        private static readonly Color _colorDarkBackColor = Color.FromArgb(37, 37, 38);       // VS编辑器背景

        // 当前配色
        private Color _colorToolbarBg;
        private Color _colorBorder;
        private Color _colorTextPrimary;
        private Color _colorTextSecondary;

        // ==================== 数据类 ====================
        private class DirInfo { public string Name = ""; public string FullName = ""; }
        private class FileItemTag { public string FullName = ""; public bool IsDir; }
        /// <summary>下拉选项中的分隔线占位</summary>
        private class SeparatorItem
        {
            public override string ToString() => "──────────";
        }

        // ==================== 自然排序比较器 ====================
        /// <summary>按字符逐个比较，数字部分按数值比较</summary>
        private class NaturalSortComparer : IComparer<string>
        {
            public int Compare(string? x, string? y)
            {
                if (x == null && y == null) return 0;
                if (x == null) return -1;
                if (y == null) return 1;

                int i = 0, j = 0;
                while (i < x.Length && j < y.Length)
                {
                    char cx = x[i];
                    char cy = y[j];

                    // 如果都是数字，按数值比较
                    if (char.IsDigit(cx) && char.IsDigit(cy))
                    {
                        // 提取完整的数字
                        int numX = 0, numY = 0;
                        while (i < x.Length && char.IsDigit(x[i]))
                        {
                            numX = numX * 10 + (x[i] - '0');
                            i++;
                        }
                        while (j < y.Length && char.IsDigit(y[j]))
                        {
                            numY = numY * 10 + (y[j] - '0');
                            j++;
                        }
                        if (numX != numY)
                            return numX.CompareTo(numY);
                    }
                    else
                    {
                        // 非数字字符，忽略大小写比较
                        int result = char.ToUpperInvariant(cx).CompareTo(char.ToUpperInvariant(cy));
                        if (result != 0)
                            return result;
                        i++;
                        j++;
                    }
                }

                // 比较剩余长度
                return x.Length.CompareTo(y.Length);
            }
        }

        // ==================== 属性 ====================
        public string CurrentDirectory => _currentDir;
        public string FilterText => _filterText;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int PanelIndex { get; set; }

        // ==================== 构造 & 初始化 ====================
        public FileBrowserPanel() : this(-1) { }

        public FileBrowserPanel(int panelIndex) : this(panelIndex, skipAutoNavigate: false) { }

        /// <summary>
        /// 完整构造函数
        /// </summary>
        /// <param name="panelIndex">面板索引</param>
        /// <param name="skipAutoNavigate">是否跳过自动导航到上次目录（用于启动优化，先显示UI再加载数据）</param>
        public FileBrowserPanel(int panelIndex, bool skipAutoNavigate)
        {
            PanelIndex = panelIndex;
            // 默认使用浅色主题
            SetThemeColors(false);
            InitializeComponent();
            InitImageList();
            CreateControls();
            RebuildPathComboItems();

            // ==== 第一步：恢复视图模式 ====
            var savedViewMode = CsvSettings.GetPaneViewMode(panelIndex);
            _viewMode = (FileListViewMode)Math.Max(0, Math.Min(4, savedViewMode));
            // 面板句柄创建后延迟恢复视图（确保子控件句柄也已就绪）
            HandleCreated += (_, _) => BeginInvoke(PopulateListView);
            _suppressEvents = true;
            if (_cboViewMode != null)
                _cboViewMode.SelectedIndex = (int)_viewMode;
            _suppressEvents = false;

            // ==== 第二步：恢复过滤关键词 ====
            var savedFilter = CsvSettings.GetPaneFilter(panelIndex);
            if (!string.IsNullOrEmpty(savedFilter))
            {
                _filterText = savedFilter;
                if (_txtFilter != null)
                {
                    _txtFilter.Text = savedFilter;
                    _txtFilter.ForeColor = _colorTextPrimary;
                }
            }

            // ==== 第三步：导航到上次目录（可被跳过） ====
            if (!skipAutoNavigate)
            {
                var startPath = CsvSettings.GetPaneDirectory(panelIndex);
                if (!string.IsNullOrEmpty(startPath) && Directory.Exists(startPath))
                    NavigateTo(startPath, restoreFilter: false, saveState: false);
                else
                    NavigateTo(GetFirstDrive(), restoreFilter: false, saveState: false);
            }
        }

        /// <summary>
        /// 延迟初始化：在面板UI已显示后，异步恢复导航到保存的目录。
        /// 配合 skipAutoNavigate=true 使用，实现启动时先显示主界面再加载数据的效果。
        /// </summary>
        public void InitializeDeferred()
        {
            var startPath = CsvSettings.GetPaneDirectory(PanelIndex);
            if (!string.IsNullOrEmpty(startPath) && Directory.Exists(startPath))
                NavigateTo(startPath, restoreFilter: false, saveState: false);
            else
                NavigateTo(GetFirstDrive(), restoreFilter: false, saveState: false);
        }

        /// <summary>设置主题颜色</summary>
        private void SetThemeColors(bool isDark)
        {
            if (isDark)
            {
                _colorToolbarBg = _colorDarkToolbarBg;
                _colorBorder = _colorDarkBorder;
                _colorTextPrimary = _colorDarkTextPrimary;
                _colorTextSecondary = _colorDarkTextSecondary;
            }
            else
            {
                _colorToolbarBg = _colorLightToolbarBg;
                _colorBorder = _colorLightBorder;
                _colorTextPrimary = _colorLightTextPrimary;
                _colorTextSecondary = _colorLightTextSecondary;
            }
        }

        /// <summary>应用主题（公共方法，由主窗体调用）</summary>
        public void ApplyTheme(bool isDark)
        {
            _isDark = isDark;
            SetThemeColors(isDark);

            // 更新背景色
            BackColor = isDark ? _colorDarkBackColor : _colorLightBackColor;

            // 更新工具栏
            if (_pnlToolbar != null)
                _pnlToolbar.BackColor = _colorToolbarBg;

            // 更新底部面板
            if (_pnlBottom != null)
                _pnlBottom.BackColor = _colorToolbarBg;

            // 更新路径下拉框
            if (_cboPath != null)
            {
                _cboPath.BackColor = isDark ? _colorDarkBackColor : Color.White;
                _cboPath.ForeColor = _colorTextPrimary;
            }

            // 更新过滤文本框
            if (_txtFilter != null)
            {
                _txtFilter.BackColor = isDark ? _colorDarkBackColor : Color.White;
                _txtFilter.ForeColor = _filterText == FilterPlaceholder ? _colorTextSecondary : _colorTextPrimary;
                _txtFilter.BorderStyle = BorderStyle.FixedSingle;
            }

            // 更新下拉框（视图模式）
            if (_cboViewMode != null)
            {
                _cboViewMode.BackColor = isDark ? _colorDarkToolbarBg : Color.White;
                _cboViewMode.ForeColor = _colorTextPrimary;
            }

            // 更新数量标签
            if (_lblFileCount != null)
            {
                _lblFileCount.ForeColor = _colorTextSecondary;
                _lblFileCount.BackColor = _colorToolbarBg;
            }

            // 更新列表视图
            _lstFiles.BackColor = isDark ? _colorDarkBackColor : Color.White;
            _lstFiles.ForeColor = _colorTextPrimary;

            // 设置 ListView 主题（深色模式使用 Explorer 主题）
            if (isDark)
            {
                SetWindowTheme(_lstFiles.Handle, "Explorer", null);
            }
            else
            {
                SetWindowTheme(_lstFiles.Handle, null, null);
            }

            // 更新按钮（上级按钮和打开目录按钮统一使用深色模式的蓝色）
            var btnBackColor = Color.FromArgb(66, 165, 245);  // 天蓝色按钮
            var btnForeColor = Color.White;                   // 始终使用白色文字

            if (_btnUp != null)
            {
                _btnUp.BackColor = btnBackColor;
                _btnUp.ForeColor = btnForeColor;
            }

            if (_btnOpenExplorer != null)
            {
                _btnOpenExplorer.BackColor = btnBackColor;
                _btnOpenExplorer.ForeColor = btnForeColor;
            }

            if (_btnOpenCsvReader != null)
            {
                _btnOpenCsvReader.BackColor = btnBackColor;
                _btnOpenCsvReader.ForeColor = btnForeColor;
            }

            if (_cboViewMode != null)
            {
                _cboViewMode.BackColor = isDark ? _colorDarkToolbarBg : Color.White;
                _cboViewMode.ForeColor = _colorTextPrimary;
            }

            Invalidate();
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

            BackColor = _colorLightBackColor; // 默认浅色背景
            BorderStyle = BorderStyle.FixedSingle;
            Padding = new Padding(1);

            // ===== 工具栏面板（顶部）=====
            _pnlToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 36,
                BackColor = _colorToolbarBg,
                Padding = new Padding(6, 4, 6, 4)
            };

            _cboPath = new ComboBox
            {
                Location = new Point(6, 6),
                Width = 200,
                Height = 23,
                Font = new Font("Microsoft YaHei UI", 9f),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.FileSystem
            };
            _cboPath.KeyDown += OnPathKeyDown;
            _cboPath.SelectedIndexChanged += OnPathComboSelectedIndexChanged;

            _btnUp = new Button
            {
                Text = "上级",
                Location = new Point(332, 5),
                Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            _btnUp.Click += (s, e) => GoUp();

            // 视图模式选择 ComboBox（替换刷新按钮位置）
            _cboViewMode = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(386, 5),
                Size = new Size(90, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            _cboViewMode.Items.AddRange(new object[] { "详细信息", "列表", "小图标", "大图标", "平铺" });
            _cboViewMode.SelectedIndex = (int)_viewMode;
            _cboViewMode.SelectedIndexChanged += OnViewModeChanged;

            _pnlToolbar.Controls.AddRange(new Control[] { _cboPath, _btnUp, _cboViewMode });

            // 工具栏大小改变时调整路径下拉框宽度
            _pnlToolbar.Resize += (s, e) =>
            {
                if (_cboPath != null && _btnUp != null && _cboViewMode != null)
                {
                    int buttonsWidth = _btnUp.Width + _cboViewMode.Width + 12;
                    int newWidth = _pnlToolbar.Width - _cboPath.Left - buttonsWidth;
                    if (newWidth > 100)
                        _cboPath.Width = newWidth;
                    _btnUp.Left = _pnlToolbar.Width - _btnUp.Width - _cboViewMode.Width - 10;
                    _cboViewMode.Left = _pnlToolbar.Width - _cboViewMode.Width - 6;
                }
            };

            // ===== 文件列表（中间区域）=====
            _lstFiles.Dock = DockStyle.Fill;
            _lstFiles.View = View.Details;
            _lstFiles.FullRowSelect = true;
            _lstFiles.GridLines = false; // 去掉网格线
            _lstFiles.HeaderStyle = ColumnHeaderStyle.Clickable;
            _lstFiles.SmallImageList = _imgList;
            _lstFiles.Font = new Font("Microsoft YaHei UI", 9.5f);
            _lstFiles.MultiSelect = true;
            _lstFiles.ShowItemToolTips = true;

            // 添加列：名称列自适应宽度，其他列固定宽度
            _lstFiles.Columns.Add("名称", 220, HorizontalAlignment.Left);
            _lstFiles.Columns.Add("修改日期", 130, HorizontalAlignment.Left);
            _lstFiles.Columns.Add("类型", 100, HorizontalAlignment.Left);
            _lstFiles.Columns.Add("大小", 70, HorizontalAlignment.Right);

            // 监听列宽调整事件，保持名称列自适应
            _lstFiles.ColumnWidthChanged += OnColumnWidthChanged;
            _lstFiles.Resize += OnListViewResize;

            _lstFiles.DoubleClick += OnFileDoubleClick;
            _lstFiles.MouseUp += OnListViewMouseUp;
            _lstFiles.SelectedIndexChanged += (s, e) =>
            {
                if (_lstFiles.SelectedItems.Count > 0) _lstFiles.Focus();
            };

            // ===== 底部面板（包含过滤框和打开资源管理器按钮）=====
            _pnlBottom = new Panel
            {
                Height = 36,
                Dock = DockStyle.Bottom,
                BackColor = _colorToolbarBg
            };

            // 打开CSV阅读器按钮 - 使用 Anchor 定位在右侧
            _btnOpenCsvReader = new Button
            {
                Width = 90,
                Height = 26,
                Location = new Point(_pnlBottom.Width - 180, 5),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Text = "阅读器",
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
            };
            _btnOpenCsvReader.Click += OnOpenCsvReaderClick;

            // 打开资源管理器按钮 - 使用 Anchor 定位在右侧
            _btnOpenExplorer = new Button
            {
                Width = 80,
                Height = 26,
                Location = new Point(_pnlBottom.Width - 84, 5),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Text = "打开目录",
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
            };
            _btnOpenExplorer.Click += OnOpenExplorerClick;

            // 文件/文件夹数量标签
            _lblFileCount = new Label
            {
                Width = 130,
                Height = 26,
                Location = new Point(4, 5),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                Text = "文件夹: 0  文件: 0",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorTextSecondary,
                BackColor = _colorToolbarBg
            };

            // 过滤框 - 使用 Anchor 填充剩余空间
            _txtFilter = new TextBox
            {
                Height = 29,
                Location = new Point(136, 4),
                Size = new Size(_pnlBottom.Width - 328, 29),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle
            };

            SetPlaceholder();

            // 创建过滤防抖定时器（300ms 延迟）
            _filterDebounceTimer = new System.Windows.Forms.Timer { Interval = 300 };
            _filterDebounceTimer.Tick += (s, e) =>
            {
                _filterDebounceTimer!.Stop();
                ApplyFilter();
            };

            // 创建保存防抖定时器（3秒延迟）
            _saveDebounceTimer = new System.Windows.Forms.Timer { Interval = 3000 };
            _saveDebounceTimer.Tick += (s, e) =>
            {
                _saveDebounceTimer!.Stop();
                if (_hasUnsavedChanges)
                {
                    SavePanelState();
                    _hasUnsavedChanges = false;
                }
            };

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

            _pnlBottom.Controls.Add(_lblFileCount);
            _pnlBottom.Controls.Add(_txtFilter);
            _pnlBottom.Controls.Add(_btnOpenCsvReader);
            _pnlBottom.Controls.Add(_btnOpenExplorer);

            // 底部面板大小改变时调整过滤框宽度
            _pnlBottom.Resize += (s, e) =>
            {
                if (_txtFilter != null && _btnOpenCsvReader != null)
                {
                    int newWidth = _pnlBottom.Width - _lblFileCount!.Width - _btnOpenCsvReader.Width - _btnOpenExplorer!.Width - 21;
                    if (newWidth > 50)
                        _txtFilter.Width = newWidth;
                    _txtFilter.Top = (_pnlBottom.Height - _txtFilter.Height) / 2;
                }
            };

            Controls.Add(_lstFiles);
            Controls.Add(_pnlBottom);
            Controls.Add(_pnlToolbar);

            MinimumSize = new Size(280, 280);
            ResumeLayout(false);
        }

        private const string FilterPlaceholder = "输入关键词过滤 (支持 || 或 / && 与)";

        private void SetPlaceholder()
        {
            if (_txtFilter == null) return;
            _txtFilter.Text = FilterPlaceholder;
            _txtFilter.ForeColor = _colorTextSecondary;

            _txtFilter.GotFocus -= Filter_GotFocus;
            _txtFilter.GotFocus += Filter_GotFocus;
            _txtFilter.LostFocus -= Filter_LostFocus;
            _txtFilter.LostFocus += Filter_LostFocus;
        }

        private void Filter_GotFocus(object? sender, EventArgs e)
        {
            if (_txtFilter != null && _txtFilter.Text == FilterPlaceholder)
            {
                _txtFilter.Text = "";
                _txtFilter.ForeColor = _colorTextPrimary;
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

            // 大图标列表（48×48）
            _imgListLarge.ImageSize = new Size(48, 48);
            _imgListLarge.ColorDepth = ColorDepth.Depth32Bit;

            var largeFolderShfi = default(SHFILEINFO);
            SHGetFileInfo(@"C:\Windows", FILE_ATTRIBUTE_DIRECTORY, out largeFolderShfi, (uint)Marshal.SizeOf(largeFolderShfi),
                          SHGFI_ICON | SHGFI_USEFILEATTRIBUTES);
            _imgListLarge.Images.Add("folder", Icon.FromHandle(largeFolderShfi.hIcon));

            var largeFileShfi = default(SHFILEINFO);
            SHGetFileInfo(".txt", 0, out largeFileShfi, (uint)Marshal.SizeOf(largeFileShfi),
                          SHGFI_ICON | SHGFI_USEFILEATTRIBUTES);
            _imgListLarge.Images.Add("file_default", Icon.FromHandle(largeFileShfi.hIcon));
        }

        // ==================== 路径下拉框 ====================
        /// <summary>重建路径下拉框选项：磁盘盘符 + 最近访问路径</summary>
        private void RebuildPathComboItems()
        {
            if (_cboPath == null) return;
            _suppressEvents = true;
            _cboPath.Items.Clear();

            // 添加磁盘盘符
            foreach (var drive in DriveInfo.GetDrives().Where(d => d.IsReady))
                _cboPath.Items.Add(drive.RootDirectory.FullName);

            // 添加分隔线
            _cboPath.Items.Add(new SeparatorItem());

            // 添加最近访问路径
            var recents = CsvSettings.GetRecentPaths();
            foreach (var p in recents)
            {
                if (Directory.Exists(p) && !_cboPath.Items.Contains(p))
                    _cboPath.Items.Add(p);
            }

            _suppressEvents = false;
        }

        private string GetFirstDrive()
        {
            // 从路径下拉框中找第一个有效的盘符选项
            if (_cboPath != null)
            {
                foreach (var item in _cboPath.Items)
                {
                    if (item is string s && Directory.Exists(s) && s.Length >= 2 && s[1] == ':')
                        return s;
                }
            }
            return "C:\\";
        }

        /// <summary>下拉框选中项变化——当用户从下拉列表中选择时触发</summary>
        private void OnPathComboSelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_suppressEvents || _cboPath?.SelectedItem == null) return;

            // 忽略分隔线选中
            if (_cboPath.SelectedItem is SeparatorItem) return;

            var selected = _cboPath.SelectedItem.ToString();
            if (!string.IsNullOrEmpty(selected) && Directory.Exists(selected))
                NavigateTo(selected);
        }

        // ==================== 导航操作 ====================
        public void NavigateTo(string path, bool restoreFilter = true, bool saveState = true)
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
                // 记录到最近访问历史
                CsvSettings.AddRecentPath(_currentDir);
                // 重建下拉框选项（刷新驱动器和最近路径）
                RebuildPathComboItems();
                RefreshUI();
                if (saveState)
                {
                    // 标记需要保存，重启保存定时器
                    _hasUnsavedChanges = true;
                    _saveDebounceTimer?.Stop();
                    _saveDebounceTimer?.Start();
                }
                // 异步加载文件列表
                _ = LoadFileListAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"无法访问路径:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                if (_cboPath == null) return;
                var input = _cboPath.Text.Trim();
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

            if (_cboPath != null)
            {
                _cboPath.Text = _currentDir;
                _pathToolTip.SetToolTip(_cboPath, _currentDir);
            }

            // 同步选中项：尝试在下拉框中找到匹配项
            if (_cboPath != null)
            {
                int targetIdx = -1;
                for (int i = 0; i < _cboPath.Items.Count; i++)
                {
                    if (_cboPath.Items[i] is SeparatorItem) continue;
                    var itemText = _cboPath.Items[i].ToString() ?? "";
                    if (_currentDir.Equals(itemText, StringComparison.OrdinalIgnoreCase))
                    {
                        targetIdx = i;
                        break;
                    }
                }
                if (targetIdx >= 0 && _cboPath.SelectedIndex != targetIdx)
                    _cboPath.SelectedIndex = targetIdx;
            }

            _suppressEvents = false;
        }

        private CancellationTokenSource? _loadCts;

        /// <summary>异步加载文件列表，避免界面卡顿</summary>
        private async Task LoadFileListAsync()
        {
            // 取消之前的加载任务
            _loadCts?.Cancel();
            _loadCts = new CancellationTokenSource();
            var token = _loadCts.Token;

            ShowLoading();

            try
            {
                var dirs = new List<DirInfo>();
                var files = new List<FileInfo>();

                // 在后台线程加载目录
                await Task.Run(() =>
                {
                    if (token.IsCancellationRequested) return;

                    try
                    {
                        var showHidden = CsvSettings.ShowHiddenFiles;
                        foreach (var dir in Directory.GetDirectories(_currentDir))
                        {
                            if (token.IsCancellationRequested) return;
                            // 不显示隐藏文件夹时，跳过隐藏属性的目录
                            if (!showHidden)
                            {
                                try { if ((File.GetAttributes(dir) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                            }
                            try { dirs.Add(new DirInfo { Name = Path.GetFileName(dir), FullName = dir }); } catch { }
                        }
                        dirs.Sort((a, b) => new NaturalSortComparer().Compare(a.Name, b.Name));

                        if (token.IsCancellationRequested) return;

                        foreach (var file in Directory.GetFiles(_currentDir))
                        {
                            if (token.IsCancellationRequested) return;
                            // 不显示隐藏文件时，跳过隐藏属性的文件
                            if (!showHidden)
                            {
                                try { if ((File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                            }
                            try { files.Add(new FileInfo(file)); } catch { }
                        }
                        files.Sort((a, b) => new NaturalSortComparer().Compare(a.Name, b.Name));
                    }
                    catch (UnauthorizedAccessException) { }
                }, token);

                if (token.IsCancellationRequested) return;

                // 更新数据
                _allDirs = dirs;
                _allFiles = files;

                // 在 UI 线程更新列表
                if (!token.IsCancellationRequested)
                {
                    if (IsHandleCreated && !IsDisposed)
                    {
                        Invoke(() => ApplyFilter());
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // 任务被取消，忽略
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加载文件列表失败: {ex.Message}");
            }
            finally
            {
                HideLoading();
            }
        }

        /// <summary>显示加载提示</summary>
        private void ShowLoading()
        {
            if (_lblLoading == null)
            {
                _lblLoading = new Label
                {
                    Text = "正在加载...",
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Microsoft YaHei UI", 10f),
                    ForeColor = _colorTextSecondary,
                    BackColor = Color.Transparent,
                    Visible = false
                };
                // 添加到列表视图上层，居中显示
                _lstFiles.Controls.Add(_lblLoading);
                _lblLoading.BringToFront();
            }

            _lblLoading.Visible = true;
            _lblLoading.Size = _lstFiles.ClientSize;
            _lblLoading.Location = new Point(0, 0);
            _isLoading = true;
        }

        /// <summary>隐藏加载提示</summary>
        private void HideLoading()
        {
            if (_lblLoading != null)
            {
                _lblLoading.Visible = false;
            }
            _isLoading = false;
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
                    MessageBox.Show(this.FindForm(), $"无法打开资源管理器: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // ==================== 打开CSV阅读器 ====================
        private void OnOpenCsvReaderClick(object? sender, EventArgs e)
        {
            try
            {
                var csvForm = new CsvReaderForm(_currentDir, PanelIndex);
                csvForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this.FindForm(), $"无法打开CSV阅读器: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ==================== 过滤逻辑 ====================
        private void OnFilterTextChanged(object? sender, EventArgs e)
        {
            if (_suppressEvents || _txtFilter == null) return;
            var txt = _txtFilter.Text;
            if (txt == FilterPlaceholder)
            {
                _filterText = "";
            }
            else
            {
                _filterText = txt ?? "";
            }

            // 标记需要保存，重启保存定时器
            _hasUnsavedChanges = true;
            _saveDebounceTimer?.Stop();
            _saveDebounceTimer?.Start();

            // 防抖：重启定时器，300ms 内无新输入才真正过滤
            _filterDebounceTimer?.Stop();
            _filterDebounceTimer?.Start();
        }

        private void PopulateListView()
        {
            if (_lstFiles == null || _lstFiles.IsDisposed) return;

            _lstFiles.BeginUpdate();

            // ===== 先清空并重置到 Details，清除原生控件之前残留的间距/布局状态 =====
            _lstFiles.Items.Clear();
            _lstFiles.View = View.Details;
            _lstFiles.SmallImageList = null;
            _lstFiles.LargeImageList = null;
            _lstFiles.HeaderStyle = ColumnHeaderStyle.Clickable;

            // ===== 根据当前视图模式配置 ListView =====
            switch (_viewMode)
            {
                case FileListViewMode.Details:
                    _lstFiles.View = View.Details;
                    _lstFiles.HeaderStyle = ColumnHeaderStyle.Clickable;
                    _lstFiles.SmallImageList = _imgList;
                    break;
                case FileListViewMode.List:
                    _lstFiles.View = View.List;
                    _lstFiles.HeaderStyle = ColumnHeaderStyle.None;
                    _lstFiles.SmallImageList = _imgList;
                    break;
                case FileListViewMode.LargeIcon:
                    _lstFiles.View = View.LargeIcon;
                    _lstFiles.HeaderStyle = ColumnHeaderStyle.None;
                    _lstFiles.LargeImageList = _imgListLarge;
                    break;
                case FileListViewMode.Tile:
                    _lstFiles.View = View.Tile;
                    _lstFiles.HeaderStyle = ColumnHeaderStyle.None;
                    _lstFiles.LargeImageList = _imgListLarge;
                    _lstFiles.TileSize = new Size(180, 56);
                    break;
                default: // SmallIcon — 先用 List 绕过布局 bug，填充完再切
                    _lstFiles.View = View.List;
                    _lstFiles.HeaderStyle = ColumnHeaderStyle.None;
                    _lstFiles.SmallImageList = _imgList;
                    break;
            }

            // ===== 填充数据 =====
            int dirCount = 0;
            int fileCount = 0;

            foreach (var dir in FilterItems(_allDirs.Select(d => d.Name)))
            {
                var info = _allDirs.First(d => d.Name.Equals(dir, StringComparison.OrdinalIgnoreCase));
                AddListViewItem(info.Name, info.FullName, isDir: true);
                dirCount++;
            }

            foreach (var fn in FilterItems(_allFiles.Select(f => f.Name)))
            {
                var fi = _allFiles.First(f => f.Name.Equals(fn, StringComparison.OrdinalIgnoreCase));
                AddListViewItem(fi.Name, fi.FullName, isDir: false, fi.LastWriteTime, fi.Length);
                fileCount++;
            }

            // ===== 更新统计 =====
            if (_lblFileCount != null)
                _lblFileCount.Text = $"文件夹: {dirCount}  文件: {fileCount}";

            // ===== 列宽调整（仅 Details 模式）=====
            if (_viewMode == FileListViewMode.Details)
                AutoResizeColumns();

            _lstFiles.EndUpdate();

            // ===== SmallIcon：填充完后切回 SmallIcon，触发原生控件正确布局（用 BeginUpdate 避免闪烁）=====
            if (_viewMode == FileListViewMode.SmallIcon)
            {
                _lstFiles.BeginUpdate();
                _lstFiles.View = View.SmallIcon;
                _lstFiles.EndUpdate();
            }
        }

        private void ApplyFilter()
        {
            PopulateListView();
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
        private void AddListViewItem(string name, string fullName, bool isDir,
                                DateTime lastWrite = default, long size = 0)
        {
            var item = new ListViewItem(name);
            item.Tag = new FileItemTag { FullName = fullName, IsDir = isDir };
            item.ToolTipText = fullName;

            if (isDir)
            {
                item.ImageKey = "folder";
                // 确保大图标列表也有 folder
                if (!_imgListLarge.Images.ContainsKey("folder"))
                {
                    var shfi = default(SHFILEINFO);
                    SHGetFileInfo(@"C:\Windows", FILE_ATTRIBUTE_DIRECTORY, out shfi, (uint)Marshal.SizeOf(shfi),
                                  SHGFI_ICON | SHGFI_USEFILEATTRIBUTES);
                    if (shfi.hIcon != IntPtr.Zero)
                        _imgListLarge.Images.Add("folder", Icon.FromHandle(shfi.hIcon));
                }
                item.SubItems.Add("");
                item.SubItems.Add("文件夹");
                item.SubItems.Add("");
            }
            else
            {
                var iconKey = GetFileIconKey(fullName);
                item.ImageKey = iconKey;
                // 同步加载大图标
                if (!_imgListLarge.Images.ContainsKey(iconKey))
                {
                    var largeKey = Path.GetExtension(fullName).ToLowerInvariant();
                    if (largeKey != iconKey) largeKey = iconKey;
                    if (!_imgListLarge.Images.ContainsKey(largeKey))
                    {
                        var shfi = default(SHFILEINFO);
                        SHGetFileInfo(fullName, 0, out shfi, (uint)Marshal.SizeOf(shfi), SHGFI_ICON);
                        if (shfi.hIcon != IntPtr.Zero)
                            _imgListLarge.Images.Add(largeKey, Icon.FromHandle(shfi.hIcon));
                    }
                }
                item.SubItems.Add(lastWrite == default ? "" : lastWrite.ToString("yyyy-MM-dd HH:mm"));
                item.SubItems.Add(GetFileType(fullName));
                item.SubItems.Add(size > 0 ? FormatFileSize(size) : "");
            }

            _lstFiles.Items.Add(item);
        }

        private string GetFileIconKey(string filePath)
        {
            return GetFileIconKeyInternal(filePath, small: true);
        }

        private string GetFileIconKeyInternal(string filePath, bool small)
        {
            try
            {
                var key = Path.GetExtension(filePath).ToLowerInvariant();
                var list = small ? _imgList : _imgListLarge;
                if (!list.Images.ContainsKey(key))
                {
                    uint flags = SHGFI_ICON;
                    if (small) flags |= SHGFI_SMALLICON;
                    var shfi = default(SHFILEINFO);
                    SHGetFileInfo(filePath, 0, out shfi, (uint)Marshal.SizeOf(shfi), flags);
                    if (shfi.hIcon != IntPtr.Zero)
                    {
                        list.Images.Add(key, Icon.FromHandle(shfi.hIcon));
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
            // 名称列铺满剩余宽度，其他列保持固定宽度
            OnListViewResize(null, EventArgs.Empty);
        }

        // ==================== 视图模式切换 ====================
        private void OnViewModeChanged(object? sender, EventArgs e)
        {
            if (_suppressEvents || _cboViewMode == null) return;
            _viewMode = (FileListViewMode)_cboViewMode.SelectedIndex;
            // 用 PopulateListView 重新渲染列表（switch 内部根据 _viewMode 切换布局）
            PopulateListView();

            // 立即持久化视图模式
            try { CsvSettings.SetPaneViewMode(PanelIndex, (int)_viewMode); }
            catch { }
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

        /// <summary>处理 ListView 大小变化，使名称列自适应剩余宽度</summary>
        private void OnListViewResize(object? sender, EventArgs e)
        {
            if (_lstFiles.Columns.Count < 4) return;

            // 计算其他列的总宽度
            int otherColumnsWidth = 0;
            for (int i = 1; i < _lstFiles.Columns.Count; i++)
            {
                otherColumnsWidth += _lstFiles.Columns[i].Width;
            }

            // 名称列宽度 = 列表宽度 - 其他列宽度 - 滚动条宽度(如果有)
            int newWidth = _lstFiles.ClientSize.Width - otherColumnsWidth - 4;
            if (newWidth < 100) newWidth = 100; // 最小宽度

            _lstFiles.Columns[0].Width = newWidth;
        }

        /// <summary>处理列宽手动调整，保持名称列自适应</summary>
        private void OnColumnWidthChanged(object? sender, ColumnWidthChangedEventArgs e)
        {
            // 如果调整的不是名称列（第0列），则重新计算名称列宽度
            if (e.ColumnIndex != 0)
            {
                OnListViewResize(sender, EventArgs.Empty);
            }
        }

        /// <summary>处理 ListView 鼠标释放事件（用于右键菜单）</summary>
        private void OnListViewMouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;

            Debug.WriteLine($"MouseUp事件触发: X={e.X}, Y={e.Y}");

            // 直接使用 MouseUp 事件处理右键，避免 MouseClick 在空白处不触发的问题
            var hit = _lstFiles.HitTest(e.Location);

            // 检测列头区域
            if (IsPointOnColumnHeader(e.Location))
            {
                Debug.WriteLine("MouseUp: 点击列头区域");
                ShowSystemContextMenuForPaths(new[] { _currentDir });
                return;
            }

            // 如果点击的是列表项
            if (hit.Item != null)
            {
                int columnIndex = 0;
                if (hit.SubItem != null)
                    columnIndex = hit.Item.SubItems.IndexOf(hit.SubItem);

                // 如果点击的项未选中，则只选中它（取消其他选择）
                if (!hit.Item.Selected)
                {
                    _lstFiles.SelectedItems.Clear();
                    hit.Item.Selected = true;
                }
                // 如果点击的项已选中，则保持当前多选状态不变

                ShowContextMenuForSelectedItems();
            }
            else
            {
                // 点击空白区域：取消所有选择，显示目录背景菜单
                _lstFiles.SelectedItems.Clear();
                Debug.WriteLine("MouseUp: 点击空白区域");
                if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
                    ShowSystemContextMenuForPaths(new[] { _currentDir });
            }
        }

        /// <summary>获取当前选中项的完整路径列表</summary>
        private List<string> GetSelectedFullPaths()
        {
            var paths = new List<string>();
            foreach (ListViewItem item in _lstFiles.SelectedItems)
            {
                if (item.Tag is FileItemTag t)
                    paths.Add(t.FullName);
            }
            return paths;
        }

        /// <summary>为当前选中项显示右键菜单</summary>
        private void ShowContextMenuForSelectedItems()
        {
            var paths = GetSelectedFullPaths();
            if (paths.Count == 0)
            {
                if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
                    ShowSystemContextMenuForPaths(new[] { _currentDir });
                return;
            }

            if (paths.Count == 1)
            {
                // 单个文件/文件夹：使用系统原生右键菜单
                ShowSystemContextMenu(paths[0]);
            }
            else
            {
                // 多个文件/文件夹：使用内置右键菜单（拷贝、删除、复制路径等）
                ShowMultiSelectContextMenu(paths);
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

        /// <summary>显示系统右键上下文菜单（单个路径）</summary>
        private void ShowSystemContextMenu(string itemPath)
        {
            // 检查控件状态，防止在已释放状态下处理事件
            if (this.IsDisposed || !this.IsHandleCreated) return;

            Debug.WriteLine($"ShowSystemContextMenu: path='{itemPath}', exists={Directory.Exists(itemPath) || File.Exists(itemPath)}");

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

        /// <summary>显示系统右键上下文菜单（支持多路径——仅当仅一个路径时使用 Shell API）</summary>
        private void ShowSystemContextMenuForPaths(string[] paths)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            if (paths.Length == 0) return;

            if (paths.Length == 1)
            {
                ShowSystemContextMenu(paths[0]);
            }
            else
            {
                ShowMultiSelectContextMenu(new List<string>(paths));
            }
        }

        /// <summary>多选右键菜单：复制、剪切、删除、复制路径</summary>
        private void ShowMultiSelectContextMenu(List<string> paths)
        {
            try
            {
                var menu = new ContextMenuStrip();
                menu.Font = new Font("Microsoft YaHei UI", 9f);

                ToolStripMenuItem itemCopy = new ToolStripMenuItem("复制", null, (s, e) => CopyFilesToClipboard(paths, false));
                ToolStripMenuItem itemCut = new ToolStripMenuItem("剪切", null, (s, e) => CopyFilesToClipboard(paths, true));
                ToolStripMenuItem itemDelete = new ToolStripMenuItem("删除", null, (s, e) => DeleteFiles(paths));
                ToolStripMenuItem itemCopyPath = new ToolStripMenuItem("复制路径", null, (s, e) => CopyPathsToClipboard(paths));

                menu.Items.Add(itemCopy);
                menu.Items.Add(itemCut);
                menu.Items.Add("-");
                menu.Items.Add(itemDelete);
                menu.Items.Add("-");
                menu.Items.Add(itemCopyPath);

                menu.Show(Cursor.Position);
            }
            catch
            {
                // 忽略
            }
        }

        /// <summary>复制或剪切文件到剪贴板</summary>
        private void CopyFilesToClipboard(List<string> paths, bool cut)
        {
            try
            {
                var collection = new System.Collections.Specialized.StringCollection();
                collection.AddRange(paths.ToArray());
                Clipboard.SetFileDropList(collection);
                if (cut)
                {
                    // 标记剪切操作（通过设置包含 Preferred DropEffect 的 MemoryStream）
                    byte[] dropEffect = new byte[] { (byte)DragDropEffects.Move, 0, 0, 0 };
                    using (var ms = new System.IO.MemoryStream())
                    {
                        ms.Write(dropEffect, 0, dropEffect.Length);
                        var data = new DataObject();
                        data.SetFileDropList(collection);
                        data.SetData("Preferred DropEffect", ms);
                        Clipboard.SetDataObject(data);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this.FindForm(), $"复制到剪贴板失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>复制路径列表到剪贴板</summary>
        private void CopyPathsToClipboard(List<string> paths)
        {
            try
            {
                Clipboard.SetText(string.Join(Environment.NewLine, paths));
            }
            catch { }
        }

        /// <summary>删除选中文件/文件夹（使用系统 Shell API，显示原生进度对话框）</summary>
        private void DeleteFiles(List<string> paths)
        {
            if (paths.Count == 0) return;

            string msg = paths.Count == 1
                ? $"确定要删除 \"{Path.GetFileName(paths[0])}\" 吗？"
                : $"确定要删除选中的 {paths.Count} 个项目吗？";

            var result = MessageBox.Show(this.FindForm(), msg, "确认删除",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result != DialogResult.Yes) return;

            try
            {
                // 用 \0 分隔文件路径，最后以 \0\0 结尾
                var fromBuf = string.Join("\0", paths) + "\0\0";

                var fileOp = new SHFILEOPSTRUCT
                {
                    hwnd = this.Handle,
                    wFunc = FO_DELETE,
                    pFrom = fromBuf,
                    fFlags = FOF_ALLOWUNDO | FOF_SIMPLEPROGRESS | FOF_WANTNUKEWARNING
                };

                int ret = SHFileOperation(ref fileOp);
                if (ret != 0)
                {
                    // 用户取消不算错误
                    if (fileOp.fAnyOperationsAborted == 0 && ret != 0x78) // 0x78 = DE_SAMEFILE
                    {
                        // 只提示可能部分失败，不弹错误框干扰用户
                        Debug.WriteLine($"SHFileOperation 返回: {ret}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"系统删除 API 失败，使用逐文件删除: {ex.Message}");
                // 降级：同步逐文件删除（后台线程避免卡UI）
                Task.Run(() =>
                {
                    foreach (var p in paths)
                    {
                        try
                        {
                            if (Directory.Exists(p))
                                Directory.Delete(p, true);
                            else if (File.Exists(p))
                                File.Delete(p);
                        }
                        catch { }
                    }
                }).ContinueWith(_ =>
                {
                    if (!this.IsDisposed)
                    {
                        try { BeginInvoke(() => RefreshCurrent()); }
                        catch { }
                    }
                });
                return;
            }

            // 完成后刷新
            RefreshCurrent();
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

                // 如果是文件，使用 cmd /c start 打开，模拟资源管理器双击效果
                if (File.Exists(filePath))
                {
                    var workingDir = Path.GetDirectoryName(filePath) ?? "";
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = $"/c start \"\" \"{filePath}\"",
                        WorkingDirectory = workingDir,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    Process.Start(startInfo);
                    return;
                }

                // 如果是文件夹，使用 ShellExecuteEx
                if (Directory.Exists(filePath))
                {
                    var info = new SHELLEXECUTEINFO
                    {
                        cbSize = Marshal.SizeOf<SHELLEXECUTEINFO>(),
                        fMask = SEE_MASK_FLAG_NO_UI,
                        hwnd = this.Handle,
                        lpVerb = verb,
                        lpFile = filePath,
                        nShow = 1
                    };
                    ShellExecuteEx(ref info);
                }
            }
            catch (ObjectDisposedException)
            {
                // 忽略已释放的对象异常
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this.FindForm(), $"无法打开:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ==================== 焦点高亮 ====================
        private static FileBrowserPanel? _focusedPanel = null;

        private bool HasFocus => _focusedPanel == this;

        protected override void OnEnter(EventArgs e)
        {
            base.OnEnter(e);
            var prev = _focusedPanel;
            _focusedPanel = this;
            if (prev != null && prev != this) prev.Invalidate();
            Invalidate();
        }

        protected override void OnLeave(EventArgs e)
        {
            base.OnLeave(e);
            if (_focusedPanel == this) _focusedPanel = null;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            // 在内置边框内侧绘制焦点高亮
            if (HasFocus)
            {
                ControlPaint.DrawBorder(e.Graphics, ClientRectangle,
                    SystemColors.Highlight, 2, ButtonBorderStyle.Inset,
                    SystemColors.Highlight, 2, ButtonBorderStyle.Inset,
                    SystemColors.Highlight, 2, ButtonBorderStyle.Inset,
                    SystemColors.Highlight, 2, ButtonBorderStyle.Inset);
            }
        }

        // ==================== 键盘快捷键 ====================
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.F5)
            {
                RefreshCurrent();
                return true;
            }
            if (keyData == Keys.F2)
            {
                StartRename();
                return true;
            }
            if (keyData == Keys.Delete)
            {
                var paths = GetSelectedFullPaths();
                if (paths.Count > 0)
                {
                    DeleteFiles(paths);
                    return true;
                }
            }
            if (keyData == (Keys.Control | Keys.A))
            {
                // 全选
                if (_lstFiles.Items.Count > 0)
                {
                    foreach (ListViewItem item in _lstFiles.Items)
                        item.Selected = true;
                }
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ==================== F2 重命名 ====================
        private void StartRename()
        {
            if (_lstFiles.SelectedItems.Count == 0) return;
            var item = _lstFiles.SelectedItems[0];
            var tag = item.Tag as FileItemTag;
            if (tag == null) return;

            var tb = new TextBox
            {
                Text = item.Text,
                Bounds = item.Bounds,
                Left = item.Bounds.Left + 18,
                Width = item.Bounds.Width - 20,
                Font = _lstFiles.Font,
                BackColor = _isDark ? _colorDarkBackColor : Color.White,
                ForeColor = _isDark ? _colorDarkTextPrimary : _colorTextPrimary,
                BorderStyle = BorderStyle.FixedSingle,
            };

            void CommitRename(object? s, EventArgs e)
            {
                var newName = tb.Text.Trim();
                if (!string.IsNullOrEmpty(newName) && newName != item.Text)
                {
                    var dir = Path.GetDirectoryName(tag.FullName)!;
                    var newPath = Path.Combine(dir, newName);
                    if (!File.Exists(newPath) && !Directory.Exists(newPath))
                    {
                        try
                        {
                            if (tag.IsDir)
                                Directory.Move(tag.FullName, newPath);
                            else
                                File.Move(tag.FullName, newPath);
                            _ = LoadFileListAsync();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(this.FindForm(), $"重命名失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show(this.FindForm(), "该名称已存在。", "重命名", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                tb.LostFocus -= CommitRename;
                tb.KeyDown -= OnRenameKeyDown;
                tb.Dispose();
            }

            void OnRenameKeyDown(object? s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    CommitRename(tb, EventArgs.Empty);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    tb.LostFocus -= CommitRename;
                    tb.KeyDown -= OnRenameKeyDown;
                    tb.Dispose();
                }
            }

            tb.LostFocus += CommitRename;
            tb.KeyDown += OnRenameKeyDown;
            _lstFiles.Controls.Add(tb);
            tb.Focus();
            tb.SelectAll();
        }

        // ==================== 公共接口 & 状态持久化 ====================
        /// <summary>刷新当前目录（通常用于显示隐藏设置变更后重新加载）</summary>
        public void RefreshCurrentDirectory()
        {
            if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
                NavigateTo(_currentDir, restoreFilter: false);
        }

        public void RefreshCurrent()
        {
            RefreshCurrentDirectory();
        }

        private void SavePanelState()
        {
            try
            {
                // 使用ConfigurationManager保存各窗格的目录和过滤条件
                var filter = _filterText == FilterPlaceholder ? "" : _filterText;
                System.Diagnostics.Debug.WriteLine($"SavePanelState: PanelIndex={PanelIndex}, Filter={filter}, _filterText={_filterText}");
                CsvSettings.SetPaneDirectory(PanelIndex, _currentDir);
                CsvSettings.SetPaneFilter(PanelIndex, filter);
                CsvSettings.SetPaneViewMode(PanelIndex, (int)_viewMode);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SavePanelState Error: {ex.Message}");
            }
        }

        /// <summary>强制立即保存当前面板状态（用于手动保存按钮）</summary>
        public void ForceSaveState()
        {
            // 停止定时器
            _saveDebounceTimer?.Stop();
            _hasUnsavedChanges = false;
            // 立即保存
            SavePanelState();
        }

        protected override void Dispose(bool disposing)
        {
            // 停止定时器并立即保存
            _saveDebounceTimer?.Stop();
            _filterDebounceTimer?.Stop();
            if (_hasUnsavedChanges)
            {
                SavePanelState();
                _hasUnsavedChanges = false;
            }
            if (disposing)
            {
                _filterDebounceTimer?.Dispose();
                _saveDebounceTimer?.Dispose();
                _pathToolTip?.Dispose();
                _imgList?.Dispose();
                _imgListLarge?.Dispose();
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
