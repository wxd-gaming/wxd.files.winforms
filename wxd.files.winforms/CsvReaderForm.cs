using System.ComponentModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using System.Diagnostics;

namespace wxd.files.winforms
{
    /// <summary>CSV/Excel阅读器窗体</summary>
    public partial class CsvReaderForm : Form
    {
        // ==================== 控件声明 ====================
        private SplitContainer _splitContainer = new();
        private TreeView _treeFiles = new();
        private TextBox _txtFileFilter = new();
        private TabControl _tabContent = new();
        private Panel _pnlLeft = new();
        private Panel _pnlLeftTop = new();
        private Panel _pnlLeftBottom = new();   // 底部状态栏
        private Label _lblFilter = new();
        private Button _btnRefresh = new();
        private Button _btnSelectFolder = new();
        private Label _lblFileCount = new();
        private ImageList _imgList = new();

        // ==================== 高级内容搜索控件 ====================
        private TextBox _txtContentSearch = new();
        private Button _btnContentSearch = new();
        private Button _btnCancelSearch = new();
        private ProgressBar _progressBar = new();
        private Panel _pnlSearchProgress = new();
        private RoundedPanel _pnlSearchContainer = new();

        // ==================== 状态字段 ====================
        private string _currentDir = "";
        private List<FileInfo> _allFiles = new();
        private CancellationTokenSource? _loadingCts;
        private readonly object _lockObj = new();
        private System.Windows.Forms.Timer? _filterDebounceTimer;
        private int _paneIndex = -1;                    // 所属窗格索引（-1表示独立打开，使用全局设置）
        private Dictionary<string, string> _fileRemarks = new(); // 文件备注映射（key: 相对路径）

        // ==================== 高级内容搜索状态 ====================
        private volatile bool _contentSearchCancelled;
        private volatile bool _isContentSearching = false;
        private string _lastContentSearchKeyword = "";

        // ==================== VCS 状态字段 ====================
        private VcsType _currentVcsType = VcsType.None;
        private bool _showConflictOnly = false;            // 是否只显示冲突文件

        // ==================== VCS 操作面板控件 ====================
        private RoundedPanel _pnlVcsActions = new();
        private Button _btnVcsUpdate = new();
        private CheckBox _chkShowConflictOnly = new();

        // ==================== 数据类 ====================
        private class FileInfo
        {
            public string Name { get; set; } = "";
            public string FullPath { get; set; } = "";
            public string RelativePath { get; set; } = "";
            public FileType Type { get; set; }
        }

        private enum FileType
        {
            Csv,
            Excel
        }

        // ==================== 圆角 Panel（用于搜索容器卡片）====================
        private class RoundedPanel : Panel
        {
            private int _radius = 6;

            [DefaultValue(6)]
            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public int Radius
            {
                get => _radius;
                set { _radius = value; Invalidate(); }
            }

            public RoundedPanel()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                         ControlStyles.ResizeRedraw | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.SupportsTransparentBackColor, true);
                BackColor = Color.FromArgb(227, 242, 253);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                var rect = new Rectangle(0, 0, Width - 1, Height - 1);
                using var path = GetRoundRect(rect, _radius);
                using var brush = new SolidBrush(BackColor);
                using var pen = new Pen(Color.FromArgb(173, 209, 237), 1);

                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }

            private static GraphicsPath GetRoundRect(Rectangle rect, int radius)
            {
                var path = new GraphicsPath();
                var r = radius;
                path.AddArc(rect.X, rect.Y, r, r, 180, 90);
                path.AddArc(rect.Right - r, rect.Y, r, r, 270, 90);
                path.AddArc(rect.Right - r, rect.Bottom - r, r, r, 0, 90);
                path.AddArc(rect.X, rect.Bottom - r, r, r, 90, 90);
                path.CloseFigure();
                return path;
            }
        }

        // ==================== 配色方案（浅色天蓝色调 · 1px圆角边框）====================
        private static readonly Color _colorBg = Color.FromArgb(235, 248, 255);
        private static readonly Color _colorLightBg = Color.FromArgb(235, 248, 255);
        private static readonly Color _colorSurface = Color.White;
        private static readonly Color _colorPanelBg = Color.FromArgb(245, 250, 255);
        private static readonly Color _colorBorder = Color.FromArgb(173, 209, 237);
        private static readonly Color _colorLightBorder = Color.FromArgb(200, 225, 245);
        private static readonly Color _colorDarkBorder = Color.FromArgb(130, 180, 220);
        private static readonly Color _colorLightText = Color.FromArgb(20, 60, 100);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(90, 130, 170);
        private static readonly Color _colorPrimary = Color.FromArgb(66, 165, 245);
        private static readonly Color _colorPrimaryLight = Color.FromArgb(227, 242, 253);
        private static readonly Color _colorPrimaryDark = Color.FromArgb(30, 136, 229);
        private static readonly Color _colorSuccess = Color.FromArgb(22, 163, 74);
        private static readonly Color _colorHoverBg = Color.FromArgb(235, 245, 255);
        private static readonly Color _colorHeaderBg = Color.FromArgb(240, 246, 252);
        private static readonly Color _colorAltRow = Color.FromArgb(248, 252, 255);
        private static readonly Color _colorSelectedBg = Color.FromArgb(227, 242, 253);
        private static readonly Color _colorGridLine = Color.FromArgb(200, 225, 245);
        private static readonly Color _colorRowNumBg = Color.FromArgb(240, 246, 252);

        // ==================== Win32 API（系统图标）====================
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr SHGetFileInfo(
            string pszPath, uint dwFileAttributes,
            out SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
        }

        private const uint SHGFI_ICON = 0x100;
        private const uint SHGFI_SMALLICON = 0x000000001;
        private const uint SHGFI_USEFILEATTRIBUTES = 0x10;

        public CsvReaderForm(string? initialDir = null, int paneIndex = -1)
        {
            _currentDir = initialDir ?? "";
            _paneIndex = paneIndex;
            InitializeComponent();
            CreateControls();
            ApplyTheme();

            // 窗体加载完成后设置splitter距离，并补充 VCS 状态初始化
            this.Shown += async (s, e) =>
            {
                await Task.Delay(200);
                if (_splitContainer.SplitterDistance < 300)
                {
                    _splitContainer.SplitterDistance = 300;
                }
                // 确保 VCS 状态已初始化（构造函数中可能因 Handle 未创建而跳过）
                if (_currentVcsType == VcsType.None && !string.IsNullOrEmpty(_currentDir))
                {
                    _ = InitializeVcsStatusAsync();
                }
            };

            // 如果有初始目录，自动加载
            if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
            {
                LoadFilesAsync();
            }

            UpdateTitle();
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            ClientSize = new Size(1200, 800);
            Text = "CSV/Excel阅读器";
            MinimumSize = new Size(800, 800);
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = _colorBg;
            Padding = new Padding(0);

            ResumeLayout(false);
        }

        private void CreateControls()
        {
            SuspendLayout();

            // ==================== 左侧面板 ====================
            _pnlLeft.Dock = DockStyle.Fill;
            _pnlLeft.BackColor = _colorSurface;
            _pnlLeft.Padding = new Padding(0);

            // 左侧顶部工具栏
            _pnlLeftTop.Height = 196;
            _pnlLeftTop.Dock = DockStyle.Top;
            _pnlLeftTop.BackColor = _colorSurface;
            _pnlLeftTop.Padding = new Padding(8, 6, 8, 4);

            // 选择目录按钮
            _btnSelectFolder = new Button
            {
                Text = "选择目录",
                Location = new Point(8, 6),
                Size = new Size(70, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = _colorPrimary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 1, BorderColor = _colorPrimaryDark, MouseOverBackColor = _colorPrimaryDark }
            };
            _btnSelectFolder.Click += OnSelectFolderClick;

            // 刷新按钮
            _btnRefresh = new Button
            {
                Text = "↻",
                Location = new Point(84, 6),
                Size = new Size(32, 28),
                Font = new Font("Segoe UI", 12f),
                BackColor = _colorHoverBg,
                ForeColor = _colorLightTextSecondary,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 1, BorderColor = _colorLightBorder, MouseOverBackColor = Color.FromArgb(200, 225, 245) }
            };
            _btnRefresh.Click += OnRefreshClick;

            // 文件数量标签（放在左侧底部状态栏）
            _lblFileCount = new Label
            {
                Text = "共 0 个文件",
                AutoSize = false,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 0, 0, 0)
            };

            // 过滤文本框
            _txtFileFilter = new TextBox
            {
                Location = new Point(8, 42),
                Size = new Size(250, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                PlaceholderText = "支持 || 或 && 运算符",
                BackColor = Color.FromArgb(242, 249, 255)
            };
            // 创建过滤防抖定时器（300ms 延迟）
            _filterDebounceTimer = new System.Windows.Forms.Timer { Interval = 300 };
            _filterDebounceTimer.Tick += (s, e) =>
            {
                _filterDebounceTimer!.Stop();
                ApplyFileFilter();
            };

            _txtFileFilter.TextChanged += OnFileFilterChanged;
            _txtFileFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    _txtFileFilter.Text = "";
                    e.Handled = true;
                }
            };

            // 使用说明按钮（放在刷新按钮右侧）
            var btnHelp = new Button
            {
                Text = "使用说明",
                Location = new Point(122, 6),
                Size = new Size(75, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.FromArgb(120, 144, 156),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(96, 125, 139) }
            };
            btnHelp.Click += OnHelpClick;

            // ==================== 高级内容搜索（第三行） ====================
            // ==================== 高级内容搜索（第三行 - 圆角容器） ====================
            _pnlSearchContainer = new RoundedPanel
            {
                Location = new Point(8, 72),
                Size = new Size(260, 70),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };

            // 内容搜索输入框
            _txtContentSearch = new TextBox
            {
                Location = new Point(8, 8),
                Size = new Size(140, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                PlaceholderText = "搜索文件内容...",
                BackColor = Color.FromArgb(242, 249, 255)
            };

            // 搜索按钮（靠右排列）
            _btnContentSearch = new Button
            {
                Text = "🔍搜索",
                Location = new Point(0, 4),
                Size = new Size(64, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = _colorPrimary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                FlatAppearance = { BorderSize = 1, BorderColor = _colorPrimaryDark, MouseOverBackColor = _colorPrimaryDark }
            };
            _btnContentSearch.Click += OnContentSearchClick;

            // 搜索框回车键触发搜索
            _txtContentSearch.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    OnContentSearchClick(_btnContentSearch, EventArgs.Empty);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };

            // ==================== 搜索进度面板（进度条+取消按钮同一行，初始隐藏） ====================
            _pnlSearchProgress = new Panel
            {
                Location = new Point(8, 38),
                Size = new Size(260, 28),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Visible = false,
                BackColor = Color.Transparent
            };

            _progressBar = new ProgressBar
            {
                Location = new Point(0, 5),
                Size = new Size(200, 18),
                Style = ProgressBarStyle.Continuous,
                ForeColor = _colorPrimary,
                BackColor = Color.FromArgb(235, 248, 255),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            // 取消搜索按钮（默认隐藏，放在进度面板内部右侧）
            _btnCancelSearch = new Button
            {
                Text = "清除",
                Size = new Size(64, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.FromArgb(239, 68, 68),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Visible = false,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(220, 38, 38), MouseOverBackColor = Color.FromArgb(220, 38, 38) }
            };
            _btnCancelSearch.Click += OnCancelSearchClick;

            _pnlSearchProgress.Controls.Add(_progressBar);
            _pnlSearchProgress.Controls.Add(_btnCancelSearch);

            _pnlSearchContainer.Controls.Add(_txtContentSearch);
            _pnlSearchContainer.Controls.Add(_btnContentSearch);
            _pnlSearchContainer.Controls.Add(_pnlSearchProgress);

            // ==================== VCS 操作面板（第四行 - 圆角容器，初始隐藏） ====================
            _pnlVcsActions = new RoundedPanel
            {
                Location = new Point(8, 148),
                Size = new Size(260, 36),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Visible = false,
                BackColor = Color.FromArgb(245, 250, 240)
            };

            _btnVcsUpdate = new Button
            {
                Text = "🔄 更新",
                Location = new Point(8, 4),
                Size = new Size(100, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.FromArgb(52, 168, 83),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(40, 140, 60), MouseOverBackColor = Color.FromArgb(40, 140, 60) }
            };
            _btnVcsUpdate.Click += OnVcsUpdateClick;

            _chkShowConflictOnly = new CheckBox
            {
                Text = "仅显示冲突",
                Location = new Point(116, 6),
                Size = new Size(100, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                Checked = false,
                AutoCheck = false
            };
            _chkShowConflictOnly.Click += OnShowConflictOnlyClick;

            _pnlVcsActions.Controls.Add(_btnVcsUpdate);
            _pnlVcsActions.Controls.Add(_chkShowConflictOnly);

            _pnlLeftTop.Controls.AddRange(new Control[] { _btnSelectFolder, _btnRefresh, _txtFileFilter, btnHelp, _pnlSearchContainer, _pnlVcsActions });

            // 文件树
            // 初始化图标列表
            _imgList = new ImageList();
            _imgList.ImageSize = new Size(16, 16);
            _imgList.ColorDepth = ColorDepth.Depth32Bit;
            // 添加默认文件夹图标（使用系统图标）
            var folderShfi = default(SHFILEINFO);
            SHGetFileInfo("folder", FILE_ATTRIBUTE_DIRECTORY, out folderShfi, (uint)Marshal.SizeOf(folderShfi),
                          SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
            if (folderShfi.hIcon != IntPtr.Zero)
                _imgList.Images.Add("folder", Icon.FromHandle(folderShfi.hIcon));
            // 添加默认CSV文件图标
            var csvShfi = default(SHFILEINFO);
            SHGetFileInfo(".csv", 0, out csvShfi, (uint)Marshal.SizeOf(csvShfi),
                          SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
            if (csvShfi.hIcon != IntPtr.Zero)
                _imgList.Images.Add("default", Icon.FromHandle(csvShfi.hIcon));

            _treeFiles = new TreeView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f),
                FullRowSelect = true,
                HideSelection = false,
                ImageList = _imgList,
                BackColor = _colorSurface,
                ForeColor = Color.FromArgb(30, 30, 30),
                ShowPlusMinus = true,
                ShowLines = true,
                ShowRootLines = true,
                Indent = 19,
                ItemHeight = 22,
                PathSeparator = "\\",
                DrawMode = TreeViewDrawMode.OwnerDrawText
            };
            // 节点展开时动态加载子目录
            _treeFiles.BeforeExpand += OnTreeBeforeExpand;
            _treeFiles.NodeMouseClick += OnTreeViewClick;
            _treeFiles.NodeMouseDoubleClick += OnTreeViewDoubleClick;
            _treeFiles.MouseDown += OnTreeViewMouseDown;
            _treeFiles.DrawNode += OnTreeViewDrawNode;

            // 左侧底部状态栏（必须先创建再添加）
            _pnlLeftBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 26,
                BackColor = Color.FromArgb(240, 244, 248),
                Padding = new Padding(0)
            };
            _pnlLeftBottom.Controls.Add(_lblFileCount);

            // 添加TreeView到左侧面板（注意顺序：Top > Fill > Bottom）
            _pnlLeft.Controls.Add(_pnlLeftBottom);
            _pnlLeft.Controls.Add(_treeFiles);
            _pnlLeft.Controls.Add(_pnlLeftTop);

            // ==================== 右侧标签页 ====================
            _tabContent = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9f),
                ItemSize = new Size(200, 28),
                SizeMode = TabSizeMode.Fixed,
                Multiline = true,
                BackColor = _colorSurface,
                Padding = new Point(12, 4)
            };

            // ▼ 下拉按钮（列出所有打开的标签页）
            var btnTabMenu = new Button
            {
                Text = "▼",
                Size = new Size(24, 24),
                Location = new Point(0, 2),
                Font = new Font("Microsoft YaHei UI", 8f),
                FlatStyle = FlatStyle.Flat,
                BackColor = _colorSurface,
                ForeColor = Color.FromArgb(107, 114, 128),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Padding = new Padding(0),
                Margin = new Padding(0),
                FlatAppearance = { BorderSize = 0 }
            };
            btnTabMenu.Click += (s, e) => ShowTabMenu(btnTabMenu);

            // 标签页面板：TabControl + ▼按钮（覆盖在右上角）
            var tabPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = _colorSurface
            };
            tabPanel.SizeChanged += (s, e) =>
            {
                btnTabMenu.Location = new Point(tabPanel.Width - 28, 2);
            };
            tabPanel.Controls.Add(_tabContent);
            tabPanel.Controls.Add(btnTabMenu);
            btnTabMenu.BringToFront();

            // ==================== 分割容器 ====================
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 400,
                FixedPanel = FixedPanel.Panel1,
                BackColor = _colorLightBorder
            };
            _splitContainer.Panel1.Controls.Add(_pnlLeft);
            _splitContainer.Panel2.Controls.Add(tabPanel);

            Controls.Add(_splitContainer);

            // 调整大小处理
            _pnlLeftTop.Resize += (s, e) =>
            {
                var w = _pnlLeftTop.Width - 16;
                _txtFileFilter.Width = Math.Max(100, w);

                // 圆角搜索容器自适应宽度
                _pnlSearchContainer.Width = Math.Max(180, w);

                // 搜索框和按钮在容器内部靠右布局
                var right = _pnlSearchContainer.Width - 8;
                _btnContentSearch.Location = new Point(right - 64, 6);
                _txtContentSearch.Width = Math.Max(60, right - 72 - 8);

                // 进度条面板在容器内部自适应
                _pnlSearchProgress.Width = Math.Max(160, _pnlSearchContainer.Width - 16);
                _progressBar.Width = Math.Max(60, _txtContentSearch.Width);
                _btnCancelSearch.Location = new Point(_pnlSearchProgress.Width - 64, 1);

                // VCS 操作面板自适应宽度
                _pnlVcsActions.Width = Math.Max(180, w);
                // 更新按钮和冲突复选框在容器内部靠左布局
                _btnVcsUpdate.Location = new Point(8, 4);
                _chkShowConflictOnly.Location = new Point(_btnVcsUpdate.Right + 8, 5);
            };

            ResumeLayout(false);
        }

        private void ApplyTheme()
        {
            BackColor = _colorBg;
            _pnlLeft.BackColor = _colorSurface;
            _treeFiles.BackColor = _colorSurface;
            _treeFiles.ForeColor = Color.FromArgb(30, 30, 30);
            _tabContent.BackColor = _colorSurface;
        }

        // ==================== 目录图标常量 ====================
        private const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;

        // ==================== 树形展开 ====================
        /// <summary>获取系统图标的索引（返回ImageList中的索引）</summary>
        private int GetSystemIconIndex(string path, bool isDir)
        {
            var shfi = default(SHFILEINFO);
            uint flags = SHGFI_ICON | SHGFI_SMALLICON;
            string searchPath;

            if (isDir)
            {
                // 文件夹使用特殊路径 + FILE_ATTRIBUTE_DIRECTORY 获取
                searchPath = path;
                SHGetFileInfo(searchPath, FILE_ATTRIBUTE_DIRECTORY, out shfi, (uint)Marshal.SizeOf(shfi), flags);
            }
            else
            {
                // 文件使用扩展名获取
                var ext = Path.GetExtension(path).ToLowerInvariant();
                searchPath = string.IsNullOrEmpty(ext) ? ".csv" : ext;
                SHGetFileInfo(searchPath, 0, out shfi, (uint)Marshal.SizeOf(shfi), flags | SHGFI_USEFILEATTRIBUTES);
            }

            if (shfi.hIcon != IntPtr.Zero)
            {
                var key = isDir ? "dir_" + path.GetHashCode() : "file_" + Path.GetExtension(path).ToLowerInvariant();
                if (!_imgList.Images.ContainsKey(key))
                {
                    _imgList.Images.Add(key, Icon.FromHandle(shfi.hIcon));
                }
                return _imgList.Images.IndexOfKey(key);
            }
            return isDir ? _imgList.Images.IndexOfKey("folder") : _imgList.Images.IndexOfKey("default");
        }

        /// <summary>展开节点前，动态加载其子目录和文件</summary>
        private void OnTreeBeforeExpand(object? sender, TreeViewCancelEventArgs e)
        {
            var node = e.Node;
            var fullPath = node.Tag as string ?? "";

            // 如果节点已经有子节点且第一个子节点是占位的虚节点，则清除并真正加载
            if (node.Nodes.Count == 1 && node.Nodes[0].Tag is string placeholder && placeholder == "__placeholder__")
            {
                node.Nodes.Clear();

                _treeFiles.BeginUpdate();
                try
                {
                    LoadChildrenToNode(node, fullPath);
                }
                finally
                {
                    _treeFiles.EndUpdate();
                }

                // 加载子节点后：
                // 只刷新当前展开目录本身的 VCS 状态（子文件夹展开时再单独刷新）
                // 注意：VCS 状态刷新是异步的，这里不立即更新节点文本
                // 等 OnVcsStatusRefreshed 事件触发后再更新 UI，避免同步 File.Exists 开销
                if (!string.IsNullOrEmpty(fullPath) && Directory.Exists(fullPath))
                {
                    _ = VcsStatusCache.RefreshAsync(fullPath);
                }
            }
        }

        /// <summary>获取文件在树中的显示文本（有备注时显示"备注名 (文件名)"）</summary>
        private string GetFileDisplayText(string fileFullPath)
        {
            var fileName = Path.GetFileName(fileFullPath);
            var relativePath = GetRelativePath(_currentDir, fileFullPath);
            if (_fileRemarks.TryGetValue(relativePath, out var remark) && !string.IsNullOrEmpty(remark))
                return remark + " (" + fileName + ")";
            return fileName;
        }

        /// <summary>递归加载目录到树节点</summary>
        private void LoadChildrenToNode(TreeNode parentNode, string dirPath)
        {
            try
            {
                var showHidden = CsvSettings.ShowHiddenFiles;

                // ★ 一次性获取所有子目录
                var subDirs = GetSubDirectories(dirPath);

                // ★ 只扫描匹配扩展名的文件（减少不相关文件的扫描）
                var csvFiles = new List<string>();
                var xlsFiles = new List<string>();
                var xlsxFiles = new List<string>();

                foreach (var file in Directory.EnumerateFiles(dirPath, "*.csv"))
                {
                    if (!showHidden) { try { if ((File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { } }
                    csvFiles.Add(file);
                }
                foreach (var file in Directory.EnumerateFiles(dirPath, "*.xls"))
                {
                    if (!showHidden) { try { if ((File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { } }
                    xlsFiles.Add(file);
                }
                foreach (var file in Directory.EnumerateFiles(dirPath, "*.xlsx"))
                {
                    if (!showHidden) { try { if ((File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { } }
                    xlsxFiles.Add(file);
                }

                // 加载子目录 — 对每个子目录轻量检查是否有任何内容（文件或子目录）
                foreach (var subDir in subDirs)
                {
                    var dirName = Path.GetFileName(subDir);
                    var dirNode = new TreeNode(dirName)
                    {
                        Tag = subDir,
                        ImageKey = "folder",
                        SelectedImageKey = "folder"
                    };
                    dirNode.ImageIndex = GetSystemIconIndex(subDir, true);
                    dirNode.SelectedImageIndex = dirNode.ImageIndex;

                    // 轻量检查子目录是否有任何条目（文件或子目录），有则添加占位节点
                    // EnumerateFileSystemEntries.Any() 会在找到第一个条目时停止枚举，非常高效
                    try
                    {
                        if (Directory.EnumerateFileSystemEntries(subDir).Any())
                            dirNode.Nodes.Add(new TreeNode("__placeholder__") { Tag = "__placeholder__" });
                    }
                    catch { }

                    parentNode.Nodes.Add(dirNode);
                }

                // 加载CSV文件
                foreach (var file in csvFiles)
                {
                    var fileNode = new TreeNode(GetFileDisplayText(file))
                    {
                        Tag = file,
                        ImageKey = "default",
                        SelectedImageKey = "default"
                    };
                    fileNode.ImageIndex = GetSystemIconIndex(file, false);
                    fileNode.SelectedImageIndex = fileNode.ImageIndex;
                    parentNode.Nodes.Add(fileNode);
                }

                // 加载Excel文件
                foreach (var file in xlsFiles.Concat(xlsxFiles))
                {
                    var fileNode = new TreeNode(GetFileDisplayText(file))
                    {
                        Tag = file,
                        ImageKey = "default",
                        SelectedImageKey = "default"
                    };
                    fileNode.ImageIndex = GetSystemIconIndex(file, false);
                    fileNode.SelectedImageIndex = fileNode.ImageIndex;
                    parentNode.Nodes.Add(fileNode);
                }
            }
            catch (UnauthorizedAccessException) { }
            catch (Exception) { }
        }

        /// <summary>获取目录下的子目录，自动忽略 .svn、.git、.wxd-work 等非内容文件夹，并根据全局设置过滤隐藏文件夹</summary>
        private static string[] GetSubDirectories(string path)
        {
            if (!Directory.Exists(path)) return Array.Empty<string>();
            var showHidden = CsvSettings.ShowHiddenFiles;
            return Directory.GetDirectories(path)
                .Where(d =>
                {
                    var name = Path.GetFileName(d);
                    if (IsVcsDir(name)) return false;
                    if (!showHidden)
                    {
                        try { return (File.GetAttributes(d) & FileAttributes.Hidden) != FileAttributes.Hidden; }
                        catch { return true; }
                    }
                    return true;
                })
                .ToArray();
        }

        /// <summary>检查目录名是否为版本控制目录或项目配置目录</summary>
        private static bool IsVcsDir(string dirName)
        {
            return dirName == ".svn" || dirName == ".git" || dirName == ".wxd-work";
        }

        /// <summary>检查目录下是否有符合条件的文件或子目录</summary>
        private bool HasChildFilesOrDirectories(string dirPath)
        {
            try
            {
                if (GetSubDirectories(dirPath).Length > 0)
                    return true;
                var showHidden = CsvSettings.ShowHiddenFiles;
                if (Directory.GetFiles(dirPath, "*.csv").Any(f => showHidden || !IsFileHidden(f)))
                    return true;
                if (Directory.GetFiles(dirPath, "*.xls").Any(f => showHidden || !IsFileHidden(f)))
                    return true;
                if (Directory.GetFiles(dirPath, "*.xlsx").Any(f => showHidden || !IsFileHidden(f)))
                    return true;
            }
            catch { }
            return false;
        }

        private static bool IsFileHidden(string path)
        {
            try { return (File.GetAttributes(path) & FileAttributes.Hidden) == FileAttributes.Hidden; }
            catch { return false; }
        }

        private void OnSelectFolderClick(object? sender, EventArgs e)
        {
            // 退出冲突模式
            if (_showConflictOnly)
            {
                _showConflictOnly = false;
                _chkShowConflictOnly.Checked = false;
            }

            using var dialog = new FolderBrowserDialog
            {
                Description = "选择包含CSV/Excel文件的目录",
                ShowNewFolderButton = false
            };

            if (!string.IsNullOrEmpty(_currentDir))
                dialog.SelectedPath = _currentDir;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                VcsStatusCache.Invalidate(_currentDir); // 清除旧目录缓存
                _currentDir = dialog.SelectedPath;
                UpdateTitle();
                LoadFilesAsync();
            }
        }

        private void UpdateTitle()
        {
            if (!string.IsNullOrEmpty(_currentDir))
                Text = $"CSV/Excel阅读器 - {_currentDir}";
            else
                Text = "CSV/Excel阅读器";
        }

        // ==================== VCS 状态集成 ====================
        /// <summary>初始化 VCS 状态检测（异步刷新，事件驱动 UI 更新）</summary>
        private async Task InitializeVcsStatusAsync()
        {
            if (string.IsNullOrEmpty(_currentDir)) return;
            if (IsDisposed) return;

            _currentVcsType = VersionControlHelper.DetectVcs(_currentDir);

            // 使用事件派发器 AddListener/RemoveListener 替代 static event +=/-=，
            // 避免多个 CsvReaderForm 实例间误删其他实例的订阅
            VcsStatusCache.RemoveListener(OnVcsStatusRefreshed);

            if (_currentVcsType != VcsType.None)
            {
                // 显示 VCS 操作面板
                ShowVcsPanel(true);

                // 先订阅事件再触发刷新，避免竞态条件（刷新完成但事件还没订阅）
                VcsStatusCache.AddListener(OnVcsStatusRefreshed);
                _ = VcsStatusCache.RefreshAsync(_currentDir);
            }
            else
            {
                // 无 VCS，隐藏 VCS 操作面板
                ShowVcsPanel(false);
            }
        }

        /// <summary>显示/隐藏 VCS 操作面板</summary>
        private void ShowVcsPanel(bool show)
        {
            if (InvokeRequired)
            {
                try { Invoke(() => ShowVcsPanel(show)); } catch { }
                return;
            }
            _pnlVcsActions.Visible = show;
            if (show)
            {
                // 更新按钮文本
                _btnVcsUpdate.Text = _currentVcsType == VcsType.Svn ? "🔄 SVN更新" : "🔄 Git拉取";
            }
        }

        /// <summary>
        /// VCS 状态刷新完成回调 — 更新 TreeView 节点文本。
        /// 不仅处理 _currentDir（根目录）的刷新，也处理子文件夹的独立缓存刷新。
        /// </summary>
        private void OnVcsStatusRefreshed(object? sender, VcsDirectoryStatus status)
        {
            var refreshDir = status.DirectoryPath;

            // 只关心当前树中已展开的目录（_currentDir 及其子目录）
            if (!refreshDir.StartsWith(_currentDir, StringComparison.OrdinalIgnoreCase))
                return;

            if (InvokeRequired)
            {
                try { Invoke(() => OnVcsStatusRefreshed(sender, status)); }
                catch { }
                return;
            }

            if (IsDisposed || _treeFiles.IsDisposed) return;

            // 如果刷新失败，输出调试信息
            if (!string.IsNullOrEmpty(status.ErrorMessage))
            {
                Debug.WriteLine($"[VcsStatusCache] 刷新失败: {status.ErrorMessage}");
                return;
            }

            // 检测是否有冲突文件
            bool hasConflictedFiles = VcsStatusCache.HasAnyConflictedFiles(_currentDir);
            _chkShowConflictOnly.Visible = hasConflictedFiles || _showConflictOnly;

            // 如果在冲突过滤模式，刷新冲突列表
            if (_showConflictOnly && string.Equals(refreshDir, _currentDir, StringComparison.OrdinalIgnoreCase))
            {
                ApplyFileFilter();
            }

            // 找到对应的树节点（从根开始查找）
            var targetNode = FindTreeNodeByPath(_treeFiles.Nodes, refreshDir);
            if (targetNode == null)
            {
                // 如果是根目录刷新，更新整个树
                if (string.Equals(refreshDir, _currentDir, StringComparison.OrdinalIgnoreCase))
                {
                    _treeFiles.BeginUpdate();
                    UpdateNodeTextWithVcsStatus(_treeFiles.Nodes);
                    _treeFiles.EndUpdate();
                    _treeFiles.Invalidate(); // 触发 DrawNode 重绘，使用新颜色
                }
                // 非根目录找不到节点（节点可能还未展开），不做更新
                return;
            }

            // 更新该节点下的所有文件节点的文本（添加 VCS 状态前缀）
            _treeFiles.BeginUpdate();
            UpdateNodeTextWithVcsStatus(targetNode.Nodes);
            _treeFiles.EndUpdate();

            // 同时更新该节点本身的文本（如果是文件夹且没有直接文件，文本不会变化，这没问题）
            UpdateNodeTextWithVcsStatusForSingleNode(targetNode);
            _treeFiles.Invalidate(); // 触发 DrawNode 重绘，使用新颜色
        }

        /// <summary>更新单个树节点的文本（如果是文件则添加 VCS 状态前缀）</summary>
        private void UpdateNodeTextWithVcsStatusForSingleNode(TreeNode node)
        {
            var fullPath = node.Tag as string;
            if (!string.IsNullOrEmpty(fullPath) && File.Exists(fullPath))
            {
                var status = VcsStatusCache.GetFileStatus(fullPath);
                var prefix = VcsStatusCache.GetStatusPrefix(status);
                var baseText = GetFileDisplayText(fullPath);
                node.Text = prefix + baseText;
            }
        }

        /// <summary>在树节点集合中查找指定路径的节点（递归）</summary>
        private static TreeNode? FindTreeNodeByPath(TreeNodeCollection nodes, string targetPath)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Tag is string nodePath &&
                    string.Equals(nodePath, targetPath, StringComparison.OrdinalIgnoreCase))
                    return node;
                if (node.Nodes.Count > 0)
                {
                    var found = FindTreeNodeByPath(node.Nodes, targetPath);
                    if (found != null) return found;
                }
            }
            return null;
        }

        /// <summary>递归更新树节点文本，添加 VCS 状态前缀</summary>
        private void UpdateNodeTextWithVcsStatus(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                var fullPath = node.Tag as string;
                if (!string.IsNullOrEmpty(fullPath) && (File.Exists(fullPath) || Directory.Exists(fullPath)))
                {
                    var status = VcsStatusCache.GetFileStatus(fullPath);
                    var prefix = VcsStatusCache.GetStatusPrefix(status);
                    var baseText = GetFileDisplayText(fullPath);
                    node.Text = prefix + baseText;
                }
                if (node.Nodes.Count > 0)
                {
                    UpdateNodeTextWithVcsStatus(node.Nodes);
                }
            }
        }

        /// <summary>OwnerDrawText 绘制节点 — 根据 VCS 状态设置文字颜色</summary>
        private void OnTreeViewDrawNode(object? sender, DrawTreeNodeEventArgs e)
        {
            // 如果节点没有 tag 或者不是文件路径，使用默认绘制
            // 注意：避免在绘制中调用 File.Exists（文件节点由 LoadChildrenToNode 添加，存在性已确认）
            if (e.Node.Tag is not string fullPath)
            {
                e.DrawDefault = true;
                return;
            }

            // 背景：选中/非选中
            var backColor = (e.State & TreeNodeStates.Selected) != 0
                ? _colorSelectedBg
                : _treeFiles.BackColor;

            // 获取 VCS 状态对应的颜色
            var status = VcsStatusCache.GetFileStatus(fullPath);
            var textColor = VcsStatusCache.GetStatusColor(status);

            // 绘制背景
            using var bgBrush = new SolidBrush(backColor);
            e.Graphics.FillRectangle(bgBrush, e.Bounds);

            // 绘制文字
            var text = e.Node.Text;
            var textRect = e.Bounds;
            // 缩进图标宽度
            textRect.X += 2;
            textRect.Width -= 2;

            TextRenderer.DrawText(e.Graphics, text, _treeFiles.Font, textRect, textColor,
                backColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
                TextFormatFlags.NoPrefix | TextFormatFlags.PreserveGraphicsClipping);
        }

        private void OnRefreshClick(object? sender, EventArgs e)
        {
            // 退出冲突模式
            if (_showConflictOnly)
            {
                _showConflictOnly = false;
                _chkShowConflictOnly.Checked = false;
            }
            if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
            {
                LoadFilesAsync();
            }
        }

        private void OnHelpClick(object? sender, EventArgs e)
        {
            const string helpText = @"📊 CSV/Excel 阅读器 - 使用说明
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

◆ 基本操作

  · 【选择目录】—— 选择包含 CSV 或 Excel 文件的文件夹
    （程序会自动递归扫描所有子目录中的表格文件）

  · 【刷新 ↻】—— 重新加载当前目录的文件列表
    （当目录内文件有变动时使用）

  · 【过滤文件】—— 输入关键词实时筛选文件名
    ├ 支持 || 运算符：匹配任意关键词（或）
    ├ 支持 && 运算符：同时匹配所有关键词（与）
    ├ 同时匹配文件名和备注名
    └ 按 Esc 键快速清空过滤，恢复完整目录树


◆ 高级内容搜索

  · 在文件列表上方的搜索框中输入关键词，点击搜索按钮
  · 异步遍历所有 CSV/Excel 文件内容，查找包含关键词的文件
  · 搜索过程中显示进度条，可随时点击取消按钮中断搜索
  · 搜索结果直接替换文件树，只显示匹配的文件
  · 取消搜索后自动恢复完整文件列表
  · 支持回车键快捷触发搜索


◆ 文件树操作

  · 以树形结构展示目录及子目录下所有 CSV/Excel 文件
  · 展开文件夹节点时异步加载子目录（延迟加载，避免卡顿）
  · 自动忽略 .svn / .git 等版本控制目录
  · 单击/双击文件节点在右侧标签页中打开文件

  · 右键菜单功能：
    ├ 打开文件：在右侧标签页中打开
    ├ 编辑文件：用系统默认程序打开（如 Excel）
    ├ 打开所在文件夹：在资源管理器中定位文件
    └ 设置备注名：为文件设置易识别的别名


◆ 支持的文件格式

  · CSV 文件（.csv）
    ├ 自动识别分隔符：逗号、制表符、分号等
    └ 集成 ExcelDataReader，支持多种编码格式

  · Excel 文件（.xls / .xlsx）
    └ 读取第一个工作表的内容


◆ 标签页操作

  · 点击标签页标题切换已打开的文件
  · 点击标签页上的 × 关闭文件
  · 右键标签页弹出菜单：关闭当前 / 关闭其它 / 关闭全部
  · 点击右上角 ▼ 按钮可快速定位到任意打开的标签页
  · 支持同时打开多个文件，无数量限制
  · 有备注名的文件，标签页标题显示备注名


◆ 表格操作

  · 【排序】—— 点击列标题按该列升序/降序排列
    （再次点击切换排序方向）

  · 【行过滤】—— 在表格上方过滤框中输入关键词筛选行
    ├ 支持 || 运算符（或）
    └ 支持 && 运算符（与）

  · 【行号列】—— 左侧显示数据行号，方便定位

  · 【列选择】—— 点击列选择按钮勾选要显示/隐藏的列
    └ 列显隐状态自动保存到目录级配置中

  · 【重新加载】—— 点击重新加载按钮刷新当前表格数据

  · 【右键菜单】（点击单元格弹出）
    ├ 复制单元格：复制当前单元格内容
    ├ 编辑单元格：直接编辑单元格内容
    ├ 查看关联表：以独立窗口打开关联的目标文件并高亮匹配行
    └ 固定Tips：将当前单元格的关联提示信息钉在表格上


◆ 列关联（跨文件列关联分析）

  · 在列选择对话框中为每列配置关联目标
  · 设置关联的目标文件、目标列，以及可选的显示列
  · 显示列映射：将关联列的 ID 映射为目标文件中易读的名称
    （例如：技能ID 4002 映射为 火球术）
  · 鼠标悬停单元格时弹出 tooltip，显示关联信息
  · 匹配的单元格以绿色高亮，不匹配的以红色高亮
  · 关联表可通过右键菜单在独立窗口打开，支持高亮定位


◆ 版本历史对比

  · 打开 SVN 或 Git 管理的 CSV/Excel 文件时
  · 表格顶部自动显示历史记录按钮
  · 点击后打开版本对比窗口，展示所有提交历史
  · 选中两条记录即可对比任意两个版本的差异
  · 差异单元格以高亮颜色标记


◆ 文件备注

  · 为文件设置别名备注，便于识别
  · 树形列表和标签页标题都会显示备注名


◆ 目录级配置

  · 冻结行、表头行、列关联、列显隐、文件备注等配置信息
  · 自动存储在目录下的 .wxd-work/config.json 文件中
  · 可随文件夹一起移植到其他位置，保持配置不变";

            HelpForm.ShowHelp(this, "使用说明 - CSV/Excel 阅读器", helpText, CsvSettings.ThemeMode == 0);
        }

        // ==================== VCS 操作按钮事件 ====================

        /// <summary>点击"更新"按钮 — 执行 SVN update 或 Git pull</summary>
        private async void OnVcsUpdateClick(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentDir) || _currentVcsType == VcsType.None)
            {
                AlertForm.Show(this, "当前目录不受 SVN 或 Git 管控，无法执行更新操作。",
                    "提示", AlertType.Information);
                return;
            }

            _btnVcsUpdate.Enabled = false;
            _btnVcsUpdate.Text = "正在更新...";

            try
            {
                var vcsName = _currentVcsType == VcsType.Svn ? "SVN" : "Git";
                var result = await VersionControlHelper.UpdateAsync(_currentDir, _currentVcsType);

                if (!result.Started)
                {
                    AlertForm.Show(this, result.ErrorMessage ?? "更新失败", "错误", AlertType.Error);
                    return;
                }

                string msg;
                AlertType alertType;

                if (result.HasConflicts)
                {
                    var conflictSummary = string.Join(GlobalConst.NewLine, result.ConflictLines.Take(20));
                    if (result.ConflictLines.Count > 20)
                        conflictSummary += $"{GlobalConst.NewLine}... 以及另外 {result.ConflictLines.Count - 20} 行冲突信息";

                    msg = $"{vcsName} 更新完成，但检测到 {result.ConflictLines.Count} 个冲突。{GlobalConst.NewLine}请手动解决冲突后提交。{GlobalConst.NewLine}{GlobalConst.NewLine}冲突详情：{GlobalConst.NewLine}{conflictSummary}";
                    alertType = AlertType.Warning;
                }
                else if (result.Success)
                {
                    msg = $"{vcsName} 更新成功!{GlobalConst.NewLine}{GlobalConst.NewLine}{result.Output}";
                    if (result.Output.Length > 2000)
                        msg = msg.Substring(0, 2000) + $"{GlobalConst.NewLine}{GlobalConst.NewLine}...(输出过长已截断)";
                    alertType = AlertType.Information;
                }
                else
                {
                    var errMsg = $"{vcsName} 更新失败:{GlobalConst.NewLine}{result.Error}";
                    AlertForm.Show(this, errMsg, "更新失败", AlertType.Warning);
                    return;
                }

                if (result.HasConflicts)
                {
                    _chkShowConflictOnly.Visible = true;
                }

                AlertForm.Show(this, msg, "更新完成", alertType);

                // 刷新 VCS 状态和文件列表
                VcsStatusCache.Invalidate(_currentDir);
                await VcsStatusCache.RefreshAsync(_currentDir);
                OnRefreshClick(null, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                AlertForm.Show(this, $"更新操作异常: {ex.Message}", "错误", AlertType.Error, exception: ex);
            }
            finally
            {
                _btnVcsUpdate.Enabled = true;
                _btnVcsUpdate.Text = "🔄 更新";
            }
        }

        /// <summary>点击"仅显示冲突"复选框 — 切换冲突过滤模式</summary>
        private async void OnShowConflictOnlyClick(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentDir) || _currentVcsType == VcsType.None)
            {
                AlertForm.Show(this, "当前目录不受 SVN 或 Git 管控。",
                    "提示", AlertType.Information);
                return;
            }

            // 切换状态
            if (_showConflictOnly)
            {
                // 退出冲突模式
                _showConflictOnly = false;
                _chkShowConflictOnly.Checked = false;
                _txtFileFilter.Text = "";
                ApplyFileFilter();
                return;
            }

            // 进入冲突模式：检查是否有冲突文件
            bool hasConflicted = false;

            foreach (var file in _allFiles)
            {
                var status = VcsStatusCache.GetFileStatus(file.FullPath);
                if (status == FileVcsStatus.Conflicted)
                {
                    hasConflicted = true;
                    break;
                }
            }

            if (!hasConflicted)
            {
                AlertForm.Show(this, "当前目录下没有检测到冲突文件。",
                    "提示", AlertType.Information);
                return;
            }

            _showConflictOnly = true;
            _chkShowConflictOnly.Checked = true;
            _txtFileFilter.Text = "";
            ApplyFileFilter();
        }

        // ==================== 在 TabPage 中打开冲突解决面板 ====================
        /// <summary>在右侧 TabPage 中打开冲突解决面板</summary>
        private async Task OpenConflictResolveInTab(string filePath)
        {
            if (!System.IO.File.Exists(filePath)) return;

            // 检查是否已打开该文件的冲突解决页签
            var conflictTag = $"__conflict__{filePath}";
            foreach (TabPage tab in _tabContent.TabPages)
            {
                if (tab.Tag is string tag && tag == conflictTag)
                {
                    _tabContent.SelectedTab = tab;
                    return;
                }
            }

            var dir = Path.GetDirectoryName(filePath) ?? "";
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            int frozenRows, headerRowIndex, uniqueColumnIndex, mappingRowIndex;
            if (ext == ".csv")
            {
                frozenRows = CsvSettings.GetCsvFrozenRowsByDirectory(dir);
                headerRowIndex = CsvSettings.GetCsvHeaderRowIndexByDirectory(dir);
                uniqueColumnIndex = CsvSettings.GetCsvUniqueColumnIndexByDirectory(dir);
                mappingRowIndex = CsvSettings.GetCsvMappingRowIndexByDirectory(dir);
            }
            else
            {
                frozenRows = CsvSettings.GetExcelFrozenRowsByDirectory(dir);
                headerRowIndex = CsvSettings.GetExcelHeaderRowIndexByDirectory(dir);
                uniqueColumnIndex = CsvSettings.GetExcelUniqueColumnIndexByDirectory(dir);
                mappingRowIndex = CsvSettings.GetExcelMappingRowIndexByDirectory(dir);
            }

            try
            {
                // 加载本地版本
                var localData = await Task.Run(() => ConflictResolveControl.LoadDataTableFromFile(filePath));
                if (localData == null)
                {
                    AlertForm.Show(this, $"无法读取文件: {filePath}", "错误", AlertType.Error);
                    return;
                }

                // 加载远程版本
                var tempDir = Path.Combine(Path.GetTempPath(), $"wxd_conflict_load_{Guid.NewGuid():N}");
                Directory.CreateDirectory(tempDir);

                string? remoteRevision = null;
                if (_currentVcsType == VcsType.Git)
                    remoteRevision = "MERGE_HEAD";
                else if (_currentVcsType == VcsType.Svn)
                    remoteRevision = "BASE";

                DataTable? remoteData = null;
                if (!string.IsNullOrEmpty(remoteRevision))
                {
                    var exportResult = await VersionControlHelper.ExportRevisionAsync(
                        filePath, _currentVcsType, remoteRevision, tempDir);
                    if (exportResult.Success && exportResult.TempFile != null && File.Exists(exportResult.TempFile))
                    {
                        remoteData = await Task.Run(() => ConflictResolveControl.LoadDataTableFromFile(exportResult.TempFile));
                    }
                }

                if (remoteData == null)
                    remoteData = localData.Copy();

                // 创建冲突解决控件
                var control = new ConflictResolveControl(filePath, _currentVcsType,
                    remoteData, localData,
                    frozenRows, headerRowIndex,
                    rootDir: _currentDir,
                    uniqueColumnIndex: uniqueColumnIndex,
                    mappingRowIndex: mappingRowIndex)
                {
                    Dock = DockStyle.Fill
                };

                var tabPage = new TabPage($"⚠️ 冲突: {Path.GetFileName(filePath)}")
                {
                    Tag = conflictTag,
                    ToolTipText = filePath
                };
                tabPage.Controls.Add(control);
                _tabContent.TabPages.Add(tabPage);
                _tabContent.SelectedTab = tabPage;
                AddCloseButtonToTab(tabPage);

                // 监听保存完成事件
                control.SaveCompleted += (s, e) =>
                {
                    // 刷新 VCS 状态
                    VcsStatusCache.Invalidate(_currentDir);
                    _ = VcsStatusCache.RefreshAsync(_currentDir);

                    // 如果在冲突模式，刷新冲突列表
                    if (_showConflictOnly)
                    {
                        ApplyFileFilter();
                    }
                };

                // 异步加载对比数据
                await control.LoadComparisonAsync();

                // 清理临时文件
                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }
            catch (Exception ex)
            {
                AlertForm.Show(this, $"打开冲突解决面板失败: {ex.Message}",
                    "错误", AlertType.Error, exception: ex);
            }
        }

        // ==================== 异步加载文件 ====================
        private async void LoadFilesAsync()
        {
            // 取消之前的加载任务
            _loadingCts?.Cancel();
            _loadingCts = new CancellationTokenSource();
            var token = _loadingCts.Token;

            _treeFiles.Nodes.Clear();
            _allFiles.Clear();

            try
            {
                await Task.Run(() =>
                {
                    // 扫描所有文件（保持与之前一致，用于过滤和关联表）
                    var files = new List<FileInfo>();
                    ScanDirectory(_currentDir, _currentDir, files, token);
                    lock (_lockObj)
                    {
                        _allFiles = files;
                    }
                }, token);

                if (!token.IsCancellationRequested)
                {
                    this.Invoke(() =>
                    {
                        _fileRemarks = CsvSettings.GetFileRemarksByDirectory(_currentDir);
                        BuildTreeFromRoot();
                        ApplyFileFilter();
                        // VCS 状态检测
                        _ = InitializeVcsStatusAsync();
                    });
                }
            }
            catch (OperationCanceledException)
            {
                // 忽略取消异常
            }
            catch (Exception ex)
            {
                this.Invoke(() =>
                {
                    AlertForm.Show(this, $"加载文件失败:" + GlobalConst.NewLine + ex.Message, "错误", AlertType.Error, exception: ex);
                });
            }
        }

        private void ScanFiles(string rootDir, CancellationToken token)
        {
            var files = new List<FileInfo>();
            ScanDirectory(rootDir, rootDir, files, token);
            lock (_lockObj)
            {
                _allFiles = files;
            }
        }

        private void ScanDirectory(string currentDir, string rootDir, List<FileInfo> files, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            try
            {
                var showHidden = CsvSettings.ShowHiddenFiles;

                // 获取当前目录的CSV文件
                foreach (var file in Directory.GetFiles(currentDir, "*.csv"))
                {
                    token.ThrowIfCancellationRequested();
                    if (!showHidden)
                    {
                        try { if ((File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                    }
                    var relativePath = GetRelativePath(rootDir, file);
                    files.Add(new FileInfo
                    {
                        Name = Path.GetFileName(file),
                        FullPath = file,
                        RelativePath = relativePath,
                        Type = FileType.Csv
                    });
                }

                // 获取当前目录的Excel文件
                foreach (var ext in new[] { "*.xls", "*.xlsx" })
                {
                    foreach (var file in Directory.GetFiles(currentDir, ext))
                    {
                        token.ThrowIfCancellationRequested();
                        if (!showHidden)
                        {
                            try { if ((File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                        }
                        var relativePath = GetRelativePath(rootDir, file);
                        files.Add(new FileInfo
                        {
                            Name = Path.GetFileName(file),
                            FullPath = file,
                            RelativePath = relativePath,
                            Type = FileType.Excel
                        });
                    }
                }

                // 递归子目录（GetSubDirectories 已内置隐藏过滤）
                foreach (var subDir in GetSubDirectories(currentDir))
                {
                    token.ThrowIfCancellationRequested();
                    ScanDirectory(subDir, rootDir, files, token);
                }
            }
            catch (UnauthorizedAccessException) { }
            catch (Exception) { }
        }

        private string GetRelativePath(string basePath, string fullPath)
        {
            if (fullPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
            {
                var relative = fullPath.Substring(basePath.Length).TrimStart(Path.DirectorySeparatorChar);
                return relative;
            }
            return fullPath;
        }

        // ==================== 构建树形结构 ====================
        private void BuildTreeFromRoot()
        {
            _treeFiles.BeginUpdate();
            _treeFiles.Nodes.Clear();

            if (string.IsNullOrEmpty(_currentDir) || !Directory.Exists(_currentDir))
            {
                _treeFiles.EndUpdate();
                _lblFileCount.Text = "共 0 个文件";
                return;
            }

            // 创建根节点（使用目录名）
            var rootNode = new TreeNode(Path.GetFileName(_currentDir))
            {
                Tag = _currentDir,
                ImageKey = "folder",
                SelectedImageKey = "folder"
            };
            rootNode.ImageIndex = GetSystemIconIndex(_currentDir, true);
            rootNode.SelectedImageIndex = rootNode.ImageIndex;

            // 添加占位节点以显示展开箭头（延迟加载）
            rootNode.Nodes.Add(new TreeNode("__placeholder__") { Tag = "__placeholder__" });

            _treeFiles.Nodes.Add(rootNode);
            rootNode.Expand();

            _treeFiles.EndUpdate();

            _lblFileCount.Text = $"共 {_allFiles.Count} 个文件";
        }

        // ==================== 文件过滤 ====================
        private void OnFileFilterChanged(object? sender, EventArgs e)
        {
            _filterDebounceTimer?.Stop();
            _filterDebounceTimer?.Start();
        }

        private void ApplyFileFilter()
        {
            var filter = _txtFileFilter.Text.Trim();

            // 冲突过滤模式
            if (_showConflictOnly)
            {
                // 过滤出所有冲突文件
                var conflictedFiles = _allFiles
                    .Where(f => VcsStatusCache.GetFileStatus(f.FullPath) == FileVcsStatus.Conflicted)
                    .OrderBy(f => f.RelativePath)
                    .ToList();

                if (conflictedFiles.Count == 0)
                {
                    _treeFiles.BeginUpdate();
                    _treeFiles.Nodes.Clear();
                    _treeFiles.EndUpdate();
                    _lblFileCount.Text = "无冲突文件";
                    return;
                }

                // 如果还有文本过滤，取交集
                if (!string.IsNullOrEmpty(filter))
                {
                    var processedFilter = filter.Replace("||", " | ").Replace("&&", " & ");
                    conflictedFiles = conflictedFiles
                        .Where(f =>
                        {
                            var remark = _fileRemarks.TryGetValue(f.RelativePath, out var r) ? r : "";
                            return MatchesFilter(f.RelativePath, processedFilter) ||
                                   (!string.IsNullOrEmpty(remark) && MatchesFilter(remark, processedFilter));
                        })
                        .ToList();
                }

                BuildFilteredTreeFromFiles(conflictedFiles);
                _lblFileCount.Text = $"冲突文件: {conflictedFiles.Count} 个（总{_allFiles.Count}）";
                return;
            }

            // 正常过滤逻辑
            if (string.IsNullOrEmpty(filter))
            {
                BuildTreeFromRoot();
            }
            else
            {
                var processedFilter = filter.Replace("||", " | ").Replace("&&", " & ");
                var matchedFiles = _allFiles
                    .Where(f =>
                    {
                        var remark = _fileRemarks.TryGetValue(f.RelativePath, out var r) ? r : "";
                        return MatchesFilter(f.RelativePath, processedFilter) ||
                               (!string.IsNullOrEmpty(remark) && MatchesFilter(remark, processedFilter));
                    })
                    .OrderBy(f => f.RelativePath)
                    .ToList();

                BuildFilteredTreeFromFiles(matchedFiles);
            }
        }

        /// <summary>根据过滤后的文件列表构建全展开的过滤树</summary>
        private void BuildFilteredTreeFromFiles(List<FileInfo> matchedFiles)
        {
            _treeFiles.BeginUpdate();
            _treeFiles.Nodes.Clear();

            if (matchedFiles.Count == 0)
            {
                _treeFiles.EndUpdate();
                _lblFileCount.Text = $"共 0 个文件（总{_allFiles.Count}）";
                return;
            }

            // 使用字典按目录路径组织节点，key 为目录路径
            var dirNodes = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
            var rootDir = _currentDir;

            // 创建根节点
            var rootName = Path.GetFileName(rootDir);
            var rootNode = new TreeNode(rootName)
            {
                Tag = rootDir,
                ImageKey = "folder",
                SelectedImageKey = "folder"
            };
            rootNode.ImageIndex = GetSystemIconIndex(rootDir, true);
            rootNode.SelectedImageIndex = rootNode.ImageIndex;
            _treeFiles.Nodes.Add(rootNode);
            dirNodes[rootDir] = rootNode;

            foreach (var file in matchedFiles)
            {
                var dirPath = Path.GetDirectoryName(file.FullPath) ?? rootDir;
                TreeNode parentNode = GetOrCreateParentNode(dirPath, rootDir, dirNodes);

                var fileNode = new TreeNode(GetFileDisplayText(file.FullPath))
                {
                    Tag = file.FullPath,
                    ImageKey = "default",
                    SelectedImageKey = "default"
                };
                fileNode.ImageIndex = GetSystemIconIndex(file.FullPath, false);
                fileNode.SelectedImageIndex = fileNode.ImageIndex;
                parentNode.Nodes.Add(fileNode);
            }

            // 展开所有节点
            _treeFiles.ExpandAll();

            _treeFiles.EndUpdate();
            _lblFileCount.Text = $"共 {matchedFiles.Count} 个文件（总{_allFiles.Count}）";
        }

        /// <summary>获取或创建文件所在的父目录节点（递归创建缺失的中间目录）</summary>
        private TreeNode GetOrCreateParentNode(string dirPath, string rootDir, Dictionary<string, TreeNode> dirNodes)
        {
            if (dirNodes.TryGetValue(dirPath, out var node))
                return node;

            // 递归创建父节点
            var parentDir = Path.GetDirectoryName(dirPath);
            if (parentDir == null || string.Equals(dirPath, rootDir, StringComparison.OrdinalIgnoreCase))
                return dirNodes[rootDir];

            // 如果父目录不在根目录范围内，挂在根下
            if (!parentDir.StartsWith(rootDir, StringComparison.OrdinalIgnoreCase))
                return dirNodes[rootDir];

            var parentNode = GetOrCreateParentNode(parentDir, rootDir, dirNodes);

            var thisNode = new TreeNode(Path.GetFileName(dirPath))
            {
                Tag = dirPath,
                ImageKey = "folder",
                SelectedImageKey = "folder"
            };
            thisNode.ImageIndex = GetSystemIconIndex(dirPath, true);
            thisNode.SelectedImageIndex = thisNode.ImageIndex;
            parentNode.Nodes.Add(thisNode);
            dirNodes[dirPath] = thisNode;
            return thisNode;
        }

        private bool MatchesFilter(string text, string filter)
        {
            if (string.IsNullOrEmpty(filter))
                return true;

            if (filter.Contains('|'))
            {
                var terms = filter.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                return terms.Any(t => text.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            else if (filter.Contains('&'))
            {
                var terms = filter.Split('&', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                return terms.All(t => text.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            else
            {
                return text.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        // ==================== 树文件选择处理 ====================
        private void OnTreeViewClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
            // 只响应左键点击，右键由 OnTreeViewMouseDown 处理
            if (e.Button != MouseButtons.Left) return;

            var fullPath = e.Node.Tag as string ?? "";
            if (System.IO.File.Exists(fullPath))
            {
                if (_showConflictOnly)
                {
                    // 冲突模式下点击文件打开冲突解决面板
                    _ = OpenConflictResolveInTab(fullPath);
                }
                else
                {
                    var fileType = DetectFileType(fullPath);
                    OpenFile(fullPath, fileType);
                }
            }
        }

        private void OnTreeViewDoubleClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
            // 只响应左键双击
            if (e.Button != MouseButtons.Left) return;

            var fullPath = e.Node.Tag as string ?? "";
            if (System.IO.File.Exists(fullPath))
            {
                if (_showConflictOnly)
                {
                    // 冲突模式下双击打开冲突解决面板
                    _ = OpenConflictResolveInTab(fullPath);
                }
                else
                {
                    var fileType = DetectFileType(fullPath);
                    OpenFile(fullPath, fileType);
                }
            }
        }

        /// <summary>文件树节点右键菜单</summary>
        private void OnTreeViewMouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            var hitInfo = _treeFiles.HitTest(e.X, e.Y);
            if (hitInfo?.Node == null) return;
            var node = hitInfo.Node;
            var fullPath = node.Tag as string ?? "";

            // 选中当前右键点击的节点
            _treeFiles.SelectedNode = node;

            var ctx = new ContextMenuStrip();
            ctx.Font = new Font("Microsoft YaHei UI", 9f);

            // 1. 设置别名（仅文件节点可用）
            var isFile = System.IO.File.Exists(fullPath);
            var mnuRemark = new ToolStripMenuItem("设置别名");
            mnuRemark.Enabled = isFile;
            mnuRemark.Click += (s, args) =>
            {
                if (!isFile) return;
                var relativePath = GetRelativePath(_currentDir, fullPath);
                var currentRemark = _fileRemarks.TryGetValue(relativePath, out var r) ? r : "";
                using var inputForm = new Form
                {
                    Text = "设置别名",
                    ClientSize = new Size(380, 170),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    Font = new Font("Microsoft YaHei UI", 9f)
                };
                var lbl = new Label
                {
                    Text = "为文件 \"" + Path.GetFileName(fullPath) + "\" 设置别名：（留空则清除别名）",
                    Location = new Point(12, 12), Size = new Size(356, 36),
                    AutoSize = false
                };
                var txt = new TextBox
                {
                    Text = currentRemark,
                    Location = new Point(12, 54), Size = new Size(356, 23),
                    BorderStyle = BorderStyle.FixedSingle
                };
                var btnOk = new Button
                {
                    Text = "确定", DialogResult = DialogResult.OK,
                    Location = new Point(190, 95), Size = new Size(80, 25),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(66, 165, 245), ForeColor = Color.White
                };
                var btnCancel = new Button
                {
                    Text = "清除", DialogResult = DialogResult.Cancel,
                    Location = new Point(280, 95), Size = new Size(80, 25),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(235, 248, 255), ForeColor = Color.FromArgb(20, 60, 100)
                };
                inputForm.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCancel });
                inputForm.AcceptButton = btnOk;
                inputForm.CancelButton = btnCancel;
                if (inputForm.ShowDialog(this) == DialogResult.OK)
                {
                    var trimmed = txt.Text.Trim();
                    if (trimmed != currentRemark)
                    {
                        CsvSettings.SetFileRemark(_currentDir, relativePath, trimmed);
                        if (string.IsNullOrEmpty(trimmed))
                            _fileRemarks.Remove(relativePath);
                        else
                            _fileRemarks[relativePath] = trimmed;
                        // 刷新树显示
                        _treeFiles.BeginUpdate();
                        node.Text = GetFileDisplayText(fullPath);
                        _treeFiles.EndUpdate();
                        // 同步更新已打开的页签标题
                        foreach (TabPage tab in _tabContent.TabPages)
                        {
                            if (tab.Tag is string tabPath && tabPath.Equals(fullPath, StringComparison.OrdinalIgnoreCase))
                            {
                                tab.Text = GetFileDisplayText(fullPath);
                                break;
                            }
                        }
                    }
                }
            };

            // 2. 打开表格（和左键单击一样）
            var mnuOpen = new ToolStripMenuItem("打开表格");
            mnuOpen.Enabled = System.IO.File.Exists(fullPath);
            mnuOpen.Click += (s, args) =>
            {
                if (System.IO.File.Exists(fullPath))
                    OpenFile(fullPath, DetectFileType(fullPath));
            };

            // 3. 编辑表格（用系统默认编辑器打开）
            var mnuEdit = new ToolStripMenuItem("编辑表格");
            mnuEdit.Enabled = System.IO.File.Exists(fullPath);
            mnuEdit.Click += (s, args) => OpenFileInSystemEditor(fullPath);

            // 4. 复制文件名（不含后缀）
            var mnuCopyFileName = new ToolStripMenuItem("复制文件名");
            mnuCopyFileName.Enabled = isFile;
            mnuCopyFileName.Click += (s, args) => CopyFileNameToClipboard(fullPath);

            // 5. 复制文件路径（完整路径）
            var mnuCopyFilePath = new ToolStripMenuItem("复制文件路径");
            mnuCopyFilePath.Enabled = isFile;
            mnuCopyFilePath.Click += (s, args) => CopyFilePathToClipboard(fullPath);

            // 6. 打开文件夹（定位到该文件）
            var mnuOpenFolder = new ToolStripMenuItem("打开所在文件夹");
            mnuOpenFolder.Enabled = !string.IsNullOrEmpty(fullPath);
            mnuOpenFolder.Click += (s, args) =>
            {
                try
                {
                    string folderPath = System.IO.File.Exists(fullPath)
                        ? System.IO.Path.GetDirectoryName(fullPath) ?? ""
                        : fullPath;
                    if (!string.IsNullOrEmpty(folderPath) && System.IO.Directory.Exists(folderPath))
                    {
                        // explorer /select,"filepath" 可定位到具体文件
                        string arg = System.IO.File.Exists(fullPath)
                            ? $"/select,\"{fullPath}\""
                            : $"\"{folderPath}\"";
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = "explorer",
                            Arguments = arg,
                            UseShellExecute = true
                        });
                    }
                }
                catch (Exception ex)
                {
                    AlertForm.Show(this, $"无法打开文件夹: {ex.Message}", "错误", AlertType.Error, exception: ex);
                }
            };

            // 7. 历史记录对比（始终显示，点击时检测 VCS 类型）
            var mnuHistory = new ToolStripMenuItem("历史记录对比");
            mnuHistory.Enabled = isFile;
            mnuHistory.Click += (s, args) =>
            {
                if (!System.IO.File.Exists(fullPath)) return;
                var vcsType = VersionControlHelper.DetectVcs(fullPath);
                if (vcsType == VcsType.None)
                {
                    AlertForm.Show(this,
                        "此文件不受 SVN 或 Git 管控，无法进行历史记录对比。",
                        "非受控文件", AlertType.Information);
                    return;
                }

                var ext = Path.GetExtension(fullPath).ToLowerInvariant();
                int frozenRows, headerRowIndex, uniqueColumnIndex, mappingRowIndex;
                if (ext == ".csv")
                {
                    frozenRows = CsvSettings.GetCsvFrozenRowsByDirectory(_currentDir);
                    headerRowIndex = CsvSettings.GetCsvHeaderRowIndexByDirectory(_currentDir);
                    uniqueColumnIndex = CsvSettings.GetCsvUniqueColumnIndexByDirectory(_currentDir);
                    mappingRowIndex = CsvSettings.GetCsvMappingRowIndexByDirectory(_currentDir);
                }
                else
                {
                    frozenRows = CsvSettings.GetExcelFrozenRowsByDirectory(_currentDir);
                    headerRowIndex = CsvSettings.GetExcelHeaderRowIndexByDirectory(_currentDir);
                    uniqueColumnIndex = CsvSettings.GetExcelUniqueColumnIndexByDirectory(_currentDir);
                    mappingRowIndex = CsvSettings.GetExcelMappingRowIndexByDirectory(_currentDir);
                }
                // 获取列可见性配置
                string relPath;
                if (!string.IsNullOrEmpty(_currentDir) && fullPath.StartsWith(_currentDir, StringComparison.OrdinalIgnoreCase))
                    relPath = fullPath.Substring(_currentDir.Length).TrimStart(Path.DirectorySeparatorChar);
                else
                    relPath = Path.GetFileName(fullPath);
                var columnVisibility = CsvSettings.GetColumnVisibility(_currentDir, relPath);
                var form = new VersionCompareForm(fullPath, vcsType, frozenRows, headerRowIndex,
                    rootDir: _currentDir, columnVisibility: columnVisibility,
                    uniqueColumnIndex: uniqueColumnIndex, mappingRowIndex: mappingRowIndex);
                form.Show();
            };

            // 8. 编辑冲突（仅冲突文件可用）
            var mnuEditConflict = new ToolStripMenuItem("编辑冲突");
            mnuEditConflict.Enabled = isFile && VcsStatusCache.GetFileStatus(fullPath) == FileVcsStatus.Conflicted;
            mnuEditConflict.Click += async (s, args) =>
            {
                if (!isFile) return;
                await OpenConflictResolveInTab(fullPath);
            };

            ctx.Items.AddRange(new ToolStripItem[]
            {
                mnuRemark,
                new ToolStripSeparator(),
                mnuOpen,
                mnuEdit,
                new ToolStripSeparator(),
                mnuCopyFileName,
                mnuCopyFilePath,
                new ToolStripSeparator(),
                mnuOpenFolder,
                mnuHistory,
                new ToolStripSeparator(),
                mnuEditConflict,
            });
            ctx.Show(_treeFiles, e.Location);
        }

        private static FileType DetectFileType(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".csv") return FileType.Csv;
            return FileType.Excel;
        }

        // ==================== 文件打开处理 ====================
        private void OpenFile(string filePath, FileType fileType)
        {
            // 检查是否已经打开
            foreach (TabPage tab in _tabContent.TabPages)
            {
                if (tab.Tag is string path && path.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                {
                    _tabContent.SelectedTab = tab;
                    return;
                }
            }

            var options = new TableBrowserOptions(filePath)
            {
                PaneIndex = _paneIndex,
                RootDir = _currentDir,
                InitialContentFilter = _lastContentSearchKeyword,
            };
            Control viewer = new TableBrowserControl(options) { Dock = DockStyle.Fill };

            // 创建新标签页（有备注时显示"备注名 (文件名)"）
            var tabPage = new TabPage(GetFileDisplayText(filePath))
            {
                Tag = filePath,
                ToolTipText = filePath
            };
            tabPage.Controls.Add(viewer);
            _tabContent.TabPages.Add(tabPage);
            _tabContent.SelectedTab = tabPage;
            AddCloseButtonToTab(tabPage);

            // 注册关联表事件
            if (viewer is TableBrowserControl browserControl)
            {
                browserControl.LinkTableRequested += OnViewerLinkTableRequested;
                browserControl.EditLinkTableRequested += OnViewerEditLinkTableRequested;
            }
        }

        /// <summary>处理查看控件的关联表请求 — 以独立新窗口打开关联文件</summary>
        private void OnViewerLinkTableRequested(object? sender, EventArgs e)
        {
            using var dialog = new LinkTableSelectorDialog(
                _allFiles.Select(f =>
                {
                    var display = GetFileDisplayText(f.FullPath);
                    return (display, f.FullPath);
                }).ToList());
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                var linkedPath = dialog.SelectedFilePath;
                if (string.IsNullOrEmpty(linkedPath) || !File.Exists(linkedPath)) return;

                // 在新独立窗口中打开关联文件，标题栏也使用备注名显示
                var displayName = GetFileDisplayText(linkedPath);
                var linkOptions = new TableBrowserOptions(linkedPath) { RootDir = _currentDir };
                var standaloneForm = new CsvReaderStandaloneForm(linkOptions, new LinkOpenOptions { DisplayName = displayName });
                standaloneForm.Show(this);
            }
        }

        /// <summary>处理编辑关联表请求 — 在当前窗体的标签页中打开关联文件</summary>
        private void OnViewerEditLinkTableRequested(object? sender, TableBrowserControl.EditLinkTableEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.TargetFilePath) && File.Exists(e.TargetFilePath))
            {
                OpenFile(e.TargetFilePath, DetectFileType(e.TargetFilePath));
            }
        }

        private void AddCloseButtonToTab(TabPage tabPage)
        {
            // 事件处理程序只在第一次添加标签页时注册一次
            if (!_tabCloseHandlerAdded)
            {
                _tabCloseHandlerAdded = true;

                // 使用TabControl的MouseDown事件处理关闭与右键菜单
                _tabContent.MouseDown += (s, e) =>
                {
                    for (int i = 0; i < _tabContent.TabPages.Count; i++)
                    {
                        var tabRect = _tabContent.GetTabRect(i);
                        var closeRect = new Rectangle(tabRect.Right - 16, tabRect.Top + 4, 12, 12);

                        if (e.Button == MouseButtons.Left && closeRect.Contains(e.Location))
                        {
                            CloseTab(i);
                            return;
                        }
                        if (e.Button == MouseButtons.Right && tabRect.Contains(e.Location))
                        {
                            ShowTabContextMenu(_tabContent.PointToScreen(e.Location), i);
                            return;
                        }
                    }
                };

                // 重绘标签页以显示关闭按钮和圆角
                _tabContent.DrawMode = TabDrawMode.OwnerDrawFixed;
                _tabContent.DrawItem += (s, e) =>
                {
                    var tabPage = _tabContent.TabPages[e.Index];
                    var tabRect = _tabContent.GetTabRect(e.Index);
                    bool isSelected = _tabContent.SelectedIndex == e.Index;

                    // 选中页签底部多延伸2px，与内容区无缝衔接
                    int bottomExtra = isSelected ? 2 : 0;
                    var roundRect = new Rectangle(
                        tabRect.X + 2, tabRect.Y + 2,
                        tabRect.Width - 4, tabRect.Height - 2 + bottomExtra);
                    int radius = 6;

                    Color bgColor, borderColor;
                    if (isSelected)
                    {
                        bgColor = Color.FromArgb(66, 165, 245);
                        borderColor = Color.FromArgb(30, 136, 229);
                    }
                    else
                    {
                        bgColor = Color.FromArgb(227, 242, 253);
                        borderColor = Color.FromArgb(173, 209, 237);
                    }

                    using (var path = new GraphicsPath())
                    {
                        // 顶部左圆角
                        path.AddArc(roundRect.X, roundRect.Y, radius * 2, radius * 2, 180, 90);
                        // 顶部右圆角
                        path.AddArc(roundRect.Right - radius * 2, roundRect.Y, radius * 2, radius * 2, 270, 90);
                        // 右侧直线
                        path.AddLine(roundRect.Right, roundRect.Y + radius, roundRect.Right, roundRect.Bottom);
                        // 底部直线（无圆角）
                        path.AddLine(roundRect.Right, roundRect.Bottom, roundRect.X, roundRect.Bottom);
                        // 左侧直线
                        path.AddLine(roundRect.X, roundRect.Bottom, roundRect.X, roundRect.Y + radius);
                        path.CloseFigure();

                        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

                        using (var brush = new SolidBrush(bgColor))
                        {
                            e.Graphics.FillPath(brush, path);
                        }

                        // 绘制圆角边框
                        using (var pen = new Pen(borderColor, 1f))
                        {
                            e.Graphics.DrawPath(pen, path);
                        }
                    }

                    // 绘制文本
                    var textRect = new Rectangle(tabRect.X + 10, tabRect.Y + 4, tabRect.Width - 32, tabRect.Height - 8);
                    Color textColor = isSelected ? Color.White : Color.FromArgb(20, 60, 100);
                    TextRenderer.DrawText(e.Graphics, tabPage.Text, _tabContent.Font, textRect, textColor,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

                    // 绘制关闭按钮
                    var closeRect = new Rectangle(tabRect.Right - 18, tabRect.Top + 6, 12, 12);
                    Color closeColor = isSelected ? Color.FromArgb(200, 220, 255) : Color.FromArgb(107, 114, 128);
                    using (var pen = new Pen(closeColor, 1.5f))
                    {
                        e.Graphics.DrawLine(pen, closeRect.Left + 2, closeRect.Top + 2, closeRect.Right - 2, closeRect.Bottom - 2);
                        e.Graphics.DrawLine(pen, closeRect.Right - 2, closeRect.Top + 2, closeRect.Left + 2, closeRect.Bottom - 2);
                    }

                    // 选中页签底部遮盖 TabControl 默认的分隔线
                    if (isSelected)
                    {
                        using (var coverPen = new Pen(bgColor, 2f))
                        {
                            e.Graphics.DrawLine(coverPen,
                                roundRect.X + 1, roundRect.Bottom,
                                roundRect.Right - 1, roundRect.Bottom);
                        }
                    }
                };
            }
        }

        private bool _tabCloseHandlerAdded = false;

        /// <summary>关闭指定页签，返回是否成功关闭（false = 用户取消）</summary>
        private bool CloseTab(int tabIndex)
        {
            if (tabIndex < 0 || tabIndex >= _tabContent.TabPages.Count)
                return false;

            var tabPage = _tabContent.TabPages[tabIndex];

            // ★ 关闭前检查是否有未保存的修改
            foreach (Control ctrl in tabPage.Controls)
            {
                if (ctrl is TableBrowserControl browser)
                {
                    if (!browser.ForceFlushBeforeClose())
                        return false;  // 用户选择不关闭
                    break;
                }
            }

            // 释放标签页中的控件资源
            foreach (Control ctrl in tabPage.Controls)
            {
                ctrl.Dispose();
            }

            _tabContent.TabPages.RemoveAt(tabIndex);

            // 清除TreeView选中状态，允许重新打开同一文件
            _treeFiles.SelectedNode = null;
            return true;
        }

        /// <summary>▼按钮下拉菜单 — 列出所有打开的标签页</summary>
        private void ShowTabMenu(Button btn)
        {
            var menu = new ContextMenuStrip();
            menu.Font = new Font("Microsoft YaHei UI", 9f);
            for (int i = 0; i < _tabContent.TabPages.Count; i++)
            {
                var tp = _tabContent.TabPages[i];
                bool isCurrent = (i == _tabContent.SelectedIndex);
                var item = new ToolStripMenuItem(tp.Text)
                {
                    Tag = i,
                    Checked = isCurrent,
                    CheckState = isCurrent ? CheckState.Checked : CheckState.Unchecked,
                    Font = new Font("Microsoft YaHei UI", 9f, isCurrent ? FontStyle.Bold : FontStyle.Regular)
                };
                item.Click += (s, e) =>
                {
                    if (s is ToolStripMenuItem mi && mi.Tag is int idx)
                        _tabContent.SelectedIndex = idx;
                };
                menu.Items.Add(item);
            }
            menu.Show(btn, new Point(0, btn.Height));
        }

        /// <summary>右键菜单 — 关闭页签</summary>
        private void ShowTabContextMenu(Point screenPoint, int tabIndex)
        {
            var menu = new ContextMenuStrip();
            menu.Font = new Font("Microsoft YaHei UI", 9f);

            // 获取当前页签的文件路径
            var tabPage = _tabContent.TabPages[tabIndex];
            var filePath = tabPage.Tag as string;
            var fileExists = !string.IsNullOrEmpty(filePath) && System.IO.File.Exists(filePath);

            // 1. 复制文件名（不含后缀）
            var mnuCopyFileName = new ToolStripMenuItem("复制文件名");
            mnuCopyFileName.Enabled = fileExists;
            mnuCopyFileName.Click += (s, e) => CopyFileNameToClipboard(filePath!);
            menu.Items.Add(mnuCopyFileName);

            // 2. 复制文件路径（完整路径）
            var mnuCopyFilePath = new ToolStripMenuItem("复制文件路径");
            mnuCopyFilePath.Enabled = fileExists;
            mnuCopyFilePath.Click += (s, e) => CopyFilePathToClipboard(filePath!);
            menu.Items.Add(mnuCopyFilePath);

            // 3. 编辑文件（用系统默认编辑器打开）
            var mnuEdit = new ToolStripMenuItem("编辑文件");
            mnuEdit.Enabled = fileExists;
            mnuEdit.Click += (s, e) => OpenFileInSystemEditor(filePath!);
            menu.Items.Add(mnuEdit);

            menu.Items.Add(new ToolStripSeparator());

            // 4. 关闭当前页签
            var itemClose = new ToolStripMenuItem("关闭当前页签");
            itemClose.Click += (s, e) => CloseTab(tabIndex);
            menu.Items.Add(itemClose);

            // 5. 关闭其它页签
            var itemCloseOthers = new ToolStripMenuItem("关闭其它页签");
            itemCloseOthers.Enabled = _tabContent.TabPages.Count > 1;
            itemCloseOthers.Click += (s, e) =>
            {
                for (int i = _tabContent.TabPages.Count - 1; i >= 0; i--)
                {
                    if (i != tabIndex)
                    {
                        if (!CloseTab(i))
                            break;  // 用户取消，停止关闭
                    }
                }
            };
            menu.Items.Add(itemCloseOthers);

            // 6. 关闭全部页签
            var itemCloseAll = new ToolStripMenuItem("关闭全部页签");
            itemCloseAll.Enabled = _tabContent.TabPages.Count > 0;
            itemCloseAll.Click += (s, e) => CloseAllTabs();
            menu.Items.Add(itemCloseAll);

            menu.Show(screenPoint);
        }

        private void CloseAllTabs()
        {
            while (_tabContent.TabPages.Count > 0)
            {
                if (!CloseTab(0))
                    break;  // 用户取消，停止关闭
            }
        }

        // ==================== 公共辅助方法 ====================

        /// <summary>复制文件名（不含后缀）到剪贴板</summary>
        private void CopyFileNameToClipboard(string filePath)
        {
            try
            {
                Clipboard.SetText(Path.GetFileNameWithoutExtension(filePath));
            }
            catch (Exception ex)
            {
                AlertForm.Show(this, $"复制失败: {ex.Message}", "错误", AlertType.Error, exception: ex);
            }
        }

        /// <summary>复制完整文件路径到剪贴板</summary>
        private void CopyFilePathToClipboard(string filePath)
        {
            try
            {
                Clipboard.SetText(filePath);
            }
            catch (Exception ex)
            {
                AlertForm.Show(this, $"复制失败: {ex.Message}", "错误", AlertType.Error, exception: ex);
            }
        }

        /// <summary>用系统默认编辑器打开文件</summary>
        private void OpenFileInSystemEditor(string filePath)
        {
            if (!System.IO.File.Exists(filePath)) return;
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                AlertForm.Show(this, $"无法打开文件: {ex.Message}", "错误", AlertType.Error, exception: ex);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 取消静态事件订阅，避免已关闭的实例仍被后台线程回调
            VcsStatusCache.RemoveListener(OnVcsStatusRefreshed);

            _loadingCts?.Cancel();
            _contentSearchCancelled = true;

            // 检查所有打开的标签页是否有未保存的修改
            for (int i = _tabContent.TabPages.Count - 1; i >= 0; i--)
            {
                var tabPage = _tabContent.TabPages[i];
                foreach (Control ctrl in tabPage.Controls)
                {
                    if (ctrl is TableBrowserControl browser)
                    {
                        if (!browser.ForceFlushBeforeClose())
                        {
                            e.Cancel = true;  // 用户取消关闭
                            return;
                        }
                        break;
                    }
                }
            }

            base.OnFormClosing(e);
        }

        // ==================== 高级内容搜索 ====================

        /// <summary>点击搜索按钮 — 异步遍历所有文件内容查找关键词</summary>
        private async void OnContentSearchClick(object? sender, EventArgs e)
        {
            var keyword = _txtContentSearch.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                AlertForm.Show(this, "请输入要搜索的内容关键词", "提示", AlertType.Information);
                return;
            }
            if (_allFiles.Count == 0)
            {
                AlertForm.Show(this, "请先选择目录加载文件", "提示", AlertType.Information);
                return;
            }
            if (_isContentSearching) return;

            _isContentSearching = true;
            _contentSearchCancelled = false;

            // 进入搜索模式：禁用搜索按钮，显示进度条和取消按钮
            _btnContentSearch.Enabled = false;
            _pnlSearchProgress.Visible = true;
            _btnCancelSearch.Visible = true;
            _progressBar.Value = 0;

            _lastContentSearchKeyword = keyword;

            var totalFiles = _allFiles.Count;
            var matchedFiles = new List<FileInfo>();
            var processed = 0;

            try
            {
                await Task.Run(() =>
                {
                    foreach (var file in _allFiles)
                    {
                        if (_contentSearchCancelled) break;

                        var current = Interlocked.Increment(ref processed);
                        this.Invoke(() =>
                        {
                            if (!_contentSearchCancelled)
                                UpdateSearchProgress(current, totalFiles);
                        });

                        if (FileContainsTextSimple(file.FullPath, file.Type, keyword))
                        {
                            lock (matchedFiles) matchedFiles.Add(file);
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                AlertForm.Show(this, $"搜索出错: {ex.Message}", "错误", AlertType.Error, exception: ex);
            }

            // 检查是否被取消
            if (_contentSearchCancelled)
            {
                // OnCancelSearchClick 已处理 UI
                _isContentSearching = false;
                return;
            }

            // 搜索正常完成：进度条到100%，恢复搜索按钮，进度条和取消按钮保持显示
            _progressBar.Value = 100;
            _btnContentSearch.Enabled = true;

            if (matchedFiles.Count > 0)
            {
                BuildFilteredTreeFromFiles(matchedFiles);
                _lblFileCount.Text = $"内容搜索: 匹配 {matchedFiles.Count} 个（总{_allFiles.Count}）";
            }
            else
            {
                _treeFiles.BeginUpdate();
                _treeFiles.Nodes.Clear();
                _treeFiles.EndUpdate();
                _lblFileCount.Text = $"内容搜索: 未找到包含\"{keyword}\"的文件（总{_allFiles.Count}）";
                AlertForm.Show(this, $"未找到包含\"{keyword}\"的文件", "搜索结果", AlertType.Information);
            }

            _isContentSearching = false;
        }

        /// <summary>点击取消按钮 — 取消正在进行的搜索，清空搜索框，恢复文件列表初始状态</summary>
        private void OnCancelSearchClick(object? sender, EventArgs e)
        {
            _contentSearchCancelled = true;
            // 清空搜索框
            _txtContentSearch.Text = "";
            _lastContentSearchKeyword = "";
            // 进度条归0，进度条和取消按钮保持显示
            _progressBar.Value = 0;
            // 恢复搜索按钮可用
            _btnContentSearch.Enabled = true;
            // 恢复文件数量标签
            _lblFileCount.Text = $"共 {_allFiles.Count} 个文件";
            _isContentSearching = false;
            // 文件列表还原初始状态
            BuildTreeFromRoot();
        }

        /// <summary>更新搜索进度（仅更新进度条）</summary>
        private void UpdateSearchProgress(int current, int total)
        {
            if (total <= 0) return;
            var pct = Math.Min(100, (int)((double)current / total * 100));
            _progressBar.Value = pct;
        }

        /// <summary>判断文件内容是否包含关键词（简化版，不依赖 CancellationToken）</summary>
        private static bool FileContainsTextSimple(string filePath, FileType fileType, string keyword)
        {
            try
            {
                if (fileType == FileType.Csv)
                {
                    return SearchCsvContentSimple(filePath, keyword);
                }
                else if (fileType == FileType.Excel)
                {
                    return SearchExcelContent(filePath, keyword);
                }
            }
            catch
            {
                // 跳过无法读取的文件
            }
            return false;
        }

        /// <summary>逐行搜索CSV文件内容（简化版，优先使用缓存）</summary>
        private static bool SearchCsvContentSimple(string filePath, string keyword)
        {
            // ★ 优先尝试从缓存读取
            try
            {
                var readOnlyData = FileDataCache.GetDataTableReadOnly(filePath);

                if (readOnlyData != null)
                {
                    foreach (DataRow row in readOnlyData.Rows)
                    {
                        foreach (var item in row.ItemArray)
                        {
                            var val = item?.ToString() ?? "";
                            if (val.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                                return true;
                        }
                    }
                    return false;
                }
            }
            catch { }

            // 回退到磁盘读取
            var encodings = new[] { Encoding.UTF8, Encoding.GetEncoding("GBK"), Encoding.Default };
            foreach (var encoding in encodings)
            {
                try
                {
                    using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var reader = new StreamReader(fs, encoding, true);
                    string? line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                    return false;
                }
                catch { continue; }
            }
            return false;
        }

        /// <summary>搜索Excel文件内容（优先使用缓存）</summary>
        private static bool SearchExcelContent(string filePath, string keyword)
        {
            // ★ 优先尝试从缓存读取
            try
            {
                var readOnlyData = FileDataCache.GetDataTableReadOnly(filePath);

                if (readOnlyData != null)
                {
                    foreach (DataRow row in readOnlyData.Rows)
                    {
                        foreach (var item in row.ItemArray)
                        {
                            var val = item?.ToString() ?? "";
                            if (val.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                                return true;
                        }
                    }
                    return false;
                }
            }
            catch { }

            // 回退到磁盘读取
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(fs);
            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = false
                }
            });

            if (dataSet.Tables.Count == 0)
                return false;

            var sheet = dataSet.Tables[0];
            for (int r = 0; r < sheet.Rows.Count; r++)
            {
                for (int c = 0; c < sheet.Columns.Count; c++)
                {
                    var val = sheet.Rows[r][c]?.ToString() ?? "";
                    if (val.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
            }
            return false;
        }
    }
}
