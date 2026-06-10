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

            // 窗体加载完成后设置splitter距离
            this.Shown += async (s, e) =>
            {
                await Task.Delay(200);
                if (_splitContainer.SplitterDistance < 300)
                {
                    _splitContainer.SplitterDistance = 300;
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
            _pnlLeftTop.Height = 150;
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

            _pnlLeftTop.Controls.AddRange(new Control[] { _btnSelectFolder, _btnRefresh, _txtFileFilter, btnHelp, _pnlSearchContainer });

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
                ForeColor = _colorLightText,
                ShowPlusMinus = true,
                ShowLines = true,
                ShowRootLines = true,
                Indent = 19,
                ItemHeight = 22,
                PathSeparator = "\\"
            };
            // 节点展开时动态加载子目录
            _treeFiles.BeforeExpand += OnTreeBeforeExpand;
            _treeFiles.NodeMouseClick += OnTreeViewClick;
            _treeFiles.NodeMouseDoubleClick += OnTreeViewDoubleClick;
            _treeFiles.MouseDown += OnTreeViewMouseDown;

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
            };

            ResumeLayout(false);
        }

        private void ApplyTheme()
        {
            BackColor = _colorBg;
            _pnlLeft.BackColor = _colorSurface;
            _treeFiles.BackColor = _colorSurface;
            _treeFiles.ForeColor = _colorLightText;
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
                LoadChildrenToNode(node, fullPath);
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
                // 加载子目录
                foreach (var subDir in GetSubDirectories(dirPath))
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

                    // 检查子目录下是否有符合条件的文件或子目录，有则添加占位节点以实现展开箭头
                    if (HasChildFilesOrDirectories(subDir))
                    {
                        dirNode.Nodes.Add(new TreeNode("__placeholder__") { Tag = "__placeholder__" });
                    }

                    parentNode.Nodes.Add(dirNode);
                }

                // 加载CSV文件
                var showHidden = CsvSettings.ShowHiddenFiles;
                foreach (var file in Directory.GetFiles(dirPath, "*.csv"))
                {
                    if (!showHidden)
                    {
                        try { if ((File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                    }
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
                foreach (var ext in new[] { "*.xls", "*.xlsx" })
                {
                    foreach (var file in Directory.GetFiles(dirPath, ext))
                    {
                        if (!showHidden)
                        {
                            try { if ((File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                        }
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
            }
            catch (UnauthorizedAccessException) { }
            catch (Exception) { }
        }

        /// <summary>获取目录下的子目录，自动忽略 .svn 和 .git 等版本控制文件夹，并根据全局设置过滤隐藏文件夹</summary>
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

        /// <summary>检查目录名是否为版本控制目录</summary>
        private static bool IsVcsDir(string dirName)
        {
            return dirName == ".svn" || dirName == ".git";
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
            using var dialog = new FolderBrowserDialog
            {
                Description = "选择包含CSV/Excel文件的目录",
                ShowNewFolderButton = false
            };

            if (!string.IsNullOrEmpty(_currentDir))
                dialog.SelectedPath = _currentDir;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
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

        private void OnRefreshClick(object? sender, EventArgs e)
        {
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
    ├ 打开关联表：以独立窗口打开关联的目标文件并高亮匹配行
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
                    MessageBox.Show(this, $"加载文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            if (string.IsNullOrEmpty(filter))
            {
                // 无过滤：重建正常的延迟加载树
                BuildTreeFromRoot();
            }
            else
            {
                // 有过滤：根据匹配文件重建内部全展开树（同时匹配文件名和备注名）
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
                var fileType = DetectFileType(fullPath);
                OpenFile(fullPath, fileType);
            }
        }

        private void OnTreeViewDoubleClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
            // 只响应左键双击
            if (e.Button != MouseButtons.Left) return;

            var fullPath = e.Node.Tag as string ?? "";
            if (System.IO.File.Exists(fullPath))
            {
                var fileType = DetectFileType(fullPath);
                OpenFile(fullPath, fileType);
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

            // 1. 打开文件（和左键单击一样）
            var mnuOpen = new ToolStripMenuItem("打开文件");
            mnuOpen.Enabled = System.IO.File.Exists(fullPath);
            mnuOpen.Click += (s, args) =>
            {
                if (System.IO.File.Exists(fullPath))
                    OpenFile(fullPath, DetectFileType(fullPath));
            };

            // 2. 编辑文件（用系统默认编辑器打开）
            var mnuEdit = new ToolStripMenuItem("编辑文件");
            mnuEdit.Enabled = System.IO.File.Exists(fullPath);
            mnuEdit.Click += (s, args) =>
            {
                if (System.IO.File.Exists(fullPath))
                {
                    try
                    {
                        // 直接用系统默认程序打开文件（无cmd窗口闪烁）
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = fullPath,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, $"无法打开文件: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };

            // 3. 打开文件夹（定位到该文件）
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
                    MessageBox.Show(this, $"无法打开文件夹: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // 4. 历史记录对比（始终显示，点击时检测 VCS 类型）
            var isFile = System.IO.File.Exists(fullPath);
            var mnuHistory = new ToolStripMenuItem("历史记录对比");
            mnuHistory.Enabled = isFile;
            mnuHistory.Click += (s, args) =>
            {
                if (!System.IO.File.Exists(fullPath)) return;
                var vcsType = VersionControlHelper.DetectVcs(fullPath);
                if (vcsType == VcsType.None)
                {
                    MessageBox.Show(this,
                        "此文件不受 SVN 或 Git 管控，无法进行历史记录对比。",
                        "非受控文件",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                var dir = Path.GetDirectoryName(fullPath) ?? "";
                var ext = Path.GetExtension(fullPath).ToLowerInvariant();
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
                // 获取列可见性配置
                string relPath;
                if (!string.IsNullOrEmpty(dir) && fullPath.StartsWith(dir, StringComparison.OrdinalIgnoreCase))
                    relPath = fullPath.Substring(dir.Length).TrimStart(Path.DirectorySeparatorChar);
                else
                    relPath = Path.GetFileName(fullPath);
                var columnVisibility = CsvSettings.GetColumnVisibility(dir, relPath);
                var form = new VersionCompareForm(fullPath, vcsType, frozenRows, headerRowIndex,
                    rootDir: _currentDir, columnVisibility: columnVisibility,
                    uniqueColumnIndex: uniqueColumnIndex, mappingRowIndex: mappingRowIndex);
                form.Show();
            };

            // 5. 设置备注名（仅文件节点可用）
            var mnuRemark = new ToolStripMenuItem("设置备注名");
            mnuRemark.Enabled = isFile;
            mnuRemark.Click += (s, args) =>
            {
                if (!isFile) return;
                var relativePath = GetRelativePath(_currentDir, fullPath);
                var currentRemark = _fileRemarks.TryGetValue(relativePath, out var r) ? r : "";
                using var inputForm = new Form
                {
                    Text = "设置备注名",
                    ClientSize = new Size(380, 170),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    Font = new Font("Microsoft YaHei UI", 9f)
                };
                var lbl = new Label
                {
                    Text = "为文件 \"" + Path.GetFileName(fullPath) + "\" 设置备注名：（留空则清除备注）",
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

            ctx.Items.AddRange(new ToolStripItem[] { mnuOpen, mnuEdit, mnuHistory, mnuOpenFolder, mnuRemark });
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
                browserControl.LinkTableRequested += OnViewerLinkTableRequested;
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

        private void CloseTab(int tabIndex)
        {
            if (tabIndex < 0 || tabIndex >= _tabContent.TabPages.Count)
                return;

            var tabPage = _tabContent.TabPages[tabIndex];

            // 释放标签页中的控件资源
            foreach (Control ctrl in tabPage.Controls)
            {
                ctrl.Dispose();
            }

            _tabContent.TabPages.RemoveAt(tabIndex);

            // 清除TreeView选中状态，允许重新打开同一文件
            _treeFiles.SelectedNode = null;
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

            var itemClose = new ToolStripMenuItem("关闭当前页签");
            itemClose.Click += (s, e) => CloseTab(tabIndex);
            menu.Items.Add(itemClose);

            var itemCloseOthers = new ToolStripMenuItem("关闭其它页签");
            itemCloseOthers.Enabled = _tabContent.TabPages.Count > 1;
            itemCloseOthers.Click += (s, e) =>
            {
                for (int i = _tabContent.TabPages.Count - 1; i >= 0; i--)
                {
                    if (i != tabIndex)
                        CloseTab(i);
                }
            };
            menu.Items.Add(itemCloseOthers);

            var itemCloseAll = new ToolStripMenuItem("关闭全部页签");
            itemCloseAll.Enabled = _tabContent.TabPages.Count > 0;
            itemCloseAll.Click += (s, e) => CloseAllTabs();
            menu.Items.Add(itemCloseAll);

            menu.Show(screenPoint);
        }

        private void CloseAllTabs()
        {
            while (_tabContent.TabPages.Count > 0)
                CloseTab(0);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _loadingCts?.Cancel();
            _contentSearchCancelled = true;
            base.OnFormClosing(e);
        }

        // ==================== 高级内容搜索 ====================

        /// <summary>点击搜索按钮 — 异步遍历所有文件内容查找关键词</summary>
        private async void OnContentSearchClick(object? sender, EventArgs e)
        {
            var keyword = _txtContentSearch.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                MessageBox.Show(this, "请输入要搜索的内容关键词", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_allFiles.Count == 0)
            {
                MessageBox.Show(this, "请先选择目录加载文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                MessageBox.Show(this, $"搜索出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show(this, $"未找到包含\"{keyword}\"的文件", "搜索结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        /// <summary>逐行搜索CSV文件内容（简化版）</summary>
        private static bool SearchCsvContentSimple(string filePath, string keyword)
        {
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

        /// <summary>搜索Excel文件内容</summary>
        private static bool SearchExcelContent(string filePath, string keyword)
        {
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


    /// <summary>自然字符串比较器，支持数字排序</summary>
    public class NaturalStringComparer : IComparer<string?>
    {
        private readonly SortOrder _sortOrder;

        public NaturalStringComparer(SortOrder sortOrder)
        {
            _sortOrder = sortOrder;
        }

        public int Compare(string? x, string? y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return _sortOrder == SortOrder.Ascending ? -1 : 1;
            if (y == null) return _sortOrder == SortOrder.Ascending ? 1 : -1;

            // 尝试数字比较
            if (double.TryParse(x, out double numX) && double.TryParse(y, out double numY))
            {
                var result = numX.CompareTo(numY);
                return _sortOrder == SortOrder.Ascending ? result : -result;
            }

            // 字符串比较
            var strResult = string.Compare(x, y, StringComparison.OrdinalIgnoreCase);
            return _sortOrder == SortOrder.Ascending ? strResult : -strResult;
        }
    }

    /// <summary>单元格编辑对话框</summary>
    public class CellEditDialog : Form
    {
        private RichTextBox _txtValue = new();
        private Button _btnSave = new();
        private Button _btnCancel = new();
        private Button _btnDefault = new();
        private Button _btnRaw = new();
        private Panel _pnlBottom = new();
        private Panel _pnlTop = new();
        private Label _lblSaveTime = new();  // 保存时间标签
        private Label _lblReadOnlyHint = new(); // 只读提示标签
        private readonly Action<string>? _onSave;

        public string CellValue => _txtValue.Text.Replace("\r\n", "\n");

        /// <summary>
        /// 将各种换行符转换为 Windows 格式 (\r\n)，供 RichTextBox 多行模式正确显示
        /// </summary>
        private static string NormalizeLineBreaksForDisplay(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            // 先将已有的 \r\n 保护，再处理 \r 和 \n，最后统一输出 \r\n
            return text.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
        }

        public CellEditDialog(string initialValue, int row, int col, string fileName = "", string colName = "",
            Action<string>? onSave = null, bool readOnly = false, string readOnlyMessage = "")
        {
            _onSave = onSave;
            InitializeComponent(readOnly, readOnlyMessage);
            _txtValue.Text = NormalizeLineBreaksForDisplay(initialValue);

            string colDisplay = string.IsNullOrEmpty(colName)
                ? (col >= 0 && col < 26 ? ((char)('A' + col)).ToString() : $"Col{col}")
                : colName;
            if (!string.IsNullOrEmpty(fileName))
                Text = $"单元格编辑 — {Path.GetFileName(fileName)} — 列：{colDisplay}，第 {row + 1} 行";
            else
                Text = $"编辑单元格 — 行 {row + 1}  列 {colDisplay}({col})";
        }

        private static void SetActiveButton(Button active, Button inactive)
        {
            active.BackColor = Color.FromArgb(66, 165, 245);
            active.ForeColor = Color.White;
            active.FlatAppearance.BorderSize = 1;
            active.FlatAppearance.BorderColor = Color.FromArgb(30, 136, 229);
            inactive.BackColor = Color.FromArgb(235, 248, 255);
            inactive.ForeColor = Color.FromArgb(20, 60, 100);
            inactive.FlatAppearance.BorderSize = 1;
            inactive.FlatAppearance.BorderColor = Color.FromArgb(173, 209, 237);
        }

        private void InitializeComponent(bool readOnly = false, string readOnlyMessage = "")
        {
            SuspendLayout();

            Text = "编辑单元格";
            ClientSize = new Size(800, 600);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = true;
            MaximizeBox = true;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);

            // ==================== 顶部格式切换面板 ====================
            _pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = Color.FromArgb(235, 245, 255),
                Padding = new Padding(8, 7, 8, 7)
            };

            _btnDefault = new Button
            {
                Text = "默认格式",
                Location = new Point(8, 7),
                Size = new Size(90, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 0, MouseOverBackColor = Color.FromArgb(227, 242, 253) }
            };

            _btnRaw = new Button
            {
                Text = "自动换行",
                Location = new Point(106, 7),
                Size = new Size(90, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 0, MouseOverBackColor = Color.FromArgb(227, 242, 253) }
            };

            _btnDefault.Click += (s, e) =>
            {
                SetActiveButton(_btnDefault, _btnRaw);
                _txtValue.WordWrap = false;
                _txtValue.ScrollBars = RichTextBoxScrollBars.Both;
            };
            _btnRaw.Click += (s, e) =>
            {
                SetActiveButton(_btnRaw, _btnDefault);
                _txtValue.WordWrap = true;
                _txtValue.ScrollBars = RichTextBoxScrollBars.Both;
            };

            // 默认选中"默认格式"
            SetActiveButton(_btnDefault, _btnRaw);

            _pnlTop.Controls.Add(_btnDefault);
            _pnlTop.Controls.Add(_btnRaw);

            // ==================== 底部按钮面板 ====================
            _pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 52,
                BackColor = Color.FromArgb(235, 245, 255)
            };

            // 只读提示标签（左侧）
            _lblReadOnlyHint = new Label
            {
                Text = readOnly ? readOnlyMessage : "",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                ForeColor = Color.Red,
                Location = new Point(10, 14)
            };

            // 保存按钮
            _btnSave = new Button
            {
                Text = "保存",
                Size = new Size(80, 30),
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229), MouseOverBackColor = Color.FromArgb(30, 136, 229) },
                Enabled = !readOnly
            };
            _btnSave.Click += (s, e) =>
            {
                _lblSaveTime.Text = $"已保存 {DateTime.Now:HH:mm:ss}";
                _onSave?.Invoke(CellValue);
            };

            // 保存时间标签
            _lblSaveTime = new Label
            {
                Text = "",
                Size = new Size(140, 28),
                Font = new Font("Microsoft YaHei UI", 8.5f),
                ForeColor = Color.FromArgb(76, 175, 80),
                TextAlign = ContentAlignment.MiddleRight
            };

            // 取消（关闭）按钮
            _btnCancel = new Button
            {
                Text = "关闭",
                Size = new Size(80, 30),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237), MouseOverBackColor = Color.FromArgb(200, 225, 245) }
            };
            _btnCancel.Click += (s, e) => { Close(); };

            _pnlBottom.Controls.Add(_lblReadOnlyHint);
            _pnlBottom.Controls.Add(_lblSaveTime);
            _pnlBottom.Controls.Add(_btnSave);
            _pnlBottom.Controls.Add(_btnCancel);

            // 窗口调整大小时重新定位按钮（保持在右下角）
            void LayoutButtons()
            {
                _btnCancel.Left = _pnlBottom.Width - _btnCancel.Width - 10;
                _btnSave.Left = _btnCancel.Left - _btnSave.Width - 8;
                _lblSaveTime.Left = _btnSave.Left - _lblSaveTime.Width - 8;
                _lblSaveTime.Top = 12;
                _btnSave.Top = 10;
                _btnCancel.Top = 10;
                _lblReadOnlyHint.Top = 14;
            }
            LayoutButtons();
            _pnlBottom.Resize += (s, e) => LayoutButtons();

            // ==================== 编辑区（RichTextBox）====================
            _txtValue = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10f),
                BorderStyle = BorderStyle.None,
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both,
                AcceptsTab = true,
                BackColor = Color.White
            };

            Controls.Add(_txtValue);
            Controls.Add(_pnlBottom);
            Controls.Add(_pnlTop);

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;

            ResumeLayout(false);
        }
    }


    /// <summary>列管理对话框 — ListView列显隐 + 关联设置</summary>
    public class ColumnManagerDialog : Form
    {
        private ListView _lvColumns = new();
        private Button _btnSelectAll = new();
        private Button _btnDeselectAll = new();
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private Button _btnLinkManage = new();
        private int _lockedColumnIndex = -1;
        private bool[] _checkedStates;
        private string _rootDir = "";
        private string _currentFilePath = "";
        private int _frozenRows = 0;
        private int _headerRowIndex = -1;
        private int _mappingRowIndex = -1;
        private List<(string display, string fullPath)> _fileEntries = new();
        private List<ColumnLinkConfig> _columnLinks = new();
        private DataTable _dataTable;

        public ColumnManagerDialog(DataTable dataTable, bool[] checkedStates, int frozenRows, int lockedColumnIndex,
            string rootDir = "", string currentFilePath = "", List<(string display, string fullPath)>? fileEntries = null,
            List<ColumnLinkConfig>? columnLinks = null, int headerRowIndex = -1,
            int mappingRowIndex = -1)
        {
            _dataTable = dataTable;
            _lockedColumnIndex = lockedColumnIndex;
            _checkedStates = (bool[])checkedStates.Clone();
            _rootDir = rootDir;
            _currentFilePath = currentFilePath;
            _frozenRows = frozenRows;
            _headerRowIndex = headerRowIndex;
            _mappingRowIndex = mappingRowIndex;
            _fileEntries = fileEntries ?? new();
            _columnLinks = columnLinks != null ? new List<ColumnLinkConfig>(columnLinks) : new();
            InitializeComponent();
            PopulateList(dataTable, frozenRows);
        }

        public bool[] GetCheckedStates() => _checkedStates;
        public List<ColumnLinkConfig> GetColumnLinks() => _columnLinks;

        private string GetRelativePath(string fullPath)
        {
            if (string.IsNullOrEmpty(_rootDir) || string.IsNullOrEmpty(fullPath))
                return Path.GetFileName(fullPath);
            if (fullPath.StartsWith(_rootDir, StringComparison.OrdinalIgnoreCase))
                return fullPath.Substring(_rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(fullPath);
        }

        private void PopulateList(DataTable dataTable, int frozenRows)
        {
            _lvColumns.BeginUpdate();
            _lvColumns.Items.Clear();
            _lvColumns.Columns.Clear();

            int frozenCount = Math.Min(frozenRows, dataTable.Rows.Count);

            // 第一列：选择
            _lvColumns.Columns.Add("选择", 50, System.Windows.Forms.HorizontalAlignment.Center);

            // 字段名列：显示表头行对应的字段名
            if (_headerRowIndex >= 0 && _headerRowIndex < dataTable.Rows.Count && _headerRowIndex < frozenCount)
                _lvColumns.Columns.Add($"字段名(行{_headerRowIndex + 1})", 140, System.Windows.Forms.HorizontalAlignment.Left);

            // 映射名列：如果设置了映射行且与表头行不同，显示映射行的值
            if (_mappingRowIndex >= 0 && _mappingRowIndex < dataTable.Rows.Count && _mappingRowIndex < frozenCount
                && _mappingRowIndex != _headerRowIndex)
                _lvColumns.Columns.Add($"映射名(行{_mappingRowIndex + 1})", 140, System.Windows.Forms.HorizontalAlignment.Left);

            // 其余冻结行
            for (int r = 0; r < frozenCount; r++)
            {
                if (r == _headerRowIndex || r == _mappingRowIndex) continue;
                _lvColumns.Columns.Add($"行{r + 1}", 130, System.Windows.Forms.HorizontalAlignment.Left);
            }

            // 关联列
            _lvColumns.Columns.Add("关联", 100, System.Windows.Forms.HorizontalAlignment.Center);

            string currentRelPath = GetRelativePath(_currentFilePath);

            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                var item = new ListViewItem("")
                {
                    Tag = col,
                    Checked = _checkedStates[col]
                };

                // 字段名
                if (_headerRowIndex >= 0 && _headerRowIndex < dataTable.Rows.Count && _headerRowIndex < frozenCount)
                {
                    var headerVal = dataTable.Rows[_headerRowIndex][col]?.ToString() ?? "";
                    item.SubItems.Add(headerVal);
                }

                // 映射名
                if (_mappingRowIndex >= 0 && _mappingRowIndex < dataTable.Rows.Count && _mappingRowIndex < frozenCount
                    && _mappingRowIndex != _headerRowIndex)
                {
                    var mappingVal = dataTable.Rows[_mappingRowIndex][col]?.ToString() ?? "";
                    item.SubItems.Add(mappingVal);
                }

                // 其余冻结行的值
                for (int r = 0; r < frozenCount; r++)
                {
                    if (r == _headerRowIndex || r == _mappingRowIndex) continue;
                    var val = dataTable.Rows[r][col]?.ToString() ?? "";
                    item.SubItems.Add(val);
                }

                var link = _columnLinks.Find(l =>
                    l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                    l.SourceColumnIndex == col);
                item.SubItems.Add(link != null ? $"⚡{Path.GetFileNameWithoutExtension(link.TargetFile)}" : "");

                _lvColumns.Items.Add(item);
            }

            _lvColumns.EndUpdate();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "列管理";
            ClientSize = new Size(980, 520);
            MinimumSize = new Size(600, 300);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false;
            MaximizeBox = false;

            // 全选/取消按钮
            _btnSelectAll = new Button
            {
                Text = "全选",
                Location = new Point(10, 8),
                Size = new Size(70, 26),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) },
                BackColor = Color.FromArgb(235, 245, 255),
                ForeColor = Color.FromArgb(20, 60, 100)
            };
            _btnSelectAll.Click += (s, e) => SetAllChecked(true);

            _btnDeselectAll = new Button
            {
                Text = "取消全选",
                Location = new Point(88, 8),
                Size = new Size(80, 26),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) },
                BackColor = Color.FromArgb(235, 245, 255),
                ForeColor = Color.FromArgb(20, 60, 100)
            };
            _btnDeselectAll.Click += (s, e) => SetAllChecked(false);

            // ListView — 使用原生 CheckBoxes
            _lvColumns = new ListView
            {
                Location = new Point(10, 42),
                Size = new Size(960, 420),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(20, 60, 100),
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                GridLines = true,
                CheckBoxes = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            // 阻止第一列取消选中
            _lvColumns.ItemCheck += (s, e) =>
            {
                if (e.Index == _lockedColumnIndex && e.NewValue == CheckState.Unchecked)
                {
                    e.NewValue = CheckState.Checked;
                }
            };
            // 同步选中状态
            _lvColumns.ItemChecked += (s, e) =>
            {
                if (e.Item.Index < _checkedStates.Length)
                    _checkedStates[e.Item.Index] = e.Item.Checked;
            };

            // 关联管理按钮
            _btnLinkManage = new Button
            {
                Text = "关联设置",
                Location = new Point(600, 468),
                Size = new Size(80, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(22, 163, 74),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(21, 128, 61), MouseOverBackColor = Color.FromArgb(21, 128, 61) },
                Cursor = Cursors.Hand
            };
            _btnLinkManage.Click += OnLinkManageClick;

            // OK 按钮
            _btnOK = new Button
            {
                Text = "确定",
                Location = new Point(690, 468),
                Size = new Size(80, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
            };

            // 取消按钮
            _btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(780, 468),
                Size = new Size(70, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };

            Controls.AddRange(new Control[] { _btnSelectAll, _btnDeselectAll, _lvColumns, _btnLinkManage, _btnOK, _btnCancel });
            AcceptButton = _btnOK;
            CancelButton = _btnCancel;

            ResumeLayout(false);
        }

        private void OnLinkManageClick(object? sender, EventArgs e)
        {
            if (_lvColumns.SelectedItems.Count == 0)
            {
                MessageBox.Show(this, "请先在列表中选择要设置关联的列（点击行选中）", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var selectedItem = _lvColumns.SelectedItems[0];
            int colIndex = (int)selectedItem.Tag;
            string currentRelPath = GetRelativePath(_currentFilePath);

            using var setupDialog = new ColumnLinkSetupDialog(
                _fileEntries, currentRelPath, colIndex, _rootDir, _frozenRows,
                _columnLinks.FindAll(l =>
                    l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                    l.SourceColumnIndex == colIndex),
                _headerRowIndex);
            if (setupDialog.ShowDialog(this) == DialogResult.OK)
            {
                _columnLinks.RemoveAll(l =>
                    l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                    l.SourceColumnIndex == colIndex);
                var newLink = setupDialog.GetResult();
                if (newLink != null)
                    _columnLinks.Add(newLink);

                UpdateLinkColumnInListView();
            }
        }

        private void UpdateLinkColumnInListView()
        {
            string currentRelPath = GetRelativePath(_currentFilePath);
            foreach (ListViewItem item in _lvColumns.Items)
            {
                if (item.Tag is int col)
                {
                    var link = _columnLinks.Find(l =>
                        l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                        l.SourceColumnIndex == col);
                    var linkText = item.SubItems[item.SubItems.Count - 1];

                    string displayText = "";
                    if (link != null)
                    {
                        if (!string.IsNullOrEmpty(link.TargetColumnName))
                            displayText = $"{Path.GetFileNameWithoutExtension(link.TargetFile)} [{link.TargetColumnName}]";
                        else
                            displayText = $"{Path.GetFileNameWithoutExtension(link.TargetFile)}";
                    }
                    linkText.Text = displayText;
                }
            }
        }

        private void SetAllChecked(bool check)
        {
            for (int i = 0; i < _lvColumns.Items.Count; i++)
            {
                if (i == _lockedColumnIndex)
                {
                    _lvColumns.Items[i].Checked = true;
                    _checkedStates[i] = true;
                }
                else
                {
                    _lvColumns.Items[i].Checked = check;
                    _checkedStates[i] = check;
                }
            }
        }
    }

    /// <summary>列关联设置对话框 — 双击文件后弹出带固定行数据的单选框面板</summary>
    public class ColumnLinkSetupDialog : Form
    {
        private ListView _lstFiles = new();
        private TextBox _txtFilter = new();
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private Button _btnDelete = new();
        private Label _lblSelection = new();
        private List<(string display, string fullPath)> _fileEntries = new();
        private string _sourceFile = "";
        private int _sourceColumnIndex;
        private string _rootDir = "";
        private int _frozenRows = 0;
        private int _targetHeaderRowIndex = -1;       // 目标文件的表头行索引
        private ColumnLinkConfig? _existingLink;
        private string _selectedTargetPath = "";
        private int _selectedTargetColIndex = -1;
        private string _selectedTargetColumnName = "";  // 目标列的表头名称（用于显示）
        private int _selectedDisplayColIndex = -1;      // 显示列的索引（-1 表示不使用）
        private string _selectedDisplayColumnName = ""; // 显示列的表头名称
        private Button _btnSetDisplayCol = new();       // 设置显示列的按钮
        public int SelectedColumnIndex => _selectedTargetColIndex;
        public string SelectedColumnName => _selectedTargetColumnName ?? "";
        // 值提取规则控件
        private CheckBox _chkExtractEnabled = new();
        private Label _lblDelimiter = new();
        private TextBox _txtDelimiter = new();
        private Label _lblPartIndex = new();
        private NumericUpDown _nudPartIndex = new();
        // 多值分隔符控件
        private Label _lblMultiValueDelimiter = new();
        private TextBox _txtMultiValueDelimiter = new();

        public ColumnLinkSetupDialog(
            List<(string display, string fullPath)> fileEntries,
            string sourceFile, int sourceColumnIndex,
            string rootDir, int frozenRows,
            List<ColumnLinkConfig> existingLinks,
            int targetHeaderRowIndex = -1)
        {
            _fileEntries = fileEntries;
            _sourceFile = sourceFile;
            _sourceColumnIndex = sourceColumnIndex;
            _rootDir = rootDir;
            _frozenRows = frozenRows;
            _targetHeaderRowIndex = targetHeaderRowIndex;
            _existingLink = existingLinks.FirstOrDefault();

            if (_existingLink != null)
            {
                // 尝试解析为全路径
                string resolved = _existingLink.TargetFile;
                if (!string.IsNullOrEmpty(rootDir))
                {
                    string candidate = Path.Combine(rootDir, _existingLink.TargetFile);
                    if (File.Exists(candidate))
                    {
                        resolved = candidate;
                    }
                    else
                    {
                        // 尝试在 fileEntries 中按文件名匹配
                        var match = _fileEntries.FirstOrDefault(f =>
                            Path.GetFileName(f.fullPath).Equals(Path.GetFileName(_existingLink.TargetFile), StringComparison.OrdinalIgnoreCase));
                        if (match.fullPath != null)
                            resolved = match.fullPath;
                    }
                }
                _selectedTargetPath = resolved;
                _selectedTargetColIndex = _existingLink.TargetColumnIndex;

                // 恢复目标列的表头名称（如果有保存）
                if (!string.IsNullOrEmpty(_existingLink.TargetColumnName))
                    _selectedTargetColumnName = _existingLink.TargetColumnName;
                // 恢复显示列设置
                _selectedDisplayColIndex = _existingLink.TargetDisplayColumnIndex;
                _selectedDisplayColumnName = _existingLink.TargetDisplayColumnName;
            }

            InitializeComponent();
            PopulateFiles("");
            UpdateSelectionLabel();

            // 加载已有配置（必须在 InitializeComponent 之后，否则控件会被重新创建覆盖）
            if (_existingLink != null)
            {
                _chkExtractEnabled.Checked = _existingLink.ValueExtractionEnabled;
                _txtDelimiter.Text = _existingLink.ExtractDelimiter ?? "";
                _nudPartIndex.Value = Math.Max(0, _existingLink.ExtractPartIndex);
                _txtMultiValueDelimiter.Text = _existingLink.MultiValueDelimiter ?? "";
            }

            if (_existingLink != null)
            {
                string targetRel = _existingLink.TargetFile.Replace("\\", "/");
                foreach (ListViewItem item in _lstFiles.Items)
                {
                    if (item.Tag is string fp)
                    {
                        string fpNorm = fp.Replace("\\", "/");
                        if (fpNorm.EndsWith(targetRel, StringComparison.OrdinalIgnoreCase) ||
                            Path.GetFileName(fp).Equals(Path.GetFileName(_existingLink.TargetFile), StringComparison.OrdinalIgnoreCase))
                        {
                            item.Selected = true;
                            _lstFiles.EnsureVisible(item.Index);
                            break;
                        }
                    }
                }
            }
        }

        public ColumnLinkConfig? GetResult()
        {
            if (string.IsNullOrEmpty(_selectedTargetPath) || _selectedTargetColIndex < 0)
                return null;
            string relPath;
            if (!string.IsNullOrEmpty(_rootDir) && _selectedTargetPath.StartsWith(_rootDir, StringComparison.OrdinalIgnoreCase))
                relPath = _selectedTargetPath.Substring(_rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            else
                relPath = Path.GetFileName(_selectedTargetPath);
            return new ColumnLinkConfig
            {
                SourceFile = _sourceFile,
                SourceColumnIndex = _sourceColumnIndex,
                TargetFile = relPath,
                TargetColumnIndex = _selectedTargetColIndex,
                TargetColumnName = _selectedTargetColumnName,  // 保存表头名称
                TargetDisplayColumnIndex = _selectedDisplayColIndex,
                TargetDisplayColumnName = _selectedDisplayColumnName,
                // 值提取规则
                ValueExtractionEnabled = _chkExtractEnabled.Checked,
                ExtractDelimiter = _txtDelimiter.Text.Trim(),
                ExtractPartIndex = (int)_nudPartIndex.Value,
                MultiValueDelimiter = _txtMultiValueDelimiter.Text.Trim()
            };
        }

        private void UpdateSelectionLabel()
        {
            if (!string.IsNullOrEmpty(_selectedTargetPath) && _selectedTargetColIndex >= 0)
            {
                string displayPath = _selectedTargetPath;
                if (!string.IsNullOrEmpty(_rootDir) && _selectedTargetPath.StartsWith(_rootDir, StringComparison.OrdinalIgnoreCase))
                    displayPath = _selectedTargetPath.Substring(_rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
                else
                    displayPath = Path.GetFileName(_selectedTargetPath);
                string colInfo = !string.IsNullOrEmpty(_selectedTargetColumnName)
                    ? $"[{_selectedTargetColumnName}]"
                    : $"列{_selectedTargetColIndex}";
                string displayInfo = _selectedDisplayColIndex >= 0
                    ? (!string.IsNullOrEmpty(_selectedDisplayColumnName) ? $" (显示: {_selectedDisplayColumnName})" : $" (显示: 列{_selectedDisplayColIndex})")
                    : "";
                _lblSelection.Text = $"已选: {displayPath} → {colInfo}{displayInfo}";
            }
            else
                _lblSelection.Text = "未选择关联列（双击目标文件设置）";
        }

        private void PopulateFiles(string filter)
        {
            _lstFiles.BeginUpdate();
            _lstFiles.Items.Clear();
            var query = string.IsNullOrEmpty(filter)
                ? _fileEntries
                : _fileEntries.Where(f => f.display.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
            foreach (var entry in query.OrderBy(f => f.display))
            {
                _lstFiles.Items.Add(new ListViewItem(entry.display) { Tag = entry.fullPath, ToolTipText = entry.fullPath });
            }
            _lstFiles.EndUpdate();
        }

        /// <summary>双击文件，弹出列选择面板（带固定行数据 + 单选框）</summary>
        private void OnFileDoubleClick(object? sender, EventArgs e)
        {
            if (_lstFiles.SelectedItems.Count == 0) return;
            var fullPath = _lstFiles.SelectedItems[0].Tag as string ?? "";
            if (!File.Exists(fullPath))
            {
                MessageBox.Show(this, "文件不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 读取目标文件固定行数据用于展示
            DataTable targetFrozenTable;
            try
            {
                var ext = Path.GetExtension(fullPath).ToLowerInvariant();
                if (ext == ".csv")
                {
                    string contentData = File.ReadAllText(fullPath, Encoding.UTF8);
                    var lines = SplitCsvLinesStatic(contentData);
                    var dt = new DataTable();
                    // 以第一行为参照确定列数
                    if (lines.Count > 0)
                    {
                        var firstParts = ParseCsvLineStatic(lines[0]);
                        for (int i = 0; i < firstParts.Length; i++)
                            dt.Columns.Add($"Col{i}", typeof(string));
                        int maxRows = Math.Min(_frozenRows > 0 ? _frozenRows : 2, lines.Count);
                        for (int r = 0; r < maxRows; r++)
                        {
                            var parts = ParseCsvLineStatic(lines[r]);
                            var row = dt.NewRow();
                            for (int i = 0; i < Math.Min(parts.Length, dt.Columns.Count); i++)
                                row[i] = parts[i];
                            dt.Rows.Add(row);
                        }
                    }
                    targetFrozenTable = dt;
                }
                else
                {
                    try
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using var fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        using var excelReader = ExcelReaderFactory.CreateReader(fs);
                        var dataSet = excelReader.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
                        });
                        if (dataSet.Tables.Count > 0)
                        {
                            var sheet = dataSet.Tables[0];
                            int maxCols = sheet.Columns.Count;
                            var dt = new DataTable();
                            for (int i = 0; i < maxCols; i++)
                                dt.Columns.Add($"Col{i}", typeof(string));
                            int maxRows = Math.Min(_frozenRows > 0 ? _frozenRows : 2, sheet.Rows.Count);
                            for (int r = 0; r < maxRows; r++)
                            {
                                var dtRow = dt.NewRow();
                                for (int i = 0; i < maxCols; i++)
                                {
                                    dtRow[i] = sheet.Rows[r][i]?.ToString() ?? "";
                                }
                                dt.Rows.Add(dtRow);
                            }
                            targetFrozenTable = dt;
                        }
                        else targetFrozenTable = new DataTable();
                    }
                    catch
                    {
                        targetFrozenTable = new DataTable();
                    }
                }
            }
            catch
            {
                MessageBox.Show(this, "无法读取目标文件数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int initialSelection = -1;
            if (_existingLink != null)
            {
                string targetRel = _existingLink.TargetFile.Replace("\\", "/");
                string fullNorm = fullPath.Replace("\\", "/");
                if (fullNorm.EndsWith(targetRel, StringComparison.OrdinalIgnoreCase) ||
                    Path.GetFileName(fullPath).Equals(Path.GetFileName(_existingLink.TargetFile), StringComparison.OrdinalIgnoreCase))
                {
                    initialSelection = _existingLink.TargetColumnIndex;
                }
            }

            using var colDlg = new TargetColumnSelectDialog(targetFrozenTable, _frozenRows, initialSelection,
                _targetHeaderRowIndex >= 0 ? _targetHeaderRowIndex : GetTargetHeaderRowIndexStatic(fullPath));
            if (colDlg.ShowDialog(this) == DialogResult.OK && colDlg.SelectedColumnIndex >= 0)
            {
                _selectedTargetPath = fullPath;
                _selectedTargetColIndex = colDlg.SelectedColumnIndex;
                _selectedTargetColumnName = colDlg.SelectedColumnName;  // 保存表头名称
                UpdateSelectionLabel();
            }
        }

        private void OnDeleteClick(object? sender, EventArgs e)
        {
            _selectedTargetPath = "";
            _selectedTargetColIndex = -1;
            _selectedDisplayColIndex = -1;
            _selectedDisplayColumnName = "";
            _existingLink = null;
            _btnDelete.Visible = false;
            UpdateSelectionLabel();
        }

        /// <summary>设置显示列 — 在目标文件中再选一列作为显示名</summary>
        private void OnSetDisplayColClick(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedTargetPath) || _selectedTargetColIndex < 0)
            {
                MessageBox.Show(this, "请先选择关联列（双击目标文件）", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!File.Exists(_selectedTargetPath))
            {
                MessageBox.Show(this, "目标文件不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 读取目标文件固定行数据用于展示
            DataTable targetFrozenTable;
            try
            {
                var ext = Path.GetExtension(_selectedTargetPath).ToLowerInvariant();
                if (ext == ".csv")
                {
                    string contentData = File.ReadAllText(_selectedTargetPath, Encoding.UTF8);
                    var lines = SplitCsvLinesStatic(contentData);
                    var dt = new DataTable();
                    if (lines.Count > 0)
                    {
                        var firstParts = ParseCsvLineStatic(lines[0]);
                        for (int i = 0; i < firstParts.Length; i++)
                            dt.Columns.Add($"Col{i}", typeof(string));
                        int maxRows = Math.Min(_frozenRows > 0 ? _frozenRows : 2, lines.Count);
                        for (int r = 0; r < maxRows; r++)
                        {
                            var parts = ParseCsvLineStatic(lines[r]);
                            var row = dt.NewRow();
                            for (int i = 0; i < Math.Min(parts.Length, dt.Columns.Count); i++)
                                row[i] = parts[i];
                            dt.Rows.Add(row);
                        }
                    }
                    targetFrozenTable = dt;
                }
                else
                {
                    try
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using var fs = new FileStream(_selectedTargetPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        using var excelReader = ExcelReaderFactory.CreateReader(fs);
                        var dataSet = excelReader.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
                        });
                        if (dataSet.Tables.Count > 0)
                        {
                            var sheet = dataSet.Tables[0];
                            int maxCols = sheet.Columns.Count;
                            var dt = new DataTable();
                            for (int i = 0; i < maxCols; i++)
                                dt.Columns.Add($"Col{i}", typeof(string));
                            int maxRows = Math.Min(_frozenRows > 0 ? _frozenRows : 2, sheet.Rows.Count);
                            for (int r = 0; r < maxRows; r++)
                            {
                                var dtRow = dt.NewRow();
                                for (int i = 0; i < maxCols; i++)
                                    dtRow[i] = sheet.Rows[r][i]?.ToString() ?? "";
                                dt.Rows.Add(dtRow);
                            }
                            targetFrozenTable = dt;
                        }
                        else targetFrozenTable = new DataTable();
                    }
                    catch { targetFrozenTable = new DataTable(); }
                }
            }
            catch
            {
                MessageBox.Show(this, "无法读取目标文件数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var colDlg = new TargetColumnSelectDialog(targetFrozenTable, _frozenRows, _selectedDisplayColIndex,
                _targetHeaderRowIndex >= 0 ? _targetHeaderRowIndex : GetTargetHeaderRowIndexStatic(_selectedTargetPath),
                title: "选择显示列 — 用于提示中显示中文名称");
            if (colDlg.ShowDialog(this) == DialogResult.OK && colDlg.SelectedColumnIndex >= 0)
            {
                _selectedDisplayColIndex = colDlg.SelectedColumnIndex;
                _selectedDisplayColumnName = colDlg.SelectedColumnName;
                _btnSetDisplayCol.Text = !string.IsNullOrEmpty(_selectedDisplayColumnName)
                    ? $"显示列: {_selectedDisplayColumnName}"
                    : $"显示列: 列{_selectedDisplayColIndex}";
                UpdateSelectionLabel();
            }
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            Text = "列关联设置 - 列 " + _sourceColumnIndex;
            ClientSize = new Size(500, 560);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false; MaximizeBox = false;

            Controls.Add(new Label { Text = "选择关联的目标文件（双击设置列）:", Location = new Point(10, 10), Size = new Size(480, 20),
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(31, 41, 55) });

            _txtFilter = new TextBox { Location = new Point(10, 32), Size = new Size(480, 23),
                Font = new Font("Microsoft YaHei UI", 9f), BorderStyle = BorderStyle.FixedSingle, PlaceholderText = "输入关键字筛选文件..." };
            _txtFilter.TextChanged += (s, e) => PopulateFiles(_txtFilter.Text.Trim());
            Controls.Add(_txtFilter);

            _lstFiles = new ListView { Location = new Point(10, 62), Size = new Size(480, 200), BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f), View = View.Details, FullRowSelect = true, HideSelection = false,
                HeaderStyle = ColumnHeaderStyle.None, MultiSelect = false };
            _lstFiles.Columns.Add("文件", 460);
            _lstFiles.DoubleClick += OnFileDoubleClick;
            Controls.Add(_lstFiles);

            _lblSelection = new Label { Text = "未选择关联列（双击目标文件设置）", Location = new Point(10, 272),
                Size = new Size(480, 40), Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = Color.FromArgb(31, 41, 55), BackColor = Color.FromArgb(235, 248, 255),
                BorderStyle = BorderStyle.FixedSingle, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(4, 0, 0, 0) };
            Controls.Add(_lblSelection);

            // 设置显示列按钮
            _btnSetDisplayCol = new Button
            {
                Text = _selectedDisplayColIndex >= 0
                    ? (!string.IsNullOrEmpty(_selectedDisplayColumnName)
                        ? $"显示列: {_selectedDisplayColumnName}"
                        : $"显示列: 列{_selectedDisplayColIndex}")
                    : "设置显示列（可选）",
                Location = new Point(10, 320), Size = new Size(200, 28),
                Font = new Font("Microsoft YaHei UI", 8f),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) },
                Cursor = Cursors.Hand
            };
            _btnSetDisplayCol.Click += OnSetDisplayColClick;
            Controls.Add(_btnSetDisplayCol);

            Controls.Add(new Label { Text = "提示: 关联基于列的位置索引。双击目标文件后在弹出面板中单选目标列。",
                Location = new Point(220, 322), Size = new Size(270, 30), Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(107, 114, 128) });

            // ===== 值提取规则区域 =====
            var grpExtract = new GroupBox
            {
                Text = "值提取规则（可选）— 用于从复合格式中提取部分内容匹配",
                Location = new Point(10, 355), Size = new Size(480, 105),
                Font = new Font("Microsoft YaHei UI", 8f), ForeColor = Color.FromArgb(55, 65, 81)
            };

            _chkExtractEnabled = new CheckBox
            {
                Text = "启用提取（如 \"4002#1\" → 提取 \"4002\" 匹配）",
                Location = new Point(10, 20), Size = new Size(300, 22),
                Font = new Font("Microsoft YaHei UI", 8f), Checked = false,
                ForeColor = Color.FromArgb(31, 41, 55)
            };
            _chkExtractEnabled.CheckedChanged += (s, e) =>
            {
                _lblDelimiter.Enabled = _chkExtractEnabled.Checked;
                _txtDelimiter.Enabled = _chkExtractEnabled.Checked;
                _lblPartIndex.Enabled = _chkExtractEnabled.Checked;
                _nudPartIndex.Enabled = _chkExtractEnabled.Checked;
                _lblMultiValueDelimiter.Enabled = _chkExtractEnabled.Checked;
                _txtMultiValueDelimiter.Enabled = _chkExtractEnabled.Checked;
            };
            grpExtract.Controls.Add(_chkExtractEnabled);

            _lblDelimiter = new Label { Text = "分隔符:", Location = new Point(10, 46), Size = new Size(50, 18),
                Font = new Font("Microsoft YaHei UI", 8f), Enabled = false, ForeColor = Color.FromArgb(55, 65, 81) };
            grpExtract.Controls.Add(_lblDelimiter);

            _txtDelimiter = new TextBox { Text = "#", Location = new Point(62, 44), Size = new Size(40, 21),
                Font = new Font("Microsoft YaHei UI", 8f), BorderStyle = BorderStyle.FixedSingle,
                Enabled = false, PlaceholderText = "# - | ..." };
            grpExtract.Controls.Add(_txtDelimiter);

            _lblPartIndex = new Label { Text = "取第几段(0起):", Location = new Point(110, 46), Size = new Size(85, 18),
                Font = new Font("Microsoft YaHei UI", 8f), Enabled = false, ForeColor = Color.FromArgb(55, 65, 81) };
            grpExtract.Controls.Add(_lblPartIndex);

            _nudPartIndex = new NumericUpDown { Value = 0, Minimum = 0, Maximum = 99, Location = new Point(198, 44),
                Size = new Size(45, 21), Font = new Font("Microsoft YaHei UI", 8f), Enabled = false,
                TextAlign = System.Windows.Forms.HorizontalAlignment.Right };
            grpExtract.Controls.Add(_nudPartIndex);

            var lblExample = new Label { Text = "例: \"4002#1\" 用 # 分隔 → 取段0得 \"4002\"",
                Location = new Point(250, 45), Size = new Size(220, 20),
                Font = new Font("Microsoft YaHei UI", 7.5f), ForeColor = Color.FromArgb(140, 145, 150) };
            grpExtract.Controls.Add(lblExample);

            // 多值分隔符
            _lblMultiValueDelimiter = new Label { Text = "多值分隔符:", Location = new Point(10, 72), Size = new Size(72, 18),
                Font = new Font("Microsoft YaHei UI", 8f), Enabled = false, ForeColor = Color.FromArgb(55, 65, 81) };
            grpExtract.Controls.Add(_lblMultiValueDelimiter);

            _txtMultiValueDelimiter = new TextBox { Text = "|", Location = new Point(85, 70), Size = new Size(40, 21),
                Font = new Font("Microsoft YaHei UI", 8f), BorderStyle = BorderStyle.FixedSingle,
                Enabled = false };
            grpExtract.Controls.Add(_txtMultiValueDelimiter);

            var lblMultiExample = new Label { Text = "例: \"100500201#1|100500202#1\" 先用 | 切分多组 → 各组再用分隔符提取",
                Location = new Point(130, 71), Size = new Size(340, 20),
                Font = new Font("Microsoft YaHei UI", 7.5f), ForeColor = Color.FromArgb(140, 145, 150) };
            grpExtract.Controls.Add(lblMultiExample);
            // ===== 结束提取规则区域 =====

            Controls.Add(grpExtract);

            _btnDelete = new Button { Text = "删除关联", Location = new Point(10, 470), Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(239, 68, 68), ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(220, 50, 50) }, Cursor = Cursors.Hand,
                Visible = _existingLink != null || !string.IsNullOrEmpty(_selectedTargetPath) };
            _btnDelete.Click += OnDeleteClick;
            Controls.Add(_btnDelete);

            _btnOK = new Button { Text = "确定", Location = new Point(330, 470), Size = new Size(80, 28), FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(66, 165, 245), ForeColor = Color.White, FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                DialogResult = DialogResult.OK, Cursor = Cursors.Hand };
            Controls.Add(_btnOK);

            _btnCancel = new Button { Text = "取消", Location = new Point(420, 470), Size = new Size(70, 28), FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel, BackColor = Color.FromArgb(235, 248, 255), ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) } };
            Controls.Add(_btnCancel);

            AcceptButton = _btnOK; CancelButton = _btnCancel;
            ResumeLayout(false);
        }

        private static List<string> SplitCsvLinesStatic(string data)
        {
            var lines = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < data.Length; i++)
            {
                char c = data[i];
                if (c == '\u0022') inQuotes = !inQuotes;
                else if (!inQuotes && c == '\u000A')
                {
                    lines.Add(sb.ToString());
                    sb.Clear();
                }
                else if (!inQuotes && c == '\u000D')
                {
                    if (i + 1 < data.Length && data[i + 1] == '\u000A') i++;
                    lines.Add(sb.ToString());
                    sb.Clear();
                }
                else sb.Append(c);
            }
            if (sb.Length > 0) lines.Add(sb.ToString());
            return lines;
        }

        private static string[] ParseCsvLineStatic(string line)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\u0022')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '\u0022') { sb.Append('\u0022'); i++; }
                    else inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes) { result.Add(sb.ToString().Trim()); sb.Clear(); }
                else sb.Append(c);
            }
            result.Add(sb.ToString());
            return result.ToArray();
        }

        /// <summary>根据目标文件路径获取其表头行索引（从设置中读取）</summary>
        private static int GetTargetHeaderRowIndexStatic(string targetFullPath)
        {
            if (string.IsNullOrEmpty(targetFullPath)) return -1;
            string targetDir = Path.GetDirectoryName(targetFullPath) ?? "";
            // 如果在同一目录下，使用该目录的设置
            var ext = Path.GetExtension(targetFullPath).ToLowerInvariant();
            if (ext == ".csv")
                return CsvSettings.GetCsvHeaderRowIndexByDirectory(targetDir);
            else if (ext == ".xls" || ext == ".xlsx")
                return CsvSettings.GetExcelHeaderRowIndexByDirectory(targetDir);
            return -1;
        }
    }

    /// <summary>目标列选择对话框 — 带固定行数据的单选面板（类似列选择面板，但使用单选框，支持表头行显示）</summary>
    public class TargetColumnSelectDialog : Form
    {
        private ListView _lvColumns = new();
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private int _selectedIndex = -1;
        private string _selectedColumnName = "";        // 选中的列的表头名称
        private int _headerRowIndex = -1;              // 表头行索引
        private bool _hasHeaderRow = false;
        private string _customTitle = "选择目标列";

        public int SelectedColumnIndex => _selectedIndex;
        public string SelectedColumnName => _selectedColumnName ?? "";

        public TargetColumnSelectDialog(DataTable frozenData, int frozenRows, int initialSelection, int headerRowIndex = -1, string title = "选择目标列")
        {
            _selectedIndex = initialSelection;
            _headerRowIndex = headerRowIndex;
            _customTitle = title;
            InitializeComponent();
            PopulateList(frozenData, frozenRows);
            if (_selectedIndex >= 0 && _selectedIndex < _lvColumns.Items.Count)
            {
                _lvColumns.Items[_selectedIndex].Selected = true;
                _lvColumns.EnsureVisible(_selectedIndex);
            }
        }

        private void PopulateList(DataTable frozenData, int frozenRows)
        {
            _lvColumns.BeginUpdate();
            _lvColumns.Items.Clear();
            _lvColumns.Columns.Clear();

            int frozenCount = Math.Min(frozenRows > 0 ? frozenRows : 2, frozenData.Rows.Count);

            // 表头行标志：如果有设置表头行，用于显示和获取列名
            _hasHeaderRow = (_headerRowIndex >= 0 && _headerRowIndex < frozenData.Rows.Count && _headerRowIndex < frozenCount);

            // 第一列：选择
            _lvColumns.Columns.Add("选择", 50, System.Windows.Forms.HorizontalAlignment.Center);

            // 表头名列：如果有设置表头行，优先显示字段名
            if (_hasHeaderRow)
                _lvColumns.Columns.Add($"字段名(行{_headerRowIndex + 1})", 160, System.Windows.Forms.HorizontalAlignment.Left);

            // 其余冻结行
            for (int r = 0; r < frozenCount; r++)
            {
                if (r == _headerRowIndex) continue; // 已作为字段名列显示
                _lvColumns.Columns.Add($"行{r + 1}", 140, System.Windows.Forms.HorizontalAlignment.Left);
            }

            for (int col = 0; col < frozenData.Columns.Count; col++)
            {
                var item = new ListViewItem("")
                {
                    Tag = col
                };

                // 先添加字段名（如果设置了表头行）
                if (_hasHeaderRow)
                {
                    var headerVal = frozenData.Rows[_headerRowIndex][col]?.ToString() ?? "";
                    item.SubItems.Add(headerVal);
                }

                // 添加其余冻结行的值
                for (int r = 0; r < frozenCount; r++)
                {
                    if (r == _headerRowIndex) continue;
                    var val = frozenData.Rows[r][col]?.ToString() ?? "";
                    item.SubItems.Add(val);
                }

                _lvColumns.Items.Add(item);
            }

            // 高亮标记表头行所在的列头（ColumnHeader 不支持 BackColor，跳过）
            
            _lvColumns.EndUpdate();
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            SuspendLayout();
            Text = _customTitle;
            ClientSize = new Size(780, 520);
            MinimumSize = new Size(560, 400);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false; MaximizeBox = false;
            AutoScaleMode = AutoScaleMode.Font;

            Controls.Add(new Label { Text = "请点击行选择目标关联列（单选）:", Location = new Point(10, 10), Size = new Size(760, 20),
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(31, 41, 55) });

            _lvColumns = new ListView
            {
                Location = new Point(10, 38),
                Size = new Size(760, 430),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(20, 60, 100),
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                GridLines = true,
                CheckBoxes = false,
                MultiSelect = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            _lvColumns.SelectedIndexChanged += (s, e) =>
            {
                if (_lvColumns.SelectedItems.Count > 0)
                {
                    var selectedItem = _lvColumns.SelectedItems[0];
                    _selectedIndex = (int)selectedItem.Tag;
                    // 获取选中的列的显示名称：有表头行用表头名称，否则用第一行数据作为列名
                    if (selectedItem.SubItems.Count > 1)
                        _selectedColumnName = selectedItem.SubItems[1]?.Text ?? "";
                    else
                        _selectedColumnName = "";
                }
            };

            _btnOK = new Button { Text = "确定", Location = new Point(590, 478),
                Size = new Size(80, 28), Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White, FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                DialogResult = DialogResult.OK, Cursor = Cursors.Hand };
            Controls.Add(_btnOK);

            _btnCancel = new Button { Text = "取消", Location = new Point(680, 478),
                Size = new Size(70, 28), Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat, DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(235, 248, 255), ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) } };
            Controls.Add(_btnCancel);

            Controls.Add(_lvColumns);
            AcceptButton = _btnOK; CancelButton = _btnCancel;
            ResumeLayout(false);
        }
    }

    internal static class ColumnLinkChecker_SimpleCsv
    {
        public static string[] ParseLine(string line)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\u0022')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '\u0022') { sb.Append('\u0022'); i++; }
                    else inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes) { result.Add(sb.ToString().Trim()); sb.Clear(); }
                else sb.Append(c);
            }
            result.Add(sb.ToString());
            return result.ToArray();
        }

        /// <summary>根据目标文件路径获取其表头行索引（从设置中读取）</summary>
        private static int GetTargetHeaderRowIndexStatic(string targetFullPath)
        {
            if (string.IsNullOrEmpty(targetFullPath)) return -1;
            string targetDir = Path.GetDirectoryName(targetFullPath) ?? "";
            // 如果在同一目录下，使用该目录的设置
            var ext = Path.GetExtension(targetFullPath).ToLowerInvariant();
            if (ext == ".csv")
                return CsvSettings.GetCsvHeaderRowIndexByDirectory(targetDir);
            else if (ext == ".xls" || ext == ".xlsx")
                return CsvSettings.GetExcelHeaderRowIndexByDirectory(targetDir);
            return -1;
        }
    }

    /// <summary>关联表选择对话框 — 展示左侧已加载的文件列表，带搜索筛选</summary>
    public class LinkTableSelectorDialog : Form
    {
        private ListView _lstFiles = new();
        private TextBox _txtFilter = new();
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private List<(string display, string fullPath)> _fileEntries = new();
        private List<(string display, string fullPath)> _filteredEntries = new();

        public string SelectedFilePath { get; private set; } = "";

        public LinkTableSelectorDialog(List<(string display, string fullPath)> fileEntries)
        {
            _fileEntries = fileEntries;
            InitializeComponent();
            PopulateList("");
        }

        private void PopulateList(string filter)
        {
            _lstFiles.BeginUpdate();
            _lstFiles.Items.Clear();
            _filteredEntries.Clear();

            var query = string.IsNullOrEmpty(filter)
                ? _fileEntries
                : _fileEntries.Where(f => f.display.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var entry in query.OrderBy(f => f.display))
            {
                _filteredEntries.Add(entry);
                var item = new ListViewItem(entry.display)
                {
                    Tag = entry.fullPath,
                    ToolTipText = entry.fullPath
                };
                _lstFiles.Items.Add(item);
            }

            _lstFiles.EndUpdate();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "选择关联表";
            ClientSize = new Size(500, 400);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false;
            MaximizeBox = false;

            // 筛选框
            _txtFilter = new TextBox
            {
                Location = new Point(10, 10),
                Size = new Size(480, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                PlaceholderText = "输入关键字筛选文件..."
            };
            _txtFilter.TextChanged += (s, e) => PopulateList(_txtFilter.Text.Trim());

            // 文件列表
            _lstFiles = new ListView
            {
                Location = new Point(10, 40),
                Size = new Size(480, 310),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f),
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                GridLines = false,
                HeaderStyle = ColumnHeaderStyle.None,
                MultiSelect = false
            };
            _lstFiles.Columns.Add("文件", 460);
            _lstFiles.DoubleClick += (s, e) =>
            {
                if (_lstFiles.SelectedItems.Count > 0)
                {
                    SelectedFilePath = _lstFiles.SelectedItems[0].Tag as string ?? "";
                    DialogResult = DialogResult.OK;
                    Close();
                }
            };

            // OK 按钮
            _btnOK = new Button
            {
                Text = "确定",
                Location = new Point(320, 360),
                Size = new Size(80, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
            };
            _btnOK.Click += (s, e) =>
            {
                if (_lstFiles.SelectedItems.Count > 0)
                {
                    SelectedFilePath = _lstFiles.SelectedItems[0].Tag as string ?? "";
                    DialogResult = DialogResult.OK;
                    Close();
                }
            };

            // 取消按钮
            _btnCancel = new Button
            {
                Text = "清除",
                Location = new Point(408, 360),
                Size = new Size(70, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };

            Controls.AddRange(new Control[] { _txtFilter, _lstFiles, _btnOK, _btnCancel });
            AcceptButton = _btnOK;
            CancelButton = _btnCancel;

            ResumeLayout(false);
        }
    }
}
