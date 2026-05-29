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
        private Label _lblFilter = new();
        private Button _btnRefresh = new();
        private Button _btnSelectFolder = new();
        private Label _lblFileCount = new();
        private ImageList _imgList = new();

        // ==================== 状态字段 ====================
        private string _currentDir = "";
        private List<FileInfo> _allFiles = new();
        private CancellationTokenSource? _loadingCts;
        private readonly object _lockObj = new();
        private System.Windows.Forms.Timer? _filterDebounceTimer;
        private int _paneIndex = -1;                    // 所属窗格索引（-1表示独立打开，使用全局设置）
        private Dictionary<string, string> _fileRemarks = new(); // 文件备注映射（key: 相对路径）

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
            _pnlLeftTop.Height = 80;
            _pnlLeftTop.Dock = DockStyle.Top;
            _pnlLeftTop.BackColor = _colorSurface;
            _pnlLeftTop.Padding = new Padding(8, 6, 8, 6);

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

            // 文件数量标签（放在第一行右侧）
            _lblFileCount = new Label
            {
                Text = "共 0 个文件",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Location = new Point(124, 10)
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

            _pnlLeftTop.Controls.AddRange(new Control[] { _btnSelectFolder, _btnRefresh, _txtFileFilter, _lblFileCount });

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

            // 添加TreeView到左侧面板
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
                _txtFileFilter.Width = Math.Max(100, _pnlLeftTop.Width - 16);
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

            // 4. 设置备注名（仅文件节点可用）
            var mnuRemark = new ToolStripMenuItem("设置备注名");
            var isFile = System.IO.File.Exists(fullPath);
            mnuRemark.Enabled = isFile;
            mnuRemark.Click += (s, args) =>
            {
                if (!isFile) return;
                var relativePath = GetRelativePath(_currentDir, fullPath);
                var currentRemark = _fileRemarks.TryGetValue(relativePath, out var r) ? r : "";
                using var inputForm = new Form
                {
                    Text = "设置备注名",
                    ClientSize = new Size(380, 150),
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
                    Location = new Point(12, 54), Size = new Size(260, 23),
                    BorderStyle = BorderStyle.FixedSingle
                };
                var btnOk = new Button
                {
                    Text = "确定", DialogResult = DialogResult.OK,
                    Location = new Point(282, 53), Size = new Size(80, 25),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(66, 165, 245), ForeColor = Color.White
                };
                var btnCancel = new Button
                {
                    Text = "取消", DialogResult = DialogResult.Cancel,
                    Location = new Point(282, 84), Size = new Size(80, 25),
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

            ctx.Items.AddRange(new ToolStripItem[] { mnuOpen, mnuEdit, mnuOpenFolder, mnuRemark });
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

            Control viewer;
            if (fileType == FileType.Excel)
                viewer = new ExcelViewerControl(filePath, _paneIndex, _currentDir) { Dock = DockStyle.Fill };
            else
                viewer = new CsvViewerControl(filePath, _paneIndex, _currentDir) { Dock = DockStyle.Fill };

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
            if (viewer is CsvViewerControl csvViewer)
                csvViewer.LinkTableRequested += OnViewerLinkTableRequested;
            else if (viewer is ExcelViewerControl excelViewer)
                excelViewer.LinkTableRequested += OnViewerLinkTableRequested;
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
                var standaloneForm = new CsvReaderStandaloneForm(linkedPath, displayName: displayName);
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
            base.OnFormClosing(e);
        }
    }

    /// <summary>CSV内容查看控件</summary>
    public class CsvViewerControl : UserControl
    {
        private DataGridView _dataGridHeader = new();  // 冻结的表头行
        private DataGridView _dataGrid = new();        // 数据区域
        private SplitContainer _pnlGridContainer = new(); // 容器面板（使用SplitContainer实现冻结效果）
        private Button _btnSelectColumns = new();     // 列选择按钮
        private TextBox _txtContentFilter = new();     // 内容过滤框
        private TextBox _txtAdvancedFilter = new();    // 高级表达式过滤框
        private Label _lblInfo = new();
        private Label _lblCellPos = new();             // 当前编辑单元格行列提示
        private Panel _pnlLoading = new();             // 加载遮罩层
        private Panel _pnlTop = new();
        private DataTable _dataTable = new();
        private DataTable _headerTable = new();        // 冻结的表头数据
        private string _filePath = "";
        /// <summary>当前文件完整路径</summary>
        public string FilePath => _filePath;
        private char _delimiter = ',';
        private int _frozenRows = 0;                   // 冻结行数
        private bool _frozenRowsOverridden = false;     // 是否通过构造函数强制指定了冻结行数
        private int _headerRowIndex = -1;               // 表头行索引（-1=未设置，0-based，必须 < frozenRows）
        private int _paneIndex = -1;                    // 所属窗格索引（-1表示独立打开，使用全局设置）
        private Dictionary<string, string> _fileRemarks = new(); // 文件备注映射（key: 相对路径）
        private bool[] _columnCheckStates = Array.Empty<bool>(); // 列选中状态
        private List<ColumnLinkConfig> _columnLinks = new();
        private List<(string display, string fullPath)> _fileEntries = new();

        /// <summary>请求关联打开另一张表事件</summary>
        public event EventHandler? LinkTableRequested;

        // 关联表跳转高亮
        private int _highlightColumnIndex = -1;
        private string _highlightValue = "";
        private string _highlightColumnName = "";
        private HashSet<int> _highlightRows = new();
        private bool _highlightsReady = false;
        private List<string> _highlightValues = new(); // 多值高亮（支持 OR 匹配）

        public void SetHighlight(int colIndex, string value, string colName = "", ColumnLinkConfig? linkConfig = null)
        {
            _highlightColumnIndex = colIndex;
            _highlightValue = value ?? "";
            _highlightColumnName = colName ?? "";
            // 根据 linkConfig 提取多值用于 OR 匹配
            _highlightValues.Clear();
            if (linkConfig != null && linkConfig.ValueExtractionEnabled && !string.IsNullOrEmpty(value))
            {
                _highlightValues = linkConfig.ExtractValues(value);
                // 兜底：如果提取后只有一个值，但原始值中包含多值分隔符，尝试手动按多值分隔符拆分
                if (_highlightValues.Count == 1 && !string.IsNullOrEmpty(linkConfig.MultiValueDelimiter) && value.Contains(linkConfig.MultiValueDelimiter))
                {
                    var groups = value.Split(new[] { linkConfig.MultiValueDelimiter }, StringSplitOptions.RemoveEmptyEntries);
                    var fallback = new List<string>();
                    foreach (var g in groups)
                    {
                        var trimmed = g.Trim();
                        if (!string.IsNullOrEmpty(trimmed))
                        {
                            // 对每组也用 ExtractDelimiter 提取
                            if (!string.IsNullOrEmpty(linkConfig.ExtractDelimiter))
                            {
                                var parts = trimmed.Split(new[] { linkConfig.ExtractDelimiter }, StringSplitOptions.None);
                                if (linkConfig.ExtractPartIndex >= 0 && linkConfig.ExtractPartIndex < parts.Length)
                                    fallback.Add(parts[linkConfig.ExtractPartIndex].Trim());
                                else
                                    fallback.Add(trimmed);
                            }
                            else
                            {
                                fallback.Add(trimmed);
                            }
                        }
                    }
                    if (fallback.Count > 1)
                        _highlightValues = fallback.Distinct().ToList();
                }
            }
            else if (!string.IsNullOrEmpty(value))
            {
                _highlightValues = new List<string> { value };
            }
        }

        // 配色（与CsvReaderForm一致的浅色天蓝色调）
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

        public CsvViewerControl(string filePath, int paneIndex = -1, string? rootDir = null, int frozenRows = -1, int headerRowIndex = -1)
        {
            _filePath = filePath;
            _paneIndex = paneIndex;
            _rootDir = rootDir;
            if (frozenRows >= 0) { _frozenRows = frozenRows; _frozenRowsOverridden = true; }
            if (headerRowIndex >= 0) _headerRowIndex = headerRowIndex;
            InitializeComponent();
            LoadCsvAsync();
        }

        private string? _rootDir;
        /// <summary>当前浏览根目录（用于文件夹级设置存储）</summary>
        private string FrozenDir => !string.IsNullOrEmpty(_rootDir) ? _rootDir : Path.GetDirectoryName(_filePath) ?? "";

        private void InitializeComponent()
        {
            SuspendLayout();
            BackColor = _colorSurface;

            // 顶部工具栏 - 使用 TableLayoutPanel 实现两行布局
            var topTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                BackColor = _colorSurface,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(0),
                Margin = new Padding(0),
                GrowStyle = TableLayoutPanelGrowStyle.AddRows
            };
            topTable.RowStyles.Clear();
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 行0: 原有控件
            topTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 1)); // 行1: 分隔线
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 行2: 内容过滤
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 行3: 高级过滤

            // 第一行 - 原有 FlowLayoutPanel
            var flowTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Height = 40,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8, 5, 8, 5),
                AutoSize = true
            };

            // 冻结行数选择
            var lblFrozenRows = new Label
            {
                Text = "冻结行:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            var cboFrozenRows = new ComboBox
            {
                Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 10, 0)
            };
            cboFrozenRows.Items.AddRange(new object[] { "0", "1", "2", "3", "4", "5", "6", "7", "8" });
            // 从设置中读取冻结行数（按目录存储，子文件共享），若构造函数已强制指定则跳过
            if (!_frozenRowsOverridden)
                _frozenRows = !string.IsNullOrEmpty(FrozenDir) ? CsvSettings.GetCsvFrozenRowsByDirectory(FrozenDir) : CsvSettings.DefaultFrozenRows;
            cboFrozenRows.SelectedIndex = _frozenRows;
            cboFrozenRows.SelectedIndexChanged += (s, e) =>
            {
                _frozenRows = cboFrozenRows.SelectedIndex;
                // 保存设置（按目录存储）
                if (!string.IsNullOrEmpty(FrozenDir))
                    CsvSettings.SetCsvFrozenRowsByDirectory(FrozenDir, _frozenRows);
                else
                    CsvSettings.DefaultFrozenRows = _frozenRows;
                // 联动：如果表头行索引超出新的冻结行范围，自动调整
                if (_headerRowIndex >= _frozenRows)
                {
                    _headerRowIndex = _frozenRows > 0 ? 0 : -1;
                    UpdateHeaderRowCombo();
                    if (!string.IsNullOrEmpty(FrozenDir))
                        CsvSettings.SetCsvHeaderRowIndexByDirectory(FrozenDir, _headerRowIndex);
                    else
                        CsvSettings.DefaultFrozenRows = _headerRowIndex; // 复用保存机制
                }
                CsvSettings.Save();
                RefreshGridLayout();
            };

            // 表头行选择（从固定行中选择哪一行是字段名列）
            var lblHeaderRow = new Label
            {
                Text = "表头行:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            var cboHeaderRow = new ComboBox
            {
                Size = new Size(70, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 10, 0)
            };
            // 从设置读取表头行索引
            if (!_frozenRowsOverridden && !string.IsNullOrEmpty(FrozenDir))
                _headerRowIndex = CsvSettings.GetCsvHeaderRowIndexByDirectory(FrozenDir);
            PopulateHeaderRowCombo(cboHeaderRow);
            cboHeaderRow.SelectedIndexChanged += (s, e) =>
            {
                _headerRowIndex = cboHeaderRow.SelectedIndex >= 0 ? cboHeaderRow.SelectedIndex : -1; // selectedIndex 0=第1行
                // 保存设置（按目录存储）
                if (!string.IsNullOrEmpty(FrozenDir))
                    CsvSettings.SetCsvHeaderRowIndexByDirectory(FrozenDir, _headerRowIndex);
                CsvSettings.Save();
                RefreshGridLayout();
            };
            // 引用保持，供联动使用
            Tag = new { HeaderRowCombo = cboHeaderRow };

            // 表头过滤 - 列选择按钮
            var lblHeaderFilter = new Label
            {
                Text = "列选择:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            _btnSelectColumns = new Button
            {
                Size = new Size(120, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                Text = "选择列 (全部)",
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = _colorHoverBg,
                ForeColor = _colorLightText,
                FlatAppearance = { BorderSize = 1, BorderColor = _colorBorder },
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 2, 10, 0)
            };
            _btnSelectColumns.Click += OnSelectColumnsClick;

            // 信息标签
            _lblInfo = new Label
            {
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 0, 0)
            };

            flowTop.Controls.AddRange(new Control[] { lblFrozenRows, cboFrozenRows, lblHeaderRow, cboHeaderRow, lblHeaderFilter, _btnSelectColumns });

            // 重新加载按钮（保留过滤状态刷新文件内容）
            var btnReload = new Button
            {
                Text = "重新加载",
                Size = new Size(75, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 2, 10, 0)
            };
            btnReload.Click += (s, e) => ReloadContentAsync();
            flowTop.Controls.Add(btnReload);

            // 关联表按钮（已移除）

            // 行列位置标签
            var lblCellPos = new Label
            {
                Text = "",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                ForeColor = _colorPrimary,
                Margin = new Padding(10, 5, 10, 0)
            };
            flowTop.Controls.Add(lblCellPos);
            _lblCellPos = lblCellPos;

            // 信息标签放到最后
            flowTop.Controls.Add(_lblInfo);

            // 第二行（分隔线）放入 topTable
            var separatorLine = new Panel
            {
                Height = 1,
                Dock = DockStyle.Top,
                BackColor = Color.FromArgb(200, 225, 245),
                Margin = new Padding(0),
                Padding = new Padding(0)
            };


            // 内容过滤
            var lblContentFilter = new Label
            {
                Text = "内容过滤:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightText,
                Margin = new Padding(3, 5, 3, 0)
            };

            var txtContentFilter = new TextBox
            {
                Size = new Size(200, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 10, 0)
            };
            txtContentFilter.TextChanged += (s, e) => ApplyFilters();
            txtContentFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    txtContentFilter.Text = "";
                    e.Handled = true;
                }
            };

            var lblAdvFilter = new Label
            {
                Text = "高级过滤:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightText,
                Margin = new Padding(3, 5, 3, 0)
            };

            var txtAdvancedFilter = new TextBox
            {
                Size = new Size(200, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 10, 0),
                BackColor = Color.FromArgb(242, 249, 255)
            };
            txtAdvancedFilter.TextChanged += (s, e) => ApplyFilters();
            txtAdvancedFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    txtAdvancedFilter.Text = "";
                    e.Handled = true;
                }
            };

            var lblAdvHint = new Label
            {
                Text = "如: id=123 or name like test",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = _colorPrimary,
                Margin = new Padding(3, 5, 10, 0)
            };

            // 第三行 - 内容过滤
            var flowContent = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Height = 30,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8, 1, 8, 0),
                AutoSize = true
            };
            flowContent.Controls.AddRange(new Control[] { lblContentFilter, txtContentFilter });

            // 第四行 - 高级表达式过滤框
            var flowAdv = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Height = 30,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8, 1, 8, 5),
                AutoSize = true
            };
            flowAdv.Controls.AddRange(new Control[] { lblAdvFilter, txtAdvancedFilter, lblAdvHint });

            // 组装 topTable
            topTable.RowStyles.Clear();
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 行0: 原有控件
            topTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 1)); // 行1: 分隔线
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 行2: 内容过滤
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 行3: 高级过滤
            topTable.Controls.Add(flowTop, 0, 0);
            topTable.Controls.Add(separatorLine, 0, 1);
            topTable.Controls.Add(flowContent, 0, 2);
            topTable.Controls.Add(flowAdv, 0, 3);

            // 保存引用
            _pnlTop = topTable;
            _txtContentFilter = txtContentFilter;
            _txtAdvancedFilter = txtAdvancedFilter;

            // 过滤框宽度自适应：宽度 = 控件宽度 - 280px（给右侧提示文本留空间）
            topTable.Resize += (s, e) =>
            {
                int w = topTable.ClientSize.Width - 280;
                if (w < 100) w = 100;
                txtContentFilter.Width = w;
                txtAdvancedFilter.Width = w;
            };

            // 表格容器 - 使用SplitContainer来精确控制布局
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 28,
                FixedPanel = FixedPanel.Panel1,
                IsSplitterFixed = true,
                BackColor = _colorSurface,
                SplitterWidth = 1,
                BorderStyle = BorderStyle.FixedSingle
            };

            // 冻结的表头表格（显示冻结的行）
            _dataGridHeader = CreateDataGridView();
            _dataGridHeader.Dock = DockStyle.Fill;
            _dataGridHeader.ScrollBars = ScrollBars.None;
            _dataGridHeader.RowHeadersVisible = true;
            _dataGridHeader.RowHeadersWidth = 50;
            _dataGridHeader.ColumnHeadersVisible = true;
            _dataGridHeader.ReadOnly = true;
            _dataGridHeader.AllowUserToResizeColumns = false;
            _dataGridHeader.AllowUserToResizeRows = false;
            _dataGridHeader.RowPostPaint += OnFrozenRowPostPaint;
            _dataGridHeader.DefaultCellStyle.SelectionBackColor = _dataGridHeader.DefaultCellStyle.BackColor;
            _dataGridHeader.DefaultCellStyle.SelectionForeColor = _dataGridHeader.DefaultCellStyle.ForeColor;
            _dataGridHeader.DefaultCellStyle.BackColor = Color.FromArgb(245, 250, 255);
            _dataGridHeader.TabStop = false;
            _dataGridHeader.CellMouseDown += (s, e) =>
            {
                // 鼠标点击时同步重置，阻止滚动偏移发生
                if (e.RowIndex >= 0 && _dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;

                // 右键菜单：复制 + 编辑
                if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    var cellValue = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    var ctx = new ContextMenuStrip();
                    ctx.Font = new Font("Microsoft YaHei UI", 9f);

                    var copyItem = new ToolStripMenuItem("复制");
                    copyItem.Click += (s2, args) =>
                    {
                        if (!string.IsNullOrEmpty(cellValue))
                            Clipboard.SetText(cellValue);
                    };
                    ctx.Items.Add(copyItem);

                    var editItem = new ToolStripMenuItem("编辑");
                    editItem.Click += (s2, args) =>
                    {
                        var cell = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        _editingFromFrozen = true;
                        EditCell(cell, _dataGridHeader);
                    };
                    ctx.Items.Add(editItem);

                    ctx.Show(_dataGridHeader, _dataGridHeader.PointToClient(Cursor.Position));
                }
            };
            _dataGridHeader.CurrentCellChanged += (s, e) =>
            {
                // 键盘导航等情况的兜底处理
                if (_dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;
            };
            _dataGridHeader.CellDoubleClick += OnFrozenCellDoubleClick;
            _dataGridHeader.CellClick += OnCellClick;

            // 数据表格（主表格）
            _dataGrid = CreateDataGridView();
            _dataGrid.Dock = DockStyle.Fill;
            _dataGrid.RowHeadersWidth = 50;
            _dataGrid.RowPostPaint += OnRowPostPaint;
            _dataGrid.ColumnHeadersVisible = false;
            _dataGrid.CellMouseDown += OnGridCellMouseDown;

            // 同步水平滚动
            _dataGrid.Scroll += (s, e) =>
            {
                _dataGridHeader.HorizontalScrollingOffset = _dataGrid.HorizontalScrollingOffset;
            };

            // 同步列宽调整
            _dataGrid.ColumnWidthChanged += (s, e) =>
            {
                // 考虑主表格可能有隐藏的 __RowNum__ 列
                int dataGridColumnIndex = e.Column.Index;
                int headerColumnIndex = dataGridColumnIndex;

                // 如果主表格有隐藏列，需要调整索引
                if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__" && dataGridColumnIndex > 0)
                {
                    headerColumnIndex = dataGridColumnIndex - 1;
                }

                // 同步列宽到冻结表格
                if (headerColumnIndex >= 0 && headerColumnIndex < _dataGridHeader.Columns.Count)
                {
                    _dataGridHeader.Columns[headerColumnIndex].Width = _dataGrid.Columns[dataGridColumnIndex].Width;
                }
            };

            splitContainer.Panel1.Controls.Add(_dataGridHeader);
            splitContainer.Panel2.Controls.Add(_dataGrid);

            // 初始隐藏冻结面板
            splitContainer.Panel1Collapsed = true;
            _pnlGridContainer = splitContainer;

            Controls.Add(_pnlGridContainer);
            Controls.Add(_pnlTop);

            // 调整大小处理
            _pnlTop.Resize += (s, e) =>
            {
                _lblInfo.Left = _pnlTop.Width - _lblInfo.Width - 8;
            };

            ResumeLayout(false);
        }

        private DataGridView CreateDataGridView()
        {
            var dgv = new DataGridView
            {
                BackgroundColor = _colorSurface,
                BorderStyle = BorderStyle.None,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 28,
                RowHeadersVisible = true,
                RowHeadersWidth = 50,
                RowHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    BackColor = _colorRowNumBg,
                    ForeColor = _colorLightTextSecondary,
                    Font = new Font("Microsoft YaHei UI", 8.5f)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _colorAltRow
                },
                GridColor = _colorGridLine,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                EnableHeadersVisualStyles = false,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _colorHeaderBg,
                    ForeColor = _colorLightText,
                    Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    SelectionBackColor = _colorHeaderBg,
                    SelectionForeColor = _colorLightText
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _colorSurface,
                    ForeColor = _colorLightText,
                    Font = new Font("Microsoft YaHei UI", 9f),
                    SelectionBackColor = _colorSelectedBg,
                    SelectionForeColor = Color.FromArgb(20, 60, 100),
                    Padding = new Padding(2, 0, 2, 0)
                },
                RowTemplate = new DataGridViewRow { Height = 24 }
            };
            dgv.ColumnHeaderMouseClick += OnColumnHeaderClick;
            dgv.CellFormatting += OnCellFormatting;
            dgv.CellDoubleClick += OnCellDoubleClick;
            dgv.CellClick += OnCellClick;
            dgv.CellToolTipTextNeeded += OnCellToolTipTextNeeded;
            dgv.SelectionChanged += (s, e) => dgv.Invalidate();
            return dgv;
        }

        // ==================== 点击行列提示 + 双击弹窗编辑 ====================
        private bool _editingFromFrozen = false;

        private void OnCellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var dgv = (DataGridView)sender!;

            int rowOff = (dgv == _dataGridHeader) ? 0 : _frozenRows;
            int colIdx = e.ColumnIndex;
            if (dgv != _dataGridHeader && dgv.Columns.Count > 0 && dgv.Columns[0].Name == "__RowNum__")
                colIdx = Math.Max(0, colIdx - 1);

            string colName = colIdx >= 0 && colIdx < 26
                ? ((char)('A' + colIdx)).ToString()
                : $"Col{colIdx}";
            _lblCellPos.Text = $"行 {e.RowIndex + 1 + rowOff}  列 {colName}({colIdx})";
        }

        private void EditCell(DataGridViewCell cell, DataGridView dgv)
        {
            var rowIndex = cell.RowIndex;
            var columnIndex = cell.ColumnIndex;
            var currentValue = cell.Value?.ToString() ?? "";
            var (originalRowIndex, originalColumnIndex) = GetOriginalIndices(rowIndex, columnIndex);

            // 获取列名
            string colName = "";
            if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && originalColumnIndex >= 0 && originalColumnIndex < _dataTable.Columns.Count)
                colName = _dataTable.Rows[_headerRowIndex][originalColumnIndex]?.ToString() ?? "";

            var dialog = new CellEditDialog(currentValue, originalRowIndex, originalColumnIndex, _filePath, colName,
                onSave: newValue =>
                {
                    if (originalRowIndex >= 0 && originalRowIndex < _dataTable.Rows.Count &&
                        originalColumnIndex >= 0 && originalColumnIndex < _dataTable.Columns.Count)
                    {
                        _dataTable.Rows[originalRowIndex][originalColumnIndex] = newValue;
                        SaveCsvFile();
                        RefreshGridLayout();
                    }
                });
            dialog.Show(this);
        }

        private void OnCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var dgv = (DataGridView)sender!;
            if (dgv == _dataGridHeader) return;
            var cell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
            _editingFromFrozen = false;
            EditCell(cell, dgv);
        }

        private void OnFrozenCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var cell = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex];
            _editingFromFrozen = true;
            EditCell(cell, _dataGridHeader);
        }


        /// <summary>获取原始数据表中的行列索引</summary>
        private (int rowIndex, int columnIndex) GetOriginalIndices(int displayRowIndex, int displayColumnIndex)
        {
            // 如果来自冻结表格，行号直接就是原始行号（0-based）
            if (_editingFromFrozen)
            {
                int originalColumnIndex = displayColumnIndex;
                // 列索引转换（考虑可见列过滤）
                if (_visibleColumnIndices.Count > 0 && displayColumnIndex < _dataGridHeader.Columns.Count)
                {
                    if (displayColumnIndex < _visibleColumnIndices.Count)
                    {
                        originalColumnIndex = _visibleColumnIndices[displayColumnIndex];
                    }
                }
                return (displayRowIndex, originalColumnIndex);
            }

            // 如果使用了过滤显示，需要转换索引
            if (_dataGrid.DataSource is DataTable filteredTable && filteredTable.Columns.Count > 0)
            {
                // 获取原始行号（从隐藏列）
                int originalRowIndex = displayRowIndex;
                if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__")
                {
                    var cellValue = _dataGrid.Rows[displayRowIndex].Cells[0].Value;
                    if (cellValue != null)
                    {
                        originalRowIndex = Convert.ToInt32(cellValue) - 1; // 转换为0-based索引
                    }
                }

                // 列索引转换
                int originalColumnIndex = displayColumnIndex;
                if (_visibleColumnIndices.Count > 0 && displayColumnIndex < _dataGrid.Columns.Count)
                {
                    // 考虑隐藏列（__RowNum__）
                    var actualDisplayIndex = displayColumnIndex;
                    if (_dataGrid.Columns[0].Name == "__RowNum__")
                    {
                        actualDisplayIndex = displayColumnIndex - 1;
                    }

                    if (actualDisplayIndex >= 0 && actualDisplayIndex < _visibleColumnIndices.Count)
                    {
                        originalColumnIndex = _visibleColumnIndices[actualDisplayIndex];
                    }
                }

                return (originalRowIndex, originalColumnIndex);
            }

            // 无过滤时直接返回
            return (displayRowIndex + _frozenRows, displayColumnIndex);
        }

        /// <summary>保存CSV文件</summary>
        private void SaveCsvFile()
        {
            try
            {
                var sb = new StringBuilder();

                // 写入数据行
                foreach (DataRow row in _dataTable.Rows)
                {
                    var values = new List<string>();
                    foreach (var item in row.ItemArray)
                    {
                        var value = item?.ToString() ?? "";
                        // 如果值包含分隔符或引号，需要转义
                        if (value.Contains(_delimiter) || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
                        {
                            value = "\"" + value.Replace("\"", "\"\"") + "\"";
                        }
                        values.Add(value);
                    }
                    sb.AppendLine(string.Join(_delimiter.ToString(), values));
                }

                // 使用FileStream写入，允许读取共享（避免与其他进程冲突）
                var content = sb.ToString();
                var bytes = Encoding.UTF8.GetBytes(content);
                using (var fs = new FileStream(_filePath, FileMode.Create, FileAccess.Write, FileShare.Read))
                {
                    fs.Write(bytes, 0, bytes.Length);
                    fs.Flush();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"保存文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>填充表头行下拉框选项（根据当前冻结行数动态生成）</summary>
        private void PopulateHeaderRowCombo(ComboBox cbo)
        {
            cbo.Items.Clear();
            if (_frozenRows > 0)
            {
                for (int i = 0; i < _frozenRows; i++)
                    cbo.Items.Add($"第{i + 1}行");
                // 设置选中项
                if (_headerRowIndex >= 0 && _headerRowIndex < _frozenRows)
                    cbo.SelectedIndex = _headerRowIndex;
                else
                    cbo.SelectedIndex = -1;
            }
        }

        /// <summary>更新表头行下拉框（联动调用，供冻结行变更时使用）</summary>
        private void UpdateHeaderRowCombo()
        {
            // 通过 Tag 获取 ComboBox 引用
            if (Tag is ComboBox cboHdr)
            {
                PopulateHeaderRowCombo(cboHdr);
            }
        }

        /// <summary>获取表头行的列名列表（用于过滤、关联等场景）</summary>
        public List<string> GetHeaderColumnNames()
        {
            var names = new List<string>();
            if (_headerRowIndex < 0 || _headerRowIndex >= _dataTable.Rows.Count)
                return names;
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                names.Add(_dataTable.Rows[_headerRowIndex][i]?.ToString() ?? $"Col{i}");
            }
            return names;
        }

        /// <summary>根据字段名查找对应的原始列索引</summary>
        public int FindColumnIndexByHeaderName(string headerName)
        {
            if (_headerRowIndex < 0 || _headerRowIndex >= _dataTable.Rows.Count) return -1;
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                var val = _dataTable.Rows[_headerRowIndex][i]?.ToString() ?? "";
                if (val.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        private void RefreshGridLayout()
        {
            // 如果数据表还没有加载完成，直接返回
            if (_dataTable.Columns.Count == 0)
            {
                return;
            }

            if (_frozenRows > 0 && _dataTable.Rows.Count > 0)
            {
                // 先应用过滤，获取可见列索引
                _visibleColumnIndices = GetVisibleColumnIndices();

                // 创建冻结行数据表，只包含可见列
                _headerTable = new DataTable();
                foreach (var colIndex in _visibleColumnIndices)
                {
                    var sourceCol = _dataTable.Columns[colIndex];
                    _headerTable.Columns.Add(sourceCol.ColumnName, sourceCol.DataType);
                }

                // 从主数据表复制前N行（冻结行）到冻结行表
                for (int i = 0; i < Math.Min(_frozenRows, _dataTable.Rows.Count); i++)
                {
                    var sourceRow = _dataTable.Rows[i];
                    var targetRow = _headerTable.NewRow();
                    for (int j = 0; j < _visibleColumnIndices.Count; j++)
                    {
                        targetRow[j] = sourceRow[_visibleColumnIndices[j]];
                    }
                    _headerTable.Rows.Add(targetRow);
                }

                // 显示冻结面板并设置高度
                _pnlGridContainer.Panel1Collapsed = false;
                _pnlGridContainer.SplitterDistance = _frozenRows * 24 + 28; // 行高24 + 列标题28

                // 绑定数据到冻结表格
                _dataGridHeader.DataSource = _headerTable;

                // 设置冻结表格的列宽和列标题（确保按原始顺序显示）
                for (int i = 0; i < _dataGridHeader.Columns.Count; i++)
                {
                    var colName = _headerTable.Columns[i].ColumnName;
                    var dgvCol = _dataGridHeader.Columns[colName];
                    if (dgvCol != null)
                    {
                        dgvCol.DisplayIndex = i; // 强制按原始顺序显示
                        dgvCol.Width = 120;
                        dgvCol.HeaderText = GetExcelColumnName(i);
                        dgvCol.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                }

                // 应用内容过滤（设置主表格的数据源）
                var contentFilter = _txtContentFilter.Text.Trim();
                var visibleRows = GetVisibleRows(contentFilter);
                RefreshFilteredDisplay(visibleRows);
            }
            else
            {
                // 隐藏冻结面板，显示所有数据
                _pnlGridContainer.Panel1Collapsed = true;
                _dataGrid.DataSource = _dataTable;

                // 设置列宽和列标题
                for (int i = 0; i < _dataGrid.Columns.Count; i++)
                {
                    _dataGrid.Columns[i].Width = 120;
                    _dataGrid.Columns[i].HeaderText = GetExcelColumnName(i);
                    _dataGrid.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
                }
            }
        }

        /// <summary>重新加载文件内容，保留过滤条件和列选择状态</summary>
        private async void ReloadContentAsync()
        {
            if (!System.IO.File.Exists(_filePath)) return;

            // 保存当前过滤/选择状态
            var savedFilter = _txtContentFilter.Text;
            var savedAdvFilter = _txtAdvancedFilter.Text;
            var savedColumnStates = _columnCheckStates;
            var savedFrozenRows = _frozenRows;

            try
            {
                if (!this.IsHandleCreated)
                {
                    var tcs = new TaskCompletionSource<bool>();
                    EventHandler? handler = null;
                    handler = (s, e) =>
                    {
                        this.HandleCreated -= handler;
                        tcs.TrySetResult(true);
                    };
                    this.HandleCreated += handler;
                    await tcs.Task;
                }

                await Task.Run(() => ReloadCsvFile());
            }
            catch (Exception ex)
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(() =>
                    {
                        MessageBox.Show(this, $"重新加载CSV文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
                // 恢复状态
                _columnCheckStates = savedColumnStates;
                _frozenRows = savedFrozenRows;
                _txtContentFilter.Text = savedFilter;
                _txtAdvancedFilter.Text = savedAdvFilter;
                UpdateColumnSelectorButton();
                return;
            }

            // 重新应用列选择状态（如果列数未变）
            if (savedColumnStates.Length == _dataTable.Columns.Count)
            {
                _columnCheckStates = savedColumnStates;
            }
            else
            {
                // 列数变化，重置为全选
                _columnCheckStates = new bool[_dataTable.Columns.Count];
                for (int i = 0; i < _columnCheckStates.Length; i++)
                    _columnCheckStates[i] = true;
            }
            _frozenRows = savedFrozenRows;
            _txtContentFilter.Text = savedFilter;
            _txtAdvancedFilter.Text = savedAdvFilter;
            UpdateColumnSelectorButton();
            RefreshGridLayout();
        }

        /// <summary>重新读取CSV文件内容（后台线程执行）</summary>
        private void ReloadCsvFile()
        {
            if (!File.Exists(_filePath)) return;

            string content;
            using (var fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(fs, Encoding.UTF8, true))
            {
                content = reader.ReadToEnd();
            }

            var delimiter = DetectDelimiter(content);
            var lines = SplitCsvLines(content);
            if (lines.Count == 0) return;

            var headers = ParseCsvLine(lines[0]);
            var dataTable = new DataTable();
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var header in headers)
            {
                var columnName = header;
                int suffix = 1;
                while (usedNames.Contains(columnName))
                {
                    columnName = $"{header}_{suffix}";
                    suffix++;
                }
                usedNames.Add(columnName);
                dataTable.Columns.Add(columnName, typeof(string));
            }

            for (int i = 0; i < lines.Count; i++)
            {
                var values = ParseCsvLine(lines[i]);
                var row = dataTable.NewRow();
                for (int j = 0; j < Math.Min(values.Length, headers.Length); j++)
                    row[j] = values[j];
                dataTable.Rows.Add(row);
            }

            this.Invoke(() =>
            {
                _delimiter = delimiter;
                _dataTable = dataTable;
                _headerTable = dataTable.Clone();
            });
        }

        private async void LoadCsvAsync()
        {
            try
            {
                // 等待句柄创建完成
                if (!this.IsHandleCreated)
                {
                    var tcs = new TaskCompletionSource<bool>();
                    EventHandler? handler = null;
                    handler = (s, e) =>
                    {
                        this.HandleCreated -= handler;
                        tcs.TrySetResult(true);
                    };
                    this.HandleCreated += handler;
                    await tcs.Task;
                }

                // 显示加载遮罩
                var loadingStart = DateTime.UtcNow;
                ShowLoadingOverlay();
                // 短暂延迟确保 loading 遮罩渲染到屏幕上
                await Task.Delay(50);

                await Task.Run(() => LoadCsvFile());

                // 确保 loading 至少显示了 200ms，避免一闪而过
                var elapsed = (DateTime.UtcNow - loadingStart).TotalMilliseconds;
                if (elapsed < 200)
                    await Task.Delay((int)(200 - elapsed));

                // 隐藏加载遮罩
                HideLoadingOverlay();
            }
            catch (Exception ex)
            {
                HideLoadingOverlay();
                if (this.IsHandleCreated)
                {
                    this.Invoke(() =>
                    {
                        MessageBox.Show(this, $"加载CSV文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            }
        }

        private void ShowLoadingOverlay()
        {
            if (!this.IsHandleCreated) return;
            this.Invoke(() =>
            {
                if (_pnlLoading?.IsDisposed != false || _pnlLoading.Parent != null) return;
                _pnlLoading = new Panel
                {
                    BackColor = Color.FromArgb(180, 255, 255, 255),
                    Dock = DockStyle.Fill,
                    Visible = true
                };
                var lbl = new Label
                {
                    Text = "正在加载数据，请稍候...",
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Microsoft YaHei UI", 14f, FontStyle.Bold),
                    ForeColor = Color.FromArgb(30, 136, 229),
                    Dock = DockStyle.Fill
                };
                _pnlLoading.Controls.Add(lbl);
                this.Controls.Add(_pnlLoading);
                _pnlLoading.BringToFront();
                Application.DoEvents();
            });
        }

        private void HideLoadingOverlay()
        {
            if (!this.IsHandleCreated) return;
            this.Invoke(() =>
            {
                if (_pnlLoading != null && !_pnlLoading.IsDisposed && _pnlLoading.Parent == this)
                {
                    this.Controls.Remove(_pnlLoading);
                    _pnlLoading.Dispose();
                    _pnlLoading = null;
                }
            });
        }

        private void LoadCsvFile()
        {
            if (!File.Exists(_filePath))
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, "CSV文件不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error));
                return;
            }

            // 读取文件内容以检测分隔符
            string content;
            using (var fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(fs, Encoding.UTF8, true))
            {
                content = reader.ReadToEnd();
            }

            // 检测分隔符
            _delimiter = DetectDelimiter(content);

            // 解析CSV
            var lines = SplitCsvLines(content);
            if (lines.Count == 0)
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, "CSV文件为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning));
                return;
            }

            // 解析表头
            var headers = ParseCsvLine(lines[0]);

            // 准备数据表结构
            var dataTable = new DataTable();
            var headerTable = new DataTable();
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var header in headers)
            {
                var columnName = header;
                int suffix = 1;
                // 处理重复列名
                while (usedNames.Contains(columnName))
                {
                    columnName = $"{header}_{suffix}";
                    suffix++;
                }
                usedNames.Add(columnName);
                dataTable.Columns.Add(columnName, typeof(string));
                headerTable.Columns.Add(columnName, typeof(string));
            }

            // 解析数据行（包括第一行，因为第一行可能是数据而非表头）
            for (int i = 0; i < lines.Count; i++)
            {
                var values = ParseCsvLine(lines[i]);
                var row = dataTable.NewRow();

                for (int j = 0; j < Math.Min(values.Length, headers.Length); j++)
                {
                    row[j] = values[j];
                }

                dataTable.Rows.Add(row);
            }

            // 在UI线程更新控件
            this.Invoke(() =>
            {
                _dataTable = dataTable;
                _headerTable = headerTable.Clone();

                // 初始化列选中状态（默认全选，第一列始终选中）
                _columnCheckStates = new bool[_dataTable.Columns.Count];
                for (int i = 0; i < _columnCheckStates.Length; i++)
                    _columnCheckStates[i] = true;

                _columnLinks = CsvSettings.GetColumnLinksByDirectory(FrozenDir);
                ColumnLinkChecker.BuildLinkLookups(_columnLinks, FrozenDir, _filePath);

                UpdateColumnSelectorButton();
                RebuildHighlightRows();

                // 更新信息
                _lblInfo.Text = $"分隔符: '{_delimiter}' | 行数: {_dataTable.Rows.Count} | 列数: {_dataTable.Columns.Count}";

                // 根据是否有冻结行来决定如何显示数据
                if (_frozenRows > 0)
                {
                    RefreshGridLayout();
                }
                else
                {
                    // 无冻结行时，直接绑定所有数据到主表格
                    _dataGrid.DataSource = _dataTable;

                    // 设置列宽和Excel样式列标题
                    for (int i = 0; i < _dataGrid.Columns.Count; i++)
                    {
                        var col = _dataGrid.Columns[i];
                        col.Width = 120;
                        col.SortMode = DataGridViewColumnSortMode.Programmatic;
                        col.HeaderText = GetExcelColumnName(i);
                    }
                }
            });
        }

        private char DetectDelimiter(string content)
        {
            var lines = SplitCsvLines(content);
            if (lines.Count == 0) return ',';

            var firstLine = lines[0];

            // 统计各分隔符出现次数
            var counts = new Dictionary<char, int>
            {
                { ',', firstLine.Count(c => c == ',') },
                { '\t', firstLine.Count(c => c == '\t') },
                { ';', firstLine.Count(c => c == ';') },
                { '|', firstLine.Count(c => c == '|') }
            };

            // 选择出现次数最多的分隔符
            return counts.OrderByDescending(kv => kv.Value).First().Key;
        }

        /// <summary>
        /// 按RFC 4180标准解析CSV内容为多行，兼容单元格内引号包裹的换行。
        /// 行内的 "" 转义为单个双引号保留在结果中。
        /// </summary>
        private List<string> SplitCsvLines(string content)
        {
            var lines = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < content.Length; i++)
            {
                char c = content[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < content.Length && content[i + 1] == '"')
                    {
                        // 转义双引号: "" -> "
                        sb.Append('"');
                        i++;
                        continue;
                    }
                    inQuotes = !inQuotes;
                    sb.Append(c);
                }
                else if (!inQuotes && (c == '\n' || c == '\r'))
                {
                    // 在引号外的换行符才是真正的行结束符
                    // 跳过 \r\n 中的 
                    if (c == '\r' && i + 1 < content.Length && content[i + 1] == '\n')
                    {
                        i++; // 跳过 

                    }
                    lines.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }

            // 最后一段
            if (sb.Length > 0)
            {
                lines.Add(sb.ToString());
            }

            return lines;
        }

        private string[] ParseCsvLine(string line)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // 转义的引号
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == _delimiter && !inQuotes)
                {
                    result.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }

            result.Add(sb.ToString());
            return result.ToArray();
        }

        // ==================== 排序功能 ====================
        private int _sortColumnIndex = -1;
        private SortOrder _sortOrder = SortOrder.None;

        private void OnColumnHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.ColumnIndex >= _dataGrid.Columns.Count)
                return;

            // 切换排序方向
            if (_sortColumnIndex == e.ColumnIndex)
            {
                _sortOrder = _sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                _sortColumnIndex = e.ColumnIndex;
                _sortOrder = SortOrder.Ascending;
            }

            // 执行排序（只排序数据部分，冻结行保持固定）
            var columnName = _dataGrid.Columns[e.ColumnIndex].DataPropertyName;

            // 分离冻结行和数据行
            var frozenRows = _dataTable.AsEnumerable().Take(_frozenRows).ToList();
            var dataRows = _dataTable.AsEnumerable().Skip(_frozenRows).ToList();

            // 只排序数据行
            var sortedDataRows = dataRows
                .OrderBy(row => row.Field<string>(columnName), new NaturalStringComparer(_sortOrder))
                .ToList();

            // 重新组合：冻结行 + 排序后的数据行
            var allRows = frozenRows.Concat(sortedDataRows).ToArray();
            _dataTable = allRows.CopyToDataTable();

            // 刷新布局以重新分配冻结行和数据行
            RefreshGridLayout();

            // 更新列头显示
            UpdateColumnHeaderSortGlyph();
        }

        private void UpdateColumnHeaderSortGlyph()
        {
            foreach (DataGridViewColumn col in _dataGrid.Columns)
            {
                if (col.Index == _sortColumnIndex)
                {
                    col.HeaderCell.SortGlyphDirection = _sortOrder;
                }
                else
                {
                    col.HeaderCell.SortGlyphDirection = SortOrder.None;
                }
            }
        }

        // ==================== 过滤功能 ====================
        private List<int> _visibleColumnIndices = new(); // 当前可见的列索引（第一列始终可见）

        private void ApplyFilters()
        {
            // 统一使用 RefreshGridLayout 处理所有情况
            RefreshGridLayout();
        }

        /// <summary>获取可见的列索引（第一列始终可见，基于多选框状态）</summary>
        private List<int> GetVisibleColumnIndices()
        {
            var visibleIndices = new List<int>();

            // 第一列始终可见
            if (_dataTable.Columns.Count > 0)
            {
                visibleIndices.Add(0);
            }

            // 根据 _columnCheckStates 过滤列
            for (int i = 1; i < _dataTable.Columns.Count; i++)
            {
                if (i < _columnCheckStates.Length && _columnCheckStates[i])
                {
                    visibleIndices.Add(i);
                }
            }

            return visibleIndices;
        }

        /// <summary>点击"选择列"按钮</summary>
        private void OnSelectColumnsClick(object? sender, EventArgs e)
        {
            if (_dataTable.Columns.Count == 0) return;

            EnsureFileEntriesLoaded();

            using var dialog = new ColumnSelectorDialog(_dataTable, _columnCheckStates, _frozenRows, 0,
                _rootDir ?? "", _filePath, _fileEntries, _columnLinks, _headerRowIndex);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _columnCheckStates = dialog.GetCheckedStates();
                _columnLinks = dialog.GetColumnLinks();
                CsvSettings.SetColumnLinksByDirectory(FrozenDir, _columnLinks);
                ColumnLinkChecker.BuildLinkLookups(_columnLinks, FrozenDir, _filePath);
                UpdateColumnSelectorButton();
                RefreshGridLayout();
            }
        }

        /// <summary>确保文件条目列表已加载（从 _rootDir 扫描）</summary>
        private void EnsureFileEntriesLoaded()
        {
            if (_fileEntries.Count > 0) return;
            var root = FrozenDir;
            if (string.IsNullOrEmpty(root) || !Directory.Exists(root)) return;
            _fileEntries.Clear();
            try
            {
                var showHidden = CsvSettings.ShowHiddenFiles;
                foreach (var f in Directory.GetFiles(root, "*.csv", SearchOption.AllDirectories))
                {
                    if (!showHidden)
                    {
                        try { if ((File.GetAttributes(f) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                    }
                    var rel = f.StartsWith(root, StringComparison.OrdinalIgnoreCase)
                        ? f.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar) : Path.GetFileName(f);
                    _fileEntries.Add((rel, f));
                }
                foreach (var ext in new[] { "*.xls", "*.xlsx" })
                {
                    foreach (var f in Directory.GetFiles(root, ext, SearchOption.AllDirectories))
                    {
                        if (!showHidden)
                        {
                            try { if ((File.GetAttributes(f) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                        }
                        var rel = f.StartsWith(root, StringComparison.OrdinalIgnoreCase)
                            ? f.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar) : Path.GetFileName(f);
                        _fileEntries.Add((rel, f));
                    }
                }
            }
            catch { }
        }

        /// <summary>更新选择列按钮文本</summary>
        private void UpdateColumnSelectorButton()
        {
            if (_dataTable.Columns.Count == 0) return;
            int checkedCount = 0;
            for (int i = 0; i < _columnCheckStates.Length; i++)
                if (_columnCheckStates[i]) checkedCount++;

            // 第一列不参与计数
            int dataColumns = Math.Max(0, _columnCheckStates.Length - 1);
            int checkedData = Math.Max(0, checkedCount - 1);
            _btnSelectColumns.Text = $"选择列 ({checkedData}/{dataColumns})";
        }

        /// <summary>获取可见的行（从冻结行之后开始）</summary>
        private List<DataRow> GetVisibleRows(string contentFilter)
        {
            var visibleRows = new List<DataRow>();
            var advFilter = _txtAdvancedFilter.Text.Trim();

            // 从冻结行之后开始遍历
            for (int i = _frozenRows; i < _dataTable.Rows.Count; i++)
            {
                var row = _dataTable.Rows[i];

                // 1) 简单内容过滤（基于 MatchesFilter 的文本匹配）
                bool contentMatch = string.IsNullOrEmpty(contentFilter); // 空则通过
                if (!contentMatch)
                {
                    var processedFilter = contentFilter.Replace("||", " | ").Replace("&&", " & ");

                    foreach (var colIndex in _visibleColumnIndices)
                    {
                        var cellValue = row[colIndex]?.ToString() ?? "";
                        if (MatchesFilter(cellValue, processedFilter))
                        {
                            contentMatch = true;
                            break;
                        }
                    }
                }

                // 2) 高级表达式过滤
                bool advMatch = string.IsNullOrEmpty(advFilter);
                if (!advMatch)
                {
                    var columnsToSearch = _visibleColumnIndices;
                    advMatch = EvaluateExpression(row, advFilter, columnsToSearch);
                }

                // 两个过滤条件同时满足（AND 逻辑）
                if (contentMatch && advMatch)
                {
                    visibleRows.Add(row);
                }
            }

            return visibleRows;
        }

        /// <summary>刷新过滤后的显示</summary>
        private void RefreshFilteredDisplay(List<DataRow> visibleRows)
        {
            // 创建新的数据表，只包含可见列，并添加原始行号列
            var filteredTable = new DataTable();

            // 添加原始行号列（第一列，用于显示）
            filteredTable.Columns.Add("__RowNum__", typeof(int));

            // 添加可见列
            foreach (var colIndex in _visibleColumnIndices)
            {
                var sourceCol = _dataTable.Columns[colIndex];
                filteredTable.Columns.Add(sourceCol.ColumnName, sourceCol.DataType);
            }

            // 添加可见行
            foreach (var sourceRow in visibleRows)
            {
                var newRow = filteredTable.NewRow();
                // 保存原始行号（从冻结行之后开始计数）
                var originalRowIndex = _dataTable.Rows.IndexOf(sourceRow) + 1;
                newRow[0] = originalRowIndex;

                for (int i = 0; i < _visibleColumnIndices.Count; i++)
                {
                    newRow[i + 1] = sourceRow[_visibleColumnIndices[i]];
                }
                filteredTable.Rows.Add(newRow);
            }

            // 绑定到主表格
            _dataGrid.DataSource = filteredTable;

            // 隐藏原始行号列（我们使用行头显示）
            if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__")
            {
                _dataGrid.Columns[0].Visible = false;
            }

            // 设置列宽和列标题（确保按原始顺序显示）
            // filteredTable的列顺序是: __RowNum__, 可见列1, 可见列2, ...
            // 可见列的顺序由_visibleColumnIndices决定
            for (int i = 1; i < _dataGrid.Columns.Count; i++)
            {
                var colName = filteredTable.Columns[i].ColumnName;
                var dgvCol = _dataGrid.Columns[colName];
                if (dgvCol != null)
                {
                    dgvCol.DisplayIndex = i; // 强制按原始顺序显示
                    dgvCol.Width = 120;
                    dgvCol.HeaderText = GetExcelColumnName(i - 1); // -1 因为第0列是行号
                    dgvCol.SortMode = DataGridViewColumnSortMode.Programmatic;
                }
            }

            // 更新信息
            _lblInfo.Text = $"显示 {_visibleColumnIndices.Count} 列, {visibleRows.Count} 行";
        }

        /// <summary>递归评估布尔表达式</summary>
        private bool EvaluateExpression(DataRow row, string expr, List<int> columnsToSearch)
        {
            expr = expr.Trim();
            if (string.IsNullOrEmpty(expr))
                return true;

            // 去除外层括号
            while (expr.StartsWith("(") && expr.EndsWith(")") && IsBalanced(expr))
            {
                expr = expr.Substring(1, expr.Length - 2).Trim();
            }

            // 找出外部最顶层的 or/and（不在括号内）
            int parenDepth = 0;
            int orPos = -1;
            int andPos = -1;

            for (int i = 0; i < expr.Length; i++)
            {
                if (expr[i] == '(')
                    parenDepth++;
                else if (expr[i] == ')')
                    parenDepth--;
                else if (parenDepth == 0)
                {
                    // 匹配 or（不区分大小写，注意单词边界）
                    if (i + 2 < expr.Length &&
                        (expr[i] == 'o' || expr[i] == 'O') &&
                        (expr[i + 1] == 'r' || expr[i + 1] == 'R') &&
                        (i + 2 >= expr.Length || !char.IsLetterOrDigit(expr[i + 2])) &&
                        (i == 0 || !char.IsLetterOrDigit(expr[i - 1])))
                    {
                        orPos = i;
                        break; // or 优先级低于 and，优先分割
                    }

                    // 匹配 and
                    if (i + 3 < expr.Length &&
                        (expr[i] == 'a' || expr[i] == 'A') &&
                        (expr[i + 1] == 'n' || expr[i + 1] == 'N') &&
                        (expr[i + 2] == 'd' || expr[i + 2] == 'D') &&
                        (i + 3 >= expr.Length || !char.IsLetterOrDigit(expr[i + 3])) &&
                        (i == 0 || !char.IsLetterOrDigit(expr[i - 1])))
                    {
                        // 先记录 and，但继续搜索。如果找到 or，优先用 or 分割
                        if (andPos < 0)
                            andPos = i;
                        i += 3; // skip "and"
                    }
                }
            }

            if (orPos >= 0)
            {
                // 按 or 分割
                var left = expr.Substring(0, orPos).Trim();
                var right = expr.Substring(orPos + 2).Trim();
                return EvaluateExpression(row, left, columnsToSearch) ||
                       EvaluateExpression(row, right, columnsToSearch);
            }

            if (andPos >= 0)
            {
                var left = expr.Substring(0, andPos).Trim();
                var right = expr.Substring(andPos + 3).Trim();
                return EvaluateExpression(row, left, columnsToSearch) &&
                       EvaluateExpression(row, right, columnsToSearch);
            }

            // 基础条件，解析比较操作
            return EvaluateCondition(row, expr, columnsToSearch);
        }

        /// <summary>检查括号是否平衡</summary>
        private bool IsBalanced(string s)
        {
            int d = 0;
            foreach (char c in s)
            {
                if (c == '(') d++;
                else if (c == ')') d--;
                if (d < 0) return false;
            }
            return d == 0;
        }

        /// <summary>评估单个条件表达式</summary>
        private bool EvaluateCondition(DataRow row, string condition, List<int> columnsToSearch)
        {
            condition = condition.Trim();

            // 检测比较操作符: like, =, !=, &lt;&gt;, &gt;, &gt;=, &lt;, &lt;=
            // 注意 like 优先级高于 =
            string op = "";
            int opPos = -1;

            // 先找 like (不区分大小写)
            var likeIdx = condition.IndexOf(" like ", StringComparison.OrdinalIgnoreCase);
            if (likeIdx >= 0)
            {
                op = "like";
                opPos = likeIdx;
            }
            else
            {
                // 找 !=
                var neqIdx = FindOperator(condition, "!=");
                var ltgtIdx = FindOperator(condition, "<>");
                var geIdx = FindOperator(condition, ">=");
                var leIdx = FindOperator(condition, "<=");
                var gtIdx = FindOperator(condition, ">");
                var ltIdx = FindOperator(condition, "<");
                var eqIdx = FindOperator(condition, "=");

                // 按优先级选择（先长的操作符以避免错误匹配）
                if (geIdx >= 0) { op = ">="; opPos = geIdx; }
                else if (leIdx >= 0) { op = "<="; opPos = leIdx; }
                else if (neqIdx >= 0) { op = "!="; opPos = neqIdx; }
                else if (ltgtIdx >= 0) { op = "<>"; opPos = ltgtIdx; }
                else if (gtIdx >= 0) { op = ">"; opPos = gtIdx; }
                else if (ltIdx >= 0) { op = "<"; opPos = ltIdx; }
                else if (eqIdx >= 0) { op = "="; opPos = eqIdx; }
            }

            if (opPos < 0)
            {
                // 没有操作符，使用全文搜索（在所有列中搜索整个条件文本）
                foreach (var colIndex in columnsToSearch)
                {
                    var cellValue = row[colIndex]?.ToString() ?? "";
                    if (cellValue.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
                return false;
            }

            // 提取列名和值
            var rawColumnName = condition.Substring(0, opPos).Trim();
            // like 匹配的是 " like "（6个字符），其他操作符按实际长度
            int opRealLen = op == "like" ? " like ".Length : op.Length;
            var rawValue = condition.Substring(opPos + opRealLen).Trim();

            var columnName = rawColumnName.Trim();

            // 处理值（移除可能的引号）
            var value = rawValue;
            if ((value.StartsWith("'") && value.EndsWith("'")) ||
                (value.StartsWith("\"") && value.EndsWith("\"")))
            {
                value = value.Substring(1, value.Length - 2);
            }

            // 检查是否是指定列名过滤
            if (!string.IsNullOrEmpty(columnName) && char.IsLetter(columnName[0]))
            {
                // 在指定列中查找
                int colIndex = ResolveColumnIndex(columnName);

                if (colIndex >= 0 && columnsToSearch.Contains(colIndex))
                {
                    var cellValue = row[colIndex]?.ToString() ?? "";
                    return CompareValue(cellValue, value, op);
                }

                // 列名未找到，回退到所有列搜索（但将列名视为普通搜索文本）
                foreach (var ci in columnsToSearch)
                {
                    var cellValue = row[ci]?.ToString() ?? "";
                    if (cellValue.IndexOf(condition.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
                return false;
            }

            // 没有列名指定，在所有列中搜索
            foreach (var colIndex in columnsToSearch)
            {
                var cellValue = row[colIndex]?.ToString() ?? "";
                if (CompareValue(cellValue, columnName + op + rawValue, op))
                    return true;
                // 也试一下全文匹配
                if (cellValue.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return false;
        }

        /// <summary>根据列名字符串解析对应的列索引，优先匹配表头行，再匹配冻结行范围内的单元格内容</summary>
        private int ResolveColumnIndex(string columnName)
        {
            // 1. 尝试精确匹配 ColumnName（DataTable 的原始列名）
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                if (_dataTable.Columns[i].ColumnName.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            // 2. 如果设置了表头行，优先在表头行搜索
            if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count)
            {
                for (int i = 0; i < _dataTable.Columns.Count; i++)
                {
                    var cellValue = _dataTable.Rows[_headerRowIndex][i]?.ToString() ?? "";
                    if (cellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        return i;
                    }
                }
            }

            // 3. 回退：在冻结行范围内所有行搜索
            int maxHeaderScan = Math.Min(_frozenRows, _dataTable.Rows.Count);
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                for (int row = 0; row < maxHeaderScan; row++)
                {
                    if (row == _headerRowIndex) continue; // 已在上面搜索过
                    var cellValue = _dataTable.Rows[row][i]?.ToString() ?? "";
                    if (cellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        return i;
                    }
                }
            }

            return -1;
        }

        /// <summary>在字符串中查找操作符（确保不是单词中间）</summary>
        private int FindOperator(string s, string op)
        {
            int idx = s.IndexOf(op, StringComparison.Ordinal);
            while (idx >= 0)
            {
                int nextIdx = idx + op.Length;
                // 检查后面是否跟着 =，避免 > 错误匹配 >= 中的 >
                if (nextIdx < s.Length && op.Length == 1 && (op[0] == '>' || op[0] == '<' || op[0] == '!') && s[nextIdx] == '=')
                {
                    idx = s.IndexOf(op, nextIdx + 1, StringComparison.Ordinal);
                    continue;
                }
                return idx;
            }
            return -1;
        }

        /// <summary>比较单元格值和条件值</summary>
        private bool CompareValue(string cellValue, string filterValue, string op)
        {
            // 尝试数字比较
            bool numericCompare = false;
            double cellNum = 0, filterNum = 0;
            if (op != "like" && double.TryParse(cellValue, out cellNum) && double.TryParse(filterValue, out filterNum))
            {
                numericCompare = true;
            }

            switch (op.ToLower())
            {
                case "like":
                    // like: 支持 * 和 % 通配符，默认包含匹配
                    if (filterValue.Contains("*") || filterValue.Contains("%"))
                    {
                        var pattern = Regex.Escape(filterValue)
                            .Replace("\\*", ".*")
                            .Replace("\\%", ".*");
                        return Regex.IsMatch(cellValue, "^" + pattern + "$", RegexOptions.IgnoreCase);
                    }
                    else
                    {
                        // 无通配符时，默认包含匹配（两端自动加 %）
                        return cellValue.IndexOf(filterValue, StringComparison.OrdinalIgnoreCase) >= 0;
                    }

                case "=":
                    return numericCompare
                        ? Math.Abs(cellNum - filterNum) < 0.000001
                        : cellValue.Equals(filterValue, StringComparison.OrdinalIgnoreCase);

                case "!=":
                case "<>":
                    return numericCompare
                        ? Math.Abs(cellNum - filterNum) >= 0.000001
                        : !cellValue.Equals(filterValue, StringComparison.OrdinalIgnoreCase);

                case ">":
                    return numericCompare ? cellNum > filterNum
                        : string.Compare(cellValue, filterValue, StringComparison.OrdinalIgnoreCase) > 0;

                case ">=":
                    return numericCompare ? cellNum >= filterNum
                        : string.Compare(cellValue, filterValue, StringComparison.OrdinalIgnoreCase) >= 0;

                case "<":
                    return numericCompare ? cellNum < filterNum
                        : string.Compare(cellValue, filterValue, StringComparison.OrdinalIgnoreCase) < 0;

                case "<=":
                    return numericCompare ? cellNum <= filterNum
                        : string.Compare(cellValue, filterValue, StringComparison.OrdinalIgnoreCase) <= 0;

                default:
                    return cellValue.IndexOf(filterValue, StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        /// <summary>简单的文本匹配（用于列名/列内容的表头过滤）</summary>
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

        private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            // 确保行高一致
            if (e.RowIndex >= 0 && e.RowIndex < _dataGrid.Rows.Count)
            {
                _dataGrid.Rows[e.RowIndex].Height = 24;
            }
            // 列关联颜色检查
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                ColumnLinkChecker.CheckCellColor(e, _visibleColumnIndices, _dataGrid, _filePath, FrozenDir, _frozenRows, _columnLinks);
                // 关联表跳转高亮 — 整行变色
                if (_highlightsReady && _highlightRows.Contains(e.RowIndex))
                {
                    e.CellStyle.BackColor = Color.FromArgb(255, 255, 200);
                    e.CellStyle.SelectionBackColor = Color.FromArgb(255, 240, 150);
                }
            }
        }

        /// <summary>单元格提示 — 关联字段值无效时显示红色原因</summary>
        private void OnCellToolTipTextNeeded(object? sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            e.ToolTipText = ColumnLinkChecker.GetInvalidLinkToolTip(
                _filePath, FrozenDir, _visibleColumnIndices, _dataGrid, e.RowIndex, e.ColumnIndex, _frozenRows, _columnLinks);
        }

        /// <summary>右键菜单 — 复制单元格内容 + 打开关联表</summary>
        private void OnGridCellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right || e.RowIndex < 0 || e.ColumnIndex < 0) return;

            int adjCol = e.ColumnIndex;
            if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__")
                adjCol = e.ColumnIndex - 1;
            int origCol = adjCol;
            if (adjCol >= 0 && adjCol < _visibleColumnIndices.Count)
                origCol = _visibleColumnIndices[adjCol];

            string currentRel = GetCurrentRelativePath();
            var link = _columnLinks.FirstOrDefault(l =>
                l.SourceFile.Equals(currentRel, StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == origCol);

            var cellValue = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
            var ctx = new ContextMenuStrip();
            ctx.Font = new Font("Microsoft YaHei UI", 9f);

            // 1. 复制当前单元格的全部内容
            var copyItem = new ToolStripMenuItem("复制");
            copyItem.Click += (s, args) =>
            {
                if (!string.IsNullOrEmpty(cellValue))
                {
                    Clipboard.SetText(cellValue);
                }
            };
            ctx.Items.Add(copyItem);

            // 1.5 编辑单元格
            var editItem = new ToolStripMenuItem("编辑");
            editItem.Click += (s, args) =>
            {
                var cell = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                _editingFromFrozen = false;
                EditCell(cell, _dataGrid);
            };
            ctx.Items.Add(editItem);

            // 2. 打开关联表（有链接才显示，但根据值有效性决定是否可用）
            var linkItem = new ToolStripMenuItem("打开关联表");
            if (link == null)
            {
                linkItem.Enabled = false;
            }
            else
            {
                bool isValid = !string.IsNullOrEmpty(cellValue) && ColumnLinkChecker.IsLinkValueValid(_filePath, FrozenDir, origCol, cellValue, _columnLinks);
                linkItem.Enabled = isValid;
                linkItem.Click += (s, args) =>
                {
                    string targetFullPath = Path.Combine(FrozenDir, link.TargetFile);
                    if (!File.Exists(targetFullPath))
                    {
                        MessageBox.Show(this, "关联文件不存在: " + link.TargetFile, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    var remarks = CsvSettings.GetFileRemarksByDirectory(FrozenDir);
                    var displayName = Path.GetFileName(targetFullPath);
                    if (remarks.TryGetValue(link.TargetFile, out var remark) && !string.IsNullOrEmpty(remark))
                        displayName = remark + " (" + Path.GetFileName(targetFullPath) + ")";
                    var standaloneForm = new CsvReaderStandaloneForm(targetFullPath, link.TargetColumnIndex, cellValue, _frozenRows, link.TargetColumnName, _headerRowIndex, isLinkedFromCell: true, linkConfig: link, displayName: displayName);
                    standaloneForm.Show(this);
                };
            }
            ctx.Items.Add(linkItem);

            ctx.Show(_dataGrid, _dataGrid.PointToClient(Cursor.Position));
        }

        private string GetCurrentRelativePath()
        {
            if (!string.IsNullOrEmpty(FrozenDir) && _filePath.StartsWith(FrozenDir, StringComparison.OrdinalIgnoreCase))
                return _filePath.Substring(FrozenDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(_filePath);
        }

        private void OnRowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            // 绘制行号（Excel样式）
            // 从隐藏列中读取原始行号
            int rowIndex;
            if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__" && e.RowIndex >= 0)
            {
                // 使用保存的原始行号
                var cellValue = _dataGrid.Rows[e.RowIndex].Cells[0].Value;
                rowIndex = cellValue != null ? Convert.ToInt32(cellValue) : e.RowIndex + 1 + _frozenRows;
            }
            else
            {
                // 默认计算
                rowIndex = e.RowIndex + 1 + _frozenRows;
            }

            var rowHeaderBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, _dataGrid.RowHeadersWidth, e.RowBounds.Height);

            using (var brush = new SolidBrush(_dataGrid.RowHeadersDefaultCellStyle.ForeColor))
            {
                var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                e.Graphics.DrawString(rowIndex.ToString(), _dataGrid.Font, brush, rowHeaderBounds, format);
            }
        }

        private void OnFrozenRowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            // 绘制冻结表格的行号（1到frozenRows）
            var rowIndex = e.RowIndex + 1; // 冻结行从1开始
            var rowHeaderBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, _dataGridHeader.RowHeadersWidth, e.RowBounds.Height);

            using (var brush = new SolidBrush(_dataGridHeader.RowHeadersDefaultCellStyle.ForeColor))
            {
                var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                e.Graphics.DrawString(rowIndex.ToString(), _dataGridHeader.Font, brush, rowHeaderBounds, format);
            }
        }

        /// <summary>将列索引转换为Excel样式的列名（0->A, 1->B, 25->Z, 26->AA）</summary>
        private static string GetExcelColumnName(int columnIndex)
        {
            var result = new StringBuilder();
            int index = columnIndex;

            do
            {
                result.Insert(0, (char)('A' + (index % 26)));
                index = index / 26 - 1;
            } while (index >= 0);

            return result.ToString();
        }

        /// <summary>重建高亮行集合 — 扫描数据表找出关联列匹配的行，并自动滚动到首行（支持多值 OR 匹配）</summary>
        private void RebuildHighlightRows()
        {
            _highlightRows.Clear();
            if (_highlightColumnIndex < 0 || _highlightValues.Count == 0 || _dataTable.Columns.Count == 0)
            {
                _highlightsReady = false;
                return;
            }
            for (int i = _frozenRows; i < _dataTable.Rows.Count; i++)
            {
                if (_highlightColumnIndex < _dataTable.Columns.Count)
                {
                    var val = _dataTable.Rows[i][_highlightColumnIndex]?.ToString()?.Trim() ?? "";
                    if (_highlightValues.Any(v => val.Equals(v, StringComparison.OrdinalIgnoreCase)))
                        _highlightRows.Add(i);
                }
            }
            _highlightsReady = true;
            // 自动填入高级过滤：列名=值 进行精准匹配（支持多值 OR）
            if (_highlightRows.Count > 0)
            {
                string colRef = !string.IsNullOrEmpty(_highlightColumnName)
                    ? _highlightColumnName : $"Col{_highlightColumnIndex}";
                string filterValue = _highlightValues.Count == 1
                    ? $"{colRef}='{_highlightValues[0]}'"
                    : string.Join(" or ", _highlightValues.Select(v => $"{colRef}='{v}'"));
                _txtAdvancedFilter.Text = filterValue;
            }
        }

        /// <summary>使用ExcelDataReader读取Excel文件(.xls/.xlsx)为DataTable</summary>
        private static DataTable? ReadExcelToDataTable(string filePath)
        {
            try
            {
                // 注册编码提供者（ExcelDataReader 需要）
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
                    return null;

                var sheet = dataSet.Tables[0];
                var dataTable = new DataTable();

                int maxCols = sheet.Columns.Count;
                for (int i = 0; i < maxCols; i++)
                {
                    var columnName = ConvertToColumnLetter(i);
                    dataTable.Columns.Add(columnName, typeof(string));
                }

                // 读取行数据
                for (int r = 0; r < sheet.Rows.Count; r++)
                {
                    var newRow = dataTable.NewRow();
                    for (int c = 0; c < maxCols; c++)
                    {
                        newRow[c] = sheet.Rows[r][c]?.ToString() ?? "";
                    }
                    dataTable.Rows.Add(newRow);
                }

                return dataTable;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>将数字索引转换为Excel列字母（A, B, ..., Z, AA, AB, ...）</summary>
        private static string ConvertToColumnLetter(int index)
        {
            var result = new StringBuilder();
            int num = index;
            do
            {
                result.Insert(0, (char)('A' + (num % 26)));
                num = num / 26 - 1;
            } while (num >= 0);
            return result.ToString();
        }

        /// <summary>获取ExcelDataReader单元格的字符串值（已弃用，保留以保持兼容）</summary>
        private static string GetCellStringValue(object cell)
        {
            return cell?.ToString() ?? "";
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

    /// <summary>Excel内容查看控件</summary>
    public class ExcelViewerControl : UserControl
    {
        private DataGridView _dataGridHeader = new();
        private DataGridView _dataGrid = new();
        private SplitContainer _pnlGridContainer = new();
        private Button _btnSelectColumns = new();
        private TextBox _txtContentFilter = new();
        private TextBox _txtAdvancedFilter = new();
        private Label _lblInfo = new();
        private Label _lblCellPos = new();             // 当前编辑单元格行列提示
        private Panel _pnlTop = new();
        private Panel _pnlLoading = new();             // 加载遮罩层
        private DataTable _dataTable = new();
        private DataTable _headerTable = new();
        private string _filePath = "";
        /// <summary>当前文件完整路径</summary>
        public string FilePath => _filePath;
        private int _frozenRows = 0;
        private bool _frozenRowsOverridden = false;     // 是否通过构造函数强制指定了冻结行数
        private int _headerRowIndex = -1;               // 表头行索引（-1=未设置，0-based，必须 < frozenRows）
        private int _paneIndex = -1;                    // 所属窗格索引（-1表示独立打开，使用全局设置）
        private Dictionary<string, string> _fileRemarks = new(); // 文件备注映射（key: 相对路径）
        private List<int> _visibleColumnIndices = new();
        private bool[] _columnCheckStates = Array.Empty<bool>();
        private List<ColumnLinkConfig> _columnLinks = new();
        private List<(string display, string fullPath)> _fileEntries = new();

        /// <summary>请求关联打开另一张表事件</summary>
        public event EventHandler? LinkTableRequested;

        // 关联表跳转高亮
        private int _highlightColumnIndex = -1;
        private string _highlightValue = "";
        private string _highlightColumnName = "";
        private HashSet<int> _highlightRows = new();
        private bool _highlightsReady = false;
        private List<string> _highlightValues = new(); // 多值高亮（支持 OR 匹配）

        public void SetHighlight(int colIndex, string value, string colName = "", ColumnLinkConfig? linkConfig = null)
        {
            _highlightColumnIndex = colIndex;
            _highlightValue = value ?? "";
            _highlightColumnName = colName ?? "";
            // 根据 linkConfig 提取多值用于 OR 匹配
            _highlightValues.Clear();
            if (linkConfig != null && linkConfig.ValueExtractionEnabled && !string.IsNullOrEmpty(value))
            {
                _highlightValues = linkConfig.ExtractValues(value);
                // 兜底：如果提取后只有一个值，但原始值中包含多值分隔符，尝试手动按多值分隔符拆分
                if (_highlightValues.Count == 1 && !string.IsNullOrEmpty(linkConfig.MultiValueDelimiter) && value.Contains(linkConfig.MultiValueDelimiter))
                {
                    var groups = value.Split(new[] { linkConfig.MultiValueDelimiter }, StringSplitOptions.RemoveEmptyEntries);
                    var fallback = new List<string>();
                    foreach (var g in groups)
                    {
                        var trimmed = g.Trim();
                        if (!string.IsNullOrEmpty(trimmed))
                        {
                            // 对每组也用 ExtractDelimiter 提取
                            if (!string.IsNullOrEmpty(linkConfig.ExtractDelimiter))
                            {
                                var parts = trimmed.Split(new[] { linkConfig.ExtractDelimiter }, StringSplitOptions.None);
                                if (linkConfig.ExtractPartIndex >= 0 && linkConfig.ExtractPartIndex < parts.Length)
                                    fallback.Add(parts[linkConfig.ExtractPartIndex].Trim());
                                else
                                    fallback.Add(trimmed);
                            }
                            else
                            {
                                fallback.Add(trimmed);
                            }
                        }
                    }
                    if (fallback.Count > 1)
                        _highlightValues = fallback.Distinct().ToList();
                }
            }
            else if (!string.IsNullOrEmpty(value))
            {
                _highlightValues = new List<string> { value };
            }
        }

        // 配色（浅色天蓝色调）
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

        public ExcelViewerControl(string filePath, int paneIndex = -1, string? rootDir = null, int frozenRows = -1, int headerRowIndex = -1)
        {
            _filePath = filePath;
            _paneIndex = paneIndex;
            _rootDir = rootDir;
            if (frozenRows >= 0) { _frozenRows = frozenRows; _frozenRowsOverridden = true; }
            if (headerRowIndex >= 0) _headerRowIndex = headerRowIndex;
            InitializeComponent();
            LoadExcelAsync();
        }

        private string? _rootDir;
        /// <summary>当前浏览根目录（用于文件夹级设置存储）</summary>
        private string FrozenDir => !string.IsNullOrEmpty(_rootDir) ? _rootDir : Path.GetDirectoryName(_filePath) ?? "";

        private void InitializeComponent()
        {
            SuspendLayout();
            BackColor = _colorSurface;

            // 顶部工具栏 - 使用 FlowLayoutPanel 自动排列
            var flowTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8, 5, 8, 5),
                AutoSize = true
            };

            // 冻结行数选择
            var lblFrozenRows = new Label
            {
                Text = "冻结行:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            var cboFrozenRows = new ComboBox
            {
                Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 10, 0)
            };
            cboFrozenRows.Items.AddRange(new object[] { "0", "1", "2", "3", "4", "5", "6", "7", "8" });
            if (!_frozenRowsOverridden)
                _frozenRows = !string.IsNullOrEmpty(FrozenDir) ? CsvSettings.GetExcelFrozenRowsByDirectory(FrozenDir) : CsvSettings.ExcelDefaultFrozenRows;
            cboFrozenRows.SelectedIndex = _frozenRows;
            cboFrozenRows.SelectedIndexChanged += (s, e) =>
            {
                _frozenRows = cboFrozenRows.SelectedIndex;
                if (!string.IsNullOrEmpty(FrozenDir))
                    CsvSettings.SetExcelFrozenRowsByDirectory(FrozenDir, _frozenRows);
                else
                    CsvSettings.ExcelDefaultFrozenRows = _frozenRows;
                // 联动：如果表头行索引超出新的冻结行范围，自动调整
                if (_headerRowIndex >= _frozenRows)
                    _headerRowIndex = _frozenRows > 0 ? 0 : -1;
                UpdateHeaderRowCombo();
                if (!string.IsNullOrEmpty(FrozenDir))
                    CsvSettings.SetExcelHeaderRowIndexByDirectory(FrozenDir, _headerRowIndex);
                CsvSettings.Save();
                RefreshGridLayout();
            };

            // 表头行选择（从固定行中选择哪一行是字段名列）
            var lblHeaderRow = new Label
            {
                Text = "表头行:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            var cboHeaderRow = new ComboBox
            {
                Size = new Size(70, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 10, 0)
            };
            // 从设置读取表头行索引
            if (!_frozenRowsOverridden && !string.IsNullOrEmpty(FrozenDir))
                _headerRowIndex = CsvSettings.GetExcelHeaderRowIndexByDirectory(FrozenDir);
            PopulateHeaderRowCombo(cboHeaderRow);
            cboHeaderRow.SelectedIndexChanged += (s, e) =>
            {
                _headerRowIndex = cboHeaderRow.SelectedIndex >= 0 ? cboHeaderRow.SelectedIndex : -1;
                if (!string.IsNullOrEmpty(FrozenDir))
                    CsvSettings.SetExcelHeaderRowIndexByDirectory(FrozenDir, _headerRowIndex);
                CsvSettings.Save();
                RefreshGridLayout();
            };
            this.Tag = cboHeaderRow; // 保持引用

            // 表头过滤 - 列选择按钮
            var lblHeaderFilter = new Label
            {
                Text = "列选择:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            _btnSelectColumns = new Button
            {
                Size = new Size(120, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                Text = "选择列 (全部)",
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = _colorHoverBg,
                ForeColor = _colorLightText,
                FlatAppearance = { BorderSize = 1, BorderColor = _colorBorder },
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 2, 10, 0)
            };
            _btnSelectColumns.Click += OnSelectColumnsClick;

            // 信息标签
            _lblInfo = new Label
            {
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 0, 0)
            };

            flowTop.Controls.AddRange(new Control[] { lblFrozenRows, cboFrozenRows, lblHeaderRow, cboHeaderRow, lblHeaderFilter, _btnSelectColumns });

            // 重新加载按钮（保留过滤状态刷新文件内容）
            var btnReload = new Button
            {
                Text = "重新加载",
                Size = new Size(75, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 2, 10, 0)
            };
            btnReload.Click += (s, e) => ReloadContentAsync();
            flowTop.Controls.Add(btnReload);

            // 关联表按钮（已移除）

            // 行列位置标签
            var lblCellPos = new Label
            {
                Text = "",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                ForeColor = _colorPrimary,
                Margin = new Padding(10, 5, 10, 0)
            };
            flowTop.Controls.Add(lblCellPos);
            _lblCellPos = lblCellPos;

            // 信息标签放到最后
            flowTop.Controls.Add(_lblInfo);

            // ====== 第二行: 分隔线 ======
            var separatorLine = new Panel
            {
                Height = 1,
                Dock = DockStyle.Top,
                BackColor = Color.FromArgb(200, 225, 245),
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            // ====== 第三行: 内容过滤 + 高级表达式过滤框 并列 ======
            var lblContentFilter = new Label
            {
                Text = "内容过滤:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightText,
                Margin = new Padding(3, 5, 3, 0)
            };

            var txtContentFilter = new TextBox
            {
                Size = new Size(200, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 10, 0)
            };
            txtContentFilter.TextChanged += (s, e) => ApplyFilters();
            txtContentFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    txtContentFilter.Text = "";
                    e.Handled = true;
                }
            };

            // ====== 第三行: 内容过滤 ======
            var flowContent = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Height = 30,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8, 1, 8, 0),
                AutoSize = true
            };
            flowContent.Controls.AddRange(new Control[] { lblContentFilter, txtContentFilter });

            // ====== 第四行: 高级过滤 ======
            var flowAdv = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Height = 30,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8, 3, 8, 5),
                AutoSize = true
            };

            var lblAdvFilter = new Label
            {
                Text = "高级过滤:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightText,
                Margin = new Padding(3, 5, 3, 0)
            };

            var txtAdvancedFilter = new TextBox
            {
                Size = new Size(200, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 10, 0),
                BackColor = Color.FromArgb(242, 249, 255)
            };
            txtAdvancedFilter.TextChanged += (s, e) => ApplyFilters();
            txtAdvancedFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    txtAdvancedFilter.Text = "";
                    e.Handled = true;
                }
            };

            var lblAdvHint = new Label
            {
                Text = "如: id=123 or name like test",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = _colorPrimary,
                Margin = new Padding(3, 5, 10, 0)
            };

            flowAdv.Controls.AddRange(new Control[] { lblAdvFilter, txtAdvancedFilter, lblAdvHint });

            // 组装 topTable 三行布局
            var topTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                BackColor = _colorSurface,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(0),
                Margin = new Padding(0),
                GrowStyle = TableLayoutPanelGrowStyle.AddRows
            };
            topTable.RowStyles.Clear();
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 行0: 原有控件
            topTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 1)); // 行1: 分隔线
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 行2: 内容过滤
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 行3: 高级过滤
            topTable.Controls.Add(flowTop, 0, 0);
            topTable.Controls.Add(separatorLine, 0, 1);
            topTable.Controls.Add(flowContent, 0, 2);
            topTable.Controls.Add(flowAdv, 0, 3);

            _pnlTop = topTable;
            _txtContentFilter = txtContentFilter;
            _txtAdvancedFilter = txtAdvancedFilter;

            // 过滤框宽度自适应：宽度 = 控件宽度 - 280px（给右侧提示文本留空间）
            topTable.Resize += (s, e) =>
            {
                int w = topTable.ClientSize.Width - 280;
                if (w < 100) w = 100;
                txtContentFilter.Width = w;
                txtAdvancedFilter.Width = w;
            };

            // 表格容器 - 使用SplitContainer来精确控制布局
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 28,
                FixedPanel = FixedPanel.Panel1,
                IsSplitterFixed = true,
                BackColor = _colorSurface,
                SplitterWidth = 1,
                BorderStyle = BorderStyle.FixedSingle
            };

            // 冻结的表头表格（显示冻结的行）
            _dataGridHeader = CreateDataGridView();
            _dataGridHeader.Dock = DockStyle.Fill;
            _dataGridHeader.ScrollBars = ScrollBars.None;
            _dataGridHeader.RowHeadersVisible = true;
            _dataGridHeader.RowHeadersWidth = 50;
            _dataGridHeader.ColumnHeadersVisible = true;
            _dataGridHeader.ReadOnly = true;
            _dataGridHeader.AllowUserToResizeColumns = false;
            _dataGridHeader.AllowUserToResizeRows = false;
            _dataGridHeader.RowPostPaint += OnFrozenRowPostPaint;
            _dataGridHeader.DefaultCellStyle.SelectionBackColor = _dataGridHeader.DefaultCellStyle.BackColor;
            _dataGridHeader.DefaultCellStyle.SelectionForeColor = _dataGridHeader.DefaultCellStyle.ForeColor;
            _dataGridHeader.DefaultCellStyle.BackColor = Color.FromArgb(245, 250, 255);
            _dataGridHeader.TabStop = false;
            _dataGridHeader.CellMouseDown += (s, e) =>
            {
                // 鼠标点击时同步重置，阻止滚动偏移发生
                if (e.RowIndex >= 0 && _dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;

                // 右键菜单：复制 + 编辑
                if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    var cellValue = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    var ctx = new ContextMenuStrip();
                    ctx.Font = new Font("Microsoft YaHei UI", 9f);

                    var copyItem = new ToolStripMenuItem("复制");
                    copyItem.Click += (s2, args) =>
                    {
                        if (!string.IsNullOrEmpty(cellValue))
                            Clipboard.SetText(cellValue);
                    };
                    ctx.Items.Add(copyItem);

                    var editItem = new ToolStripMenuItem("编辑");
                    editItem.Click += (s2, args) =>
                    {
                        var cell = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        _editingFromFrozen = true;
                        EditCell(cell, _dataGridHeader);
                    };
                    ctx.Items.Add(editItem);

                    ctx.Show(_dataGridHeader, _dataGridHeader.PointToClient(Cursor.Position));
                }
            };
            _dataGridHeader.CurrentCellChanged += (s, e) =>
            {
                // 键盘导航等情况的兜底处理
                if (_dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;
            };
            _dataGridHeader.CellDoubleClick += OnFrozenCellDoubleClick;
            _dataGridHeader.CellClick += OnCellClick; // 双击编辑

            // 数据表格（主表格）
            _dataGrid = CreateDataGridView();
            _dataGrid.Dock = DockStyle.Fill;
            _dataGrid.RowHeadersWidth = 50;
            _dataGrid.RowPostPaint += OnRowPostPaint;
            _dataGrid.ColumnHeadersVisible = false;
            _dataGrid.CellMouseDown += OnGridCellMouseDown;

            // 同步水平滚动
            _dataGrid.Scroll += (s, e) =>
            {
                _dataGridHeader.HorizontalScrollingOffset = _dataGrid.HorizontalScrollingOffset;
            };

            // 同步列宽调整
            _dataGrid.ColumnWidthChanged += (s, e) =>
            {
                int dataGridColumnIndex = e.Column.Index;
                int headerColumnIndex = dataGridColumnIndex;

                if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__" && dataGridColumnIndex > 0)
                {
                    headerColumnIndex = dataGridColumnIndex - 1;
                }

                if (headerColumnIndex >= 0 && headerColumnIndex < _dataGridHeader.Columns.Count)
                {
                    _dataGridHeader.Columns[headerColumnIndex].Width = _dataGrid.Columns[dataGridColumnIndex].Width;
                }
            };

            splitContainer.Panel1.Controls.Add(_dataGridHeader);
            splitContainer.Panel2.Controls.Add(_dataGrid);

            splitContainer.Panel1Collapsed = true;
            _pnlGridContainer = splitContainer;

            Controls.Add(_pnlGridContainer);
            Controls.Add(_pnlTop);

            ResumeLayout(false);
        }

        private DataGridView CreateDataGridView()
        {
            var dgv = new DataGridView
            {
                BackgroundColor = _colorSurface,
                BorderStyle = BorderStyle.None,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 28,
                RowHeadersVisible = true,
                RowHeadersWidth = 50,
                RowHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    BackColor = _colorRowNumBg,
                    ForeColor = _colorLightTextSecondary,
                    Font = new Font("Microsoft YaHei UI", 8.5f)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _colorAltRow
                },
                GridColor = _colorGridLine,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                EnableHeadersVisualStyles = false,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _colorHeaderBg,
                    ForeColor = _colorLightText,
                    Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    SelectionBackColor = _colorHeaderBg,
                    SelectionForeColor = _colorLightText
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _colorSurface,
                    ForeColor = _colorLightText,
                    Font = new Font("Microsoft YaHei UI", 9f),
                    SelectionBackColor = _colorSelectedBg,
                    SelectionForeColor = Color.FromArgb(20, 60, 100),
                    Padding = new Padding(2, 0, 2, 0)
                },
                RowTemplate = new DataGridViewRow { Height = 24 }
            };
            dgv.ColumnHeaderMouseClick += OnColumnHeaderClick;
            dgv.CellFormatting += OnCellFormatting;
            dgv.CellDoubleClick += OnCellDoubleClick;
            dgv.CellClick += OnCellClick;
            dgv.CellToolTipTextNeeded += OnCellToolTipTextNeeded;
            dgv.SelectionChanged += (s, e) => dgv.Invalidate();
            return dgv;
        }

        // ==================== 点击行列提示 + 双击弹窗编辑 ====================
        private bool _editingFromFrozen = false;

        private void OnCellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var dgv = (DataGridView)sender!;

            int rowOff = (dgv == _dataGridHeader) ? 0 : _frozenRows;
            int colIdx = e.ColumnIndex;
            if (dgv != _dataGridHeader && dgv.Columns.Count > 0 && dgv.Columns[0].Name == "__RowNum__")
                colIdx = Math.Max(0, colIdx - 1);

            string colName = colIdx >= 0 && colIdx < 26
                ? ((char)('A' + colIdx)).ToString()
                : $"Col{colIdx}";
            _lblCellPos.Text = $"行 {e.RowIndex + 1 + rowOff}  列 {colName}({colIdx})";
        }

        private void EditCell(DataGridViewCell cell, DataGridView dgv)
        {
            var rowIndex = cell.RowIndex;
            var columnIndex = cell.ColumnIndex;
            var currentValue = cell.Value?.ToString() ?? "";
            var (originalRowIndex, originalColumnIndex) = GetOriginalIndices(rowIndex, columnIndex);

            // 获取列名
            string colName = "";
            if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && originalColumnIndex >= 0 && originalColumnIndex < _dataTable.Columns.Count)
                colName = _dataTable.Rows[_headerRowIndex][originalColumnIndex]?.ToString() ?? "";

            // Excel文件：只读模式，保存按钮置灰并显示提示
            var dialog = new CellEditDialog(currentValue, originalRowIndex, originalColumnIndex, _filePath, colName,
                onSave: null, readOnly: true, readOnlyMessage: "不支持Excel文件修改");
            dialog.Show(this);
        }

        private void OnCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var dgv = (DataGridView)sender!;
            if (dgv == _dataGridHeader) return;
            var cell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
            _editingFromFrozen = false;
            EditCell(cell, dgv);
        }

        private void OnFrozenCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var cell = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex];
            _editingFromFrozen = true;
            EditCell(cell, _dataGridHeader);
        }

        /// <summary>获取原始数据表中的行列索引</summary>
        private (int rowIndex, int columnIndex) GetOriginalIndices(int displayRowIndex, int displayColumnIndex)
        {
            // 如果来自冻结表格，行号直接就是原始行号（0-based）
            if (_editingFromFrozen)
            {
                int originalColumnIndex = displayColumnIndex;
                // 列索引转换（考虑可见列过滤）
                if (_visibleColumnIndices.Count > 0 && displayColumnIndex < _dataGridHeader.Columns.Count)
                {
                    if (displayColumnIndex < _visibleColumnIndices.Count)
                    {
                        originalColumnIndex = _visibleColumnIndices[displayColumnIndex];
                    }
                }
                return (displayRowIndex, originalColumnIndex);
            }

            // 如果使用了过滤显示，需要转换索引
            if (_dataGrid.DataSource is DataTable filteredTable && filteredTable.Columns.Count > 0)
            {
                // 获取原始行号（从隐藏列）
                int originalRowIndex = displayRowIndex;
                if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__" && displayRowIndex >= 0)
                {
                    var cellValue = _dataGrid.Rows[displayRowIndex].Cells[0].Value;
                    if (cellValue != null)
                    {
                        originalRowIndex = Convert.ToInt32(cellValue) - 1; // 转换为0-based索引
                    }
                }

                // 列索引转换
                int originalColumnIndex = displayColumnIndex;
                if (_visibleColumnIndices.Count > 0 && displayColumnIndex < _dataGrid.Columns.Count)
                {
                    // 考虑隐藏列（__RowNum__）
                    var actualDisplayIndex = displayColumnIndex;
                    if (_dataGrid.Columns[0].Name == "__RowNum__")
                    {
                        actualDisplayIndex = displayColumnIndex - 1;
                    }

                    if (actualDisplayIndex >= 0 && actualDisplayIndex < _visibleColumnIndices.Count)
                    {
                        originalColumnIndex = _visibleColumnIndices[actualDisplayIndex];
                    }
                }

                return (originalRowIndex, originalColumnIndex);
            }

            // 无过滤时直接返回
            return (displayRowIndex + _frozenRows, displayColumnIndex);
        }

        /// <summary>重新加载Excel文件内容，保留过滤条件和列选择状态</summary>
        private async void ReloadContentAsync()
        {
            if (!System.IO.File.Exists(_filePath)) return;

            var savedFilter = _txtContentFilter.Text;
            var savedAdvFilter = _txtAdvancedFilter.Text;
            var savedColumnStates = _columnCheckStates;
            var savedFrozenRows = _frozenRows;

            try
            {
                if (!this.IsHandleCreated)
                {
                    var tcs = new TaskCompletionSource<bool>();
                    EventHandler? handler = null;
                    handler = (s, e) =>
                    {
                        this.HandleCreated -= handler;
                        tcs.TrySetResult(true);
                    };
                    this.HandleCreated += handler;
                    await tcs.Task;
                }

                await Task.Run(() => ReloadExcelFile());
            }
            catch (Exception ex)
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(() =>
                    {
                        MessageBox.Show(this, $"重新加载Excel文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
                _columnCheckStates = savedColumnStates;
                _frozenRows = savedFrozenRows;
                _txtContentFilter.Text = savedFilter;
                _txtAdvancedFilter.Text = savedAdvFilter;
                UpdateColumnSelectorButton();
                return;
            }

            if (savedColumnStates.Length == _dataTable.Columns.Count)
            {
                _columnCheckStates = savedColumnStates;
            }
            else
            {
                _columnCheckStates = new bool[_dataTable.Columns.Count];
                for (int i = 0; i < _columnCheckStates.Length; i++)
                    _columnCheckStates[i] = true;
            }
            _frozenRows = savedFrozenRows;
            _txtContentFilter.Text = savedFilter;
            _txtAdvancedFilter.Text = savedAdvFilter;
            UpdateColumnSelectorButton();
            RefreshGridLayout();
        }

        /// <summary>重新读取Excel文件内容（后台线程执行）</summary>
        private void ReloadExcelFile()
        {
            if (!System.IO.File.Exists(_filePath)) return;

            var dataTable = ReadExcelToDataTable(_filePath);
            if (dataTable == null) return;

            this.Invoke(() =>
            {
                _dataTable = dataTable;
                _headerTable = dataTable.Clone();
            });
        }

        private async void LoadExcelAsync()
        {
            try
            {
                if (!this.IsHandleCreated)
                {
                    var tcs = new TaskCompletionSource<bool>();
                    EventHandler? handler = null;
                    handler = (s, e) =>
                    {
                        this.HandleCreated -= handler;
                        tcs.TrySetResult(true);
                    };
                    this.HandleCreated += handler;
                    await tcs.Task;
                }

                // 显示加载遮罩
                var loadingStart = DateTime.UtcNow;
                ShowLoadingOverlay();
                // 短暂延迟确保 loading 遮罩渲染到屏幕上
                await Task.Delay(50);

                await Task.Run(() => LoadExcelFile());

                // 确保 loading 至少显示了 200ms，避免一闪而过
                var elapsed = (DateTime.UtcNow - loadingStart).TotalMilliseconds;
                if (elapsed < 200)
                    await Task.Delay((int)(200 - elapsed));

                // 隐藏加载遮罩
                HideLoadingOverlay();
            }
            catch (Exception ex)
            {
                HideLoadingOverlay();
                if (this.IsHandleCreated)
                {
                    this.Invoke(() =>
                    {
                        MessageBox.Show(this, $"加载Excel文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            }
        }

        private void ShowLoadingOverlay()
        {
            if (!this.IsHandleCreated) return;
            this.Invoke(() =>
            {
                if (_pnlLoading?.IsDisposed != false || _pnlLoading.Parent != null) return;
                _pnlLoading = new Panel
                {
                    BackColor = Color.FromArgb(180, 255, 255, 255),
                    Dock = DockStyle.Fill,
                    Visible = true
                };
                var lbl = new Label
                {
                    Text = "正在加载数据，请稍候...",
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Microsoft YaHei UI", 14f, FontStyle.Bold),
                    ForeColor = Color.FromArgb(30, 136, 229),
                    Dock = DockStyle.Fill
                };
                _pnlLoading.Controls.Add(lbl);
                this.Controls.Add(_pnlLoading);
                _pnlLoading.BringToFront();
                Application.DoEvents();
            });
        }

        private void HideLoadingOverlay()
        {
            if (!this.IsHandleCreated) return;
            this.Invoke(() =>
            {
                if (_pnlLoading != null && !_pnlLoading.IsDisposed && _pnlLoading.Parent == this)
                {
                    this.Controls.Remove(_pnlLoading);
                    _pnlLoading.Dispose();
                    _pnlLoading = null;
                }
            });
        }

        private void LoadExcelFile()
        {
            if (!System.IO.File.Exists(_filePath))
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, "Excel文件不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error));
                return;
            }

            var dataTable = ReadExcelToDataTable(_filePath);
            if (dataTable == null)
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, "Excel文件没有工作表", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning));
                return;
            }

            this.Invoke(() =>
            {
                _dataTable = dataTable;
                _headerTable = dataTable.Clone();

                // 初始化列选中状态（默认全选，第一列始终选中）
                _columnCheckStates = new bool[_dataTable.Columns.Count];
                for (int i = 0; i < _columnCheckStates.Length; i++)
                    _columnCheckStates[i] = true;

                _columnLinks = CsvSettings.GetColumnLinksByDirectory(FrozenDir);
                ColumnLinkChecker.BuildLinkLookups(_columnLinks, FrozenDir, _filePath);

                UpdateColumnSelectorButton();
                RebuildHighlightRows();

                _lblInfo.Text = $"行数: {_dataTable.Rows.Count} | 列数: {_dataTable.Columns.Count}";

                if (_frozenRows > 0)
                {
                    RefreshGridLayout();
                }
                else
                {
                    _dataGrid.DataSource = _dataTable;

                    for (int i = 0; i < _dataGrid.Columns.Count; i++)
                    {
                        var col = _dataGrid.Columns[i];
                        col.Width = 120;
                        col.SortMode = DataGridViewColumnSortMode.Programmatic;
                        col.HeaderText = GetExcelColumnName(i);
                    }
                }
            });
        }

        /// <summary>填充表头行下拉框选项（根据当前冻结行数动态生成）</summary>
        private void PopulateHeaderRowCombo(ComboBox cbo)
        {
            cbo.Items.Clear();
            if (_frozenRows > 0)
            {
                for (int i = 0; i < _frozenRows; i++)
                    cbo.Items.Add($"第{i + 1}行");
                // 设置选中项
                if (_headerRowIndex >= 0 && _headerRowIndex < _frozenRows)
                    cbo.SelectedIndex = _headerRowIndex;
                else
                    cbo.SelectedIndex = -1;
            }
        }

        /// <summary>更新表头行下拉框（联动调用，供冻结行变更时使用）</summary>
        private void UpdateHeaderRowCombo()
        {
            // 通过 Tag 获取 ComboBox 引用
            if (Tag is ComboBox cboHdr)
            {
                PopulateHeaderRowCombo(cboHdr);
            }
        }

        /// <summary>获取表头行的列名列表（用于过滤、关联等场景）</summary>
        public List<string> GetHeaderColumnNames()
        {
            var names = new List<string>();
            if (_headerRowIndex < 0 || _headerRowIndex >= _dataTable.Rows.Count)
                return names;
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                names.Add(_dataTable.Rows[_headerRowIndex][i]?.ToString() ?? $"Col{i}");
            }
            return names;
        }

        /// <summary>根据字段名查找对应的原始列索引</summary>
        public int FindColumnIndexByHeaderName(string headerName)
        {
            if (_headerRowIndex < 0 || _headerRowIndex >= _dataTable.Rows.Count) return -1;
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                var val = _dataTable.Rows[_headerRowIndex][i]?.ToString() ?? "";
                if (val.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        private void RefreshGridLayout()
        {
            if (_dataTable.Columns.Count == 0)
            {
                return;
            }

            if (_frozenRows > 0 && _dataTable.Rows.Count > 0)
            {
                _visibleColumnIndices = GetVisibleColumnIndices();

                _headerTable = new DataTable();
                foreach (var colIndex in _visibleColumnIndices)
                {
                    var sourceCol = _dataTable.Columns[colIndex];
                    _headerTable.Columns.Add(sourceCol.ColumnName, sourceCol.DataType);
                }

                for (int i = 0; i < Math.Min(_frozenRows, _dataTable.Rows.Count); i++)
                {
                    var sourceRow = _dataTable.Rows[i];
                    var targetRow = _headerTable.NewRow();
                    for (int j = 0; j < _visibleColumnIndices.Count; j++)
                    {
                        targetRow[j] = sourceRow[_visibleColumnIndices[j]];
                    }
                    _headerTable.Rows.Add(targetRow);
                }

                _pnlGridContainer.Panel1Collapsed = false;
                _pnlGridContainer.SplitterDistance = _frozenRows * 24 + 28;

                _dataGridHeader.DataSource = _headerTable;

                for (int i = 0; i < _dataGridHeader.Columns.Count; i++)
                {
                    var colName = _headerTable.Columns[i].ColumnName;
                    var dgvCol = _dataGridHeader.Columns[colName];
                    if (dgvCol != null)
                    {
                        dgvCol.DisplayIndex = i;
                        dgvCol.Width = 120;
                        dgvCol.HeaderText = GetExcelColumnName(i);
                        dgvCol.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                }

                var contentFilter = _txtContentFilter.Text.Trim();
                var visibleRows = GetVisibleRows(contentFilter);
                RefreshFilteredDisplay(visibleRows);
            }
            else
            {
                _pnlGridContainer.Panel1Collapsed = true;
                _dataGrid.DataSource = _dataTable;

                for (int i = 0; i < _dataGrid.Columns.Count; i++)
                {
                    _dataGrid.Columns[i].Width = 120;
                    _dataGrid.Columns[i].HeaderText = GetExcelColumnName(i);
                    _dataGrid.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
                }
            }
        }

        private List<int> GetVisibleColumnIndices()
        {
            var visibleIndices = new List<int>();

            if (_dataTable.Columns.Count > 0)
            {
                visibleIndices.Add(0);
            }

            for (int i = 1; i < _dataTable.Columns.Count; i++)
            {
                if (i < _columnCheckStates.Length && _columnCheckStates[i])
                {
                    visibleIndices.Add(i);
                }
            }

            return visibleIndices;
        }

        private void OnSelectColumnsClick(object? sender, EventArgs e)
        {
            if (_dataTable.Columns.Count == 0) return;

            EnsureFileEntriesLoaded();

            using var dialog = new ColumnSelectorDialog(_dataTable, _columnCheckStates, _frozenRows, 0,
                _rootDir ?? "", _filePath, _fileEntries, _columnLinks, _headerRowIndex);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _columnCheckStates = dialog.GetCheckedStates();
                _columnLinks = dialog.GetColumnLinks();
                CsvSettings.SetColumnLinksByDirectory(FrozenDir, _columnLinks);
                ColumnLinkChecker.BuildLinkLookups(_columnLinks, FrozenDir, _filePath);
                UpdateColumnSelectorButton();
                RefreshGridLayout();
            }
        }

        /// <summary>确保文件条目列表已加载（从 _rootDir 扫描）</summary>
        private void EnsureFileEntriesLoaded()
        {
            if (_fileEntries.Count > 0) return;
            var root = FrozenDir;
            if (string.IsNullOrEmpty(root) || !Directory.Exists(root)) return;
            _fileEntries.Clear();
            try
            {
                var showHidden = CsvSettings.ShowHiddenFiles;
                foreach (var f in Directory.GetFiles(root, "*.csv", SearchOption.AllDirectories))
                {
                    if (!showHidden)
                    {
                        try { if ((File.GetAttributes(f) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                    }
                    var rel = f.StartsWith(root, StringComparison.OrdinalIgnoreCase)
                        ? f.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar) : Path.GetFileName(f);
                    _fileEntries.Add((rel, f));
                }
                foreach (var ext in new[] { "*.xls", "*.xlsx" })
                {
                    foreach (var f in Directory.GetFiles(root, ext, SearchOption.AllDirectories))
                    {
                        if (!showHidden)
                        {
                            try { if ((File.GetAttributes(f) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                        }
                        var rel = f.StartsWith(root, StringComparison.OrdinalIgnoreCase)
                            ? f.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar) : Path.GetFileName(f);
                        _fileEntries.Add((rel, f));
                    }
                }
            }
            catch { }
        }

        private void UpdateColumnSelectorButton()
        {
            if (_dataTable.Columns.Count == 0) return;
            int checkedCount = 0;
            for (int i = 0; i < _columnCheckStates.Length; i++)
                if (_columnCheckStates[i]) checkedCount++;

            int dataColumns = Math.Max(0, _columnCheckStates.Length - 1);
            int checkedData = Math.Max(0, checkedCount - 1);
            _btnSelectColumns.Text = $"选择列 ({checkedData}/{dataColumns})";
        }

        private List<DataRow> GetVisibleRows(string contentFilter)
        {
            var visibleRows = new List<DataRow>();
            var advFilter = _txtAdvancedFilter.Text.Trim();

            for (int i = _frozenRows; i < _dataTable.Rows.Count; i++)
            {
                var row = _dataTable.Rows[i];

                // 1) 简单内容过滤
                bool contentMatch = string.IsNullOrEmpty(contentFilter);
                if (!contentMatch)
                {
                    var processedFilter = contentFilter.Replace("||", " | ").Replace("&&", " & ");

                    foreach (var colIndex in _visibleColumnIndices)
                    {
                        var cellValue = row[colIndex]?.ToString() ?? "";
                        if (MatchesFilter(cellValue, processedFilter))
                        {
                            contentMatch = true;
                            break;
                        }
                    }
                }

                // 2) 高级表达式过滤
                bool advMatch = string.IsNullOrEmpty(advFilter);
                if (!advMatch)
                {
                    advMatch = EvaluateExpression(row, advFilter, _visibleColumnIndices);
                }

                // AND 逻辑
                if (contentMatch && advMatch)
                {
                    visibleRows.Add(row);
                }
            }

            return visibleRows;
        }

        private void RefreshFilteredDisplay(List<DataRow> visibleRows)
        {
            var filteredTable = new DataTable();

            filteredTable.Columns.Add("__RowNum__", typeof(int));

            foreach (var colIndex in _visibleColumnIndices)
            {
                var sourceCol = _dataTable.Columns[colIndex];
                filteredTable.Columns.Add(sourceCol.ColumnName, sourceCol.DataType);
            }

            foreach (var sourceRow in visibleRows)
            {
                var newRow = filteredTable.NewRow();
                var originalRowIndex = _dataTable.Rows.IndexOf(sourceRow) + 1;
                newRow[0] = originalRowIndex;

                for (int i = 0; i < _visibleColumnIndices.Count; i++)
                {
                    newRow[i + 1] = sourceRow[_visibleColumnIndices[i]];
                }
                filteredTable.Rows.Add(newRow);
            }

            _dataGrid.DataSource = filteredTable;

            if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__")
            {
                _dataGrid.Columns[0].Visible = false;
            }

            for (int i = 1; i < _dataGrid.Columns.Count; i++)
            {
                var colName = filteredTable.Columns[i].ColumnName;
                var dgvCol = _dataGrid.Columns[colName];
                if (dgvCol != null)
                {
                    dgvCol.DisplayIndex = i;
                    dgvCol.Width = 120;
                    dgvCol.HeaderText = GetExcelColumnName(i - 1);
                    dgvCol.SortMode = DataGridViewColumnSortMode.Programmatic;
                }
            }

            _lblInfo.Text = $"显示 {_visibleColumnIndices.Count} 列, {visibleRows.Count} 行";
        }

        private void ApplyFilters()
        {
            RefreshGridLayout();
        }

        /// <summary>递归评估布尔表达式</summary>
        private bool EvaluateExpression(DataRow row, string expr, List<int> columnsToSearch)
        {
            expr = expr.Trim();
            if (string.IsNullOrEmpty(expr))
                return true;

            while (expr.StartsWith("(") && expr.EndsWith(")") && IsBalanced(expr))
            {
                expr = expr.Substring(1, expr.Length - 2).Trim();
            }

            int parenDepth = 0;
            int orPos = -1;
            int andPos = -1;

            for (int i = 0; i < expr.Length; i++)
            {
                if (expr[i] == '(')
                    parenDepth++;
                else if (expr[i] == ')')
                    parenDepth--;
                else if (parenDepth == 0)
                {
                    if (i + 2 < expr.Length &&
                        (expr[i] == 'o' || expr[i] == 'O') &&
                        (expr[i + 1] == 'r' || expr[i + 1] == 'R') &&
                        (i + 2 >= expr.Length || !char.IsLetterOrDigit(expr[i + 2])) &&
                        (i == 0 || !char.IsLetterOrDigit(expr[i - 1])))
                    {
                        orPos = i;
                        break;
                    }

                    if (i + 3 < expr.Length &&
                        (expr[i] == 'a' || expr[i] == 'A') &&
                        (expr[i + 1] == 'n' || expr[i + 1] == 'N') &&
                        (expr[i + 2] == 'd' || expr[i + 2] == 'D') &&
                        (i + 3 >= expr.Length || !char.IsLetterOrDigit(expr[i + 3])) &&
                        (i == 0 || !char.IsLetterOrDigit(expr[i - 1])))
                    {
                        if (andPos < 0)
                            andPos = i;
                        i += 3;
                    }
                }
            }

            if (orPos >= 0)
            {
                var left = expr.Substring(0, orPos).Trim();
                var right = expr.Substring(orPos + 2).Trim();
                return EvaluateExpression(row, left, columnsToSearch) ||
                       EvaluateExpression(row, right, columnsToSearch);
            }

            if (andPos >= 0)
            {
                var left = expr.Substring(0, andPos).Trim();
                var right = expr.Substring(andPos + 3).Trim();
                return EvaluateExpression(row, left, columnsToSearch) &&
                       EvaluateExpression(row, right, columnsToSearch);
            }

            return EvaluateCondition(row, expr, columnsToSearch);
        }

        private bool IsBalanced(string s)
        {
            int d = 0;
            foreach (char c in s)
            {
                if (c == '(') d++;
                else if (c == ')') d--;
                if (d < 0) return false;
            }
            return d == 0;
        }

        private bool EvaluateCondition(DataRow row, string condition, List<int> columnsToSearch)
        {
            condition = condition.Trim();

            string op = "";
            int opPos = -1;

            var likeIdx = condition.IndexOf(" like ", StringComparison.OrdinalIgnoreCase);
            if (likeIdx >= 0)
            {
                op = "like";
                opPos = likeIdx;
            }
            else
            {
                var neqIdx = FindOperator(condition, "!=");
                var ltgtIdx = FindOperator(condition, "<>");
                var geIdx = FindOperator(condition, ">=");
                var leIdx = FindOperator(condition, "<=");
                var gtIdx = FindOperator(condition, ">");
                var ltIdx = FindOperator(condition, "<");
                var eqIdx = FindOperator(condition, "=");

                if (geIdx >= 0) { op = ">="; opPos = geIdx; }
                else if (leIdx >= 0) { op = "<="; opPos = leIdx; }
                else if (neqIdx >= 0) { op = "!="; opPos = neqIdx; }
                else if (ltgtIdx >= 0) { op = "<>"; opPos = ltgtIdx; }
                else if (gtIdx >= 0) { op = ">"; opPos = gtIdx; }
                else if (ltIdx >= 0) { op = "<"; opPos = ltIdx; }
                else if (eqIdx >= 0) { op = "="; opPos = eqIdx; }
            }

            if (opPos < 0)
            {
                foreach (var colIndex in columnsToSearch)
                {
                    var cellValue = row[colIndex]?.ToString() ?? "";
                    if (cellValue.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
                return false;
            }

            var rawColumnName = condition.Substring(0, opPos).Trim();
            // like 匹配的是 " like "（6个字符），其他操作符按实际长度
            int opRealLen = op == "like" ? " like ".Length : op.Length;
            var rawValue = condition.Substring(opPos + opRealLen).Trim();
            var columnName = rawColumnName.Trim();

            var value = rawValue;
            if ((value.StartsWith("'") && value.EndsWith("'")) ||
                (value.StartsWith("\"") && value.EndsWith("\"")))
            {
                value = value.Substring(1, value.Length - 2);
            }

            if (!string.IsNullOrEmpty(columnName) && char.IsLetter(columnName[0]))
            {
                int colIndex = ResolveColumnIndex(columnName);

                if (colIndex >= 0 && columnsToSearch.Contains(colIndex))
                {
                    var cellValue = row[colIndex]?.ToString() ?? "";
                    return CompareValue(cellValue, value, op);
                }

                foreach (var ci in columnsToSearch)
                {
                    var cellValue = row[ci]?.ToString() ?? "";
                    if (cellValue.IndexOf(condition.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
                return false;
            }

            foreach (var colIndex in columnsToSearch)
            {
                var cellValue = row[colIndex]?.ToString() ?? "";
                if (CompareValue(cellValue, columnName + op + rawValue, op))
                    return true;
                if (cellValue.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return false;
        }

        /// <summary>根据列名字符串解析对应的列索引，优先匹配表头行，再匹配冻结行范围内的单元格内容</summary>
        private int ResolveColumnIndex(string columnName)
        {
            // 1. 尝试精确匹配 ColumnName（DataTable 的原始列名）
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                if (_dataTable.Columns[i].ColumnName.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            // 2. 如果设置了表头行，优先在表头行搜索
            if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count)
            {
                for (int i = 0; i < _dataTable.Columns.Count; i++)
                {
                    var cellValue = _dataTable.Rows[_headerRowIndex][i]?.ToString() ?? "";
                    if (cellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        return i;
                    }
                }
            }

            // 3. 回退：在冻结行范围内所有行搜索
            int maxHeaderScan = Math.Min(_frozenRows, _dataTable.Rows.Count);
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                for (int row = 0; row < maxHeaderScan; row++)
                {
                    if (row == _headerRowIndex) continue; // 已在上面搜索过
                    var cellValue = _dataTable.Rows[row][i]?.ToString() ?? "";
                    if (cellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        private int FindOperator(string s, string op)
        {
            int idx = s.IndexOf(op, StringComparison.Ordinal);
            while (idx >= 0)
            {
                int nextIdx = idx + op.Length;
                // 检查后面是否跟着 =，避免 > 错误匹配 >= 中的 >
                if (nextIdx < s.Length && op.Length == 1 && (op[0] == '>' || op[0] == '<' || op[0] == '!') && s[nextIdx] == '=')
                {
                    idx = s.IndexOf(op, nextIdx + 1, StringComparison.Ordinal);
                    continue;
                }
                return idx;
            }
            return -1;
        }

        private bool CompareValue(string cellValue, string filterValue, string op)
        {
            bool numericCompare = false;
            double cellNum = 0, filterNum = 0;
            if (op != "like" && double.TryParse(cellValue, out cellNum) && double.TryParse(filterValue, out filterNum))
            {
                numericCompare = true;
            }

            switch (op.ToLower())
            {
                case "like":
                    // like: 支持 * 和 % 通配符，默认包含匹配
                    if (filterValue.Contains("*") || filterValue.Contains("%"))
                    {
                        var pattern = Regex.Escape(filterValue).Replace("\\*", ".*").Replace("\\%", ".*");
                        return Regex.IsMatch(cellValue, "^" + pattern + "$", RegexOptions.IgnoreCase);
                    }
                    else
                    {
                        // 无通配符时，默认包含匹配（两端自动加 %）
                        return cellValue.IndexOf(filterValue, StringComparison.OrdinalIgnoreCase) >= 0;
                    }

                case "=":
                    return numericCompare
                        ? Math.Abs(cellNum - filterNum) < 0.000001
                        : cellValue.Equals(filterValue, StringComparison.OrdinalIgnoreCase);

                case "!=":
                case "<>":
                    return numericCompare
                        ? Math.Abs(cellNum - filterNum) >= 0.000001
                        : !cellValue.Equals(filterValue, StringComparison.OrdinalIgnoreCase);

                case ">":
                    return numericCompare ? cellNum > filterNum
                        : string.Compare(cellValue, filterValue, StringComparison.OrdinalIgnoreCase) > 0;

                case ">=":
                    return numericCompare ? cellNum >= filterNum
                        : string.Compare(cellValue, filterValue, StringComparison.OrdinalIgnoreCase) >= 0;

                case "<":
                    return numericCompare ? cellNum < filterNum
                        : string.Compare(cellValue, filterValue, StringComparison.OrdinalIgnoreCase) < 0;

                case "<=":
                    return numericCompare ? cellNum <= filterNum
                        : string.Compare(cellValue, filterValue, StringComparison.OrdinalIgnoreCase) <= 0;

                default:
                    return cellValue.IndexOf(filterValue, StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        /// <summary>简单的文本匹配（用于列名/列内容的表头过滤）</summary>
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

        private int _sortColumnIndex = -1;
        private SortOrder _sortOrder = SortOrder.None;

        private void OnColumnHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.ColumnIndex >= _dataGrid.Columns.Count)
                return;

            if (_sortColumnIndex == e.ColumnIndex)
            {
                _sortOrder = _sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                _sortColumnIndex = e.ColumnIndex;
                _sortOrder = SortOrder.Ascending;
            }

            var columnName = _dataGrid.Columns[e.ColumnIndex].DataPropertyName;

            var frozenRows = _dataTable.AsEnumerable().Take(_frozenRows).ToList();
            var dataRows = _dataTable.AsEnumerable().Skip(_frozenRows).ToList();

            var sortedDataRows = dataRows
                .OrderBy(row => row.Field<string>(columnName), new NaturalStringComparer(_sortOrder))
                .ToList();

            var allRows = frozenRows.Concat(sortedDataRows).ToArray();
            _dataTable = allRows.CopyToDataTable();

            RefreshGridLayout();

            UpdateColumnHeaderSortGlyph();
        }

        private void UpdateColumnHeaderSortGlyph()
        {
            foreach (DataGridViewColumn col in _dataGrid.Columns)
            {
                if (col.Index == _sortColumnIndex)
                {
                    col.HeaderCell.SortGlyphDirection = _sortOrder;
                }
                else
                {
                    col.HeaderCell.SortGlyphDirection = SortOrder.None;
                }
            }
        }

        private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < _dataGrid.Rows.Count)
            {
                _dataGrid.Rows[e.RowIndex].Height = 24;
            }
            // 列关联颜色检查
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                ColumnLinkChecker.CheckCellColor(e, _visibleColumnIndices, _dataGrid, _filePath, FrozenDir, _frozenRows, _columnLinks);
                // 关联表跳转高亮 — 整行变色
                if (_highlightsReady && _highlightRows.Contains(e.RowIndex))
                {
                    e.CellStyle.BackColor = Color.FromArgb(255, 255, 200);
                    e.CellStyle.SelectionBackColor = Color.FromArgb(255, 240, 150);
                }
            }
        }

        /// <summary>单元格提示 — 关联字段值无效时显示红色原因</summary>
        private void OnCellToolTipTextNeeded(object? sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            e.ToolTipText = ColumnLinkChecker.GetInvalidLinkToolTip(
                _filePath, FrozenDir, _visibleColumnIndices, _dataGrid, e.RowIndex, e.ColumnIndex, _frozenRows, _columnLinks);
        }

        /// <summary>右键菜单 — 复制单元格内容 + 打开关联表</summary>
        private void OnGridCellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right || e.RowIndex < 0 || e.ColumnIndex < 0) return;

            int adjCol = e.ColumnIndex;
            if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__")
                adjCol = e.ColumnIndex - 1;
            int origCol = adjCol;
            if (adjCol >= 0 && adjCol < _visibleColumnIndices.Count)
                origCol = _visibleColumnIndices[adjCol];

            string currentRel = GetCurrentRelativePath();
            var link = _columnLinks.FirstOrDefault(l =>
                l.SourceFile.Equals(currentRel, StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == origCol);

            var cellValue = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
            var ctx = new ContextMenuStrip();
            ctx.Font = new Font("Microsoft YaHei UI", 9f);

            // 1. 复制当前单元格的全部内容
            var copyItem = new ToolStripMenuItem("复制");
            copyItem.Click += (s, args) =>
            {
                if (!string.IsNullOrEmpty(cellValue))
                {
                    Clipboard.SetText(cellValue);
                }
            };
            ctx.Items.Add(copyItem);

            // 1.5 编辑单元格
            var editItem = new ToolStripMenuItem("编辑");
            editItem.Click += (s, args) =>
            {
                var cell = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                _editingFromFrozen = false;
                EditCell(cell, _dataGrid);
            };
            ctx.Items.Add(editItem);

            // 2. 打开关联表（有链接才显示，但根据值有效性决定是否可用）
            var linkItem = new ToolStripMenuItem("打开关联表");
            if (link == null)
            {
                linkItem.Enabled = false;
            }
            else
            {
                bool isValid = !string.IsNullOrEmpty(cellValue) && ColumnLinkChecker.IsLinkValueValid(_filePath, FrozenDir, origCol, cellValue, _columnLinks);
                linkItem.Enabled = isValid;
                linkItem.Click += (s, args) =>
                {
                    string targetFullPath = Path.Combine(FrozenDir, link.TargetFile);
                    if (!File.Exists(targetFullPath))
                    {
                        MessageBox.Show(this, "关联文件不存在: " + link.TargetFile, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    var remarks = CsvSettings.GetFileRemarksByDirectory(FrozenDir);
                    var displayName = Path.GetFileName(targetFullPath);
                    if (remarks.TryGetValue(link.TargetFile, out var remark) && !string.IsNullOrEmpty(remark))
                        displayName = remark + " (" + Path.GetFileName(targetFullPath) + ")";
                    var standaloneForm = new CsvReaderStandaloneForm(targetFullPath, link.TargetColumnIndex, cellValue, _frozenRows, link.TargetColumnName, _headerRowIndex, isLinkedFromCell: true, linkConfig: link, displayName: displayName);
                    standaloneForm.Show(this);
                };
            }
            ctx.Items.Add(linkItem);

            ctx.Show(_dataGrid, _dataGrid.PointToClient(Cursor.Position));
        }

        private string GetCurrentRelativePath()
        {
            if (!string.IsNullOrEmpty(FrozenDir) && _filePath.StartsWith(FrozenDir, StringComparison.OrdinalIgnoreCase))
                return _filePath.Substring(FrozenDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(_filePath);
        }

        private void OnRowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            int rowIndex;
            if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__" && e.RowIndex >= 0)
            {
                var cellValue = _dataGrid.Rows[e.RowIndex].Cells[0].Value;
                rowIndex = cellValue != null ? Convert.ToInt32(cellValue) : e.RowIndex + 1 + _frozenRows;
            }
            else
            {
                rowIndex = e.RowIndex + 1 + _frozenRows;
            }

            var rowHeaderBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, _dataGrid.RowHeadersWidth, e.RowBounds.Height);

            using (var brush = new SolidBrush(_dataGrid.RowHeadersDefaultCellStyle.ForeColor))
            {
                var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                e.Graphics.DrawString(rowIndex.ToString(), _dataGrid.Font, brush, rowHeaderBounds, format);
            }
        }

        private void OnFrozenRowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            var rowIndex = e.RowIndex + 1;
            var rowHeaderBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, _dataGridHeader.RowHeadersWidth, e.RowBounds.Height);

            using (var brush = new SolidBrush(_dataGridHeader.RowHeadersDefaultCellStyle.ForeColor))
            {
                var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                e.Graphics.DrawString(rowIndex.ToString(), _dataGridHeader.Font, brush, rowHeaderBounds, format);
            }
        }

        private static string GetExcelColumnName(int columnIndex)
        {
            var result = new StringBuilder();
            int index = columnIndex;

            do
            {
                result.Insert(0, (char)('A' + (index % 26)));
                index = index / 26 - 1;
            } while (index >= 0);

            return result.ToString();
        }

        /// <summary>重建高亮行集合 — 扫描数据表找出关联列匹配的行，并自动滚动到首行（支持多值 OR 匹配）</summary>
        private void RebuildHighlightRows()
        {
            _highlightRows.Clear();
            if (_highlightColumnIndex < 0 || _highlightValues.Count == 0 || _dataTable.Columns.Count == 0)
            {
                _highlightsReady = false;
                return;
            }
            for (int i = _frozenRows; i < _dataTable.Rows.Count; i++)
            {
                if (_highlightColumnIndex < _dataTable.Columns.Count)
                {
                    var val = _dataTable.Rows[i][_highlightColumnIndex]?.ToString()?.Trim() ?? "";
                    if (_highlightValues.Any(v => val.Equals(v, StringComparison.OrdinalIgnoreCase)))
                        _highlightRows.Add(i);
                }
            }
            _highlightsReady = true;
            // 自动填入高级过滤：列名=值 进行精准匹配（支持多值 OR）
            if (_highlightRows.Count > 0)
            {
                string colRef = !string.IsNullOrEmpty(_highlightColumnName)
                    ? _highlightColumnName : $"Col{_highlightColumnIndex}";
                string filterValue = _highlightValues.Count == 1
                    ? $"{colRef}='{_highlightValues[0]}'"
                    : string.Join(" or ", _highlightValues.Select(v => $"{colRef}='{v}'"));
                _txtAdvancedFilter.Text = filterValue;
            }
        }

        /// <summary>使用ExcelDataReader读取Excel文件(.xls/.xlsx)为DataTable</summary>
        private static DataTable? ReadExcelToDataTable(string filePath)
        {
            try
            {
                // 注册编码提供者（ExcelDataReader 需要）
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
                    return null;

                var sheet = dataSet.Tables[0];
                var dataTable = new DataTable();

                int maxCols = sheet.Columns.Count;
                for (int i = 0; i < maxCols; i++)
                {
                    var columnName = ConvertToColumnLetter(i);
                    dataTable.Columns.Add(columnName, typeof(string));
                }

                // 读取行数据
                for (int r = 0; r < sheet.Rows.Count; r++)
                {
                    var newRow = dataTable.NewRow();
                    for (int c = 0; c < maxCols; c++)
                    {
                        newRow[c] = sheet.Rows[r][c]?.ToString() ?? "";
                    }
                    dataTable.Rows.Add(newRow);
                }

                return dataTable;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>将数字索引转换为Excel列字母（A, B, ..., Z, AA, AB, ...）</summary>
        private static string ConvertToColumnLetter(int index)
        {
            var result = new StringBuilder();
            int num = index;
            do
            {
                result.Insert(0, (char)('A' + (num % 26)));
                num = num / 26 - 1;
            } while (num >= 0);
            return result.ToString();
        }

        /// <summary>获取ExcelDataReader单元格的字符串值（已弃用，保留以保持兼容）</summary>
        private static string GetCellStringValue(object cell)
        {
            return cell?.ToString() ?? "";
        }
    }

    /// <summary>列选择对话框 — ListView展示冻结行数据，复选框选择列，新增关联按钮（加宽180px）</summary>
    public class ColumnSelectorDialog : Form
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
        private int _headerRowIndex = -1;              // 表头行索引（-1=未设置）
        private List<(string display, string fullPath)> _fileEntries = new();
        private List<ColumnLinkConfig> _columnLinks = new();

        public ColumnSelectorDialog(DataTable dataTable, bool[] checkedStates, int frozenRows, int lockedColumnIndex,
            string rootDir = "", string currentFilePath = "", List<(string display, string fullPath)>? fileEntries = null,
            List<ColumnLinkConfig>? columnLinks = null, int headerRowIndex = -1)
        {
            _lockedColumnIndex = lockedColumnIndex;
            _checkedStates = (bool[])checkedStates.Clone();
            _rootDir = rootDir;
            _currentFilePath = currentFilePath;
            _frozenRows = frozenRows;
            _headerRowIndex = headerRowIndex;  // 接收表头行索引
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

            // 表头名列：显示表头行对应的字段名（如果有设置表头行）
            if (_headerRowIndex >= 0 && _headerRowIndex < dataTable.Rows.Count && _headerRowIndex < frozenCount)
                _lvColumns.Columns.Add($"字段名(行{_headerRowIndex + 1})", 140, System.Windows.Forms.HorizontalAlignment.Left);

            // 其余冻结行
            for (int r = 0; r < frozenCount; r++)
            {
                if (r == _headerRowIndex) continue; // 已作为字段名列显示
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

                // 先添加字段名（如果设置了表头行）
                if (_headerRowIndex >= 0 && _headerRowIndex < dataTable.Rows.Count && _headerRowIndex < frozenCount)
                {
                    var headerVal = dataTable.Rows[_headerRowIndex][col]?.ToString() ?? "";
                    item.SubItems.Add(headerVal);
                }

                // 添加其余冻结行的值
                for (int r = 0; r < frozenCount; r++)
                {
                    if (r == _headerRowIndex) continue;
                    var val = dataTable.Rows[r][col]?.ToString() ?? "";
                    item.SubItems.Add(val);
                }

                var link = _columnLinks.Find(l =>
                    l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                    l.SourceColumnIndex == col);
                item.SubItems.Add(link != null ? $"⚡{Path.GetFileNameWithoutExtension(link.TargetFile)}" : "");

                _lvColumns.Items.Add(item);
            }

            // 高亮标记表头行所在的列头（ColumnHeader 不支持 BackColor，跳过）
            
            _lvColumns.EndUpdate();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "选择显示列";
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

                // 刷新ListView中的关联列显示
                UpdateLinkColumnInListView();
            }
        }

        /// <summary>刷新ListView中关联列的显示</summary>
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
                    // 更新关联文本（sub-items 最后一个）
                    var linkText = item.SubItems[item.SubItems.Count - 1];

                    // 构建关联显示文本：优先显示表头名称
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
                _lblSelection.Text = !string.IsNullOrEmpty(_selectedTargetColumnName)
                    ? $"已选: {displayPath} → [{_selectedTargetColumnName}]"
                    : $"已选: {displayPath} → 列{_selectedTargetColIndex}";
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
            _existingLink = null;
            _btnDelete.Visible = false;
            UpdateSelectionLabel();
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

            Controls.Add(new Label { Text = "提示: 关联基于列的位置索引。双击目标文件后在弹出面板中单选目标列。",
                Location = new Point(10, 320), Size = new Size(480, 30), Font = new Font("Microsoft YaHei UI", 8f),
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

        public int SelectedColumnIndex => _selectedIndex;
        public string SelectedColumnName => _selectedColumnName ?? "";

        public TargetColumnSelectDialog(DataTable frozenData, int frozenRows, int initialSelection, int headerRowIndex = -1)
        {
            _selectedIndex = initialSelection;
            _headerRowIndex = headerRowIndex;
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
            Text = "选择目标列";
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
                Text = "取消",
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
