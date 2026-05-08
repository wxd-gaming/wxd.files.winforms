using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    public partial class LargeFileViewer : Form
    {
        // ==================== 控件 ====================
        private SplitContainer _splitContainer = new();
        private ListView _lstFiles = new();
        private TextBox _txtFileFilter = new();
        private TabControl _tabContent = new();
        private Panel _pnlLeft = new();
        private Panel _pnlLeftTop = new();
        private Button _btnSelectFolder = new();
        private Button _btnRefresh = new();
        private Label _lblFileCount = new();
        private ImageList _imgList = new();

        // ==================== Win32 API（系统图标）====================
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr SHGetFileInfo(
            string pszPath, uint dwFileAttributes,
            out SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int    iIcon;
            public uint   dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
        }

        private const uint SHGFI_ICON              = 0x100;
        private const uint SHGFI_SMALLICON         = 0x000000001;
        private const uint SHGFI_USEFILEATTRIBUTES = 0x10;
        private const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;

        // ==================== 状态 ====================
        private string _currentDir = "";
        private List<FileInfo> _allFiles = new();
        private CancellationTokenSource? _loadingCts;
        private readonly object _lockObj = new();
        private System.Windows.Forms.Timer? _filterDebounceTimer;
        private bool _tabCloseHandlerAdded = false;

        private class FileInfo
        {
            public string Name { get; set; } = "";
            public string FullPath { get; set; } = "";
            public string RelativePath { get; set; } = "";
        }

        // ==================== 配色 ====================
        private static readonly Color _colorLightBg = Color.FromArgb(250, 250, 250);
        private static readonly Color _colorLightBorder = Color.FromArgb(218, 220, 224);
        private static readonly Color _colorLightText = Color.FromArgb(32, 33, 36);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(95, 99, 104);
        private static readonly Color _colorBtnBg = Color.FromArgb(0, 122, 204);

        public LargeFileViewer(string? initialDir = null)
        {
            _currentDir = initialDir ?? "";
            InitializeComponent();
            CreateControls();
            ApplyTheme();

            Shown += async (s, e) =>
            {
                await Task.Delay(200);
                if (_splitContainer.SplitterDistance < 300)
                    _splitContainer.SplitterDistance = 300;
                if (_lstFiles.Columns.Count > 0)
                    _lstFiles.Columns[0].Width = 280;
                _lstFiles.Refresh();
            };

            if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
            {
                LoadFilesAsync();
            }
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            ClientSize = new Size(1200, 800);
            Text = "大文件阅读器";
            MinimumSize = new Size(800, 600);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9f);
            try { Icon = new Icon(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.ico")); } catch { }
            ResumeLayout(false);
        }

        private void CreateControls()
        {
            SuspendLayout();

            // ==================== 左侧面板 ====================
            _pnlLeft.Dock = DockStyle.Fill;
            _pnlLeft.BackColor = Color.White;
            _pnlLeft.Padding = new Padding(0);

            // 左侧顶部工具栏
            _pnlLeftTop.Height = 80;
            _pnlLeftTop.Dock = DockStyle.Top;
            _pnlLeftTop.BackColor = _colorLightBg;
            _pnlLeftTop.Padding = new Padding(6, 4, 6, 4);

            // 选择目录按钮
            _btnSelectFolder = new Button
            {
                Text = "选择目录",
                Location = new Point(6, 6),
                Size = new Size(80, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = _colorBtnBg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnSelectFolder.FlatAppearance.BorderSize = 0;
            _btnSelectFolder.Click += OnSelectFolderClick;

            // 刷新按钮
            _btnRefresh = new Button
            {
                Text = "刷新",
                Location = new Point(90, 6),
                Size = new Size(50, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = _colorBtnBg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnRefresh.FlatAppearance.BorderSize = 0;
            _btnRefresh.Click += OnRefreshClick;

            // 过滤文本框
            _txtFileFilter = new TextBox
            {
                Location = new Point(6, 40),
                Size = new Size(250, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                PlaceholderText = "过滤文件名..."
            };

            _lblFileCount = new Label
            {
                Text = "共 0 个文件",
                Location = new Point(6, 66),
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary
            };

            _filterDebounceTimer = new System.Windows.Forms.Timer { Interval = 300 };
            _filterDebounceTimer.Tick += (s, e) =>
            {
                _filterDebounceTimer!.Stop();
                ApplyFileFilter();
            };
            _txtFileFilter.TextChanged += (s, e) =>
            {
                _filterDebounceTimer?.Stop();
                _filterDebounceTimer?.Start();
            };
            _txtFileFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    _txtFileFilter.Text = "";
                    e.Handled = true;
                }
            };

            _pnlLeftTop.Controls.AddRange(new Control[] { _btnSelectFolder, _btnRefresh, _txtFileFilter, _lblFileCount });

            // 文件列表
            _imgList = new ImageList();
            _imgList.ImageSize = new Size(16, 16);
            _imgList.ColorDepth = ColorDepth.Depth32Bit;
            // 使用系统文件图标作为默认图标
            var defaultShfi = default(SHFILEINFO);
            SHGetFileInfo(".txt", 0, out defaultShfi, (uint)Marshal.SizeOf(defaultShfi),
                          SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
            _imgList.Images.Add("file", Icon.FromHandle(defaultShfi.hIcon));

            _lstFiles = new ListView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new Font("Microsoft YaHei UI", 9f),
                View = View.Details,
                FullRowSelect = true,
                GridLines = false,
                HideSelection = false,
                SmallImageList = _imgList,
                MultiSelect = false
            };
            _lstFiles.Columns.Add("文件", 380, HorizontalAlignment.Left);
            _lstFiles.Click += OnListViewClick;
            _lstFiles.DoubleClick += OnListViewDoubleClick;
            _lstFiles.ColumnWidthChanging += (s, e) =>
            {
                if (e.ColumnIndex == 0 && e.NewWidth < 280)
                {
                    e.NewWidth = 280;
                    e.Cancel = false;
                }
            };

            _pnlLeft.Controls.Add(_lstFiles);
            _pnlLeft.Controls.Add(_pnlLeftTop);

            // ==================== 右侧标签页 ====================
            _tabContent = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9f),
                ItemSize = new Size(200, 28),
                SizeMode = TabSizeMode.Fixed
            };
            _tabContent.SelectedIndexChanged += OnTabSelectedChanged;

            // 初始空白页
            var welcomePage = new TabPage("欢迎")
            {
                BackColor = _colorLightBg
            };
            var welcomeLabel = new Label
            {
                Text = "在左侧文件列表中选择文件打开\n或点击\"打开文件\"按钮",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei UI", 12f),
                ForeColor = _colorLightTextSecondary
            };
            welcomePage.Controls.Add(welcomeLabel);
            _tabContent.TabPages.Add(welcomePage);

            // ==================== 分割容器 ====================
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 400,
                FixedPanel = FixedPanel.Panel1,
                BackColor = _colorLightBorder
            };
            _splitContainer.Panel1.Controls.Add(_pnlLeft);
            _splitContainer.Panel2.Controls.Add(_tabContent);

            _splitContainer.SplitterMoved += (s, e) =>
            {
                if (_lstFiles.Columns.Count > 0)
                    _lstFiles.Columns[0].Width = Math.Max(380, _splitContainer.Panel1.Width - 20);
            };

            Controls.Add(_splitContainer);

            _pnlLeftTop.Resize += (s, e) =>
            {
                _txtFileFilter.Width = Math.Max(100, _pnlLeftTop.Width - 16);
            };

            _lstFiles.Resize += (s, e) =>
            {
                if (_lstFiles.Columns.Count > 0 && _lstFiles.ClientSize.Width > 380)
                    _lstFiles.Columns[0].Width = _lstFiles.ClientSize.Width - 4;
            };

            ResumeLayout(false);
        }

        private void ApplyTheme()
        {
            BackColor = _colorLightBg;
            _pnlLeft.BackColor = Color.White;
            _lstFiles.BackColor = Color.White;
            _lstFiles.ForeColor = _colorLightText;
            _tabContent.BackColor = Color.White;
        }

        // ==================== 目录选择与文件加载 ====================
        private void OnSelectFolderClick(object? sender, EventArgs e)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "选择包含文本文件的目录",
                ShowNewFolderButton = false
            };

            if (!string.IsNullOrEmpty(_currentDir))
                dialog.SelectedPath = _currentDir;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _currentDir = dialog.SelectedPath;
                LoadFilesAsync();
            }
        }

        private void OnRefreshClick(object? sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
                LoadFilesAsync();
        }

        private async void LoadFilesAsync()
        {
            _loadingCts?.Cancel();
            _loadingCts = new CancellationTokenSource();
            var token = _loadingCts.Token;

            _lblFileCount.Text = "加载中...";

            try
            {
                await Task.Run(() => ScanFiles(_currentDir, token), token);

                if (!token.IsCancellationRequested)
                {
                    this.Invoke(() =>
                    {
                        RefreshFileList();
                        ApplyFileFilter();
                    });
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                this.Invoke(() =>
                {
                    MessageBox.Show(this, $"加载文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }

        private static readonly HashSet<string> _textExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".txt", ".log", ".csv", ".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf",
            ".cs", ".cpp", ".c", ".h", ".hpp", ".py", ".js", ".ts", ".jsx", ".tsx",
            ".java", ".rb", ".go", ".rs", ".swift", ".kt", ".dart",
            ".md", ".rst", ".html", ".css", ".scss", ".less",
            ".sql", ".sh", ".bat", ".ps1", ".vbs",
            ".cfg", ".config", ".props", ".targets",
            ".sln", ".csproj", ".vcxproj",
            ".json", ".xml", ".yml", ".yaml", ".toml"
        };

        private static bool IsTextFile(string path)
        {
            var ext = Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext)) return true;
            return _textExtensions.Contains(ext) || ext.Length <= 5; // 短扩展名也尝试打开
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
                foreach (var file in Directory.GetFiles(currentDir))
                {
                    token.ThrowIfCancellationRequested();

                    if (!IsTextFile(file))
                        continue;

                    var relativePath = GetRelativePath(rootDir, file);
                    files.Add(new FileInfo
                    {
                        Name = Path.GetFileName(file),
                        FullPath = file,
                        RelativePath = relativePath
                    });
                }

                foreach (var subDir in Directory.GetDirectories(currentDir))
                {
                    token.ThrowIfCancellationRequested();
                    ScanDirectory(subDir, rootDir, files, token);
                }
            }
            catch (UnauthorizedAccessException) { }
            catch { }
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

        private void RefreshFileList()
        {
            _lstFiles.BeginUpdate();
            _lstFiles.Items.Clear();

            if (_lstFiles.Columns.Count > 0)
                _lstFiles.Columns[0].Width = 380;

            if (_allFiles.Count == 0)
            {
                _lstFiles.Items.Add(new ListViewItem("未找到文本文件") { ForeColor = _colorLightTextSecondary });
                _lstFiles.EndUpdate();
                _lblFileCount.Text = "共 0 个文件";
                return;
            }

            foreach (var file in _allFiles.OrderBy(f => f.RelativePath))
            {
                var item = new ListViewItem(file.RelativePath)
                {
                    Tag = file,
                    ToolTipText = file.FullPath,
                    ImageKey = GetFileIconKey(file.FullPath)
                };
                _lstFiles.Items.Add(item);
            }

            _lstFiles.EndUpdate();
            _lblFileCount.Text = $"共 {_lstFiles.Items.Count} 个文件";
        }

        private void ApplyFileFilter()
        {
            var filter = _txtFileFilter.Text.Trim();

            _lstFiles.BeginUpdate();
            _lstFiles.Items.Clear();

            if (_lstFiles.Columns.Count > 0)
                _lstFiles.Columns[0].Width = 380;

            if (_allFiles.Count == 0)
            {
                _lstFiles.Items.Add(new ListViewItem("未找到文本文件") { ForeColor = _colorLightTextSecondary });
                _lstFiles.EndUpdate();
                _lblFileCount.Text = "共 0 个文件";
                return;
            }

            if (filter.Contains("||"))
            {
                var terms = filter.Split(new[] { "||" }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var file in _allFiles.OrderBy(f => f.RelativePath))
                {
                    if (!terms.Any(t => file.RelativePath.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0))
                        continue;
                    _lstFiles.Items.Add(new ListViewItem(file.RelativePath) { Tag = file, ToolTipText = file.FullPath, ImageKey = GetFileIconKey(file.FullPath) });
                }
            }
            else if (filter.Contains("&&"))
            {
                var terms = filter.Split(new[] { "&&" }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var file in _allFiles.OrderBy(f => f.RelativePath))
                {
                    if (!terms.All(t => file.RelativePath.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0))
                        continue;
                    _lstFiles.Items.Add(new ListViewItem(file.RelativePath) { Tag = file, ToolTipText = file.FullPath, ImageKey = GetFileIconKey(file.FullPath) });
                }
            }
            else if (!string.IsNullOrEmpty(filter))
            {
                foreach (var file in _allFiles.OrderBy(f => f.RelativePath))
                {
                    if (file.RelativePath.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    _lstFiles.Items.Add(new ListViewItem(file.RelativePath) { Tag = file, ToolTipText = file.FullPath, ImageKey = GetFileIconKey(file.FullPath) });
                }
            }
            else
            {
                foreach (var file in _allFiles.OrderBy(f => f.RelativePath))
                {
                    _lstFiles.Items.Add(new ListViewItem(file.RelativePath) { Tag = file, ToolTipText = file.FullPath, ImageKey = GetFileIconKey(file.FullPath) });
                }
            }

            _lstFiles.EndUpdate();
            _lblFileCount.Text = $"共 {_lstFiles.Items.Count} 个文件";
        }

        // ==================== 系统文件图标 ====================
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
            return "file";
        }

        // ==================== 文件选择与打开 ====================
        private void OnListViewClick(object? sender, EventArgs e)
        {
            if (_lstFiles.SelectedItems.Count > 0)
            {
                var item = _lstFiles.SelectedItems[0];
                if (item.Tag is FileInfo fileInfo && System.IO.File.Exists(fileInfo.FullPath))
                {
                    OpenFile(fileInfo.FullPath);
                }
            }
        }

        private void OnListViewDoubleClick(object? sender, EventArgs e)
        {
            if (_lstFiles.SelectedItems.Count > 0)
            {
                var item = _lstFiles.SelectedItems[0];
                if (item.Tag is FileInfo fileInfo && System.IO.File.Exists(fileInfo.FullPath))
                {
                    OpenFile(fileInfo.FullPath);
                }
            }
        }

        private void OpenFile(string filePath)
        {
            // 检查是否已经打开
            foreach (TabPage tab in _tabContent.TabPages)
            {
                // 跳过"欢迎"页
                if (tab.Text == "欢迎") continue;
                if (tab.Tag is string path && path.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                {
                    _tabContent.SelectedTab = tab;
                    return;
                }
            }

            // 如果欢迎页是唯一标签页，移除它
            if (_tabContent.TabPages.Count == 1 && _tabContent.TabPages[0].Text == "欢迎")
            {
                _tabContent.TabPages.Clear();
            }

            // 创建新的标签页
            var tabPage = new TabPage(Path.GetFileName(filePath))
            {
                Tag = filePath,
                ToolTipText = filePath
            };

            var reader = new LargeFileReaderControl
            {
                Dock = DockStyle.Fill
            };
            tabPage.Controls.Add(reader);
            _tabContent.TabPages.Add(tabPage);
            _tabContent.SelectedTab = tabPage;

            // 添加关闭按钮
            AddCloseButtonToTab();
            reader.FileClosed += (_, _) => CloseTab(_tabContent.SelectedIndex);
            reader.FileExported += (_, _) =>
            {
                if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
                    LoadFilesAsync();
            };

            // 立即打开文件
            Application.DoEvents();
            reader.Open(filePath);
        }

        private void AddCloseButtonToTab()
        {
            if (!_tabCloseHandlerAdded)
            {
                _tabCloseHandlerAdded = true;

                _tabContent.MouseDown += (s, e) =>
                {
                    for (int i = 0; i < _tabContent.TabPages.Count; i++)
                    {
                        var tabRect = _tabContent.GetTabRect(i);
                        var closeRect = new Rectangle(tabRect.Right - 16, tabRect.Top + 4, 12, 12);

                        if (closeRect.Contains(e.Location))
                        {
                            CloseTab(i);
                            return;
                        }
                    }
                };

                _tabContent.DrawMode = TabDrawMode.OwnerDrawFixed;
                _tabContent.DrawItem += (s, e) =>
                {
                    var tabPage = _tabContent.TabPages[e.Index];
                    var tabRect = _tabContent.GetTabRect(e.Index);

                    bool isSelected = _tabContent.SelectedIndex == e.Index;
                    using (var brush = new SolidBrush(isSelected ? Color.White : Color.FromArgb(240, 240, 240)))
                    {
                        e.Graphics.FillRectangle(brush, tabRect);
                    }

                    var textRect = new Rectangle(tabRect.X + 4, tabRect.Y + 4, tabRect.Width - 24, tabRect.Height - 8);
                    TextRenderer.DrawText(e.Graphics, tabPage.Text, _tabContent.Font, textRect, _colorLightText,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

                    var closeRect = new Rectangle(tabRect.Right - 16, tabRect.Top + 6, 10, 10);
                    using (var pen = new Pen(_colorLightTextSecondary, 1.5f))
                    {
                        e.Graphics.DrawLine(pen, closeRect.Left, closeRect.Top, closeRect.Right, closeRect.Bottom);
                        e.Graphics.DrawLine(pen, closeRect.Right, closeRect.Top, closeRect.Left, closeRect.Bottom);
                    }

                    using (var pen = new Pen(_colorLightBorder))
                    {
                        e.Graphics.DrawLine(pen, tabRect.Left, tabRect.Bottom - 1, tabRect.Right, tabRect.Bottom - 1);
                    }
                };
            }
        }

        private void CloseTab(int tabIndex)
        {
            if (tabIndex < 0 || tabIndex >= _tabContent.TabPages.Count)
                return;

            var tabPage = _tabContent.TabPages[tabIndex];

            foreach (Control ctrl in tabPage.Controls)
            {
                if (ctrl is LargeFileReaderControl reader)
                {
                    reader.Cleanup();
                }
                ctrl.Dispose();
            }

            _tabContent.TabPages.RemoveAt(tabIndex);
            _lstFiles.SelectedItems.Clear();

            // 如果没有标签页了，恢复欢迎页
            if (_tabContent.TabPages.Count == 0)
            {
                var welcomePage = new TabPage("欢迎")
                {
                    BackColor = _colorLightBg
                };
                var welcomeLabel = new Label
                {
                    Text = "在左侧文件列表中选择文件打开\n或点击\"打开文件\"按钮",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Microsoft YaHei UI", 12f),
                    ForeColor = _colorLightTextSecondary
                };
                welcomePage.Controls.Add(welcomeLabel);
                _tabContent.TabPages.Add(welcomePage);
            }
        }

        private void OnTabSelectedChanged(object? sender, EventArgs e)
        {
            if (_tabContent.SelectedTab?.Tag is string path)
            {
                Text = $"大文件阅读器 - {Path.GetFileName(path)}";
            }
            else
            {
                Text = "大文件阅读器";
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _loadingCts?.Cancel();
            base.OnFormClosing(e);
        }
    }
}
