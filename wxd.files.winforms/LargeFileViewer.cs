using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
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
        private ThemeField _fldFilter = new();
        private TextBox _txtFileFilter => _fldFilter.Edit;
        private TabControl _tabContent = new();
        private Panel _pnlLeft = new();
        private Panel _pnlLeftTop = new();
        private ThemeButton _btnSelectFolder = new();
        private ThemeButton _btnRefresh = new();
        private ThemeButton _btnHelp = new();
        private ThemeButton _btnTabMenu = new();
        private Panel _tabPanel = new();
        private Panel _tabStripCover = new();
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

        public LargeFileViewer(string? initialDir = null)
        {
            _currentDir = initialDir ?? "";
            InitializeComponent();
            CreateControls();
            ApplyTheme();
            // 独立窗口，主窗翻明暗时遍历不到，只能自己订阅
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;

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

            UpdateTitle();
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            ClientSize = new Size(1200, 800);
            Text = "日志分析";
            MinimumSize = new Size(800, 600);
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;
            Font = new Font("Microsoft YaHei UI", 9f);
            try { Icon = new Icon(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.ico")); } catch { }
            ResumeLayout(false);
        }

        private void CreateControls()
        {
            SuspendLayout();

            // ==================== 左侧面板 ====================
            _pnlLeft.Dock = DockStyle.Fill;
            _pnlLeft.Padding = new Padding(0);

            // 左侧顶部工具栏：按钮一行、过滤框一行、文件数一行，坐标由 LayoutLeftTop 算
            _pnlLeftTop.Dock = DockStyle.Top;
            _pnlLeftTop.Padding = new Padding(0);

            _btnSelectFolder = new ThemeButton { Text = "选择目录", Primary = true };
            _btnSelectFolder.Click += OnSelectFolderClick;

            _btnRefresh = new ThemeButton { Text = "刷新" };
            _btnRefresh.Click += OnRefreshClick;

            _btnHelp = new ThemeButton { Text = "使用说明" };
            _btnHelp.Click += OnHelpClick;

            _fldFilter = new ThemeField { Placeholder = "过滤文件名..." };

            _lblFileCount = new Label { Text = "共 0 个文件", AutoSize = true };

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

            _pnlLeftTop.Controls.AddRange(new Control[] { _btnSelectFolder, _btnRefresh, _btnHelp, _fldFilter, _lblFileCount });
            LayoutLeftTop();

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
                SizeMode = TabSizeMode.Fixed,
                Multiline = true
            };
            _tabContent.SelectedIndexChanged += OnTabSelectedChanged;

            // ▼ 下拉按钮（列出所有打开的标签页）
            _btnTabMenu = new ThemeButton { Text = "▼", Ghost = true };
            _btnTabMenu.Click += (s, e) => ShowTabMenu(_btnTabMenu);

            // 标签页面板：TabControl + 页签条盖板 + ▼按钮（后两个盖在右上角）
            _tabPanel = new Panel { Dock = DockStyle.Fill };
            _tabPanel.SizeChanged += (s, e) =>
            {
                int size = Theme.P.ButtonHeight;
                _btnTabMenu.SetBounds(_tabPanel.Width - size - 6, 4, size, size);
                LayoutTabStripCover();
            };
            _tabContent.SelectedIndexChanged += (s, e) => LayoutTabStripCover();
            _tabContent.ControlAdded += (s, e) => LayoutTabStripCover();
            _tabContent.ControlRemoved += (s, e) => LayoutTabStripCover();
            _tabPanel.Controls.Add(_tabContent);
            _tabPanel.Controls.Add(_tabStripCover);
            _tabPanel.Controls.Add(_btnTabMenu);
            _tabStripCover.BringToFront();
            _btnTabMenu.BringToFront();

            // 初始空白页
            _tabContent.TabPages.Add(MakeWelcomePage());

            // ==================== 分割容器 ====================
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 400,
                FixedPanel = FixedPanel.Panel1
            };
            _splitContainer.Panel1.Controls.Add(_pnlLeft);
            _splitContainer.Panel2.Controls.Add(_tabPanel);

            _splitContainer.SplitterMoved += (s, e) =>
            {
                if (_lstFiles.Columns.Count > 0)
                    _lstFiles.Columns[0].Width = Math.Max(380, _splitContainer.Panel1.Width - 20);
            };

            Controls.Add(_splitContainer);

            _pnlLeftTop.Resize += (s, e) => LayoutLeftTop();

            _lstFiles.Resize += (s, e) =>
            {
                if (_lstFiles.Columns.Count > 0 && _lstFiles.ClientSize.Width > 380)
                    _lstFiles.Columns[0].Width = _lstFiles.ClientSize.Width - 4;
            };

            ResumeLayout(false);
        }

        /// <summary>
        /// 左侧工具区逐行算坐标：按钮、过滤框、文件数各一行。
        /// 过滤框跟着面板宽度铺满，不再像原来那样固定 250 宽被截断。
        /// </summary>
        private void LayoutLeftTop()
        {
            // 边距和间距与阅读器工具条用同一套数字，两个区域才看得出是同一风格
            int margin = ThemeMetrics.Margin, gap = ThemeMetrics.Gap;
            int rowH = ThemeMetrics.RowH;
            int w = Math.Max(120, _pnlLeftTop.Width - margin * 2);
            int y = 8;

            int x = margin;
            int wSelect = ThemeButton.FitWidth(_btnSelectFolder.Text);
            int wRefresh = ThemeButton.FitWidth(_btnRefresh.Text);
            int wHelp = ThemeButton.FitWidth(_btnHelp.Text);
            _btnSelectFolder.SetBounds(x, y, wSelect, rowH); x += wSelect + gap;
            _btnRefresh.SetBounds(x, y, wRefresh, rowH); x += wRefresh + gap;
            _btnHelp.SetBounds(x, y, wHelp, rowH);

            y += rowH + gap;
            _fldFilter.SetBounds(margin, y, w, rowH);

            y += rowH + gap - 2;
            _lblFileCount.Location = new Point(margin + 2, y);
            _pnlLeftTop.Height = _lblFileCount.Bottom + 8;
        }

        /// <summary>
        /// 页签条右侧空白是 comctl32 拿系统灰画死的，BackColor 和 SetWindowTheme 都压不住，
        /// 所以用一块同色面板只盖空白段，页签本体仍由 TabControl 自绘；多行时空白只出现在最后一行。
        /// </summary>
        private void LayoutTabStripCover()
        {
            int n = _tabContent.TabCount;
            if (n == 0 || !_tabContent.IsHandleCreated)
            {
                _tabStripCover.Visible = false;
                return;
            }
            var last = _tabContent.GetTabRect(n - 1);
            int w = _tabContent.Width - last.Right;
            if (w <= 1)
            {
                _tabStripCover.Visible = false;
                return;
            }
            // 从页签条顶端开始盖，顺带盖住上面那条系统色描边和下面的分隔线
            // 但页签换行时 last 已经在下面一行，从 0 盖会把第一行右侧的页签一起涂掉，只盖它自己那一行
            var first = _tabContent.GetTabRect(0);
            int top = last.Top > first.Top ? last.Top : 0;
            _tabStripCover.SetBounds(last.Right, top, w, last.Bottom - top + 4);
            _tabStripCover.Visible = true;
        }

        /// <summary>没有页签时的占位页，两种皮肤都靠令牌取色。</summary>
        private TabPage MakeWelcomePage()
        {
            var page = new TabPage("欢迎") { BackColor = Theme.P.Window };
            var label = new Label
            {
                Text = "在左侧文件列表中选择文件打开\n或点击\"打开文件\"按钮",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = Theme.Title,
                ForeColor = Theme.P.Muted,
                BackColor = Color.Transparent,
                Tag = "welcome"
            };
            page.Controls.Add(label);
            return page;
        }

        /// <summary>取色全部走当前皮肤令牌；换皮肤或翻明暗时由 Theme.Changed 重新走一遍。</summary>
        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _splitContainer.BackColor = p.Divider;

            _pnlLeft.BackColor = p.Surface;
            _pnlLeftTop.BackColor = p.Chrome;
            foreach (var b in new[] { _btnSelectFolder, _btnRefresh, _btnHelp })
            {
                b.BackColor = p.Chrome;
                b.Font = Theme.Body;
                b.Invalidate();
            }
            _fldFilter.ApplyTheme();
            _lblFileCount.Font = Theme.Caption;
            _lblFileCount.ForeColor = p.Muted;
            _lblFileCount.BackColor = p.Chrome;

            _lstFiles.Font = Theme.List;
            _lstFiles.BackColor = p.Surface;
            _lstFiles.ForeColor = p.Text;

            _tabPanel.BackColor = p.Window;
            _tabContent.Font = Theme.Body;
            _tabContent.BackColor = p.Window;
            _tabContent.ItemSize = new Size(180, p.ButtonHeight + 4);
            _tabStripCover.BackColor = p.Window;
            _btnTabMenu.BackColor = p.Window;
            _btnTabMenu.Font = Theme.Caption;
            _btnTabMenu.Invalidate();

            // 欢迎页是程序自己造的占位页，页签里带的文件阅读器由控件自己响应主题
            foreach (TabPage page in _tabContent.TabPages)
            {
                bool welcome = false;
                foreach (Control c in page.Controls)
                    if (c.Tag as string == "welcome") { c.ForeColor = p.Muted; welcome = true; }
                if (welcome) page.BackColor = p.Window;
            }

            LayoutLeftTop();
            LayoutTabStripCover();
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
                UpdateTitle();
            }
        }

        private void UpdateTitle()
        {
            if (!string.IsNullOrEmpty(_currentDir))
                Text = $"日志分析 - {_currentDir}";
            else
                Text = "日志分析";
        }

        private void OnRefreshClick(object? sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentDir) && Directory.Exists(_currentDir))
                LoadFilesAsync();
        }

        private void OnHelpClick(object? sender, EventArgs e)
        {
            const string helpText = @"📄 日志分析 - 使用说明
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

◆ 基本操作

  · 【选择目录】—— 选择包含文本文件的文件夹
    （程序会自动递归扫描所有子目录中的文本文件）

  · 【刷新】—— 重新加载当前目录的文件列表
    （当目录内文件有变动时使用）

  · 【过滤文件】—— 在过滤框中输入关键词实时筛选
    ├ 支持 || 运算符：匹配任意关键词（或）
    ├ 支持 && 运算符：同时匹配所有关键词（与）
    └ 按 Esc 键快速清空过滤


◆ 文件列表操作

  · 单击文件名 —— 在右侧标签页中预览/打开该文件
  · 双击文件名 —— 在新标签页中打开该文件
  · 显示文件相对路径，方便区分不同子目录下的同名文件
  · 自动识别并显示系统文件图标


◆ 支持的文件格式

  · 常见文本格式：.txt .log .csv .json .xml .yaml .yml
  · 配置文件：.ini .cfg .conf .config .props .targets
  · 代码文件：.cs .cpp .c .h .py .js .ts .java .go 等
  · Markdown：.md .rst
  · 网页相关：.html .css .scss .less
  · 脚本文件：.sh .bat .ps1 .sql
  · 项目文件：.sln .csproj .vcxproj
  · 其他短扩展名文件也会自动尝试打开


◆ 标签页操作

  · 点击标签页标题切换已打开的文件
  · 点击标签页上的 × 关闭文件
  · 右键标签页弹出菜单：关闭当前 / 关闭其它 / 关闭全部
  · 点击右上角 ▼ 按钮可快速定位到任意打开的标签页
  · 支持同时打开多个文件，无数量限制


◆ 阅读控制（大文件阅读核心功能）

  · 【虚拟滚动】
    ├ 高性能浏览超大文件，流畅不卡顿
    └ 支持 10GB+ 文件秒级打开

  · 【行号跳转】
    ├ 输入行号直接跳转到指定行
    └ 支持行号输入框和回车快捷跳转

  · 【全文搜索】
    ├ 输入关键词在文件中搜索
    ├ Aa 复选框：区分大小写搜索
    ├ .* 复选框：使用正则表达式搜索
    ├ |< 按钮：跳转到第一个搜索结果
    ├ 上一个 / 下一个：逐条跳转搜索结果
    ├ >| 按钮：跳转到最后一个搜索结果
    └ 搜索结果高亮显示，搜索计数实时更新

  · 【上下文行】
    ├ 通过 NumericUpDown 控件设置搜索时显示的上下文行数（0~999）
    ├ 搜索时可显示匹配行的前后文内容
    └ 便于理解匹配内容的上下文语境

  · 【导出功能】
    ├ 将搜索结果导出为新文件
    └ 导出时自动包含设定的上下文行数，方便查看匹配行的前后文


◆ 性能特点

  · 流式读取：不将整个文件加载到内存
  · 后台异步构建行偏移索引，支持快速行号定位
  · 内存占用严格控制在合理范围内
  · 大文件打开速度与文件大小无关，秒级响应";

            HelpForm.ShowHelp(this, "使用说明 - 日志分析", helpText);
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
                    AlertForm.Show(this, $"加载文件失败:" + GlobalConst.NewLine + ex.Message, "错误", AlertType.Error, exception: ex);
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
                var showHidden = CsvSettings.ShowHiddenFiles;

                foreach (var file in Directory.GetFiles(currentDir))
                {
                    token.ThrowIfCancellationRequested();

                    if (!showHidden)
                    {
                        try { if ((File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                    }

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
                    if (!showHidden)
                    {
                        try { if ((File.GetAttributes(subDir) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                    }
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
                _lstFiles.Items.Add(new ListViewItem("未找到文本文件") { ForeColor = Theme.P.Muted });
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
                _lstFiles.Items.Add(new ListViewItem("未找到文本文件") { ForeColor = Theme.P.Muted });
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
                    // 使用扩展名 + USEFILEATTRIBUTES，避免 Shell 缓存锁定文件/目录
                    SHGetFileInfo(key, 0, out shfi, (uint)Marshal.SizeOf(shfi), SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
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

                _tabContent.DrawMode = TabDrawMode.OwnerDrawFixed;
                _tabContent.DrawItem += (s, e) =>
                {
                    var tabPage = _tabContent.TabPages[e.Index];
                    var tabRect = _tabContent.GetTabRect(e.Index);
                    bool isSelected = _tabContent.SelectedIndex == e.Index;

                    int bottomExtra = isSelected ? 2 : 0;
                    var roundRect = new Rectangle(
                        tabRect.X + 2, tabRect.Y + 2,
                        tabRect.Width - 4, tabRect.Height - 2 + bottomExtra);
                    int radius = Theme.P.Radius + 2;

                    using (var path = new GraphicsPath())
                    {
                        path.AddArc(roundRect.X, roundRect.Y, radius * 2, radius * 2, 180, 90);
                        path.AddArc(roundRect.Right - radius * 2, roundRect.Y, radius * 2, radius * 2, 270, 90);
                        path.AddLine(roundRect.Right, roundRect.Y + radius, roundRect.Right, roundRect.Bottom);
                        path.AddLine(roundRect.Right, roundRect.Bottom, roundRect.X, roundRect.Bottom);
                        path.AddLine(roundRect.X, roundRect.Bottom, roundRect.X, roundRect.Y + radius);
                        path.CloseFigure();

                        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

                        var bgColor = isSelected ? Theme.P.Accent : Theme.P.Selection;
                        var borderColor = isSelected ? Theme.P.Accent : Theme.P.Border;

                        using (var brush = new SolidBrush(bgColor))
                        {
                            e.Graphics.FillPath(brush, path);
                        }
                        using (var pen = new Pen(borderColor, 1f))
                        {
                            e.Graphics.DrawPath(pen, path);
                        }
                    }

                    var textRect = new Rectangle(tabRect.X + 10, tabRect.Y + 4, tabRect.Width - 32, tabRect.Height - 8);
                    Color textColor = isSelected ? Theme.P.AccentText : Theme.P.Text;
                    TextRenderer.DrawText(e.Graphics, tabPage.Text, _tabContent.Font, textRect, textColor,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

                    var closeRect = new Rectangle(tabRect.Right - 18, tabRect.Top + 6, 12, 12);
                    Color closeColor = isSelected ? Theme.Mix(Theme.P.Accent, Color.White, 0.5f) : Theme.P.Muted;
                    using (var pen = new Pen(closeColor, 1.5f))
                    {
                        e.Graphics.DrawLine(pen, closeRect.Left + 2, closeRect.Top + 2, closeRect.Right - 2, closeRect.Bottom - 2);
                        e.Graphics.DrawLine(pen, closeRect.Right - 2, closeRect.Top + 2, closeRect.Left + 2, closeRect.Bottom - 2);
                    }

                    if (isSelected)
                    {
                        using (var coverPen = new Pen(Theme.P.Accent, 2f))
                        {
                            e.Graphics.DrawLine(coverPen,
                                roundRect.X + 1, roundRect.Bottom,
                                roundRect.Right - 1, roundRect.Bottom);
                        }
                    }
                };
            }
        }

        private void CloseTab(int tabIndex)
        {
            if (tabIndex < 0 || tabIndex >= _tabContent.TabPages.Count)
                return;

            var tabPage = _tabContent.TabPages[tabIndex];

            // ★ 先清理 reader 资源（文件流、缓存等），再移除标签页
            foreach (Control ctrl in tabPage.Controls)
            {
                if (ctrl is LargeFileReaderControl reader)
                {
                    reader.Cleanup();
                }
            }

            _tabContent.TabPages.RemoveAt(tabIndex);
            _lstFiles.SelectedItems.Clear();

            // 如果没有标签页了，恢复欢迎页
            if (_tabContent.TabPages.Count == 0)
                _tabContent.TabPages.Add(MakeWelcomePage());
        }

        /// <summary>▼按钮下拉菜单 — 列出所有打开的标签页</summary>
        private void ShowTabMenu(Button btn)
        {
            var menu = new ContextMenuStrip
            {
                Font = Theme.List,
                ShowImageMargin = false,
                ShowCheckMargin = true
            };
            for (int i = 0; i < _tabContent.TabPages.Count; i++)
            {
                var tp = _tabContent.TabPages[i];
                bool isCurrent = (i == _tabContent.SelectedIndex);
                var item = new ToolStripMenuItem(tp.Text)
                {
                    Tag = i,
                    Checked = isCurrent,
                    CheckState = isCurrent ? CheckState.Checked : CheckState.Unchecked,
                    Font = isCurrent ? Theme.ListBold : Theme.List
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
            var menu = new ContextMenuStrip
            {
                Font = Theme.List,
                ShowImageMargin = false
            };

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
            // ★ 先记录当前标签页数量，避免 CloseTab 中添加的"欢迎"页导致无限循环
            int count = _tabContent.TabPages.Count;
            for (int i = 0; i < count; i++)
                CloseTab(0);
        }

        private void OnTabSelectedChanged(object? sender, EventArgs e)
        {
            var dirText = !string.IsNullOrEmpty(_currentDir) ? $" - {_currentDir}" : "";
            if (_tabContent.SelectedTab?.Tag is string path)
            {
                Text = $"日志分析{dirText} - {Path.GetFileName(path)}";
            }
            else
            {
                Text = $"日志分析{dirText}";
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 防抖计时器带着指向本窗体的闭包，不 Dispose 就等于窗体关闭后还被消息循环攥着一份
            try
            {
                _filterDebounceTimer?.Stop();
                _filterDebounceTimer?.Dispose();
                _filterDebounceTimer = null;
            }
            catch { }
            try { _loadingCts?.Cancel(); _loadingCts?.Dispose(); _loadingCts = null; } catch { }
            // ★ 主动清理所有标签页，避免 Dispose 递归时堆积导致 UI 卡顿
            CloseAllTabs();
            base.OnFormClosing(e);
        }
    }
}
