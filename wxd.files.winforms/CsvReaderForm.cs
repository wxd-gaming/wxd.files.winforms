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
        private ListView _lstFiles = new();
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

        // ==================== 配色方案（简洁清晰的Office风格）====================
        private static readonly Color _colorBg = Color.FromArgb(240, 242, 245);
        private static readonly Color _colorLightBg = Color.FromArgb(240, 242, 245);
        private static readonly Color _colorSurface = Color.White;
        private static readonly Color _colorPanelBg = Color.FromArgb(247, 248, 250);
        private static readonly Color _colorBorder = Color.FromArgb(209, 213, 219);
        private static readonly Color _colorLightBorder = Color.FromArgb(224, 228, 234);
        private static readonly Color _colorDarkBorder = Color.FromArgb(179, 184, 194);
        private static readonly Color _colorLightText = Color.FromArgb(31, 41, 55);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(107, 114, 128);
        private static readonly Color _colorPrimary = Color.FromArgb(37, 99, 235);
        private static readonly Color _colorPrimaryLight = Color.FromArgb(219, 234, 254);
        private static readonly Color _colorPrimaryDark = Color.FromArgb(29, 78, 216);
        private static readonly Color _colorSuccess = Color.FromArgb(22, 163, 74);
        private static readonly Color _colorHoverBg = Color.FromArgb(243, 244, 246);
        private static readonly Color _colorHeaderBg = Color.FromArgb(243, 244, 246);
        private static readonly Color _colorAltRow = Color.FromArgb(249, 250, 251);
        private static readonly Color _colorSelectedBg = Color.FromArgb(219, 234, 254);
        private static readonly Color _colorGridLine = Color.FromArgb(209, 213, 219);
        private static readonly Color _colorRowNumBg = Color.FromArgb(243, 244, 246);

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
            // 注册ExcelDataReader的编码支持
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            _currentDir = initialDir ?? "";
            _paneIndex = paneIndex;
            InitializeComponent();
            CreateControls();
            ApplyTheme();

            // 窗体加载完成后设置列宽
            this.Shown += async (s, e) =>
            {
                // 延迟设置，确保布局已完成
                await Task.Delay(200);

                // 强制设置SplitContainer宽度
                if (_splitContainer.SplitterDistance < 300)
                {
                    _splitContainer.SplitterDistance = 300;
                }

                // 强制设置列宽
                if (_lstFiles.Columns.Count > 0)
                {
                    _lstFiles.Columns[0].Width = 280;
                }

                // 强制刷新
                _lstFiles.Refresh();
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
                FlatAppearance = { BorderSize = 0, MouseOverBackColor = _colorPrimaryDark }
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
                FlatAppearance = { BorderSize = 0, MouseOverBackColor = Color.FromArgb(224, 228, 234) }
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
                BackColor = Color.FromArgb(248, 249, 251)
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
            // 使用系统文件图标作为默认图标
            var csvShfi = default(SHFILEINFO);
            SHGetFileInfo(".csv", 0, out csvShfi, (uint)Marshal.SizeOf(csvShfi),
                          SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
            _imgList.Images.Add("default", Icon.FromHandle(csvShfi.hIcon));

            _lstFiles = new ListView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f),
                View = View.Details,
                FullRowSelect = true,
                GridLines = false,
                HideSelection = false,
                SmallImageList = _imgList,
                MultiSelect = false,
                BackColor = _colorSurface,
                ForeColor = _colorLightText,
                OwnerDraw = true
            };
            _lstFiles.DrawItem += OnListViewDrawItem;
            _lstFiles.DrawSubItem += OnListViewDrawSubItem;
            _lstFiles.DrawColumnHeader += OnListViewDrawColumnHeader;
            _lstFiles.Columns.Add("文件", 380, System.Windows.Forms.HorizontalAlignment.Left);
            _lstFiles.Click += OnListViewClick;
            _lstFiles.DoubleClick += OnListViewDoubleClick;

            // 阻止列宽被自动调整
            _lstFiles.ColumnWidthChanging += (s, e) =>
            {
                if (e.ColumnIndex == 0 && e.NewWidth < 280)
                {
                    e.NewWidth = 280;
                    e.Cancel = false;
                }
            };

            // 添加ListView到左侧面板
            _pnlLeft.Controls.Add(_lstFiles);
            _pnlLeft.Controls.Add(_pnlLeftTop);

            // ==================== 右侧标签页 ====================
            _tabContent = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9f),
                ItemSize = new Size(200, 28),
                SizeMode = TabSizeMode.Fixed,
                BackColor = _colorSurface,
                Padding = new Point(12, 4)
            };

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

            // SplitContainer调整大小时更新列宽
            _splitContainer.SplitterMoved += (s, e) =>
            {
                if (_lstFiles.Columns.Count > 0)
                {
                    _lstFiles.Columns[0].Width = Math.Max(380, _splitContainer.Panel1.Width - 20);
                }
            };

            Controls.Add(_splitContainer);

            // 调整大小处理
            _pnlLeftTop.Resize += (s, e) =>
            {
                _txtFileFilter.Width = Math.Max(100, _pnlLeftTop.Width - 16);
            };

            // ListView大小改变时保持列宽
            _lstFiles.Resize += (s, e) =>
            {
                if (_lstFiles.Columns.Count > 0 && _lstFiles.ClientSize.Width > 380)
                {
                    _lstFiles.Columns[0].Width = _lstFiles.ClientSize.Width - 4;
                }
            };

            ResumeLayout(false);
        }

        private void ApplyTheme()
        {
            BackColor = _colorBg;
            _pnlLeft.BackColor = _colorSurface;
            _lstFiles.BackColor = _colorSurface;
            _lstFiles.ForeColor = _colorLightText;
            _tabContent.BackColor = _colorSurface;
        }

        // ==================== 自定义ListView绘制（现代风格）====================
        private void OnListViewDrawColumnHeader(object? sender, DrawListViewColumnHeaderEventArgs e)
        {
            using var headerBrush = new SolidBrush(_colorHeaderBg);
            e.Graphics.FillRectangle(headerBrush, e.Bounds);
            TextRenderer.DrawText(e.Graphics, e.Header.Text, new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                e.Bounds, _colorLightTextSecondary, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            using var bottomPen = new Pen(_colorDarkBorder);
            e.Graphics.DrawLine(bottomPen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            using var rightPen = new Pen(_colorLightBorder);
            e.Graphics.DrawLine(rightPen, e.Bounds.Right - 1, e.Bounds.Top, e.Bounds.Right - 1, e.Bounds.Bottom - 1);
        }

        private void OnListViewDrawItem(object? sender, DrawListViewItemEventArgs e)
        {
            if (e.Item.Selected)
            {
                e.Graphics.FillRectangle(new SolidBrush(_colorSelectedBg), e.Bounds);
                using var pen = new Pen(_colorPrimaryLight);
                e.Graphics.DrawRectangle(pen, e.Bounds.X, e.Bounds.Y, e.Bounds.Width - 1, e.Bounds.Height - 1);
            }
            else if (e.ItemIndex % 2 == 0)
            {
                e.Graphics.FillRectangle(new SolidBrush(_colorAltRow), e.Bounds);
            }
            else
            {
                e.Graphics.FillRectangle(new SolidBrush(_colorSurface), e.Bounds);
            }

            // 绘制图标和文本（OwnerDraw 模式下需要手动绘制）
            if (e.Item.ImageKey != null && _imgList.Images.ContainsKey(e.Item.ImageKey))
            {
                int imgIdx = _imgList.Images.IndexOfKey(e.Item.ImageKey);
                if (imgIdx >= 0)
                    _imgList.Draw(e.Graphics, e.Bounds.X + 3, e.Bounds.Y + (e.Bounds.Height - 16) / 2, 16, 16, imgIdx);
            }
            var textRect = new Rectangle(e.Bounds.X + 22, e.Bounds.Y + 2, e.Bounds.Width - 26, e.Bounds.Height - 4);
            TextRenderer.DrawText(e.Graphics, e.Item.Text, _lstFiles.Font, textRect,
                e.Item.Selected ? Color.FromArgb(30, 64, 175) : _colorLightText,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

            // 底部行线
            using var linePen = new Pen(_colorLightBorder);
            e.Graphics.DrawLine(linePen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
        }

        private void OnListViewDrawSubItem(object? sender, DrawListViewSubItemEventArgs e)
        {
            // DrawSubItem 由 DrawItem 处理，这里留空避免重复绘制
        }

        // ==================== 目录选择 ====================
        // ==================== 打开单个文件 ====================
        private void OnOpenCsvClick(object? sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog
            {
                Title = "选择CSV文件",
                Filter = "CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*",
                Multiselect = false
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                OpenFile(dialog.FileName, FileType.Csv);
            }
        }

        private void OnOpenExcelClick(object? sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog
            {
                Title = "选择Excel文件",
                Filter = "Excel文件 (*.xls;*.xlsx)|*.xls;*.xlsx|所有文件 (*.*)|*.*",
                Multiselect = false
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                OpenFile(dialog.FileName, FileType.Excel);
            }
        }

        // ==================== 目录选择 ====================
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

            _lstFiles.Items.Clear();
            _allFiles.Clear();

            // 显示加载中
            _lstFiles.Items.Add(new ListViewItem("加载中..."));

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
                // 获取当前目录的CSV文件
                foreach (var file in Directory.GetFiles(currentDir, "*.csv"))
                {
                    token.ThrowIfCancellationRequested();

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
                foreach (var file in Directory.GetFiles(currentDir, "*.xls"))
                {
                    token.ThrowIfCancellationRequested();

                    var relativePath = GetRelativePath(rootDir, file);
                    files.Add(new FileInfo
                    {
                        Name = Path.GetFileName(file),
                        FullPath = file,
                        RelativePath = relativePath,
                        Type = FileType.Excel
                    });
                }

                foreach (var file in Directory.GetFiles(currentDir, "*.xlsx"))
                {
                    token.ThrowIfCancellationRequested();

                    var relativePath = GetRelativePath(rootDir, file);
                    files.Add(new FileInfo
                    {
                        Name = Path.GetFileName(file),
                        FullPath = file,
                        RelativePath = relativePath,
                        Type = FileType.Excel
                    });
                }

                // 递归子目录
                foreach (var subDir in Directory.GetDirectories(currentDir))
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
        // ==================== 刷新文件列表 ====================
        private void RefreshFileList()
        {
            _lstFiles.BeginUpdate();
            _lstFiles.Items.Clear();

            // 确保列宽设置正确
            if (_lstFiles.Columns.Count > 0)
            {
                _lstFiles.Columns[0].Width = 380;
            }

            if (_allFiles.Count == 0)
            {
                var emptyItem = new ListViewItem("未找到CSV/Excel文件")
                {
                    ForeColor = _colorLightTextSecondary
                };
                _lstFiles.Items.Add(emptyItem);
                _lstFiles.EndUpdate();
                _lblFileCount.Text = "共 0 个文件";
                return;
            }

            foreach (var file in _allFiles.OrderBy(f => f.RelativePath))
            {
                // 显示相对路径作为文件名（包含目录结构）
                var item = new ListViewItem(file.RelativePath)
                {
                    Tag = file,
                    ToolTipText = file.FullPath,
                    ImageKey = GetFileIconKey(file.FullPath)
                };
                _lstFiles.Items.Add(item);
            }

            _lstFiles.EndUpdate();

            // 更新文件数量标签
            _lblFileCount.Text = $"共 {_lstFiles.Items.Count} 个文件";
        }

        // ==================== 文件过滤 ====================
        private void OnFileFilterChanged(object? sender, EventArgs e)
        {
            // 防抖：重启定时器，300ms 内无新输入才真正过滤
            _filterDebounceTimer?.Stop();
            _filterDebounceTimer?.Start();
        }

        private void ApplyFileFilter()
        {
            var filter = _txtFileFilter.Text.Trim();

            _lstFiles.BeginUpdate();
            _lstFiles.Items.Clear();

            // 确保列宽设置正确
            if (_lstFiles.Columns.Count > 0)
            {
                _lstFiles.Columns[0].Width = 380;
            }

            if (_allFiles.Count == 0)
            {
                var emptyItem = new ListViewItem("未找到CSV/Excel文件")
                {
                    ForeColor = _colorLightTextSecondary
                };
                _lstFiles.Items.Add(emptyItem);
                _lstFiles.EndUpdate();
                _lblFileCount.Text = "共 0 个文件";
                return;
            }

            // 处理运算符
            var processedFilter = filter.Replace("||", " | ").Replace("&&", " & ");

            foreach (var file in _allFiles.OrderBy(f => f.RelativePath))
            {
                // 如果有过滤条件，检查是否匹配
                if (!string.IsNullOrEmpty(filter))
                {
                    if (!MatchesFilter(file.RelativePath, processedFilter))
                        continue;
                }

                var item = new ListViewItem(file.RelativePath)
                {
                    Tag = file,
                    ToolTipText = file.FullPath,
                    ImageKey = GetFileIconKey(file.FullPath)
                };
                _lstFiles.Items.Add(item);
            }

            _lstFiles.EndUpdate();

            // 更新文件数量标签
            _lblFileCount.Text = $"共 {_lstFiles.Items.Count} 个文件";
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

        // ==================== 系统文件图标 ====================
        private string GetFileIconKey(string filePath)
        {
            try
            {
                var ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (!_imgList.Images.ContainsKey(ext))
                {
                    var shfi = default(SHFILEINFO);
                    SHGetFileInfo(filePath, 0, out shfi, (uint)Marshal.SizeOf(shfi), SHGFI_ICON | SHGFI_SMALLICON);
                    if (shfi.hIcon != IntPtr.Zero)
                    {
                        _imgList.Images.Add(ext, Icon.FromHandle(shfi.hIcon));
                        return ext;
                    }
                }
                else return ext;
            }
            catch { }
            return "default";
        }

        // ==================== 文件选择处理 ====================
        private void OnListViewClick(object? sender, EventArgs e)
        {
            if (_lstFiles.SelectedItems.Count > 0)
            {
                var item = _lstFiles.SelectedItems[0];
                if (item.Tag is FileInfo fileInfo && System.IO.File.Exists(fileInfo.FullPath))
                {
                    OpenFile(fileInfo.FullPath, fileInfo.Type);
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
                    OpenFile(fileInfo.FullPath, fileInfo.Type);
                }
            }
        }

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
                viewer = new ExcelViewerControl(filePath, _paneIndex) { Dock = DockStyle.Fill };
            else
                viewer = new CsvViewerControl(filePath, _paneIndex) { Dock = DockStyle.Fill };

            // 创建新标签页
            var tabPage = new TabPage(Path.GetFileName(filePath))
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
                _allFiles.Select(f => (f.RelativePath, f.FullPath)).ToList());
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                var linkedPath = dialog.SelectedFilePath;
                if (string.IsNullOrEmpty(linkedPath) || !File.Exists(linkedPath)) return;

                // 在新独立窗口中打开关联文件
                var standaloneForm = new CsvReaderStandaloneForm(linkedPath);
                standaloneForm.Show(this);
            }
        }

        private void AddCloseButtonToTab(TabPage tabPage)
        {
            // 事件处理程序只在第一次添加标签页时注册一次
            if (!_tabCloseHandlerAdded)
            {
                _tabCloseHandlerAdded = true;

                // 使用TabControl的MouseDown事件处理关闭
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
                        bgColor = Color.FromArgb(37, 99, 235);
                        borderColor = Color.FromArgb(37, 99, 235);
                    }
                    else
                    {
                        bgColor = Color.FromArgb(219, 234, 254);
                        borderColor = Color.FromArgb(191, 219, 254);
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
                    Color textColor = isSelected ? Color.White : Color.FromArgb(30, 64, 175);
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

            // 清除ListView选中状态，允许重新打开同一文件
            _lstFiles.SelectedItems.Clear();
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
        private Panel _pnlTop = new();
        private DataTable _dataTable = new();
        private DataTable _headerTable = new();        // 冻结的表头数据
        private string _filePath = "";
        /// <summary>当前文件完整路径</summary>
        public string FilePath => _filePath;
        private char _delimiter = ',';
        private int _frozenRows = 0;                   // 冻结行数
        private int _paneIndex = -1;                    // 所属窗格索引（-1表示独立打开，使用全局设置）
        private bool[] _columnCheckStates = Array.Empty<bool>(); // 列选中状态

        /// <summary>请求关联打开另一张表事件</summary>
        public event EventHandler? LinkTableRequested;

        // 配色
        private static readonly Color _colorBg = Color.FromArgb(240, 242, 245);
        private static readonly Color _colorLightBg = Color.FromArgb(240, 242, 245);
        private static readonly Color _colorSurface = Color.White;
        private static readonly Color _colorPanelBg = Color.FromArgb(247, 248, 250);
        private static readonly Color _colorBorder = Color.FromArgb(209, 213, 219);
        private static readonly Color _colorLightBorder = Color.FromArgb(224, 228, 234);
        private static readonly Color _colorDarkBorder = Color.FromArgb(179, 184, 194);
        private static readonly Color _colorLightText = Color.FromArgb(31, 41, 55);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(107, 114, 128);
        private static readonly Color _colorPrimary = Color.FromArgb(37, 99, 235);
        private static readonly Color _colorPrimaryLight = Color.FromArgb(219, 234, 254);
        private static readonly Color _colorPrimaryDark = Color.FromArgb(29, 78, 216);
        private static readonly Color _colorSuccess = Color.FromArgb(22, 163, 74);
        private static readonly Color _colorHoverBg = Color.FromArgb(243, 244, 246);
        private static readonly Color _colorHeaderBg = Color.FromArgb(243, 244, 246);
        private static readonly Color _colorAltRow = Color.FromArgb(249, 250, 251);
        private static readonly Color _colorSelectedBg = Color.FromArgb(219, 234, 254);
        private static readonly Color _colorGridLine = Color.FromArgb(209, 213, 219);
        private static readonly Color _colorRowNumBg = Color.FromArgb(243, 244, 246);

        public CsvViewerControl(string filePath, int paneIndex = -1)
        {
            _filePath = filePath;
            _paneIndex = paneIndex;
            InitializeComponent();
            LoadCsvAsync();
        }

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
            // 从设置中读取冻结行数（窗格独立记录）
            _frozenRows = _paneIndex >= 0 ? CsvSettings.GetPaneCsvFrozenRows(_paneIndex) : CsvSettings.DefaultFrozenRows;
            cboFrozenRows.SelectedIndex = _frozenRows;
            cboFrozenRows.SelectedIndexChanged += (s, e) =>
            {
                _frozenRows = cboFrozenRows.SelectedIndex;
                // 保存设置（窗格独立记录）
                if (_paneIndex >= 0)
                    CsvSettings.SetPaneCsvFrozenRows(_paneIndex, _frozenRows);
                else
                    CsvSettings.DefaultFrozenRows = _frozenRows;
                CsvSettings.Save();
                RefreshGridLayout();
            };

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

            flowTop.Controls.AddRange(new Control[] { lblFrozenRows, cboFrozenRows, lblHeaderFilter, _btnSelectColumns, _lblInfo });

            // 关联表按钮
            var btnLinkTable = new Button
            {
                Text = "关联表",
                Size = new Size(60, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 0 },
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 2, 10, 0)
            };
            btnLinkTable.Click += (s, e) => LinkTableRequested?.Invoke(this, EventArgs.Empty);
            flowTop.Controls.Add(btnLinkTable);

            // 第二行（分隔线）放入 topTable
            var separatorLine = new Panel
            {
                Height = 1,
                Dock = DockStyle.Top,
                BackColor = Color.FromArgb(224, 228, 234),
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
                Size = new Size(400, 23),
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
                Size = new Size(400, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 10, 0),
                BackColor = Color.FromArgb(248, 249, 251)
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
            _dataGridHeader.DefaultCellStyle.BackColor = Color.FromArgb(249, 250, 252);
            _dataGridHeader.TabStop = false;
            _dataGridHeader.CellMouseDown += (s, e) =>
            {
                // 鼠标点击时同步重置，阻止滚动偏移发生
                if (e.RowIndex >= 0 && _dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;
            };
            _dataGridHeader.CurrentCellChanged += (s, e) =>
            {
                // 键盘导航等情况的兜底处理
                if (_dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;
            };
            _dataGridHeader.CellDoubleClick += OnFrozenCellDoubleClick;

            // 数据表格（主表格）
            _dataGrid = CreateDataGridView();
            _dataGrid.Dock = DockStyle.Fill;
            _dataGrid.RowHeadersWidth = 50;
            _dataGrid.RowPostPaint += OnRowPostPaint;
            _dataGrid.ColumnHeadersVisible = false;

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
                    SelectionForeColor = Color.FromArgb(30, 64, 175),
                    Padding = new Padding(2, 0, 2, 0)
                },
                RowTemplate = new DataGridViewRow { Height = 24 }
            };
            dgv.ColumnHeaderMouseClick += OnColumnHeaderClick;
            dgv.CellFormatting += OnCellFormatting;
            dgv.CellDoubleClick += OnCellDoubleClick;
            return dgv;
        }

        // ==================== 双击编辑功能 ====================
        private bool _editingFromFrozen = false; // 标记是否来自冻结表格
        private bool _isEditing = false; // 防止重复触发编辑

        private void OnCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (_isEditing) return; // 防止重复触发
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var dgv = sender as DataGridView;
                if (dgv == null) return;

                // 如果双击的是冻结表格，不处理（由 OnFrozenCellDoubleClick 处理）
                if (dgv == _dataGridHeader) return;

                // 获取双击的单元格
                var cell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                _editingFromFrozen = false;

                // 打开编辑对话框
                EditCell(cell);
            }
        }

        private void OnFrozenCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (_isEditing) return; // 防止重复触发
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // 获取冻结表格的双击单元格
                var cell = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex];
                _editingFromFrozen = true;

                // 打开编辑对话框
                EditCell(cell);
            }
        }

        private void EditCell(DataGridViewCell cell)
        {
            _isEditing = true; // 设置编辑标志，防止重复触发

            try
            {
                var rowIndex = cell.RowIndex;
                var columnIndex = cell.ColumnIndex;
                var currentValue = cell.Value?.ToString() ?? "";

                // 获取原始数据表中的行列索引
                var (originalRowIndex, originalColumnIndex) = GetOriginalIndices(rowIndex, columnIndex);

                using var dialog = new CellEditDialog(currentValue);
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var newValue = dialog.CellValue;

                    // 更新原始数据表
                    if (originalRowIndex >= 0 && originalRowIndex < _dataTable.Rows.Count &&
                        originalColumnIndex >= 0 && originalColumnIndex < _dataTable.Columns.Count)
                    {
                        _dataTable.Rows[originalRowIndex][originalColumnIndex] = newValue;

                        // 保存到文件
                        SaveCsvFile();

                        // 刷新显示
                        RefreshGridLayout();
                    }
                }
            }
            finally
            {
                // 使用延迟重置标志，确保双击事件完全处理完毕
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 200;
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    _isEditing = false;
                    _editingFromFrozen = false; // 重置冻结行编辑标志
                };
                timer.Start();
            }
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

                File.WriteAllText(_filePath, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"保存文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

                await Task.Run(() => LoadCsvFile());
            }
            catch (Exception ex)
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(() =>
                    {
                        MessageBox.Show(this, $"加载CSV文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            }
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
            using (var reader = new StreamReader(_filePath, Encoding.UTF8, true))
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

                UpdateColumnSelectorButton();

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
        /// 逐行解析CSV内容，兼容单元格内引号包裹的换行
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
                    // 引号转义
                    if (i + 1 < content.Length && content[i + 1] == '"')
                    {
                        sb.Append('"');
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
                    result.Add(sb.ToString().Trim());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }

            result.Add(sb.ToString().Trim());
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

            using var dialog = new ColumnSelectorDialog(_dataTable, _columnCheckStates, _frozenRows, 0);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _columnCheckStates = dialog.GetCheckedStates();
                UpdateColumnSelectorButton();
                RefreshGridLayout();
            }
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

        /// <summary>根据列名字符串解析对应的列索引，优先匹配 ColumnName，再匹配冻结行范围内的单元格内容</summary>
        private int ResolveColumnIndex(string columnName)
        {
            // 1. 尝试精确匹配 ColumnName
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                if (_dataTable.Columns[i].ColumnName.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            // 2. 在冻结行范围内的各列单元格值中搜索列名
            // 因为用户设置的冻结行已经包含了表头行
            int maxHeaderScan = Math.Min(_frozenRows, _dataTable.Rows.Count);
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                for (int row = 0; row < maxHeaderScan; row++)
                {
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

        public CellEditDialog(string initialValue)
        {
            InitializeComponent();
            _txtValue.Text = NormalizeLineBreaksForDisplay(initialValue);
        }

        private static void SetActiveButton(Button active, Button inactive)
        {
            active.BackColor = Color.FromArgb(0, 120, 212);
            active.ForeColor = Color.White;
            active.FlatAppearance.BorderSize = 0;
            inactive.BackColor = Color.FromArgb(240, 242, 245);
            inactive.ForeColor = Color.FromArgb(55, 65, 81);
            inactive.FlatAppearance.BorderSize = 0;
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "编辑单元格";
            ClientSize = new Size(800, 600);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = true;
            MaximizeBox = true;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(248, 249, 251);

            // ==================== 顶部格式切换面板 ====================
            _pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = Color.FromArgb(245, 246, 248),
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
                BackColor = Color.FromArgb(245, 246, 248)
            };

            // 保存按钮
            _btnSave = new Button
            {
                Text = "保存",
                Size = new Size(80, 30),
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK,
                FlatAppearance = { BorderSize = 0, MouseOverBackColor = Color.FromArgb(16, 110, 190) }
            };

            // 取消按钮
            _btnCancel = new Button
            {
                Text = "取消",
                Size = new Size(80, 30),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.FromArgb(240, 242, 245),
                ForeColor = Color.FromArgb(55, 65, 81),
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(209, 213, 219), MouseOverBackColor = Color.FromArgb(229, 231, 235) }
            };

            _pnlBottom.Controls.Add(_btnSave);
            _pnlBottom.Controls.Add(_btnCancel);

            // 窗口调整大小时重新定位按钮（保持在右下角）
            void LayoutButtons()
            {
                _btnCancel.Left = _pnlBottom.Width - _btnCancel.Width - 10;
                _btnSave.Left = _btnCancel.Left - _btnSave.Width - 8;
                _btnSave.Top = 10;
                _btnCancel.Top = 10;
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
        private Panel _pnlTop = new();
        private DataTable _dataTable = new();
        private DataTable _headerTable = new();
        private string _filePath = "";
        /// <summary>当前文件完整路径</summary>
        public string FilePath => _filePath;
        private int _frozenRows = 0;
        private int _paneIndex = -1;                    // 所属窗格索引（-1表示独立打开，使用全局设置）
        private List<int> _visibleColumnIndices = new();
        private bool[] _columnCheckStates = Array.Empty<bool>();

        /// <summary>请求关联打开另一张表事件</summary>
        public event EventHandler? LinkTableRequested;

        // 配色
        private static readonly Color _colorBg = Color.FromArgb(240, 242, 245);
        private static readonly Color _colorLightBg = Color.FromArgb(240, 242, 245);
        private static readonly Color _colorSurface = Color.White;
        private static readonly Color _colorPanelBg = Color.FromArgb(247, 248, 250);
        private static readonly Color _colorBorder = Color.FromArgb(209, 213, 219);
        private static readonly Color _colorLightBorder = Color.FromArgb(224, 228, 234);
        private static readonly Color _colorDarkBorder = Color.FromArgb(179, 184, 194);
        private static readonly Color _colorLightText = Color.FromArgb(31, 41, 55);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(107, 114, 128);
        private static readonly Color _colorPrimary = Color.FromArgb(37, 99, 235);
        private static readonly Color _colorPrimaryLight = Color.FromArgb(219, 234, 254);
        private static readonly Color _colorPrimaryDark = Color.FromArgb(29, 78, 216);
        private static readonly Color _colorSuccess = Color.FromArgb(22, 163, 74);
        private static readonly Color _colorHoverBg = Color.FromArgb(243, 244, 246);
        private static readonly Color _colorHeaderBg = Color.FromArgb(243, 244, 246);
        private static readonly Color _colorAltRow = Color.FromArgb(249, 250, 251);
        private static readonly Color _colorSelectedBg = Color.FromArgb(219, 234, 254);
        private static readonly Color _colorGridLine = Color.FromArgb(209, 213, 219);
        private static readonly Color _colorRowNumBg = Color.FromArgb(243, 244, 246);

        public ExcelViewerControl(string filePath, int paneIndex = -1)
        {
            _filePath = filePath;
            _paneIndex = paneIndex;
            InitializeComponent();
            LoadExcelAsync();
        }

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
            _frozenRows = _paneIndex >= 0 ? CsvSettings.GetPaneExcelFrozenRows(_paneIndex) : CsvSettings.ExcelDefaultFrozenRows;
            cboFrozenRows.SelectedIndex = _frozenRows;
            cboFrozenRows.SelectedIndexChanged += (s, e) =>
            {
                _frozenRows = cboFrozenRows.SelectedIndex;
                if (_paneIndex >= 0)
                    CsvSettings.SetPaneExcelFrozenRows(_paneIndex, _frozenRows);
                else
                    CsvSettings.ExcelDefaultFrozenRows = _frozenRows;
                CsvSettings.Save();
                RefreshGridLayout();
            };

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

            flowTop.Controls.AddRange(new Control[] { lblFrozenRows, cboFrozenRows, lblHeaderFilter, _btnSelectColumns, _lblInfo });

            // 关联表按钮
            var btnLinkTable = new Button
            {
                Text = "关联表",
                Size = new Size(60, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 0 },
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 2, 10, 0)
            };
            btnLinkTable.Click += (s, e) => LinkTableRequested?.Invoke(this, EventArgs.Empty);
            flowTop.Controls.Add(btnLinkTable);

            // ====== 第二行: 分隔线 ======
            var separatorLine = new Panel
            {
                Height = 1,
                Dock = DockStyle.Top,
                BackColor = Color.FromArgb(224, 228, 234),
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
                Size = new Size(400, 23),
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
                Size = new Size(400, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 10, 0),
                BackColor = Color.FromArgb(248, 249, 251)
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
            _dataGridHeader.DefaultCellStyle.BackColor = Color.FromArgb(249, 250, 252);
            _dataGridHeader.TabStop = false;
            _dataGridHeader.CellMouseDown += (s, e) =>
            {
                // 鼠标点击时同步重置，阻止滚动偏移发生
                if (e.RowIndex >= 0 && _dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;
            };
            _dataGridHeader.CurrentCellChanged += (s, e) =>
            {
                // 键盘导航等情况的兜底处理
                if (_dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;
            };
            _dataGridHeader.CellDoubleClick += OnFrozenCellDoubleClick; // 双击编辑

            // 数据表格（主表格）
            _dataGrid = CreateDataGridView();
            _dataGrid.Dock = DockStyle.Fill;
            _dataGrid.RowHeadersWidth = 50;
            _dataGrid.RowPostPaint += OnRowPostPaint;
            _dataGrid.ColumnHeadersVisible = false;

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
                    SelectionForeColor = Color.FromArgb(30, 64, 175),
                    Padding = new Padding(2, 0, 2, 0)
                },
                RowTemplate = new DataGridViewRow { Height = 24 }
            };
            dgv.ColumnHeaderMouseClick += OnColumnHeaderClick;
            dgv.CellFormatting += OnCellFormatting;
            dgv.CellDoubleClick += OnCellDoubleClick;
            return dgv;
        }

        // ==================== 双击编辑功能 ====================
        private bool _editingFromFrozen = false; // 标记是否来自冻结表格
        private bool _isEditing = false; // 防止重复触发编辑

        private void OnCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (_isEditing) return; // 防止重复触发
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var dgv = sender as DataGridView;
                if (dgv == null) return;

                // 如果双击的是冻结表格，不处理（由 OnFrozenCellDoubleClick 处理）
                if (dgv == _dataGridHeader) return;

                // 获取双击的单元格
                var cell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                _editingFromFrozen = false;

                // 打开编辑对话框
                EditCell(cell);
            }
        }

        private void OnFrozenCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (_isEditing) return; // 防止重复触发
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // 获取冻结表格的双击单元格
                var cell = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex];
                _editingFromFrozen = true;

                // 打开编辑对话框
                EditCell(cell);
            }
        }

        private void EditCell(DataGridViewCell cell)
        {
            _isEditing = true; // 设置编辑标志，防止重复触发

            try
            {
                var rowIndex = cell.RowIndex;
                var columnIndex = cell.ColumnIndex;
                var currentValue = cell.Value?.ToString() ?? "";

                // 获取原始数据表中的行列索引
                var (originalRowIndex, originalColumnIndex) = GetOriginalIndices(rowIndex, columnIndex);

                using var dialog = new CellEditDialog(currentValue);
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var newValue = dialog.CellValue;

                    // 更新原始数据表
                    if (originalRowIndex >= 0 && originalRowIndex < _dataTable.Rows.Count &&
                        originalColumnIndex >= 0 && originalColumnIndex < _dataTable.Columns.Count)
                    {
                        _dataTable.Rows[originalRowIndex][originalColumnIndex] = newValue;

                        // 刷新显示
                        RefreshGridLayout();
                    }
                }
            }
            finally
            {
                // 使用延迟重置标志，确保双击事件完全处理完毕
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 200;
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    _isEditing = false;
                    _editingFromFrozen = false; // 重置冻结行编辑标志
                };
                timer.Start();
            }
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

                await Task.Run(() => LoadExcelFile());
            }
            catch (Exception ex)
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(() =>
                    {
                        MessageBox.Show(this, $"加载Excel文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            }
        }

        private void LoadExcelFile()
        {
            if (!System.IO.File.Exists(_filePath))
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, "Excel文件不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error));
                return;
            }

            using var stream = System.IO.File.Open(_filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();

            if (result.Tables.Count == 0)
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, "Excel文件没有工作表", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning));
                return;
            }

            // 使用第一个工作表
            var table = result.Tables[0];
            var dataTable = new DataTable();

            // 添加列
            for (int i = 0; i < table.Columns.Count; i++)
            {
                var columnName = table.Columns[i].ColumnName;
                if (string.IsNullOrEmpty(columnName))
                {
                    columnName = $"Column{i + 1}";
                }
                dataTable.Columns.Add(columnName, typeof(string));
            }

            // 添加行
            foreach (DataRow row in table.Rows)
            {
                var newRow = dataTable.NewRow();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    var value = row[i];
                    newRow[i] = value?.ToString() ?? "";
                }
                dataTable.Rows.Add(newRow);
            }

            this.Invoke(() =>
            {
                _dataTable = dataTable;
                _headerTable = dataTable.Clone();

                // 初始化列选中状态（默认全选，第一列始终选中）
                _columnCheckStates = new bool[_dataTable.Columns.Count];
                for (int i = 0; i < _columnCheckStates.Length; i++)
                    _columnCheckStates[i] = true;

                UpdateColumnSelectorButton();

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

            using var dialog = new ColumnSelectorDialog(_dataTable, _columnCheckStates, _frozenRows, 0);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _columnCheckStates = dialog.GetCheckedStates();
                UpdateColumnSelectorButton();
                RefreshGridLayout();
            }
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

        /// <summary>根据列名字符串解析对应的列索引，优先匹配 ColumnName，再匹配冻结行范围内的单元格内容</summary>
        private int ResolveColumnIndex(string columnName)
        {
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                if (_dataTable.Columns[i].ColumnName.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            // 在冻结行范围内的各列单元格值中搜索列名
            // 因为用户设置的冻结行已经包含了表头行
            int maxHeaderScan = Math.Min(_frozenRows, _dataTable.Rows.Count);
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                for (int row = 0; row < maxHeaderScan; row++)
                {
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
    }

    /// <summary>列选择对话框 — ListView展示冻结行数据，复选框选择列</summary>
    public class ColumnSelectorDialog : Form
    {
        private ListView _lvColumns = new();
        private Button _btnSelectAll = new();
        private Button _btnDeselectAll = new();
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private int _lockedColumnIndex = -1;
        private bool[] _checkedStates;

        public ColumnSelectorDialog(DataTable dataTable, bool[] checkedStates, int frozenRows, int lockedColumnIndex)
        {
            _lockedColumnIndex = lockedColumnIndex;
            _checkedStates = (bool[])checkedStates.Clone();
            InitializeComponent();
            PopulateList(dataTable, frozenRows);
        }

        public bool[] GetCheckedStates() => _checkedStates;

        private void PopulateList(DataTable dataTable, int frozenRows)
        {
            _lvColumns.BeginUpdate();
            _lvColumns.Items.Clear();
            _lvColumns.Columns.Clear();

            int frozenCount = Math.Min(frozenRows, dataTable.Rows.Count);

            // 列头：选择 + 每行冻结行的内容
            _lvColumns.Columns.Add("选择", 50, System.Windows.Forms.HorizontalAlignment.Center);
            for (int r = 0; r < frozenCount; r++)
            {
                _lvColumns.Columns.Add($"行{r + 1}", 160, System.Windows.Forms.HorizontalAlignment.Left);
            }

            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                var item = new ListViewItem("")
                {
                    Tag = col,
                    Checked = _checkedStates[col]
                };

                for (int r = 0; r < frozenCount; r++)
                {
                    var val = dataTable.Rows[r][col]?.ToString() ?? "";
                    item.SubItems.Add(val);
                }

                _lvColumns.Items.Add(item);
            }

            _lvColumns.EndUpdate();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "选择显示列";
            ClientSize = new Size(800, 520);
            MinimumSize = new Size(600, 300);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(248, 249, 251);
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
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(209, 213, 219) },
                BackColor = Color.FromArgb(243, 244, 246),
                ForeColor = Color.FromArgb(31, 41, 55)
            };
            _btnSelectAll.Click += (s, e) => SetAllChecked(true);

            _btnDeselectAll = new Button
            {
                Text = "取消全选",
                Location = new Point(88, 8),
                Size = new Size(80, 26),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(209, 213, 219) },
                BackColor = Color.FromArgb(243, 244, 246),
                ForeColor = Color.FromArgb(31, 41, 55)
            };
            _btnDeselectAll.Click += (s, e) => SetAllChecked(false);

            // ListView — 使用原生 CheckBoxes
            _lvColumns = new ListView
            {
                Location = new Point(10, 42),
                Size = new Size(780, 420),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(31, 41, 55),
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

            // OK 按钮
            _btnOK = new Button
            {
                Text = "确定",
                Location = new Point(612, 468),
                Size = new Size(80, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK,
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 0 }
            };

            // 取消按钮
            _btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(700, 468),
                Size = new Size(70, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(240, 242, 245),
                ForeColor = Color.FromArgb(55, 65, 81),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(209, 213, 219) }
            };

            Controls.AddRange(new Control[] { _btnSelectAll, _btnDeselectAll, _lvColumns, _btnOK, _btnCancel });
            AcceptButton = _btnOK;
            CancelButton = _btnCancel;

            ResumeLayout(false);
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
            BackColor = Color.FromArgb(248, 249, 251);
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
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 0 }
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
                BackColor = Color.FromArgb(240, 242, 245),
                ForeColor = Color.FromArgb(55, 65, 81),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(209, 213, 219) }
            };

            Controls.AddRange(new Control[] { _txtFilter, _lstFiles, _btnOK, _btnCancel });
            AcceptButton = _btnOK;
            CancelButton = _btnCancel;

            ResumeLayout(false);
        }
    }
}
