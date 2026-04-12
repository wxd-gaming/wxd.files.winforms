using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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
        private Button _btnOpenExcel = new();
        private Button _btnOpenCsv = new();
        private ImageList _imgList = new();

        // ==================== 状态字段 ====================
        private string _currentDir = "";
        private List<FileInfo> _allFiles = new();
        private CancellationTokenSource? _loadingCts;
        private readonly object _lockObj = new();

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

        // ==================== 配色方案 ====================
        private static readonly Color _colorLightBg = Color.FromArgb(250, 250, 250);
        private static readonly Color _colorLightBorder = Color.FromArgb(218, 220, 224);
        private static readonly Color _colorLightText = Color.FromArgb(32, 33, 36);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(95, 99, 104);
        private static readonly Color _colorBtnBg = Color.FromArgb(0, 122, 204);

        public CsvReaderForm(string? initialDir = null)
        {
            // 注册ExcelDataReader的编码支持
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            _currentDir = initialDir ?? "";
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
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            ClientSize = new Size(1200, 800);
            Text = "CSV/Excel阅读器";
            MinimumSize = new Size(800, 800);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9f);

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

            // 打开CSV文件按钮
            _btnOpenCsv = new Button
            {
                Text = "打开CSV",                
                Location = new Point(6, 6),
                Size = new Size(70, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = _colorBtnBg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnOpenCsv.FlatAppearance.BorderSize = 0;
            _btnOpenCsv.Click += OnOpenCsvClick;

            // 打开Excel文件按钮
            _btnOpenExcel = new Button
            {
                Text = "打开Excel",
                Location = new Point(80, 6),
                Size = new Size(75, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.FromArgb(0, 150, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnOpenExcel.FlatAppearance.BorderSize = 0;
            _btnOpenExcel.Click += OnOpenExcelClick;

            // 选择目录按钮
            _btnSelectFolder = new Button
            {
                Text = "选择目录",
                Location = new Point(159, 6),
                Size = new Size(70, 28),
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
                Location = new Point(233, 6),
                Size = new Size(50, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = _colorBtnBg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnRefresh.FlatAppearance.BorderSize = 0;
            _btnRefresh.Click += OnRefreshClick;

            // 过滤标签
            _lblFilter = new Label
            {
                Text = "过滤:",
                Location = new Point(6, 42),
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary
            };

            // 过滤文本框
            _txtFileFilter = new TextBox
            {
                Location = new Point(45, 40),
                Size = new Size(238, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
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

            // 提示标签
            var lblHint = new Label
            {
                Text = "支持 || 或 && 运算符",
                Location = new Point(6, 64),
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = _colorLightTextSecondary
            };

            _pnlLeftTop.Controls.AddRange(new Control[] { _btnOpenCsv, _btnOpenExcel, _btnSelectFolder, _btnRefresh, _lblFilter, _txtFileFilter, lblHint });

            // 文件树
            // 初始化图标列表
            _imgList = new ImageList();
            _imgList.ImageSize = new Size(16, 16);
            _imgList.ColorDepth = ColorDepth.Depth32Bit;
            // 添加默认图标
            _imgList.Images.Add("csv", SystemIcons.Application);
            _imgList.Images.Add("excel", SystemIcons.Application);

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
                ItemSize = new Size(200, 28),  // 增加标签页宽度到200
                SizeMode = TabSizeMode.Fixed   // 固定宽度模式
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
                _txtFileFilter.Width = _pnlLeftTop.Width - _txtFileFilter.Left - 8;
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
            BackColor = _colorLightBg;
            _pnlLeft.BackColor = Color.White;
            _lstFiles.BackColor = Color.White;
            _lstFiles.ForeColor = _colorLightText;
            _tabContent.BackColor = Color.White;
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
                LoadFilesAsync();
            }
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
                    MessageBox.Show($"加载文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                return;
            }

            foreach (var file in _allFiles.OrderBy(f => f.RelativePath))
            {
                // 显示相对路径作为文件名（包含目录结构）
                var item = new ListViewItem(file.RelativePath)
                {
                    Tag = file,
                    ToolTipText = file.FullPath,
                    ImageKey = file.Type == FileType.Csv ? "csv" : "excel"
                };
                _lstFiles.Items.Add(item);
            }
        }

        // ==================== 文件过滤 ====================
        private void OnFileFilterChanged(object? sender, EventArgs e)
        {
            ApplyFileFilter();
        }

        private void ApplyFileFilter()
        {
            var filter = _txtFileFilter.Text.Trim();

            // 重新加载列表
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
                    ImageKey = file.Type == FileType.Csv ? "csv" : "excel"
                };
                _lstFiles.Items.Add(item);
            }
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

            // 创建新的标签页
            var tabPage = new TabPage(Path.GetFileName(filePath))
            {
                Tag = filePath,
                ToolTipText = filePath
            };

            // 根据文件类型创建相应的查看控件
            Control viewer;
            if (fileType == FileType.Excel)
            {
                viewer = new ExcelViewerControl(filePath)
                {
                    Dock = DockStyle.Fill
                };
            }
            else
            {
                viewer = new CsvViewerControl(filePath)
                {
                    Dock = DockStyle.Fill
                };
            }

            tabPage.Controls.Add(viewer);
            _tabContent.TabPages.Add(tabPage);
            _tabContent.SelectedTab = tabPage;

            // 添加关闭按钮
            AddCloseButtonToTab(tabPage);
        }

        private void AddCloseButtonToTab(TabPage tabPage)
        {
            // 使用TabControl的MouseDown事件处理关闭
            _tabContent.MouseDown += (s, e) =>
            {
                for (int i = 0; i < _tabContent.TabPages.Count; i++)
                {
                    var tabRect = _tabContent.GetTabRect(i);
                    var closeRect = new Rectangle(tabRect.Right - 16, tabRect.Top + 4, 12, 12);

                    if (closeRect.Contains(e.Location))
                    {
                        _tabContent.TabPages.RemoveAt(i);
                        // 清除ListView选中状态，允许重新打开同一文件
                        _lstFiles.SelectedItems.Clear();
                        return;
                    }
                }
            };

            // 重绘标签页以显示关闭按钮
            _tabContent.DrawMode = TabDrawMode.OwnerDrawFixed;
            _tabContent.DrawItem += (s, e) =>
            {
                var tabPage = _tabContent.TabPages[e.Index];
                var tabRect = _tabContent.GetTabRect(e.Index);

                // 绘制背景
                bool isSelected = _tabContent.SelectedIndex == e.Index;
                using (var brush = new SolidBrush(isSelected ? Color.White : Color.FromArgb(240, 240, 240)))
                {
                    e.Graphics.FillRectangle(brush, tabRect);
                }

                // 绘制文本（增加文本区域宽度）
                var textRect = new Rectangle(tabRect.X + 4, tabRect.Y + 4, tabRect.Width - 24, tabRect.Height - 8);
                TextRenderer.DrawText(e.Graphics, tabPage.Text, _tabContent.Font, textRect, _colorLightText,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

                // 绘制关闭按钮
                var closeRect = new Rectangle(tabRect.Right - 16, tabRect.Top + 6, 10, 10);
                using (var pen = new Pen(_colorLightTextSecondary, 1.5f))
                {
                    e.Graphics.DrawLine(pen, closeRect.Left, closeRect.Top, closeRect.Right, closeRect.Bottom);
                    e.Graphics.DrawLine(pen, closeRect.Right, closeRect.Top, closeRect.Left, closeRect.Bottom);
                }

                // 绘制边框
                using (var pen = new Pen(_colorLightBorder))
                {
                    e.Graphics.DrawLine(pen, tabRect.Left, tabRect.Bottom - 1, tabRect.Right, tabRect.Bottom - 1);
                }
            };
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
        private TextBox _txtHeaderFilter = new();      // 表头过滤框
        private TextBox _txtContentFilter = new();     // 内容过滤框
        private Label _lblInfo = new();
        private Panel _pnlTop = new();
        private DataTable _dataTable = new();
        private DataTable _headerTable = new();        // 冻结的表头数据
        private string _filePath = "";
        private char _delimiter = ',';
        private int _frozenRows = 0;                   // 冻结行数

        // 配色
        private static readonly Color _colorLightBg = Color.FromArgb(250, 250, 250);
        private static readonly Color _colorLightText = Color.FromArgb(32, 33, 36);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(95, 99, 104);
        private static readonly Color _colorBtnBg = Color.FromArgb(0, 122, 204);

        public CsvViewerControl(string filePath)
        {
            _filePath = filePath;
            InitializeComponent();
            LoadCsvAsync();
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            BackColor = Color.White;            
            // 顶部工具栏 - 使用 FlowLayoutPanel 自动排列
            var flowTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = _colorLightBg,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(5, 5, 5, 5),
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
            // 从设置中读取默认冻结行数，默认为2
            _frozenRows = CsvSettings.DefaultFrozenRows;
            cboFrozenRows.SelectedIndex = _frozenRows;
            cboFrozenRows.SelectedIndexChanged += (s, e) =>
            {
                _frozenRows = cboFrozenRows.SelectedIndex;
                // 保存设置
                CsvSettings.DefaultFrozenRows = _frozenRows;
                CsvSettings.Save();
                RefreshGridLayout();
            };

            // 表头过滤框
            var lblHeaderFilter = new Label
            {
                Text = "表头过滤:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            var txtHeaderFilter = new TextBox
            {
                Size = new Size(160, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 10, 0)
            };
            txtHeaderFilter.TextChanged += (s, e) => ApplyFilters();
            txtHeaderFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    txtHeaderFilter.Text = "";
                    e.Handled = true;
                }
            };

            // 内容过滤框
            var lblContentFilter = new Label
            {
                Text = "内容过滤:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            var txtContentFilter = new TextBox
            {
                Size = new Size(160, 23),
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

            // 提示标签
            var lblHint = new Label
            {
                Text = "支持 || 或 &&",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 10, 0)
            };

            // 信息标签
            _lblInfo = new Label
            {
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 0, 0)
            };

            flowTop.Controls.AddRange(new Control[] { lblFrozenRows, cboFrozenRows, lblHeaderFilter, txtHeaderFilter, lblContentFilter, txtContentFilter, lblHint, _lblInfo });

            // 保存引用
            _pnlTop = flowTop;
            _txtHeaderFilter = txtHeaderFilter;
            _txtContentFilter = txtContentFilter;

            // 表格容器 - 使用SplitContainer来精确控制布局
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 28, // 初始只有列标题高度
                FixedPanel = FixedPanel.Panel1,
                IsSplitterFixed = true,
                BackColor = Color.White
            };

            // 冻结的表头表格（显示冻结的行）
            _dataGridHeader = CreateDataGridView();
            _dataGridHeader.Dock = DockStyle.Fill;
            _dataGridHeader.ScrollBars = ScrollBars.None; // 不显示滚动条
            _dataGridHeader.RowHeadersVisible = true; // 显示行头（用于对齐）
            _dataGridHeader.RowHeadersWidth = 50; // 与数据表格相同的行头宽度
            _dataGridHeader.ColumnHeadersVisible = true; // 显示列标题
            _dataGridHeader.ReadOnly = true;
            _dataGridHeader.AllowUserToResizeColumns = false; // 禁止调整列宽，由主表格控制
            _dataGridHeader.AllowUserToResizeRows = false;
            _dataGridHeader.RowPostPaint += OnFrozenRowPostPaint; // 绘制冻结行的行号
            _dataGridHeader.DefaultCellStyle.SelectionBackColor = _dataGridHeader.DefaultCellStyle.BackColor; // 禁用选中高亮
            _dataGridHeader.DefaultCellStyle.SelectionForeColor = _dataGridHeader.DefaultCellStyle.ForeColor;
            _dataGridHeader.CellDoubleClick += OnFrozenCellDoubleClick; // 双击编辑

            // 数据表格（主表格）
            _dataGrid = CreateDataGridView();
            _dataGrid.Dock = DockStyle.Fill;
            _dataGrid.RowHeadersWidth = 50;
            _dataGrid.RowPostPaint += OnRowPostPaint;
            _dataGrid.ColumnHeadersVisible = false; // 隐藏列标题（由冻结表格显示）

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
                BackgroundColor = Color.White,
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
                    BackColor = Color.FromArgb(240, 240, 240),
                    ForeColor = Color.FromArgb(32, 33, 36)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(248, 249, 250)
                }
            };
            dgv.ColumnHeaderMouseClick += OnColumnHeaderClick;
            dgv.CellFormatting += OnCellFormatting;
            dgv.CellDoubleClick += OnCellDoubleClick; // 双击编辑
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
                MessageBox.Show($"保存文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                var headerFilter = _txtHeaderFilter.Text.Trim();
                _visibleColumnIndices = GetVisibleColumnIndices(headerFilter);

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
                var visibleRows = GetVisibleRows(contentFilter, headerFilter);
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

        private void SyncColumnSettings()
        {
            // 同步两个表格的列宽和列标题
            for (int i = 0; i < Math.Min(_dataGrid.Columns.Count, _dataGridHeader.Columns.Count); i++)
            {
                _dataGridHeader.Columns[i].Width = _dataGrid.Columns[i].Width;
                _dataGridHeader.Columns[i].HeaderText = _dataGrid.Columns[i].HeaderText;
                _dataGridHeader.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
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
                        MessageBox.Show($"加载CSV文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            }
        }

        private void LoadCsvFile()
        {
            if (!File.Exists(_filePath))
                throw new FileNotFoundException("CSV文件不存在", _filePath);

            // 读取文件内容以检测分隔符
            string content;
            using (var reader = new StreamReader(_filePath, Encoding.UTF8, true))
            {
                content = reader.ReadToEnd();
            }

            // 检测分隔符
            _delimiter = DetectDelimiter(content);

            // 解析CSV
            var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0)
                throw new InvalidDataException("CSV文件为空");

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
            for (int i = 0; i < lines.Length; i++)
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
            var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return ',';

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

        /// <summary>获取可见的列索引（第一列始终可见，保持原始顺序）</summary>
        private List<int> GetVisibleColumnIndices(string headerFilter)
        {
            var visibleIndices = new List<int>();

            // 第一列始终可见
            if (_dataTable.Columns.Count > 0)
            {
                visibleIndices.Add(0);
            }

            // 如果没有表头过滤，所有列都可见（按原始顺序）
            if (string.IsNullOrEmpty(headerFilter))
            {
                for (int i = 1; i < _dataTable.Columns.Count; i++)
                {
                    visibleIndices.Add(i);
                }
                return visibleIndices;
            }

            // 处理运算符
            var processedFilter = headerFilter.Replace("||", " | ").Replace("&&", " & ");

            // 检查每一列（从第2列开始，第1列始终可见）
            // 按原始顺序检查，保持列的顺序
            for (int i = 1; i < _dataTable.Columns.Count; i++)
            {
                bool isMatch = false;

                // 检查列名是否匹配
                var colName = _dataTable.Columns[i].ColumnName;
                if (MatchesFilter(colName, processedFilter))
                {
                    isMatch = true;
                }

                // 检查数据表前N行（冻结行）的数据是否匹配
                if (!isMatch && _frozenRows > 0)
                {
                    for (int row = 0; row < Math.Min(_frozenRows, _dataTable.Rows.Count); row++)
                    {
                        var cellValue = _dataTable.Rows[row][i]?.ToString() ?? "";
                        if (MatchesFilter(cellValue, processedFilter))
                        {
                            isMatch = true;
                            break;
                        }
                    }
                }

                // 如果匹配或者是第一列，添加到可见列表（保持原始顺序）
                if (isMatch)
                {
                    visibleIndices.Add(i);
                }
            }

            return visibleIndices;
        }

        /// <summary>获取可见的行（从冻结行之后开始）</summary>
        private List<DataRow> GetVisibleRows(string contentFilter, string headerFilter)
        {
            var visibleRows = new List<DataRow>();

            // 从冻结行之后开始遍历
            for (int i = _frozenRows; i < _dataTable.Rows.Count; i++)
            {
                var row = _dataTable.Rows[i];

                // 如果没有内容过滤，所有行都可见
                if (string.IsNullOrEmpty(contentFilter))
                {
                    visibleRows.Add(row);
                    continue;
                }

                // 处理运算符
                var processedFilter = contentFilter.Replace("||", " | ").Replace("&&", " & ");

                // 检查该行是否满足内容过滤条件
                bool isMatch = false;

                // 如果有表头过滤，只在可见列中搜索
                var columnsToSearch = string.IsNullOrEmpty(headerFilter)
                    ? Enumerable.Range(0, _dataTable.Columns.Count).ToList()
                    : _visibleColumnIndices;

                foreach (var colIndex in columnsToSearch)
                {
                    var cellValue = row[colIndex]?.ToString() ?? "";
                    if (MatchesFilter(cellValue, processedFilter))
                    {
                        isMatch = true;
                        break;
                    }
                }

                if (isMatch)
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

        private string BuildRowFilterExpression(string filter)
        {
            if (string.IsNullOrEmpty(filter))
                return "";

            var conditions = new List<string>();

            if (filter.Contains('|'))
            {
                // OR 条件
                var terms = filter.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var term in terms)
                {
                    conditions.Add(BuildSingleCondition(term.Trim()));
                }
                return string.Join(" OR ", conditions);
            }
            else if (filter.Contains('&'))
            {
                // AND 条件
                var terms = filter.Split('&', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var term in terms)
                {
                    conditions.Add(BuildSingleCondition(term.Trim()));
                }
                return string.Join(" AND ", conditions);
            }
            else
            {
                return BuildSingleCondition(filter);
            }
        }

        private string BuildSingleCondition(string term)
        {
            // 对所有列进行OR搜索
            var columnConditions = new List<string>();
            foreach (DataColumn col in _dataTable.Columns)
            {
                columnConditions.Add($"CONVERT([{col.ColumnName}], 'System.String') LIKE '%{term.Replace("'", "''")}%'");
            }
            return $"({string.Join(" OR ", columnConditions)})";
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
    public class NaturalStringComparer : IComparer<string>
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
        private TextBox _txtValue = new();
        private Button _btnSave = new();
        private Button _btnCancel = new();

        public string CellValue => _txtValue.Text;

        public CellEditDialog(string initialValue)
        {
            InitializeComponent();
            _txtValue.Text = initialValue;
            _txtValue.SelectAll();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "编辑单元格";
            ClientSize = new Size(400, 150);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.White;

            // 文本框
            _txtValue = new TextBox
            {
                Location = new Point(12, 20),
                Size = new Size(360, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            // 保存按钮
            _btnSave = new Button
            {
                Text = "保存",
                Location = new Point(220, 80),
                Size = new Size(70, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK
            };
            _btnSave.FlatAppearance.BorderSize = 0;

            // 取消按钮
            _btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(302, 80),
                Size = new Size(70, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(32, 33, 36),
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel
            };
            _btnCancel.FlatAppearance.BorderSize = 0;

            Controls.Add(_txtValue);
            Controls.Add(_btnSave);
            Controls.Add(_btnCancel);

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
        private TextBox _txtHeaderFilter = new();
        private TextBox _txtContentFilter = new();
        private Label _lblInfo = new();
        private Panel _pnlTop = new();
        private DataTable _dataTable = new();
        private DataTable _headerTable = new();
        private string _filePath = "";
        private int _frozenRows = 0;
        private List<int> _visibleColumnIndices = new();

        // 配色
        private static readonly Color _colorLightBg = Color.FromArgb(250, 250, 250);
        private static readonly Color _colorLightText = Color.FromArgb(32, 33, 36);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(95, 99, 104);
        private static readonly Color _colorBtnBg = Color.FromArgb(0, 122, 204);

        public ExcelViewerControl(string filePath)
        {
            _filePath = filePath;
            InitializeComponent();
            LoadExcelAsync();
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            BackColor = Color.White;

            // 顶部工具栏 - 使用 FlowLayoutPanel 自动排列
            var flowTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = _colorLightBg,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(5, 5, 5, 5),
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
            _frozenRows = CsvSettings.DefaultFrozenRows;
            cboFrozenRows.SelectedIndex = _frozenRows;
            cboFrozenRows.SelectedIndexChanged += (s, e) =>
            {
                _frozenRows = cboFrozenRows.SelectedIndex;
                CsvSettings.DefaultFrozenRows = _frozenRows;
                CsvSettings.Save();
                RefreshGridLayout();
            };

            // 表头过滤框
            var lblHeaderFilter = new Label
            {
                Text = "表头过滤:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            var txtHeaderFilter = new TextBox
            {
                Size = new Size(160, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 10, 0)
            };
            txtHeaderFilter.TextChanged += (s, e) => ApplyFilters();
            txtHeaderFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    txtHeaderFilter.Text = "";
                    e.Handled = true;
                }
            };

            // 内容过滤框
            var lblContentFilter = new Label
            {
                Text = "内容过滤:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };

            var txtContentFilter = new TextBox
            {
                Size = new Size(160, 23),
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

            // 提示标签
            var lblHint = new Label
            {
                Text = "支持 || 或 &&",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 10, 0)
            };

            // 信息标签
            _lblInfo = new Label
            {
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 0, 0)
            };

            flowTop.Controls.AddRange(new Control[] { lblFrozenRows, cboFrozenRows, lblHeaderFilter, txtHeaderFilter, lblContentFilter, txtContentFilter, lblHint, _lblInfo });

            _pnlTop = flowTop;
            _txtHeaderFilter = txtHeaderFilter;
            _txtContentFilter = txtContentFilter;

            // 表格容器 - 使用SplitContainer来精确控制布局
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 28,
                FixedPanel = FixedPanel.Panel1,
                IsSplitterFixed = true,
                BackColor = Color.White
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
                BackgroundColor = Color.White,
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
                    BackColor = Color.FromArgb(240, 240, 240),
                    ForeColor = Color.FromArgb(32, 33, 36)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(248, 249, 250)
                }
            };
            dgv.ColumnHeaderMouseClick += OnColumnHeaderClick;
            dgv.CellFormatting += OnCellFormatting;
            dgv.CellDoubleClick += OnCellDoubleClick; // 双击编辑
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
                        MessageBox.Show($"加载Excel文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            }
        }

        private void LoadExcelFile()
        {
            if (!System.IO.File.Exists(_filePath))
                throw new FileNotFoundException("Excel文件不存在", _filePath);

            using var stream = System.IO.File.Open(_filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();

            if (result.Tables.Count == 0)
                throw new InvalidDataException("Excel文件没有工作表");

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
                var headerFilter = _txtHeaderFilter.Text.Trim();
                _visibleColumnIndices = GetVisibleColumnIndices(headerFilter);

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
                var visibleRows = GetVisibleRows(contentFilter, headerFilter);
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

        private List<int> GetVisibleColumnIndices(string headerFilter)
        {
            var visibleIndices = new List<int>();

            if (_dataTable.Columns.Count > 0)
            {
                visibleIndices.Add(0);
            }

            if (string.IsNullOrEmpty(headerFilter))
            {
                for (int i = 1; i < _dataTable.Columns.Count; i++)
                {
                    visibleIndices.Add(i);
                }
                return visibleIndices;
            }

            var processedFilter = headerFilter.Replace("||", " | ").Replace("&&", " & ");

            for (int i = 1; i < _dataTable.Columns.Count; i++)
            {
                bool isMatch = false;

                var colName = _dataTable.Columns[i].ColumnName;
                if (MatchesFilter(colName, processedFilter))
                {
                    isMatch = true;
                }

                if (!isMatch && _frozenRows > 0)
                {
                    for (int row = 0; row < Math.Min(_frozenRows, _dataTable.Rows.Count); row++)
                    {
                        var cellValue = _dataTable.Rows[row][i]?.ToString() ?? "";
                        if (MatchesFilter(cellValue, processedFilter))
                        {
                            isMatch = true;
                            break;
                        }
                    }
                }

                if (isMatch)
                {
                    visibleIndices.Add(i);
                }
            }

            return visibleIndices;
        }

        private List<DataRow> GetVisibleRows(string contentFilter, string headerFilter)
        {
            var visibleRows = new List<DataRow>();

            for (int i = _frozenRows; i < _dataTable.Rows.Count; i++)
            {
                var row = _dataTable.Rows[i];

                if (string.IsNullOrEmpty(contentFilter))
                {
                    visibleRows.Add(row);
                    continue;
                }

                var processedFilter = contentFilter.Replace("||", " | ").Replace("&&", " & ");
                bool isMatch = false;

                var columnsToSearch = string.IsNullOrEmpty(headerFilter)
                    ? Enumerable.Range(0, _dataTable.Columns.Count).ToList()
                    : _visibleColumnIndices;

                foreach (var colIndex in columnsToSearch)
                {
                    var cellValue = row[colIndex]?.ToString() ?? "";
                    if (MatchesFilter(cellValue, processedFilter))
                    {
                        isMatch = true;
                        break;
                    }
                }

                if (isMatch)
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
}
