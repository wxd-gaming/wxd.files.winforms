using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;

namespace wxd.files.winforms
{
    /// <summary>版本对比窗体：顶部历史记录表格 + 底部左右 TableBrowserControl 对比表格</summary>
    public partial class VersionCompareForm : Form
    {
        // ==================== 控件 ====================
        private DataGridView _dgvHistory;
        private CheckBox _chkDiffOnly;
        private CheckBox _chkAlwaysLatest;
        private Label _lblInfo;
        private SplitContainer _splitMain;
        private SplitContainer _splitCompare;
        private Panel _pnlTop;
        private Label _lblLeftHeader;
        private Label _lblRightHeader;
        private LoadingForm? _loadingForm;

        // ==================== 数据 ====================
        private string _filePath;
        private VcsType _vcsType;
        private List<VersionHistoryEntry> _history = new();
        private string _tempDir;
        private int _selectedIndex = 0;
        private int _frozenRows;
        private int _headerRowIndex;
        private bool[]? _columnVisibility;

        // 左右两个 TableBrowserControl（复用，不每次创建销毁）
        private TableBrowserControl? _leftView;
        private TableBrowserControl? _rightView;
        private string? _leftRevision;
        private string? _rightRevision;

        // 加载守卫
        private bool _isLoading;

        // DataTable 缓存：key = revision, value = DataTable
        private readonly Dictionary<string, DataTable> _dataCache = new();

        // 差异数据 —— 用行号集合记录原始DataTable中的差异行索引
        // 保留这些字段用于恢复全部显示（保存原始完整 FileTableData）
        private FileTableData? _leftFullFileData;
        private FileTableData? _rightFullFileData;

        // ★ 差异行原始行号集合（0-based，对应 DataTable.Rows 的索引，包含固定行）
        private HashSet<int> _leftDiffRows = new();
        private HashSet<int> _rightDiffRows = new();
        private HashSet<int> _leftDiffCells = new();    // (row << 16) | col 编码的差异单元格集合（左）
        private HashSet<int> _rightDiffCells = new();   // (row << 16) | col 编码的差异单元格集合（右）

        // ★ DGV 行索引 → 原始完整 DataTable 行索引（包含固定行）的映射
        private List<int> _leftRowToOriginalIndex = new();
        private List<int> _rightRowToOriginalIndex = new();

        // ★ DGV 列索引 → 原始完整 DataTable 列索引的映射（解决列隐藏后列号转换问题）
        private List<int> _leftColToOriginalIndex = new();
        private List<int> _rightColToOriginalIndex = new();

        // ★ 完整列数的列隐藏状态（阶段3.5会覆盖 _columnCheckStates 为投影后列数，需要独立保存）
        private bool[]? _savedLeftColStates = null;
        private bool[]? _savedRightColStates = null;

        private int _diffRowCount = 0;
        private bool _diffComputed = false;
        private bool _syncingScroll = false;

        // 配色
        private static readonly Color _colorSurface = Color.White;
        private static readonly Color _colorDiffCellBg = Color.FromArgb(255, 200, 200);   // 差异单元格：红色背景
        private static readonly Color _colorDiffRowBg = Color.FromArgb(255, 253, 220);    // 差异行：浅黄色背景
        private static readonly Color _colorHeaderBg = Color.FromArgb(240, 246, 252);
        private static readonly Color _colorLightText = Color.FromArgb(20, 60, 100);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(90, 130, 170);
        private static readonly Color _colorSelectedBg = Color.FromArgb(227, 242, 253);

        // ==================== 构造函数 ====================
        public VersionCompareForm(string filePath, VcsType vcsType,
            int frozenRows, int headerRowIndex,
            string? rootDir = null,
            bool[]? columnVisibility = null)
        {
            _filePath = filePath;
            _vcsType = vcsType;

            // ✅ 优先使用调用方传入的值（调用方已正确获取配置）
            _frozenRows = frozenRows;
            _headerRowIndex = headerRowIndex;
            _columnVisibility = columnVisibility;

            // 只有调用方未传入列可见性时，才从配置回退读取（兼容旧调用）
            if (_columnVisibility == null)
            {
                var dir = rootDir ?? Path.GetDirectoryName(filePath) ?? "";
                var ext = Path.GetExtension(filePath).ToLowerInvariant();
                string relPath;
                if (!string.IsNullOrEmpty(dir) && filePath.StartsWith(dir, StringComparison.OrdinalIgnoreCase))
                    relPath = filePath.Substring(dir.Length).TrimStart(Path.DirectorySeparatorChar);
                else
                    relPath = Path.GetFileName(filePath);
                _columnVisibility = CsvSettings.GetColumnVisibility(dir, relPath);
            }

            _tempDir = Path.Combine(Path.GetTempPath(),
                $"wxd_vcs_diff_{Guid.NewGuid():N}");
            Directory.CreateDirectory(_tempDir);

            InitializeComponent();
            _ = LoadHistoryAsync();

            FormClosed += (s, e) => CleanupTempFiles();
        }

        private void InitializeComponent()
        {
            Text = $"版本对比 - {Path.GetFileName(_filePath)} " +
                   $"[{(_vcsType == VcsType.Svn ? "SVN" : "Git")}]";
            ClientSize = new Size(1400, 900);
            WindowState = FormWindowState.Maximized;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = _colorSurface;
            KeyPreview = true;
            KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                    Close();
            };

            // === 顶部控制面板 ===
            _pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = Color.FromArgb(245, 250, 255),
                Padding = new Padding(8, 6, 8, 4)
            };

            _chkDiffOnly = new CheckBox
            {
                Text = "仅显示差异行",
                AutoSize = true,
                Location = new Point(8, 8),
                ForeColor = _colorLightText,
                Font = new Font("Microsoft YaHei UI", 9f)
            };
            _chkDiffOnly.CheckedChanged += (s, e) => RefreshDiffView();

            _chkAlwaysLatest = new CheckBox
            {
                Text = "右侧永远显示最新版本",
                AutoSize = true,
                Location = new Point(140, 8),
                ForeColor = _colorLightText,
                Font = new Font("Microsoft YaHei UI", 9f)
            };
            _chkAlwaysLatest.CheckedChanged += async (s, e) =>
                await LoadComparisonAsync(_selectedIndex);

            _lblInfo = new Label
            {
                AutoSize = true,
                Location = new Point(340, 10),
                ForeColor = _colorLightTextSecondary,
                Font = new Font("Microsoft YaHei UI", 9f)
            };

            _pnlTop.Controls.AddRange(new Control[] {
                _chkDiffOnly, _chkAlwaysLatest, _lblInfo });

            // === 上下分割: 历史表格 / 对比区域 ===
            _splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 160,
                Panel1MinSize = 120,
                FixedPanel = FixedPanel.Panel1,
                BackColor = Color.FromArgb(200, 225, 245),
                SplitterWidth = 2
            };

            _dgvHistory = CreateHistoryGrid();
            _splitMain.Panel1.Controls.Add(_dgvHistory);

            // === 对比区域：左右标签 + TableBrowserControl ===
            _splitCompare = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 50,
                BackColor = Color.FromArgb(200, 225, 245),
                SplitterWidth = 3
            };

            // 左侧标题 + 控件占位（用 TableLayoutPanel 确保不重叠）
            _lblLeftHeader = new Label
            {
                Text = "旧版本",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(255, 245, 245),
                ForeColor = Color.FromArgb(200, 50, 50),
                Font = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold),
                BorderStyle = BorderStyle.FixedSingle
            };
            var leftPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            leftPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            leftPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            leftPanel.Controls.Add(_lblLeftHeader, 0, 0);
            _splitCompare.Panel1.Controls.Add(leftPanel);

            // 右侧标题 + 控件占位
            _lblRightHeader = new Label
            {
                Text = "新版本",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(245, 255, 245),
                ForeColor = Color.FromArgb(50, 160, 50),
                Font = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold),
                BorderStyle = BorderStyle.FixedSingle
            };
            var rightPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            rightPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            rightPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            rightPanel.Controls.Add(_lblRightHeader, 0, 0);
            _splitCompare.Panel2.Controls.Add(rightPanel);

            _splitMain.Panel2.Controls.Add(_splitCompare);

            Controls.Add(_splitMain);
            Controls.Add(_pnlTop);

            // 调整左右对比各占 50%，历史列表固定高度
            Shown += (s, e) =>
            {
                _splitMain.SplitterDistance = 160;
                _splitCompare.SplitterDistance = _splitCompare.Width / 2;
            };
            _splitCompare.Resize += (s, e) =>
            {
                if (_leftView != null || _rightView != null)
                    _splitCompare.SplitterDistance = _splitCompare.Width / 2;
            };
        }

        private DataGridView CreateHistoryGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = _colorSurface,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeight = 30,
                RowTemplate = new DataGridViewRow { Height = 24 },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _colorHeaderBg,
                    ForeColor = _colorLightText,
                    Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold)
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Font = new Font("Microsoft YaHei UI", 9f),
                    ForeColor = _colorLightText,
                    SelectionBackColor = _colorSelectedBg,
                    SelectionForeColor = _colorLightText
                }
            };

            dgv.Columns.Add("Revision", "版本号");
            dgv.Columns.Add("Date", "提交时间");
            dgv.Columns.Add("Author", "提交人");
            dgv.Columns.Add("Message", "提交信息");

            dgv.Columns["Revision"].FillWeight = 15;
            dgv.Columns["Date"].FillWeight = 25;
            dgv.Columns["Author"].FillWeight = 15;
            dgv.Columns["Message"].FillWeight = 45;

            dgv.SelectionChanged += OnHistorySelectionChanged;

            return dgv;
        }

        // ==================== Loading 对话框 ====================
        private Thread? _loadingThread;

        private void ShowLoading(string? detail = null)
        {
            if (_loadingForm != null || _loadingThread != null) return;

            // ★ 在独立 STA 线程中运行 LoadingForm，拥有自己的消息泵
            //    这样即使主 UI 线程被同步操作阻塞，旋转动画也能持续转动
            var detailCopy = detail;
            var ownerHandle = this.Handle; // 在主线程获取句柄
            var ownerLocation = this.Location;
            var ownerSize = this.Size;

            _loadingThread = new Thread(() =>
            {
                _loadingForm = new LoadingForm();
                if (!string.IsNullOrEmpty(detailCopy))
                    _loadingForm.SetDetail(detailCopy);

                // 居中于父窗体
                _loadingForm.Load += (s, e) =>
                {
                    _loadingForm.Location = new Point(
                        ownerLocation.X + (ownerSize.Width - _loadingForm.Width) / 2,
                        ownerLocation.Y + (ownerSize.Height - _loadingForm.Height) / 2);
                };

                _loadingForm.FormClosed += (s, e) =>
                {
                    _loadingForm = null;
                };

                Application.Run(_loadingForm);
            });
            _loadingThread.SetApartmentState(ApartmentState.STA);
            _loadingThread.IsBackground = true;
            _loadingThread.Start();

            // 禁用父窗体，阻止用户交互
            Enabled = false;
        }

        private void HideLoading()
        {
            // 恢复父窗体交互
            Enabled = true;

            // ★ 跨线程关闭 LoadingForm
            _loadingForm?.BeginInvoke(new Action(() =>
            {
                _loadingForm?.Close();
            }));

            _loadingThread?.Join(1000); // 等待线程结束
            _loadingForm?.Dispose();
            _loadingForm = null;
            _loadingThread = null;
        }

        /// <summary>独立的 Loading 提示窗体（旋转圆弧动画 + 文字）</summary>
        private class LoadingForm : Form
        {
            private readonly System.Windows.Forms.Timer _animTimer;
            private int _angle;
            private readonly Label _lblTitle;
            private readonly Label _lblDetail;

            public LoadingForm()
            {
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;
                ShowInTaskbar = false;
                Size = new Size(340, 160);
                BackColor = Color.White;
                TopMost = true;

                // 双缓冲减少闪烁
                SetStyle(ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.UserPaint |
                         ControlStyles.DoubleBuffer |
                         ControlStyles.ResizeRedraw, true);
                UpdateStyles();

                // 标题文字
                _lblTitle = new Label
                {
                    Text = "正在加载对比数据...",
                    AutoSize = false,
                    Size = new Size(300, 28),
                    Location = new Point(20, 95),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Microsoft YaHei UI", 12f, FontStyle.Bold),
                    ForeColor = Color.FromArgb(30, 136, 229)
                };

                // 详细信息（版本号等，可动态更新）
                _lblDetail = new Label
                {
                    Text = "",
                    AutoSize = false,
                    Size = new Size(300, 20),
                    Location = new Point(20, 125),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Microsoft YaHei UI", 9f),
                    ForeColor = Color.FromArgb(120, 160, 200)
                };

                Controls.Add(_lblTitle);
                Controls.Add(_lblDetail);

                // 绘制圆角边框 + 旋转圆弧
                Paint += OnPaintAll;

                // 旋转动画：~60 FPS
                _animTimer = new System.Windows.Forms.Timer { Interval = 16 };
                _animTimer.Tick += OnAnimTick;
            }

            /// <summary>更新详细信息文字</summary>
            public void SetDetail(string text)
            {
                if (_lblDetail.InvokeRequired)
                {
                    _lblDetail.Invoke(() => _lblDetail.Text = text);
                    return;
                }
                _lblDetail.Text = text;
            }

            protected override void OnShown(EventArgs e)
            {
                base.OnShown(e);
                // 居中于父窗体
                if (Owner != null)
                {
                    Location = new Point(
                        Owner.Left + (Owner.Width - Width) / 2,
                        Owner.Top + (Owner.Height - Height) / 2);
                }
                _animTimer.Start();
            }

            protected override void OnFormClosed(FormClosedEventArgs e)
            {
                _animTimer.Stop();
                _animTimer.Dispose();
                base.OnFormClosed(e);
            }

            private void OnAnimTick(object? sender, EventArgs e)
            {
                _angle = (_angle + 6) % 360;
                Invalidate(new Rectangle(20, 15, 72, 72));
            }

            /// <summary>统一绘制：圆角边框 + 旋转圆弧</summary>
            private void OnPaintAll(object? sender, PaintEventArgs e)
            {
                e.Graphics.SmoothingMode =
                    System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                // 圆角矩形边框
                var rect = new Rectangle(0, 0, Width - 1, Height - 1);
                using var pen = new Pen(Color.FromArgb(200, 210, 225), 1f);
                using var path = GetRoundedRectPath(rect, 8);
                e.Graphics.DrawPath(pen, path);

                // 旋转圆弧
                DrawSpinner(e.Graphics);
            }

            private void DrawSpinner(Graphics g)
            {
                int cx = 56, cy = 50, radius = 22;
                float penWidth = 3.5f;

                // 背景圆环
                using var bgPen = new Pen(Color.FromArgb(230, 235, 240), penWidth);
                g.DrawEllipse(bgPen, cx - radius, cy - radius, radius * 2, radius * 2);

                // 旋转蓝色弧线
                using var arcPen = new Pen(Color.FromArgb(30, 136, 229), penWidth)
                {
                    StartCap = System.Drawing.Drawing2D.LineCap.Round,
                    EndCap = System.Drawing.Drawing2D.LineCap.Round
                };
                g.DrawArc(arcPen,
                    cx - radius, cy - radius, radius * 2, radius * 2,
                    _angle, 270);
            }

            private static System.Drawing.Drawing2D.GraphicsPath GetRoundedRectPath(
                Rectangle rect, int radius)
            {
                var path = new System.Drawing.Drawing2D.GraphicsPath();
                int d = radius * 2;
                path.AddArc(rect.X, rect.Y, d, d, 180, 90);
                path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
                path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
                path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }
        }

        // ==================== 加载历史 ====================
        private async Task LoadHistoryAsync()
        {
            try
            {
                _lblInfo.Text = "正在加载历史记录...";

                _history = _vcsType == VcsType.Svn
                    ? await VersionControlHelper.GetSvnHistoryAsync(_filePath)
                    : await VersionControlHelper.GetGitHistoryAsync(_filePath);

                if (_history.Count == 0)
                {
                    _lblInfo.Text = "未找到历史记录";
                    return;
                }

                var table = new DataTable();
                table.Columns.Add("Revision", typeof(string));
                table.Columns.Add("Date", typeof(string));
                table.Columns.Add("Author", typeof(string));
                table.Columns.Add("Message", typeof(string));

                foreach (var entry in _history)
                {
                    table.Rows.Add(
                        entry.Revision,
                        entry.CommitDate.ToString("yyyy-MM-dd HH:mm"),
                        entry.Author,
                        entry.Message.Length > 80
                            ? entry.Message.Substring(0, 77) + "..."
                            : entry.Message
                    );
                }

                _dgvHistory.DataSource = table;
                _lblInfo.Text = $"共 {_history.Count} 条历史记录";

                if (_dgvHistory.Rows.Count > 0)
                {
                    // 先移除事件，手动设置选择后再恢复，避免重复触发
                    _dgvHistory.SelectionChanged -= OnHistorySelectionChanged;
                    _selectedIndex = 0;
                    _dgvHistory.Rows[0].Selected = true;
                    _dgvHistory.SelectionChanged += OnHistorySelectionChanged;

                    // 显式触发加载
                    _ = LoadComparisonAsync(0);
                }
            }
            catch (Exception ex)
            {
                _lblInfo.Text = $"加载失败: {ex.Message}";
            }
        }

        // ==================== 历史选择变更 ====================
        private async void OnHistorySelectionChanged(object? sender, EventArgs e)
        {
            if (_dgvHistory.SelectedRows.Count == 0) return;
            var newIndex = _dgvHistory.SelectedRows[0].Index;
            if (newIndex == _selectedIndex) return;
            _selectedIndex = newIndex;
            await LoadComparisonAsync(_selectedIndex);
        }

        // ==================== 加载对比 ====================
        /// <summary>
        /// 版本对比逻辑：
        /// - 左边 = 选中的提交记录
        /// - 右边 = 更新的版本（如果有下一个提交记录则用下一个，否则用当前目录文件）
        /// - 勾选"右侧永远最新"时，右边始终用当前目录文件（最新）
        /// </summary>
        private async Task LoadComparisonAsync(int selectedIndex)
        {
            // 加载中不允许切换
            if (_isLoading) return;
            if (selectedIndex >= _history.Count) return;

            var selected = _history[selectedIndex];

            // 左边：选中的提交记录
            string leftRevision = selected.FullRevision;
            string leftLabel = $"r{selected.Revision} ({selected.CommitDate:yyyy-MM-dd HH:mm})";

            // 右边：更新的版本
            string rightRevision;
            string rightLabel;

            if (_chkAlwaysLatest.Checked)
            {
                // 勾选"右侧永远最新"：右边用当前目录文件
                rightRevision = "";  // 空字符串表示使用当前文件
                rightLabel = "当前文件 (最新)";
            }
            else if (selectedIndex > 0)
            {
                // 有更新的版本（selectedIndex=0 是最新，没有更新的）
                var newer = _history[selectedIndex - 1];
                rightRevision = newer.FullRevision;
                rightLabel = $"r{newer.Revision} ({newer.CommitDate:yyyy-MM-dd HH:mm})";
            }
            else
            {
                // 已是最新版本，右边用当前目录文件
                rightRevision = "";
                rightLabel = "当前文件 (最新)";
            }

            // 去重：和当前一样则不重新加载
            if (_leftRevision == leftRevision && _rightRevision == rightRevision)
                return;

            _isLoading = true;

            // ★ 阶段0：先从界面中移除 DataGridView 控件，防止设置 DataSource 时逐行绘制
            EnsureViewsCreated();
            _leftView?.DetachGridFromParent();
            _rightView?.DetachGridFromParent();

            ShowLoading($"{leftLabel} ← → {rightLabel}");
            // 给消息循环一点时间处理，确保 LoadingForm 被绘制
            await Task.Delay(50);

            try
            {
                _leftRevision = leftRevision;
                _rightRevision = rightRevision;

                _lblInfo.Text = $"正在加载对比: {leftLabel} ← → {rightLabel}";
                _lblLeftHeader.Text = $"旧版本 - {leftLabel}";
                _lblRightHeader.Text = $"新版本 - {rightLabel}";

                // 阶段1：加载 DataTable（从缓存或导出文件）—— 后台并行执行
                DataTable? leftData = null;
                DataTable? rightData = null;

                var loadTasks = new List<Task>();
                // 左边：从 SVN/Git 导出选中版本
                loadTasks.Add(Task.Run(async () =>
                {
                    leftData = await LoadDataTableAsync(leftRevision);
                }));
                // 右边：导出更新版本，或读取当前文件
                loadTasks.Add(Task.Run(async () =>
                {
                    rightData = await LoadDataTableAsync(rightRevision);
                }));
                await Task.WhenAll(loadTasks);

                // 阶段2：更新 TableBrowserControl 数据（控件已移除，不会逐行绘制）
                // ★ 在 SetDataTable 之前，用完整列数的保存状态注入 _externalColumnStates
                // （_externalColumnStates 在第一次 SetDataTable 后被清空；阶段3.5会覆盖 _columnCheckStates，
                //   所以不能依赖 GetColumnCheckStates()，需要独立保存完整列数的状态）
                if (_leftView != null && _savedLeftColStates != null)
                    _leftView.SetColumnVisibility(_savedLeftColStates);
                else if (_leftView != null && _columnVisibility != null)
                    _leftView.SetColumnVisibility(_columnVisibility);

                if (_rightView != null && _savedRightColStates != null)
                    _rightView.SetColumnVisibility(_savedRightColStates);
                else if (_rightView != null && _columnVisibility != null)
                    _rightView.SetColumnVisibility(_columnVisibility);

                if (leftData != null)
                    _leftView!.SetDataTable(leftData,
                        $"旧版本 {leftLabel} | 行数: {leftData.Rows.Count}");
                else if (_leftView != null)
                    _leftView.SetDataTable(new DataTable(), "无数据");

                if (rightData != null)
                    _rightView!.SetDataTable(rightData,
                        $"新版本 {rightLabel} | 行数: {rightData.Rows.Count}");
                else if (_rightView != null)
                    _rightView.SetDataTable(new DataTable(), "无数据");

                // ★ 保存原始 FileTableData（用于筛选后恢复全部显示）
                _leftFullFileData = _leftView?.GetFileData();
                _rightFullFileData = _rightView?.GetFileData();

                // ★ 保存完整列数的列隐藏状态（阶段3.5会覆盖 _columnCheckStates，需要独立保存）
                _savedLeftColStates = _leftView?.GetColumnCheckStates();
                _savedRightColStates = _rightView?.GetColumnCheckStates();

                _splitCompare.SplitterDistance = _splitCompare.Width / 2;

                // 阶段3：后台计算差异
                _loadingForm?.SetDetail("正在计算差异...");
                await ComputeDiffAsync(leftData, rightData);

                // ★ 阶段3.1：将差异结果回写到 _leftFullFileData / _rightFullFileData 的 TableRow 中
                // （BuildProjectedFileData 在 diffOnly:true 时依赖 TableRow.IsDiffRow 判断，必须回写）
                if (_leftFullFileData != null)
                {
                    foreach (var row in _leftFullFileData.Rows)
                    {
                        row.IsDiffRow = _leftDiffRows.Contains(row.OriginalRowIndex);
                    }
                }
                if (_rightFullFileData != null)
                {
                    foreach (var row in _rightFullFileData.Rows)
                    {
                        row.IsDiffRow = _rightDiffRows.Contains(row.OriginalRowIndex);
                    }
                }

                // 阶段3.5：★ 根据当前列显隐状态重新构建投影后的 FileTableData → DataTable
                // ★ 先保存当前列显隐状态（基于完整列数），后续 SetDataTable 会重置 _columnCheckStates
                var leftColStatesBefore = _leftView?.GetColumnCheckStates();
                var rightColStatesBefore = _rightView?.GetColumnCheckStates();
                var leftVisibleCols = _leftView?.GetVisibleColumnIndices() ?? new List<int> { };
                var rightVisibleCols = _rightView?.GetVisibleColumnIndices() ?? new List<int> { };

                if (_leftFullFileData != null && leftVisibleCols.Count > 0 && _leftView != null)
                {
                    var leftProjected = BuildProjectedFileData(_leftFullFileData, leftVisibleCols, diffOnly: false);
                    _leftView.SetColumnVisibility(leftColStatesBefore);
                    _leftView.SetDataTable(BuildDataTableFromFileData(leftProjected), _lblLeftHeader.Text);
                    // ★ 构建行/列索引映射（DGV行/列→原始行/列），不再依赖 Tag 或 GetVisibleColumnIndices
                    _leftRowToOriginalIndex = BuildRowIndexMap(leftProjected);
                    _leftColToOriginalIndex = new List<int>(leftVisibleCols);
                }
                if (_rightFullFileData != null && rightVisibleCols.Count > 0 && _rightView != null)
                {
                    var rightProjected = BuildProjectedFileData(_rightFullFileData, rightVisibleCols, diffOnly: false);
                    _rightView.SetColumnVisibility(rightColStatesBefore);
                    _rightView.SetDataTable(BuildDataTableFromFileData(rightProjected), _lblRightHeader.Text);
                    _rightRowToOriginalIndex = BuildRowIndexMap(rightProjected);
                    _rightColToOriginalIndex = new List<int>(rightVisibleCols);
                }

                // 阶段4：应用高亮 + 差异行筛选
                ApplyDiffHighlight(_leftView?.GetDataGridView(), true);
                ApplyDiffHighlight(_rightView?.GetDataGridView(), false);

                // ★ 如果勾选了"仅显示差异行"，应用筛选
                if (_chkDiffOnly.Checked)
                {
                    if (_leftFullFileData != null && _leftView != null && leftVisibleCols.Count > 0)
                    {
                        var leftFiltered = BuildProjectedFileData(_leftFullFileData, leftVisibleCols, diffOnly: true);
                        _leftView.SetColumnVisibility(leftColStatesBefore);
                        _leftView.SetDataTable(BuildDataTableFromFileData(leftFiltered), _lblLeftHeader.Text);
                        _leftRowToOriginalIndex = BuildRowIndexMap(leftFiltered);
                        _leftColToOriginalIndex = new List<int>(leftVisibleCols);
                    }
                    if (_rightFullFileData != null && _rightView != null && rightVisibleCols.Count > 0)
                    {
                        var rightFiltered = BuildProjectedFileData(_rightFullFileData, rightVisibleCols, diffOnly: true);
                        _rightView.SetColumnVisibility(rightColStatesBefore);
                        _rightView.SetDataTable(BuildDataTableFromFileData(rightFiltered), _lblRightHeader.Text);
                        _rightRowToOriginalIndex = BuildRowIndexMap(rightFiltered);
                        _rightColToOriginalIndex = new List<int>(rightVisibleCols);
                    }
                    // 筛选后重新绑定差异高亮事件
                    ApplyDiffHighlight(_leftView?.GetDataGridView(), true);
                    ApplyDiffHighlight(_rightView?.GetDataGridView(), false);
                }

                SetupSyncScroll();

                _lblInfo.Text = $"对比: {leftLabel} ← → {rightLabel}  " +
                               $"(差异行: {_diffRowCount})";
            }
            catch (Exception ex)
            {
                _lblInfo.Text = $"加载失败: {ex.Message}";
            }
            finally
            {
                // ★ 数据全部就绪后，把 DataGridView 加回界面，一次性渲染
                _leftView?.AttachGridToParent();
                _rightView?.AttachGridToParent();
                HideLoading();
                _isLoading = false;
            }
        }

        /// <summary>确保左右 TableBrowserControl 已创建（只创建一次）。
        /// 传入空路径避免自动加载当前文件，数据由后续 SetDataTable 注入。</summary>
        private void EnsureViewsCreated()
        {
            var leftPanel = _splitCompare.Panel1.Controls[0] as TableLayoutPanel;
            var rightPanel = _splitCompare.Panel2.Controls[0] as TableLayoutPanel;

            if (_leftView == null && leftPanel != null)
            {
                // 传入空字符串：跳过构造函数的自动文件加载，数据由 SetDataTable 注入
                _leftView = new TableBrowserControl("",
                    frozenRows: _frozenRows, headerRowIndex: _headerRowIndex,
                    lightweight: false)
                { Dock = DockStyle.Fill };
                if (_columnVisibility != null)
                    _leftView.SetColumnVisibility(_columnVisibility);
                leftPanel.Controls.Add(_leftView, 0, 1);
            }

            if (_rightView == null && rightPanel != null)
            {
                _rightView = new TableBrowserControl("",
                    frozenRows: _frozenRows, headerRowIndex: _headerRowIndex,
                    lightweight: false)
                { Dock = DockStyle.Fill };
                if (_columnVisibility != null)
                    _rightView.SetColumnVisibility(_columnVisibility);
                rightPanel.Controls.Add(_rightView, 0, 1);
            }
        }

        /// <summary>加载指定版本的 DataTable（带缓存）。
        /// revision 为空字符串时，直接读取当前目录文件（不通过 VCS 导出）。</summary>
        private async Task<DataTable?> LoadDataTableAsync(string revision)
        {
            // 缓存 key：空字符串表示"当前文件"
            var cacheKey = string.IsNullOrEmpty(revision) ? "__CURRENT_FILE__" : revision;

            // 检查缓存
            if (_dataCache.TryGetValue(cacheKey, out var cached))
                return cached;

            string? file;
            if (string.IsNullOrEmpty(revision))
            {
                // 直接使用当前文件（无需导出）
                file = _filePath;
            }
            else
            {
                // 从 SVN/Git 导出指定版本
                file = await VersionControlHelper.ExportRevisionAsync(
                    _filePath, _vcsType, revision, _tempDir);
            }

            if (file == null || !File.Exists(file))
                return null;

            // 解析文件到 DataTable
            var dataTable = await Task.Run(() => ParseFileToDataTable(file));

            if (dataTable != null)
            {
                // 缓存（限制最多 10 个）
                if (_dataCache.Count >= 10)
                {
                    var oldestKey = _dataCache.Keys.First();
                    _dataCache.Remove(oldestKey);
                }
                _dataCache[cacheKey] = dataTable;
            }

            return dataTable;
        }

        /// <summary>解析文件为 DataTable（CSV 或 Excel）</summary>
        private static DataTable? ParseFileToDataTable(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".xls" || ext == ".xlsx")
                return ReadExcelToDataTable(filePath);
            else
                return ReadCsvToDataTable(filePath);
        }

        private static DataTable? ReadCsvToDataTable(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return null;

                string content;
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(fs, Encoding.UTF8, true))
                {
                    content = reader.ReadToEnd();
                }

                var delimiter = DetectCsvDelimiter(content);
                var lines = SplitCsvLines(content);
                if (lines.Count == 0) return null;

                var headers = ParseCsvLine(lines[0], delimiter);
                var dataTable = new DataTable();
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var header in headers)
                {
                    var columnName = header;
                    int suffix = 1;
                    while (usedNames.Contains(columnName)) { columnName = $"{header}_{suffix}"; suffix++; }
                    usedNames.Add(columnName);
                    dataTable.Columns.Add(columnName, typeof(string));
                }

                for (int i = 0; i < lines.Count; i++)
                {
                    var values = ParseCsvLine(lines[i], delimiter);
                    var row = dataTable.NewRow();
                    for (int j = 0; j < Math.Min(values.Length, headers.Length); j++)
                        row[j] = values[j];
                    dataTable.Rows.Add(row);
                }

                return dataTable;
            }
            catch { return null; }
        }

        private static DataTable? ReadExcelToDataTable(string filePath)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(fs);
                var dataSet = reader.AsDataSet(new ExcelDataReader.ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataReader.ExcelDataTableConfiguration { UseHeaderRow = false }
                });
                if (dataSet.Tables.Count == 0) return null;
                var sheet = dataSet.Tables[0];
                var dataTable = new DataTable();
                int maxCols = sheet.Columns.Count;
                for (int i = 0; i < maxCols; i++)
                    dataTable.Columns.Add(GetExcelColumnName(i), typeof(string));
                for (int r = 0; r < sheet.Rows.Count; r++)
                {
                    var newRow = dataTable.NewRow();
                    for (int c = 0; c < maxCols; c++)
                        newRow[c] = sheet.Rows[r][c]?.ToString() ?? "";
                    dataTable.Rows.Add(newRow);
                }
                return dataTable;
            }
            catch { return null; }
        }

        private static char DetectCsvDelimiter(string content)
        {
            var lines = SplitCsvLines(content);
            if (lines.Count == 0) return ',';
            var firstLine = lines[0];
            var counts = new Dictionary<char, int>
            {
                { ',', firstLine.Count(c => c == ',') },
                { '\t', firstLine.Count(c => c == '\t') },
                { ';', firstLine.Count(c => c == ';') },
                { '|', firstLine.Count(c => c == '|') }
            };
            return counts.OrderByDescending(kv => kv.Value).First().Key;
        }

        private static List<string> SplitCsvLines(string content)
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
                        sb.Append('"');
                        i++;
                        continue;
                    }
                    inQuotes = !inQuotes;
                    sb.Append(c);
                }
                else if (!inQuotes && (c == '\n' || c == '\r'))
                {
                    if (c == '\r' && i + 1 < content.Length && content[i + 1] == '\n') i++;
                    lines.Add(sb.ToString());
                    sb.Clear();
                }
                else { sb.Append(c); }
            }
            if (sb.Length > 0) lines.Add(sb.ToString());
            return lines;
        }

        private static string[] ParseCsvLine(string line, char delimiter)
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
                    { sb.Append('"'); i++; }
                    else inQuotes = !inQuotes;
                }
                else if (c == delimiter && !inQuotes)
                { result.Add(sb.ToString()); sb.Clear(); }
                else sb.Append(c);
            }
            result.Add(sb.ToString());
            return result.ToArray();
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

        // ==================== 同步滚动 ====================
        private void SetupSyncScroll()
        {
            var leftDgv = _leftView?.GetDataGridView();
            var rightDgv = _rightView?.GetDataGridView();

            if (leftDgv != null)
            {
                leftDgv.Scroll -= OnLeftGridScroll;
                leftDgv.Scroll += OnLeftGridScroll;
            }
            if (rightDgv != null)
            {
                rightDgv.Scroll -= OnRightGridScroll;
                rightDgv.Scroll += OnRightGridScroll;
            }
        }

        private void OnLeftGridScroll(object? sender, ScrollEventArgs e)
        {
            if (_syncingScroll) return;
            var rightDgv = _rightView?.GetDataGridView();
            var leftDgv = _leftView?.GetDataGridView();
            if (rightDgv == null || leftDgv == null) return;

            _syncingScroll = true;
            try
            {
                if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
                    rightDgv.FirstDisplayedScrollingRowIndex = leftDgv.FirstDisplayedScrollingRowIndex;
                else if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                    rightDgv.HorizontalScrollingOffset = leftDgv.HorizontalScrollingOffset;
            }
            finally { _syncingScroll = false; }
        }

        private void OnRightGridScroll(object? sender, ScrollEventArgs e)
        {
            if (_syncingScroll) return;
            var leftDgv = _leftView?.GetDataGridView();
            var rightDgv = _rightView?.GetDataGridView();
            if (leftDgv == null || rightDgv == null) return;

            _syncingScroll = true;
            try
            {
                if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
                    leftDgv.FirstDisplayedScrollingRowIndex = rightDgv.FirstDisplayedScrollingRowIndex;
                else if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                    leftDgv.HorizontalScrollingOffset = rightDgv.HorizontalScrollingOffset;
            }
            finally { _syncingScroll = false; }
        }

        // ==================== 差异计算（异步） ====================
        private async Task ComputeDiffAsync(DataTable? leftData, DataTable? rightData)
        {
            _diffRowCount = 0;
            _diffComputed = false;

            // ★ 清空旧差异数据
            _leftDiffRows.Clear();
            _rightDiffRows.Clear();
            _leftDiffCells.Clear();
            _rightDiffCells.Clear();

            if (leftData == null || rightData == null) return;

            await Task.Run(() =>
            {
                int maxRows = Math.Max(leftData.Rows.Count, rightData.Rows.Count);
                int maxCols = Math.Max(leftData.Columns.Count, rightData.Columns.Count);
                int diffCount = 0;

                for (int row = 0; row < maxRows; row++)
                {
                    bool rowHasDiff = false;
                    HashSet<int>? leftDiffCols = null;
                    HashSet<int>? rightDiffCols = null;

                    for (int col = 0; col < maxCols; col++)
                    {
                        string? leftVal = (row < leftData.Rows.Count && col < leftData.Columns.Count)
                            ? leftData.Rows[row][col]?.ToString() : null;
                        string? rightVal = (row < rightData.Rows.Count && col < rightData.Columns.Count)
                            ? rightData.Rows[row][col]?.ToString() : null;

                        if (leftVal != rightVal)
                        {
                            rowHasDiff = true;
                            if (leftVal != null && rightVal == null)
                                { leftDiffCols ??= new(); leftDiffCols.Add(col); }
                            else if (leftVal == null && rightVal != null)
                                { rightDiffCols ??= new(); rightDiffCols.Add(col); }
                            else
                            {
                                leftDiffCols ??= new(); leftDiffCols.Add(col);
                                rightDiffCols ??= new(); rightDiffCols.Add(col);
                            }
                        }
                    }

                    if (rowHasDiff)
                    {
                        diffCount++;
                        // ★ 仅记录原始行号到集合（原始行号 = DataTable 行索引，包含固定行）
                        if (row < leftData.Rows.Count)
                        {
                            _leftDiffRows.Add(row);
                            foreach (var c in leftDiffCols ?? Enumerable.Empty<int>())
                                _leftDiffCells.Add((row << 16) | c);
                        }
                        if (row < rightData.Rows.Count)
                        {
                            _rightDiffRows.Add(row);
                            foreach (var c in rightDiffCols ?? Enumerable.Empty<int>())
                                _rightDiffCells.Add((row << 16) | c);
                        }
                    }
                }

                _diffRowCount = diffCount;
            });

            _diffComputed = true;
        }

        private void ApplyDiffHighlight(DataGridView? dgv, bool isLeft)
        {
            if (dgv == null) return;

            // ★ 差异数据已在 TableRow 上，不需要再构建 DiffTagData
            // 只绑定 OnDiffCellFormatting 事件即可

            var view = isLeft ? _leftView : _rightView;
            if (view != null)
            {
                view.OnExternalCellFormatting -= OnDiffCellFormatting;
                view.OnExternalCellFormatting += OnDiffCellFormatting;
            }
        }

        private void OnDiffCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (sender is not DataGridView dgv) return;
            if (e.RowIndex < 0) return;

            // ★ 判断是左表还是右表
            bool isLeft = dgv == _leftView?.GetDataGridView();
            var rowMap = isLeft ? _leftRowToOriginalIndex : _rightRowToOriginalIndex;
            var diffRows = isLeft ? _leftDiffRows : _rightDiffRows;
            var diffCells = isLeft ? _leftDiffCells : _rightDiffCells;

            // ★ 通过行映射获取原始行索引
            if (e.RowIndex >= rowMap.Count) return;
            int originalRowIndex = rowMap[e.RowIndex];

            // ★ 用原始行索引判断是否为差异行
            if (!diffRows.Contains(originalRowIndex)) return;

            Color backColor, selBackColor;
            int displayCol = e.ColumnIndex;

            // RowHeader 处理
            if (displayCol < 0)
            {
                backColor = _colorDiffRowBg;
                selBackColor = Color.FromArgb(240, 230, 180);
            }
            else
            {
                // ★ 计算当前显示列对应的原始列索引（使用固定的列映射，不依赖运行时 GetVisibleColumnIndices）
                var colMap = isLeft ? _leftColToOriginalIndex : _rightColToOriginalIndex;
                int originalColIndex = displayCol;
                if (displayCol < colMap.Count)
                    originalColIndex = colMap[displayCol];

                // ★ 用原始行索引和原始列索引判断是否为差异单元格
                int cellKey = (originalRowIndex << 16) | originalColIndex;
                bool cellHasDiff = diffCells.Contains(cellKey);

                if (cellHasDiff)
                {
                    backColor = _colorDiffCellBg;       // 红色：差异单元格
                    selBackColor = Color.FromArgb(255, 170, 170);
                }
                else
                {
                    backColor = _colorDiffRowBg;        // 黄色：差异行的非差异单元格
                    selBackColor = Color.FromArgb(240, 230, 180);
                }
            }

            // ★ 设置事件参数样式
            e.CellStyle.BackColor = backColor;
            e.CellStyle.SelectionBackColor = selBackColor;
            e.FormattingApplied = true;
        }

        // ==================== 差异视图刷新 ====================
        private void RefreshDiffView()
        {
            if (_leftView == null || _rightView == null) return;
            if (!_diffComputed) return;

            var leftDgv = _leftView.GetDataGridView();
            var rightDgv = _rightView.GetDataGridView();

            // ★ 获取当前列显隐状态（基于投影后列数的状态）
            var currentLeftColStates = _leftView.GetColumnCheckStates();
            var currentRightColStates = _rightView.GetColumnCheckStates();

            // ★ 关键修复：不能直接使用 GetVisibleColumnIndices()（投影后索引），
            //    需要使用 _leftColToOriginalIndex / _rightColToOriginalIndex 中的原始列索引。
            //    GetVisibleColumnIndices() 返回的是当前投影后 DataTable 中的列索引(0~N-1)，
            //    而 _leftFullFileData 保存的是完整列数的数据，需要用原始列索引来访问。
            var leftOriginalCols = new List<int>(_leftColToOriginalIndex);
            var rightOriginalCols = new List<int>(_rightColToOriginalIndex);

            // ★ 根据当前列显隐状态，从原始列索引中筛选出当前可见的列
            //    currentLeftColStates 维度 = 投影后列数，与原列索引一一对应
            var leftVisibleOriginalCols = new List<int>();
            for (int i = 0; i < leftOriginalCols.Count && i < currentLeftColStates.Length; i++)
            {
                if (currentLeftColStates[i])
                    leftVisibleOriginalCols.Add(leftOriginalCols[i]);
            }
            var rightVisibleOriginalCols = new List<int>();
            for (int i = 0; i < rightOriginalCols.Count && i < currentRightColStates.Length; i++)
            {
                if (currentRightColStates[i])
                    rightVisibleOriginalCols.Add(rightOriginalCols[i]);
            }

            bool diffOnly = _chkDiffOnly.Checked;

            // ★ 左表：使用原始列索引构建投影数据
            if (_leftFullFileData != null && leftVisibleOriginalCols.Count > 0)
            {
                var leftProjected = BuildProjectedFileData(_leftFullFileData, leftVisibleOriginalCols, diffOnly);
                _leftView.SetColumnVisibility(currentLeftColStates);
                _leftView.SetDataTable(BuildDataTableFromFileData(leftProjected), _lblLeftHeader.Text);
                _leftRowToOriginalIndex = BuildRowIndexMap(leftProjected);
                _leftColToOriginalIndex = new List<int>(leftVisibleOriginalCols);
            }
            else
            {
                _leftRowToOriginalIndex.Clear();
                _leftColToOriginalIndex.Clear();
            }

            // ★ 右表：同上
            if (_rightFullFileData != null && rightVisibleOriginalCols.Count > 0)
            {
                var rightProjected = BuildProjectedFileData(_rightFullFileData, rightVisibleOriginalCols, diffOnly);
                _rightView.SetColumnVisibility(currentRightColStates);
                _rightView.SetDataTable(BuildDataTableFromFileData(rightProjected), _lblRightHeader.Text);
                _rightRowToOriginalIndex = BuildRowIndexMap(rightProjected);
                _rightColToOriginalIndex = new List<int>(rightVisibleOriginalCols);
            }
            else
            {
                _rightRowToOriginalIndex.Clear();
                _rightColToOriginalIndex.Clear();
            }

            // 重新绑定差异高亮事件
            ApplyDiffHighlight(leftDgv, true);
            ApplyDiffHighlight(rightDgv, false);
        }

        /// <summary>从原始 FileTableData 按可见列索引构建投影后的 FileTableData</summary>
        private FileTableData BuildProjectedFileData(FileTableData source, List<int> visibleColIndices, bool diffOnly = false)
        {
            var result = new FileTableData
            {
                FilePath = source.FilePath,
                Type = source.Type,
                Delimiter = source.Delimiter
            };

            // 构建新列名列表（只包含可见列）
            for (int i = 0; i < visibleColIndices.Count; i++)
            {
                int srcIdx = visibleColIndices[i];
                if (srcIdx < source.ColumnNames.Count)
                    result.ColumnNames.Add(source.ColumnNames[srcIdx]);
                else
                    result.ColumnNames.Add($"Column{i}");
            }

            // 构建行数据（不再需要 DiffColumns 转换，差异信息由 _leftDiffRows/_leftDiffCells 独立管理）
            foreach (var srcRow in source.Rows)
            {
                // ★ 固定行（冻结行）始终保留，不受 diffOnly 影响
                bool isFrozenRow = srcRow.OriginalRowIndex < _frozenRows;
                if (diffOnly && !srcRow.IsDiffRow && !isFrozenRow) continue;

                var newRow = new TableRow
                {
                    OriginalRowIndex = srcRow.OriginalRowIndex,
                    IsDiffRow = srcRow.IsDiffRow,
                    DiffColumns = null
                };

                // 只复制可见列的单元格值
                for (int i = 0; i < visibleColIndices.Count; i++)
                {
                    int srcIdx = visibleColIndices[i];
                    newRow.Cells.Add(srcIdx < srcRow.Cells.Count ? srcRow.Cells[srcIdx] : "");
                }

                result.Rows.Add(newRow);
            }

            return result;
        }

        /// <summary>从投影后的 FileTableData 构建 DGV行索引→原始行索引 的映射（冻结行模式下跳过固定行）</summary>
        private List<int> BuildRowIndexMap(FileTableData projectedData)
        {
            var map = new List<int>();
            // ★ 冻结行模式下，DGV 不显示冻结行，跳过冻结行的映射
            for (int i = _frozenRows; i < projectedData.Rows.Count; i++)
                map.Add(projectedData.Rows[i].OriginalRowIndex);
            return map;
        }

        /// <summary>从 FileTableData 构建 DataTable</summary>
        private static DataTable BuildDataTableFromFileData(FileTableData fd)
        {
            var dt = new DataTable();
            foreach (var colName in fd.ColumnNames)
                dt.Columns.Add(colName, typeof(string));

            foreach (var row in fd.Rows)
            {
                var dr = dt.NewRow();
                for (int c = 0; c < Math.Min(row.Cells.Count, dt.Columns.Count); c++)
                    dr[c] = row.Cells[c];
                dt.Rows.Add(dr);
            }
            return dt;
        }

        // ==================== 清理 ====================
        private void CleanupTempFiles()
        {
            try
            {
                if (Directory.Exists(_tempDir))
                    Directory.Delete(_tempDir, true);
            }
            catch { }
        }

        private void CleanupViews()
        {
            if (_leftView != null)
            {
                var leftDgv = _leftView.GetDataGridView();
                if (leftDgv != null)
                    leftDgv.Scroll -= OnLeftGridScroll;
                _leftView.Dispose();
                _leftView = null;
            }
            if (_rightView != null)
            {
                var rightDgv = _rightView.GetDataGridView();
                if (rightDgv != null)
                    rightDgv.Scroll -= OnRightGridScroll;
                _rightView.Dispose();
                _rightView = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                CleanupViews();
                CleanupTempFiles();
                _dataCache.Clear();
            }
            base.Dispose(disposing);
        }
    }
}
