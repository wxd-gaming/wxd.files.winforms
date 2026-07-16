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
        private CheckBox _chkIdAlign;
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
        private int _mappingRowIndex = -1;
        private bool[]? _columnVisibility;

        // 左右两个 TableBrowserControl（每次切换版本都销毁重建，避免列状态残留）
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

        // ★ DGV 行索引 → 行属性（DiffType + IsBlankRow，用于 OnDiffCellFormatting 直接判断行类型）
        private List<(DiffRowType diffType, bool isBlankRow)> _leftRowProperties = new();
        private List<(DiffRowType diffType, bool isBlankRow)> _rightRowProperties = new();

        // ★ DGV 列索引 → 原始完整 DataTable 列索引的映射（解决列隐藏后列号转换问题）
        private List<int> _leftColToOriginalIndex = new();
        private List<int> _rightColToOriginalIndex = new();

        // ★ _columnVisibility 是唯一权威来源（只读配置，版本对比中不允许修改列显隐）
        // 已删除 _savedLeftColStates / _savedRightColStates，全部从 _columnVisibility 派生

        private int _diffRowCount = 0;
        private bool _diffComputed = false;
        private bool _syncingScroll = false;

        // ★ ID 对齐相关
        private int _uniqueColumnIndex = -1;                   // 从配置读取的唯一列索引（调用方传入）
        private int _idColumnIndex = -1;                       // ID 列索引（-1 表示未检测到或未启用）
        private HashSet<int> _leftDeletedRows = new();        // 左表删除行集合（右表无对应 ID）
        private HashSet<int> _rightAddedRows = new();         // 右表新增行集合（左表无对应 ID）

        // 配色
        private static readonly Color _colorSurface = Color.White;
        private static readonly Color _colorDiffCellBg = Color.FromArgb(255, 200, 200);   // 差异单元格：红色背景
        private static readonly Color _colorDiffRowBg = Color.FromArgb(255, 253, 220);    // 差异行：浅黄色背景
        private static readonly Color _colorHeaderBg = Color.FromArgb(240, 246, 252);
        private static readonly Color _colorLightText = Color.FromArgb(20, 60, 100);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(90, 130, 170);
        private static readonly Color _colorSelectedBg = Color.FromArgb(227, 242, 253);
        private static readonly Color _colorAddedRowBg = Color.FromArgb(213, 245, 227);    // 新增行：浅绿色
        private static readonly Color _colorDeletedRowBg = Color.FromArgb(250, 219, 216);  // 删除行：浅橙色
        private static readonly Color _colorBlankRowForAdded = Color.FromArgb(232, 250, 240);   // 空白行（对面为新增行）：更浅绿
        private static readonly Color _colorBlankRowForDeleted = Color.FromArgb(253, 235, 233);  // 空白行（对面为删除行）：更浅橙

        // ==================== 构造函数 ====================
        public VersionCompareForm(string filePath, VcsType vcsType,
            int frozenRows, int headerRowIndex,
            string? rootDir = null,
            bool[]? columnVisibility = null,
            int uniqueColumnIndex = -1,
            int mappingRowIndex = -1)
        {
            _filePath = filePath;
            _vcsType = vcsType;

            // ✅ 优先使用调用方传入的值（调用方已正确获取配置）
            _frozenRows = frozenRows;
            _headerRowIndex = headerRowIndex;
            _mappingRowIndex = mappingRowIndex;
            _columnVisibility = columnVisibility;
            _uniqueColumnIndex = uniqueColumnIndex;

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
            // LoadHistoryAsync 移到 Shown 事件中延迟执行，确保窗体先显示

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

            _chkIdAlign = new CheckBox
            {
                Text = "ID对齐匹配",
                AutoSize = true,
                Location = new Point(310, 8),
                Checked = true,
                ForeColor = _colorLightText,
                Font = new Font("Microsoft YaHei UI", 9f)
            };
            _chkIdAlign.CheckedChanged += async (s, e) =>
                await LoadComparisonAsync(_selectedIndex);

            _lblInfo = new Label
            {
                AutoSize = true,
                Location = new Point(440, 10),
                ForeColor = _colorLightTextSecondary,
                Font = new Font("Microsoft YaHei UI", 9f)
            };

            _pnlTop.Controls.AddRange(new Control[] {
                _chkDiffOnly, _chkAlwaysLatest, _chkIdAlign, _lblInfo });

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
            Shown += async (s, e) =>
            {
                _splitMain.SplitterDistance = 160;
                _splitCompare.SplitterDistance = _splitCompare.Width / 2;

                // ★ 延迟 500ms，确保窗体完成渲染后再检测 SVN/Git
                await Task.Delay(500);
                await LoadHistoryAsync();
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

        // ==================== 错误提示 ====================
        /// <summary>显示错误信息：红色背景标签 + MessageBox 弹窗</summary>
        private void ShowError(string message)
        {
            _lblInfo.BackColor = Color.FromArgb(200, 50, 50);
            _lblInfo.ForeColor = Color.White;
            _lblInfo.Text = message;

            // ★ 延迟弹出 MessageBox，确保差异化窗口先显示
            BeginInvoke(new Action(() =>
            {
                AlertForm.Show(this, message, "错误", AlertType.Error);
            }));
        }

        /// <summary>恢复 _lblInfo 为正常样式</summary>
        private void ResetInfoStyle()
        {
            _lblInfo.BackColor = Color.Transparent;
            _lblInfo.ForeColor = _colorLightTextSecondary;
        }

        // ==================== 加载历史 ====================
        private async Task LoadHistoryAsync()
        {
            try
            {
                _lblInfo.Text = "正在加载历史记录...";

                var result = _vcsType == VcsType.Svn
                    ? await VersionControlHelper.GetSvnHistoryAsync(_filePath)
                    : await VersionControlHelper.GetGitHistoryAsync(_filePath);

                if (!result.Success)
                {
                    ShowError(result.ErrorMessage!);
                    return;
                }

                _history = result.Entries;

                if (_history.Count == 0)
                {
                    ShowError("该文件没有提交历史记录");
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
                ResetInfoStyle();
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
                ShowError($"加载失败: {ex.Message}");
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

            // ★ 阶段0：销毁旧控件并重建新的 TableBrowserControl（避免列状态残留）
            RecreateViews();

            ShowLoading($"{leftLabel} ← → {rightLabel}");
            // 给消息循环一点时间处理，确保 LoadingForm 被绘制
            await Task.Delay(50);

            try
            {
                _leftRevision = leftRevision;
                _rightRevision = rightRevision;

                ResetInfoStyle();
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
                // ★ 每次都是新控件，直接用 _columnVisibility（唯一权威来源）
                if (_leftView != null && _columnVisibility != null)
                    _leftView.SetColumnVisibility(_columnVisibility);

                if (_rightView != null && _columnVisibility != null)
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
                        if (_leftDeletedRows.Contains(row.OriginalRowIndex))
                            row.DiffType = DiffRowType.Deleted;
                        else if (row.IsDiffRow)
                            row.DiffType = DiffRowType.Modified;
                        else
                            row.DiffType = DiffRowType.None;
                    }
                }
                if (_rightFullFileData != null)
                {
                    foreach (var row in _rightFullFileData.Rows)
                    {
                        row.IsDiffRow = _rightDiffRows.Contains(row.OriginalRowIndex);
                        if (_rightAddedRows.Contains(row.OriginalRowIndex))
                            row.DiffType = DiffRowType.Added;
                        else if (row.IsDiffRow)
                            row.DiffType = DiffRowType.Modified;
                        else
                            row.DiffType = DiffRowType.None;
                    }
                }

                // ★ 阶段3.2：ID对齐模式下，构建对齐后的左右 FileTableData（插入空白行使行一一对应）
                if (_chkIdAlign.Checked && _idColumnIndex >= 0 &&
                    _leftFullFileData != null && _rightFullFileData != null)
                {
                    var (alignedLeft, alignedRight) = BuildAlignedTables(
                        _leftFullFileData, _rightFullFileData,
                        leftData!, rightData!);
                    _leftFullFileData = alignedLeft;
                    _rightFullFileData = alignedRight;
                }

                // 阶段3.5：★ 根据当前列显隐状态重新构建投影后的 FileTableData → DataTable
                // ★ 直接使用 _columnVisibility（只读配置，唯一权威来源），不再依赖控件内部状态
                var leftColStatesBefore = _columnVisibility;
                var rightColStatesBefore = _columnVisibility;
                var leftVisibleCols = GetVisibleColsFromVisibility();
                var rightVisibleCols = GetVisibleColsFromVisibility();

                if (_leftFullFileData != null && leftVisibleCols.Count > 0 && _leftView != null)
                {
                    var leftProjected = BuildProjectedFileData(_leftFullFileData, leftVisibleCols, diffOnly: false);
                    _leftView.SetColumnVisibility(leftColStatesBefore);
                    _leftView.SetDataTable(BuildDataTableFromFileData(leftProjected), _lblLeftHeader.Text);
                    // ★ 构建行/列索引映射（DGV行/列→原始行/列），不再依赖 Tag 或 GetVisibleColumnIndices
                    _leftRowToOriginalIndex = BuildRowIndexMap(leftProjected, out _leftRowProperties);
                    _leftColToOriginalIndex = new List<int>(leftVisibleCols);
                }
                if (_rightFullFileData != null && rightVisibleCols.Count > 0 && _rightView != null)
                {
                    var rightProjected = BuildProjectedFileData(_rightFullFileData, rightVisibleCols, diffOnly: false);
                    _rightView.SetColumnVisibility(rightColStatesBefore);
                    _rightView.SetDataTable(BuildDataTableFromFileData(rightProjected), _lblRightHeader.Text);
                    _rightRowToOriginalIndex = BuildRowIndexMap(rightProjected, out _rightRowProperties);
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
                        _leftRowToOriginalIndex = BuildRowIndexMap(leftFiltered, out _leftRowProperties);
                        _leftColToOriginalIndex = new List<int>(leftVisibleCols);
                    }
                    if (_rightFullFileData != null && _rightView != null && rightVisibleCols.Count > 0)
                    {
                        var rightFiltered = BuildProjectedFileData(_rightFullFileData, rightVisibleCols, diffOnly: true);
                        _rightView.SetColumnVisibility(rightColStatesBefore);
                        _rightView.SetDataTable(BuildDataTableFromFileData(rightFiltered), _lblRightHeader.Text);
                        _rightRowToOriginalIndex = BuildRowIndexMap(rightFiltered, out _rightRowProperties);
                        _rightColToOriginalIndex = new List<int>(rightVisibleCols);
                    }
                    // 筛选后重新绑定差异高亮事件
                    ApplyDiffHighlight(_leftView?.GetDataGridView(), true);
                    ApplyDiffHighlight(_rightView?.GetDataGridView(), false);
                }

                SetupSyncScroll();

                ResetInfoStyle();
                _lblInfo.Text = $"对比: {leftLabel} ← → {rightLabel}  " +
                               $"(差异行: {_diffRowCount})";
            }
            catch (Exception ex)
            {
                ShowError($"加载失败: {ex.Message}");
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
        /// <summary>销毁旧的左右控件并重建新的 TableBrowserControl（每次切换版本时调用，避免列状态残留）。
        /// 初始注入 _columnVisibility 并禁用列管理按钮。</summary>
        private void RecreateViews()
        {
            var leftPanel = _splitCompare.Panel1.Controls[0] as TableLayoutPanel;
            var rightPanel = _splitCompare.Panel2.Controls[0] as TableLayoutPanel;

            // ★ 销毁旧控件（从父容器移除并 Dispose）
            if (_leftView != null)
            {
                var parent = _leftView.Parent;
                if (parent != null) parent.Controls.Remove(_leftView);
                _leftView.Dispose();
                _leftView = null;
            }
            if (_rightView != null)
            {
                var parent = _rightView.Parent;
                if (parent != null) parent.Controls.Remove(_rightView);
                _rightView.Dispose();
                _rightView = null;
            }

            // ★ 创建新控件
            var rootDir = Path.GetDirectoryName(_filePath);

            if (leftPanel != null)
            {
                var leftOptions = new TableBrowserOptions(_filePath)
                {
                    RootDir = rootDir,
                    FrozenRows = _frozenRows,
                    HeaderRowIndex = _headerRowIndex,
                    UniqueColumnIndex = _uniqueColumnIndex,
                    MappingRowIndex = _mappingRowIndex,
                    SkipAutoLoad = true,
                };
                _leftView = new TableBrowserControl(leftOptions) { Dock = DockStyle.Fill };
                if (_columnVisibility != null)
                    _leftView.SetColumnVisibility(_columnVisibility);
                _leftView.DisableColumnManagement();
                leftPanel.Controls.Add(_leftView, 0, 1);
            }

            if (rightPanel != null)
            {
                var rightOptions = new TableBrowserOptions(_filePath)
                {
                    RootDir = rootDir,
                    FrozenRows = _frozenRows,
                    HeaderRowIndex = _headerRowIndex,
                    UniqueColumnIndex = _uniqueColumnIndex,
                    MappingRowIndex = _mappingRowIndex,
                    SkipAutoLoad = true,
                };
                _rightView = new TableBrowserControl(rightOptions) { Dock = DockStyle.Fill };
                if (_columnVisibility != null)
                    _rightView.SetColumnVisibility(_columnVisibility);
                _rightView.DisableColumnManagement();
                rightPanel.Controls.Add(_rightView, 0, 1);
            }
        }

        /// <summary>从 _columnVisibility 派生可见列索引列表（第0列始终包含）</summary>
        private List<int> GetVisibleColsFromVisibility()
        {
            var result = new List<int>();
            if (_columnVisibility == null) return result;
            result.Add(0); // 第0列（序号列）始终显示
            for (int i = 1; i < _columnVisibility.Length; i++)
                if (_columnVisibility[i]) result.Add(i);
            return result;
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
                var exportResult = await VersionControlHelper.ExportRevisionAsync(
                    _filePath, _vcsType, revision, _tempDir);
                if (!exportResult.Success)
                {
                    ShowError(exportResult.ErrorMessage!);
                    return null;
                }
                file = exportResult.TempFile;
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

        /// <summary>解析文件为 DataTable（CSV 或 Excel），使用统一的 FileDataCache.LoadFileDirect</summary>
        private static DataTable? ParseFileToDataTable(string filePath)
        {
            try
            {
                var (dataTable, _) = FileDataCache.LoadFileDirect(filePath);
                return dataTable;
            }
            catch { return null; }
        }

        // ==================== 清理临时文件 ====================
        private void CleanupTempFiles()
        {
            try
            {
                if (Directory.Exists(_tempDir))
                    Directory.Delete(_tempDir, true);
            }
            catch { }
        }

        // ==================== 差异视图刷新（仅显示差异行开关） ====================
        private void RefreshDiffView()
        {
            if (_leftFullFileData == null || _rightFullFileData == null) return;
            if (_leftView == null || _rightView == null) return;

            var leftColStatesBefore = _columnVisibility;
            var rightColStatesBefore = _columnVisibility;
            var leftVisibleCols = _leftColToOriginalIndex;
            var rightVisibleCols = _rightColToOriginalIndex;

            if (leftVisibleCols.Count == 0 || rightVisibleCols.Count == 0) return;

            if (_chkDiffOnly.Checked)
            {
                // 仅显示差异行
                var leftFiltered = BuildProjectedFileData(_leftFullFileData, leftVisibleCols, diffOnly: true);
                _leftView.SetColumnVisibility(leftColStatesBefore);
                _leftView.SetDataTable(BuildDataTableFromFileData(leftFiltered), _lblLeftHeader.Text);
                _leftRowToOriginalIndex = BuildRowIndexMap(leftFiltered, out _leftRowProperties);
                _leftColToOriginalIndex = new List<int>(leftVisibleCols);

                var rightFiltered = BuildProjectedFileData(_rightFullFileData, rightVisibleCols, diffOnly: true);
                _rightView.SetColumnVisibility(rightColStatesBefore);
                _rightView.SetDataTable(BuildDataTableFromFileData(rightFiltered), _lblRightHeader.Text);
                _rightRowToOriginalIndex = BuildRowIndexMap(rightFiltered, out _rightRowProperties);
                _rightColToOriginalIndex = new List<int>(rightVisibleCols);
            }
            else
            {
                // 显示全部行
                var leftProjected = BuildProjectedFileData(_leftFullFileData, leftVisibleCols, diffOnly: false);
                _leftView.SetColumnVisibility(leftColStatesBefore);
                _leftView.SetDataTable(BuildDataTableFromFileData(leftProjected), _lblLeftHeader.Text);
                _leftRowToOriginalIndex = BuildRowIndexMap(leftProjected, out _leftRowProperties);
                _leftColToOriginalIndex = new List<int>(leftVisibleCols);

                var rightProjected = BuildProjectedFileData(_rightFullFileData, rightVisibleCols, diffOnly: false);
                _rightView.SetColumnVisibility(rightColStatesBefore);
                _rightView.SetDataTable(BuildDataTableFromFileData(rightProjected), _lblRightHeader.Text);
                _rightRowToOriginalIndex = BuildRowIndexMap(rightProjected, out _rightRowProperties);
                _rightColToOriginalIndex = new List<int>(rightVisibleCols);
            }

            ApplyDiffHighlight(_leftView?.GetDataGridView(), true);
            ApplyDiffHighlight(_rightView?.GetDataGridView(), false);
            SetupSyncScroll();
        }

        // ==================== 差异计算 ====================
        private async Task ComputeDiffAsync(DataTable? leftData, DataTable? rightData)
        {
            _leftDiffRows.Clear();
            _rightDiffRows.Clear();
            _leftDiffCells.Clear();
            _rightDiffCells.Clear();
            _leftDeletedRows.Clear();
            _rightAddedRows.Clear();
            _diffRowCount = 0;
            _diffComputed = false;

            if (leftData == null || rightData == null) return;

            bool useIdAlign = DetectIdColumn();
            int dataStartRow = _frozenRows;
            int maxCols = Math.Max(leftData.Columns.Count, rightData.Columns.Count);

            await Task.Run(() =>
            {
                if (useIdAlign && _idColumnIndex >= 0)
                {
                    int idCol = _idColumnIndex;

                    var leftIdMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    var rightIdMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    for (int r = dataStartRow; r < leftData.Rows.Count; r++)
                    {
                        string id = leftData.Rows[r][idCol]?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(id) && !leftIdMap.ContainsKey(id))
                            leftIdMap[id] = r;
                    }
                    for (int r = dataStartRow; r < rightData.Rows.Count; r++)
                    {
                        string id = rightData.Rows[r][idCol]?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(id) && !rightIdMap.ContainsKey(id))
                            rightIdMap[id] = r;
                    }

                    var allIds = new HashSet<string>(leftIdMap.Keys, StringComparer.OrdinalIgnoreCase);
                    allIds.UnionWith(rightIdMap.Keys);

                    foreach (var id in allIds)
                    {
                        bool leftExists = leftIdMap.TryGetValue(id, out int leftRow);
                        bool rightExists = rightIdMap.TryGetValue(id, out int rightRow);

                        if (leftExists && !rightExists)
                        {
                            // 左有右无 → 左侧删除行
                            _leftDiffRows.Add(leftRow);
                            _leftDeletedRows.Add(leftRow);
                            for (int c = 0; c < leftData.Columns.Count; c++)
                                _leftDiffCells.Add((leftRow << 16) | c);
                        }
                        else if (!leftExists && rightExists)
                        {
                            // 右有左无 → 右侧新增行
                            _rightDiffRows.Add(rightRow);
                            _rightAddedRows.Add(rightRow);
                            for (int c = 0; c < rightData.Columns.Count; c++)
                                _rightDiffCells.Add((rightRow << 16) | c);
                        }
                        else
                        {
                            bool rowHasDiff = false;
                            var leftDiffCols = new HashSet<int>();
                            var rightDiffCols = new HashSet<int>();

                            for (int col = 0; col < maxCols; col++)
                            {
                                string? leftVal = (col < leftData.Columns.Count)
                                    ? leftData.Rows[leftRow][col]?.ToString() : null;
                                string? rightVal = (col < rightData.Columns.Count)
                                    ? rightData.Rows[rightRow][col]?.ToString() : null;

                                if (leftVal != rightVal)
                                {
                                    rowHasDiff = true;
                                    leftDiffCols.Add(col);
                                    rightDiffCols.Add(col);
                                }
                            }

                            if (rowHasDiff)
                            {
                                _leftDiffRows.Add(leftRow);
                                foreach (var c in leftDiffCols)
                                    _leftDiffCells.Add((leftRow << 16) | c);
                                _rightDiffRows.Add(rightRow);
                                foreach (var c in rightDiffCols)
                                    _rightDiffCells.Add((rightRow << 16) | c);
                            }
                        }
                    }
                }
                else
                {
                    int maxRows = Math.Max(leftData.Rows.Count, rightData.Rows.Count);

                    for (int row = dataStartRow; row < maxRows; row++)
                    {
                        bool rowHasDiff = false;
                        HashSet<int> leftDiffCols = new();
                        HashSet<int> rightDiffCols = new();

                        for (int col = 0; col < maxCols; col++)
                        {
                            string? leftVal = (row < leftData.Rows.Count && col < leftData.Columns.Count)
                                ? leftData.Rows[row][col]?.ToString() : null;
                            string? rightVal = (row < rightData.Rows.Count && col < rightData.Columns.Count)
                                ? rightData.Rows[row][col]?.ToString() : null;

                            if (leftVal != rightVal)
                            {
                                rowHasDiff = true;
                                if (row < leftData.Rows.Count) leftDiffCols.Add(col);
                                if (row < rightData.Rows.Count) rightDiffCols.Add(col);
                            }
                        }

                        if (rowHasDiff)
                        {
                            if (row < leftData.Rows.Count)
                            {
                                _leftDiffRows.Add(row);
                                foreach (var c in leftDiffCols)
                                    _leftDiffCells.Add((row << 16) | c);
                            }
                            if (row < rightData.Rows.Count)
                            {
                                _rightDiffRows.Add(row);
                                foreach (var c in rightDiffCols)
                                    _rightDiffCells.Add((row << 16) | c);
                            }
                        }
                    }
                }
            });

            _diffRowCount = _leftDiffRows.Count + _rightDiffRows.Count;
            _diffComputed = true;
        }

        // ==================== ID列检测 ====================
        private bool DetectIdColumn()
        {
            // ★ 优先使用配置的唯一列索引
            if (_uniqueColumnIndex >= 0)
            {
                _idColumnIndex = _uniqueColumnIndex;
                return true;
            }

            // 当 _uniqueColumnIndex=-1 时，默认使用第0列
            _idColumnIndex = 0;
            return true;
        }

        // ==================== 构建对齐后的 FileTableData（ID对齐模式） ====================
        private (FileTableData leftAligned, FileTableData rightAligned) BuildAlignedTables(
            FileTableData leftFd, FileTableData rightFd,
            DataTable leftData, DataTable rightData)
        {
            int dataStartRow = _frozenRows;
            int idCol = _idColumnIndex;

            // 构建 ID → FileTableData 行索引 的映射
            var leftIdMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var rightIdMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < leftFd.Rows.Count; i++)
            {
                var row = leftFd.Rows[i];
                if (row.OriginalRowIndex < dataStartRow) continue;
                string? id = idCol < row.Cells.Count ? row.Cells[idCol]?.Trim() : null;
                if (!string.IsNullOrEmpty(id) && !leftIdMap.ContainsKey(id))
                    leftIdMap[id] = i;
            }
            for (int i = 0; i < rightFd.Rows.Count; i++)
            {
                var row = rightFd.Rows[i];
                if (row.OriginalRowIndex < dataStartRow) continue;
                string? id = idCol < row.Cells.Count ? row.Cells[idCol]?.Trim() : null;
                if (!string.IsNullOrEmpty(id) && !rightIdMap.ContainsKey(id))
                    rightIdMap[id] = i;
            }

            var leftAligned = new FileTableData();
            var rightAligned = new FileTableData();

            // 复制列名
            for (int c = 0; c < leftData.Columns.Count; c++)
                leftAligned.ColumnNames.Add(leftData.Columns[c].ColumnName);
            for (int c = 0; c < rightData.Columns.Count; c++)
                rightAligned.ColumnNames.Add(rightData.Columns[c].ColumnName);

            // ★ 先复制固定行/表头行（原样保留，不对齐）
            for (int i = 0; i < leftFd.Rows.Count; i++)
            {
                if (leftFd.Rows[i].OriginalRowIndex < dataStartRow)
                    leftAligned.Rows.Add(CloneRow(leftFd.Rows[i], leftFd.ColumnNames.Count));
            }
            for (int i = 0; i < rightFd.Rows.Count; i++)
            {
                if (rightFd.Rows[i].OriginalRowIndex < dataStartRow)
                    rightAligned.Rows.Add(CloneRow(rightFd.Rows[i], rightFd.ColumnNames.Count));
            }

            // 合并左右 ID 为有序去重列表
            var leftIdsOrdered = SortIds(leftIdMap.Keys);
            var rightIdsOrdered = SortIds(rightIdMap.Keys);
            var allIdsOrdered = new List<string>();
            int lp = 0, rp = 0;
            while (lp < leftIdsOrdered.Count || rp < rightIdsOrdered.Count)
            {
                if (lp >= leftIdsOrdered.Count)
                    allIdsOrdered.Add(rightIdsOrdered[rp++]);
                else if (rp >= rightIdsOrdered.Count)
                    allIdsOrdered.Add(leftIdsOrdered[lp++]);
                else
                {
                    int cmp = string.Compare(leftIdsOrdered[lp], rightIdsOrdered[rp], StringComparison.OrdinalIgnoreCase);
                    if (cmp == 0) { allIdsOrdered.Add(leftIdsOrdered[lp]); lp++; rp++; }
                    else if (cmp < 0) allIdsOrdered.Add(leftIdsOrdered[lp++]);
                    else allIdsOrdered.Add(rightIdsOrdered[rp++]);
                }
            }

            // 按有序 ID 遍历
            int colCount = Math.Max(leftData.Columns.Count, rightData.Columns.Count);

            foreach (var id in allIdsOrdered)
            {
                bool leftHas = leftIdMap.ContainsKey(id);
                bool rightHas = rightIdMap.ContainsKey(id);

                if (leftHas && rightHas)
                {
                    int leftFdIdx = leftIdMap[id];
                    int rightFdIdx = rightIdMap[id];
                    leftAligned.Rows.Add(CloneRow(leftFd.Rows[leftFdIdx], colCount));
                    rightAligned.Rows.Add(CloneRow(rightFd.Rows[rightFdIdx], colCount));
                }
                else if (leftHas)
                {
                    // 左有右无 → 删除行（左数据行 + 右空白行）
                    int leftFdIdx = leftIdMap[id];
                    var leftRow = CloneRow(leftFd.Rows[leftFdIdx], colCount);
                    leftRow.DiffType = DiffRowType.Deleted;
                    leftRow.IsDiffRow = true;
                    leftAligned.Rows.Add(leftRow);
                    rightAligned.Rows.Add(CreateBlankRow(leftFd.Rows[leftFdIdx].OriginalRowIndex, colCount, DiffRowType.Deleted));
                }
                else
                {
                    // 右有左无 → 新增行（左空白行 + 右数据行）
                    int rightFdIdx = rightIdMap[id];
                    leftAligned.Rows.Add(CreateBlankRow(rightFd.Rows[rightFdIdx].OriginalRowIndex, colCount, DiffRowType.Added));
                    var rightRow = CloneRow(rightFd.Rows[rightFdIdx], colCount);
                    rightRow.DiffType = DiffRowType.Added;
                    rightRow.IsDiffRow = true;
                    rightAligned.Rows.Add(rightRow);
                }
            }

            return (leftAligned, rightAligned);
        }

        /// <summary>按数值或字典序对ID列表排序</summary>
        private static List<string> SortIds(IEnumerable<string> ids)
        {
            var list = ids.ToList();
            if (list.Count == 0) return list;
            if (list.All(id => int.TryParse(id, out _)))
                return list.OrderBy(id => int.Parse(id)).ToList();
            return list.OrderBy(id => id, StringComparer.OrdinalIgnoreCase).ToList();
        }

        /// <summary>克隆 TableRow</summary>
        private static TableRow CloneRow(TableRow src, int colCount)
        {
            var row = new TableRow
            {
                OriginalRowIndex = src.OriginalRowIndex,
                IsDiffRow = src.IsDiffRow,
                DiffType = src.DiffType,
                IsBlankRow = src.IsBlankRow,
                DiffColumns = src.DiffColumns != null ? new HashSet<int>(src.DiffColumns) : null
            };
            for (int i = 0; i < colCount; i++)
                row.Cells.Add(i < src.Cells.Count ? src.Cells[i] : "");
            return row;
        }

        /// <summary>创建空白占位行</summary>
        private static TableRow CreateBlankRow(int originalRowIndex, int colCount, DiffRowType oppositeDiffType)
        {
            var row = new TableRow
            {
                OriginalRowIndex = -1,
                IsDiffRow = false,
                DiffType = oppositeDiffType,
                IsBlankRow = true
            };
            for (int i = 0; i < colCount; i++)
                row.Cells.Add("");
            return row;
        }

        // ==================== 构建投影后的 FileTableData（列显隐 + 差异行筛选） ====================
        private FileTableData BuildProjectedFileData(FileTableData source, List<int> visibleCols, bool diffOnly)
        {
            var projected = new FileTableData();

            // 复制可见列名
            foreach (var colIdx in visibleCols)
            {
                if (colIdx < source.ColumnNames.Count)
                    projected.ColumnNames.Add(source.ColumnNames[colIdx]);
                else
                    projected.ColumnNames.Add($"Col{colIdx}");
            }

            foreach (var row in source.Rows)
            {
                // ★ 固定行（冻结行）始终保留
                bool isFrozenRow = row.OriginalRowIndex < _frozenRows && row.OriginalRowIndex >= 0;
                // ★ 空白行始终保留（左右行数必须一致）
                if (isFrozenRow || row.IsBlankRow)
                {
                    var projRow = new TableRow
                    {
                        OriginalRowIndex = row.OriginalRowIndex,
                        IsDiffRow = row.IsDiffRow,
                        DiffType = row.DiffType,
                        IsBlankRow = row.IsBlankRow,
                        DiffColumns = row.DiffColumns
                    };
                    foreach (var colIdx in visibleCols)
                        projRow.Cells.Add(colIdx < row.Cells.Count ? row.Cells[colIdx] : "");
                    projected.Rows.Add(projRow);
                    continue;
                }

                // ★ diffOnly 模式：只保留差异行
                if (diffOnly && !row.IsDiffRow)
                    continue;

                var dataProjRow = new TableRow
                {
                    OriginalRowIndex = row.OriginalRowIndex,
                    IsDiffRow = row.IsDiffRow,
                    DiffType = row.DiffType,
                    IsBlankRow = row.IsBlankRow,
                    DiffColumns = row.DiffColumns
                };
                foreach (var colIdx in visibleCols)
                    dataProjRow.Cells.Add(colIdx < row.Cells.Count ? row.Cells[colIdx] : "");
                projected.Rows.Add(dataProjRow);
            }

            return projected;
        }

        /// <summary>从FileTableData构建DataTable</summary>
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

        /// <summary>构建DGV行索引→原始行索引映射</summary>
        private List<int> BuildRowIndexMap(FileTableData fd, out List<(DiffRowType diffType, bool isBlankRow)> rowProps)
        {
            var map = new List<int>();
            rowProps = new List<(DiffRowType, bool)>();
            // 跳过前 _frozenRows 行（固定行在 DGV 中不显示）
            for (int i = _frozenRows; i < fd.Rows.Count; i++)
            {
                int origIdx = fd.Rows[i].OriginalRowIndex;
                map.Add(origIdx);
                rowProps.Add((fd.Rows[i].DiffType, fd.Rows[i].IsBlankRow));
            }
            return map;
        }

        // ==================== 差异高亮绑定 ====================
        private void ApplyDiffHighlight(DataGridView? dgv, bool isLeft)
        {
            if (dgv == null) return;

            var view = isLeft ? _leftView : _rightView;
            if (view == null) return;

            view.OnExternalCellFormatting -= OnDiffCellFormatting;
            view.OnExternalCellFormatting += OnDiffCellFormatting;
        }

        private void OnDiffCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (sender is not DataGridView dgv) return;
            if (e.RowIndex < 0) return;

            bool isLeft = dgv == _leftView?.GetDataGridView();
            var rowProps = isLeft ? _leftRowProperties : _rightRowProperties;
            var rowMap = isLeft ? _leftRowToOriginalIndex : _rightRowToOriginalIndex;
            var diffCells = isLeft ? _leftDiffCells : _rightDiffCells;
            var diffRows = isLeft ? _leftDiffRows : _rightDiffRows;

            if (e.RowIndex >= rowProps.Count) return;
            var (diffType, isBlankRow) = rowProps[e.RowIndex];

            // ★ 空白行 → 根据对面行类型选择背景色
            if (isBlankRow)
            {
                if (diffType == DiffRowType.Deleted)
                {
                    // 左侧数据行删除 → 右侧空白行（更浅橙）
                    e.CellStyle.BackColor = _colorBlankRowForDeleted;
                    e.CellStyle.SelectionBackColor = Color.FromArgb(253, 235, 233);
                }
                else if (diffType == DiffRowType.Added)
                {
                    // 右侧数据行新增 → 左侧空白行（更浅绿）
                    e.CellStyle.BackColor = _colorBlankRowForAdded;
                    e.CellStyle.SelectionBackColor = Color.FromArgb(232, 250, 240);
                }
                else
                {
                    e.CellStyle.BackColor = Color.FromArgb(245, 245, 245);
                    e.CellStyle.SelectionBackColor = Color.FromArgb(230, 230, 230);
                }
                e.FormattingApplied = true;
                return;
            }

            // ★ 删除行（只在左表存在）
            if (diffType == DiffRowType.Deleted)
            {
                e.CellStyle.BackColor = _colorDeletedRowBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(250, 219, 216);
                e.FormattingApplied = true;
                return;
            }

            // ★ 新增行（只在右表存在）
            if (diffType == DiffRowType.Added)
            {
                e.CellStyle.BackColor = _colorAddedRowBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(213, 245, 227);
                e.FormattingApplied = true;
                return;
            }

            // ★ 修改行 → 逐单元格判断差异
            if (diffType != DiffRowType.Modified && diffType != DiffRowType.None) return;

            if (e.RowIndex >= rowMap.Count) return;
            int originalRow = rowMap[e.RowIndex];

            if (!diffRows.Contains(originalRow)) return;

            int cellKey = (originalRow << 16) | e.ColumnIndex;
            bool isDiffCell = diffCells.Contains(cellKey);

            if (isDiffCell)
            {
                e.CellStyle.BackColor = _colorDiffCellBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(255, 170, 170);
            }
            else
            {
                e.CellStyle.BackColor = _colorDiffRowBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(255, 253, 220);
            }

            e.FormattingApplied = true;
        }

        // ==================== 滚动同步 ====================
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
            var leftDgv = _leftView?.GetDataGridView();
            var rightDgv = _rightView?.GetDataGridView();
            if (leftDgv == null || rightDgv == null) return;

            _syncingScroll = true;
            try
            {
                if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
                {
                    int idx = leftDgv.FirstDisplayedScrollingRowIndex;
                    if (rightDgv.RowCount > 0)
                        rightDgv.FirstDisplayedScrollingRowIndex = Math.Min(idx, rightDgv.RowCount - 1);
                }
                else if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                {
                    rightDgv.HorizontalScrollingOffset = leftDgv.HorizontalScrollingOffset;
                }
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
                {
                    int idx = rightDgv.FirstDisplayedScrollingRowIndex;
                    if (leftDgv.RowCount > 0)
                        leftDgv.FirstDisplayedScrollingRowIndex = Math.Min(idx, leftDgv.RowCount - 1);
                }
                else if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                {
                    leftDgv.HorizontalScrollingOffset = rightDgv.HorizontalScrollingOffset;
                }
            }
            finally { _syncingScroll = false; }
        }
    }
}