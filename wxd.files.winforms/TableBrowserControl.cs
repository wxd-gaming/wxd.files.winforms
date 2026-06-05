using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
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
    /// <summary>统一的表格浏览控件（支持CSV和Excel）</summary>
    public class TableBrowserControl : UserControl
    {
        // ==================== 控件声明 ====================
        private DataGridView _dataGridHeader = new();
        private DataGridView _dataGrid = new();
        private SplitContainer _pnlGridContainer = new();
        private Button _btnSelectColumns = new();
        private TextBox _txtContentFilter = new();
        private TextBox _txtAdvancedFilter = new();
        private Label _lblInfo = new();
        private Label _lblCellPos = new();
        private Panel _pnlLoading = new();
        private Panel _pnlTop = new();
        private DataTable _dataTable = new();
        private DataTable _headerTable = new();
        private string _filePath = "";
        private TableFileType _fileType;
        private char _delimiter = ',';
        private int _frozenRows = 0;
        private bool _frozenRowsOverridden = false;
        private int _headerRowIndex = -1;
        private int _paneIndex = -1;
        private Dictionary<string, string> _fileRemarks = new();
        private bool[] _columnCheckStates = Array.Empty<bool>();
        private string _initialContentFilter = "";
        private List<ColumnLinkConfig> _columnLinks = new();
        private List<(string display, string fullPath)> _fileEntries = new();
        private List<int> _visibleColumnIndices = new();
        private bool _editingFromFrozen = false;
        private DataGridLinkToolTip? _linkToolTip;
        private FlowLayoutPanel? _flowTop;
        private Button? _btnReload;
        private Button? _btnHistory;

        // ==================== 公开属性 ====================
        public string FilePath => _filePath;
        public TableFileType FileType => _fileType;

        /// <summary>数据加载完成事件（供版本对比等外部功能使用）</summary>
        public event EventHandler? DataLoaded;

        /// <summary>请求关联打开另一张表事件</summary>
        public event EventHandler? LinkTableRequested;

        // 关联表跳转高亮
        private int _highlightColumnIndex = -1;
        private string _highlightValue = "";
        private string _highlightColumnName = "";
        private HashSet<int> _highlightRows = new();
        private bool _highlightsReady = false;
        private List<string> _highlightValues = new();

        // 排序
        private int _sortColumnIndex = -1;
        private SortOrder _sortOrder = SortOrder.None;

        // ==================== 配色 ====================
        private static readonly Color _colorSurface = Color.White;
        private static readonly Color _colorBorder = Color.FromArgb(173, 209, 237);
        private static readonly Color _colorLightText = Color.FromArgb(20, 60, 100);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(90, 130, 170);
        private static readonly Color _colorPrimary = Color.FromArgb(66, 165, 245);
        private static readonly Color _colorHoverBg = Color.FromArgb(235, 245, 255);
        private static readonly Color _colorHeaderBg = Color.FromArgb(240, 246, 252);
        private static readonly Color _colorAltRow = Color.FromArgb(248, 252, 255);
        private static readonly Color _colorSelectedBg = Color.FromArgb(227, 242, 253);
        private static readonly Color _colorGridLine = Color.FromArgb(200, 225, 245);
        private static readonly Color _colorRowNumBg = Color.FromArgb(240, 246, 252);

        private bool _lightweight;
        private string? _rootDir;
        private string FrozenDir => !string.IsNullOrEmpty(_rootDir) ? _rootDir : Path.GetDirectoryName(_filePath) ?? "";

        // ==================== 构造函数 ====================
        public TableBrowserControl(string filePath, int paneIndex = -1, string? rootDir = null,
            int frozenRows = -1, int headerRowIndex = -1, string initialContentFilter = "",
            bool lightweight = false)
        {
            _filePath = filePath;
            _paneIndex = paneIndex;
            _rootDir = rootDir;
            _lightweight = lightweight;
            if (frozenRows >= 0) { _frozenRows = frozenRows; _frozenRowsOverridden = true; }
            if (headerRowIndex >= 0) _headerRowIndex = headerRowIndex;
            _initialContentFilter = initialContentFilter;

            var dir = FrozenDir;
            if (!string.IsNullOrEmpty(dir))
                _fileRemarks = CsvSettings.GetFileRemarksByDirectory(dir);

            InitializeComponent();

            // 如果文件路径为空或文件不存在，跳过自动加载（由外部通过 SetDataTable 注入数据）
            if (!string.IsNullOrEmpty(_filePath) && File.Exists(_filePath))
            {
                _fileType = DetectFileType(filePath);
                if (_fileType == TableFileType.Csv)
                    LoadCsvAsync();
                else
                    LoadExcelAsync();
            }
            else
            {
                // 空数据模式：创建一个空的 DataTable，等待外部 SetDataTable
                _dataTable = new DataTable();
                _headerTable = _dataTable.Clone();
                _fileType = TableFileType.Csv; // 默认类型
            }
        }

        private static TableFileType DetectFileType(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".xls" || ext == ".xlsx")
                return TableFileType.Excel;
            return TableFileType.Csv;
        }

        // ==================== 高亮设置 ====================
        public void SetHighlight(int colIndex, string value, string colName = "", ColumnLinkConfig? linkConfig = null)
        {
            _highlightColumnIndex = colIndex;
            _highlightValue = value ?? "";
            _highlightColumnName = colName ?? "";
            _highlightValues.Clear();
            if (linkConfig != null && linkConfig.ValueExtractionEnabled && !string.IsNullOrEmpty(value))
            {
                _highlightValues = linkConfig.ExtractValues(value);
                if (_highlightValues.Count == 1 && !string.IsNullOrEmpty(linkConfig.MultiValueDelimiter) && value.Contains(linkConfig.MultiValueDelimiter))
                {
                    var groups = value.Split(new[] { linkConfig.MultiValueDelimiter }, StringSplitOptions.RemoveEmptyEntries);
                    var fallback = new List<string>();
                    foreach (var g in groups)
                    {
                        var trimmed = g.Trim();
                        if (!string.IsNullOrEmpty(trimmed))
                        {
                            if (!string.IsNullOrEmpty(linkConfig.ExtractDelimiter))
                            {
                                var parts = trimmed.Split(new[] { linkConfig.ExtractDelimiter }, StringSplitOptions.None);
                                if (linkConfig.ExtractPartIndex >= 0 && linkConfig.ExtractPartIndex < parts.Length)
                                    fallback.Add(parts[linkConfig.ExtractPartIndex].Trim());
                                else
                                    fallback.Add(trimmed);
                            }
                            else { fallback.Add(trimmed); }
                        }
                    }
                    if (fallback.Count > 1) _highlightValues = fallback.Distinct().ToList();
                }
            }
            else if (!string.IsNullOrEmpty(value))
                _highlightValues = new List<string> { value };
        }

        /// <summary>获取表头行的列名列表</summary>
        public List<string> GetHeaderColumnNames()
        {
            var names = new List<string>();
            if (_headerRowIndex < 0 || _headerRowIndex >= _dataTable.Rows.Count) return names;
            for (int i = 0; i < _dataTable.Columns.Count; i++)
                names.Add(_dataTable.Rows[_headerRowIndex][i]?.ToString() ?? $"Col{i}");
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

        // ==================== UI初始化 ====================
        private void InitializeComponent()
        {
            SuspendLayout();
            BackColor = _colorSurface;

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
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            topTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 1));
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // 第一行工具栏
            _flowTop = new FlowLayoutPanel
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
            cboFrozenRows.Items.AddRange(new object[] { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10" });
            if (!_frozenRowsOverridden)
            {
                _frozenRows = !string.IsNullOrEmpty(FrozenDir)
                    ? (_fileType == TableFileType.Csv
                        ? CsvSettings.GetCsvFrozenRowsByDirectory(FrozenDir)
                        : CsvSettings.GetExcelFrozenRowsByDirectory(FrozenDir))
                    : (_fileType == TableFileType.Csv
                        ? CsvSettings.DefaultFrozenRows
                        : CsvSettings.ExcelDefaultFrozenRows);
            }
            cboFrozenRows.SelectedIndex = _frozenRows;
            cboFrozenRows.SelectedIndexChanged += (s, e) =>
            {
                _frozenRows = cboFrozenRows.SelectedIndex;
                if (!string.IsNullOrEmpty(FrozenDir))
                {
                    if (_fileType == TableFileType.Csv)
                        CsvSettings.SetCsvFrozenRowsByDirectory(FrozenDir, _frozenRows);
                    else
                        CsvSettings.SetExcelFrozenRowsByDirectory(FrozenDir, _frozenRows);
                }
                else
                {
                    if (_fileType == TableFileType.Csv)
                        CsvSettings.DefaultFrozenRows = _frozenRows;
                    else
                        CsvSettings.ExcelDefaultFrozenRows = _frozenRows;
                }
                CsvSettings.Save();
                RefreshGridLayout();
            };

            // 表头行选择
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
            if (!_frozenRowsOverridden && !string.IsNullOrEmpty(FrozenDir))
            {
                _headerRowIndex = _fileType == TableFileType.Csv
                    ? CsvSettings.GetCsvHeaderRowIndexByDirectory(FrozenDir)
                    : CsvSettings.GetExcelHeaderRowIndexByDirectory(FrozenDir);
            }
            PopulateHeaderRowCombo(cboHeaderRow);
            cboHeaderRow.SelectedIndexChanged += (s, e) =>
            {
                _headerRowIndex = cboHeaderRow.SelectedIndex >= 0 ? cboHeaderRow.SelectedIndex : -1;
                if (!string.IsNullOrEmpty(FrozenDir))
                {
                    if (_fileType == TableFileType.Csv)
                        CsvSettings.SetCsvHeaderRowIndexByDirectory(FrozenDir, _headerRowIndex);
                    else
                        CsvSettings.SetExcelHeaderRowIndexByDirectory(FrozenDir, _headerRowIndex);
                }
                CsvSettings.Save();
                RefreshGridLayout();
            };
            Tag = cboHeaderRow;

            // 列选择按钮
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
                Size = new Size(120, 27),
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

            _lblInfo = new Label
            {
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 0, 0)
            };

            _flowTop.Controls.AddRange(new Control[] { lblFrozenRows, cboFrozenRows, lblHeaderRow, cboHeaderRow, lblHeaderFilter, _btnSelectColumns });

            // 重新加载按钮
            _btnReload = new Button
            {
                Text = "重新加载",
                Size = new Size(75, 27),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 2, 10, 0)
            };
            _btnReload.Click += (s, e) => ReloadContentAsync();
            _flowTop.Controls.Add(_btnReload);

            var lblCellPos = new Label
            {
                Text = "",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                ForeColor = _colorPrimary,
                Margin = new Padding(10, 5, 10, 0)
            };
            _flowTop.Controls.Add(lblCellPos);
            _lblCellPos = lblCellPos;
            _flowTop.Controls.Add(_lblInfo);

            // 分隔线
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
                if (e.KeyCode == Keys.Escape) { txtContentFilter.Text = ""; e.Handled = true; }
            };

            // 高级过滤
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
                if (e.KeyCode == Keys.Escape) { txtAdvancedFilter.Text = ""; e.Handled = true; }
            };
            var lblAdvHint = new Label
            {
                Text = "如: id=123 or name like test",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = _colorPrimary,
                Margin = new Padding(3, 5, 10, 0)
            };

            var flowContent = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, Height = 30,
                BackColor = _colorSurface, FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false, Padding = new Padding(8, 1, 8, 0), AutoSize = true
            };
            flowContent.Controls.AddRange(new Control[] { lblContentFilter, txtContentFilter });

            var flowAdv = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, Height = 30,
                BackColor = _colorSurface, FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false, Padding = new Padding(8, 1, 8, 5), AutoSize = true
            };
            flowAdv.Controls.AddRange(new Control[] { lblAdvFilter, txtAdvancedFilter, lblAdvHint });

            topTable.Controls.Add(_flowTop, 0, 0);
            topTable.Controls.Add(separatorLine, 0, 1);
            topTable.Controls.Add(flowContent, 0, 2);
            topTable.Controls.Add(flowAdv, 0, 3);

            _pnlTop = topTable;
            _txtContentFilter = txtContentFilter;
            _txtAdvancedFilter = txtAdvancedFilter;

            topTable.Resize += (s, e) =>
            {
                int w = topTable.ClientSize.Width - 280;
                if (w < 100) w = 100;
                txtContentFilter.Width = w;
                txtAdvancedFilter.Width = w;
            };

            // 表格容器
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
                if (e.RowIndex >= 0 && _dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;
                if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    var cellValue = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    var ctx = new ContextMenuStrip();
                    ctx.Font = new Font("Microsoft YaHei UI", 9f);
                    var copyItem = new ToolStripMenuItem("复制");
                    copyItem.Click += (s2, args) => { if (!string.IsNullOrEmpty(cellValue)) Clipboard.SetText(cellValue); };
                    ctx.Items.Add(copyItem);
                    var editItem = new ToolStripMenuItem("编辑");
                    editItem.Click += (s2, args) =>
                    {
                        var cell = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        _editingFromFrozen = true;
                        EditCell(cell, _dataGridHeader);
                    };
                    ctx.Items.Add(editItem);
                    // 固定Tips
                    var pinItem2 = new ToolStripMenuItem("固定Tips");
                    int frozenCol = e.ColumnIndex;
                    pinItem2.Click += (s2, args) =>
                    {
                        if (_linkToolTip?.GetToolTipContent != null)
                        {
                            var content = _linkToolTip.GetToolTipContent(_dataGridHeader, e.RowIndex, frozenCol);
                            if (content.title != null || content.lines.Count > 0 || !string.IsNullOrEmpty(content.rawText))
                            {
                                var cellRect = _dataGridHeader.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, false);
                                var screenPos = _dataGridHeader.PointToScreen(new Point(cellRect.Right + 6, cellRect.Top));
                                _linkToolTip.CreatePinnedWindow(content.title, content.lines, content.rawText, screenPos);
                            }
                        }
                    };
                    ctx.Items.Add(pinItem2);
                    ctx.Show(_dataGridHeader, _dataGridHeader.PointToClient(Cursor.Position));
                }
            };
            _dataGridHeader.CurrentCellChanged += (s, e) =>
            {
                if (_dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;
            };
            _dataGridHeader.CellDoubleClick += OnFrozenCellDoubleClick;
            _dataGridHeader.CellClick += OnCellClick;

            _dataGrid = CreateDataGridView();
            _dataGrid.Dock = DockStyle.Fill;
            _dataGrid.RowHeadersWidth = 50;
            _dataGrid.RowPostPaint += OnRowPostPaint;
            _dataGrid.ColumnHeadersVisible = false;
            _dataGrid.CellMouseDown += OnGridCellMouseDown;

            _dataGrid.Scroll += (s, e) =>
            {
                _dataGridHeader.HorizontalScrollingOffset = _dataGrid.HorizontalScrollingOffset;
            };
            _dataGrid.ColumnWidthChanged += (s, e) =>
            {
                int dataGridColumnIndex = e.Column.Index;
                int headerColumnIndex = dataGridColumnIndex;
                if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__" && dataGridColumnIndex > 0)
                    headerColumnIndex = dataGridColumnIndex - 1;
                if (headerColumnIndex >= 0 && headerColumnIndex < _dataGridHeader.Columns.Count)
                    _dataGridHeader.Columns[headerColumnIndex].Width = _dataGrid.Columns[dataGridColumnIndex].Width;
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
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = _colorAltRow },
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

        // ==================== 数据加载 ====================
        private async void LoadCsvAsync()
        {
            try
            {
                if (!this.IsHandleCreated)
                {
                    var tcs = new TaskCompletionSource<bool>();
                    EventHandler? handler = null;
                    handler = (s, e) => { this.HandleCreated -= handler; tcs.TrySetResult(true); };
                    this.HandleCreated += handler;
                    await tcs.Task;
                }
                var loadingStart = DateTime.UtcNow;
                ShowLoadingOverlay();
                await Task.Delay(50);
                await Task.Run(() => LoadCsvFile());
                var elapsed = (DateTime.UtcNow - loadingStart).TotalMilliseconds;
                if (elapsed < 200) await Task.Delay((int)(200 - elapsed));
                HideLoadingOverlay();
            }
            catch (Exception ex)
            {
                HideLoadingOverlay();
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, $"加载CSV文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error));
            }
        }

        private async void LoadExcelAsync()
        {
            try
            {
                if (!this.IsHandleCreated)
                {
                    var tcs = new TaskCompletionSource<bool>();
                    EventHandler? handler = null;
                    handler = (s, e) => { this.HandleCreated -= handler; tcs.TrySetResult(true); };
                    this.HandleCreated += handler;
                    await tcs.Task;
                }
                var loadingStart = DateTime.UtcNow;
                ShowLoadingOverlay();
                await Task.Delay(50);
                await Task.Run(() => LoadExcelFile());
                var elapsed = (DateTime.UtcNow - loadingStart).TotalMilliseconds;
                if (elapsed < 200) await Task.Delay((int)(200 - elapsed));
                HideLoadingOverlay();
            }
            catch (Exception ex)
            {
                HideLoadingOverlay();
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, $"加载Excel文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error));
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

            string content;
            using (var fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(fs, Encoding.UTF8, true))
            {
                content = reader.ReadToEnd();
            }

            _delimiter = DetectDelimiter(content);
            var lines = SplitCsvLines(content);
            if (lines.Count == 0)
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, "CSV文件为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning));
                return;
            }

            var headers = ParseCsvLine(lines[0]);
            var dataTable = new DataTable();
            var headerTable = new DataTable();
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var header in headers)
            {
                var columnName = header;
                int suffix = 1;
                while (usedNames.Contains(columnName)) { columnName = $"{header}_{suffix}"; suffix++; }
                usedNames.Add(columnName);
                dataTable.Columns.Add(columnName, typeof(string));
                headerTable.Columns.Add(columnName, typeof(string));
            }

            for (int i = 0; i < lines.Count; i++)
            {
                var values = ParseCsvLine(lines[i]);
                var row = dataTable.NewRow();
                for (int j = 0; j < Math.Min(values.Length, headers.Length); j++)
                    row[j] = values[j];
                dataTable.Rows.Add(row);
            }

            this.Invoke(() => OnDataLoaded(dataTable, headerTable, $"分隔符: '{_delimiter}' | 行数: {dataTable.Rows.Count} | 列数: {dataTable.Columns.Count}"));
        }

        private void LoadExcelFile()
        {
            if (!File.Exists(_filePath))
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

            this.Invoke(() => OnDataLoaded(dataTable, dataTable.Clone(), $"行数: {dataTable.Rows.Count} | 列数: {dataTable.Columns.Count}"));
        }

        private void OnDataLoaded(DataTable dataTable, DataTable headerTable, string infoText)
        {
            _dataTable = dataTable;
            _headerTable = headerTable.Clone();

            _columnCheckStates = new bool[_dataTable.Columns.Count];
            var relPath = GetCurrentRelativePath();
            var savedStates = CsvSettings.GetColumnVisibility(FrozenDir, relPath);
            if (savedStates != null && savedStates.Length == _columnCheckStates.Length)
            {
                _columnCheckStates = savedStates;
                _columnCheckStates[0] = true;
            }
            else
            {
                for (int i = 0; i < _columnCheckStates.Length; i++)
                    _columnCheckStates[i] = true;
            }

            _columnLinks = CsvSettings.GetColumnLinksByDirectory(FrozenDir);
            ColumnLinkChecker.BuildLinkLookups(_columnLinks, FrozenDir, _filePath);

            UpdateColumnSelectorButton();
            RebuildHighlightRows();
            _lblInfo.Text = infoText;

            if (_lightweight)
            {
                // 轻量模式：直接绑定 DataSource，不做额外初始化
                _dataGrid.DataSource = _dataTable;
                for (int i = 0; i < _dataGrid.Columns.Count; i++)
                {
                    _dataGrid.Columns[i].Width = 120;
                    _dataGrid.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
                    _dataGrid.Columns[i].HeaderText = GetExcelColumnName(i);
                }
                // 轻量模式：隐藏顶部工具栏
                _pnlTop.Visible = false;
            }
            else if (_frozenRows > 0)
                RefreshGridLayout();
            else
            {
                _dataGrid.DataSource = _dataTable;
                for (int i = 0; i < _dataGrid.Columns.Count; i++)
                {
                    _dataGrid.Columns[i].Width = 120;
                    _dataGrid.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
                    _dataGrid.Columns[i].HeaderText = GetExcelColumnName(i);
                }
            }

            if (!string.IsNullOrEmpty(_initialContentFilter))
                _txtContentFilter.Text = _initialContentFilter;

            // 初始化自定义 tooltip（只初始化一次；轻量模式跳过）
            if (!_lightweight && _linkToolTip == null)
            {
                _linkToolTip = new DataGridLinkToolTip(_dataGrid);
                _linkToolTip.RegisterGrid(_dataGridHeader);
                _linkToolTip.GetToolTipContent = GetLinkToolTipContent;
            }

            // 触发数据加载完成事件
            DataLoaded?.Invoke(this, EventArgs.Empty);

            // 检测是否 SVN/Git 管理，如果是则显示"历史记录"按钮（轻量模式跳过）
            if (!_lightweight)
                ShowHistoryButtonIfVcs();
        }

        /// <summary>为自定义 tooltip 提供内容（grid: 当前鼠标所在的 DataGridView）</summary>
        private (string? title, List<DataGridLinkToolTip.TipLine> lines, string? rawText) GetLinkToolTipContent(DataGridView grid, int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || columnIndex < 0)
                return (null, new List<DataGridLinkToolTip.TipLine>(), null);

            bool isHeader = grid == _dataGridHeader;

            // 计算原始列索引
            int originalColIndex = columnIndex;
            if (_visibleColumnIndices.Count > 0 && columnIndex < grid.Columns.Count)
            {
                int adjustedCol = columnIndex;
                if (grid.Columns.Count > 0 && grid.Columns[0].Name == "__RowNum__")
                    adjustedCol = columnIndex - 1;
                if (adjustedCol >= 0 && adjustedCol < _visibleColumnIndices.Count)
                    originalColIndex = _visibleColumnIndices[adjustedCol];
            }

            // 获取单元格值
            var cellVal = grid.Rows[rowIndex]?.Cells[columnIndex]?.Value?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(cellVal)) return (null, new List<DataGridLinkToolTip.TipLine>(), null);

            // 计算真实行号：_dataGrid 的行索引从 _frozenRows 开始
            int realRowIndex = isHeader ? rowIndex : rowIndex + _frozenRows;

            // 固定行直接显示单元格内容，不走关联关系逻辑
            if (isHeader)
            {
                string headerName = "";
                if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
                    headerName = _dataTable.Rows[_headerRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(headerName))
                    headerName = GetExcelColumnName(originalColIndex);

                string cellTitle = $"{headerName}  --  第{realRowIndex + 1}行";
                var cellLines = new List<DataGridLinkToolTip.TipLine>
                {
                    new DataGridLinkToolTip.TipLine { Text = cellVal, IsWarning = false }
                };
                return (cellTitle, cellLines, null);
            }

            // 获取当前文件相对路径
            string currentRel = GetCurrentRelativePath();

            // 查找关联配置
            var linkConfig = _columnLinks.FirstOrDefault(l =>
                l.SourceFile.EndsWith(Path.GetFileName(_filePath), StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == originalColIndex);

            if (linkConfig == null)
            {
                // 无关联配置时，显示单元格内容和表头信息
                string headerName = "";
                if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
                    headerName = _dataTable.Rows[_headerRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(headerName))
                    headerName = GetExcelColumnName(originalColIndex);

                string cellTitle = $"{headerName}  --  第{realRowIndex + 1}行";
                var cellLines = new List<DataGridLinkToolTip.TipLine>
                {
                    new DataGridLinkToolTip.TipLine { Text = cellVal, IsWarning = false }
                };
                return (cellTitle, cellLines, null);
            }

            // 提取匹配值
            var matchValues = (linkConfig.ValueExtractionEnabled)
                ? linkConfig.ExtractValues(cellVal) : new List<string> { cellVal };

            // 查找缓存
            string prefix = currentRel + "|" + originalColIndex + "->";
            LinkLookupEntry? lookup = null;
            foreach (var kvp in ColumnLinkChecker.GetLookups())
            {
                if (kvp.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    lookup = kvp.Value;
                    break;
                }
            }

            if (lookup == null) return (null, new List<DataGridLinkToolTip.TipLine>(), null);

            // 构建标题
            string displayTableName = Path.GetFileName(lookup.TargetFile);
            if (_fileRemarks.TryGetValue(lookup.TargetFile, out var remark) && !string.IsNullOrEmpty(remark))
                displayTableName = $"{remark} ({Path.GetFileName(lookup.TargetFile)})";
            else if (!string.IsNullOrEmpty(lookup.TargetColumnName))
                displayTableName = $"{Path.GetFileName(lookup.TargetFile)} [{lookup.TargetColumnName}]";

            string title = $"📋 {displayTableName}";

            // 构建映射行
            var lines = new List<DataGridLinkToolTip.TipLine>();
            foreach (var mv in matchValues)
            {
                if (lookup.Values.Contains(mv))
                {
                    if (lookup.DisplayNames.TryGetValue(mv, out var displayName) && !string.IsNullOrEmpty(displayName))
                        lines.Add(new DataGridLinkToolTip.TipLine { Text = $"  {mv}  →  {displayName}", IsWarning = false });
                    else
                        lines.Add(new DataGridLinkToolTip.TipLine { Text = $"  {mv}  →  (无显示名)", IsWarning = false });
                }
                else
                {
                    lines.Add(new DataGridLinkToolTip.TipLine { Text = $"  {mv}  →  ❌ 未找到", IsWarning = true });
                }
            }

            // 排序：警告在前
            var orderedLines = lines
                .Where(l => l.IsWarning)
                .Concat(lines.Where(l => !l.IsWarning))
                .ToList();

            return (title, orderedLines, null);
        }

        // ==================== CSV解析 ====================
        private char DetectDelimiter(string content)
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
                    { sb.Append('"'); i++; }
                    else inQuotes = !inQuotes;
                }
                else if (c == _delimiter && !inQuotes)
                { result.Add(sb.ToString()); sb.Clear(); }
                else sb.Append(c);
            }
            result.Add(sb.ToString());
            return result.ToArray();
        }

        // ==================== 加载遮罩 ====================
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

        // ==================== 重新加载 ====================
        private async void ReloadContentAsync()
        {
            if (!File.Exists(_filePath)) return;
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
                    handler = (s, e) => { this.HandleCreated -= handler; tcs.TrySetResult(true); };
                    this.HandleCreated += handler;
                    await tcs.Task;
                }
                await Task.Run(() =>
                {
                    if (_fileType == TableFileType.Csv)
                        ReloadCsvFile();
                    else
                        LoadExcelFile();
                });
            }
            catch (Exception ex)
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => MessageBox.Show(this, $"重新加载失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error));
                _columnCheckStates = savedColumnStates;
                _frozenRows = savedFrozenRows;
                _txtContentFilter.Text = savedFilter;
                _txtAdvancedFilter.Text = savedAdvFilter;
                UpdateColumnSelectorButton();
                return;
            }

            if (savedColumnStates.Length == _dataTable.Columns.Count)
                _columnCheckStates = savedColumnStates;
            else
            {
                _columnCheckStates = new bool[_dataTable.Columns.Count];
                for (int i = 0; i < _columnCheckStates.Length; i++) _columnCheckStates[i] = true;
            }
            _frozenRows = savedFrozenRows;
            _txtContentFilter.Text = savedFilter;
            _txtAdvancedFilter.Text = savedAdvFilter;
            UpdateColumnSelectorButton();
            ColumnLinkChecker.BuildLinkLookups(_columnLinks, FrozenDir, _filePath);
            RebuildHighlightRows();
            RefreshGridLayout();
        }

        private void ReloadCsvFile()
        {
            if (!File.Exists(_filePath)) return;
            string content;
            using (var fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(fs, Encoding.UTF8, true))
                content = reader.ReadToEnd();

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
                while (usedNames.Contains(columnName)) { columnName = $"{header}_{suffix}"; suffix++; }
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

            this.Invoke(() => { _delimiter = delimiter; _dataTable = dataTable; _headerTable = dataTable.Clone(); });
        }

        // ==================== 表头行下拉 ====================
        private void PopulateHeaderRowCombo(ComboBox cbo)
        {
            cbo.Items.Clear();
            for (int i = 0; i < 10; i++) cbo.Items.Add($"第{i + 1}行");
            if (_headerRowIndex >= 0 && _headerRowIndex < 10)
                cbo.SelectedIndex = _headerRowIndex;
            else
                cbo.SelectedIndex = -1;
        }

        // ==================== 布局刷新 ====================
        private void RefreshGridLayout()
        {
            if (_dataTable.Columns.Count == 0) return;

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
                        targetRow[j] = sourceRow[_visibleColumnIndices[j]];
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

        // ==================== 列选择 ====================
        private List<int> GetVisibleColumnIndices()
        {
            var visibleIndices = new List<int>();
            if (_dataTable.Columns.Count > 0) visibleIndices.Add(0);
            for (int i = 1; i < _dataTable.Columns.Count; i++)
            {
                if (i < _columnCheckStates.Length && _columnCheckStates[i])
                    visibleIndices.Add(i);
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
                CsvSettings.SetColumnVisibility(FrozenDir, GetCurrentRelativePath(), _columnCheckStates);
                UpdateColumnSelectorButton();
                RefreshGridLayout();
            }
        }

        private void EnsureFileEntriesLoaded()
        {
            if (_fileEntries.Count > 0) return;
            var root = FrozenDir;
            if (string.IsNullOrEmpty(root) || !Directory.Exists(root)) return;
            _fileRemarks = CsvSettings.GetFileRemarksByDirectory(root);
            _fileEntries.Clear();
            try
            {
                var showHidden = CsvSettings.ShowHiddenFiles;
                foreach (var ext in new[] { "*.csv", "*.xls", "*.xlsx" })
                {
                    foreach (var f in Directory.GetFiles(root, ext, SearchOption.AllDirectories))
                    {
                        if (!showHidden)
                        {
                            try { if ((File.GetAttributes(f) & FileAttributes.Hidden) == FileAttributes.Hidden) continue; } catch { }
                        }
                        var rel = f.StartsWith(root, StringComparison.OrdinalIgnoreCase)
                            ? f.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar) : Path.GetFileName(f);
                        var display = _fileRemarks.TryGetValue(rel, out var remark) && !string.IsNullOrEmpty(remark)
                            ? remark + " (" + Path.GetFileName(f) + ")" : rel;
                        _fileEntries.Add((display, f));
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
            _btnSelectColumns.ForeColor = checkedData < dataColumns ? Color.Red : _colorLightText;
        }

        // ==================== 行过滤 ====================
        private void ApplyFilters() => RefreshGridLayout();

        private List<DataRow> GetVisibleRows(string contentFilter)
        {
            var visibleRows = new List<DataRow>();
            var advFilter = _txtAdvancedFilter.Text.Trim();
            for (int i = _frozenRows; i < _dataTable.Rows.Count; i++)
            {
                var row = _dataTable.Rows[i];
                bool contentMatch = string.IsNullOrEmpty(contentFilter);
                if (!contentMatch)
                {
                    var processedFilter = contentFilter.Replace("||", " | ").Replace("&&", " & ");
                    foreach (var colIndex in _visibleColumnIndices)
                    {
                        var cellValue = row[colIndex]?.ToString() ?? "";
                        if (MatchesFilter(cellValue, processedFilter))
                        { contentMatch = true; break; }
                    }
                }
                bool advMatch = string.IsNullOrEmpty(advFilter);
                if (!advMatch) advMatch = EvaluateExpression(row, advFilter, _visibleColumnIndices);
                if (contentMatch && advMatch) visibleRows.Add(row);
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
                    newRow[i + 1] = sourceRow[_visibleColumnIndices[i]];
                filteredTable.Rows.Add(newRow);
            }
            _dataGrid.DataSource = filteredTable;
            if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__")
                _dataGrid.Columns[0].Visible = false;
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

        // ==================== 排序 ====================
        private void OnColumnHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.ColumnIndex >= _dataGrid.Columns.Count) return;
            if (_sortColumnIndex == e.ColumnIndex)
                _sortOrder = _sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            else { _sortColumnIndex = e.ColumnIndex; _sortOrder = SortOrder.Ascending; }
            var columnName = _dataGrid.Columns[e.ColumnIndex].DataPropertyName;
            var frozenRows = _dataTable.AsEnumerable().Take(_frozenRows).ToList();
            var dataRows = _dataTable.AsEnumerable().Skip(_frozenRows).ToList();
            var sortedDataRows = dataRows.OrderBy(row => row.Field<string>(columnName), new NaturalStringComparer(_sortOrder)).ToList();
            _dataTable = frozenRows.Concat(sortedDataRows).ToArray().CopyToDataTable();
            RefreshGridLayout();
            UpdateColumnHeaderSortGlyph();
        }

        private void UpdateColumnHeaderSortGlyph()
        {
            foreach (DataGridViewColumn col in _dataGrid.Columns)
                col.HeaderCell.SortGlyphDirection = col.Index == _sortColumnIndex ? _sortOrder : SortOrder.None;
        }

        // ==================== 单元格事件 ====================
        private void OnGridKeyDown(object? sender, KeyEventArgs e)
        {
            // 预留快捷键扩展
        }

        private void OnCellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var dgv = (DataGridView)sender!;
            int rowOff = (dgv == _dataGridHeader) ? 0 : _frozenRows;
            int colIdx = e.ColumnIndex;
            if (dgv != _dataGridHeader && dgv.Columns.Count > 0 && dgv.Columns[0].Name == "__RowNum__")
                colIdx = Math.Max(0, colIdx - 1);
            string colName = colIdx >= 0 && colIdx < 26 ? ((char)('A' + colIdx)).ToString() : $"Col{colIdx}";
            _lblCellPos.Text = $"行 {e.RowIndex + 1 + rowOff}  列 {colName}({colIdx})";
        }

        private void EditCell(DataGridViewCell cell, DataGridView dgv)
        {
            var rowIndex = cell.RowIndex;
            var columnIndex = cell.ColumnIndex;
            var currentValue = cell.Value?.ToString() ?? "";
            var (originalRowIndex, originalColumnIndex) = GetOriginalIndices(rowIndex, columnIndex);
            string colName = "";
            if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && originalColumnIndex >= 0 && originalColumnIndex < _dataTable.Columns.Count)
                colName = _dataTable.Rows[_headerRowIndex][originalColumnIndex]?.ToString() ?? "";

            if (_fileType == TableFileType.Csv)
            {
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
            else
            {
                // Excel：只读模式
                var dialog = new CellEditDialog(currentValue, originalRowIndex, originalColumnIndex, _filePath, colName,
                    onSave: null, readOnly: true, readOnlyMessage: "不支持Excel文件修改");
                dialog.Show(this);
            }
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

        private (int rowIndex, int columnIndex) GetOriginalIndices(int displayRowIndex, int displayColumnIndex)
        {
            if (_editingFromFrozen)
            {
                int originalColumnIndex = displayColumnIndex;
                if (_visibleColumnIndices.Count > 0 && displayColumnIndex < _dataGridHeader.Columns.Count)
                {
                    if (displayColumnIndex < _visibleColumnIndices.Count)
                        originalColumnIndex = _visibleColumnIndices[displayColumnIndex];
                }
                return (displayRowIndex, originalColumnIndex);
            }
            if (_dataGrid.DataSource is DataTable filteredTable && filteredTable.Columns.Count > 0)
            {
                int originalRowIndex = displayRowIndex;
                if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__")
                {
                    var cellValue = _dataGrid.Rows[displayRowIndex].Cells[0].Value;
                    if (cellValue != null) originalRowIndex = Convert.ToInt32(cellValue) - 1;
                }
                int originalColumnIndex = displayColumnIndex;
                if (_visibleColumnIndices.Count > 0 && displayColumnIndex < _dataGrid.Columns.Count)
                {
                    var actualDisplayIndex = displayColumnIndex;
                    if (_dataGrid.Columns[0].Name == "__RowNum__") actualDisplayIndex = displayColumnIndex - 1;
                    if (actualDisplayIndex >= 0 && actualDisplayIndex < _visibleColumnIndices.Count)
                        originalColumnIndex = _visibleColumnIndices[actualDisplayIndex];
                }
                return (originalRowIndex, originalColumnIndex);
            }
            return (displayRowIndex + _frozenRows, displayColumnIndex);
        }

        private void SaveCsvFile()
        {
            try
            {
                var sb = new StringBuilder();
                foreach (DataRow row in _dataTable.Rows)
                {
                    var values = new List<string>();
                    foreach (var item in row.ItemArray)
                    {
                        var value = item?.ToString() ?? "";
                        if (value.Contains(_delimiter) || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
                            value = "\"" + value.Replace("\"", "\"\"") + "\"";
                        values.Add(value);
                    }
                    sb.AppendLine(string.Join(_delimiter.ToString(), values));
                }
                var content = sb.ToString();
                var bytes = Encoding.UTF8.GetBytes(content);
                using (var fs = new FileStream(_filePath, FileMode.Create, FileAccess.Write, FileShare.Read))
                { fs.Write(bytes, 0, bytes.Length); fs.Flush(); }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"保存文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (sender == _dataGridHeader) return;
            if (e.RowIndex >= 0 && e.RowIndex < _dataGrid.Rows.Count)
                _dataGrid.Rows[e.RowIndex].Height = 24;
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                ColumnLinkChecker.CheckCellColor(e, _visibleColumnIndices, _dataGrid, _filePath, FrozenDir, _columnLinks);
                if (_highlightsReady && _highlightRows.Contains(e.RowIndex))
                {
                    e.CellStyle.BackColor = Color.FromArgb(255, 255, 200);
                    e.CellStyle.SelectionBackColor = Color.FromArgb(255, 240, 150);
                }
            }
            // 转发给外部事件处理器（如版本对比的差异高亮）
            OnExternalCellFormatting?.Invoke(sender, e);
        }

        private void OnCellToolTipTextNeeded(object? sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            // 系统 tooltip 已由 DataGridLinkToolTip 取代，此处不再设置文本
            if (sender == _dataGridHeader || e.RowIndex < 0 || e.ColumnIndex < 0) return;
            // 不设置 e.ToolTipText，保持为空（自定义 tooltip 会接管）
        }

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

            var copyItem = new ToolStripMenuItem("复制");
            copyItem.Click += (s, args) => { if (!string.IsNullOrEmpty(cellValue)) Clipboard.SetText(cellValue); };
            ctx.Items.Add(copyItem);

            var editItem = new ToolStripMenuItem("编辑");
            editItem.Click += (s, args) =>
            {
                var cell = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                _editingFromFrozen = false;
                EditCell(cell, _dataGrid);
            };
            ctx.Items.Add(editItem);

            var linkItem = new ToolStripMenuItem("打开关联表");
            if (link == null) { linkItem.Enabled = false; }
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

            // 固定Tips
            var pinItem = new ToolStripMenuItem("固定Tips");
            pinItem.Click += (s, args) =>
            {
                // 获取当前单元格的 tooltip 内容
                int adjCol2 = e.ColumnIndex;
                if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__")
                    adjCol2 = e.ColumnIndex - 1;
                if (adjCol2 >= 0 && _linkToolTip?.GetToolTipContent != null)
                {
                    var content = _linkToolTip.GetToolTipContent(_dataGrid, e.RowIndex, adjCol2);
                    if (content.title != null || content.lines.Count > 0 || !string.IsNullOrEmpty(content.rawText))
                    {
                        var cellRect = _dataGrid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, false);
                        var screenPos = _dataGrid.PointToScreen(new Point(cellRect.Right + 6, cellRect.Top));
                        _linkToolTip.CreatePinnedWindow(content.title, content.lines, content.rawText, screenPos);
                    }
                }
            };
            ctx.Items.Add(pinItem);

            ctx.Show(_dataGrid, _dataGrid.PointToClient(Cursor.Position));
        }

        private string GetCurrentRelativePath()
        {
            if (!string.IsNullOrEmpty(FrozenDir) && _filePath.StartsWith(FrozenDir, StringComparison.OrdinalIgnoreCase))
                return _filePath.Substring(FrozenDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(_filePath);
        }

        // ==================== 行号绘制 ====================
        private void OnRowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            int rowIndex;
            if (_dataGrid.Columns.Count > 0 && _dataGrid.Columns[0].Name == "__RowNum__" && e.RowIndex >= 0)
            {
                var cellValue = _dataGrid.Rows[e.RowIndex].Cells[0].Value;
                rowIndex = cellValue != null ? Convert.ToInt32(cellValue) : e.RowIndex + 1 + _frozenRows;
            }
            else rowIndex = e.RowIndex + 1 + _frozenRows;
            var rowHeaderBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, _dataGrid.RowHeadersWidth, e.RowBounds.Height);
            using (var brush = new SolidBrush(_dataGrid.RowHeadersDefaultCellStyle.ForeColor))
            {
                var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                e.Graphics.DrawString(rowIndex.ToString(), _dataGrid.Font, brush, rowHeaderBounds, format);
            }
        }

        private void OnFrozenRowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            var rowIndex = e.RowIndex + 1;
            var rowHeaderBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, _dataGridHeader.RowHeadersWidth, e.RowBounds.Height);
            using (var brush = new SolidBrush(_dataGridHeader.RowHeadersDefaultCellStyle.ForeColor))
            {
                var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                e.Graphics.DrawString(rowIndex.ToString(), _dataGridHeader.Font, brush, rowHeaderBounds, format);
            }
        }

        // ==================== 高亮重建 ====================
        private void RebuildHighlightRows()
        {
            _highlightRows.Clear();
            if (_highlightColumnIndex < 0 || _highlightValues.Count == 0 || _dataTable.Columns.Count == 0)
            { _highlightsReady = false; return; }
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
            if (_highlightRows.Count > 0)
            {
                string colRef = !string.IsNullOrEmpty(_highlightColumnName) ? _highlightColumnName : $"Col{_highlightColumnIndex}";
                string filterValue = _highlightValues.Count == 1
                    ? $"{colRef}='{_highlightValues[0]}'"
                    : string.Join(" or ", _highlightValues.Select(v => $"{colRef}='{v}'"));
                _txtAdvancedFilter.Text = filterValue;
            }
        }

        // ==================== 表达式过滤 ====================
        private bool EvaluateExpression(DataRow row, string expr, List<int> columnsToSearch)
        {
            expr = expr.Trim();
            if (string.IsNullOrEmpty(expr)) return true;
            while (expr.StartsWith("(") && expr.EndsWith(")") && IsBalanced(expr))
                expr = expr.Substring(1, expr.Length - 2).Trim();

            int parenDepth = 0, orPos = -1, andPos = -1;
            for (int i = 0; i < expr.Length; i++)
            {
                if (expr[i] == '(') parenDepth++;
                else if (expr[i] == ')') parenDepth--;
                else if (parenDepth == 0)
                {
                    if (i + 2 < expr.Length && (expr[i] == 'o' || expr[i] == 'O') && (expr[i + 1] == 'r' || expr[i + 1] == 'R')
                        && (i + 2 >= expr.Length || !char.IsLetterOrDigit(expr[i + 2])) && (i == 0 || !char.IsLetterOrDigit(expr[i - 1])))
                    { orPos = i; break; }
                    if (i + 3 < expr.Length && (expr[i] == 'a' || expr[i] == 'A') && (expr[i + 1] == 'n' || expr[i + 1] == 'N')
                        && (expr[i + 2] == 'd' || expr[i + 2] == 'D') && (i + 3 >= expr.Length || !char.IsLetterOrDigit(expr[i + 3]))
                        && (i == 0 || !char.IsLetterOrDigit(expr[i - 1])))
                    { if (andPos < 0) andPos = i; i += 3; }
                }
            }
            if (orPos >= 0)
            {
                return EvaluateExpression(row, expr.Substring(0, orPos).Trim(), columnsToSearch) ||
                       EvaluateExpression(row, expr.Substring(orPos + 2).Trim(), columnsToSearch);
            }
            if (andPos >= 0)
            {
                return EvaluateExpression(row, expr.Substring(0, andPos).Trim(), columnsToSearch) &&
                       EvaluateExpression(row, expr.Substring(andPos + 3).Trim(), columnsToSearch);
            }
            return EvaluateCondition(row, expr, columnsToSearch);
        }

        private bool IsBalanced(string s)
        {
            int d = 0;
            foreach (char c in s) { if (c == '(') d++; else if (c == ')') d--; if (d < 0) return false; }
            return d == 0;
        }

        private bool EvaluateCondition(DataRow row, string condition, List<int> columnsToSearch)
        {
            condition = condition.Trim();
            string op = ""; int opPos = -1;
            var likeIdx = condition.IndexOf(" like ", StringComparison.OrdinalIgnoreCase);
            if (likeIdx >= 0) { op = "like"; opPos = likeIdx; }
            else
            {
                var geIdx = FindOperator(condition, ">=");
                var leIdx = FindOperator(condition, "<=");
                var neqIdx = FindOperator(condition, "!=");
                var ltgtIdx = FindOperator(condition, "<>");
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
                    if (cellValue.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0) return true;
                }
                return false;
            }

            var rawColumnName = condition.Substring(0, opPos).Trim();
            int opRealLen = op == "like" ? " like ".Length : op.Length;
            var rawValue = condition.Substring(opPos + opRealLen).Trim();
            var columnName = rawColumnName.Trim();
            var value = rawValue;
            if ((value.StartsWith("'") && value.EndsWith("'")) || (value.StartsWith("\"") && value.EndsWith("\"")))
                value = value.Substring(1, value.Length - 2);

            if (!string.IsNullOrEmpty(columnName) && char.IsLetter(columnName[0]))
            {
                int colIndex = ResolveColumnIndex(columnName);
                if (colIndex >= 0 && columnsToSearch.Contains(colIndex))
                    return CompareValue(row[colIndex]?.ToString() ?? "", value, op);
                foreach (var ci in columnsToSearch)
                {
                    var cellValue = row[ci]?.ToString() ?? "";
                    if (cellValue.IndexOf(condition.Trim(), StringComparison.OrdinalIgnoreCase) >= 0) return true;
                }
                return false;
            }
            foreach (var colIndex in columnsToSearch)
            {
                var cellValue = row[colIndex]?.ToString() ?? "";
                if (cellValue.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            return false;
        }

        private int ResolveColumnIndex(string columnName)
        {
            for (int i = 0; i < _dataTable.Columns.Count; i++)
                if (_dataTable.Columns[i].ColumnName.Equals(columnName, StringComparison.OrdinalIgnoreCase)) return i;
            if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count)
            {
                for (int i = 0; i < _dataTable.Columns.Count; i++)
                {
                    var cellValue = _dataTable.Rows[_headerRowIndex][i]?.ToString() ?? "";
                    if (cellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase)) return i;
                }
            }
            int maxHeaderScan = Math.Min(_frozenRows, _dataTable.Rows.Count);
            for (int i = 0; i < _dataTable.Columns.Count; i++)
            {
                for (int row = 0; row < maxHeaderScan; row++)
                {
                    if (row == _headerRowIndex) continue;
                    var cellValue = _dataTable.Rows[row][i]?.ToString() ?? "";
                    if (cellValue.Equals(columnName, StringComparison.OrdinalIgnoreCase)) return i;
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
                if (nextIdx < s.Length && op.Length == 1 && (op[0] == '>' || op[0] == '<' || op[0] == '!') && s[nextIdx] == '=')
                { idx = s.IndexOf(op, nextIdx + 1, StringComparison.Ordinal); continue; }
                return idx;
            }
            return -1;
        }

        private bool CompareValue(string cellValue, string filterValue, string op)
        {
            bool numericCompare = false;
            double cellNum = 0, filterNum = 0;
            if (op != "like" && double.TryParse(cellValue, out cellNum) && double.TryParse(filterValue, out filterNum))
                numericCompare = true;

            switch (op.ToLower())
            {
                case "like":
                    if (filterValue.Contains("*") || filterValue.Contains("%"))
                    {
                        var pattern = Regex.Escape(filterValue).Replace("\\*", ".*").Replace("\\%", ".*");
                        return Regex.IsMatch(cellValue, "^" + pattern + "$", RegexOptions.IgnoreCase);
                    }
                    return cellValue.IndexOf(filterValue, StringComparison.OrdinalIgnoreCase) >= 0;
                case "=":
                    return numericCompare ? Math.Abs(cellNum - filterNum) < 0.000001
                        : cellValue.Equals(filterValue, StringComparison.OrdinalIgnoreCase);
                case "!=":
                case "<>":
                    return numericCompare ? Math.Abs(cellNum - filterNum) >= 0.000001
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

        private bool MatchesFilter(string text, string filter)
        {
            if (string.IsNullOrEmpty(filter)) return true;
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
            return text.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // ==================== 工具方法 ====================
        private static string GetExcelColumnName(int columnIndex)
        {
            var result = new StringBuilder();
            int index = columnIndex;
            do { result.Insert(0, (char)('A' + (index % 26))); index = index / 26 - 1; } while (index >= 0);
            return result.ToString();
        }

        private static DataTable? ReadExcelToDataTable(string filePath)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var reader = ExcelReaderFactory.CreateReader(fs);
                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
                });
                if (dataSet.Tables.Count == 0) return null;
                var sheet = dataSet.Tables[0];
                var dataTable = new DataTable();
                int maxCols = sheet.Columns.Count;
                for (int i = 0; i < maxCols; i++)
                    dataTable.Columns.Add(ConvertToColumnLetter(i), typeof(string));
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

        private static string ConvertToColumnLetter(int index)
        {
            var result = new StringBuilder();
            int num = index;
            do { result.Insert(0, (char)('A' + (num % 26))); num = num / 26 - 1; } while (num >= 0);
            return result.ToString();
        }

        // ==================== 版本历史按钮 ====================
        private void ShowHistoryButtonIfVcs()
        {
            try
            {
                if (_btnHistory != null) return;
                if (string.IsNullOrEmpty(_filePath) || !File.Exists(_filePath)) return;

                var vcs = VersionControlHelper.DetectVcs(_filePath);
                if (vcs == VcsType.None) return;

                _btnHistory = new Button
                {
                    Text = vcs == VcsType.Svn ? "SVN历史" : "Git历史",
                    Size = new Size(75, 27),
                    Font = new Font("Microsoft YaHei UI", 9f),
                    FlatStyle = FlatStyle.Flat,
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = Color.FromArgb(22, 163, 74),
                    ForeColor = Color.White,
                    FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(21, 128, 61) },
                    Cursor = Cursors.Hand,
                    Margin = new Padding(0, 2, 10, 0)
                };

                var vcsType = vcs;
                var frozen = _frozenRows;
                var headerRow = _headerRowIndex;
                var filePath = _filePath;

                _btnHistory.Click += (s, e) =>
                {
                    var form = new VersionCompareForm(filePath, vcsType, frozen, headerRow);
                    form.Show(this);
                };

                _flowTop?.Controls.Add(_btnHistory);
            }
            catch { }
        }

        /// <summary>获取内部 DataTable（供版本对比等外部功能使用）</summary>
        public DataTable GetDataTable() => _dataTable;

        /// <summary>获取内部 DataGridView（供版本对比等外部功能使用）</summary>
        public DataGridView GetDataGridView() => _dataGrid;

        /// <summary>
        /// 直接设置 DataTable 并刷新显示（供版本对比等外部功能复用控件实例）。
        /// 跳过文件加载、tooltip 初始化、VCS 检测等。
        /// 隐藏"重新加载"和"历史记录"按钮，避免版本对比中无限嵌套。
        /// </summary>
        /// <summary>从父容器中移除 DataGridView，防止设置 DataSource 时逐行绘制</summary>
        public void DetachGridFromParent()
        {
            var gridParent = _dataGrid.Parent;
            var gridHeaderParent = _dataGridHeader.Parent;

            if (gridParent != null)
                gridParent.Controls.Remove(_dataGrid);
            if (gridHeaderParent != null)
                gridHeaderParent.Controls.Remove(_dataGridHeader);
        }

        /// <summary>将 DataGridView 重新加入父容器（在数据绑定完成后调用）</summary>
        public void AttachGridToParent()
        {
            // _dataGrid 的父容器是 _pnlGridContainer.Panel2
            // _dataGridHeader 的父容器是 _pnlGridContainer.Panel1
            if (!_pnlGridContainer.Panel2.Controls.Contains(_dataGrid))
                _pnlGridContainer.Panel2.Controls.Add(_dataGrid);
            if (!_pnlGridContainer.Panel1Collapsed && !_pnlGridContainer.Panel1.Controls.Contains(_dataGridHeader))
                _pnlGridContainer.Panel1.Controls.Add(_dataGridHeader);
        }

        public void SetDataTable(DataTable dataTable, string infoText = "")
        {
            if (dataTable == null) return;

            // ★ 彻底清理旧绑定，避免多次切换后事件/引用累积导致卡顿
            _dataGrid.DataSource = null;
            _dataGridHeader.DataSource = null;
            // 清理可能由外部附加的 CellFormatting 事件（如版本对比的差异高亮）
            _dataGrid.CellFormatting -= OnExternalCellFormatting;
            // 清理高亮状态
            _highlightRows.Clear();
            _highlightsReady = false;

            _dataTable = dataTable;
            _headerTable = dataTable.Clone();

            _columnCheckStates = new bool[_dataTable.Columns.Count];
            for (int i = 0; i < _columnCheckStates.Length; i++)
                _columnCheckStates[i] = true;

            RebuildHighlightRows();
            _lblInfo.Text = string.IsNullOrEmpty(infoText)
                ? $"行数: {dataTable.Rows.Count} | 列数: {dataTable.Columns.Count}"
                : infoText;

            // 隐藏"重新加载"和"历史记录"按钮（避免版本对比中无限嵌套）
            if (_btnReload != null) _btnReload.Visible = false;
            if (_btnHistory != null) _btnHistory.Visible = false;

            if (_frozenRows > 0)
                RefreshGridLayout();
            else
            {
                _pnlGridContainer.Panel1Collapsed = true;
                _dataGrid.DataSource = _dataTable;
                for (int i = 0; i < _dataGrid.Columns.Count; i++)
                {
                    _dataGrid.Columns[i].Width = 120;
                    _dataGrid.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
                    _dataGrid.Columns[i].HeaderText = GetExcelColumnName(i);
                }
            }
        }

        /// <summary>供外部（版本对比等）移除的 CellFormatting 事件占位</summary>
        internal event DataGridViewCellFormattingEventHandler? OnExternalCellFormatting;

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _linkToolTip?.Dispose();
                _linkToolTip = null;
                _btnHistory?.Dispose();
                _btnHistory = null;
            }
            base.Dispose(disposing);
        }
    }
}
