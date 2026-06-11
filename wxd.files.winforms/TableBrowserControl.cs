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
using MiniExcelLibs;

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
        private Label _lblLock = new();
        private Label _lblFileName = new();
        private Label _lblInfo = new();
        private Label _lblCellPos = new();
        private Panel _pnlLoading = new();
        private Panel _pnlTop = new();
        private Panel _pnlStatus = new();
        private DataTable _dataTable = new();
        private DataTable _headerTable = new();
        private FileTableData _fileData = new();  // 统一数据模型，用于行号追踪
        private string _filePath = "";
        private TableFileType _fileType;
        private char _delimiter = ',';
        private int _frozenRows = 0;
        private bool _frozenRowsOverridden = false;
        private int _headerRowIndex = -1;
        private int _uniqueColumnIndex = -1;
        private int _mappingRowIndex = -1;
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
        private ComboBox? _cboUniqueColumn;
        private bool[]? _externalColumnStates = null;
        private bool _isEmbeddedMode; // 是否为外部嵌入模式（版本对比等），禁止编辑
        private bool _isEditLocked = false; // Excel默认锁定编辑，CSV默认解锁

        // ==================== 统一缓存 ====================
        private DirectoryCache? _dirCache; // 目录级缓存（由 FrozenDir 决定）
        private bool _hasUnsavedChanges = false;
        private System.Windows.Forms.Timer? _retrySaveTimer;
        private const int SaveDebounceMs = 3000; // 3秒无操作后自动保存

        // ==================== 行编辑模式 ====================
        private int _editingRowIndex = -1;           // 当前正在编辑的显示行索引（-1=无）
        private string[]? _editingRowOriginalValues;  // 编辑前原始值快照（用于Esc取消）
        private bool _enteringEditMode = false;       // 防止事件重入标志
        private bool _isNewRowForEdit = false;        // 是否为新增行（Esc时删除该行）
        private int _editingRowOrigIndex = -1;        // 编辑行的原始行索引
        private bool _syncingColumnWidth = false;     // 列宽同步防重入
        private bool _syncingScroll = false;           // 滚动同步防重入

        // ==================== 过滤防抖 ====================
        private CancellationTokenSource? _filterDebounceCts;
        private const int FilterDebounceMs = 300;

        /// <summary>唯一列重复警告文本（null = 无警告），用于状态栏持久提示</summary>
        private string? _duplicateWarningText = null;
        /// <summary>唯一列重复值完整列表（供详情窗体使用）</summary>
        private List<DuplicateGroup> _duplicateDetails = new();

        /// <summary>重复值分组数据</summary>
        public class DuplicateGroup
        {
            public string Value { get; set; } = "";
            public int Count { get; set; }
            public List<int> RowIndices { get; set; } = new();
        }

        // ==================== 公开属性 ====================
        public string FilePath => _filePath;
        public TableFileType FileType => _fileType;
        public int UniqueColumnIndex => _uniqueColumnIndex;
        public int MappingRowIndex => _mappingRowIndex;

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
        /// <summary>使用配置对象构造（推荐）</summary>
        public TableBrowserControl(TableBrowserOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            _filePath = options.FilePath;
            _paneIndex = options.PaneIndex;
            _rootDir = options.RootDir;
            _lightweight = options.Lightweight;
            // ★ 先确定文件类型，后续配置读取需要用到
            _fileType = DetectFileType(options.FilePath);
            _isEditLocked = (_fileType == TableFileType.Excel); // Excel默认锁定编辑
            // ★ 初始化目录级缓存
            if (!string.IsNullOrEmpty(_rootDir))
                _dirCache = FileDataCache.GetDirectoryCache(_rootDir);
            if (options.FrozenRows >= 0) { _frozenRows = options.FrozenRows; _frozenRowsOverridden = true; }
            _initialContentFilter = options.InitialContentFilter;

            var dir = FrozenDir;

            // ★ 表头行：优先使用传入值，否则从配置回退读取（不受 dir 为空限制）
            if (options.HeaderRowIndex >= 0)
                _headerRowIndex = options.HeaderRowIndex;
            else if (!string.IsNullOrEmpty(dir) && _fileType == TableFileType.Csv)
                _headerRowIndex = CsvSettings.GetCsvHeaderRowIndexByDirectory(dir);
            else if (!string.IsNullOrEmpty(dir))
                _headerRowIndex = CsvSettings.GetExcelHeaderRowIndexByDirectory(dir);

            // ★ 唯一列：传入值 > 文件级配置 > 目录级配置 > 默认 0
            if (options.UniqueColumnIndex >= 0)
            {
                _uniqueColumnIndex = options.UniqueColumnIndex;
            }
            else if (!string.IsNullOrEmpty(dir))
            {
                var relPath = GetCurrentRelativePath();
                var fileCol = CsvSettings.GetFileUniqueColumn(dir, relPath, _fileType);
                if (fileCol >= 0)
                    _uniqueColumnIndex = fileCol;
                else if (_fileType == TableFileType.Csv)
                    _uniqueColumnIndex = CsvSettings.GetCsvUniqueColumnIndexByDirectory(dir);
                else
                    _uniqueColumnIndex = CsvSettings.GetExcelUniqueColumnIndexByDirectory(dir);

                if (_uniqueColumnIndex < 0)
                    _uniqueColumnIndex = 0; // 默认第一列
            }

            // ★ 映射行：优先使用传入值，否则从配置回退读取（不受 dir 为空限制）
            if (options.MappingRowIndex >= 0)
                _mappingRowIndex = options.MappingRowIndex;
            else if (!string.IsNullOrEmpty(dir) && _fileType == TableFileType.Csv)
                _mappingRowIndex = CsvSettings.GetCsvMappingRowIndexByDirectory(dir);
            else if (!string.IsNullOrEmpty(dir))
                _mappingRowIndex = CsvSettings.GetExcelMappingRowIndexByDirectory(dir);

            // 文件备注名仍然需要从配置读取（依赖 dir）
            if (!string.IsNullOrEmpty(dir))
            {
                _fileRemarks = CsvSettings.GetFileRemarksByDirectory(dir);
            }

            InitializeComponent();

            // 如果文件路径为空或文件不存在，跳过自动加载（由外部通过 SetDataTable 注入数据）
            if (!string.IsNullOrEmpty(_filePath) && File.Exists(_filePath))
            {
                if (_fileType == TableFileType.Csv)
                    LoadCsvAsync();
                else
                    LoadExcelAsync();
            }
            else
            {
                // 空数据模式：创建一个空的 DataTable，等待外部 SetDataTable
                _dataTable = new DataTable();
                _fileData = new FileTableData();
                _headerTable = _dataTable.Clone();
                _fileType = TableFileType.Csv; // 默认类型
            }
        }

        /// <summary>旧构造函数（已废弃，委托给新构造函数）</summary>
        [Obsolete("请使用 TableBrowserControl(TableBrowserOptions) 代替")]
        public TableBrowserControl(string filePath, int paneIndex = -1, string? rootDir = null,
            int frozenRows = -1, int headerRowIndex = -1, string initialContentFilter = "",
            bool lightweight = false, int uniqueColumnIndex = -1, int mappingRowIndex = -1)
            : this(new TableBrowserOptions(filePath)
            {
                PaneIndex = paneIndex,
                RootDir = rootDir,
                FrozenRows = frozenRows,
                HeaderRowIndex = headerRowIndex,
                InitialContentFilter = initialContentFilter,
                Lightweight = lightweight,
                UniqueColumnIndex = uniqueColumnIndex,
                MappingRowIndex = mappingRowIndex,
            })
        { }

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
                // 如果 ExtractValues 只返回了一个值，但原始值包含多层切割，尝试用 ExtractSteps 手动展开
                if (_highlightValues.Count == 1 && linkConfig.ExtractSteps.Count > 0)
                {
                    var fallback = new List<string>();
                    // 只对第一层做切割展开（后续层交给 ExtractValues）
                    var firstStep = linkConfig.ExtractSteps[0];
                    if (!string.IsNullOrEmpty(firstStep.Delimiter))
                    {
                        var groups = value.Split(new[] { firstStep.Delimiter }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var g in groups)
                        {
                            var trimmed = g.Trim();
                            if (!string.IsNullOrEmpty(trimmed))
                            {
                                if (linkConfig.ExtractSteps.Count > 1)
                                {
                                    // 有后续步骤：对每组应用剩余步骤
                                    var subSteps = linkConfig.ExtractSteps.Skip(1).ToList();
                                    var tempConfig = new ColumnLinkConfig
                                    {
                                        ValueExtractionEnabled = true,
                                        ExtractSteps = subSteps
                                    };
                                    var subValues = tempConfig.ExtractValues(trimmed);
                                    fallback.AddRange(subValues);
                                }
                                else
                                {
                                    fallback.Add(trimmed);
                                }
                            }
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
                    : CsvSettings.DefaultValue;
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
            cboHeaderRow.Items.Add("未设置");
            for (int i = 0; i < 10; i++) cboHeaderRow.Items.Add($"第{i + 1}行");
            cboHeaderRow.SelectedIndex = _headerRowIndex >= 0 && _headerRowIndex < 10 ? _headerRowIndex + 1 : 0;
            cboHeaderRow.SelectedIndexChanged += (s, e) =>
            {
                _headerRowIndex = cboHeaderRow.SelectedIndex > 0 ? cboHeaderRow.SelectedIndex - 1 : -1;
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

            // 映射行选择
            var lblMappingRow = new Label
            {
                Text = "映射行:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };
            var cboMappingRow = new ComboBox
            {
                Size = new Size(70, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 10, 0)
            };
            cboMappingRow.Items.Add("未设置");
            int maxMappingRow = Math.Min(_frozenRows, 10);
            for (int i = 0; i < maxMappingRow; i++) cboMappingRow.Items.Add($"第{i + 1}行");
            cboMappingRow.SelectedIndex = _mappingRowIndex >= 0 && _mappingRowIndex < maxMappingRow ? _mappingRowIndex + 1 : 0;
            cboMappingRow.SelectedIndexChanged += (s, e) =>
            {
                _mappingRowIndex = cboMappingRow.SelectedIndex > 0 ? cboMappingRow.SelectedIndex - 1 : -1;
                if (!string.IsNullOrEmpty(FrozenDir))
                {
                    if (_fileType == TableFileType.Csv)
                        CsvSettings.SetCsvMappingRowIndexByDirectory(FrozenDir, _mappingRowIndex);
                    else
                        CsvSettings.SetExcelMappingRowIndexByDirectory(FrozenDir, _mappingRowIndex);
                }
                CsvSettings.Save();
                RefreshGridLayout();
            };

            // 唯一列选择
            var lblUniqueColumn = new Label
            {
                Text = "唯一列:",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, 5, 3, 0)
            };
            _cboUniqueColumn = new ComboBox
            {
                Size = new Size(130, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 10, 0)
            };
            RefreshUniqueColumnCombo();
            _cboUniqueColumn.SelectedIndexChanged += (s, e) =>
            {
                _uniqueColumnIndex = _cboUniqueColumn.SelectedIndex > 0 ? _cboUniqueColumn.SelectedIndex - 1 : -1;
                if (!string.IsNullOrEmpty(FrozenDir))
                {
                    var relPath = GetCurrentRelativePath();
                    CsvSettings.SetFileUniqueColumn(FrozenDir, relPath, _uniqueColumnIndex);
                }
                CsvSettings.Save();
                CheckUniqueColumnDuplicates(showMessageBox: true);
            };

            // 列管理按钮
            var lblHeaderFilter = new Label
            {
                Text = "列管理:",
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
                Text = "列管理",
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

            _flowTop.Controls.AddRange(new Control[] { lblFrozenRows, cboFrozenRows, lblHeaderRow, cboHeaderRow,
                lblMappingRow, cboMappingRow, lblUniqueColumn, _cboUniqueColumn, lblHeaderFilter, _btnSelectColumns });

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

            // ★ 初始化保存防抖定时器
            _retrySaveTimer = new System.Windows.Forms.Timer { Interval = SaveDebounceMs };
            _retrySaveTimer.Tick += (_, _) =>
            {
                _retrySaveTimer.Stop();
                TryFlushToDisk();
            };

            var lblCellPos = new Label
            {
                Text = "",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                ForeColor = _colorPrimary,
                Margin = new Padding(0, 3, 10, 3)
            };
            _lblCellPos = lblCellPos;
            // _lblCellPos 和 _lblInfo 不再添加到 _flowTop，改为放入底部状态栏 _pnlStatus

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
            txtContentFilter.TextChanged += (s, e) => DebounceFilter();
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
            txtAdvancedFilter.TextChanged += (s, e) => DebounceFilter();
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
                Dock = DockStyle.Fill,
                Height = 30,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8, 1, 8, 0),
                AutoSize = true
            };
            flowContent.Controls.AddRange(new Control[] { lblContentFilter, txtContentFilter });

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
            _dataGridHeader.AllowUserToResizeColumns = true;
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
                // 同步水平滚动偏移到内容行（键盘方向键在固定行导航时联动滚动）
                if (_dataGrid.HorizontalScrollingOffset != _dataGridHeader.HorizontalScrollingOffset)
                    _dataGrid.HorizontalScrollingOffset = _dataGridHeader.HorizontalScrollingOffset;
            };
            _dataGridHeader.CellDoubleClick += OnFrozenCellDoubleClick;
            _dataGridHeader.CellClick += OnCellClick;

            _dataGrid = CreateDataGridView();
            _dataGrid.Dock = DockStyle.Fill;
            _dataGrid.RowHeadersWidth = 50;
            _dataGrid.RowPostPaint += OnRowPostPaint;
            _dataGrid.ColumnHeadersVisible = false;
            _dataGrid.CellMouseDown += OnGridCellMouseDown;

            // ==================== 行编辑模式：点击其他行自动退出编辑 ====================
            _dataGrid.SelectionChanged += (s, e) =>
            {
                if (_editingRowIndex < 0) return;
                if (_enteringEditMode) return;

                // 如果当前选中行不是编辑行，则提交退出编辑
                if (_dataGrid.CurrentRow != null && _dataGrid.CurrentRow.Index != _editingRowIndex)
                {
                    _enteringEditMode = true;
                    try
                    {
                        // 使用 CancelEdit 而非 EndEdit，避免触发 OnCellEndEdit 中的 CurrentCell 跳转（防止重入）
                        if (_dataGrid.IsCurrentCellInEditMode)
                            _dataGrid.CancelEdit();
                        ExitRowEditMode(commit: true);
                    }
                    finally { _enteringEditMode = false; }
                }
            };

            _dataGrid.Scroll += (s, e) =>
            {
                if (_syncingScroll) return;
                _syncingScroll = true;
                try { _dataGridHeader.HorizontalScrollingOffset = _dataGrid.HorizontalScrollingOffset; }
                finally { _syncingScroll = false; }
            };
            _dataGrid.ColumnWidthChanged += (s, e) =>
            {
                if (_syncingColumnWidth) return;
                _syncingColumnWidth = true;
                try
                {
                    // 全量同步所有列宽（避免键盘导航后索引不一致导致错位）
                    for (int i = 0; i < Math.Min(_dataGrid.Columns.Count, _dataGridHeader.Columns.Count); i++)
                        _dataGridHeader.Columns[i].Width = _dataGrid.Columns[i].Width;
                }
                finally { _syncingColumnWidth = false; }
            };

            // ★ 固定行滚动同步到内容行
            _dataGridHeader.Scroll += (s, e) =>
            {
                if (_syncingScroll) return;
                _syncingScroll = true;
                try { _dataGrid.HorizontalScrollingOffset = _dataGridHeader.HorizontalScrollingOffset; }
                finally { _syncingScroll = false; }
            };

            // ★ 固定行列宽变化同步到内容行
            _dataGridHeader.ColumnWidthChanged += (s, e) =>
            {
                if (_syncingColumnWidth) return;
                _syncingColumnWidth = true;
                try
                {
                    for (int i = 0; i < Math.Min(_dataGrid.Columns.Count, _dataGridHeader.Columns.Count); i++)
                        _dataGrid.Columns[i].Width = _dataGridHeader.Columns[i].Width;
                }
                finally { _syncingColumnWidth = false; }
            };

            // 快捷键：Ctrl+D 复制当前行，Delete 删除当前行
            _dataGrid.KeyDown += (s, e) =>
            {
                if (_isEmbeddedMode) return;
                if (_dataGrid.CurrentRow == null) return;

                int displayIdx = _dataGrid.CurrentRow.Index;
                int origIdx = GetOriginalRowIndexFromDisplay(displayIdx);

                // ==================== 行编辑模式快捷键 ====================
                if (_editingRowIndex >= 0)
                {
                    if (e.KeyCode == Keys.Escape)
                    {
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                        CancelRowEditMode();
                        return;
                    }
                    if (e.KeyCode == Keys.Enter && !e.Control)
                    {
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                        // 先结束当前单元格编辑
                        if (_dataGrid.IsCurrentCellInEditMode)
                            _dataGrid.EndEdit();
                        ExitRowEditMode(commit: true);
                        return;
                    }
                    // 行编辑模式下禁用行操作快捷键
                    if (e.Control && e.KeyCode == Keys.D) { e.Handled = true; return; }
                    if (e.KeyCode == Keys.Delete && !e.Control) { e.Handled = true; return; }
                    return;
                }

                if (e.Control && e.KeyCode == Keys.D)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    DuplicateRowAt(origIdx);
                }
                else if (e.KeyCode == Keys.Delete && !e.Control)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    DeleteRowAt(origIdx);
                }
            };

            splitContainer.Panel1.Controls.Add(_dataGridHeader);
            splitContainer.Panel2.Controls.Add(_dataGrid);
            splitContainer.Panel1Collapsed = true;
            _pnlGridContainer = splitContainer;

            Controls.Add(_pnlGridContainer);
            Controls.Add(_pnlTop);

            // ============ 底部状态栏 ============
            // 状态栏顶部分隔线
            var statusSeparator = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 1,
                BackColor = Color.FromArgb(200, 225, 245),
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            Controls.Add(statusSeparator);

            _pnlStatus = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                BackColor = Color.FromArgb(245, 250, 255),
                Padding = new Padding(0)
            };

            // 使用 TableLayoutPanel 来实现四列布局
            var statusTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                BackColor = Color.Transparent,
                Padding = new Padding(4, 0, 4, 0),
                Margin = new Padding(0)
            };
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));    // 锁图标
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));  // 文件名
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));  // 行列坐标（固定宽度足够）
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));   // 重复提示（占满剩余空间）

            // 列0: 锁图标 🔒/🔓
            _lblLock = new Label
            {
                Text = "",
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = _colorLightTextSecondary,
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                Padding = new Padding(0, 0, 4, 0),
            };
            _lblLock.Click += OnLockLabelClick;
            statusTable.Controls.Add(_lblLock, 0, 0);

            // 列1: 文件名
            _lblFileName = new Label
            {
                Text = "",
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = _colorLightTextSecondary,
                Font = new Font("Microsoft YaHei UI", 9f),
                Dock = DockStyle.Fill,
            };
            statusTable.Controls.Add(_lblFileName, 1, 0);

            // 列2: _lblCellPos 行列坐标
            _lblCellPos.TextAlign = ContentAlignment.MiddleLeft;
            _lblCellPos.Dock = DockStyle.Fill;
            statusTable.Controls.Add(_lblCellPos, 2, 0);

            // 列3: _lblInfo 重复提示（左侧，紧跟行列坐标）
            _lblInfo.TextAlign = ContentAlignment.MiddleLeft;
            _lblInfo.Dock = DockStyle.Fill;
            statusTable.Controls.Add(_lblInfo, 3, 0);

            _pnlStatus.Controls.Add(statusTable);
            Controls.Add(_pnlStatus);

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
                ReadOnly = true,  // 默认只读（浏览模式），行编辑模式下临时设为false
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
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
            dgv.CellBeginEdit += OnCellBeginEdit;
            dgv.CellEndEdit += OnCellEndEdit;
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
            // ★ 同步构建 FileTableData（继承旧差异信息）
            _fileData = BuildFileData(dataTable, _fileData);
            _headerTable = headerTable.Clone();

            _columnCheckStates = new bool[_dataTable.Columns.Count];
            var relPath = GetCurrentRelativePath();
            var savedStates = CsvSettings.GetColumnVisibility(FrozenDir, relPath, _dataTable.Columns.Count);
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
            // ★ 使用目录级缓存
            if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
            // ★ 后台异步构建关联 lookup，不阻塞表格显示
            if (_dirCache != null && _columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_dirCache, _columnLinks, _filePath);

            UpdateColumnSelectorButton();
            RebuildHighlightRows();
            _lblFileName.Text = GetDisplayFileName();
            UpdateLockDisplay();
            _lblCellPos.Text = infoText;
            _lblInfo.Text = "";

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

            // ★ 数据加载完成后刷新唯一列下拉（触发 SelectedIndexChanged 自动检查重复）
            RefreshUniqueColumnCombo();

            // 触发数据加载完成事件
            DataLoaded?.Invoke(this, EventArgs.Empty);

            // 检测是否 SVN/Git 管理，如果是则显示"历史记录"按钮（轻量模式跳过）
            if (!_lightweight)
                ShowHistoryButtonIfVcs();
        }

        /// <summary>为自定义 tooltip 提供内容（grid: 当前鼠标所在的 DataGridView）</summary>
        private (string? title, List<DataGridLinkToolTip.TipLine> lines, string? rawText) GetLinkToolTipContent(DataGridView grid, int rowIndex, int columnIndex)
        {
            // 列标题悬停（rowIndex == -1）
            if (rowIndex < 0 && columnIndex >= 0)
            {
                var dgvCol = grid.Columns[columnIndex];
                var headerVal = dgvCol.HeaderText;

                // 计算原始列索引
                int colIdx = columnIndex;
                if (_visibleColumnIndices.Count > 0 && columnIndex < _visibleColumnIndices.Count)
                    colIdx = _visibleColumnIndices[columnIndex];

                // 构建列标题 ToolTip
                string tipText = headerVal;
                if (_mappingRowIndex >= 0 && _mappingRowIndex < _dataTable.Rows.Count
                    && colIdx < _dataTable.Columns.Count)
                {
                    var mappingVal = _dataTable.Rows[_mappingRowIndex][colIdx]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(mappingVal))
                        tipText = $"{mappingVal} ({headerVal}) - 第{_mappingRowIndex + 1}行";
                }

                return (null, new List<DataGridLinkToolTip.TipLine>(), tipText);
            }

            if (rowIndex < 0 || columnIndex < 0)
                return (null, new List<DataGridLinkToolTip.TipLine>(), null);

            bool isHeader = grid == _dataGridHeader;

            // 计算原始列索引
            int originalColIndex = columnIndex;
            if (_visibleColumnIndices.Count > 0 && columnIndex < grid.Columns.Count)
            {
                // ★ 不再需要 __RowNum__ 偏移
                int adjustedCol = columnIndex;
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
                // 如果设置了映射行，拼接映射名 → "映射名(表头名)"
                if (_mappingRowIndex >= 0 && _mappingRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
                {
                    var mappingVal = _dataTable.Rows[_mappingRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(mappingVal))
                    {
                        if (!string.IsNullOrEmpty(headerName))
                            headerName = $"{mappingVal}({headerName})";
                        else
                            headerName = mappingVal;
                    }
                }
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
                // 如果设置了映射行，拼接映射名 → "映射名(表头名)"
                if (_mappingRowIndex >= 0 && _mappingRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
                {
                    var mappingVal = _dataTable.Rows[_mappingRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(mappingVal))
                    {
                        if (!string.IsNullOrEmpty(headerName))
                            headerName = $"{mappingVal}({headerName})";
                        else
                            headerName = mappingVal;
                    }
                }
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

            // ★ 查找缓存（使用 _dirCache 替代 ColumnLinkChecker.GetLookups）
            if (_dirCache == null) return (null, new List<DataGridLinkToolTip.TipLine>(), null);

            // ★ 检查关联 lookup 是否正在后台构建中，返回友好提示
            if (_dirCache.IsLinkLookupBuilding(currentRel))
            {
                // 获取列名
                string loadingHeaderName = "";
                if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
                    loadingHeaderName = _dataTable.Rows[_headerRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
                if (_mappingRowIndex >= 0 && _mappingRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
                {
                    var mappingVal = _dataTable.Rows[_mappingRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(mappingVal))
                    {
                        if (!string.IsNullOrEmpty(loadingHeaderName))
                            loadingHeaderName = $"{mappingVal}({loadingHeaderName})";
                        else
                            loadingHeaderName = mappingVal;
                    }
                }
                if (string.IsNullOrEmpty(loadingHeaderName))
                    loadingHeaderName = GetExcelColumnName(originalColIndex);

                string loadingDisplayName = Path.GetFileName(linkConfig.TargetFile);
                if (_fileRemarks.TryGetValue(linkConfig.TargetFile, out var loadingRemark) && !string.IsNullOrEmpty(loadingRemark))
                    loadingDisplayName = $"{loadingRemark} ({Path.GetFileName(linkConfig.TargetFile)})";

                string loadingTitle = $"{loadingDisplayName}\n来自: {loadingHeaderName} — 第{realRowIndex + 1}行";
                var loadingLines = new List<DataGridLinkToolTip.TipLine>
                {
                    new() { Text = "  关联数据加载中，请稍候...", IsWarning = false }
                };
                return (loadingTitle, loadingLines, null);
            }

            string prefix = currentRel + "|" + originalColIndex + "->";
            LinkLookupEntry? lookup = null;
            foreach (var kvp in _dirCache.GetLookups())
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

            // 获取列名
            string colHeaderName = "";
            if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
                colHeaderName = _dataTable.Rows[_headerRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
            // 如果设置了映射行，拼接映射名 → "映射名(表头名)"
            if (_mappingRowIndex >= 0 && _mappingRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
            {
                var mappingVal = _dataTable.Rows[_mappingRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(mappingVal))
                {
                    if (!string.IsNullOrEmpty(colHeaderName))
                        colHeaderName = $"{mappingVal}({colHeaderName})";
                    else
                        colHeaderName = mappingVal;
                }
            }
            if (string.IsNullOrEmpty(colHeaderName))
                colHeaderName = GetExcelColumnName(originalColIndex);

            string title = $"{displayTableName}\n来自: {colHeaderName} — 第{realRowIndex + 1}行";

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
            // ★ 使用目录级缓存
            if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
            if (_dirCache != null && _columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_dirCache, _columnLinks, _filePath);
            RebuildHighlightRows();
            RefreshGridLayout();

            // ★ 重新加载后重新检测唯一列重复（首次加载，弹 MessageBox）
            CheckUniqueColumnDuplicates(showMessageBox: true);
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

            this.Invoke(() => { _delimiter = delimiter; _dataTable = dataTable; _fileData = BuildFileData(dataTable, _fileData); _headerTable = dataTable.Clone(); });
        }

        // ==================== 唯一列下拉刷新 ====================
        private void RefreshUniqueColumnCombo()
        {
            if (_cboUniqueColumn == null) return;

            _cboUniqueColumn.Items.Clear();
            _cboUniqueColumn.Items.Add("未设置");
            for (int col = 0; col < _dataTable.Columns.Count; col++)
            {
                string displayName;
                if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && _headerRowIndex < _frozenRows)
                    displayName = _dataTable.Rows[_headerRowIndex][col]?.ToString() ?? $"Col{col}";
                else
                    displayName = $"Col{col}";
                _cboUniqueColumn.Items.Add($"{displayName} (Col{col})");
            }
            // 根据 _uniqueColumnIndex 恢复选中状态；未设置时默认选中第一列
            _cboUniqueColumn.SelectedIndex = _uniqueColumnIndex >= 0 && _uniqueColumnIndex < _dataTable.Columns.Count
                ? _uniqueColumnIndex + 1
                : (_dataTable.Columns.Count > 0 ? 1 : 0);
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
                        // 如果设置了表头行，用表头行的值作为列标题
                        if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && i < _visibleColumnIndices.Count)
                        {
                            var srcColIdx = _visibleColumnIndices[i];
                            var headerVal = _dataTable.Rows[_headerRowIndex][srcColIdx]?.ToString() ?? GetExcelColumnName(i);
                            dgvCol.HeaderText = headerVal;
                            // 如果设置了映射行，设置 ToolTip 为 "映射名 (表头名)"
                            SetColumnHeaderToolTip(dgvCol, i, headerVal);
                        }
                        else
                        {
                            dgvCol.HeaderText = GetExcelColumnName(i);
                        }
                        dgvCol.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                }

                var contentFilter = _txtContentFilter.Text.Trim();
                var visibleRowIndices = GetVisibleRowIndices(contentFilter);
                RefreshFilteredDisplay(visibleRowIndices);
            }
            else
            {
                _pnlGridContainer.Panel1Collapsed = true;
                _dataGrid.DataSource = _dataTable;
                for (int i = 0; i < _dataGrid.Columns.Count; i++)
                {
                    _dataGrid.Columns[i].Width = 120;
                    _dataGrid.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
                    // 如果设置了表头行，用表头行的值作为列标题
                    if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count)
                    {
                        var headerVal = _dataTable.Rows[_headerRowIndex][i]?.ToString() ?? GetExcelColumnName(i);
                        _dataGrid.Columns[i].HeaderText = headerVal;
                        SetColumnHeaderToolTip(_dataGrid.Columns[i], i, headerVal);
                    }
                    else
                    {
                        _dataGrid.Columns[i].HeaderText = GetExcelColumnName(i);
                    }
                }
                // ★ 同步设置 Tag（DataSource 赋值后行已创建）
                SetRowTags();
            }
        }

        /// <summary>将 _fileData.Rows 挂到 _dataGrid 每行的 Tag 上</summary>
        private void SetRowTags()
        {
            for (int i = 0; i < _dataGrid.Rows.Count && i < _fileData.Rows.Count; i++)
                _dataGrid.Rows[i].Tag = _fileData.Rows[i];
        }

        /// <summary>列标题 ToolTip 已统一由 DataGridLinkToolTip 自定义控件处理，此方法保留为空</summary>
        private void SetColumnHeaderToolTip(DataGridViewColumn dgvCol, int displayColIndex, string headerVal)
        {
            // 不再设置原生 ToolTipText，列标题悬停由 DataGridLinkToolTip.OnGridCellMouseMove 统一处理
        }

        /// <summary>检查唯一列是否有重复值，如有则弹窗警告并在状态栏显示红色提示</summary>
        /// <returns>true = 存在重复值</returns>
        /// <param name="showMessageBox">是否弹出 MessageBox（仅首次加载/手动切换唯一列时传 true）</param>
        private bool CheckUniqueColumnDuplicates(bool showMessageBox = false)
        {
            // ★ 先清除之前的警告状态
            _duplicateWarningText = null;
            _duplicateDetails.Clear();

            if (_uniqueColumnIndex < 0 || _uniqueColumnIndex >= _dataTable.Columns.Count)
            { RestoreInfoStyle(); return false; }
            if (_dataTable.Rows.Count <= _frozenRows)
            { RestoreInfoStyle(); return false; }

            // ★ 获取列显示名（映射名(列名) 格式）
            string colDisplayName = $"Col{_uniqueColumnIndex}";
            if (_mappingRowIndex >= 0 && _mappingRowIndex < _dataTable.Rows.Count)
            {
                var mappingVal = _dataTable.Rows[_mappingRowIndex][_uniqueColumnIndex]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(mappingVal))
                {
                    if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count)
                    {
                        var headerVal = _dataTable.Rows[_headerRowIndex][_uniqueColumnIndex]?.ToString()?.Trim() ?? "";
                        colDisplayName = $"{mappingVal}({headerVal})";
                    }
                    else
                        colDisplayName = mappingVal;
                }
            }
            else if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count)
            {
                colDisplayName = _dataTable.Rows[_headerRowIndex][_uniqueColumnIndex]?.ToString()?.Trim() ?? $"Col{_uniqueColumnIndex}";
            }

            var duplicates = _dataTable.AsEnumerable()
                .Skip(_frozenRows)
                .Select((r, idx) => (Row: r, DataTableIndex: _frozenRows + idx))
                .GroupBy(x => x.Row[_uniqueColumnIndex]?.ToString()?.Trim() ?? "")
                .Where(g => !string.IsNullOrEmpty(g.Key) && g.Count() > 1)
                .OrderByDescending(g => g.Count())
                .ToList();

            if (!duplicates.Any())
            {
                RestoreInfoStyle();
                return false;
            }

            // ★ 保存完整数据
            _duplicateDetails = duplicates.Select(g => new DuplicateGroup
            {
                Value = g.Key,
                Count = g.Count(),
                RowIndices = g.Select(x => x.DataTableIndex).ToList(),
            }).ToList();

            var totalDupCount = duplicates.Count;
            var totalRedundant = duplicates.Sum(g => g.Count() - 1);

            // ★ MessageBox：只在首次加载或手动切换唯一列时弹出
            if (showMessageBox)
            {
                var sb = new StringBuilder();
                sb.AppendLine($"列 \"{colDisplayName}\" 存在重复值：");
                sb.AppendLine($"共 {totalDupCount} 个重复值，{totalRedundant} 处冗余");
                sb.AppendLine();

                int showGroups = Math.Min(3, duplicates.Count);
                for (int i = 0; i < showGroups; i++)
                {
                    var g = duplicates[i];
                    sb.AppendLine($"▸ 值 \"{g.Key}\" 重复 {g.Count()} 次：");
                    int showRows = Math.Min(5, g.Count());
                    foreach (var x in g.Take(showRows))
                    {
                        int fileRowNum = x.DataTableIndex < _fileData.Rows.Count
                            ? _fileData.Rows[x.DataTableIndex].OriginalRowIndex + 1
                            : x.DataTableIndex + 1;
                        sb.AppendLine($"    行 {fileRowNum}");
                    }
                    if (g.Count() > 5)
                        sb.AppendLine($"    ... 还有 {g.Count() - 5} 行");
                }
                if (duplicates.Count > 3)
                    sb.AppendLine($"... 还有 {duplicates.Count - 3} 个重复值组未显示");
                sb.AppendLine();
                sb.AppendLine("点击状态栏可查看完整详情");

                MessageBox.Show(this, sb.ToString(), "唯一列重复警告",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // ★ 状态栏：简短摘要 + 可点击（始终更新）
            _duplicateWarningText = totalDupCount <= 3
                ? $"⚠ 唯一列\"{colDisplayName}\"有{totalDupCount}个重复值({totalRedundant}处冗余)"
                : $"⚠ 唯一列\"{colDisplayName}\"有{totalDupCount}个重复值({totalRedundant}处冗余)…点击查看详情";

            ApplyDuplicateWarningStyle();
            return true;
        }

        /// <summary>应用重复警告样式到状态栏</summary>
        private void ApplyDuplicateWarningStyle()
        {
            if (_duplicateWarningText == null) return;
            _lblInfo.ForeColor = Color.FromArgb(200, 30, 30);
            _lblInfo.Text = _duplicateWarningText;
            _lblInfo.Cursor = Cursors.Hand;
            _lblInfo.Click += LblInfo_Click_ShowDuplicates;
            _pnlStatus.BackColor = Color.FromArgb(255, 235, 235);
        }

        /// <summary>恢复状态栏正常样式</summary>
        private void RestoreInfoStyle()
        {
            _duplicateWarningText = null;
            _duplicateDetails.Clear();
            _lblInfo.Text = "";
            _lblInfo.ForeColor = _colorLightTextSecondary;
            _lblInfo.Cursor = Cursors.Default;
            _lblInfo.Click -= LblInfo_Click_ShowDuplicates;
            _pnlStatus.BackColor = Color.FromArgb(245, 250, 255);
        }

        /// <summary>点击状态栏标签 → 弹出重复值详情窗体</summary>
        private void LblInfo_Click_ShowDuplicates(object? sender, EventArgs e)
        {
            if (_duplicateDetails.Count == 0) return;
            var form = new DuplicateDetailForm(_duplicateDetails, _fileData, this);
            form.ShowDialog(this);
        }

        /// <summary>导航到指定行（供外部调用，如重复值详情窗体）</summary>
        public void NavigateToRow(int dataTableRowIndex)
        {
            if (_dataGrid.Rows.Count == 0) return;

            for (int i = 0; i < _dataGrid.Rows.Count; i++)
            {
                var row = _dataGrid.Rows[i];
                if (row.DataBoundItem is DataRowView drv)
                {
                    var dtRow = _dataTable.Rows.IndexOf(drv.Row);
                    if (dtRow == dataTableRowIndex)
                    {
                        _dataGrid.ClearSelection();
                        _dataGrid.Rows[i].Selected = true;
                        _dataGrid.FirstDisplayedScrollingRowIndex = Math.Max(0, i - 5);
                        _dataGrid.CurrentCell = _dataGrid.Rows[i].Cells[Math.Min(1, _dataGrid.Columns.Count - 1)];
                        return;
                    }
                }
            }
        }

        // ==================== 列选择 ====================
        public List<int> GetVisibleColumnIndices()
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
            using var dialog = new ColumnManagerDialog(_dataTable, _columnCheckStates, _frozenRows, 0,
                _rootDir ?? "", _filePath, _fileEntries, _columnLinks, _headerRowIndex,
                _mappingRowIndex);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _columnCheckStates = dialog.GetCheckedStates();
                _columnLinks = dialog.GetColumnLinks();
                CsvSettings.SetColumnLinksByDirectory(FrozenDir, _columnLinks);
                // ★ 使用目录级缓存
                if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                    _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
                if (_dirCache != null && _columnLinks.Count > 0)
                    FireAndForgetBuildLinkLookups(_dirCache, _columnLinks, _filePath);
                CsvSettings.SetColumnVisibility(FrozenDir, GetCurrentRelativePath(), _columnCheckStates);
                CsvSettings.Save();
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
            _btnSelectColumns.Text = $"列管理({checkedData}/{dataColumns})";
            _btnSelectColumns.ForeColor = checkedData < dataColumns ? Color.Red : _colorLightText;
        }

        // ==================== 行过滤 ====================
        private async void DebounceFilter()
        {
            _filterDebounceCts?.Cancel();
            _filterDebounceCts = new CancellationTokenSource();
            var token = _filterDebounceCts.Token;
            try
            {
                await Task.Delay(FilterDebounceMs, token);
                if (!token.IsCancellationRequested)
                    ApplyFilters();
            }
            catch (TaskCanceledException) { /* 被新输入取消 */ }
        }

        private void ApplyFilters() => RefreshGridLayout();

        /// <summary>过滤结果：返回行号列表（直接走 FileTableData.Cells，零装箱）</summary>
        private List<int> GetVisibleRowIndices(string contentFilter)
        {
            var result = new List<int>(_fileData.Rows.Count - _frozenRows);
            var advFilter = _txtAdvancedFilter.Text.Trim();

            // ★ 预编译高级过滤表达式（只解析一次）
            FilterNode? compiledExpr = null;
            if (!string.IsNullOrEmpty(advFilter))
                compiledExpr = CompileExpression(advFilter);

            // 内容过滤预处理
            bool hasContentFilter = !string.IsNullOrEmpty(contentFilter);
            string processedContentFilter = hasContentFilter
                ? contentFilter.Replace("||", " | ").Replace("&&", " & ")
                : "";

            for (int i = _frozenRows; i < _fileData.Rows.Count; i++)
            {
                var cells = _fileData.Rows[i].Cells;

                // 内容过滤
                if (hasContentFilter)
                {
                    bool contentMatch = false;
                    foreach (var colIdx in _visibleColumnIndices)
                    {
                        if (colIdx < cells.Count &&
                            MatchesFilter(cells[colIdx], processedContentFilter))
                        { contentMatch = true; break; }
                    }
                    if (!contentMatch) continue;
                }

                // 高级过滤
                if (compiledExpr != null && !EvaluateCompiled(cells, compiledExpr))
                    continue;

                result.Add(i);
            }
            return result;
        }

        // ==================== 预编译过滤表达式 AST ====================
        /// <summary>过滤表达式 AST 节点</summary>
        private abstract class FilterNode { }

        private class BinOpNode : FilterNode
        {
            public string Op = "";          // "and" / "or"
            public FilterNode? Left;
            public FilterNode? Right;
        }

        private class CondNode : FilterNode
        {
            public string? ColumnName;    // 原始列名字符串（指定列过滤时）
            public int ColumnIndex = -1;   // 预解析的列索引，-1 表示全列搜索
            public string CompareOp = "";  // "=", "!=", ">", "<", ">=", "<=", "like", "" (全列模糊)
            public string Value = "";      // 比较值（已去引号）
            public bool IsNumeric;
            public double NumericValue;
        }

        /// <summary>将字符串表达式编译为 AST（只解析一次，每行复用）</summary>
        private FilterNode? CompileExpression(string expr)
        {
            expr = expr.Trim();
            if (string.IsNullOrEmpty(expr)) return null;

            // 去外层括号
            while (expr.StartsWith("(") && expr.EndsWith(")") && IsBalanced(expr))
                expr = expr.Substring(1, expr.Length - 2).Trim();

            // 找最外层 or / and（or 优先短路）
            int parenDepth = 0, orPos = -1, andPos = -1;
            for (int i = 0; i < expr.Length; i++)
            {
                if (expr[i] == '(') parenDepth++;
                else if (expr[i] == ')') parenDepth--;
                else if (parenDepth == 0)
                {
                    if (i + 2 < expr.Length && char.ToLower(expr[i]) == 'o' && char.ToLower(expr[i + 1]) == 'r'
                        && (i + 2 >= expr.Length || !char.IsLetterOrDigit(expr[i + 2]))
                        && (i == 0 || !char.IsLetterOrDigit(expr[i - 1])))
                    { orPos = i; break; }
                    if (i + 3 < expr.Length && char.ToLower(expr[i]) == 'a' && char.ToLower(expr[i + 1]) == 'n'
                        && char.ToLower(expr[i + 2]) == 'd'
                        && (i + 3 >= expr.Length || !char.IsLetterOrDigit(expr[i + 3]))
                        && (i == 0 || !char.IsLetterOrDigit(expr[i - 1])))
                    { if (andPos < 0) andPos = i; i += 3; }
                }
            }

            if (orPos >= 0)
            {
                return new BinOpNode
                {
                    Op = "or",
                    Left = CompileExpression(expr.Substring(0, orPos).Trim())!,
                    Right = CompileExpression(expr.Substring(orPos + 2).Trim())!
                };
            }
            if (andPos >= 0)
            {
                return new BinOpNode
                {
                    Op = "and",
                    Left = CompileExpression(expr.Substring(0, andPos).Trim())!,
                    Right = CompileExpression(expr.Substring(andPos + 3).Trim())!
                };
            }

            return CompileCondition(expr);
        }

        private CondNode CompileCondition(string condition)
        {
            condition = condition.Trim();
            string op = ""; int opPos = -1;

            var likeIdx = condition.IndexOf(" like ", StringComparison.OrdinalIgnoreCase);
            if (likeIdx >= 0) { op = "like"; opPos = likeIdx; }
            else
            {
                // 按长度降序查找，避免 >= 被 > 误匹配
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
                // 无操作符 → 全列模糊匹配
                return new CondNode { ColumnIndex = -1, CompareOp = "", Value = condition };
            }

            var rawColumnName = condition.Substring(0, opPos).Trim();
            int opRealLen = op == "like" ? " like ".Length : op.Length;
            var rawValue = condition.Substring(opPos + opRealLen).Trim();
            var value = rawValue;
            if ((value.StartsWith("'") && value.EndsWith("'")) ||
                (value.StartsWith("\"") && value.EndsWith("\"")))
                value = value.Substring(1, value.Length - 2);

            // 保存原始列名，求值时再解析列索引（因为列名可能对应表头行）
            var columnName = rawColumnName.Trim();
            return new CondNode
            {
                ColumnName = !string.IsNullOrEmpty(columnName) && char.IsLetter(columnName[0]) ? columnName : null,
                ColumnIndex = -1,  // 求值时再解析
                CompareOp = op,
                Value = value,
                IsNumeric = op != "like" && double.TryParse(value, out var nv),
                NumericValue = op != "like" && double.TryParse(value, out nv) ? nv : 0
            };
        }

        /// <summary>对 AST 求值（零装箱，直接从 List&lt;string&gt; 读取）</summary>
        private bool EvaluateCompiled(List<string> cells, FilterNode node)
        {
            switch (node)
            {
                case BinOpNode bin:
                    bool left = EvaluateCompiled(cells, bin.Left!);
                    if (bin.Op == "or" && left) return true;   // OR 短路
                    if (bin.Op == "and" && !left) return false; // AND 短路
                    return EvaluateCompiled(cells, bin.Right!);

                case CondNode cond:
                    return EvaluateConditionFast(cells, cond);
            }
            return true;
        }

        private bool EvaluateConditionFast(List<string> cells, CondNode cond)
        {
            if (string.IsNullOrEmpty(cond.CompareOp))
            {
                // 无操作符：全列模糊匹配
                foreach (var colIdx in _visibleColumnIndices)
                {
                    if (colIdx < cells.Count &&
                        cells[colIdx].IndexOf(cond.Value, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
                return false;
            }

            // 有操作符 + 指定列名 → 解析列索引
            int colIndex = -1;
            if (cond.ColumnIndex >= 0)
            {
                colIndex = cond.ColumnIndex;
            }
            else if (!string.IsNullOrEmpty(cond.ColumnName))
            {
                colIndex = ResolveColumnIndex(cond.ColumnName);
                if (colIndex < 0)
                {
                    // 列名没找到 → 尝试全列模糊匹配（兼容旧行为）
                    foreach (var ci in _visibleColumnIndices)
                    {
                        if (ci < cells.Count &&
                            cells[ci].IndexOf(cond.Value, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                    return false;
                }
                // 缓存解析结果
                cond.ColumnIndex = colIndex;
            }
            else
            {
                // 无列名但有操作符 → 全列模糊匹配
                foreach (var ci in _visibleColumnIndices)
                {
                    if (ci < cells.Count &&
                        cells[ci].IndexOf(cond.Value, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
                return false;
            }

            if (colIndex < 0 || colIndex >= cells.Count) return false;
            string cellValue = cells[colIndex];
            return CompareValue(cellValue, cond.Value, cond.CompareOp);
        }

        private void RefreshFilteredDisplay(List<int> visibleRowIndices)
        {
            // ==================== 行编辑模式：刷新前退出编辑 ====================
            if (_editingRowIndex >= 0)
            {
                _enteringEditMode = true;
                try
                {
                    if (_dataGrid.IsCurrentCellInEditMode)
                        _dataGrid.EndEdit();
                    ExitRowEditMode(commit: true);
                }
                finally { _enteringEditMode = false; }
            }

            var filteredTable = new DataTable();
            // ★ 不再添加 __RowNum__ 隐藏列；行号通过 _fileData.Rows[origIdx].OriginalRowIndex 追踪
            foreach (var colIndex in _visibleColumnIndices)
            {
                var sourceCol = _dataTable.Columns[colIndex];
                filteredTable.Columns.Add(sourceCol.ColumnName, typeof(string)); // ★ 统一 string 避免类型冲突
            }

            foreach (var origRowIdx in visibleRowIndices)
            {
                var newRow = filteredTable.NewRow();
                for (int i = 0; i < _visibleColumnIndices.Count; i++)
                {
                    int colIdx = _visibleColumnIndices[i];
                    newRow[i] = colIdx < _dataTable.Columns.Count
                        ? (_dataTable.Rows[origRowIdx][colIdx] ?? "")
                        : "";
                }
                filteredTable.Rows.Add(newRow);
            }

            _dataGrid.DataSource = filteredTable;
            // ★ 同步设置 Tag（DataSource 赋值后行已创建），使用原始行索引
            for (int i = 0; i < _dataGrid.Rows.Count && i < visibleRowIndices.Count; i++)
            {
                int origIdx = visibleRowIndices[i];
                if (origIdx >= 0 && origIdx < _fileData.Rows.Count)
                    _dataGrid.Rows[i].Tag = _fileData.Rows[origIdx];
            }
            for (int i = 0; i < _dataGrid.Columns.Count; i++)
            {
                var colName = filteredTable.Columns[i].ColumnName;
                var dgvCol = _dataGrid.Columns[colName];
                if (dgvCol != null)
                {
                    dgvCol.DisplayIndex = i;
                    dgvCol.Width = 120;
                    dgvCol.SortMode = DataGridViewColumnSortMode.Programmatic;
                    // 如果设置了表头行，用表头行的值作为列标题
                    if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && i < _visibleColumnIndices.Count)
                    {
                        var srcColIdx = _visibleColumnIndices[i];
                        var headerVal = _dataTable.Rows[_headerRowIndex][srcColIdx]?.ToString() ?? GetExcelColumnName(i);
                        dgvCol.HeaderText = headerVal;
                        // 如果设置了映射行，设置 ToolTip 为 "映射名 (表头名)"
                        SetColumnHeaderToolTip(dgvCol, i, headerVal);
                    }
                    else
                    {
                        dgvCol.HeaderText = GetExcelColumnName(i);
                    }
                }
            }
            _lblCellPos.Text = $"显示 {_visibleColumnIndices.Count} 列, {visibleRowIndices.Count} 行";
            _lblInfo.Text = "";

            // ★ 如果有重复警告，追加警告文本并保持红色
            if (_duplicateWarningText != null)
            {
                _lblInfo.ForeColor = Color.FromArgb(200, 30, 30);
                _lblInfo.Text = _duplicateWarningText;
            }
        }

        // ==================== 排序 ====================
        private void OnColumnHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.ColumnIndex >= _dataGrid.Columns.Count) return;

            // ==================== 行编辑模式：排序前退出编辑 ====================
            if (_editingRowIndex >= 0)
                ExitRowEditMode(commit: true);

            if (_sortColumnIndex == e.ColumnIndex)
                _sortOrder = _sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            else { _sortColumnIndex = e.ColumnIndex; _sortOrder = SortOrder.Ascending; }
            var columnName = _dataGrid.Columns[e.ColumnIndex].DataPropertyName;
            var frozenRows = _dataTable.AsEnumerable().Take(_frozenRows).ToList();
            var dataRows = _dataTable.AsEnumerable().Skip(_frozenRows).ToList();
            var sortedDataRows = dataRows.OrderBy(row => row.Field<string>(columnName), new NaturalStringComparer(_sortOrder)).ToList();
            _dataTable = frozenRows.Concat(sortedDataRows).ToArray().CopyToDataTable();
            // ★ 排序后同步重建 FileTableData（继承旧差异信息，按 OriginalRowIndex 匹配）
            _fileData = BuildFileData(_dataTable, _fileData);
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
            // ★ 不再需要 __RowNum__ 偏移
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
                if (_isEmbeddedMode)
                {
                    // 历史对比等嵌入式模式：只读，不允许修改
                    var dialog = new CellEditDialog(currentValue, originalRowIndex, originalColumnIndex, _filePath, colName,
                        onSave: null, readOnly: true, readOnlyMessage: "当前状态不支持修改");
                    dialog.Show(this);
                }
                else
                {
                    var dialog = new CellEditDialog(currentValue, originalRowIndex, originalColumnIndex, _filePath, colName,
                        onSave: newValue =>
                    {
                        if (originalRowIndex >= 0 && originalRowIndex < _dataTable.Rows.Count &&
                            originalColumnIndex >= 0 && originalColumnIndex < _dataTable.Columns.Count)
                        {
                            // 1. 更新内存中的数据
                            _dataTable.Rows[originalRowIndex][originalColumnIndex] = newValue;

                            // 2. 回写文件
                            SaveCsvFile();

                            // 3. 只在 DataGridView 中更新对应的显示单元格，不重建整个 Grid
                            int displayColIndex = _visibleColumnIndices.Count > 0
                                ? _visibleColumnIndices.IndexOf(originalColumnIndex)
                                : originalColumnIndex;

                            if (displayColIndex >= 0)
                            {
                                for (int i = 0; i < _dataGrid.Rows.Count; i++)
                                {
                                    if (_dataGrid.Rows[i].Tag is TableRow tagRow && tagRow.OriginalRowIndex == originalRowIndex)
                                    {
                                        if (displayColIndex < _dataGrid.Rows[i].Cells.Count)
                                            _dataGrid.Rows[i].Cells[displayColIndex].Value = newValue;
                                        break;
                                    }
                                }
                            }
                        }
                    });
                    dialog.Show(this);
                }
            }
            else
            {
                var ext = Path.GetExtension(_filePath).ToLowerInvariant();
                if (ext == ".xlsx")
                {
                    if (_isEditLocked)
                    {
                        var dialog = new CellEditDialog(currentValue, originalRowIndex, originalColumnIndex, _filePath, colName,
                            onSave: null, readOnly: true, readOnlyMessage: "当前模式不允许编辑（文件已锁定）");
                        dialog.Show(this);
                    }
                    else
                    {
                        var dialog = new CellEditDialog(currentValue, originalRowIndex, originalColumnIndex, _filePath, colName,
                            onSave: newValue =>
                            {
                                if (originalRowIndex >= 0 && originalRowIndex < _dataTable.Rows.Count &&
                                    originalColumnIndex >= 0 && originalColumnIndex < _dataTable.Columns.Count)
                                {
                                    _dataTable.Rows[originalRowIndex][originalColumnIndex] = newValue;
                                    SaveExcelFile();

                                    int displayColIndex = _visibleColumnIndices.Count > 0
                                        ? _visibleColumnIndices.IndexOf(originalColumnIndex)
                                        : originalColumnIndex;
                                    if (displayColIndex >= 0)
                                    {
                                        for (int i = 0; i < _dataGrid.Rows.Count; i++)
                                        {
                                            if (_dataGrid.Rows[i].Tag is TableRow tagRow && tagRow.OriginalRowIndex == originalRowIndex)
                                            {
                                                if (displayColIndex < _dataGrid.Rows[i].Cells.Count)
                                                    _dataGrid.Rows[i].Cells[displayColIndex].Value = newValue;
                                                break;
                                            }
                                        }
                                    }
                                }
                            });
                        dialog.Show(this);
                    }
                }
                else
                {
                    var dialog = new CellEditDialog(currentValue, originalRowIndex, originalColumnIndex, _filePath, colName,
                        onSave: null, readOnly: true, readOnlyMessage: "不支持此文件类型的修改");
                    dialog.Show(this);
                }
            }
        }

        private void OnCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var dgv = (DataGridView)sender!;
            if (dgv == _dataGridHeader) return;
            // 双击进入行编辑模式
            if (!_isEmbeddedMode && !_isEditLocked)
                EnterRowEditMode(e.RowIndex);
        }

        private void OnFrozenCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var cell = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex];
            _editingFromFrozen = true;
            EditCell(cell, _dataGridHeader);
        }

        private void OnCellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
        {
            // 嵌入式模式（历史对比等）禁止编辑
            if (_isEmbeddedMode) { e.Cancel = true; return; }
            // 编辑锁定模式禁止编辑（Excel默认锁定）
            if (_isEditLocked) { e.Cancel = true; return; }

            // ==================== 行编辑模式 ====================
            // 如果不是通过 EnterRowEditMode 触发的编辑（即用户直接点击单元格尝试编辑），
            // 且不在行编辑模式下，则自动进入行编辑模式
            if (!_enteringEditMode && _editingRowIndex < 0)
            {
                e.Cancel = true;
                EnterRowEditMode(e.RowIndex);
                return;
            }
            // 行编辑模式下，只允许编辑当前编辑行的单元格
            if (_editingRowIndex >= 0 && e.RowIndex != _editingRowIndex)
            {
                e.Cancel = true;
                return;
            }
        }

        private void OnCellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            // ==================== 行编辑模式 ====================
            if (_editingRowIndex >= 0)
            {
                // 行编辑模式下不逐单元格保存，等整行退出后统一保存
                // 使用 BeginInvoke 延迟跳转，避免在 EndEdit 调用栈内设置 CurrentCell 导致重入
                if (e.ColumnIndex < _dataGrid.Columns.Count - 1 && !_enteringEditMode)
                {
                    int nextRow = e.RowIndex;
                    int nextCol = e.ColumnIndex + 1;
                    BeginInvoke(new Action(() =>
                    {
                        // 延迟执行时再次校验：编辑行未变且行仍有效
                        if (_editingRowIndex == nextRow && nextRow < _dataGrid.Rows.Count)
                        {
                            // 检查 CurrentCell 是否已经被用户（Tab/鼠标）移到了目标位置
                            if (_dataGrid.CurrentCell != null
                                && _dataGrid.CurrentCell.RowIndex == nextRow
                                && _dataGrid.CurrentCell.ColumnIndex == nextCol)
                            {
                                // DataGridView 已经移到了目标列，只需进入编辑模式
                                _dataGrid.BeginEdit(true);
                            }
                            else
                            {
                                // CurrentCell 不在预期位置（如按 Enter 结束编辑），需要手动跳转
                                _enteringEditMode = true;
                                try
                                {
                                    _dataGrid.CurrentCell = _dataGrid.Rows[nextRow].Cells[nextCol];
                                    _dataGrid.BeginEdit(true);
                                }
                                finally { _enteringEditMode = false; }
                            }
                        }
                    }));
                }
                return;
            }

            // 非行编辑模式下的原有逻辑（正常不会走到这里，因为 ReadOnly=true）
            var cell = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex];
            var newValue = cell.Value?.ToString() ?? "";
            var (originalRowIndex, originalColumnIndex) = GetOriginalIndices(e.RowIndex, e.ColumnIndex);

            if (originalRowIndex >= 0 && originalRowIndex < _dataTable.Rows.Count &&
                originalColumnIndex >= 0 && originalColumnIndex < _dataTable.Columns.Count)
            {
                _dataTable.Rows[originalRowIndex][originalColumnIndex] = newValue;
                SaveCurrentFile();
            }
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
                // ★ 从 DataGridView 行的 Tag 直接获取原始行号
                int originalRowIndex = displayRowIndex;
                if (displayRowIndex >= 0 && displayRowIndex < _dataGrid.Rows.Count)
                {
                    var tagRow = _dataGrid.Rows[displayRowIndex].Tag as TableRow;
                    if (tagRow != null)
                        originalRowIndex = tagRow.OriginalRowIndex;
                }
                int originalColumnIndex = displayColumnIndex;
                if (_visibleColumnIndices.Count > 0 && displayColumnIndex < _dataGrid.Columns.Count)
                {
                    if (displayColumnIndex >= 0 && displayColumnIndex < _visibleColumnIndices.Count)
                        originalColumnIndex = _visibleColumnIndices[displayColumnIndex];
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
                using (var fs = new FileStream(_filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                { fs.Write(bytes, 0, bytes.Length); fs.Flush(); }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"保存文件失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // ★ 保存后重新检测唯一列重复
            CheckUniqueColumnDuplicates();
        }

        private void SaveExcelFile()
        {
            var backupPath = _filePath + "." + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".bak";
            try
            {
                // 1. 备份原始文件并删除原文件（避免 MiniExcel 报"已存在"）
                if (File.Exists(_filePath))
                {
                    File.Copy(_filePath, backupPath, overwrite: true);
                    File.Delete(_filePath);
                }

                // 2. 写入新内容
                MiniExcel.SaveAs(_filePath, _dataTable, printHeader: false);

                // 3. 写入成功，删除备份
                if (File.Exists(backupPath))
                    File.Delete(backupPath);
            }
            catch (Exception ex)
            {
                // 4. 写入失败，尝试还原原始文件
                if (File.Exists(backupPath))
                {
                    try
                    {
                        if (File.Exists(_filePath))
                            File.Delete(_filePath);
                        File.Move(backupPath, _filePath);
                    }
                    catch
                    {
                        // 还原失败也保留备份文件，用户可手动恢复
                    }
                }
                MessageBox.Show(this, $"保存Excel文件失败:{Environment.NewLine}{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // ★ 保存后重新检测唯一列重复
            CheckUniqueColumnDuplicates();
        }

        private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (sender == _dataGridHeader) return;
            if (e.RowIndex >= 0 && e.RowIndex < _dataGrid.Rows.Count)
                _dataGrid.Rows[e.RowIndex].Height = 24;
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // ==================== 行编辑模式：编辑行浅绿色背景 ====================
                if (_editingRowIndex >= 0 && e.RowIndex == _editingRowIndex)
                {
                    e.CellStyle.BackColor = Color.FromArgb(230, 255, 230);  // 浅绿色
                    e.CellStyle.SelectionBackColor = Color.FromArgb(200, 240, 200);
                }
                else
                {
                    // ★ 传入 _dirCache
                    if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                        _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
                    ColumnLinkChecker.CheckCellColor(e, _visibleColumnIndices, _dataGrid, _filePath, FrozenDir, _dirCache, _columnLinks);
                    if (_highlightsReady && _highlightRows.Contains(e.RowIndex))
                    {
                        e.CellStyle.BackColor = Color.FromArgb(255, 255, 200);
                        e.CellStyle.SelectionBackColor = Color.FromArgb(255, 240, 150);
                    }
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
            // ★ 不再需要 __RowNum__ 偏移
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
            if (_isEmbeddedMode || _isEditLocked) editItem.Enabled = false;
            else
            {
                editItem.Click += (s, args) =>
                {
                    _dataGrid.CurrentCell = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    _dataGrid.BeginEdit(true);
                };
            }
            ctx.Items.Add(editItem);

            // ==================== 编辑行菜单 ====================
            var editRowItem = new ToolStripMenuItem(_editingRowIndex >= 0 ? "退出行编辑" : "编辑行");
            if (_isEmbeddedMode) { editRowItem.Enabled = false; }
            else
            {
                int targetRow = e.RowIndex;
                editRowItem.Click += (s, args) =>
                {
                    // 先选中右键点击的行
                    _dataGrid.CurrentCell = _dataGrid.Rows[targetRow].Cells[e.ColumnIndex];
                    if (_editingRowIndex >= 0)
                    {
                        // 如果当前正在编辑，则提交退出
                        ExitRowEditMode(commit: true);
                    }
                    else
                    {
                        EnterRowEditMode(targetRow);
                    }
                };
            }
            ctx.Items.Add(editRowItem);

            var editDialogItem = new ToolStripMenuItem("弹窗编辑");
            if (_isEmbeddedMode) editDialogItem.Enabled = false;
            else
            {
                editDialogItem.Click += (s, args) =>
                {
                    var cell = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    _editingFromFrozen = false;
                    EditCell(cell, _dataGrid);
                };
            }
            ctx.Items.Add(editDialogItem);

            var linkItem = new ToolStripMenuItem("打开关联表");
            if (link == null) { linkItem.Enabled = false; }
            else
            {
                // ★ 传入 _dirCache
                if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                    _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
                bool isValid = !string.IsNullOrEmpty(cellValue) && ColumnLinkChecker.IsLinkValueValid(_filePath, FrozenDir, origCol, cellValue, _dirCache, _columnLinks);
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
                    var linkOpts = new TableBrowserOptions(targetFullPath)
                    {
                        RootDir = FrozenDir,
                        FrozenRows = _frozenRows,
                        HeaderRowIndex = _headerRowIndex,
                        MappingRowIndex = _mappingRowIndex,
                    };
                    var linkOpen = new LinkOpenOptions
                    {
                        HighlightColumnIndex = link.TargetColumnIndex,
                        HighlightValue = cellValue,
                        HighlightColumnName = link.TargetColumnName,
                        LinkConfig = link,
                        DisplayName = displayName,
                        IsLinkedFromCell = true,
                    };
                    var standaloneForm = new CsvReaderStandaloneForm(linkOpts, linkOpen);
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
                // ★ 不再需要 __RowNum__ 偏移
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

            // ==================== 行操作菜单 ====================
            ctx.Items.Add(new ToolStripSeparator());

            // 行编辑模式下禁用行操作
            bool rowOpsDisabled = _editingRowIndex >= 0 || _isEditLocked;

            var insertAboveItem = new ToolStripMenuItem("在上方插入行");
            if (rowOpsDisabled) insertAboveItem.Enabled = false;
            else
            {
                insertAboveItem.Click += (s, args) =>
                {
                    int origRowIndex = GetOriginalRowIndexFromDisplay(e.RowIndex);
                    InsertRowAt(origRowIndex - 1);
                };
            }
            ctx.Items.Add(insertAboveItem);

            var insertBelowItem = new ToolStripMenuItem("在下方插入行");
            if (rowOpsDisabled) insertBelowItem.Enabled = false;
            else
            {
                insertBelowItem.Click += (s, args) =>
                {
                    int origRowIndex = GetOriginalRowIndexFromDisplay(e.RowIndex);
                    InsertRowAt(origRowIndex);
                };
            }
            ctx.Items.Add(insertBelowItem);

            var duplicateItem = new ToolStripMenuItem("复制当前行");
            if (rowOpsDisabled) duplicateItem.Enabled = false;
            else
            {
                duplicateItem.Click += (s, args) =>
                {
                    int origRowIndex = GetOriginalRowIndexFromDisplay(e.RowIndex);
                    DuplicateRowAt(origRowIndex);
                };
            }
            ctx.Items.Add(duplicateItem);

            var deleteItem = new ToolStripMenuItem("删除当前行");
            if (rowOpsDisabled) deleteItem.Enabled = false;
            else
            {
                deleteItem.Click += (s, args) =>
                {
                    int origRowIndex = GetOriginalRowIndexFromDisplay(e.RowIndex);
                    DeleteRowAt(origRowIndex);
                };
            }
            ctx.Items.Add(deleteItem);

            ctx.Show(_dataGrid, _dataGrid.PointToClient(Cursor.Position));
        }

        private string GetCurrentRelativePath()
        {
            if (!string.IsNullOrEmpty(FrozenDir) && _filePath.StartsWith(FrozenDir, StringComparison.OrdinalIgnoreCase))
                return _filePath.Substring(FrozenDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(_filePath);
        }

        /// <summary>后台异步构建关联 lookup，不阻塞 UI（fire-and-forget，内部防重入）</summary>
        private void FireAndForgetBuildLinkLookups(DirectoryCache dirCache, List<ColumnLinkConfig> columnLinks, string filePath)
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    await dirCache.BuildLinkLookupsAsync(columnLinks, filePath);
                    // 构建完成后，通知 UI 刷新（让 tooltip 能获取到最新数据）
                    if (this.IsHandleCreated)
                        this.BeginInvoke(() => _dataGrid?.Invalidate());
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"BuildLinkLookupsAsync failed: {ex.Message}");
                }
            });
        }

        /// <summary>获取状态栏显示文件名（优先使用备注名）</summary>
        private string GetDisplayFileName()
        {
            var relPath = GetCurrentRelativePath();
            var remark = _fileRemarks.GetValueOrDefault(relPath ?? "");
            if (!string.IsNullOrEmpty(remark))
                return $"{remark} ({Path.GetFileName(_filePath)})";
            return Path.GetFileName(_filePath);
        }

        /// <summary>点击锁图标：切换编辑锁定/解锁状态</summary>
        private void OnLockLabelClick(object? sender, EventArgs e)
        {
            if (!_isEditLocked)
            {
                // 解锁状态 → 重新锁定
                // 如果正在行编辑，先退出编辑模式
                if (_editingRowIndex >= 0)
                    ExitRowEditMode(commit: true);
                _isEditLocked = true;
                UpdateLockDisplay();
                return;
            }

            // 锁定状态 → 弹出确认框解锁
            var result = MessageBox.Show(this,
                "⚠️ 警告：Excel文件如果包含公式、宏定义等，编辑后保存可能导致这些内容丢失。\n确认要解锁编辑吗？",
                "解锁编辑确认",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning);

            if (result == DialogResult.OK)
            {
                _isEditLocked = false;
                UpdateLockDisplay();
            }
        }

        /// <summary>更新状态栏锁图标和文件名显示</summary>
        private void UpdateLockDisplay()
        {
            if (_isEditLocked)
            {
                _lblLock.Text = "🔒 锁定";
                _lblLock.ForeColor = _colorLightTextSecondary;
            }
            else
            {
                _lblLock.Text = "✏️ 可编辑";
                _lblLock.ForeColor = Color.FromArgb(22, 163, 74); // 绿色，表示可编辑
            }
        }

        // ==================== 行号绘制 ====================
        private void OnRowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            // ★ 从 DataGridView 行的 Tag 直接获取原始行号
            int rowIndex;
            if (e.RowIndex >= 0 && e.RowIndex < _dataGrid.Rows.Count)
            {
                var tagRow = _dataGrid.Rows[e.RowIndex].Tag as TableRow;
                if (tagRow != null)
                    rowIndex = tagRow.OriginalRowIndex + 1;
                else
                    rowIndex = e.RowIndex + 1 + _frozenRows;
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

        private bool IsBalanced(string s)
        {
            int d = 0;
            foreach (char c in s) { if (c == '(') d++; else if (c == ')') d--; if (d < 0) return false; }
            return d == 0;
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
                    var form = new VersionCompareForm(filePath, vcsType, frozen, headerRow,
                        rootDir: FrozenDir,
                        columnVisibility: _columnCheckStates,
                        uniqueColumnIndex: _uniqueColumnIndex,
                        mappingRowIndex: _mappingRowIndex);
                    form.Show(this);
                };

                _flowTop?.Controls.Add(_btnHistory);
            }
            catch { }
        }

        /// <summary>获取内部 DataTable（供版本对比等外部功能使用）</summary>
        public DataTable GetDataTable() => _dataTable;

        /// <summary>获取内部 FileTableData（用于行号追踪、差异标记）</summary>
        public FileTableData GetFileData() => _fileData;

        /// <summary>直接设置 FileTableData 并刷新 Tag（供版本对比等场景使用，绕过 BuildFileData 继承逻辑）</summary>
        public void SetFileData(FileTableData fileData)
        {
            _fileData = fileData;
            SetRowTags();
        }

        /// <summary>获取内部 DataGridView（供版本对比等外部功能使用）</summary>
        public DataGridView GetDataGridView() => _dataGrid;

        /// <summary>获取固定行数（供版本对比计算内容表格行号映射）</summary>
        public int FrozenRows => _frozenRows;

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

        /// <summary>从 DataTable 构建 FileTableData（用于行号追踪）</summary>
        private static FileTableData BuildFileData(DataTable dt, FileTableData? inheritFrom = null)
        {
            var fd = new FileTableData();
            foreach (DataColumn col in dt.Columns)
                fd.ColumnNames.Add(col.ColumnName);

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                int origRowIdx = r;
                bool isDiff = false;
                HashSet<int>? diffCols = null;

                if (inheritFrom != null)
                {
                    // ★ 如果新旧表行数一致 → 按索引继承（排序场景）
                    if (dt.Rows.Count == inheritFrom.Rows.Count && r < inheritFrom.Rows.Count)
                    {
                        origRowIdx = inheritFrom.Rows[r].OriginalRowIndex;
                        isDiff = inheritFrom.Rows[r].IsDiffRow;
                        diffCols = inheritFrom.Rows[r].DiffColumns;
                    }
                    // ★ 如果行数不一致（筛选子集场景）→ 通过 OriginalRowIndex 反查
                    else
                    {
                        origRowIdx = r;
                        // 在 inheritFrom 中查找 OriginalRowIndex == r 的行
                        var matched = inheritFrom.Rows.FirstOrDefault(x => x.OriginalRowIndex == r);
                        if (matched != null)
                        {
                            isDiff = matched.IsDiffRow;
                            diffCols = matched.DiffColumns;
                        }
                    }
                }

                var row = new TableRow
                {
                    OriginalRowIndex = origRowIdx,
                    IsDiffRow = isDiff,
                    DiffColumns = diffCols
                };
                for (int c = 0; c < dt.Columns.Count; c++)
                    row.Cells.Add(dt.Rows[r][c]?.ToString() ?? "");

                fd.Rows.Add(row);
            }
            return fd;
        }

        public void SetDataTable(DataTable dataTable, string infoText = "")
        {
            if (dataTable == null) return;

            _isEmbeddedMode = true; // 标记为外部嵌入模式（版本对比等），禁止编辑

            // ★ 彻底清理旧绑定，避免多次切换后事件/引用累积导致卡顿
            _dataGrid.DataSource = null;
            _dataGridHeader.DataSource = null;
            // 清理可能由外部附加的 CellFormatting 事件（如版本对比的差异高亮）
            _dataGrid.CellFormatting -= OnExternalCellFormatting;
            // 清理高亮状态
            _highlightRows.Clear();
            _highlightsReady = false;

            _dataTable = dataTable;
            // ★ 同步构建 FileTableData（继承旧差异信息）
            _fileData = BuildFileData(dataTable, _fileData);
            _headerTable = dataTable.Clone();

            if (_externalColumnStates != null && _externalColumnStates.Length == _dataTable.Columns.Count)
            {
                _columnCheckStates = (bool[])_externalColumnStates.Clone();
                _externalColumnStates = null; // 只使用一次，避免影响后续调用
            }
            else
            {
                _columnCheckStates = new bool[_dataTable.Columns.Count];
                for (int i = 0; i < _columnCheckStates.Length; i++)
                    _columnCheckStates[i] = true;
            }

            RebuildHighlightRows();
            _lblCellPos.Text = string.IsNullOrEmpty(infoText)
                ? $"行数: {dataTable.Rows.Count} | 列数: {dataTable.Columns.Count}"
                : infoText;
            _lblInfo.Text = "";

            // ★ 如果有重复警告，追加警告文本并保持红色
            if (_duplicateWarningText != null)
            {
                _lblInfo.ForeColor = Color.FromArgb(200, 30, 30);
                _lblInfo.Text = _duplicateWarningText;
            }

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
                    // 如果设置了表头行，用表头行的值作为列标题
                    if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count)
                    {
                        var headerVal = _dataTable.Rows[_headerRowIndex][i]?.ToString() ?? GetExcelColumnName(i);
                        _dataGrid.Columns[i].HeaderText = headerVal;
                        SetColumnHeaderToolTip(_dataGrid.Columns[i], i, headerVal);
                    }
                    else
                    {
                        _dataGrid.Columns[i].HeaderText = GetExcelColumnName(i);
                    }
                }
                // ★ 同步设置 Tag（DataSource 赋值后行已创建）
                SetRowTags();
            }

            // ★ 数据表变更后刷新唯一列下拉选项
            RefreshUniqueColumnCombo();

            // ★ 初始化自定义 tooltip（如果尚未初始化，且非轻量模式）
            if (!_lightweight && _linkToolTip == null)
            {
                _linkToolTip = new DataGridLinkToolTip(_dataGrid);
                _linkToolTip.RegisterGrid(_dataGridHeader);
                _linkToolTip.GetToolTipContent = GetLinkToolTipContent;
            }

            // ★ 重新加载关联配置（确保 _columnLinks 和缓存正确；
            //   版本对比等 SetDataTable 场景下 _filePath 为空字符串，需要依赖 rootDir 加载配置）
            if (!string.IsNullOrEmpty(_rootDir))
            {
                _columnLinks = CsvSettings.GetColumnLinksByDirectory(_rootDir);
                if (_dirCache == null)
                    _dirCache = FileDataCache.GetDirectoryCache(_rootDir);
                if (_dirCache != null && _columnLinks.Count > 0)
                    FireAndForgetBuildLinkLookups(_dirCache, _columnLinks, _filePath);
            }
        }

        /// <summary>供外部（版本对比等）移除的 CellFormatting 事件占位</summary>
        internal event DataGridViewCellFormattingEventHandler? OnExternalCellFormatting;

        /// <summary>设置外部传入的列可见性状态（版本对比等场景使用）。下次 SetDataTable 调用时会优先使用此状态。</summary>
        public void SetColumnVisibility(bool[]? states)
        {
            _externalColumnStates = states;
        }

        /// <summary>获取当前列显隐状态的副本（用于切换数据时保持用户选择的列显示/隐藏）</summary>
        public bool[] GetColumnCheckStates()
        {
            return (bool[])_columnCheckStates.Clone();
        }

        // ==================== 行编辑模式 ====================

        /// <summary>进入行编辑模式</summary>
        /// <param name="displayRowIndex">DataGridView 显示行索引</param>
        private void EnterRowEditMode(int displayRowIndex)
        {
            if (_isEmbeddedMode) return;
            if (_isEditLocked) return;
            if (displayRowIndex < 0 || displayRowIndex >= _dataGrid.Rows.Count) return;

            // 如果正在编辑同一行，不重复进入
            if (_editingRowIndex == displayRowIndex) return;

            // 如果正在编辑另一行，先退出
            if (_editingRowIndex >= 0)
                ExitRowEditMode(commit: true);

            _enteringEditMode = true;
            try
            {
                // 快照当前行所有单元格值
                var row = _dataGrid.Rows[displayRowIndex];
                _editingRowOriginalValues = new string[row.Cells.Count];
                for (int i = 0; i < row.Cells.Count; i++)
                    _editingRowOriginalValues[i] = row.Cells[i].Value?.ToString() ?? "";

                _editingRowIndex = displayRowIndex;

                // 获取原始行索引（如果 _editingRowOrigIndex 尚未被外部设置）
                if (_editingRowOrigIndex < 0)
                {
                    if (row.Tag is TableRow tr)
                        _editingRowOrigIndex = tr.OriginalRowIndex;
                    else
                        _editingRowOrigIndex = displayRowIndex + _frozenRows;
                }

                // 临时解除只读
                _dataGrid.ReadOnly = false;

                // 刷新行高亮
                _dataGrid.InvalidateRow(displayRowIndex);

                // 自动定位到第一个可编辑列并进入编辑
                if (_dataGrid.Columns.Count > 0)
                {
                    _dataGrid.CurrentCell = row.Cells[0];
                    _dataGrid.BeginEdit(true);
                }
            }
            finally { _enteringEditMode = false; }
        }

        /// <summary>退出行编辑模式（提交修改）</summary>
        /// <param name="commit">true=提交修改保存文件，false=只退出不保存</param>
        private void ExitRowEditMode(bool commit)
        {
            if (_editingRowIndex < 0) return;

            _enteringEditMode = true;
            try
            {
                int editingRow = _editingRowIndex;

                // 先清除状态，避免事件重入
                _editingRowIndex = -1;
                var snapShot = _editingRowOriginalValues;
                _editingRowOriginalValues = null;
                int origIndex = _editingRowOrigIndex;
                _editingRowOrigIndex = -1;
                bool wasNewRow = _isNewRowForEdit;
                _isNewRowForEdit = false;

                // 恢复只读
                _dataGrid.ReadOnly = true;

                if (commit && editingRow >= 0 && editingRow < _dataGrid.Rows.Count && origIndex >= 0)
                {
                    bool hasChanges = false;
                    var row = _dataGrid.Rows[editingRow];

                    // 将 DataGridView 中的修改写回 _dataTable
                    for (int col = 0; col < row.Cells.Count; col++)
                    {
                        var newValue = row.Cells[col].Value?.ToString() ?? "";
                        var oldValue = snapShot != null && col < snapShot.Length ? snapShot[col] : "";

                        if (newValue != oldValue)
                        {
                            hasChanges = true;
                            // 计算原始列索引
                            int origCol = col;
                            if (_visibleColumnIndices.Count > 0 && col < _visibleColumnIndices.Count)
                                origCol = _visibleColumnIndices[col];

                            if (origIndex < _dataTable.Rows.Count && origCol < _dataTable.Columns.Count)
                                _dataTable.Rows[origIndex][origCol] = newValue;
                        }
                    }

                    if (hasChanges || wasNewRow)
                    {
                        // 重建 _fileData
                        _fileData = BuildFileData(_dataTable, inheritFrom: null);

                        // 重建关联缓存（不再调用 Reset，dirCache 内部管理）
                        if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                            _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
                        if (_dirCache != null && _columnLinks.Count > 0)
                            FireAndForgetBuildLinkLookups(_dirCache, _columnLinks, _filePath);

                        // 一次性保存文件
                        SaveCurrentFile();

                        // 刷新显示（重建 Grid 以反映 _dataTable 的变更）
                        RefreshGridLayout();
                    }
                    else
                    {
                        // 没有实际修改，只需刷新编辑行的高亮
                        _dataGrid.InvalidateRow(editingRow);
                    }
                }
                else
                {
                    // 不提交，刷新编辑行的高亮
                    if (editingRow >= 0 && editingRow < _dataGrid.Rows.Count)
                        _dataGrid.InvalidateRow(editingRow);
                }
            }
            finally { _enteringEditMode = false; }
        }

        /// <summary>取消行编辑模式（Esc，恢复原始值）</summary>
        private void CancelRowEditMode()
        {
            if (_editingRowIndex < 0) return;

            _enteringEditMode = true;
            try
            {
                int editingRow = _editingRowIndex;
                _editingRowIndex = -1;
                var snapShot = _editingRowOriginalValues;
                _editingRowOriginalValues = null;
                bool wasNewRow = _isNewRowForEdit;
                int origIndex = _editingRowOrigIndex;
                _editingRowOrigIndex = -1;
                _isNewRowForEdit = false;

                // 恢复只读
                _dataGrid.ReadOnly = true;

                // 先结束当前单元格编辑
                if (_dataGrid.IsCurrentCellInEditMode)
                {
                    // 不提交，直接取消编辑
                    _dataGrid.CancelEdit();
                }

                if (wasNewRow && origIndex >= 0 && origIndex < _dataTable.Rows.Count)
                {
                    // 新增行：删除该行
                    _dataTable.Rows.RemoveAt(origIndex);

                    // 重建 _fileData
                    _fileData = BuildFileData(_dataTable, inheritFrom: null);

                    // 重建关联缓存
                    if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                        _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
                    if (_dirCache != null && _columnLinks.Count > 0)
                        FireAndForgetBuildLinkLookups(_dirCache, _columnLinks, _filePath);

                    RefreshGridLayout();
                }
                else if (snapShot != null && editingRow >= 0 && editingRow < _dataGrid.Rows.Count)
                {
                    // 恢复原始值到 DataGridView
                    var row = _dataGrid.Rows[editingRow];
                    for (int i = 0; i < Math.Min(snapShot.Length, row.Cells.Count); i++)
                        row.Cells[i].Value = snapShot[i];

                    // 刷新编辑行的高亮
                    _dataGrid.InvalidateRow(editingRow);
                }
            }
            finally { _enteringEditMode = false; }
        }

        // ==================== 行操作 ====================

        /// <summary>从 DataGridView 显示行索引获取原始 DataTable 行索引</summary>
        private int GetOriginalRowIndexFromDisplay(int displayRowIndex)
        {
            if (displayRowIndex >= 0 && displayRowIndex < _dataGrid.Rows.Count)
            {
                if (_dataGrid.Rows[displayRowIndex].Tag is TableRow tr)
                    return tr.OriginalRowIndex;
            }
            // 回退：无冻结行时 displayRowIndex 即原始行索引
            return displayRowIndex;
        }

        /// <summary>在指定位置之后插入空行</summary>
        /// <param name="afterOriginalRowIndex">在此原始行索引之后插入（-1=最前面插入，>=Rows.Count=末尾追加）</param>
        private void InsertRowAt(int afterOriginalRowIndex)
        {
            if (_isEmbeddedMode)
            {
                MessageBox.Show(this, "当前状态不支持插入行", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 如果有排序，先取消排序
            if (_sortColumnIndex >= 0)
            {
                var result = MessageBox.Show(this,
                    "当前处于排序状态，插入行前需要先取消排序。是否继续？",
                    "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result != DialogResult.Yes) return;
                _sortColumnIndex = -1;
                _sortOrder = SortOrder.None;
            }

            // 保存滚动位置
            int scrollRow = _dataGrid.FirstDisplayedScrollingRowIndex;
            int selectedRow = _dataGrid.CurrentRow?.Index ?? -1;
            int selectedCol = _dataGrid.CurrentCell?.ColumnIndex ?? -1;

            // 创建新行，所有列初始化为空字符串
            var newRow = _dataTable.NewRow();
            for (int i = 0; i < _dataTable.Columns.Count; i++)
                newRow[i] = "";

            // 插入到 _dataTable
            if (afterOriginalRowIndex >= 0 && afterOriginalRowIndex < _dataTable.Rows.Count - 1)
                _dataTable.Rows.InsertAt(newRow, afterOriginalRowIndex + 1);
            else if (afterOriginalRowIndex < 0)
                _dataTable.Rows.InsertAt(newRow, 0);  // 最前面插入
            else
                _dataTable.Rows.Add(newRow);  // 末尾追加

            // 重建 _fileData
            _fileData = BuildFileData(_dataTable, inheritFrom: null);

            // 重建关联缓存
            if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
            if (_dirCache != null && _columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_dirCache, _columnLinks, _filePath);

            // 刷新显示
            RefreshGridLayout();

            // 恢复滚动位置
            if (scrollRow >= 0 && scrollRow < _dataGrid.Rows.Count)
                _dataGrid.FirstDisplayedScrollingRowIndex = scrollRow;
            if (selectedRow >= 0 && selectedRow < _dataGrid.Rows.Count && selectedCol >= 0)
                _dataGrid.CurrentCell = _dataGrid.Rows[selectedRow].Cells[selectedCol];

            // 不再即时保存，等行编辑模式退出后统一保存

            // ==================== 自动进入行编辑模式 ====================
            // 找到新插入行在显示表格中的位置
            int newOrigIndex;
            if (afterOriginalRowIndex >= 0 && afterOriginalRowIndex < _dataTable.Rows.Count - 2)
                newOrigIndex = afterOriginalRowIndex + 1;
            else if (afterOriginalRowIndex < 0)
                newOrigIndex = 0;
            else
                newOrigIndex = _dataTable.Rows.Count - 1;

            // 在 _dataGrid 中定位到新行
            int displayRow = -1;
            for (int i = 0; i < _dataGrid.Rows.Count; i++)
            {
                if (_dataGrid.Rows[i].Tag is TableRow tr && tr.OriginalRowIndex == newOrigIndex)
                { displayRow = i; break; }
            }
            if (displayRow < 0) displayRow = 0;

            _isNewRowForEdit = true;
            _editingRowOrigIndex = newOrigIndex;
            EnterRowEditMode(displayRow);
        }

        /// <summary>复制指定行并插入到其下方</summary>
        private void DuplicateRowAt(int originalRowIndex)
        {
            if (_isEmbeddedMode)
            {
                MessageBox.Show(this, "当前状态不支持复制行", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (originalRowIndex < 0 || originalRowIndex >= _dataTable.Rows.Count)
                return;

            // 如果有排序，先取消排序
            if (_sortColumnIndex >= 0)
            {
                var result = MessageBox.Show(this,
                    "当前处于排序状态，复制行前需要先取消排序。是否继续？",
                    "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result != DialogResult.Yes) return;
                _sortColumnIndex = -1;
                _sortOrder = SortOrder.None;
            }

            // 保存滚动位置
            int scrollRow = _dataGrid.FirstDisplayedScrollingRowIndex;
            int selectedRow = _dataGrid.CurrentRow?.Index ?? -1;
            int selectedCol = _dataGrid.CurrentCell?.ColumnIndex ?? -1;

            // 复制源行数据
            var newRow = _dataTable.NewRow();
            var sourceRow = _dataTable.Rows[originalRowIndex];
            for (int i = 0; i < _dataTable.Columns.Count; i++)
                newRow[i] = sourceRow[i];

            // 插入到源行下方
            _dataTable.Rows.InsertAt(newRow, originalRowIndex + 1);

            // 重建 _fileData
            _fileData = BuildFileData(_dataTable, inheritFrom: null);

            // 重建关联缓存
            if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
            if (_dirCache != null && _columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_dirCache, _columnLinks, _filePath);

            // 刷新显示
            RefreshGridLayout();

            // 恢复滚动位置
            if (scrollRow >= 0 && scrollRow < _dataGrid.Rows.Count)
                _dataGrid.FirstDisplayedScrollingRowIndex = scrollRow;
            if (selectedRow >= 0 && selectedRow < _dataGrid.Rows.Count && selectedCol >= 0)
                _dataGrid.CurrentCell = _dataGrid.Rows[selectedRow].Cells[selectedCol];

            // 不再即时保存，等行编辑模式退出后统一保存

            // ==================== 自动进入行编辑模式 ====================
            // 找到新复制行在显示表格中的位置
            int dupOrigIndex = originalRowIndex + 1;
            int dupDisplayRow = -1;
            for (int i = 0; i < _dataGrid.Rows.Count; i++)
            {
                if (_dataGrid.Rows[i].Tag is TableRow tr && tr.OriginalRowIndex == dupOrigIndex)
                { dupDisplayRow = i; break; }
            }
            if (dupDisplayRow < 0) dupDisplayRow = 0;

            _isNewRowForEdit = false;  // 复制行不是新增行，Esc不删除
            _editingRowOrigIndex = dupOrigIndex;
            EnterRowEditMode(dupDisplayRow);
        }

        /// <summary>删除指定行</summary>
        private void DeleteRowAt(int originalRowIndex)
        {
            if (_isEmbeddedMode)
            {
                MessageBox.Show(this, "当前状态不支持删除行", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (originalRowIndex < 0 || originalRowIndex >= _dataTable.Rows.Count)
                return;

            // ==================== 行编辑模式：不允许删除正在编辑的行 ====================
            if (_editingRowIndex >= 0 && _editingRowOrigIndex == originalRowIndex)
            {
                MessageBox.Show(this, "请先完成或取消当前行的编辑再删除。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 冻结行不允许删除
            if (originalRowIndex < _frozenRows)
            {
                MessageBox.Show(this, "冻结行不允许删除。请先调整冻结行设置。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = MessageBox.Show(this,
                $"确认删除第 {originalRowIndex + 1} 行？此操作不可撤销。",
                "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result != DialogResult.Yes) return;

            // 取消排序
            if (_sortColumnIndex >= 0)
            {
                _sortColumnIndex = -1;
                _sortOrder = SortOrder.None;
            }

            // 保存滚动位置
            int scrollRow = _dataGrid.FirstDisplayedScrollingRowIndex;
            int selectedRow = _dataGrid.CurrentRow?.Index ?? -1;
            int selectedCol = _dataGrid.CurrentCell?.ColumnIndex ?? -1;

            _dataTable.Rows.RemoveAt(originalRowIndex);

            // 重建 _fileData
            _fileData = BuildFileData(_dataTable, inheritFrom: null);

            // 重建关联缓存
            if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);
            if (_dirCache != null && _columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_dirCache, _columnLinks, _filePath);

            // 刷新显示
            RefreshGridLayout();

            // 恢复滚动位置
            if (scrollRow >= 0 && scrollRow < _dataGrid.Rows.Count)
                _dataGrid.FirstDisplayedScrollingRowIndex = scrollRow;
            if (selectedRow >= 0 && selectedRow < _dataGrid.Rows.Count && selectedCol >= 0)
                _dataGrid.CurrentCell = _dataGrid.Rows[selectedRow].Cells[selectedCol];

            // 即时保存
            SaveCurrentFile();
        }

        /// <summary>统一保存（根据文件类型调用 CSV 或 Excel 保存）</summary>
        private void SaveCurrentFile()
        {
            if (_fileType == TableFileType.Csv)
                SaveCsvFile();
            else
                SaveExcelFile();
        }

        // ==================== 防抖保存 ====================

        /// <summary>标记脏数据并启动防抖定时器</summary>
        private void MarkDirtyAndStartSaveTimer()
        {
            _hasUnsavedChanges = true;
            _retrySaveTimer?.Stop();
            _retrySaveTimer?.Start();
            ShowDirtyIndicator();
        }

        /// <summary>尝试写磁盘</summary>
        private void TryFlushToDisk()
        {
            if (!_hasUnsavedChanges) return;

            if (_dirCache == null && !string.IsNullOrEmpty(FrozenDir))
                _dirCache = FileDataCache.GetDirectoryCache(FrozenDir);

            var relPath = GetCurrentRelativePath();
            bool saved;

            if (_fileType == TableFileType.Csv)
                saved = _dirCache?.TrySaveCsvToDisk(relPath, _dataTable, _delimiter) ?? false;
            else
                saved = _dirCache?.TrySaveExcelToDisk(relPath, _dataTable) ?? false;

            if (saved)
            {
                _hasUnsavedChanges = false;
                // 更新缓存主副本
                _dirCache?.UpdateCacheFromMemory(relPath, _dataTable,
                    BuildFileData(_dataTable, inheritFrom: null));
                HideDirtyIndicator();
            }
            else
            {
                // 仍被占用 → 重新启动定时器，继续等待
                _retrySaveTimer?.Start();
            }
        }

        /// <summary>关闭前强制保存</summary>
        private bool ForceFlushBeforeClose()
        {
            _retrySaveTimer?.Stop();
            if (!_hasUnsavedChanges) return true;

            TryFlushToDisk();
            if (!_hasUnsavedChanges) return true;

            // 仍失败 → 弹窗让用户决定
            var result = MessageBox.Show(
                "文件被占用无法保存，修改仅保留在内存中。\n关闭后修改将丢失，是否继续？",
                "⚠ 保存失败", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            return result == DialogResult.Yes;
        }

        private void ShowDirtyIndicator()
        {
            if (_lblInfo.IsHandleCreated)
            {
                _lblInfo.Text = "⚠ 文件被占用，修改未保存（等待重试...）";
                _lblInfo.ForeColor = Color.FromArgb(200, 130, 30);
            }
        }

        private void HideDirtyIndicator()
        {
            if (_lblInfo.IsHandleCreated)
            {
                _lblInfo.Text = "";
                _lblInfo.ForeColor = _colorLightTextSecondary;
            }
        }

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
