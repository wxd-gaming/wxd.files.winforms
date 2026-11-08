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
        private Panel? _headerPanel;  // 固定行区域容器引用，避免每次 AttachGridToParent 创建新 Panel 导致空 Panel 堆积遮挡
        private ThemeButton _btnSelectColumns = new();
        private TextBox _txtContentFilter = new();
        private TextBox _txtAdvancedFilter = new();
        private Label _lblLock = new();
        private Label _lblFileName = new();
        private Label _lblInfo = new();
        private Label _lblCellPos = new();
        private Label _lblRowNumber = new();      // 选中标记行的源文件行号区间
        private Font _rowNumberFont = Theme.Mono;  // 行号区等宽字体，多位行号对齐便于扫读
        private static readonly Font AdvHelpGlyphFont = new("Microsoft YaHei UI", 12f, FontStyle.Bold);
        private ThemeCheckBox _chkFilterKeepAll = new();
        private Panel _pnlLoading = new();
        private ThemeProgressBar? _loadingProgressBar;
        private Label? _loadingLabel;
        private bool _isLoading = false;  // 防重入标志（DoEvents 期间避免二次加载）
        private Panel _pnlTop = new();
        private TableLayoutPanel _pnlFilterBar = new();
        private Panel _pnlStatus = new();
        private Panel _statusSeparator = new();
        private DataTable _dataTable = new();
        private DataTable _headerTable = new();
        private FileTableData _fileData = new();  // 统一数据模型，用于行号追踪
        private TableRow[] _rowTags = Array.Empty<TableRow>();  // 显示行 → 源行记录
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
        /// <summary>冻结列数（显示列口径，冻结最左侧 N 列；0 = 不冻结），按文件持久化</summary>
        private int _frozenColumnCount = 0;
        private bool _editingFromFrozen = false;
        private DataGridLinkToolTip? _linkToolTip;
        private FlowLayoutPanel? _flowTop;
        private Button? _btnReload;
        private Button? _btnHistory;
        private ComboBox? _cboUniqueColumn;
        private bool[]? _externalColumnStates = null;
        private int? _externalTotalColumns = null;    // 外部注入的完整列数（不含锁定可见的第0列主键列，与"列管理(x/y)"口径一致）
        private int? _externalVisibleColumns = null;   // 外部注入的可见列数（同上，不含第0列主键列）
        private bool _columnManagementDisabled = false; // 列管理被禁用（版本对比等场景，保持按钮显示但不响应点击）
        private bool _isEmbeddedMode; // 是否为外部嵌入模式（版本对比等），禁止编辑
        private bool _isEditLocked = false; // Excel默认锁定编辑，CSV默认解锁
        private ContextMenuMode _contextMenuMode = ContextMenuMode.Normal; // 右键菜单模式

        // ==================== 缓存 ====================
        /// <summary>本控件使用的表格数据缓存：宿主传入则共享（随宿主窗口释放），未传则自建（随本控件 Dispose 释放）</summary>
        private readonly FileDataCache _cache;
        private readonly bool _ownsCache;
        private bool _hasUnsavedChanges = false;
        private System.Windows.Forms.Timer? _retrySaveTimer;
        private const int SaveDebounceMs = 3000; // 3秒无操作后自动保存
        /// <summary>已修改但尚未保存到磁盘的原始行索引集合（用于单元格橙色背景标记）</summary>
        private HashSet<int> _unsavedRowIndices = new();

        // ==================== 行编辑模式 ====================
        private int _editingRowIndex = -1;           // 当前正在编辑的显示行索引（-1=无）
        private string[]? _editingRowOriginalValues;  // 编辑前原始值快照（用于Esc取消）
        private bool _enteringEditMode = false;       // 防止事件重入标志
        private bool _isNewRowForEdit = false;        // 是否为新增行（Esc时删除该行）
        private int _editingRowOrigIndex = -1;        // 编辑行的原始行索引
        /// <summary>行编辑模式中已修改的单元格集合（按原始行索引分组，key=原始行索引，value=该行已修改的列索引集合，不同行互不影响）</summary>
        private Dictionary<int, HashSet<int>> _modifiedCellsInEdit = new();
        private bool _syncingColumnWidth = false;     // 列宽同步防重入
        private bool _syncingScroll = false;           // 滚动同步防重入
        /// <summary>用户真正想要的横向滚动偏移（只由主表滚动/固定行键盘导航写入）。
        /// 固定行网格没有横向滚动条，改列宽时它的偏移会被布局夹回 0，回灌主表就表现为视图跳回最左列，
        /// 所以列宽同步结束后要用这个值把两侧复位。</summary>
        private int _preferredHorizontalOffset = 0;

        // ==================== 过滤版本控制 ====================
        private int _filterVersion = 0;        // 递增版本号，用于丢弃过期过滤结果
        private int _pendingFilterVersion = 0; // 正在后台执行的过滤版本号，防重复启动

        /// <summary>唯一列重复警告文本（null = 无警告），用于状态栏持久提示</summary>
        private string? _duplicateWarningText = null;
        /// <summary>唯一列重复值完整列表（供详情窗体使用）</summary>
        private List<DuplicateGroup> _duplicateDetails = new();

        /// <summary>内容搜索匹配的单元格集合（原始行索引 → 原始列索引集合），用于单元格红色加粗高亮，不过滤行</summary>
        private Dictionary<int, HashSet<int>> _contentSearchMatchCells = new();

        /// <summary>高级过滤匹配的单元格集合（原始行索引 → 原始列索引集合）</summary>
        private Dictionary<int, HashSet<int>> _advFilterMatchCells = new();

        /// <summary>过滤匹配的单元格集合（显示行索引 → 该行匹配的显示列索引集合），用于 _dataGrid 红色加粗高亮</summary>
        private Dictionary<int, HashSet<int>> _filterMatchCells = new();

        /// <summary>过滤匹配的单元格集合（显示行索引 → 该行匹配的显示列索引集合），用于 _dataGridHeader 红色加粗高亮</summary>
        private Dictionary<int, HashSet<int>> _headerFilterMatchCells = new();

        /// <summary>高级过滤时是否保留所有行（仅高亮匹配单元格，不过滤行），默认 false</summary>
        private bool _filterKeepAllRows = false;

        /// <summary>重复值分组数据</summary>
        public class DuplicateGroup
        {
            public string Value { get; set; } = "";
            public int Count { get; set; }
            public List<int> RowIndices { get; set; } = new();
        }

        /// <summary>编辑关联表请求事件参数</summary>
        public class EditLinkTableEventArgs : EventArgs
        {
            public string TargetFilePath { get; set; } = "";
            public string DisplayName { get; set; } = "";
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

        /// <summary>在父窗体标签页中编辑打开关联表请求事件</summary>
        public event EventHandler<EditLinkTableEventArgs>? EditLinkTableRequested;

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
        // 全部转发到 Theme 令牌：换皮肤/明暗只改 Theme.Set 的入参，这里不用逐个值调。
        // 用表达式属性而非 readonly 字段，静态上下文也能在每次取色时拿到当前皮肤。
        private static Color _colorSurface => Theme.P.Surface;
        private static Color _colorBorder => Theme.P.Border;
        private static Color _colorLightText => Theme.P.Text;
        private static Color _colorLightTextSecondary => Theme.P.Muted;
        private static Color _colorPrimary => Theme.P.Accent;
        private static Color _colorHoverBg => Theme.P.Hover;
        private static Color _colorHeaderBg => Theme.P.Header;
        private static Color _colorAltRow => Theme.P.Zebra;
        private static Color _colorSelectedBg => Theme.P.Selection;
        private static Color _colorGridLine => Theme.P.GridLine;
        private static Color _colorRowNumBg => Theme.P.RowHead;   // 比数据区更深一档，行号区和数据列一眼能分开
        private static Color _colorRowNumberText => Theme.P.Muted;
        private static Color _colorRowSelectedBorder => Theme.P.Accent; // 选中行左侧强调条
        /// <summary>标记行按源文件物理行号记账：重新绑定/过滤/排序后仍能把标记落到同一行</summary>
        private readonly HashSet<int> _markedLineNumbers = new();
        /// <summary>标记行在当前视图中的显示行索引，由 _markedLineNumbers 映射，每次重绑后必须重建</summary>
        private HashSet<int> _selectedRows = new();

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
            // ★ 缓存归属：宿主给就用宿主的（同目录多标签共享、随宿主窗口释放），没给就自持一份
            _ownsCache = options.Cache == null;
            _cache = options.Cache ?? new FileDataCache();
            // ★ 目标表保存后（本窗口或共享缓存的其它窗口），重算关联着色
            _cache.FileSaved += OnCacheFileSaved;
            _lightweight = options.Lightweight;
            _contextMenuMode = options.ContextMenuMode;
            // ★ 先确定文件类型，后续配置读取需要用到
            _fileType = DetectFileType(options.FilePath);
            _isEditLocked = (_fileType == TableFileType.Excel); // Excel默认锁定编辑
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
                _frozenColumnCount = CsvSettings.GetFileFrozenColumnCount(dir, GetCurrentRelativePath());
            }

            InitializeComponent();

            // 独立表格窗口（阅读器/版本对比/冲突）不挂在主窗的传播链上，切皮肤靠自己订阅重取令牌
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;

            // 极简 UI 模式：隐藏工具栏和底部状态栏，只留高级过滤那一行
            if (options.ShowMinimalUI)
            {
                // _pnlTop = 工具栏 + 分隔线；_pnlFilterBar = [0]分隔线 [1]内容过滤 [2]高级过滤
                _pnlTop.Visible = false;
                if (_pnlFilterBar.Controls.Count > 1) _pnlFilterBar.Controls[1].Visible = false;
                _statusSeparator.Visible = false;
                _pnlStatus.Visible = false;
            }

            // 如果文件路径为空或文件不存在，跳过自动加载（由外部通过 SetDataTable 注入数据）
            if (!options.SkipAutoLoad && !string.IsNullOrEmpty(_filePath) && File.Exists(_filePath))
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
        /// <summary>
        /// 工具条前置标签：文字比同行控件矮，FlowLayoutPanel 只按上边距对齐，
        /// 所以按当前皮肤的行高补一次居中偏移，皮肤切换后不会顶到上边。
        /// </summary>
        private static Label ToolLabel(string text) => new Label
        {
            Text = text,
            AutoSize = true,
            Font = Theme.Body,
            ForeColor = Theme.P.Muted,
            Margin = new Padding(0, Math.Max(2, (ThemeMetrics.RowH - Theme.Body.Height) / 2), 3, 0)
        };

        private void InitializeComponent()
        {
            SuspendLayout();
            BackColor = _colorSurface;

            // 顶栏只剩工具栏一行；搜索/过滤两行改放到底部（表格与状态栏之间）
            var topTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                BackColor = _colorSurface,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0),
                Margin = new Padding(0),
                GrowStyle = TableLayoutPanelGrowStyle.AddRows
            };
            topTable.RowStyles.Clear();
            topTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            topTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 1));

            // 第一行工具栏
            _flowTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Height = ThemeMetrics.RowH + 2 * ThemeMetrics.RowPad,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(ThemeMetrics.Margin, ThemeMetrics.RowPad,
                                     ThemeMetrics.Margin, ThemeMetrics.RowPad),
                AutoSize = true
            };

            // 冻结行数选择
            var lblFrozenRows = ToolLabel("冻结行:");
            var cboFrozenRows = new ThemeComboBox
            {
                Size = new Size(50, ThemeMetrics.RowH),
                Font = Theme.Body,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
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
            var lblHeaderRow = ToolLabel("表头行:");
            var cboHeaderRow = new ThemeComboBox
            {
                Size = new Size(70, ThemeMetrics.RowH),
                Font = Theme.Body,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
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
            var lblMappingRow = ToolLabel("映射行:");
            var cboMappingRow = new ThemeComboBox
            {
                Size = new Size(70, ThemeMetrics.RowH),
                Font = Theme.Body,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
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
            var lblUniqueColumn = ToolLabel("唯一列:");
            _cboUniqueColumn = new ThemeComboBox
            {
                Size = new Size(130, ThemeMetrics.RowH),
                Font = Theme.Body,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
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
            var lblHeaderFilter = ToolLabel("列管理:");
            _btnSelectColumns = new ThemeButton
            {
                Size = new Size(ThemeButton.FitWidth("列管理"), ThemeMetrics.RowH),
                Font = Theme.Body,
                Text = "列管理",
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
            };
            _btnSelectColumns.Click += OnSelectColumnsClick;

            _lblInfo = new Label
            {
                AutoSize = true,
                Font = Theme.Body,
                ForeColor = _colorLightTextSecondary,
                Margin = new Padding(3, Math.Max(2, (ThemeMetrics.RowH - Theme.Body.Height) / 2), 0, 0)
            };

            _flowTop.Controls.AddRange(new Control[] { lblFrozenRows, cboFrozenRows, lblHeaderRow, cboHeaderRow,
                lblMappingRow, cboMappingRow, lblUniqueColumn, _cboUniqueColumn, lblHeaderFilter, _btnSelectColumns });

            // 重新加载按钮
            _btnReload = new ThemeButton
            {
                Text = "重新加载",
                Size = new Size(ThemeButton.FitWidth("重新加载"), ThemeMetrics.RowH),
                Font = Theme.Body,
                Primary = true,
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
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
                Font = Theme.BodyBold,
                ForeColor = _colorPrimary,
                Margin = new Padding(0, 3, ThemeMetrics.Group, 3)
            };
            _lblCellPos = lblCellPos;
            // _lblCellPos 和 _lblInfo 不再添加到 _flowTop，改为放入底部状态栏 _pnlStatus

            // 分隔线
            var separatorLine = new Panel
            {
                Height = 1,
                Dock = DockStyle.Top,
                BackColor = Theme.P.Border,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            // 内容过滤
            var btnContentFilter = new ThemeButton
            {
                Text = "高亮匹配",
                Size = new Size(ThemeButton.FitWidth("高亮匹配"), ThemeMetrics.RowH),
                Font = Theme.Body,
                Margin = new Padding(0, 0, ThemeMetrics.Gap, 0),
                Padding = new Padding(0),
                TextAlign = ContentAlignment.MiddleCenter
            };
            btnContentFilter.Click += (s, e) => TriggerFilter();
            var fldContentFilter = new ThemeField
            {
                Size = new Size(200, ThemeMetrics.RowH),
                Placeholder = "按回车搜索...",
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
            };
            var txtContentFilter = fldContentFilter.Edit;
            txtContentFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { TriggerFilter(); e.Handled = true; e.SuppressKeyPress = true; }
                else if (e.KeyCode == Keys.Escape) { txtContentFilter.Text = ""; TriggerFilter(); e.Handled = true; }
            };

            // 高级过滤
            var btnAdvFilter = new ThemeButton
            {
                Text = "高级过滤",
                Size = new Size(ThemeButton.FitWidth("高级过滤"), ThemeMetrics.RowH),
                Font = Theme.Body,
                Margin = new Padding(0, 0, ThemeMetrics.Gap, 0),
                Padding = new Padding(0),
                TextAlign = ContentAlignment.MiddleCenter
            };
            btnAdvFilter.Click += (s, e) => TriggerFilter();
            var fldAdvancedFilter = new ThemeField
            {
                Size = new Size(200, ThemeMetrics.RowH),
                Placeholder = "按回车过滤，支持 @line=128",
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
            };
            var txtAdvancedFilter = fldAdvancedFilter.Edit;
            txtAdvancedFilter.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { TriggerFilter(); e.Handled = true; e.SuppressKeyPress = true; }
                else if (e.KeyCode == Keys.Escape) { txtAdvancedFilter.Text = ""; TriggerFilter(); e.Handled = true; }
            };
            var btnAdvHelp = new PictureBox
            {
                Size = new Size(25, ThemeMetrics.RowH),
                SizeMode = PictureBoxSizeMode.StretchImage,
                Margin = new Padding(0, 0, ThemeMetrics.Gap, 0),
                Cursor = Cursors.Help,
                BackColor = Color.Transparent
            };
            try
            {
                using var stream = typeof(TableBrowserControl).Assembly
                    .GetManifestResourceStream("wxd.files.winforms.wenhao.jpg");
                if (stream != null)
                    btnAdvHelp.Image = Image.FromStream(stream);
            }
            catch { /* 图片加载失败，用 ? 文本兜底 */ }
            if (btnAdvHelp.Image == null)
            {
                btnAdvHelp.Paint += (s, e) =>
                {
                    using var brush = new SolidBrush(_colorPrimary);
                    // 字体必须是静态只读的一份：Paint 每帧都 new 的话，重绘次数就是 GDI 句柄的泄漏数
                    e.Graphics.DrawString("?", AdvHelpGlyphFont, brush, 5, 3);
                };
            }
            var advToolTip = new ToolTip
            {
                AutoPopDelay = 30000,
                InitialDelay = 0,
                ReshowDelay = 0,
                IsBalloon = true,
                ToolTipIcon = ToolTipIcon.Info,
                ToolTipTitle = "高级过滤语法"
            };
            btnAdvHelp.Click += (s, e) =>
            {
                advToolTip.Show(@"基础操作符:
  id=10001             等于
  id!=10001            不等于
  id>10001             大于
  id>=10001            大于等于
  id<10001             小于
  id<=10001            小于等于
  name like test       包含test(忽略大小写)

通配符(仅like):
  id like *00001       以00001结尾
  id like 00001*       以00001开头
  id like *000*        包含000
  id like 00*01        00开头01结尾

全列模糊(无列名/操作符):
  10001                所有列包含10001
  test                 所有列包含test

无列名搜索(操作符+值):
  =10001               任意列等于10001
  >100                 任意列大于100
  like test            任意列包含test

多条件组合:
  id=10001 and name like test   同时满足(AND)
  id=10001 or id=10002          满足其一(OR)
  id>100 and id<200             范围查询

行号过滤(@line 是虚拟列, 比的是最左侧行号区的行号):
  @line=128                     只看源文件第128行
  @line>=100 and @line<=200     行号区间
  @line>50 and level like 暴击  行号 + 内容组合

列名通过表头行文字识别,不区分大小写", btnAdvHelp, 0, btnAdvHelp.Height + 5);
            };
            var lblAdvHint = new Label
            {
                Text = "如: id=123 or name like test",
                AutoSize = true,
                Font = Theme.Caption,
                ForeColor = _colorPrimary,
                Margin = new Padding(3, Math.Max(2, (ThemeMetrics.RowH - Theme.Caption.Height) / 2),
                                    ThemeMetrics.Group, 0)
            };

            // 这两行是叠起来的，上下留白按标准 RowPad=8 走会翻成两倍高（84px），读表时太占地方；
            // 折到 4px，整块从 85 降到 69，和单行工具条的观感仍然一致。
            const int filterRowPad = 4;

            var flowContent = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Height = ThemeMetrics.RowH + 2 * filterRowPad,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(ThemeMetrics.Margin, filterRowPad,
                                     ThemeMetrics.Margin, filterRowPad),
                AutoSize = true
            };
            flowContent.Controls.AddRange(new Control[] { btnContentFilter, fldContentFilter });

            var flowAdv = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Height = ThemeMetrics.RowH + 2 * filterRowPad,
                BackColor = _colorSurface,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(ThemeMetrics.Margin, filterRowPad,
                                     ThemeMetrics.Margin, filterRowPad),
                AutoSize = true
            };
            flowAdv.Controls.AddRange(new Control[] { btnAdvFilter, fldAdvancedFilter, btnAdvHelp, lblAdvHint });

            // 搜索/过滤条：贴在表格下方、状态栏上方，读表时视线不用回到顶部
            _pnlFilterBar = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                AutoSize = true,
                BackColor = _colorSurface,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(0),
                Margin = new Padding(0),
                GrowStyle = TableLayoutPanelGrowStyle.AddRows
            };
            _pnlFilterBar.RowStyles.Clear();
            _pnlFilterBar.RowStyles.Add(new RowStyle(SizeType.Absolute, 1));
            _pnlFilterBar.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _pnlFilterBar.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var gridDivider = new Panel
            {
                Height = 1,
                Dock = DockStyle.Fill,
                BackColor = Theme.P.Border,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            topTable.Controls.Add(_flowTop, 0, 0);
            topTable.Controls.Add(separatorLine, 0, 1);
            _pnlFilterBar.Controls.Add(gridDivider, 0, 0);
            _pnlFilterBar.Controls.Add(flowContent, 0, 1);
            _pnlFilterBar.Controls.Add(flowAdv, 0, 2);

            _pnlTop = topTable;
            _txtContentFilter = txtContentFilter;
            _txtAdvancedFilter = txtAdvancedFilter;

            // 两条过滤行右边那组控件（过滤按钮 + 帮助 + 示例文字）占掉的宽度，输入框吃掉剩下的。
            // 文案或字号改动时要跟着调：这里是唯一一个横向尺寸不是全靠 ThemeMetrics 推出来的地方。
            const int filterTailWidth = 280;

            _pnlFilterBar.Resize += (s, e) =>
            {
                int w = _pnlFilterBar.ClientSize.Width - filterTailWidth;
                if (w < 100) w = 100;
                fldContentFilter.Width = w;
                fldAdvancedFilter.Width = w;
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
            _dataGridHeader.CellFormatting += OnHeaderCellFormatting;
            StyleGrid(_dataGridHeader, frozenHeader: true);
            _dataGridHeader.TabStop = false;
            _dataGridHeader.CellMouseDown += (s, e) =>
            {
                // 列头右键：冻结到此列 / 取消冻结
                if (e.Button == MouseButtons.Right && e.RowIndex == -1 && e.ColumnIndex >= 0)
                {
                    ShowColumnFreezeMenu(e.ColumnIndex);
                    return;
                }
                ClearRowSelection();
                if (e.RowIndex >= 0 && _dataGridHeader.FirstDisplayedScrollingRowIndex > 0)
                    _dataGridHeader.FirstDisplayedScrollingRowIndex = 0;
                if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    if (_contextMenuMode == ContextMenuMode.None) return;
                    var cellValue = _dataGridHeader.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    var ctx = new ContextMenuStrip();
                    ctx.Font = Theme.Body;
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
                                _linkToolTip.CreatePinnedWindow(content.title, content.lines, content.rawText, screenPos, content.tableHeaders);
                            }
                        }
                    };
                    ctx.Items.Add(pinItem2);
                    // 冻结行区同样提供冻结列入库，入口不随面板布局变化
                    AppendFrozenColumnMenuItems(ctx, e.ColumnIndex, addSeparator: true);
                    // ★ 显示右键菜单前主动隐藏 ToolTip，防止 ToolTip 遮挡右键菜单
                    _linkToolTip?.Hide();
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
                _preferredHorizontalOffset = _dataGrid.HorizontalScrollingOffset;
            };
            _dataGridHeader.CellDoubleClick += OnFrozenCellDoubleClick;
            // CellClick 已由 CreateDataGridView 统一挂接 OnCellClick，此处不再重复绑定

            _dataGrid = CreateDataGridView();
            _dataGrid.Dock = DockStyle.Fill;
            _dataGrid.RowHeadersWidth = 50;
            _dataGrid.RowPostPaint += OnRowPostPaint;
            _dataGrid.CellMouseDown += OnGridCellMouseDown;

            _dataGrid.Scroll += (s, e) =>
            {
                if (_syncingScroll) return;
                _preferredHorizontalOffset = _dataGrid.HorizontalScrollingOffset;
                _syncingScroll = true;
                try { _dataGridHeader.HorizontalScrollingOffset = _dataGrid.HorizontalScrollingOffset; }
                finally { _syncingScroll = false; }
            };
            _dataGrid.ColumnWidthChanged += (s, e) => SyncColumnWidths(_dataGrid, _dataGridHeader);

            // ★ 固定行滚动同步到内容行
            _dataGridHeader.Scroll += (s, e) =>
            {
                if (_syncingScroll) return;
                _syncingScroll = true;
                try { _dataGrid.HorizontalScrollingOffset = _dataGridHeader.HorizontalScrollingOffset; }
                finally { _syncingScroll = false; }
            };

            // ★ 固定行列宽变化同步到内容行
            _dataGridHeader.ColumnWidthChanged += (s, e) => SyncColumnWidths(_dataGridHeader, _dataGrid);

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

            // 阻止编辑模式下方向键跳格（通过 PreviewKeyDown 在 DataGridView.ProcessDialogKey 之前拦截）
            _dataGrid.EditingControlShowing += (s, ev) =>
            {
                if (ev.Control is TextBox tb)
                {
                    tb.KeyDown -= OnEditTextBoxKeyDown;
                    tb.KeyDown += OnEditTextBoxKeyDown;
                    // PreviewKeyDown 在 ProcessDialogKey 之前触发，可阻止 DataGridView 的导航
                    tb.PreviewKeyDown -= OnEditTextBoxPreviewKeyDown;
                    tb.PreviewKeyDown += OnEditTextBoxPreviewKeyDown;
                }
            };

            var headerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 0, SystemInformation.VerticalScrollBarWidth, 0)
            };
            headerPanel.Controls.Add(_dataGridHeader);
            splitContainer.Panel1.Controls.Add(headerPanel);
            _headerPanel = headerPanel;
            splitContainer.Panel2.Controls.Add(_dataGrid);
            splitContainer.Panel1Collapsed = true;
            _pnlGridContainer = splitContainer;

            Controls.Add(_pnlGridContainer);
            Controls.Add(_pnlTop);

            // 底部同侧停靠后加的更靠边，所以从内到外依次是过滤条、分隔线、状态栏
            Controls.Add(_pnlFilterBar);

            // ============ 底部状态栏 ============
            // 状态栏顶部分隔线
            _statusSeparator = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 1,
                BackColor = Theme.P.Border,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            Controls.Add(_statusSeparator);

            _pnlStatus = new Panel
            {
                Dock = DockStyle.Bottom,
                // 和表格列头同一档高度，读表和看状态条是一条线
                Height = Theme.P.RowHeight + 4,
                BackColor = Theme.P.Header,
                Padding = new Padding(0)
            };

            // 使用 TableLayoutPanel 来实现四列布局
            var statusTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 6,
                RowCount = 1,
                BackColor = Color.Transparent,
                Padding = new Padding(4, 0, 4, 0),
                Margin = new Padding(0)
            };
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));   // 过滤保留所有行 CheckBox
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));    // 锁图标
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));  // 文件名
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));   // 行列坐标（自动适应内容宽度，避免换行）
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));   // 重复提示（自动适应内容宽度，避免换行）
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));   // 选中标记行的源文件行号区间

            // 列0: 过滤保留所有行 CheckBox
            _chkFilterKeepAll = new ThemeCheckBox
            {
                Text = "保留所有行",
                Checked = false,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Font = Theme.Body,
                Margin = new Padding(4, 2, 0, 2),
                Padding = new Padding(0),
            };
            _chkFilterKeepAll.CheckedChanged += OnFilterKeepAllChanged;
            statusTable.Controls.Add(_chkFilterKeepAll, 0, 0);

            // 列1: 锁图标 ??/??
            _lblLock = new Label
            {
                Text = "",
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = _colorLightTextSecondary,
                Font = Theme.Body,
                Cursor = Cursors.Hand,
                Padding = new Padding(0, 0, 4, 0),
            };
            _lblLock.Click += OnLockLabelClick;
            statusTable.Controls.Add(_lblLock, 1, 0);

            // 列2: 文件名
            _lblFileName = new Label
            {
                Text = "",
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = _colorLightTextSecondary,
                Font = Theme.Body,
                Dock = DockStyle.Fill,
            };
            statusTable.Controls.Add(_lblFileName, 2, 0);

            // 列3: _lblCellPos 行列坐标（AutoSize 让列宽跟随内容，避免固定 180px 导致换行）
            _lblCellPos.TextAlign = ContentAlignment.MiddleLeft;
            _lblCellPos.AutoSize = true;
            _lblCellPos.Dock = DockStyle.None;
            _lblCellPos.Margin = new Padding(2, 3, 10, 3);
            statusTable.Controls.Add(_lblCellPos, 3, 0);

            // 列4: _lblInfo 重复提示（AutoSize 让列宽跟随内容，避免换行）
            _lblInfo.TextAlign = ContentAlignment.MiddleLeft;
            _lblInfo.AutoSize = true;
            _lblInfo.Dock = DockStyle.None;
            _lblInfo.Margin = new Padding(2, 3, 4, 3);
            statusTable.Controls.Add(_lblInfo, 4, 0);

            // 列5: 选中标记行的源文件行号区间
            _lblRowNumber.Text = "";
            _lblRowNumber.TextAlign = ContentAlignment.MiddleLeft;
            _lblRowNumber.AutoSize = true;
            _lblRowNumber.Dock = DockStyle.None;
            _lblRowNumber.Font = Theme.BodyBold;
            _lblRowNumber.ForeColor = _colorLightText;
            _lblRowNumber.Margin = new Padding(2, 3, 4, 3);
            statusTable.Controls.Add(_lblRowNumber, 5, 0);

            _pnlStatus.Controls.Add(statusTable);
            Controls.Add(_pnlStatus);

            ResumeLayout(false);
        }

        /// <summary>
        /// 列宽双向同步（from 是用户正在拖的那一侧）。
        /// 固定行网格 ScrollBars=None，改列宽会把它内部的 HorizontalScrollingOffset 夹回 0；
        /// 这个 0 原本会经 Scroll 事件回灌主表，表现为"横向滚到右边后拖一下列宽，整张表跳回最左列并重绘"。
        /// 所以同步期间屏蔽 Scroll 互写，结束后按记录的用户滚动意图复位两侧偏移。
        /// </summary>
        private void SyncColumnWidths(DataGridView from, DataGridView to)
        {
            if (_syncingColumnWidth) return;
            _syncingColumnWidth = true;
            _syncingScroll = true;
            try
            {
                // 仍然全量比对（键盘导航后两侧列索引可能错位），但只写真正有差异的列
                int count = Math.Min(from.Columns.Count, to.Columns.Count);
                for (int i = 0; i < count; i++)
                {
                    int width = from.Columns[i].Width;
                    if (to.Columns[i].Width != width)
                        to.Columns[i].Width = width;
                }
            }
            finally
            {
                _syncingColumnWidth = false;
                _syncingScroll = false;
            }

            // 夹取发生在 Width setter 内部、本事件之前，所以只能认 _preferredHorizontalOffset，
            // 不能拿此刻的 _dataGrid.HorizontalScrollingOffset（已经是被拽回来的 0）。
            int offset = _preferredHorizontalOffset;
            if (_dataGrid.HorizontalScrollingOffset != offset)
                _dataGrid.HorizontalScrollingOffset = offset; // 触发 grid.Scroll，记账并带回固定行网格
            if (_dataGridHeader.HorizontalScrollingOffset != _dataGrid.HorizontalScrollingOffset)
                _dataGridHeader.HorizontalScrollingOffset = _dataGrid.HorizontalScrollingOffset;
        }

        // ==================== 冻结列 ====================

        /// <summary>列头右键菜单：冻结到此列 / 取消冻结（Excel 式冻结最左侧若干列）</summary>
        private void ShowColumnFreezeMenu(int displayColumnIndex)
        {
            if (_columnManagementDisabled || _dataGrid.Columns.Count <= 1) return;
            var ctx = new ContextMenuStrip();
            ctx.Font = Theme.Body;
            AppendFrozenColumnMenuItems(ctx, displayColumnIndex, addSeparator: false);
            _linkToolTip?.Hide();
            // 不绑定 owner：冻结行面板折叠时列头在 _dataGrid 上，菜单跟随鼠标位置显示
            ctx.Show(Cursor.Position);
        }

        /// <summary>把冻结列菜单项追加到已有右键菜单（列头菜单和单元格菜单共用）</summary>
        private void AppendFrozenColumnMenuItems(ContextMenuStrip ctx, int displayColumnIndex, bool addSeparator)
        {
            if (_columnManagementDisabled || _dataGrid.Columns.Count <= 1) return;
            if (addSeparator) ctx.Items.Add(new ToolStripSeparator());

            int desired = displayColumnIndex + 1;
            // 冻结全部列没有意义（也不留可滚动区域），最多冻结到倒数第二列
            if (desired < _dataGrid.Columns.Count)
            {
                var freezeItem = new ToolStripMenuItem($"冻结到此列（前 {desired} 列）")
                {
                    Checked = _frozenColumnCount == desired
                };
                freezeItem.Click += (s, e) => SetFrozenColumnCount(desired);
                ctx.Items.Add(freezeItem);
            }
            var unfreezeItem = new ToolStripMenuItem("取消冻结列")
            {
                Enabled = _frozenColumnCount > 0
            };
            unfreezeItem.Click += (s, e) => SetFrozenColumnCount(0);
            ctx.Items.Add(unfreezeItem);
        }

        private void SetFrozenColumnCount(int count)
        {
            // 至少保留一列可滚动区域，同时保证保存到配置里的值一定是可用的
            count = Math.Clamp(count, 0, Math.Max(0, _dataGrid.Columns.Count - 1));
            if (count == _frozenColumnCount) return;
            _frozenColumnCount = count;
            if (!_isEmbeddedMode && !_columnManagementDisabled && !string.IsNullOrEmpty(FrozenDir))
                CsvSettings.SetFileFrozenColumnCount(FrozenDir, GetCurrentRelativePath(), count);
            ApplyFrozenColumns();
        }

        /// <summary>冻结行面板显示/隐藏。隐藏时内容区自己显示列头，否则整张表一个列头都没有。</summary>
        private void SetFrozenPanelVisible(bool visible)
        {
            _pnlGridContainer.Panel1Collapsed = !visible;
            _dataGrid.ColumnHeadersVisible = !visible;
        }

        /// <summary>把 _frozenColumnCount 应用到两个网格。列数变化（隐藏列/刷新）后需重新调用</summary>
        private void ApplyFrozenColumns()
        {
            int desired = Math.Clamp(_frozenColumnCount, 0, Math.Max(0, _dataGrid.Columns.Count - 1));
            // 隐藏列后原冻结数可能越界，回写到字段，保证菜单勾选态和实际一致。
            // Columns 为空（尚未绑定）时不回写，避免把已保存的冻结数临时清成 0。
            if (_dataGrid.Columns.Count > 0)
                _frozenColumnCount = desired;
            ApplyFrozenColumnsToGrid(_dataGrid, desired);
            ApplyFrozenColumnsToGrid(_dataGridHeader, desired);

            // 冻结会缩小可滚动宽度，按用户滚动意图重新定位，并保持两侧一致
            int offset = _preferredHorizontalOffset;
            if (_dataGrid.HorizontalScrollingOffset != offset)
                _dataGrid.HorizontalScrollingOffset = offset;
            if (_dataGridHeader.HorizontalScrollingOffset != _dataGrid.HorizontalScrollingOffset)
                _dataGridHeader.HorizontalScrollingOffset = _dataGrid.HorizontalScrollingOffset;
        }

        private static void ApplyFrozenColumnsToGrid(DataGridView dgv, int count)
        {
            if (dgv.Columns.Count == 0) return;
            count = Math.Clamp(count, 0, Math.Max(0, dgv.Columns.Count - 1));
            var ordered = dgv.Columns.Cast<DataGridViewColumn>().OrderBy(c => c.DisplayIndex).ToList();
            bool matched = true;
            for (int i = 0; i < ordered.Count; i++)
            {
                if (ordered[i].Frozen != (i < count)) { matched = false; break; }
            }
            if (matched) return;
            // 先解冻范围外的列（从右往左），避免冻结区中间出现空洞导致 DGV 重排列；
            // 再按显示顺序冻结最左侧 count 列。
            for (int i = ordered.Count - 1; i >= count; i--)
            {
                if (ordered[i].Frozen) ordered[i].Frozen = false;
            }
            for (int i = 0; i < count; i++)
            {
                if (!ordered[i].Frozen) ordered[i].Frozen = true;
            }
        }

        /// <summary>DataGridView 的 DoubleBuffered 是 protected，只能靠派生类打开；不开时宽表逐格直绘，52 列下逐行滚动 91ms→12ms</summary>
        private sealed class BufferedDataGridView : DataGridView
        {
            public BufferedDataGridView() => DoubleBuffered = true;
        }

        /// <summary>
        /// 把当前皮肤的网格参数写进表格。可重入：换皮肤/明暗后对每个网格再调一次。
        /// 只动视觉属性，不碰数据绑定、冻结列滚动同步和行号记账。
        /// </summary>
        /// <param name="frozenHeader">是否为冻结列的那个网格：它不该有选中高亮，底色跟表头同档。</param>
        private void StyleGrid(DataGridView dgv, bool frozenHeader = false)
        {
            var p = Theme.P;
            _rowNumberFont = Theme.Mono;

            dgv.BackgroundColor = p.Surface;
            dgv.GridColor = p.GridLine;
            dgv.CellBorderStyle = p.ShowGridLines
                ? DataGridViewCellBorderStyle.Single
                : DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.ColumnHeadersHeight = p.RowHeight + 4;
            dgv.RowTemplate.Height = p.RowHeight;

            // 行号区：整格样式优先级低于单元格自身的 Style，所以三个态都要给
            dgv.RowHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                BackColor = p.RowHead,
                ForeColor = p.Muted,
                SelectionBackColor = p.RowHead,
                SelectionForeColor = p.Text,
                Font = _rowNumberFont
            };

            dgv.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = p.ZebraOn ? p.Zebra : p.Surface,
                SelectionBackColor = p.Selection,
                SelectionForeColor = p.Text
            };

            dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = p.HeaderFill ? p.Header : p.Surface,
                ForeColor = p.Text,
                Font = Theme.Head,
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                SelectionBackColor = p.HeaderFill ? p.Header : p.Surface,
                SelectionForeColor = p.Text
            };

            var cellBack = frozenHeader ? p.Header : p.Surface;
            dgv.DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = cellBack,
                ForeColor = p.Text,
                Font = Theme.Body,
                // 冻结网格只是主表的镜像，选中态由主表表达，这里不再叠一层高亮
                SelectionBackColor = frozenHeader ? cellBack : p.Selection,
                SelectionForeColor = p.Text,
                Padding = new Padding(p.CellPadX, 0, p.CellPadX, 0)
            };
        }

        /// <summary>
        /// 皮肤或明暗变了：网格按新令牌重刷，其余子控件交给 <see cref="ThemeRecolor"/> 平移旧令牌色。
        /// 这个界面有上百个创建时一次性取色的局部控件，逐个重赋一定会漏，所以走通用平移。
        /// </summary>
        public void ApplyTheme()
        {
            if (IsDisposed) return;
            ThemeRecolor.Apply(this);
            _rowNumberFont = Theme.Mono;
            StyleGrid(_dataGrid);
            StyleGrid(_dataGridHeader, frozenHeader: true);
            _dataGrid.Invalidate();
            _dataGridHeader.Invalidate();
            Invalidate(true);
        }

        private DataGridView CreateDataGridView()
        {
            var dgv = new BufferedDataGridView
            {
                BackgroundColor = _colorSurface,
                BorderStyle = BorderStyle.None,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,  // 默认只读（浏览模式），行编辑模式下临时设为false
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,  // 整行标记由 _selectedRows 单独绘制，不叠加 DGV 多选
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                RowHeadersVisible = true,
                RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing,
                EnableHeadersVisualStyles = false
            };
            StyleGrid(dgv);
            dgv.ColumnHeaderMouseClick += OnColumnHeaderClick;
            dgv.CellFormatting += OnCellFormatting;
            dgv.CellDoubleClick += OnCellDoubleClick;
            dgv.CellClick += OnCellClick;
            dgv.RowHeaderMouseClick += OnRowHeaderMouseClick;
            dgv.CellToolTipTextNeeded += OnCellToolTipTextNeeded;
            // 选中变化只重绘新旧当前行：整表 Invalidate 在宽表下每次约 200ms
            int lastRepaintedRow = -1;
            dgv.SelectionChanged += (s, e) =>
            {
                int current = dgv.CurrentCell?.RowIndex ?? -1;
                if (current == lastRepaintedRow) return;
                InvalidateRow(dgv, lastRepaintedRow);
                InvalidateRow(dgv, current);
                lastRepaintedRow = current;
            };
            dgv.CellBeginEdit += OnCellBeginEdit;
            dgv.CellEndEdit += OnCellEndEdit;
            dgv.CellValueChanged += OnCellValueChangedInEdit;
            return dgv;
        }

        /// <summary>行索引可能已随排序/筛选失效，越界直接忽略</summary>
        private static void InvalidateRow(DataGridView dgv, int rowIndex)
        {
            if (rowIndex >= 0 && rowIndex < dgv.Rows.Count)
                dgv.InvalidateRow(rowIndex);
        }

        // ==================== 数据加载 ====================

        private async void LoadCsvAsync()
        {
            if (_isLoading) return;  // ★ 防重入（DoEvents 期间可能触发二次加载）
            _isLoading = true;
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
                // ★ 后台线程：LoadFile（读缓存/解析）
                var result = await Task.Run(() => LoadCsvFile());
                if (result != null)
                {
                    var (dataTable, fileData, delimiter) = result.Value;
                    _delimiter = delimiter;
                    // ★ UI 线程：OnDataLoadedAsync（含分批绑定 + 进度条，冻结行场景后台构建 filteredTable）
                    await OnDataLoadedAsync(dataTable, dataTable.Clone(), fileData,
                        $"分隔符: '{_delimiter}' | 行数: {dataTable.Rows.Count} | 列数: {dataTable.Columns.Count}");
                }
                var elapsed = (DateTime.UtcNow - loadingStart).TotalMilliseconds;
                if (elapsed < 200) await Task.Delay((int)(200 - elapsed));
                // ★ 延迟 400ms 关闭遮罩，让表格完成最终渲染（ResumeLayout 后的布局重算、高亮重绘），
                //    避免遮罩消失瞬间表格还在刷新导致的闪烁
                await Task.Delay(400);
                HideLoadingOverlay();
            }
            catch (Exception ex)
            {
                HideLoadingOverlay();
                if (this.IsHandleCreated)
                    this.Invoke(() => AlertForm.Show(this, $"加载CSV文件失败:" + GlobalConst.NewLine + ex.Message, "错误", AlertType.Error, exception: ex));
            }
            finally
            {
                _isLoading = false;
            }
        }

        private async void LoadExcelAsync()
        {
            if (_isLoading) return;  // ★ 防重入
            _isLoading = true;
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
                // ★ 后台线程：LoadFile（读缓存/解析）
                var result = await Task.Run(() => LoadExcelFile());
                if (result != null)
                {
                    var (dataTable, fileData) = result.Value;
                    // ★ UI 线程：OnDataLoadedAsync（含分批绑定 + 进度条）
                    await OnDataLoadedAsync(dataTable, dataTable.Clone(), fileData,
                        $"行数: {dataTable.Rows.Count} | 列数: {dataTable.Columns.Count}");
                }
                var elapsed = (DateTime.UtcNow - loadingStart).TotalMilliseconds;
                if (elapsed < 200) await Task.Delay((int)(200 - elapsed));
                // ★ 延迟 400ms 关闭遮罩，让表格完成最终渲染，避免闪烁
                await Task.Delay(400);
                HideLoadingOverlay();
            }
            catch (Exception ex)
            {
                HideLoadingOverlay();
                if (this.IsHandleCreated)
                    this.Invoke(() => AlertForm.Show(this, $"加载Excel文件失败:" + GlobalConst.NewLine + ex.Message, "错误", AlertType.Error, exception: ex));
            }
            finally
            {
                _isLoading = false;
            }
        }

        // ★ 改为返回数据，由调用方（UI 线程）调用 OnDataLoadedAsync
        //    这样 OnDataLoadedAsync 内的 await 才能在 UI 线程正常工作
        private (DataTable dataTable, FileTableData fileData, char delimiter)? LoadCsvFile()
        {
            if (!File.Exists(_filePath))
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => AlertForm.Show(this, "CSV文件不存在", "错误", AlertType.Error));
                return null;
            }

            var (dataTable, fileData) = _cache.LoadFile(_filePath);
            return (dataTable, fileData, fileData.Delimiter);
        }

        private (DataTable dataTable, FileTableData fileData)? LoadExcelFile()
        {
            if (!File.Exists(_filePath))
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => AlertForm.Show(this, "Excel文件不存在", "错误", AlertType.Error));
                return null;
            }

            var (dataTable, fileData) = _cache.LoadFile(_filePath);
            if (dataTable.Rows.Count == 0 && dataTable.Columns.Count == 0)
            {
                if (this.IsHandleCreated)
                    this.Invoke(() => AlertForm.Show(this, "Excel文件没有工作表", "警告", AlertType.Warning));
                return null;
            }
            return (dataTable, fileData);
        }

        private async Task OnDataLoadedAsync(DataTable dataTable, DataTable headerTable, FileTableData fileData, string infoText)
        {
            _dataTable = dataTable;
            // 数据重新加载后旧标记对应的是上一份数据，直接清掉（重新加载/切换文件场景）
            ClearRowSelection();
            // ★ 直接使用 LoadFile 返回的 FileTableData（首次加载无需 inheritFrom，避免第三次全量遍历）
            //    BuildFileData 仍保留供排序/过滤/版本对比场景调用
            _fileData = fileData;
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
            // ★ 对引用模板的列，从模板加载最新配置填充运行时数据
            foreach (var link in _columnLinks)
            {
                link.RefreshFromTemplate();
            }
            // ★ 后台异步构建关联 lookup，不阻塞表格显示
            if (_columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_columnLinks, _filePath);

            UpdateColumnSelectorButton();
            RebuildHighlightRows();
            _lblFileName.Text = GetDisplayFileName();
            UpdateLockDisplay();
            _lblCellPos.Text = infoText;
            _lblInfo.Text = "";

            if (_lightweight)
            {
                // 轻量模式：直接绑定 DataSource，不做额外初始化
                BindDataTableInBatches(_dataTable, configureColumns: true);
                // 轻量模式：顶部工具栏和底部过滤条都收起来，只留表格
                _pnlTop.Visible = false;
                _pnlFilterBar.Visible = false;
            }
            else if (_frozenRows > 0)
                await RefreshGridLayoutWithProgressAsync();  // ★ 冻结行场景：异步构建 filteredTable + 进度条
            else
            {
                BindDataTableInBatches(_dataTable, configureColumns: true);
            }

            if (!string.IsNullOrEmpty(_initialContentFilter))
            {
                // 从左侧文件列表"搜索内容"匹配后打开文件时，将搜索关键词填充到高级过滤（过滤掉不匹配行），
                // 而非内容搜索（仅高亮不过滤行）。无运算符的关键词会被当作全列模糊匹配处理。
                _txtAdvancedFilter.Text = _initialContentFilter;
                TriggerFilter();
            }

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
        private (string? title, List<DataGridLinkToolTip.TipLine> lines, string? rawText, List<string>? tableHeaders) GetLinkToolTipContent(DataGridView grid, int rowIndex, int columnIndex)
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

                return (null, new List<DataGridLinkToolTip.TipLine>(), tipText, null);
            }

            if (rowIndex < 0 || columnIndex < 0)
                return (null, new List<DataGridLinkToolTip.TipLine>(), null, null);

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
            if (string.IsNullOrEmpty(cellVal)) return (null, new List<DataGridLinkToolTip.TipLine>(), null, null);

            // 计算真实行号：_dataGrid 的行索引从 _frozenRows 开始
            int realRowIndex = isHeader ? rowIndex : rowIndex + _frozenRows;
            // 行号统一取 DisplayLineNumber：过滤/排序后显示的仍是源表行号
            int? displayLine = isHeader
                ? GetFrozenDisplayLineNumber(rowIndex)
                : GetDisplayLineNumber(rowIndex);
            int lineLabel = displayLine ?? realRowIndex + 1;

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

                string cellTitle = $"{headerName}  --  第{lineLabel}行";
                var cellLines = new List<DataGridLinkToolTip.TipLine>
                {
                    new DataGridLinkToolTip.TipLine { Text = cellVal, IsWarning = false }
                };
                return (cellTitle, cellLines, null, null);
            }

            // 获取当前文件相对路径
            string currentRel = GetCurrentRelativePath();

            // 查找关联配置
            var linkConfig = _columnLinks.FirstOrDefault(l =>
                LinkSourceMatcher.Matches(l.SourceFile, currentRel, Path.GetFileName(_filePath)) &&
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

                string cellTitle = $"{headerName}  --  第{lineLabel}行";
                var cellLines = new List<DataGridLinkToolTip.TipLine>
                {
                    new DataGridLinkToolTip.TipLine { Text = cellVal, IsWarning = false }
                };
                return (cellTitle, cellLines, null, null);
            }

                        // 提取匹配值
            var matchValues = (linkConfig.ValueExtractionEnabled)
                ? linkConfig.ExtractValues(cellVal) : new List<string> { cellVal };

            // ★ 格式模式：有切割配置但没有关联目标段
            //   直接用列自身的值切割后以表格模式显示
            //   注意：不使用 ExtractValues 的结果（它会完全展开所有层级），
            //   而是从原始 cellVal 用提取规则构建表格行
            bool hasFieldLinks = linkConfig.FieldLinks != null && linkConfig.FieldLinks.Any(f => !string.IsNullOrEmpty(f.TargetFile));
            if (!hasFieldLinks)
            {
                return BuildFormatOnlyToolTipContent(linkConfig, cellVal, realRowIndex, originalColIndex, lineLabel);
            }

            // ★ 有关联配置：对每个匹配值的每个字段，分别查询关联数据
            return BuildFieldLinkedToolTipContent(linkConfig, cellVal, realRowIndex, originalColIndex, currentRel, lineLabel);
        }

        /// <summary>构建多字段关联的 tooltip 内容（每个匹配值的每个字段分别查关联）</summary>
        private (string? title, List<DataGridLinkToolTip.TipLine> lines, string? rawText, List<string>? tableHeaders)
            BuildFieldLinkedToolTipContent(ColumnLinkConfig linkConfig, string cellVal,
                int realRowIndex, int originalColIndex, string currentRel, int lineLabel)
        {
            // ★ 获取表头配置
            List<string>? headers = null;
            if (!string.IsNullOrEmpty(linkConfig.DisplayColumnHeaders))
            {
                headers = linkConfig.DisplayColumnHeaders
                    .Split(new[] { linkConfig.EffectiveDisplayDelimiter }, StringSplitOptions.None)
                    .Select(h => h.Trim())
                    .ToList();
            }

            // ★ 获取列名用于标题
            string colHeaderName = "";
            if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
                colHeaderName = _dataTable.Rows[_headerRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
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

            // ★ 构建关联标题：列出所有有关联的字段
            var linkDescs = new List<string>();
            foreach (var fl in linkConfig.FieldLinks)
            {
                if (!string.IsNullOrEmpty(fl.TargetFile))
                {
                    string displayFn = Path.GetFileName(fl.TargetFile);
                    if (_fileRemarks.TryGetValue(fl.TargetFile, out var remark) && !string.IsNullOrEmpty(remark))
                        displayFn = $"{remark} ({displayFn})";
                    string colInfo = !string.IsNullOrEmpty(fl.TargetColumnName) ? $"[{fl.TargetColumnName}]" : $"[列{fl.TargetColumnIndex}]";
                    linkDescs.Add($"{displayFn}{colInfo}");
                }
            }
            string linkTitlePart = string.Join(", ", linkDescs);
            string title = $"关联: {linkTitlePart}{GlobalConst.LfChar}来自: {colHeaderName} — 第{lineLabel}行";

            // ★ 对每个匹配值，查询所有字段的关联
            var lines = new List<DataGridLinkToolTip.TipLine>();

            // ★ 使用 ExtractRows 按行列语义展开（前 N-1 步行，最后一步列）
            var rows = linkConfig.ExtractRows(cellVal);

            // ★ 确定总列数
            int totalCols = rows.Count > 0 ? rows.Max(r => r.Count) : 1;
            if (headers != null && headers.Count > totalCols)
                totalCols = headers.Count;

            foreach (var rowCells in rows)
            {
                var cells = new List<string>();
                bool hasWarning = false;

                for (int colIdx = 0; colIdx < totalCols; colIdx++)
                {
                    string fieldVal = colIdx < rowCells.Count ? rowCells[colIdx] : "";

                    // 查找此列是否有关联配置
                    var fl = linkConfig.FieldLinks.FirstOrDefault(f => f.FieldIndex == colIdx);

                    if (fl == null || string.IsNullOrEmpty(fl.TargetFile) || string.IsNullOrEmpty(fieldVal))
                    {
                        cells.Add(fieldVal);
                        continue;
                    }

                    string fieldPrefix = $"{currentRel}|{originalColIndex}->field{fl.FieldIndex}->{fl.TargetFile}:{fl.TargetColumnIndex}";
                    var fieldLookups = _cache.GetLookupsByPrefix(fieldPrefix);
                    LinkLookupEntry? fieldLookup = fieldLookups.Count > 0 ? fieldLookups[0] : null;

                    if (fieldLookup != null && fieldLookup.Values.Contains(fieldVal))
                    {
                        string displayName = "";
                        if (fl.TargetDisplayColumnIndex >= 0 && fieldLookup.DisplayNames.TryGetValue(fieldVal, out var dn) && !string.IsNullOrEmpty(dn))
                            displayName = $"{fieldVal}({dn})";
                        else if (fieldLookup.DisplayNames.TryGetValue(fieldVal, out var dn2) && !string.IsNullOrEmpty(dn2))
                            displayName = $"{fieldVal}({dn2})";
                        else
                            displayName = fieldVal;
                        cells.Add(displayName);
                    }
                    else if (fieldLookup != null)
                    {
                        cells.Add($"{fieldVal}(未找到)");
                        hasWarning = true;
                    }
                    else
                    {
                        cells.Add(fieldVal);
                    }
                }

                lines.Add(new DataGridLinkToolTip.TipLine
                {
                    Text = string.Join("\t", cells),
                    IsWarning = hasWarning,
                    Cells = cells
                });
            }

            // ★ 排序：警告行在最前面
            var orderedLines = lines
                .Where(l => l.IsWarning)
                .Concat(lines.Where(l => !l.IsWarning))
                .ToList();

            // ★ 如果用户填写了表头，调整列数以匹配实际 cells 列数
            if (headers != null && orderedLines.Count > 0)
            {
                int actualColCount = orderedLines.Max(l => l.Cells?.Count ?? 1);
                while (headers.Count < actualColCount)
                    headers.Add($"列{headers.Count}");
            }

            return (title, orderedLines, null, headers);
        }

        /// <summary>构建格式模式（无关联目标）的 tooltip 内容，直接用列自身数据切割后以表格显示</summary>
        private (string? title, List<DataGridLinkToolTip.TipLine> lines, string? rawText, List<string>? tableHeaders)
            BuildFormatOnlyToolTipContent(ColumnLinkConfig linkConfig, string cellVal, int realRowIndex, int originalColIndex, int lineLabel)
        {
            // ★ 获取表头配置
            List<string>? headers = null;
            if (!string.IsNullOrEmpty(linkConfig.DisplayColumnHeaders))
            {
                headers = linkConfig.DisplayColumnHeaders
                    .Split(new[] { linkConfig.EffectiveDisplayDelimiter }, StringSplitOptions.None)
                    .Select(h => h.Trim())
                    .ToList();
            }

            // ★ 获取列名用于标题
            string colHeaderName = "";
            if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && originalColIndex < _dataTable.Columns.Count)
                colHeaderName = _dataTable.Rows[_headerRowIndex][originalColIndex]?.ToString()?.Trim() ?? "";
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

            string title = $"{colHeaderName}  --  第{lineLabel}行";

            // ★ 使用 ExtractRows 按行列语义展开（前 N-1 步行，最后一步列）
            var rows = linkConfig.ExtractRows(cellVal);

            // ★ 确定总列数：取所有行的最大列数
            int colCount = rows.Count > 0 ? rows.Max(r => r.Count) : 1;
            if (headers != null && headers.Count > colCount)
                colCount = headers.Count;

            var lines = new List<DataGridLinkToolTip.TipLine>();
            foreach (var rowCells in rows)
            {
                var cells = new List<string>();
                for (int j = 0; j < colCount; j++)
                    cells.Add(j < rowCells.Count ? rowCells[j] : "-"); // 不足的列用 "-" 填充

                lines.Add(new DataGridLinkToolTip.TipLine
                {
                    Text = string.Join("\t", cells),
                    IsWarning = false,
                    Cells = cells
                });
            }

            // ★ 排序
            var orderedLines = lines
                .Where(l => l.IsWarning)
                .Concat(lines.Where(l => !l.IsWarning))
                .ToList();

            // ★ 如果用户填写了表头，调整列数以匹配实际 cells 列数
            if (headers != null && orderedLines.Count > 0)
            {
                int actualColCount = orderedLines.Max(l => l.Cells?.Count ?? 1);
                while (headers.Count < actualColCount)
                    headers.Add($"列{headers.Count}");
            }

            return (title, orderedLines, null, headers);
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
                    // 原来是写死的半透明白，深色皮肤里等于糊一屏白板；遮罩跟着窗口色走
                    BackColor = Theme.P.Window,
                    Dock = DockStyle.Fill,
                    Visible = true
                };
                _loadingLabel = new Label
                {
                    Text = "正在加载数据，请稍候...",
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = Theme.TitleBold,
                    ForeColor = Theme.P.Link,
                    Dock = DockStyle.Top,
                    Height = 60
                };
                _loadingProgressBar = new ThemeProgressBar
                {
                    Width = 400,
                    Height = 24
                };
                _pnlLoading.Controls.Add(_loadingProgressBar);
                _pnlLoading.Controls.Add(_loadingLabel);
                this.Controls.Add(_pnlLoading);
                _pnlLoading.BringToFront();
                // 居中进度条
                _pnlLoading.Layout += (s, e) =>
                {
                    if (_loadingProgressBar != null)
                    {
                        _loadingProgressBar.Left = (_pnlLoading.Width - _loadingProgressBar.Width) / 2;
                        _loadingProgressBar.Top = (_pnlLoading.Height - _loadingProgressBar.Height) / 2 + 20;
                    }
                };
                Application.DoEvents();
            });
        }

        // 新增：更新加载进度（在分批绑定过程中调用，让进度条刷新）
        private void UpdateLoadingProgress(int percent, string stageText)
        {
            if (_pnlLoading == null || !_pnlLoading.Visible || _pnlLoading.IsDisposed) return;
            if (_loadingProgressBar != null)
            {
                _loadingProgressBar.Value = Math.Clamp(percent, 0, 100);
                // 自绘控件不吃父面板的 Refresh，得自己催一次重绘
                _loadingProgressBar.Invalidate();
                _loadingProgressBar.Update();
            }
            if (_loadingLabel != null)
                _loadingLabel.Text = stageText;
            _pnlLoading.Refresh();
            Application.DoEvents();
        }

        /// <summary>
        /// 分批构建 DataTable 再一次性绑定到 DataGridView，避免一次性绑定大表导致 UI 长时间冻结。
        /// 思路：先在【未绑定】的 displayTable 上分批 ImportRow（不触发 DataGridView 事件，极快），
        /// 每批之间刷新进度条，全部填充完成后一次性绑定。
        /// 注意：不要在 displayTable 已绑定 DataGridView 的情况下 ImportRow，
        /// 否则每次 ImportRow 都会触发 ListChanged → DataGridView 同步创建行对象 + 布局，
        /// 导致 UI 线程被占满，进度条无法刷新。
        /// 绑定完成后 _dataTable 引用替换为 displayTable，保证后续编辑/排序基于同一份数据。
        /// </summary>
        private void BindDataTableInBatches(DataTable sourceTable, bool configureColumns)
        {
            int totalRows = sourceTable.Rows.Count;
            // 小表直接绑定，不走分批（避免无谓开销）
            if (totalRows <= 500)
            {
                _dataGrid.DataSource = sourceTable;
                if (configureColumns)
                {
                    for (int i = 0; i < _dataGrid.Columns.Count; i++)
                    {
                        _dataGrid.Columns[i].Width = 120;
                        _dataGrid.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
                        _dataGrid.Columns[i].HeaderText = GetExcelColumnName(i);
                    }
                }
                // 行号列取的是显示行→源行映射（TableRow → 源文件物理行号），这条路径不登记映射的话
                // 首屏行号会退化成连续记录号，带引号换行的 CSV 就和源文件行号对不上
                RebuildRowMap();
                ApplyFrozenColumns();
                return;
            }

            // ★ 阶段1：创建【未绑定】的 displayTable（只有列结构），分批 ImportRow 填充数据行
            //    关键：此时 displayTable 未绑定到 DataGridView，ImportRow 不会触发任何 UI 事件，
            //    Application.DoEvents() 能真正刷新进度条（不会被 DataGridView 布局占满 UI 线程）
            UpdateLoadingProgress(5, "正在准备数据...");
            var displayTable = sourceTable.Clone();  // 克隆结构，0 行，未绑定

            int batchSize = 2000;
            int added = 0;
            for (int start = 0; start < totalRows; start += batchSize)
            {
                int end = Math.Min(start + batchSize, totalRows);
                for (int i = start; i < end; i++)
                    displayTable.ImportRow(sourceTable.Rows[i]);
                added = end;
                int percent = 5 + (int)((added / (double)totalRows) * 75);  // 5% ~ 80%
                UpdateLoadingProgress(percent, $"正在加载行数据... {added}/{totalRows}");
            }
            // ★ 阶段2：数据已全部填充到 displayTable，现在一次性绑定到 DataGridView
            //    绑定本身只要零点几秒（DataGridViewRow 是按需建的），但进度条已走到 80%，
            //    用户知道数据已加载完成、正在渲染
            UpdateLoadingProgress(85, "正在渲染表格（可能需要数秒，请稍候）...");
            _dataGrid.DataSource = displayTable;

            // ★ 阶段3：配置列宽/表头（此时 DataGridView 已有行，但列数固定，操作很快）
            //    用 SuspendLayout/ResumeLayout 包裹，避免每列设置都触发布局重算
            if (configureColumns)
            {
                _dataGrid.SuspendLayout();
                for (int i = 0; i < _dataGrid.Columns.Count; i++)
                {
                    _dataGrid.Columns[i].Width = 120;
                    _dataGrid.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
                    _dataGrid.Columns[i].HeaderText = GetExcelColumnName(i);
                }
                _dataGrid.ResumeLayout(false);
            }

            // ★ 阶段4：统一引用，保证后续编辑/排序基于同一份数据
            _dataTable = displayTable;
            RebuildRowMap();
            // 这条路径不走 RefreshGridLayout，持久化的冻结列必须在这里补一次
            ApplyFrozenColumns();
            UpdateLoadingProgress(100, "完成");
        }

        private void HideLoadingOverlay()
        {
            if (!this.IsHandleCreated) return;
            this.Invoke(() =>
            {
                if (_pnlLoading != null && !_pnlLoading.IsDisposed && _pnlLoading.Parent == this)
                {
                    var covered = _pnlLoading.Bounds;
                    this.Controls.Remove(_pnlLoading);
                    _pnlLoading.Dispose();
                    _pnlLoading = null;
                    _loadingProgressBar = null;
                    _loadingLabel = null;
                    // 遮罩是盖在整个控件上的，移除后 WinForms 不会重绘被它遮住的子控件，
                    // 冻结行区会残留遮罩的白底（看起来像表头没渲染），必须连子控件一起失效并立即重绘
                    this.Invalidate(covered, true);
                    this.Update();
                }
            });
        }

        // ==================== 重新加载 ====================
        private async void ReloadContentAsync()
        {
            if (!File.Exists(_filePath)) return;
            if (_isLoading) return;  // ★ 防重入
            _isLoading = true;
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
                ShowLoadingOverlay();
                await Task.Delay(50);
                // ★ 适配 LoadCsvFile/LoadExcelFile 新签名（返回数据）
                if (_fileType == TableFileType.Csv)
                {
                    var result = await Task.Run(() => LoadCsvFile());
                    if (result != null)
                    {
                        var (dataTable, fileData, delimiter) = result.Value;
                        _delimiter = delimiter;
                        await OnDataLoadedAsync(dataTable, dataTable.Clone(), fileData,
                            $"分隔符: '{_delimiter}' | 行数: {dataTable.Rows.Count} | 列数: {dataTable.Columns.Count}");
                    }
                }
                else
                {
                    var result = await Task.Run(() => LoadExcelFile());
                    if (result != null)
                    {
                        var (dataTable, fileData) = result.Value;
                        await OnDataLoadedAsync(dataTable, dataTable.Clone(), fileData,
                            $"行数: {dataTable.Rows.Count} | 列数: {dataTable.Columns.Count}");
                    }
                }
                // ★ 延迟 400ms 关闭遮罩，让表格完成最终渲染，避免闪烁
                await Task.Delay(400);
                HideLoadingOverlay();
            }
            catch (Exception ex)
            {
                HideLoadingOverlay();
                if (this.IsHandleCreated)
                    this.Invoke(() => AlertForm.Show(this, $"重新加载失败:" + GlobalConst.NewLine + ex.Message, "错误", AlertType.Error, exception: ex));
                _columnCheckStates = savedColumnStates;
                _frozenRows = savedFrozenRows;
                _txtContentFilter.Text = savedFilter;
                _txtAdvancedFilter.Text = savedAdvFilter;
                UpdateColumnSelectorButton();
                return;
            }
            finally
            {
                _isLoading = false;
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
            // ★ 后台异步构建关联 lookup，不阻塞表格显示
            if (_columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_columnLinks, _filePath);
            RebuildHighlightRows();
            RefreshGridLayout();

            // ★ 重新加载后重新检测唯一列重复（首次加载，弹 MessageBox）
            CheckUniqueColumnDuplicates(showMessageBox: true);
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

            // 重绑会重建所有行：先丢掉旧的显示行索引（重绑后按行号重建），
            // 但 _markedLineNumbers 保留，标记不会因为过滤/排序/编辑刷新而丢
            _selectedRows.Clear();
            UpdateSelectedRowsLabel();
            UpdateRowHeadersWidth();

            // 每次重新过滤前清空匹配高亮
            _filterMatchCells.Clear();
            _headerFilterMatchCells.Clear();
            _contentSearchMatchCells.Clear();
            _advFilterMatchCells.Clear();

            // 列可见性在刷新时统一重算，避免无冻结行路径沿用旧映射
            _visibleColumnIndices = GetVisibleColumnIndices();

            if (_frozenRows > 0 && _dataTable.Rows.Count > 0)
            {
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

                SetFrozenPanelVisible(true);
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

                // ★ 内容搜索：不过滤行，仅获取匹配单元格（所有行包括冻结行）
                _contentSearchMatchCells = GetContentMatchCells(contentFilter);

                // ★ 高级过滤：按表达式过滤行 + 匹配单元格
                var (visibleRowIndices, advMatchCells) = GetAdvVisibleRowIndices();
                // 如果保留所有行，则高级过滤不过滤行
                if (_filterKeepAllRows && !string.IsNullOrEmpty(_txtAdvancedFilter.Text.Trim()))
                {
                    visibleRowIndices = Enumerable.Range(_frozenRows, _fileData.Rows.Count - _frozenRows).ToList();
                }

                // ★ 排序只改变视图行顺序（先过滤后排序）
                visibleRowIndices = ApplySortToRowIndices(visibleRowIndices);

                // ★ 合并内容搜索 + 高级过滤的匹配单元格（原始行列索引）
                _advFilterMatchCells = advMatchCells;
                // ★ 转换匹配单元格：原始行列索引 → 显示行列索引（_dataGrid 内容区）
                _filterMatchCells.Clear();
                _headerFilterMatchCells.Clear();
                var mergedOrigMatch = MergeMatchCells(_contentSearchMatchCells, _advFilterMatchCells);
                foreach (var (origRow, origCols) in mergedOrigMatch)
                {
                    // 冻结行区域（origRow < _frozenRows）：直接存入 _headerFilterMatchCells
                    if (origRow < _frozenRows)
                    {
                        var displayCols = new HashSet<int>();
                        foreach (var origCol in origCols)
                        {
                            int displayCol = _visibleColumnIndices.IndexOf(origCol);
                            if (displayCol >= 0)
                                displayCols.Add(displayCol);
                        }
                        if (displayCols.Count > 0)
                            _headerFilterMatchCells[origRow] = displayCols;
                    }
                    else
                    {
                        // 内容行区域：转换为显示行索引存入 _filterMatchCells
                        int displayRow = visibleRowIndices.IndexOf(origRow);
                        if (displayRow < 0) continue;
                        var displayCols = new HashSet<int>();
                        foreach (var origCol in origCols)
                        {
                            int displayCol = _visibleColumnIndices.IndexOf(origCol);
                            if (displayCol >= 0)
                                displayCols.Add(displayCol);
                        }
                        if (displayCols.Count > 0)
                            _filterMatchCells[displayRow] = displayCols;
                    }
                }
                RefreshFilteredDisplay(visibleRowIndices);
            }
            else
            {
                SetFrozenPanelVisible(false);

                _contentSearchMatchCells = GetContentMatchCells(_txtContentFilter.Text.Trim());
                var (_, advMatchCellsNoFreeze) = GetAdvVisibleRowIndices();
                _advFilterMatchCells = advMatchCellsNoFreeze;
                _filterMatchCells.Clear();
                var mergedNoFreeze = MergeMatchCells(_contentSearchMatchCells, _advFilterMatchCells);

                if (_sortColumnIndex >= 0)
                {
                    // ★ 排序视图：只重排显示表，_dataTable 保持原始行序
                    var sortedRows = ApplySortToRowIndices(Enumerable.Range(0, _dataTable.Rows.Count).ToList());
                    foreach (var (origRow, origCols) in mergedNoFreeze)
                    {
                        int displayRow = sortedRows.IndexOf(origRow);
                        if (displayRow < 0) continue;
                        var displayCols = new HashSet<int>();
                        foreach (var origCol in origCols)
                        {
                            int displayCol = _visibleColumnIndices.IndexOf(origCol);
                            if (displayCol >= 0) displayCols.Add(displayCol);
                        }
                        if (displayCols.Count > 0)
                            _filterMatchCells[displayRow] = displayCols;
                    }
                    RefreshFilteredDisplay(sortedRows);
                    ApplyFrozenColumns();
                    RebuildSelectedRowIndices();
                    return;
                }

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
                // ★ 登记显示行 → 源行映射
                RebuildRowMap();

                // ★ 无冻结行模式下计算过滤匹配单元格（全表作为 DataSource，显示行 i 对应原始行 i）
                foreach (var (origRow, origCols) in mergedNoFreeze)
                {
                    if (origRow < 0 || origRow >= _dataGrid.Rows.Count) continue;
                    var displayCols = new HashSet<int>();
                    foreach (var origCol in origCols)
                    {
                        int displayCol = _visibleColumnIndices.IndexOf(origCol);
                        if (displayCol >= 0) displayCols.Add(displayCol);
                    }
                    if (displayCols.Count > 0)
                        _filterMatchCells[origRow] = displayCols;
                }
            }

            ApplyFrozenColumns();
            RebuildSelectedRowIndices();
        }

        /// <summary>登记全量行映射（显示行 i ↔ _fileData.Rows[i]）并重算行号列宽</summary>
        private void RebuildRowMap()
        {
            BuildRowTagMap(null);
            UpdateRowHeadersWidth();
        }

        /// <summary>
        /// 登记「显示行 → 源行记录」映射。
        /// 不再逐行写 DataGridViewRow.Tag：那一步会强制网格为每一行建出 Row/Cell 对象，
        /// 12 万行实测纯 UI 冻结 12.5 秒，而映射表本身只要零点几毫秒。
        /// </summary>
        private void BuildRowTagMap(IList<int>? visibleOrigIndices)
        {
            int count = visibleOrigIndices?.Count ?? _fileData.Rows.Count;
            var map = new TableRow[count];
            for (int i = 0; i < count; i++)
            {
                int origIdx = visibleOrigIndices?[i] ?? i;
                if (origIdx >= 0 && origIdx < _fileData.Rows.Count)
                    map[i] = _fileData.Rows[origIdx];
            }
            _rowTags = map;
        }

        /// <summary>取显示行对应的源行记录（越界或刷新中间态返回 null，由调用方兜底）</summary>
        private TableRow? RowTagAt(int displayRowIndex)
            => displayRowIndex >= 0 && displayRowIndex < _rowTags.Length ? _rowTags[displayRowIndex] : null;

        /// <summary>按原始行索引反查显示行，找不到返回 -1（只扫映射表，不碰网格行对象）</summary>
        private int FindDisplayRowByOrigIndex(int originalRowIndex)
        {
            for (int i = 0; i < _rowTags.Length; i++)
                if (_rowTags[i]?.OriginalRowIndex == originalRowIndex) return i;
            return -1;
        }

        /// <summary>按源文件物理行号反查显示行，找不到返回 -1</summary>
        private int FindDisplayRowByLineNumber(long lineNumber)
        {
            for (int i = 0; i < _rowTags.Length; i++)
                if (_rowTags[i]?.DisplayLineNumber == lineNumber) return i;
            return -1;
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
                    sb.AppendLine($"? 值 \"{g.Key}\" 重复 {g.Count()} 次：");
                    int showRows = Math.Min(5, g.Count());
                    foreach (var x in g.Take(showRows))
                    {
                        int fileRowNum = x.DataTableIndex < _fileData.Rows.Count
                            ? _fileData.Rows[x.DataTableIndex].DisplayLineNumber ?? x.DataTableIndex + 1
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

                AlertForm.Show(this, sb.ToString(), "唯一列重复警告", AlertType.Warning);
            }

            // ★ 状态栏：简短摘要 + 可点击（始终更新）
            _duplicateWarningText = totalDupCount <= 3
                ? $"? 唯一列\"{colDisplayName}\"有{totalDupCount}个重复值({totalRedundant}处冗余)"
                : $"? 唯一列\"{colDisplayName}\"有{totalDupCount}个重复值({totalRedundant}处冗余)…点击查看详情";

            ApplyDuplicateWarningStyle();
            return true;
        }

        /// <summary>应用重复警告样式到状态栏</summary>
        private void ApplyDuplicateWarningStyle()
        {
            if (_duplicateWarningText == null) return;
            _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 30, 30));
            _lblInfo.Text = _duplicateWarningText;
            _lblInfo.Cursor = Cursors.Hand;
            _lblInfo.Click += LblInfo_Click_ShowDuplicates;
            _pnlStatus.BackColor = Theme.CellTint(Color.FromArgb(255, 235, 235));
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
            _pnlStatus.BackColor = Theme.P.Header;
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

            // 行号以外的绑定表可能只是排序/过滤视图，必须用映射里的原始行索引
            int i = FindDisplayRowByOrigIndex(dataTableRowIndex);
            if (i < 0) return;
            _dataGrid.ClearSelection();
            _dataGrid.Rows[i].Selected = true;
            _dataGrid.FirstDisplayedScrollingRowIndex = Math.Max(0, i - 5);
            _dataGrid.CurrentCell = _dataGrid.Rows[i].Cells[Math.Min(1, _dataGrid.Columns.Count - 1)];
        }

        /// <summary>
        /// 按行号区显示的行号（源文件物理行号）定位到对应行。
        /// 找不到时区分"被过滤条件隐藏"和"确实没有这一行"，把原因写在状态栏。
        /// </summary>
        public bool NavigateToPhysicalLine(long lineNumber)
        {
            int hit = FindDisplayRowByLineNumber(lineNumber);
            if (hit >= 0)
            {
                RevealRow(_dataGrid, hit);
                return true;
            }

            // 冻结区（表头/描述行）行号 = 显示索引 + 1
            if (_frozenRows > 0 && lineNumber >= 1 && lineNumber <= _frozenRows
                && lineNumber - 1 < _dataGridHeader.Rows.Count)
            {
                RevealRow(_dataGridHeader, (int)lineNumber - 1);
                return true;
            }

            bool hiddenByFilter = _fileData.Rows.Any(r => r.DisplayLineNumber == lineNumber);
            _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 30, 30));
            _lblInfo.Text = hiddenByFilter
                ? $"行号 {lineNumber} 被当前过滤条件隐藏，清空高级过滤后再跳转"
                : $"没有行号为 {lineNumber} 的行（当前最大行号 {MaxDisplayLineNumber()}）";
            return false;
        }

        // ==================== 列选择 ====================
        public List<int> GetVisibleColumnIndices()
        {
            var visibleIndices = new List<int>();
            // 第0列是配置表的主键/唯一ID列，锁定可见（用于行识别与版本对齐）；
            // 行号不是数据列，画在 RowHeader 上，见 OnRowPostPaint
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
            if (_columnManagementDisabled) return;  // ★ 版本对比等场景，列管理被禁用
            if (_dataTable.Columns.Count == 0) return;
            EnsureFileEntriesLoaded();
            using var dialog = new ColumnManagerDialog(_dataTable, _columnCheckStates, _frozenRows, 0,
                _rootDir ?? "", _filePath, _fileEntries, _columnLinks, _headerRowIndex,
                _mappingRowIndex);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _columnCheckStates = dialog.GetCheckedStates();
                _columnLinks = dialog.GetColumnLinks();
                // ★ 保存前对模板引用的列清理冗余数据（仅保留 TemplateUid，由运行时实时加载）
                foreach (var link in _columnLinks)
                {
                    if (!string.IsNullOrEmpty(link.TemplateUid))
                    {
                        link.ValueExtractionEnabled = false;
                        link.ExtractSteps = new();
                        link.DisplayColumnHeaders = "";
                        link.FieldLinks = new();
                    }
                }
                CsvSettings.SetColumnLinksByDirectory(FrozenDir, _columnLinks);
                // ★ 保存后重新从模板填充运行时数据
                foreach (var link in _columnLinks)
                {
                    link.RefreshFromTemplate();
                }
                // ★ 后台异步构建关联 lookup
                if (_columnLinks.Count > 0)
                    FireAndForgetBuildLinkLookups(_columnLinks, _filePath);
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
            int dataColumns, checkedData;
            if (_externalTotalColumns.HasValue && _externalVisibleColumns.HasValue)
            {
                // ★ 版本对比场景：使用外部注入的完整列数信息（避免投影后 _columnCheckStates 长度不匹配）
                dataColumns = _externalTotalColumns.Value;
                checkedData = _externalVisibleColumns.Value;
            }
            else
            {
                // 正常场景：基于 _columnCheckStates 计算
                int checkedCount = 0;
                for (int i = 0; i < _columnCheckStates.Length; i++)
                    if (_columnCheckStates[i]) checkedCount++;
                dataColumns = Math.Max(0, _columnCheckStates.Length - 1);
                checkedData = Math.Max(0, checkedCount - 1);
            }
            string colmgrText = $"列管理({checkedData}/{dataColumns})";
            bool hasHidden = checkedData < dataColumns;
            if (_btnSelectColumns.Text != colmgrText)
            {
                // 按钮建好时按"列管理"量宽，加上列数后文字变长会被截断，这里跟着文本重新量宽
                _btnSelectColumns.Text = colmgrText;
                _btnSelectColumns.Width = ThemeButton.FitWidth(colmgrText);
            }
            // ThemeButton 自绘不认 ForeColor，有隐藏列时走 Danger 态把文字和描边染红
            if (_btnSelectColumns.Danger != hasHidden)
            {
                _btnSelectColumns.Danger = hasHidden;
                _btnSelectColumns.Invalidate();
            }
        }

        // ==================== 行过滤 ====================
        /// <summary>触发过滤（回车/按钮统一入口），后台执行不卡 UI，支持版本号丢弃过期结果</summary>
        private void TriggerFilter()
        {
            // ★ 表达式引用了不存在的列名 → 明确报错并保留当前视图。
            //   以前会静默退化成"全列模糊匹配"，策划看到的是错行却不报错。
            var unknownCols = FindUnknownFilterColumns(_txtAdvancedFilter.Text.Trim());
            if (unknownCols.Count > 0)
            {
                _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 30, 30));
                _lblInfo.Text = $"高级过滤里的列名不存在：{string.Join("、", unknownCols)}（可用偏左的\"列管理\"查看实际列名）";
                return;
            }

            var myVersion = Interlocked.Increment(ref _filterVersion);
            _ = RunFilterAsync(myVersion);
        }

        /// <summary>收集高级过滤表达式中引用了但不存在的列名（编译期一次性校验）</summary>
        private List<string> FindUnknownFilterColumns(string expr)
        {
            var missing = new List<string>();
            if (string.IsNullOrEmpty(expr)) return missing;
            CollectUnknownFilterColumns(CompileExpression(expr), missing);
            return missing;
        }

        private void CollectUnknownFilterColumns(FilterNode? node, List<string> missing)
        {
            switch (node)
            {
                case BinOpNode bin:
                    CollectUnknownFilterColumns(bin.Left, missing);
                    CollectUnknownFilterColumns(bin.Right, missing);
                    break;
                case CondNode cond when !cond.IsLineFilter
                    && !string.IsNullOrEmpty(cond.CompareOp) && !string.IsNullOrEmpty(cond.ColumnName):
                    if (ResolveColumnIndex(cond.ColumnName!) < 0
                        && !missing.Contains(cond.ColumnName!, StringComparer.OrdinalIgnoreCase))
                        missing.Add(cond.ColumnName!);
                    break;
            }
        }

        private async Task RunFilterAsync(int version)
        {
            if (version <= _pendingFilterVersion) return;
            _pendingFilterVersion = version;

            try
            {
                // ===== 阶段1：UI 线程捕获状态快照 =====
                var contentFilter = _txtContentFilter.Text.Trim();
                var advFilterText = _txtAdvancedFilter.Text.Trim();
                var filterKeepAll = _filterKeepAllRows;
                var frozenRows = _frozenRows;
                var visibleColIndices = _visibleColumnIndices.Count > 0
                    ? _visibleColumnIndices.ToList()
                    : GetVisibleColumnIndices();
                var dataTableRef = _dataTable;
                var fileDataRef = _fileData;
                var headerRowIndex = _headerRowIndex;
                var sortSourceColumn = _sortColumnIndex >= 0 ? GetSourceColumnIndex(_sortColumnIndex) : -1;
                var sortOrder = _sortOrder;

                // ===== 阶段2：后台线程执行耗时计算 =====
                var result = await Task.Run(() =>
                    BuildFilterResultInBackground(
                        contentFilter, advFilterText, filterKeepAll,
                        frozenRows, visibleColIndices, dataTableRef, fileDataRef,
                        headerRowIndex, sortSourceColumn, sortOrder, version));

                // ===== 阶段3：UI 线程检查版本 + 绑定 =====
                if (version != _filterVersion) return; // 已过期，丢弃

                ApplyFilterResultToUI(result);
            }
            finally
            {
                if (_pendingFilterVersion == version)
                    _pendingFilterVersion = 0;
            }
        }

        private void OnFilterKeepAllChanged(object? sender, EventArgs e)
        {
            _filterKeepAllRows = _chkFilterKeepAll.Checked;
            TriggerFilter();
        }

        private void ApplyFilters() => RefreshGridLayout();

        /// <summary>后台构建过滤结果（在后台线程执行）</summary>
        private FilterResult BuildFilterResultInBackground(
            string contentFilter, string advFilterText, bool filterKeepAll,
            int frozenRows, List<int> visibleColIndices,
            DataTable dataTable, FileTableData fileData,
            int headerRowIndex, int sortSourceColumn, SortOrder sortOrder, int version)
        {
            void CheckVersion()
            {
                if (version != _filterVersion)
                    throw new OperationCanceledException("过滤结果已过期");
            }

            CheckVersion();

            // 内容搜索匹配
            var contentMatches = GetContentMatchCellsBackground(
                fileData, visibleColIndices, contentFilter, version);
            CheckVersion();

            // 高级过滤
            var (visibleRowIndices, advMatches) = GetAdvVisibleRowIndicesBackground(
                fileData, advFilterText, filterKeepAll, frozenRows, version);
            CheckVersion();

            // ★ 排序只改变视图顺序（先过滤后排序）
            visibleRowIndices = SortRowIndices(dataTable, visibleRowIndices, sortSourceColumn, sortOrder);
            CheckVersion();

            // 构建 filteredTable
            var filteredTable = BuildFilteredTableBackground(
                dataTable, visibleRowIndices, visibleColIndices, version);
            CheckVersion();

            // 合并匹配单元格 → 转换显示行列索引
            var merged = MergeMatchCells(contentMatches, advMatches);
            var headerMatchCells = new Dictionary<int, HashSet<int>>();
            var dataMatchCells = new Dictionary<int, HashSet<int>>();

            foreach (var (origRow, origCols) in merged)
            {
                var displayCols = new HashSet<int>();
                foreach (var origCol in origCols)
                {
                    int displayCol = visibleColIndices.IndexOf(origCol);
                    if (displayCol >= 0) displayCols.Add(displayCol);
                }
                if (displayCols.Count == 0) continue;

                if (origRow < frozenRows)
                    headerMatchCells[origRow] = displayCols;
                else
                {
                    int displayRow = visibleRowIndices.IndexOf(origRow);
                    if (displayRow >= 0)
                        dataMatchCells[displayRow] = displayCols;
                }
            }

            return new FilterResult
            {
                FilteredTable = filteredTable,
                VisibleRowIndices = visibleRowIndices,
                ContentSearchMatchCells = contentMatches,
                AdvFilterMatchCells = advMatches,
                HeaderFilterMatchCells = headerMatchCells,
                DataFilterMatchCells = dataMatchCells
            };
        }

        /// <summary>后台内容搜索匹配（带版本检查）</summary>
        private Dictionary<int, HashSet<int>> GetContentMatchCellsBackground(
            FileTableData fileData, List<int> visibleColIndices,
            string contentFilter, int version)
        {
            if (string.IsNullOrEmpty(contentFilter))
                return new Dictionary<int, HashSet<int>>();

            string processedFilter = contentFilter
                .Replace("||", " | ").Replace("&&", " & ");
            var matchCells = new Dictionary<int, HashSet<int>>();

            for (int i = 0; i < fileData.Rows.Count; i++)
            {
                if (i % 500 == 0 && version != _filterVersion)
                    return matchCells;

                var cells = fileData.Rows[i].Cells;
                HashSet<int>? rowMatchCols = null;
                foreach (var colIdx in visibleColIndices)
                {
                    if (colIdx < cells.Count &&
                        MatchesFilter(cells[colIdx], processedFilter))
                    {
                        (rowMatchCols ??= new HashSet<int>()).Add(colIdx);
                    }
                }
                if (rowMatchCols != null && rowMatchCols.Count > 0)
                    matchCells[i] = rowMatchCols;
            }
            return matchCells;
        }

        /// <summary>后台高级过滤（带版本检查），从 _frozenRows 开始过滤内容行</summary>
        private (List<int> rowIndices, Dictionary<int, HashSet<int>> matchCells) GetAdvVisibleRowIndicesBackground(
            FileTableData fileData, string advFilterText, bool filterKeepAll, int frozenRows, int version)
        {
            var result = new List<int>(fileData.Rows.Count - frozenRows);
            var matchCells = new Dictionary<int, HashSet<int>>();

            FilterNode? compiledExpr = null;
            if (!string.IsNullOrEmpty(advFilterText))
                compiledExpr = CompileExpression(advFilterText);

            // 无高级过滤 → 返回所有内容行（从 frozenRows 开始）
            if (compiledExpr == null)
            {
                for (int i = frozenRows; i < fileData.Rows.Count; i++)
                    result.Add(i);
                return (result, matchCells);
            }

            for (int i = frozenRows; i < fileData.Rows.Count; i++)
            {
                if (i % 500 == 0 && version != _filterVersion)
                    return (result, matchCells); // 过期退出

                var row = fileData.Rows[i];
                bool advMatch = EvaluateCompiled(row, compiledExpr);
                if (!advMatch && filterKeepAll)
                {
                    // 保留所有行模式：不过滤行，只记录匹配的单元格
                    result.Add(i);
                    var advMatchCols = GetAdvFilterMatchCols(row.Cells, compiledExpr);
                    if (advMatchCols.Count > 0)
                        matchCells[i] = advMatchCols;
                }
                else if (advMatch)
                {
                    result.Add(i);
                    var advMatchCols = GetAdvFilterMatchCols(row.Cells, compiledExpr);
                    if (advMatchCols.Count > 0)
                        matchCells[i] = advMatchCols;
                }
            }

            // ★ 高级过滤无匹配时：
            //   - filterKeepAll == true（保留所有行模式）：回退显示所有内容行
            //   - filterKeepAll == false：result 保持空，冻结行正常显示，内容区为空
            if (result.Count == 0 && filterKeepAll)
            {
                for (int i = frozenRows; i < fileData.Rows.Count; i++)
                    result.Add(i);
            }

            return (result, matchCells);
        }

        /// <summary>后台构建 filteredTable（带版本检查）</summary>
        private DataTable BuildFilteredTableBackground(
            DataTable sourceTable, List<int> visibleRowIndices,
            List<int> visibleColIndices, int version)
        {
            var filteredTable = new DataTable();
            foreach (var colIndex in visibleColIndices)
            {
                var sourceCol = sourceTable.Columns[colIndex];
                filteredTable.Columns.Add(sourceCol.ColumnName, typeof(string));
            }

            for (int idx = 0; idx < visibleRowIndices.Count; idx++)
            {
                if (idx % 2000 == 0 && version != _filterVersion)
                    return filteredTable;

                var newRow = filteredTable.NewRow();
                int origRowIdx = visibleRowIndices[idx];
                for (int i = 0; i < visibleColIndices.Count; i++)
                {
                    int colIdx = visibleColIndices[i];
                    newRow[i] = colIdx < sourceTable.Columns.Count
                        ? (sourceTable.Rows[origRowIdx][colIdx] ?? "")
                        : "";
                }
                filteredTable.Rows.Add(newRow);
            }
            return filteredTable;
        }

        /// <summary>将后台过滤结果绑定到 UI（必须在 UI 线程调用）</summary>
        private void ApplyFilterResultToUI(FilterResult result)
        {
            // 行编辑模式：刷新前退出编辑
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

            // 清空旧匹配
            _filterMatchCells.Clear();
            _headerFilterMatchCells.Clear();
            _contentSearchMatchCells.Clear();
            _advFilterMatchCells.Clear();

            if (_frozenRows > 0)
            {
                // 冻结行场景：重建 _headerTable + 绑定两个 DataGridView
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

                SetFrozenPanelVisible(true);
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
                        if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && i < _visibleColumnIndices.Count)
                        {
                            var srcColIdx = _visibleColumnIndices[i];
                            var headerVal = _dataTable.Rows[_headerRowIndex][srcColIdx]?.ToString() ?? GetExcelColumnName(i);
                            dgvCol.HeaderText = headerVal;
                            SetColumnHeaderToolTip(dgvCol, i, headerVal);
                        }
                        else
                        {
                            dgvCol.HeaderText = GetExcelColumnName(i);
                        }
                        dgvCol.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                }

                _dataGrid.DataSource = result.FilteredTable;
                BuildRowTagMap(result.VisibleRowIndices);

                // 配置列宽/表头
                _dataGrid.SuspendLayout();
                for (int i = 0; i < _dataGrid.Columns.Count; i++)
                {
                    var colName = result.FilteredTable.Columns[i].ColumnName;
                    var dgvCol = _dataGrid.Columns[colName];
                    if (dgvCol != null)
                    {
                        dgvCol.DisplayIndex = i;
                        dgvCol.Width = 120;
                        dgvCol.SortMode = DataGridViewColumnSortMode.Programmatic;
                        if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && i < _visibleColumnIndices.Count)
                        {
                            var srcColIdx = _visibleColumnIndices[i];
                            var headerVal = _dataTable.Rows[_headerRowIndex][srcColIdx]?.ToString() ?? GetExcelColumnName(i);
                            dgvCol.HeaderText = headerVal;
                            SetColumnHeaderToolTip(dgvCol, i, headerVal);
                        }
                        else
                        {
                            dgvCol.HeaderText = GetExcelColumnName(i);
                        }
                    }
                }
                _dataGrid.ResumeLayout(false);
            }
            else
            {
                // 无冻结行场景
                SetFrozenPanelVisible(false);
                _dataGrid.DataSource = result.FilteredTable;

                // 登记行映射
                BuildRowTagMap(result.VisibleRowIndices);

                // 配置列宽/表头
                _dataGrid.SuspendLayout();
                for (int i = 0; i < _dataGrid.Columns.Count; i++)
                {
                    _dataGrid.Columns[i].Width = 120;
                    _dataGrid.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
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
                _dataGrid.ResumeLayout(false);
            }

            // 设置匹配单元格
            _contentSearchMatchCells = result.ContentSearchMatchCells;
            _advFilterMatchCells = result.AdvFilterMatchCells;
            _filterMatchCells = result.DataFilterMatchCells;
            _headerFilterMatchCells = result.HeaderFilterMatchCells;

            _lblCellPos.Text = $"显示 {_visibleColumnIndices.Count} 列, {result.VisibleRowIndices.Count} 行";
            _lblInfo.Text = "";

            // ★ 高级过滤无匹配提示
            if (result.VisibleRowIndices.Count == 0 && !string.IsNullOrEmpty(_txtAdvancedFilter.Text.Trim()))
            {
                _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 30, 30));
                _lblInfo.Text = $"高级过滤 \"{_txtAdvancedFilter.Text.Trim()}\" 未匹配到任何行";
            }
            else if (_duplicateWarningText != null)
            {
                _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 30, 30));
                _lblInfo.Text = _duplicateWarningText;
            }

            // 后台过滤是独立的重绑路径：冻结列要补套，标记要按行号映射到新的显示行索引
            ApplyFrozenColumns();
            RebuildSelectedRowIndices();
        }

        /// <summary>后台过滤结果数据容器</summary>
        private class FilterResult
        {
            public DataTable FilteredTable { get; set; } = new();
            public List<int> VisibleRowIndices { get; set; } = new();
            public Dictionary<int, HashSet<int>> ContentSearchMatchCells { get; set; } = new();
            public Dictionary<int, HashSet<int>> AdvFilterMatchCells { get; set; } = new();
            public Dictionary<int, HashSet<int>> HeaderFilterMatchCells { get; set; } = new();
            public Dictionary<int, HashSet<int>> DataFilterMatchCells { get; set; } = new();
        }

        /// <summary>内容搜索：搜索所有行（包括冻结行）的所有可见列，仅返回匹配的原始单元格位置，不过滤行</summary>
        private Dictionary<int, HashSet<int>> GetContentMatchCells(string contentFilter)
        {
            if (string.IsNullOrEmpty(contentFilter))
                return new Dictionary<int, HashSet<int>>();

            string processedContentFilter = contentFilter
                .Replace("||", " | ").Replace("&&", " & ");
            var matchCells = new Dictionary<int, HashSet<int>>();

            // 搜索所有行：从 0 到 _fileData.Rows.Count（包括冻结行）
            for (int i = 0; i < _fileData.Rows.Count; i++)
            {
                var cells = _fileData.Rows[i].Cells;
                HashSet<int>? rowMatchCols = null;
                foreach (var colIdx in _visibleColumnIndices)
                {
                    if (colIdx < cells.Count &&
                        MatchesFilter(cells[colIdx], processedContentFilter))
                    {
                        (rowMatchCols ??= new HashSet<int>()).Add(colIdx);
                    }
                }
                if (rowMatchCols != null && rowMatchCols.Count > 0)
                    matchCells[i] = rowMatchCols;
            }
            return matchCells;
        }

        /// <summary>高级过滤：仅根据高级过滤表达式返回可见行索引 + 匹配的原始单元格（不过滤内容搜索）</summary>
        private (List<int> rowIndices, Dictionary<int, HashSet<int>> matchCells) GetAdvVisibleRowIndices()
        {
            var result = new List<int>(_fileData.Rows.Count - _frozenRows);
            var matchCells = new Dictionary<int, HashSet<int>>();
            var advFilter = _txtAdvancedFilter.Text.Trim();

            FilterNode? compiledExpr = null;
            if (!string.IsNullOrEmpty(advFilter))
                compiledExpr = CompileExpression(advFilter);

            // 无高级过滤 → 返回所有内容行（从 _frozenRows 开始）
            if (compiledExpr == null)
            {
                for (int i = _frozenRows; i < _fileData.Rows.Count; i++)
                    result.Add(i);
                return (result, matchCells);
            }

            for (int i = _frozenRows; i < _fileData.Rows.Count; i++)
            {
                var row = _fileData.Rows[i];
                bool advMatch = EvaluateCompiled(row, compiledExpr);
                if (!advMatch) continue;

                result.Add(i);
                var advMatchCols = GetAdvFilterMatchCols(row.Cells, compiledExpr);
                if (advMatchCols.Count > 0)
                    matchCells[i] = advMatchCols;
            }

            // ★ 高级过滤无匹配时：
            //   - _filterKeepAllRows == true（保留所有行模式）：回退显示所有内容行
            //   - _filterKeepAllRows == false：result 保持空，冻结行正常显示，内容区为空
            if (result.Count == 0 && _filterKeepAllRows)
            {
                for (int i = _frozenRows; i < _fileData.Rows.Count; i++)
                    result.Add(i);
            }

            return (result, matchCells);
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
            public bool IsLineFilter;      // "@line" 虚拟列：比的是行号区显示的行号，不对应任何数据列
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
            // @line 是虚拟列（行号区显示的行号），不是数据列，不参与列名解析
            bool isLineFilter = op != "like" && columnName.Equals("@line", StringComparison.OrdinalIgnoreCase);
            return new CondNode
            {
                ColumnName = !isLineFilter && !string.IsNullOrEmpty(columnName) && char.IsLetter(columnName[0]) ? columnName : null,
                ColumnIndex = -1,  // 求值时再解析
                CompareOp = op,
                Value = value,
                IsLineFilter = isLineFilter,
                IsNumeric = op != "like" && double.TryParse(value, out var nv),
                NumericValue = op != "like" && double.TryParse(value, out nv) ? nv : 0
            };
        }

        /// <summary>对 AST 求值（零装箱，直接从 List&lt;string&gt; 读取）</summary>
        private bool EvaluateCompiled(TableRow row, FilterNode node)
        {
            switch (node)
            {
                case BinOpNode bin:
                    bool left = EvaluateCompiled(row, bin.Left!);
                    if (bin.Op == "or" && left) return true;   // OR 短路
                    if (bin.Op == "and" && !left) return false; // AND 短路
                    return EvaluateCompiled(row, bin.Right!);

                case CondNode cond:
                    return EvaluateConditionFast(row, cond);
            }
            return true;
        }

        private bool EvaluateConditionFast(TableRow row, CondNode cond)
        {
            var cells = row.Cells;

            if (cond.IsLineFilter)
            {
                // 行号口径和 RowHeader 完全一致；占位行（本侧不存在）没有行号，不参与 @line 匹配
                var line = row.DisplayLineNumber;
                return line.HasValue && CompareValue(line.Value.ToString(), cond.Value, cond.CompareOp);
            }

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
                // 列名找不到不会走到这里（TriggerFilter 已拦截）；兜底按不匹配处理，不再冒充全列模糊匹配
                if (colIndex < 0) return false;
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

        /// <summary>从高级过滤 AST 中提取匹配的原始列索引集合</summary>
        private HashSet<int> GetAdvFilterMatchCols(List<string> cells, FilterNode node)
        {
            var result = new HashSet<int>();
            CollectAdvMatchCols(cells, node, result);
            return result;
        }

        private void CollectAdvMatchCols(List<string> cells, FilterNode node, HashSet<int> result)
        {
            switch (node)
            {
                case BinOpNode bin:
                    CollectAdvMatchCols(cells, bin.Left!, result);
                    CollectAdvMatchCols(cells, bin.Right!, result);
                    break;

                case CondNode cond:
                    CollectCondMatchCols(cells, cond, result);
                    break;
            }
        }

        private void CollectCondMatchCols(List<string> cells, CondNode cond, HashSet<int> result)
        {
            // @line 匹配的是行号，不是数据列，没有单元格可高亮
            if (cond.IsLineFilter) return;

            if (string.IsNullOrEmpty(cond.CompareOp))
            {
                // 无操作符：全列模糊匹配
                foreach (var colIdx in _visibleColumnIndices)
                {
                    if (colIdx < cells.Count &&
                        cells[colIdx].IndexOf(cond.Value, StringComparison.OrdinalIgnoreCase) >= 0)
                        result.Add(colIdx);
                }
                return;
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
                // 列名找不到不会走到这里（TriggerFilter 已拦截）；兜底不高亮任何列
                if (colIndex < 0) return;
            }
            else
            {
                // 无列名但有操作符 → 全列模糊匹配
                foreach (var ci in _visibleColumnIndices)
                {
                    if (ci < cells.Count &&
                        cells[ci].IndexOf(cond.Value, StringComparison.OrdinalIgnoreCase) >= 0)
                        result.Add(ci);
                }
                return;
            }

            // ★ 有操作符时，必须检查该单元格的值是否真的匹配条件，否则会导致整列高亮
            if (colIndex >= 0 && colIndex < cells.Count)
            {
                string cellValue = cells[colIndex];
                if (CompareValue(cellValue, cond.Value, cond.CompareOp))
                    result.Add(colIndex);
            }
        }

        /// <summary>合并两个匹配单元格字典（key 为原始行索引，value 为原始列索引集合），内容搜索 + 高级过滤合并</summary>
        private static Dictionary<int, HashSet<int>> MergeMatchCells(
            Dictionary<int, HashSet<int>> dict1,
            Dictionary<int, HashSet<int>> dict2)
        {
            if (dict1.Count == 0) return dict2;
            if (dict2.Count == 0) return dict1;
            var merged = new Dictionary<int, HashSet<int>>(dict1);
            foreach (var (row, cols) in dict2)
            {
                if (merged.TryGetValue(row, out var existing))
                {
                    existing.UnionWith(cols);
                }
                else
                {
                    merged[row] = new HashSet<int>(cols);
                }
            }
            return merged;
        }

        /// <summary>
        /// 冻结行场景的异步加载版本（带进度条）。
        /// 把 filteredTable 构建移到后台线程，用 Progress&lt;T&gt; 报告进度（自动切 UI 线程更新进度条）。
        /// 构建完成后切回 UI 线程做绑定 + 行映射 + 列宽。
        /// 注意：其他场景（过滤/排序/列选择）仍用同步的 RefreshGridLayout，本方法只用于首次加载。
        /// </summary>
        private async Task RefreshGridLayoutWithProgressAsync()
        {
            if (_dataTable.Columns.Count == 0) return;
            if (_frozenRows <= 0 || _dataTable.Rows.Count == 0)
            {
                RefreshGridLayout();
                return;
            }

            // ==================== 阶段1：构建 _headerTable + 绑定 _dataGridHeader（UI 线程，很快） ====================
            UpdateLoadingProgress(5, "正在准备表头...");
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
            SetFrozenPanelVisible(true);
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
                    if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && i < _visibleColumnIndices.Count)
                    {
                        var srcColIdx = _visibleColumnIndices[i];
                        var headerVal = _dataTable.Rows[_headerRowIndex][srcColIdx]?.ToString() ?? GetExcelColumnName(i);
                        dgvCol.HeaderText = headerVal;
                        SetColumnHeaderToolTip(dgvCol, i, headerVal);
                    }
                    else
                    {
                        dgvCol.HeaderText = GetExcelColumnName(i);
                    }
                    dgvCol.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
            }
            await Task.Yield();  // 让 UI 线程有机会刷新一次

            // ==================== 阶段2：后台线程构建 filteredTable（Progress<T> 自动切 UI 线程更新进度条） ====================
            var contentFilter = _txtContentFilter.Text.Trim();

            // ★ 内容搜索：不过滤行，仅获取匹配单元格（所有行包括冻结行）
            _contentSearchMatchCells = GetContentMatchCells(contentFilter);

            // ★ 高级过滤：按表达式过滤行 + 匹配单元格
            var (visibleRowIndices, advMatchCells) = GetAdvVisibleRowIndices();
            if (_filterKeepAllRows && !string.IsNullOrEmpty(_txtAdvancedFilter.Text.Trim()))
            {
                visibleRowIndices = Enumerable.Range(_frozenRows, _fileData.Rows.Count - _frozenRows).ToList();
            }

            // ★ 合并内容搜索 + 高级过滤的匹配单元格
            _advFilterMatchCells = advMatchCells;
            _filterMatchCells.Clear();
            _headerFilterMatchCells.Clear();
            var mergedOrig = MergeMatchCells(_contentSearchMatchCells, _advFilterMatchCells);
            foreach (var (origRow, origCols) in mergedOrig)
            {
                if (origRow < _frozenRows)
                {
                    var displayCols = new HashSet<int>();
                    foreach (var origCol in origCols)
                    {
                        int displayCol = _visibleColumnIndices.IndexOf(origCol);
                        if (displayCol >= 0)
                            displayCols.Add(displayCol);
                    }
                    if (displayCols.Count > 0)
                        _headerFilterMatchCells[origRow] = displayCols;
                }
                else
                {
                    int displayRow = visibleRowIndices.IndexOf(origRow);
                    if (displayRow < 0) continue;
                    var displayCols = new HashSet<int>();
                    foreach (var origCol in origCols)
                    {
                        int displayCol = _visibleColumnIndices.IndexOf(origCol);
                        if (displayCol >= 0)
                            displayCols.Add(displayCol);
                    }
                    if (displayCols.Count > 0)
                        _filterMatchCells[displayRow] = displayCols;
                }
            }
            var progress = new Progress<(int percent, string text)>(p => UpdateLoadingProgress(p.percent, p.text));

            // ★ 捕获后台线程需要的只读引用（避免后台线程访问 UI 控件）
            var dataTableRef = _dataTable;
            var visibleColIndicesRef = _visibleColumnIndices;

            var filteredTable = await Task.Run(() => BuildFilteredTableInBackground(
                visibleRowIndices, visibleColIndicesRef, dataTableRef, progress));

            // ==================== 阶段3：UI 线程绑定 + 行映射 + 列宽 ====================
            // 绑定 12 万行实测只要零点二秒（DataGridViewRow 按需建），UI 线程不会明显停顿。
            UpdateLoadingProgress(85, "正在渲染表格（可能需要数秒，请稍候）...");

            // 行编辑模式：刷新前退出编辑
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

            _dataGrid.DataSource = filteredTable;

            // 显示行 → 源行映射
            BuildRowTagMap(visibleRowIndices);

            // 配置列宽/表头
            _dataGrid.SuspendLayout();
            for (int i = 0; i < _dataGrid.Columns.Count; i++)
            {
                var colName = filteredTable.Columns[i].ColumnName;
                var dgvCol = _dataGrid.Columns[colName];
                if (dgvCol != null)
                {
                    dgvCol.DisplayIndex = i;
                    dgvCol.Width = 120;
                    dgvCol.SortMode = DataGridViewColumnSortMode.Programmatic;
                    if (_headerRowIndex >= 0 && _headerRowIndex < _dataTable.Rows.Count && i < _visibleColumnIndices.Count)
                    {
                        var srcColIdx = _visibleColumnIndices[i];
                        var headerVal = _dataTable.Rows[_headerRowIndex][srcColIdx]?.ToString() ?? GetExcelColumnName(i);
                        dgvCol.HeaderText = headerVal;
                        SetColumnHeaderToolTip(dgvCol, i, headerVal);
                    }
                    else
                    {
                        dgvCol.HeaderText = GetExcelColumnName(i);
                    }
                }
            }
            _dataGrid.ResumeLayout(false);
            _lblCellPos.Text = $"显示 {_visibleColumnIndices.Count} 列, {visibleRowIndices.Count} 行";
            _lblInfo.Text = "";
            if (_duplicateWarningText != null)
            {
                _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 30, 30));
                _lblInfo.Text = _duplicateWarningText;
            }

            ApplyFrozenColumns();
            RebuildSelectedRowIndices();
            UpdateLoadingProgress(100, "完成");
        }

        /// <summary>
        /// 后台线程构建 filteredTable（纯数据操作，不触发 UI 事件）。
        /// 通过 Progress&lt;T&gt; 报告进度（自动切 UI 线程更新进度条）。
        /// 逻辑与 RefreshFilteredDisplay 中的构建部分完全一致，只是移到了后台线程执行。
        /// </summary>
        private DataTable BuildFilteredTableInBackground(
            List<int> visibleRowIndices,
            List<int> visibleColumnIndices,
            DataTable sourceTable,
            IProgress<(int percent, string text)> progress)
        {
            var filteredTable = new DataTable();
            foreach (var colIndex in visibleColumnIndices)
            {
                var sourceCol = sourceTable.Columns[colIndex];
                filteredTable.Columns.Add(sourceCol.ColumnName, typeof(string));  // 统一 string 避免类型冲突
            }

            int total = visibleRowIndices.Count;
            int batchSize = 2000;
            int added = 0;
            foreach (var origRowIdx in visibleRowIndices)
            {
                var newRow = filteredTable.NewRow();
                for (int i = 0; i < visibleColumnIndices.Count; i++)
                {
                    int colIdx = visibleColumnIndices[i];
                    newRow[i] = colIdx < sourceTable.Columns.Count
                        ? (sourceTable.Rows[origRowIdx][colIdx] ?? "")
                        : "";
                }
                filteredTable.Rows.Add(newRow);
                added++;
                if (added % batchSize == 0)
                {
                    int percent = 10 + (int)((added / (double)total) * 70);  // 10% ~ 80%
                    progress.Report((percent, $"正在加载行数据... {added}/{total}"));
                }
            }
            progress.Report((80, $"正在加载行数据... {added}/{total}"));
            return filteredTable;
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
            // ★ 行映射（DataSource 赋值后登记）
            BuildRowTagMap(visibleRowIndices);
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
                _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 30, 30));
                _lblInfo.Text = _duplicateWarningText;
            }

            // ★ 高级过滤无匹配提示（覆盖重复警告）
            if (visibleRowIndices.Count == 0 && !string.IsNullOrEmpty(_txtAdvancedFilter.Text.Trim()))
            {
                _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 30, 30));
                _lblInfo.Text = $"高级过滤 \"{_txtAdvancedFilter.Text.Trim()}\" 未匹配到任何行";
            }
        }

        // ==================== 排序 ====================
        private void OnColumnHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;  // 右键列头用于冻结列菜单，不触发排序
            if (e.ColumnIndex < 0 || e.ColumnIndex >= _dataGrid.Columns.Count) return;

            // ==================== 行编辑模式：排序前退出编辑 ====================
            if (_editingRowIndex >= 0)
                ExitRowEditMode(commit: true);

            if (_sortColumnIndex == e.ColumnIndex)
                _sortOrder = _sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            else { _sortColumnIndex = e.ColumnIndex; _sortOrder = SortOrder.Ascending; }

            // ★ 排序只作用于视图：不重排 _dataTable，也不重建 _fileData，
            //   行号区/物理行号始终跟随原始行，保存时写回的行序不会被排序影响。
            RefreshGridLayout();
            UpdateColumnHeaderSortGlyph();
        }

        /// <summary>显示列索引 → 源表列索引</summary>
        private int GetSourceColumnIndex(int displayColumnIndex)
        {
            if (_visibleColumnIndices.Count > 0 &&
                displayColumnIndex >= 0 && displayColumnIndex < _visibleColumnIndices.Count)
                return _visibleColumnIndices[displayColumnIndex];
            return displayColumnIndex;
        }

        /// <summary>对原始行索引列表应用当前排序（仅生成视图顺序，不修改数据）</summary>
        private List<int> ApplySortToRowIndices(List<int> rowIndices)
        {
            if (_sortColumnIndex < 0 || _sortOrder == SortOrder.None || rowIndices.Count == 0)
                return rowIndices;
            return SortRowIndices(_dataTable, rowIndices, GetSourceColumnIndex(_sortColumnIndex), _sortOrder);
        }

        /// <summary>按指定源列排序原始行索引（后台线程安全：只读传入的 DataTable）</summary>
        private static List<int> SortRowIndices(DataTable sourceTable, List<int> rowIndices, int sourceColumnIndex, SortOrder sortOrder)
        {
            if (sourceColumnIndex < 0 || sourceColumnIndex >= sourceTable.Columns.Count || rowIndices.Count == 0)
                return rowIndices;

            var comparer = new NaturalStringComparer(sortOrder);
            return rowIndices
                .OrderBy(i => i >= 0 && i < sourceTable.Rows.Count
                    ? sourceTable.Rows[i][sourceColumnIndex]?.ToString() ?? ""
                    : "", comparer)
                .ToList();
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

            // ★ 内容区点击任意单元格 = 点击行号：整行标记（橙红边框）。
            //   标记随行绘制，横向/纵向滚动后仍能看清当前行。
            if (dgv == _dataGrid)
            {
                bool ctrlHeld = (Control.ModifierKeys & Keys.Control) == Keys.Control;
                if (ctrlHeld) ToggleRowSelection(e.RowIndex);
                else ReplaceRowSelection(e.RowIndex);
            }
            else if (_markedLineNumbers.Count > 0)
            {
                // 冻结行区域点击：沿用旧行为，清除内容区标记
                ClearRowSelection();
            }

            int colIdx = e.ColumnIndex;
            string colName = colIdx >= 0 && colIdx < 26 ? ((char)('A' + colIdx)).ToString() : $"Col{colIdx}";
            // 状态栏行号与 RowHeader 同口径：源文件物理行号，跨行记录显示区间
            _lblCellPos.Text = $"行号 {FormatLineRange(GetRowOf(dgv, e.RowIndex))}  列 {colName}({colIdx})";
        }

        /// <summary>把整行标记替换为指定行（Ctrl 之外的单元格/行号点击）</summary>
        private void ReplaceRowSelection(int displayRow)
        {
            var line = GetMarkableLineNumber(displayRow);
            if (line == null) return;  // 占位行没有源文件行号，不参与标记
            if (_markedLineNumbers.Count == 1 && _markedLineNumbers.Contains(line.Value))
            {
                UpdateSelectedRowsLabel();
                return;
            }
            var oldSelected = _selectedRows.ToList();
            _markedLineNumbers.Clear();
            _markedLineNumbers.Add(line.Value);
            _selectedRows.Clear();
            _selectedRows.Add(displayRow);
            foreach (var r in oldSelected) InvalidateMarkedRow(r);
            InvalidateMarkedRow(displayRow);
            UpdateSelectedRowsLabel();
        }

        /// <summary>Ctrl+点击：追加/取消整行标记</summary>
        private void ToggleRowSelection(int displayRow)
        {
            var line = GetMarkableLineNumber(displayRow);
            if (line == null) return;
            if (_selectedRows.Contains(displayRow))
            {
                _markedLineNumbers.Remove(line.Value);
                _selectedRows.Remove(displayRow);
            }
            else
            {
                _markedLineNumbers.Add(line.Value);
                _selectedRows.Add(displayRow);
            }
            InvalidateMarkedRow(displayRow);
            UpdateSelectedRowsLabel();
        }

        /// <summary>取显示行可标记的源文件物理行号（占位行返回 null）</summary>
        private int? GetMarkableLineNumber(int displayRow)
        {
            if (displayRow < 0 || displayRow >= _dataGrid.Rows.Count) return null;
            return RowTagAt(displayRow)?.DisplayLineNumber;
        }

        /// <summary>重绑 DataGridView 后按 _markedLineNumbers 重算显示行索引。
        /// 被过滤隐藏的标记行保留在 _markedLineNumbers 中，过滤条件去掉后自动回到视图。</summary>
        private void RebuildSelectedRowIndices()
        {
            _selectedRows.Clear();
            if (_markedLineNumbers.Count > 0)
            {
                for (int i = 0; i < _rowTags.Length; i++)
                {
                    if (_rowTags[i]?.DisplayLineNumber is int line
                        && _markedLineNumbers.Contains(line))
                    {
                        _selectedRows.Add(i);
                    }
                }
            }
            UpdateSelectedRowsLabel();
        }

        /// <summary>重绘标记行，并多刷下一行以擦除边框溢出到下一行顶部的残留</summary>
        private void InvalidateMarkedRow(int displayRow)
        {
            if (displayRow < 0 || displayRow >= _dataGrid.Rows.Count) return;
            _dataGrid.InvalidateRow(displayRow);
            if (displayRow + 1 < _dataGrid.Rows.Count) _dataGrid.InvalidateRow(displayRow + 1);
        }

        /// <summary>RefreshGridLayout 重建行后恢复滚动位置与当前单元格（行数不变时按显示索引恢复）</summary>
        private void RestoreScrollAndCurrentCell(int scrollRow, int currentRow, int currentCol)
        {
            if (scrollRow >= 0 && scrollRow < _dataGrid.Rows.Count)
                _dataGrid.FirstDisplayedScrollingRowIndex = scrollRow;
            if (currentRow >= 0 && currentRow < _dataGrid.Rows.Count
                && currentCol >= 0 && currentCol < _dataGrid.Rows[currentRow].Cells.Count)
                _dataGrid.CurrentCell = _dataGrid.Rows[currentRow].Cells[currentCol];
        }

        private void OnRowHeaderMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            // 行号标记只对应内容区行；冻结行区域的行号不支持标记
            if (!ReferenceEquals(sender, _dataGrid)) return;
            if (e.Button != MouseButtons.Left || e.RowIndex < 0) return;
            bool ctrlHeld = (Control.ModifierKeys & Keys.Control) == Keys.Control;
            if (ctrlHeld) ToggleRowSelection(e.RowIndex);
            else ReplaceRowSelection(e.RowIndex);
        }

        private void ClearRowSelection()
        {
            if (_markedLineNumbers.Count == 0 && _selectedRows.Count == 0) return;
            var oldSelected = _selectedRows.ToList();
            _markedLineNumbers.Clear();
            _selectedRows.Clear();
            foreach (var r in oldSelected) InvalidateMarkedRow(r);
            UpdateSelectedRowsLabel();
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
                                int hitRow = FindDisplayRowByOrigIndex(originalRowIndex);
                                if (hitRow >= 0 && displayColIndex < _dataGrid.Rows[hitRow].Cells.Count)
                                    _dataGrid.Rows[hitRow].Cells[displayColIndex].Value = newValue;
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
                                        int hitRow = FindDisplayRowByOrigIndex(originalRowIndex);
                                        if (hitRow >= 0 && displayColIndex < _dataGrid.Rows[hitRow].Cells.Count)
                                            _dataGrid.Rows[hitRow].Cells[displayColIndex].Value = newValue;
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
            // 双击进入行编辑模式，定位到双击的列
            if (!_isEmbeddedMode && !_isEditLocked)
                EnterRowEditMode(e.RowIndex, e.ColumnIndex);
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
                EnterRowEditMode(e.RowIndex, e.ColumnIndex);
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
                            // 如果 CurrentCell 已被用户（鼠标点击）移到了其他位置，不再强制跳转
                            if (_dataGrid.CurrentCell != null
                                && (_dataGrid.CurrentCell.RowIndex != nextRow
                                    || _dataGrid.CurrentCell.ColumnIndex != nextCol))
                            {
                                // 用户已主动选择了其他单元格，只尝试进入编辑模式
                                if (_dataGrid.CurrentCell.RowIndex == nextRow)
                                    _dataGrid.BeginEdit(true);
                                return;
                            }

                            // CurrentCell 在预期位置（Tab/Enter 导航），进入编辑模式
                            if (_dataGrid.CurrentCell != null
                                && _dataGrid.CurrentCell.RowIndex == nextRow
                                && _dataGrid.CurrentCell.ColumnIndex == nextCol)
                            {
                                _dataGrid.BeginEdit(true);
                            }
                            else
                            {
                                // 手动跳转到下一列
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

        /// <summary>行编辑模式中单元格值变更事件：记录已修改的单元格并刷新背景色（按行分组存储，Grid 重建后仍有效）</summary>
        private void OnCellValueChangedInEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (_editingRowIndex < 0) return;
            if (e.RowIndex != _editingRowIndex) return;
            if (e.ColumnIndex < 0 || e.ColumnIndex >= (_dataGrid.Rows[e.RowIndex]?.Cells.Count ?? 0)) return;

            // 检查是否与原始值不同
            if (_editingRowOriginalValues != null && e.ColumnIndex < _editingRowOriginalValues.Length)
            {
                var currentValue = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                var originalValue = _editingRowOriginalValues[e.ColumnIndex];

                // 转换为原始行列索引
                var (origRow, origCol) = GetOriginalIndices(e.RowIndex, e.ColumnIndex);

                // 按行分组存储
                if (currentValue != originalValue)
                {
                    if (!_modifiedCellsInEdit.TryGetValue(origRow, out var cols))
                    {
                        cols = new HashSet<int>();
                        _modifiedCellsInEdit[origRow] = cols;
                    }
                    cols.Add(origCol);
                }
                else
                {
                    // 改回原始值，从该行的已修改集合中移除
                    if (_modifiedCellsInEdit.TryGetValue(origRow, out var cols))
                    {
                        cols.Remove(origCol);
                        if (cols.Count == 0)
                            _modifiedCellsInEdit.Remove(origRow);
                    }
                }

                // 刷新该单元格背景色
                _dataGrid.InvalidateCell(e.ColumnIndex, e.RowIndex);
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
                // ★ 从行标记取原始行号
                int originalRowIndex = displayRowIndex;
                var tagRow = RowTagAt(displayRowIndex);
                if (tagRow != null)
                    originalRowIndex = tagRow.OriginalRowIndex;
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

        /// <returns>true=保存成功, false=保存失败</returns>
        private bool SaveCsvFile()
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
                string friendlyMsg;
                if (ex is IOException || ex is UnauthorizedAccessException)
                {
                    friendlyMsg = "文件被其他程序占用（如Excel已打开此文件），无法保存。\n" +
                                  "请关闭其他程序后重试，或另存到其他位置。";
                }
                else
                {
                    friendlyMsg = $"保存失败：{ex.Message}";
                }
                AlertForm.Show(this, friendlyMsg, "保存失败", AlertType.Warning);

                // 标记脏数据并启动重试定时器（设置 _hasUnsavedChanges、状态栏提示、启动定时器）
                MarkDirtyAndStartSaveTimer();

                return false;
            }

            // ★ 保存后重新检测唯一列重复
            CheckUniqueColumnDuplicates();
            // ★ 本路径直接写盘，需手动让缓存失效并通知同缓存页签重算关联
            _cache.InvalidateFile(_filePath);
            _cache.NotifyFileSaved(_filePath);
            return true;
        }

        /// <returns>true=保存成功, false=保存失败</returns>
        private bool SaveExcelFile()
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
                string friendlyMsg;
                if (ex is IOException || ex is UnauthorizedAccessException)
                {
                    friendlyMsg = "文件被其他程序占用（如Excel已打开此文件），无法保存。\n" +
                                  "请关闭其他程序后重试。\n" +
                                  "备份文件已保留在：" + backupPath;
                }
                else
                {
                    friendlyMsg = $"保存失败：{ex.Message}";
                }
                AlertForm.Show(this, friendlyMsg, "保存失败", AlertType.Warning);

                // 标记脏数据并启动重试定时器（设置 _hasUnsavedChanges、状态栏提示、启动定时器）
                MarkDirtyAndStartSaveTimer();

                return false;
            }

            // ★ 保存后重新检测唯一列重复
            CheckUniqueColumnDuplicates();
            // ★ 本路径直接写盘，需手动让缓存失效并通知同缓存页签重算关联
            _cache.InvalidateFile(_filePath);
            _cache.NotifyFileSaved(_filePath);
            return true;
        }

        private Font? _boldCellStyleFont;
        private Font? _boldCellStyleFontBase;

        /// <summary>
        /// 过滤命中单元格用的加粗字体。以前每命中一格、每重绘一次就 new 一个 Font，
        /// 大表滚动一次的能量就是几千个 GDI 句柄；这里按基准字体缓存一份。
        /// 切皮肤时 Theme 会换一批 Font 对象，基准引用变了就重建（旧句柄不回收，和 Theme 里的字体同一处理，
        /// 因为可能仍被现存样式引用）。
        /// </summary>
        private Font BoldCellStyleFont(Font baseFont)
        {
            if (!ReferenceEquals(_boldCellStyleFontBase, baseFont))
            {
                _boldCellStyleFontBase = baseFont;
                _boldCellStyleFont = new Font(baseFont, FontStyle.Bold);
            }
            return _boldCellStyleFont;
        }

        /// <summary>冻结行区域（_dataGridHeader）的 CellFormatting：过滤匹配单元格红色加粗高亮</summary>
        private void OnHeaderCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (_headerFilterMatchCells.Count > 0)
            {
                if (_headerFilterMatchCells.TryGetValue(e.RowIndex, out var matchCols) && matchCols.Contains(e.ColumnIndex))
                {
                    e.CellStyle.Font = BoldCellStyleFont(e.CellStyle.Font ?? _dataGridHeader.Font);
                    e.CellStyle.ForeColor = Color.Red;
                    e.CellStyle.SelectionForeColor = Color.Red;
                }
            }
        }

        private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (sender == _dataGridHeader) return;
            if (e.RowIndex >= 0 && e.RowIndex < _dataGrid.Rows.Count)
                _dataGrid.Rows[e.RowIndex].Height = Theme.P.RowHeight;
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // ==================== 标记行：整行铺一层柔和选中底（密度由皮肤决定） ====================
                // 放在语义色之前：改了值/重复行这类信息比"被标记"更重要，后面会覆盖它
                if (_selectedRows.Contains(e.RowIndex))
                {
                    e.CellStyle.BackColor = Theme.P.Selection;
                    e.CellStyle.SelectionBackColor = Theme.Mix(Theme.P.Selection, Theme.P.Text, 0.08f);
                    e.CellStyle.SelectionForeColor = Theme.P.Text;
                }

                // ==================== 行编辑模式：编辑行浅黄色背景 + 修改的单元格红色背景 ====================
                if (_editingRowIndex >= 0 && e.RowIndex == _editingRowIndex)
                {
                    // 判断该单元格是否被修改过
                    var (modOrigRow, modOrigCol) = GetOriginalIndices(e.RowIndex, e.ColumnIndex);
                    if (_modifiedCellsInEdit.TryGetValue(modOrigRow, out var modCols) && modCols.Contains(modOrigCol))
                    {
                        e.CellStyle.BackColor = Theme.CellTint(Color.FromArgb(255, 180, 180));  // 浅红色（已修改单元格）
                        e.CellStyle.SelectionBackColor = Theme.CellTint(Color.FromArgb(255, 140, 140));
                    }
                    else
                    {
                        e.CellStyle.BackColor = Theme.CellTint(Color.FromArgb(255, 253, 210));  // 浅黄色（与未保存行颜色一致）
                        e.CellStyle.SelectionBackColor = Theme.CellTint(Color.FromArgb(255, 250, 180));
                    }
                }
                else
                {
                    // ==================== 未保存行：淡橙色背景 ====================
                    bool isUnsavedRow = false;
                    if (_unsavedRowIndices.Count > 0)
                    {
                        var tagRow = RowTagAt(e.RowIndex);
                        if (tagRow != null && _unsavedRowIndices.Contains(tagRow.OriginalRowIndex))
                        {
                            isUnsavedRow = true;
                            e.CellStyle.BackColor = Theme.CellTint(Color.FromArgb(255, 253, 210));  // 浅黄色（与编辑行一致）
                            e.CellStyle.SelectionBackColor = Theme.CellTint(Color.FromArgb(255, 250, 180));
                            // 检查该单元格是否被修改过（使用原始行列索引匹配）
                            if (_modifiedCellsInEdit.Count > 0)
                            {
                                var (modOrigRow2, modOrigCol2) = GetOriginalIndices(e.RowIndex, e.ColumnIndex);
                                if (_modifiedCellsInEdit.TryGetValue(modOrigRow2, out var modCols2) && modCols2.Contains(modOrigCol2))
                                {
                                    e.CellStyle.BackColor = Theme.CellTint(Color.FromArgb(255, 180, 180));
                                    e.CellStyle.SelectionBackColor = Theme.CellTint(Color.FromArgb(255, 140, 140));
                                }
                            }
                        }
                    }

                    if (!isUnsavedRow)
                    {
                        // ★ 列关联颜色检查
                        ColumnLinkChecker.CheckCellColor(_cache, e, _visibleColumnIndices, _dataGrid, _filePath, FrozenDir, _columnLinks);
                        if (_highlightsReady && _highlightRows.Contains(e.RowIndex))
                        {
                            e.CellStyle.BackColor = Theme.CellTint(Color.FromArgb(255, 255, 200));
                            e.CellStyle.SelectionBackColor = Theme.CellTint(Color.FromArgb(255, 240, 150));
                        }
                    }
                }
            }
            // ==================== 过滤匹配单元格高亮（红色加粗字体） ====================
            if (_filterMatchCells.Count > 0 && e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (_filterMatchCells.TryGetValue(e.RowIndex, out var matchCols) && matchCols.Contains(e.ColumnIndex))
                {
                    e.CellStyle.Font = BoldCellStyleFont(e.CellStyle.Font ?? _dataGrid.Font);
                    e.CellStyle.ForeColor = Color.Red;
                    e.CellStyle.SelectionForeColor = Color.Red;
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
            // 内容区列头右键：无冻结行时列头由 _dataGrid 自己显示，冻结列入口要和固定行网格一致
            if (e.Button == MouseButtons.Right && e.RowIndex == -1)
            {
                if (e.ColumnIndex >= 0) ShowColumnFreezeMenu(e.ColumnIndex);
                return;
            }
            if (e.Button != MouseButtons.Right || e.RowIndex < 0) return;

            // ★ 行号区（RowHeader）右键：行号操作不受 ContextMenuMode 限制，阅读器/版本对比/冲突解决三处一致
            if (e.ColumnIndex < 0)
            {
                ShowRowNumberMenu(e.RowIndex);
                return;
            }

            if (_contextMenuMode == ContextMenuMode.None) return;

            // ★ Conflict 模式：转发给外部事件处理器（如冲突解决的右键菜单）
            if (_contextMenuMode == ContextMenuMode.Conflict)
            {
                OnExternalCellMouseDown?.Invoke(sender, e);
                return;
            }

            int adjCol = e.ColumnIndex;
            // ★ 不再需要 __RowNum__ 偏移
            int origCol = adjCol;
            if (adjCol >= 0 && adjCol < _visibleColumnIndices.Count)
                origCol = _visibleColumnIndices[adjCol];

            string currentRel = GetCurrentRelativePath();
            var link = _columnLinks.FirstOrDefault(l =>
                LinkSourceMatcher.Matches(l.SourceFile, currentRel, Path.GetFileName(_filePath)) &&
                l.SourceColumnIndex == origCol);

            var cellValue = _dataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
            var ctx = new ContextMenuStrip();
            ctx.Font = Theme.Body;

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
            if (_isEmbeddedMode || _isEditLocked) { editRowItem.Enabled = false; }
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

            ctx.Items.Add(new ToolStripSeparator());

            var linkItem = new ToolStripMenuItem("查看关联表");
            if (link == null) { linkItem.Enabled = false; }
            else
            {
                bool isValid = !string.IsNullOrEmpty(cellValue) && ColumnLinkChecker.IsLinkValueValid(_cache, _filePath, FrozenDir, origCol, cellValue, _columnLinks);
                linkItem.Enabled = isValid;
                linkItem.Click += (s, args) =>
                {
                    var firstLink = link.FieldLinks?.FirstOrDefault(f => !string.IsNullOrEmpty(f.TargetFile));
                    string targetFile = firstLink?.TargetFile ?? "";
                    string targetFullPath = Path.Combine(FrozenDir, targetFile);
                    if (!File.Exists(targetFullPath))
                    {
                        AlertForm.Show(this, "关联文件不存在: " + targetFile, "错误", AlertType.Error);
                        return;
                    }
                    var remarks = CsvSettings.GetFileRemarksByDirectory(FrozenDir);
                    var displayName = Path.GetFileName(targetFullPath);
                    if (remarks.TryGetValue(targetFile, out var remark) && !string.IsNullOrEmpty(remark))
                        displayName = remark + " (" + Path.GetFileName(targetFullPath) + ")";
                    var linkOpts = new TableBrowserOptions(targetFullPath)
                    {
                        RootDir = FrozenDir,
                        FrozenRows = _frozenRows,
                        HeaderRowIndex = _headerRowIndex,
                        MappingRowIndex = _mappingRowIndex,
                        Cache = _cache,   // ★ 关联弹出的独立窗口复用本控件的缓存，不另存一份表数据
                    };
                    var linkOpen = new LinkOpenOptions
                    {
                        HighlightColumnIndex = firstLink?.TargetColumnIndex ?? -1,
                        HighlightValue = cellValue,
                        HighlightColumnName = firstLink?.TargetColumnName ?? "",
                        LinkConfig = link,
                        DisplayName = displayName,
                        IsLinkedFromCell = true,
                    };
                    var standaloneForm = new CsvReaderStandaloneForm(linkOpts, linkOpen);
                    standaloneForm.Show(this);
                };
            }
            ctx.Items.Add(linkItem);

            // "编辑关联表" - 在父窗体标签页中打开
            var editLinkItem = new ToolStripMenuItem("编辑关联表");
            if (link == null) { editLinkItem.Enabled = false; }
            else
            {
                editLinkItem.Enabled = linkItem.Enabled;
                editLinkItem.Click += (s, args) =>
                {
                    var firstLink = link.FieldLinks?.FirstOrDefault(f => !string.IsNullOrEmpty(f.TargetFile));
                    string targetFile = firstLink?.TargetFile ?? "";
                    string targetFullPath = Path.Combine(FrozenDir, targetFile);
                    if (!File.Exists(targetFullPath))
                    {
                        AlertForm.Show(this, "关联文件不存在: " + targetFile, "错误", AlertType.Error);
                        return;
                    }
                    EditLinkTableRequested?.Invoke(this, new EditLinkTableEventArgs
                    {
                        TargetFilePath = targetFullPath,
                        DisplayName = Path.GetFileName(targetFullPath),
                    });
                };
            }
            ctx.Items.Add(editLinkItem);

            ctx.Items.Add(new ToolStripSeparator());

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
                        _linkToolTip.CreatePinnedWindow(content.title, content.lines, content.rawText, screenPos, content.tableHeaders);
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

            // 清除标记：和行号右键菜单保持同一入口
            if (_markedLineNumbers.Count > 0)
            {
                var clearMarksItem = new ToolStripMenuItem($"清除标记（{_markedLineNumbers.Count} 行）");
                clearMarksItem.Click += (s, args) => ClearRowSelection();
                ctx.Items.Add(clearMarksItem);
            }

            // 冻结列（冻结行数=0 时没有列头可用，这里同样提供入口）
            AppendFrozenColumnMenuItems(ctx, e.ColumnIndex, addSeparator: true);

            // ★ 显示右键菜单前主动隐藏 ToolTip，防止 ToolTip 遮挡右键菜单
            _linkToolTip?.Hide();
            ctx.Show(_dataGrid, _dataGrid.PointToClient(Cursor.Position));
        }

        private string GetCurrentRelativePath()
        {
            if (!string.IsNullOrEmpty(FrozenDir) && _filePath.StartsWith(FrozenDir, StringComparison.OrdinalIgnoreCase))
                return _filePath.Substring(FrozenDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(_filePath);
        }

        /// <summary>后台异步构建关联 lookup，不阻塞 UI（fire-and-forget，内部防重入）</summary>
        private void FireAndForgetBuildLinkLookups(List<ColumnLinkConfig> columnLinks, string filePath)
        {
            var rootDir = FrozenDir;
            _ = Task.Run(async () =>
            {
                try
                {
                    await _cache.BuildLinkLookupsAsync(rootDir, columnLinks, filePath);
                    // 构建完成后，通知 UI 刷新（让 tooltip 能获取到最新数据）
                    if (this.IsHandleCreated)
                        this.BeginInvoke(() =>
                        {
                            _dataGrid?.Invalidate();
                            _dataGridHeader?.Invalidate();
                        });
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"BuildLinkLookupsAsync failed: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 缓存通知：某个文件已保存。若保存的是本文件，或本文件的关联配置引用了它，
        /// 则重建关联 lookup 并重绘（否则红色"未找到"提示会停留在旧数据上）。
        /// </summary>
        private void OnCacheFileSaved(string savedPath)
        {
            if (IsDisposed || Disposing) return;

            if (InvokeRequired)
            {
                try { BeginInvoke(new Action<string>(OnCacheFileSaved), savedPath); }
                catch (InvalidOperationException) { }
                return;
            }

            if (string.IsNullOrEmpty(_rootDir) || string.IsNullOrEmpty(_filePath)) return;
            if (_columnLinks.Count == 0 || string.IsNullOrEmpty(savedPath)) return;

            bool ownSaved = string.Equals(savedPath, _filePath, StringComparison.OrdinalIgnoreCase);
            bool affected = ownSaved;
            if (!affected)
            {
                var savedName = Path.GetFileName(savedPath);
                var ownName = Path.GetFileName(_filePath);
                var currentRel = GetCurrentRelativePath();
                affected = _columnLinks.Any(l =>
                    LinkSourceMatcher.Matches(l.SourceFile, currentRel, ownName) &&
                    l.FieldLinks != null &&
                    l.FieldLinks.Any(fl => !string.IsNullOrEmpty(fl.TargetFile) &&
                        string.Equals(Path.GetFileName(fl.TargetFile), savedName,
                            StringComparison.OrdinalIgnoreCase)));
            }
            if (!affected) return;

            FireAndForgetBuildLinkLookups(_columnLinks, _filePath);
            _dataGrid.Invalidate();
            _dataGridHeader.Invalidate();
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
            var result = AlertForm.ShowConfirm(this,
                "Excel文件如果包含公式、宏定义等，编辑后保存可能导致这些内容丢失。" + GlobalConst.NewLine + "确认要解锁编辑吗？",
                "解锁编辑确认", AlertType.Warning);

            if (result == DialogResult.Yes)
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
                _lblLock.Text = "已锁定";
                _lblLock.ForeColor = _colorLightTextSecondary;
            }
            else
            {
                _lblLock.Text = "可编辑";
                _lblLock.ForeColor = Theme.P.Mark; // 橙色，表示可编辑
            }
        }

        // ==================== 行号绘制 ====================

        /// <summary>
        /// 行号显示的唯一取值出口。行身份已知时用 TableRow.DisplayLineNumber（物理行号优先）；
        /// 映射登记缺失（刷新中间态）时退回显示位置推算，保证不画空。
        /// </summary>
        private static int? ResolveRowNumber(TableRow? tagRow, int displayRowIndex, int frozenRows)
            => tagRow?.DisplayLineNumber ?? displayRowIndex + 1 + frozenRows;

        /// <summary>取数据区某行的行号（主表格用）</summary>
        private int? GetDisplayLineNumber(int rowIndex)
            => ResolveRowNumber(RowTagAt(rowIndex), rowIndex, _frozenRows);

        /// <summary>取冻结区某行的行号（表头表格不在行映射里，按 _fileData 索引取）</summary>
        private int? GetFrozenDisplayLineNumber(int rowIndex)
        {
            TableRow? tagRow = rowIndex >= 0 && rowIndex < _fileData.Rows.Count
                ? _fileData.Rows[rowIndex]
                : null;
            return ResolveRowNumber(tagRow, rowIndex, 0);
        }

        private void OnRowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            // 标记行的整行底色由 OnCellFormatting 铺，这里只补左侧强调条：
            // 原先画 1px 橙红方框要跟着横向滚动算右边界，宽表下既慢又容易错位。
            if (_selectedRows.Contains(e.RowIndex) && Theme.P.LeftAccent)
            {
                using var bar = new SolidBrush(_colorRowSelectedBorder);
                e.Graphics.FillRectangle(bar, e.RowBounds.Left, e.RowBounds.Top, 3, e.RowBounds.Height);
            }
            // ★ 行号取值单点化：优先源文件物理行号
            var lineNumber = GetDisplayLineNumber(e.RowIndex);
            var rowHeaderBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, _dataGrid.RowHeadersWidth, e.RowBounds.Height);
            using (var brush = new SolidBrush(_colorRowNumberText))
            {
                var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                e.Graphics.DrawString(FormatRowNumber(lineNumber), _rowNumberFont, brush, rowHeaderBounds, format);
            }
        }

        /// <summary>行号的显示文本：null（占位行/无对应源行）画破折号</summary>
        private static string FormatRowNumber(int? lineNumber) => lineNumber?.ToString() ?? "—";

        /// <summary>取某个显示行对应的数据行（冻结区按 _fileData 索引，主体区按行映射）</summary>
        private TableRow GetRowOf(DataGridView dgv, int rowIndex)
        {
            if (dgv == _dataGridHeader && rowIndex >= 0 && rowIndex < _fileData.Rows.Count)
                return _fileData.Rows[rowIndex];
            if (dgv == _dataGrid && RowTagAt(rowIndex) is TableRow tagRow)
                return tagRow;
            // 刷新中间态映射还没登记：退回记录序号，口径和 RowHeader 兜底一致
            return new TableRow { OriginalRowIndex = rowIndex + (dgv == _dataGrid ? _frozenRows : 0) };
        }

        /// <summary>源文件行号区间文本：跨行记录显示 L12–L14，占位行显示 —</summary>
        private static string FormatLineRange(TableRow row)
        {
            var start = row.DisplayLineNumber;
            if (start == null) return "—";
            return row.PhysicalEndLine > start ? $"L{start}–L{row.PhysicalEndLine}" : $"L{start}";
        }

        /// <summary>行号区右键菜单：复制行号 / 只看该行号 / 跳转到行号 / 清除标记</summary>
        private void ShowRowNumberMenu(int displayRow)
        {
            var line = GetRowOf(_dataGrid, displayRow).DisplayLineNumber;
            var ctx = new ContextMenuStrip();
            ctx.Font = Theme.Body;

            if (line.HasValue)
            {
                string lineText = line.Value.ToString();
                var copyItem = new ToolStripMenuItem($"复制行号 L{lineText}");
                copyItem.Click += (s, args) => Clipboard.SetText(lineText);
                ctx.Items.Add(copyItem);

                var filterItem = new ToolStripMenuItem($"只看该行号（@line={lineText}）");
                filterItem.Click += (s, args) =>
                {
                    _txtAdvancedFilter.Text = $"@line={lineText}";
                    TriggerFilter();
                };
                ctx.Items.Add(filterItem);
            }
            else
            {
                // 冲突解决/版本对比里本侧不存在的行没有源行号
                ctx.Items.Add(new ToolStripMenuItem("该行没有源文件行号") { Enabled = false });
            }

            var jumpItem = new ToolStripMenuItem("跳转到行号…");
            jumpItem.Click += (s, args) =>
            {
                using var dlg = new LineNumberJumpDialog(MaxDisplayLineNumber());
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    NavigateToPhysicalLine(dlg.LineNumber);
            };
            ctx.Items.Add(jumpItem);

            if (_markedLineNumbers.Count > 0)
            {
                var clearItem = new ToolStripMenuItem($"清除标记（{_markedLineNumbers.Count} 行）");
                clearItem.Click += (s, args) => ClearRowSelection();
                ctx.Items.Add(clearItem);
            }

            ctx.Show(Cursor.Position);
        }

        /// <summary>选中并滚动到某行（只定位，不进入编辑）。当前单元格必须跟着走，否则选中状态会被原来的当前单元格抢回去。</summary>
        private static void RevealRow(DataGridView dgv, int rowIndex)
        {
            dgv.ClearSelection();
            dgv.Rows[rowIndex].Selected = true;
            dgv.FirstDisplayedScrollingRowIndex = Math.Max(0, rowIndex - 5);
            if (dgv.Columns.Count > 0)
                dgv.CurrentCell = dgv.Rows[rowIndex].Cells[Math.Min(1, dgv.Columns.Count - 1)];
        }

        /// <summary>当前表的最大行号（物理行号可能大于行数，因为一条记录能占多行）</summary>
        private int MaxDisplayLineNumber()
        {
            int maxLine = _fileData.Rows.Count + _frozenRows;
            foreach (var row in _fileData.Rows)
                if (row.PhysicalEndLine > maxLine) maxLine = row.PhysicalEndLine;
            return maxLine;
        }

        /// <summary>
        /// 行号区宽度跟随最大行号的位数，避免 5 位以上行号被裁。
        /// RowHeadersWidthSizeMode 用 DisableResizing，否则 AutoSize 会按内容重算并覆盖这里的设置。
        /// </summary>
        private void UpdateRowHeadersWidth()
        {
            int maxLine = MaxDisplayLineNumber();

            int width = Math.Max(50,
                TextRenderer.MeasureText(new string('9', maxLine.ToString().Length), _rowNumberFont).Width + 14);
            _dataGrid.RowHeadersWidth = width;
            _dataGridHeader.RowHeadersWidth = width;
        }

        /// <summary>底部状态栏的选中标记摘要（行号区间 + 行数）</summary>
        private void UpdateSelectedRowsLabel()
        {
            if (_markedLineNumbers.Count == 0)
            {
                _lblRowNumber.Text = "";
                return;
            }
            int min = _markedLineNumbers.Min();
            int max = _markedLineNumbers.Max();
            _lblRowNumber.Text = $"已标记 {_markedLineNumbers.Count} 行（源文件 L{min}–L{max}）";
        }

        private void OnFrozenRowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            var rowIndex = GetFrozenDisplayLineNumber(e.RowIndex);
            var rowHeaderBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, _dataGridHeader.RowHeadersWidth, e.RowBounds.Height);
            using (var brush = new SolidBrush(_colorRowNumberText))
            {
                var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                e.Graphics.DrawString(FormatRowNumber(rowIndex), _rowNumberFont, brush, rowHeaderBounds, format);
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

        private void ShowHistoryButtonIfVcs()
        {
            try
            {
                if (_btnHistory != null) return;
                if (string.IsNullOrEmpty(_filePath) || !File.Exists(_filePath)) return;

                var vcs = VersionControlHelper.DetectVcs(_filePath);
                if (vcs == VcsType.None) return;

                var historyText = vcs == VcsType.Svn ? "SVN历史" : "Git历史";
                _btnHistory = new ThemeButton
                {
                    Text = historyText,
                    Size = new Size(ThemeButton.FitWidth(historyText), ThemeMetrics.RowH),
                    Font = Theme.Body,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Margin = new Padding(0, 0, ThemeMetrics.Gap, 0)
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
                        mappingRowIndex: _mappingRowIndex,
                        cache: _cache);
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

        /// <summary>直接设置 FileTableData 并刷新行映射（供版本对比等场景使用，绕过 BuildFileData 继承逻辑）</summary>
        public void SetFileData(FileTableData fileData)
        {
            _fileData = fileData;
            RebuildRowMap();
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
            if (!_pnlGridContainer.Panel1Collapsed && _dataGridHeader.Parent == null)
            {
                // 复用已有的 _headerPanel，避免每次创建新 Panel 导致空 Panel 堆积遮挡 _dataGridHeader
                if (_headerPanel != null)
                {
                    // 确保旧 Panel 仍在 Panel1 中（DetachGridFromParent 只移除 _dataGridHeader，不移除 _headerPanel）
                    if (!_pnlGridContainer.Panel1.Controls.Contains(_headerPanel))
                        _pnlGridContainer.Panel1.Controls.Add(_headerPanel);
                    _headerPanel.Controls.Add(_dataGridHeader);
                }
                else
                {
                    // 首次创建（理论上不会走到这里，初始化时已创建）
                    var headerPanel = new Panel
                    {
                        Dock = DockStyle.Fill,
                        Padding = new Padding(0, 0, SystemInformation.VerticalScrollBarWidth, 0)
                    };
                    headerPanel.Controls.Add(_dataGridHeader);
                    _pnlGridContainer.Panel1.Controls.Add(headerPanel);
                    _headerPanel = headerPanel;
                }
            }
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
                int startLine = -1;
                int endLine = -1;
                bool isBlank = false;
                var diffType = DiffRowType.None;

                if (inheritFrom != null)
                {
                    // ★ 如果新旧表行数一致 → 按索引继承（排序场景）
                    if (dt.Rows.Count == inheritFrom.Rows.Count && r < inheritFrom.Rows.Count)
                    {
                        origRowIdx = inheritFrom.Rows[r].OriginalRowIndex;
                        isDiff = inheritFrom.Rows[r].IsDiffRow;
                        diffCols = inheritFrom.Rows[r].DiffColumns;
                        startLine = inheritFrom.Rows[r].PhysicalStartLine;
                        endLine = inheritFrom.Rows[r].PhysicalEndLine;
                        isBlank = inheritFrom.Rows[r].IsBlankRow;
                        diffType = inheritFrom.Rows[r].DiffType;
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
                            startLine = matched.PhysicalStartLine;
                            endLine = matched.PhysicalEndLine;
                            isBlank = matched.IsBlankRow;
                            diffType = matched.DiffType;
                        }
                    }
                }

                var row = new TableRow
                {
                    OriginalRowIndex = origRowIdx,
                    PhysicalStartLine = startLine,
                    PhysicalEndLine = endLine,
                    IsDiffRow = isDiff,
                    DiffColumns = diffCols,
                    IsBlankRow = isBlank,
                    DiffType = diffType
                };
                for (int c = 0; c < dt.Columns.Count; c++)
                    row.Cells.Add(dt.Rows[r][c]?.ToString() ?? "");

                fd.Rows.Add(row);
            }
            return fd;
        }

        public void SetDataTable(DataTable dataTable, string infoText = "", FileTableData? fileData = null)
        {
            if (dataTable == null) return;

            _isEmbeddedMode = true; // 标记为外部嵌入模式（版本对比等），禁止编辑

            // ★ 彻底清理旧绑定，避免多次切换后事件/引用累积导致卡顿
            _dataGrid.DataSource = null;
            _dataGridHeader.DataSource = null;
            // 换表后旧的标记行不再属于当前数据，清掉避免残留标记
            ClearRowSelection();
            // 清理可能由外部附加的 CellFormatting 事件（如版本对比的差异高亮）
            _dataGrid.CellFormatting -= OnExternalCellFormatting;
            // 清理高亮状态
            _highlightRows.Clear();
            _highlightsReady = false;

            _dataTable = dataTable;
            // ★ 外部给了真实 FileTableData 且行数对得上就直接用它（保住物理行号和占位行标记），
            //   否则从 DataTable 反推（继承旧差异信息）
            _fileData = fileData != null && fileData.Rows.Count == dataTable.Rows.Count
                ? fileData
                : BuildFileData(dataTable, _fileData);
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
                _externalColumnStates = null; // ★ else分支也需要清除，避免残留的旧列状态影响下次调用
            }

            RebuildHighlightRows();
            _lblCellPos.Text = string.IsNullOrEmpty(infoText)
                ? $"行数: {dataTable.Rows.Count} | 列数: {dataTable.Columns.Count}"
                : infoText;
            _lblInfo.Text = "";

            // ★ 如果有重复警告，追加警告文本并保持红色
            if (_duplicateWarningText != null)
            {
                _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 30, 30));
                _lblInfo.Text = _duplicateWarningText;
            }

            // 隐藏"重新加载"和"历史记录"按钮（避免版本对比中无限嵌套）
            if (_btnReload != null) _btnReload.Visible = false;
            if (_btnHistory != null) _btnHistory.Visible = false;

            if (_frozenRows > 0)
                RefreshGridLayout();
            else
            {
                SetFrozenPanelVisible(false);
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
                // ★ 登记显示行 → 源行映射
                RebuildRowMap();

                // ★ 无冻结行模式下计算过滤匹配单元格（全表作为 DataSource，显示行 i 对应原始行 i）
                _filterMatchCells.Clear();
                _contentSearchMatchCells = GetContentMatchCells(_txtContentFilter.Text.Trim());
                var (_, advMatchCellsNoFreeze) = GetAdvVisibleRowIndices();
                _advFilterMatchCells = advMatchCellsNoFreeze;
                var mergedNoFreeze = MergeMatchCells(_contentSearchMatchCells, _advFilterMatchCells);
                foreach (var (origRow, origCols) in mergedNoFreeze)
                {
                    if (origRow < 0 || origRow >= _dataGrid.Rows.Count) continue;
                    var displayCols = new HashSet<int>();
                    foreach (var origCol in origCols)
                    {
                        if (origCol >= 0 && origCol < _visibleColumnIndices.Count)
                            displayCols.Add(origCol);
                    }
                    if (displayCols.Count > 0)
                        _filterMatchCells[origRow] = displayCols;
                }
            }

            // 外部注入数据的路径不走 RefreshGridLayout，这里同样套用冻结列设置
            ApplyFrozenColumns();

            // ★ 数据表变更后刷新唯一列下拉选项
            RefreshUniqueColumnCombo();
            // ★ 更新列管理按钮文本（显示"列管理(x/y)"，有隐藏列时红色显示）
            UpdateColumnSelectorButton();

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
                // ★ 对引用模板的列，从模板加载最新配置填充运行时数据
                foreach (var link in _columnLinks)
                {
                    link.RefreshFromTemplate();
                }
                if (_columnLinks.Count > 0)
                    FireAndForgetBuildLinkLookups(_columnLinks, _filePath);
            }
        }

        /// <summary>供外部（版本对比等）移除的 CellFormatting 事件占位</summary>
        internal event DataGridViewCellFormattingEventHandler? OnExternalCellFormatting;

        /// <summary>供外部（冲突解决等）使用的 CellMouseDown 事件占位。当 ContextMenuMode.Conflict 时触发。</summary>
        internal event DataGridViewCellMouseEventHandler? OnExternalCellMouseDown;

        /// <summary>设置外部传入的列可见性状态（版本对比等场景使用）。下次 SetDataTable 调用时会优先使用此状态。
        /// 同时保存完整列数/可见列数信息，用于 UpdateColumnSelectorButton 显示正确的"列管理(x/y)"文本。</summary>
        public void SetColumnVisibility(bool[]? states)
        {
            _externalColumnStates = states;
            if (states != null && states.Length > 1)
            {
                _externalTotalColumns = states.Length - 1;
                int visible = 0;
                for (int i = 1; i < states.Length; i++)
                    if (states[i]) visible++;
                _externalVisibleColumns = visible;
            }
            else
            {
                _externalTotalColumns = null;
                _externalVisibleColumns = null;
            }
        }

        /// <summary>获取当前列显隐状态的副本（用于切换数据时保持用户选择的列显示/隐藏）</summary>
        public bool[] GetColumnCheckStates()
        {
            return (bool[])_columnCheckStates.Clone();
        }

        /// <summary>设置嵌入式模式（冲突解决等场景使用）。设为 false 可使表格变为可编辑。</summary>
        public void SetEmbeddedMode(bool embedded)
        {
            _isEmbeddedMode = embedded;
        }

        /// <summary>禁用列管理和所有配置控件（版本对比等场景，配置由外部传入，不允许用户修改）。
        /// 控件保持显示（列管理按钮仍显示列数量和红色提示，ComboBox 仍显示当前值），但全部禁用交互。</summary>
        public void DisableColumnManagement()
        {
            _columnManagementDisabled = true;
            // ★ 不设置 _btnSelectColumns.Enabled = false，禁用态会把文字和描边混灰，
            //   盖掉"有隐藏列"的红色警示。改为用标志位在 Click 事件中拦截，按钮外观保持正常。
            // 禁用 _flowTop 中所有 ComboBox（固定行/表头行/映射行/唯一列）
            if (_flowTop != null)
            {
                foreach (Control c in _flowTop.Controls)
                {
                    if (c is ComboBox combo)
                        combo.Enabled = false;
                }
            }
        }

        // ==================== 行编辑模式 ====================

        /// <summary>进入行编辑模式</summary>
        /// <param name="displayRowIndex">DataGridView 显示行索引</param>
        /// <param name="columnIndex">可选：进入编辑后定位到的列索引（-1=默认第一列）</param>
        private void EnterRowEditMode(int displayRowIndex, int columnIndex = -1)
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

                // 只清除当前行的已修改单元格记录（不影响其他行的标记）
                if (_editingRowOrigIndex >= 0)
                    _modifiedCellsInEdit.Remove(_editingRowOrigIndex);

                // 获取原始行索引（如果 _editingRowOrigIndex 尚未被外部设置）
                if (_editingRowOrigIndex < 0)
                {
                    if (RowTagAt(displayRowIndex) is TableRow tr)
                        _editingRowOrigIndex = tr.OriginalRowIndex;
                    else
                        _editingRowOrigIndex = displayRowIndex + _frozenRows;
                }

                // 临时解除只读 + 切换 EditMode 为 EditProgrammatically 阻止方向键导航
                _dataGrid.ReadOnly = false;
                _dataGrid.EditMode = DataGridViewEditMode.EditProgrammatically;

                // 刷新行高亮
                _dataGrid.InvalidateRow(displayRowIndex);

                // 自动定位到指定列或第一个可编辑列并进入编辑
                if (_dataGrid.Columns.Count > 0)
                {
                    int targetCol = columnIndex >= 0 && columnIndex < row.Cells.Count
                        ? columnIndex
                        : 0;
                    _dataGrid.CurrentCell = row.Cells[targetCol];
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

                // 恢复只读 + 恢复 EditMode
                _dataGrid.ReadOnly = true;
                _dataGrid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;

                // 暂不清理 _modifiedCellsInEdit，保存成功后再清理（保存失败时保留红色标记）

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
                        // 重建 _fileData；行数未变时按索引继承物理行号（编辑不应让行号区退化）
                        _fileData = BuildFileData(_dataTable, inheritFrom: _fileData);

                        // 重建关联缓存
                        if (_columnLinks.Count > 0)
                            FireAndForgetBuildLinkLookups(_columnLinks, _filePath);

                        // 一次性保存文件
                        bool saved = SaveCurrentFile();

                        if (saved)
                        {
                            _unsavedRowIndices.Clear();
                            // 保存成功，清除该行的已修改单元格标记（不影响其他行）
                            _modifiedCellsInEdit.Remove(origIndex);
                        }
                        else
                        {
                            // 保存失败，标记这一行为未保存（保留 _modifiedCellsInEdit 中的红色标记）
                            _hasUnsavedChanges = true;
                            _unsavedRowIndices.Add(origIndex);
                        }

                        // 刷新显示（重建 Grid 以反映 _dataTable 的变更）
                        int scrollRow = _dataGrid.FirstDisplayedScrollingRowIndex;
                        int currentRow = _dataGrid.CurrentRow?.Index ?? -1;
                        int currentCol = _dataGrid.CurrentCell?.ColumnIndex ?? -1;
                        RefreshGridLayout();
                        RestoreScrollAndCurrentCell(scrollRow, currentRow, currentCol);

                        // 如果保存失败，刷新后需要恢复状态栏提示（RefreshGridLayout 会清空 _lblInfo.Text）
                        if (!saved && _unsavedRowIndices.Count > 0)
                        {
                            ShowDirtyIndicator();
                        }
                    }
                    else
                    {
                        // 没有实际修改，只需刷新编辑行的高亮
                        _dataGrid.InvalidateRow(editingRow);
                    }
                }
                else
                {
                    // 不提交：把已经敲进网格的值退回进入编辑时的快照。
                    // 只重画一行的话，用户按「取消」的改动会一直留在网格里，之后任意一次保存都会把它写进源文件。
                    if (editingRow >= 0 && editingRow < _dataGrid.Rows.Count)
                    {
                        var row = _dataGrid.Rows[editingRow];
                        if (snapShot != null)
                        {
                            for (int col = 0; col < row.Cells.Count && col < snapShot.Length; col++)
                                row.Cells[col].Value = snapShot[col];
                        }
                        if (origIndex >= 0) _modifiedCellsInEdit.Remove(origIndex);
                        _dataGrid.InvalidateRow(editingRow);
                    }
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

                // 恢复只读 + 恢复 EditMode
                _dataGrid.ReadOnly = true;
                _dataGrid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;

                // 只清除当前行的已修改单元格记录（取消编辑时该行的修改被撤销）
                if (origIndex >= 0)
                    _modifiedCellsInEdit.Remove(origIndex);

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
                    if (_columnLinks.Count > 0)
                        FireAndForgetBuildLinkLookups(_columnLinks, _filePath);

                    // 行数变化后按行号映射不可靠，显式清标记；滚动位置尽量保持
                    int scrollRow = _dataGrid.FirstDisplayedScrollingRowIndex;
                    int currentRow = _dataGrid.CurrentRow?.Index ?? -1;
                    int currentCol = _dataGrid.CurrentCell?.ColumnIndex ?? -1;
                    ClearRowSelection();
                    RefreshGridLayout();
                    RestoreScrollAndCurrentCell(scrollRow, currentRow, currentCol);
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
            if (RowTagAt(displayRowIndex) is TableRow tr)
                return tr.OriginalRowIndex;
            // 回退：无冻结行时 displayRowIndex 即原始行索引
            return displayRowIndex;
        }

        /// <summary>在指定位置之后插入空行</summary>
        /// <param name="afterOriginalRowIndex">在此原始行索引之后插入（-1=最前面插入，>=Rows.Count=末尾追加）</param>
        private void InsertRowAt(int afterOriginalRowIndex)
        {
            if (_isEmbeddedMode)
            {
                AlertForm.Show(this, "当前状态不支持插入行", "提示", AlertType.Information);
                return;
            }

            // 如果有排序，先取消排序
            if (_sortColumnIndex >= 0)
            {
                var result = AlertForm.ShowConfirm(this,
                    "当前处于排序状态，插入行前需要先取消排序。是否继续？",
                    "提示", AlertType.Warning);
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
            if (_columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_columnLinks, _filePath);

            // 行数变化后物理行号口径会变化，按行号映射的标记无法可靠还原，显式清掉
            ClearRowSelection();

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

            // 在行映射中定位到新行
            int displayRow = FindDisplayRowByOrigIndex(newOrigIndex);
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
                AlertForm.Show(this, "当前状态不支持复制行", "提示", AlertType.Information);
                return;
            }

            if (originalRowIndex < 0 || originalRowIndex >= _dataTable.Rows.Count)
                return;

            // 如果有排序，先取消排序
            if (_sortColumnIndex >= 0)
            {
                var result = AlertForm.ShowConfirm(this,
                    "当前处于排序状态，复制行前需要先取消排序。是否继续？",
                    "提示", AlertType.Warning);
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
            if (_columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_columnLinks, _filePath);

            // 行数变化后物理行号口径会变化，按行号映射的标记无法可靠还原，显式清掉
            ClearRowSelection();

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
            int dupDisplayRow = FindDisplayRowByOrigIndex(dupOrigIndex);
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
                AlertForm.Show(this, "当前状态不支持删除行", "提示", AlertType.Information);
                return;
            }

            if (originalRowIndex < 0 || originalRowIndex >= _dataTable.Rows.Count)
                return;

            // ==================== 行编辑模式：不允许删除正在编辑的行 ====================
            if (_editingRowIndex >= 0 && _editingRowOrigIndex == originalRowIndex)
            {
                AlertForm.Show(this, "请先完成或取消当前行的编辑再删除。",
                    "提示", AlertType.Warning);
                return;
            }

            // 冻结行不允许删除
            if (originalRowIndex < _frozenRows)
            {
                AlertForm.Show(this, "冻结行不允许删除。请先调整冻结行设置。",
                    "提示", AlertType.Warning);
                return;
            }

            var result = AlertForm.ShowConfirm(this,
                $"确认删除第 {originalRowIndex + 1} 行？此操作不可撤销。",
                "确认删除", AlertType.Warning);
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
            if (_columnLinks.Count > 0)
                FireAndForgetBuildLinkLookups(_columnLinks, _filePath);

            // 行数变化后物理行号口径会变化，按行号映射的标记无法可靠还原，显式清掉
            ClearRowSelection();

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
        /// <returns>true=保存成功, false=保存失败</returns>
        private bool SaveCurrentFile()
        {
            if (_fileType == TableFileType.Csv)
                return SaveCsvFile();
            else
                return SaveExcelFile();
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

            bool saved = _cache.SaveFile(_filePath, _dataTable, _delimiter);

            if (saved)
            {
                _hasUnsavedChanges = false;
                _unsavedRowIndices.Clear();
                HideDirtyIndicator();
                int scrollRow = _dataGrid.FirstDisplayedScrollingRowIndex;
                int currentRow = _dataGrid.CurrentRow?.Index ?? -1;
                int currentCol = _dataGrid.CurrentCell?.ColumnIndex ?? -1;
                RefreshGridLayout();
                RestoreScrollAndCurrentCell(scrollRow, currentRow, currentCol);
                // SaveFile 已同步缓存条目并清掉指向本文件的 lookup，
                // 这里再通知同缓存页签重算关联着色
                _cache.NotifyFileSaved(_filePath);
            }
            else
            {
                _retrySaveTimer?.Start();
            }
        }

        /// <summary>关闭前强制保存（供外部 CsvReaderForm 等调用）</summary>
        public bool ForceFlushBeforeClose()
        {
            _retrySaveTimer?.Stop();
            if (!_hasUnsavedChanges) return true;

            TryFlushToDisk();
            if (!_hasUnsavedChanges) return true;

            // 仍失败 → 弹窗让用户决定
            var result = AlertForm.ShowConfirm(null,
                $"文件被占用无法保存，{_unsavedRowIndices.Count} 行修改仅保留在内存中。" + GlobalConst.NewLine + "关闭后修改将丢失，是否继续？",
                "保存失败", AlertType.Warning);
            return result == DialogResult.Yes;
        }

        private void ShowDirtyIndicator()
        {
            if (_lblInfo.IsHandleCreated)
            {
                _lblInfo.Text = $"? 文件被占用，{_unsavedRowIndices.Count} 行修改未保存（等待重试...）";
                _lblInfo.ForeColor = Theme.Signal(Color.FromArgb(200, 130, 30));
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

        /// <summary>编辑模式下阻止左右方向键在文本边界时跳格</summary>
        /// <summary>在 ProcessDialogKey 之前拦截方向键，阻止 DataGridView 导航</summary>
        private void OnEditTextBoxPreviewKeyDown(object? sender, PreviewKeyDownEventArgs e)
        {
            if (_editingRowIndex >= 0)
            {
                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right
                    || e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
                {
                    e.IsInputKey = true; // 告诉 DataGridView：这个键由编辑控件自己处理
                }
            }
        }

        /// <summary>编辑模式下阻止方向键跳格</summary>
        private void OnEditTextBoxKeyDown(object? sender, KeyEventArgs e)
        {
            if (sender is not TextBox tb) return;

            // 上下左右方向键一律阻止，防止跳格
            if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right
                || e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;

                // 左右键：在 TextBox 内模拟光标移动（上/下键不模拟，保持原地）
                if (e.KeyCode == Keys.Left && tb.SelectionStart > 0)
                    tb.Select(tb.SelectionStart - 1, 0);
                else if (e.KeyCode == Keys.Right && tb.SelectionStart < tb.Text.Length)
                    tb.Select(tb.SelectionStart + 1, 0);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _cache.FileSaved -= OnCacheFileSaved;
                _retrySaveTimer?.Stop();
                _retrySaveTimer?.Dispose();
                _retrySaveTimer = null;
                _linkToolTip?.Dispose();
                _linkToolTip = null;
                _btnHistory?.Dispose();
                _btnHistory = null;
                // ★ 只释放自建的缓存；宿主传入的缓存由宿主窗口统一释放
                if (_ownsCache) _cache.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
