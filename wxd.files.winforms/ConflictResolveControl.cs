using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;

namespace wxd.files.winforms
{
    /// <summary>冲突解决控件：可嵌入 TabPage 的冲突解决面板</summary>
    public partial class ConflictResolveControl : UserControl
    {
        // ==================== 布局控件 ====================
        private SplitContainer _splitMain;
        private SplitContainer _splitCompare;
        private Panel _pnlToolbar;
        private Label _lblInfo;

        // ==================== 对比视图 ====================
        private TableBrowserControl? _remoteView;
        private TableBrowserControl? _localView;
        /// <summary>宿主阅读器窗口的表格数据缓存；为 null 时三个视图各自持有私有缓存</summary>
        private readonly FileDataCache? _ownerCache;
        private Label _lblRemoteHeader;
        private Label _lblLocalHeader;

        // ==================== 合并结果 ====================
        private TableBrowserControl _mergedView;
        private DataTable _mergedTable;
        private Panel _pnlMergedHeader;

        // ==================== 数据 ====================
        private string _filePath;
        private VcsType _vcsType;
        private DataTable _remoteData;
        private DataTable _localData;
        // ★ 远程/本地文件的原始解析结果：带源文件物理行号，行号显示要靠它
        private FileTableData? _remoteFileData;
        private FileTableData? _localFileData;
        private string _tempDir;
        private int _frozenRows;
        private int _headerRowIndex;
        private int _mappingRowIndex = -1;
        private int _uniqueColumnIndex = -1;
        private int _idColumnIndex = -1;

        // ==================== 差异数据 ====================
        private HashSet<int> _remoteDiffRows = new();
        private HashSet<int> _localDiffRows = new();
        private HashSet<int> _remoteDiffCells = new();
        private HashSet<int> _localDiffCells = new();
        private HashSet<int> _conflictCells = new();

        private List<int> _remoteRowToOriginal = new();
        private List<int> _localRowToOriginal = new();
        private List<int> _remoteColToOriginal = new();
        private List<int> _localColToOriginal = new();

        private List<(DiffRowType diffType, bool isBlankRow)> _remoteRowProps = new();
        private List<(DiffRowType diffType, bool isBlankRow)> _localRowProps = new();

        // ID对齐数据
        private bool _useIdAlign = false;
        private HashSet<int> _remoteDeletedRows = new();
        private HashSet<int> _localAddedRows = new();

        // ==================== 选择状态 ====================
        private bool _selectionFromRemote = false;
        private HashSet<int> _selectedRemoteRows = new();
        private HashSet<int> _selectedLocalRows = new();
        private (int row, int col)? _selectedRemoteCell;
        private (int row, int col)? _selectedLocalCell;

        // ==================== 滚动同步 ====================
        private bool _syncingScroll = false;

        // ==================== 配色 ====================
        private static readonly Color _colorSurface = Color.White;
        private static readonly Color _colorBorder = Color.FromArgb(173, 209, 237);
        private static readonly Color _colorLightText = Color.FromArgb(20, 60, 100);
        private static readonly Color _colorLightTextSecondary = Color.FromArgb(90, 130, 170);
        private static readonly Color _colorPrimary = Color.FromArgb(66, 165, 245);
        private static readonly Color _colorPrimaryDark = Color.FromArgb(30, 136, 229);
        private static readonly Color _colorHoverBg = Color.FromArgb(235, 245, 255);
        private static readonly Color _colorDiffCellBg = Color.FromArgb(255, 200, 200);
        private static readonly Color _colorDiffRowBg = Color.FromArgb(255, 253, 220);
        private static readonly Color _colorAddedRowBg = Color.FromArgb(213, 245, 227);
        private static readonly Color _colorDeletedRowBg = Color.FromArgb(250, 219, 216);
        private static readonly Color _colorConflictBg = Color.FromArgb(255, 180, 180);
        private static readonly Color _colorSelectedRemoteBg = Color.FromArgb(180, 210, 255);
        private static readonly Color _colorSelectedLocalBg = Color.FromArgb(180, 255, 200);
        private static readonly Color _colorBlankRowForAdded = Color.FromArgb(232, 250, 240);
        private static readonly Color _colorBlankRowForDeleted = Color.FromArgb(253, 235, 233);
        private static readonly Color _colorHeaderBg = Color.FromArgb(240, 246, 252);

        // ==================== 事件 ====================
        /// <summary>保存完成事件（保存成功后触发）</summary>
        public event EventHandler? SaveCompleted;

        // ==================== 构造函数 ====================
        public ConflictResolveControl(string filePath, VcsType vcsType,
            DataTable remoteData, DataTable localData,
            int frozenRows, int headerRowIndex,
            string? rootDir = null,
            int uniqueColumnIndex = -1,
            int mappingRowIndex = -1,
            FileTableData? remoteFileData = null,
            FileTableData? localFileData = null,
            FileDataCache? cache = null)
        {
            _filePath = filePath;
            _vcsType = vcsType;
            _ownerCache = cache;
            _remoteData = remoteData;
            _localData = localData;
            _remoteFileData = remoteFileData;
            _localFileData = localFileData;
            _frozenRows = frozenRows;
            _headerRowIndex = headerRowIndex;
            _uniqueColumnIndex = uniqueColumnIndex;
            _mappingRowIndex = mappingRowIndex;

            _tempDir = Path.Combine(Path.GetTempPath(), $"wxd_conflict_{Guid.NewGuid():N}");
            Directory.CreateDirectory(_tempDir);

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Dock = DockStyle.Fill;
            BackColor = _colorSurface;
            Font = new Font("Microsoft YaHei UI", 9f);

            SuspendLayout();

            // === 顶部工具栏 ===
            _pnlToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 48,
                BackColor = Color.FromArgb(245, 250, 255),
                Padding = new Padding(8, 6, 8, 4)
            };

            var flowToolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                AutoSize = false,
                BackColor = Color.Transparent
            };

            var toolTip = new ToolTip { AutoPopDelay = 5000, InitialDelay = 500, ReshowDelay = 100 };

            var btnUseRemoteRow = CreateFlowButton("← 使用远程此行", _colorPrimary);
            btnUseRemoteRow.Click += (s, e) => UseRemoteSelected();
            toolTip.SetToolTip(btnUseRemoteRow, "将远程版本选中行/单元格复制到合并结果");

            var btnUseLocalRow = CreateFlowButton("使用本地此行 →", Color.FromArgb(52, 168, 83));
            btnUseLocalRow.Click += (s, e) => UseLocalSelected();
            toolTip.SetToolTip(btnUseLocalRow, "将本地版本选中行/单元格复制到合并结果");

            var btnSave = CreateFlowButton("💾 保存", _colorPrimaryDark);
            btnSave.Click += (s, e) => SaveMergedResult();

            var btnSaveAs = CreateFlowButton("另存为...", Color.FromArgb(120, 144, 156));
            btnSaveAs.Click += (s, e) => SaveMergedResultAs();

            var btnUseAllRemote = CreateFlowButton("全部使用远程", Color.FromArgb(90, 140, 210));
            btnUseAllRemote.Click += (s, e) => UseAllRemote();
            toolTip.SetToolTip(btnUseAllRemote, "将远程版本所有差异行复制到合并结果");

            var btnUseAllLocal = CreateFlowButton("全部使用本地", Color.FromArgb(90, 190, 110));
            btnUseAllLocal.Click += (s, e) => UseAllLocal();
            toolTip.SetToolTip(btnUseAllLocal, "将本地版本所有差异行复制到合并结果");

            var btnSelectAllRemote = CreateFlowButton("全选远程", Color.FromArgb(100, 150, 220));
            btnSelectAllRemote.Click += (s, e) => SelectAllRemoteRows();
            toolTip.SetToolTip(btnSelectAllRemote, "选中远程版本所有差异行");

            var btnSelectAllLocal = CreateFlowButton("全选本地", Color.FromArgb(100, 200, 120));
            btnSelectAllLocal.Click += (s, e) => SelectAllLocalRows();
            toolTip.SetToolTip(btnSelectAllLocal, "选中本地版本所有差异行");

            _lblInfo = new Label
            {
                AutoSize = true,
                Margin = new Padding(8, 6, 0, 0),
                ForeColor = _colorLightTextSecondary,
                Font = new Font("Microsoft YaHei UI", 9f),
                TextAlign = ContentAlignment.MiddleLeft
            };

            flowToolbar.Controls.AddRange(new Control[] {
                btnUseRemoteRow, btnUseLocalRow,
                btnSave, btnSaveAs,
                btnUseAllRemote, btnUseAllLocal,
                btnSelectAllRemote, btnSelectAllLocal,
                _lblInfo
            });

            _pnlToolbar.Controls.Add(flowToolbar);

            // === 上下分割 ===
            _splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 300,
                Panel1MinSize = 150,
                Panel2MinSize = 150,
                BackColor = Color.FromArgb(200, 225, 245),
                SplitterWidth = 3
            };

            // === 上半部分：左右对比 ===
            _splitCompare = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 50,
                BackColor = Color.FromArgb(200, 225, 245),
                SplitterWidth = 3
            };

            // 远程版本（左侧）
            _lblRemoteHeader = new Label
            {
                Text = "远程版本 (服务器)",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(255, 245, 245),
                ForeColor = Color.FromArgb(200, 50, 50),
                Font = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold),
                BorderStyle = BorderStyle.FixedSingle
            };
            var remotePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            remotePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            remotePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            remotePanel.Controls.Add(_lblRemoteHeader, 0, 0);
            _splitCompare.Panel1.Controls.Add(remotePanel);

            // 本地版本（右侧）
            _lblLocalHeader = new Label
            {
                Text = "本地版本 (工作区)",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(245, 255, 245),
                ForeColor = Color.FromArgb(50, 160, 50),
                Font = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold),
                BorderStyle = BorderStyle.FixedSingle
            };
            var localPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            localPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            localPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            localPanel.Controls.Add(_lblLocalHeader, 0, 0);
            _splitCompare.Panel2.Controls.Add(localPanel);

            _splitMain.Panel1.Controls.Add(_splitCompare);

            // === 下半部分：合并结果 ===
            _pnlMergedHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 30,
                BackColor = Color.FromArgb(245, 250, 255)
            };
            var lblMergeHeader = new Label
            {
                Text = "合并结果 (可编辑)",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = _colorPrimaryDark,
                Font = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold),
                BorderStyle = BorderStyle.FixedSingle
            };
            _pnlMergedHeader.Controls.Add(lblMergeHeader);

            var mergedOptions = new TableBrowserOptions(_filePath)
            {
                RootDir = Path.GetDirectoryName(_filePath),
                FrozenRows = _frozenRows,
                HeaderRowIndex = _headerRowIndex,
                UniqueColumnIndex = _uniqueColumnIndex,
                MappingRowIndex = _mappingRowIndex,
                ContextMenuMode = ContextMenuMode.None,  // 统一冲突解决模式右键菜单
                SkipAutoLoad = true,
                ShowMinimalUI = true,
                Cache = _ownerCache,
            };
            _mergedView = new TableBrowserControl(mergedOptions) { Dock = DockStyle.Fill };
            _mergedView.GetDataGridView().CellFormatting += OnMergedCellFormatting;
            _mergedView.GetDataGridView().CellEndEdit += OnMergedCellEndEdit;
            _mergedView.GetDataGridView().KeyDown += OnMergedGridKeyDown;

            var mergePanel = new Panel { Dock = DockStyle.Fill };
            mergePanel.Controls.Add(_mergedView);
            mergePanel.Controls.Add(_pnlMergedHeader);

            _splitMain.Panel2.Controls.Add(mergePanel);

            Controls.Add(_splitMain);
            Controls.Add(_pnlToolbar);

            // 初始化分割距离
            _splitCompare.Resize += (s, e) =>
            {
                if (_splitCompare.Width > 0)
                    _splitCompare.SplitterDistance = _splitCompare.Width / 2;
            };

            ResumeLayout(false);
        }

        private Button CreateFlowButton(string text, Color backColor)
        {
            var btn = new Button
            {
                Text = text,
                AutoSize = true,
                Margin = new Padding(2, 4, 2, 0),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229), MouseOverBackColor = Color.FromArgb(30, 136, 229) }
            };
            return btn;
        }

        // ==================== 新增/删除/修改行处理映射 ====================
        /// <summary>将远程新增行（本地无对应ID）的行索引集合</summary>
        private HashSet<int> _remoteNewRowOriginalIndices = new();
        /// <summary>将本地删除行（远程无对应ID）的行索引集合</summary>
        private HashSet<int> _localDeletedRowOriginalIndices = new();

        // ==================== 加载对比数据 ====================
        public async Task LoadComparisonAsync()
        {
            _lblInfo.Text = "正在计算差异...";

            try
            {
                EnsureViewsCreated();

                // ★ 阶段1：后台计算差异（只做差异标记）
                await ComputeDiffAsync();

                // ★ 阶段2：构建对齐后的 FileTableData（双指针归并，空白行插入到正确位置）
                var (alignedRemoteFd, alignedLocalFd) = BuildAlignedTables();

                // ★ 阶段3：构建对齐后的DataTable，设置到视图
                var remoteDt = BuildDataTableFromFileData(alignedRemoteFd);
                var localDt = BuildDataTableFromFileData(alignedLocalFd);

                _remoteView!.SetDataTable(remoteDt, $"远程版本 | 行数: {remoteDt.Rows.Count}", alignedRemoteFd);
                _localView!.SetDataTable(localDt, $"本地版本 | 行数: {localDt.Rows.Count}", alignedLocalFd);

                // ★ 阶段4：构建行/列索引映射（跳过前 _frozenRows 行，确保 rowProps.Count == DGV 行数）
                _remoteRowToOriginal = BuildRowIndexMap(alignedRemoteFd, out _remoteRowProps);
                _localRowToOriginal = BuildRowIndexMap(alignedLocalFd, out _localRowProps);

                // ★ 阶段5：重新绑定差异高亮事件（SetDataTable 内部会移除 OnExternalCellFormatting）
                ApplyDiffHighlight();

                // ★ 阶段6：初始化合并结果表（包含所有行）
                InitializeMergedTable();

                if (_splitCompare.Width > 0)
                    _splitCompare.SplitterDistance = _splitCompare.Width / 2;
                if (_splitMain.Height > 0)
                    _splitMain.SplitterDistance = _splitMain.Height / 2;

                int totalDiffRows = _remoteDiffRows.Count + _localDiffRows.Count;
                _lblInfo.Text = $"差异行: {totalDiffRows}  |  冲突单元格: {_conflictCells.Count}  |  提示: 点击远程/本地行头选中行，然后点击工具栏按钮合并";
            }
            catch (Exception ex)
            {
                _lblInfo.Text = $"加载失败: {ex.Message}";
                AlertForm.Show(ParentForm, $"加载失败: {ex.Message}", "错误", AlertType.Error, exception: ex);
            }
        }

        private void EnsureViewsCreated()
        {
            var remotePanel = _splitCompare.Panel1.Controls[0] as TableLayoutPanel;
            var localPanel = _splitCompare.Panel2.Controls[0] as TableLayoutPanel;

            if (_remoteView == null && remotePanel != null)
            {
                var options = new TableBrowserOptions(_filePath)
                {
                    RootDir = Path.GetDirectoryName(_filePath),
                    FrozenRows = _frozenRows,
                    HeaderRowIndex = _headerRowIndex,
                    UniqueColumnIndex = _uniqueColumnIndex,
                    MappingRowIndex = _mappingRowIndex,
                    ContextMenuMode = ContextMenuMode.Conflict,  // 统一冲突解决模式右键菜单
                    ShowMinimalUI = true,
                    SkipAutoLoad = true,                     // 跳过构造函数自动加载，由外部 SetDataTable 注入数据
                    Cache = _ownerCache,
                };
                _remoteView = new TableBrowserControl(options) { Dock = DockStyle.Fill };
                remotePanel.Controls.Add(_remoteView, 0, 1);
            }

            if (_localView == null && localPanel != null)
            {
                var options = new TableBrowserOptions(_filePath)
                {
                    RootDir = Path.GetDirectoryName(_filePath),
                    FrozenRows = _frozenRows,
                    HeaderRowIndex = _headerRowIndex,
                    UniqueColumnIndex = _uniqueColumnIndex,
                    MappingRowIndex = _mappingRowIndex,
                    ContextMenuMode = ContextMenuMode.Conflict,  // 统一冲突解决模式右键菜单
                    ShowMinimalUI = true,
                    SkipAutoLoad = true,                     // 跳过构造函数自动加载，由外部 SetDataTable 注入数据
                    Cache = _ownerCache,
                };
                _localView = new TableBrowserControl(options) { Dock = DockStyle.Fill };
                localPanel.Controls.Add(_localView, 0, 1);
            }
        }

        // ==================== 差异计算（只做差异标记，不构建对齐） ====================
        /// <summary>计算差异，填充 _remoteDiffRows / _localDiffRows / _remoteDiffCells / _localDiffCells / _conflictCells 等</summary>
        private async Task ComputeDiffAsync()
        {
            _remoteDiffRows.Clear();
            _localDiffRows.Clear();
            _remoteDiffCells.Clear();
            _localDiffCells.Clear();
            _conflictCells.Clear();
            _remoteDeletedRows.Clear();
            _localAddedRows.Clear();
            _remoteNewRowOriginalIndices.Clear();
            _localDeletedRowOriginalIndices.Clear();

            if (_remoteData == null || _localData == null) return;

            int maxCols = Math.Max(_remoteData.Columns.Count, _localData.Columns.Count);
            _useIdAlign = DetectIdColumn();
            int dataStartRow = _frozenRows;

            await Task.Run(() =>
            {
                if (_useIdAlign && _idColumnIndex >= 0)
                {
                    int idCol = _idColumnIndex;

                    var remoteIdMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    var localIdMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    for (int r = dataStartRow; r < _remoteData.Rows.Count; r++)
                    {
                        string id = _remoteData.Rows[r][idCol]?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(id) && !remoteIdMap.ContainsKey(id))
                            remoteIdMap[id] = r;
                    }
                    for (int r = dataStartRow; r < _localData.Rows.Count; r++)
                    {
                        string id = _localData.Rows[r][idCol]?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(id) && !localIdMap.ContainsKey(id))
                            localIdMap[id] = r;
                    }

                    var allIds = new HashSet<string>(remoteIdMap.Keys, StringComparer.OrdinalIgnoreCase);
                    allIds.UnionWith(localIdMap.Keys);

                    foreach (var id in allIds)
                    {
                        bool remoteExists = remoteIdMap.TryGetValue(id, out int remoteRow);
                        bool localExists = localIdMap.TryGetValue(id, out int localRow);

                        if (remoteExists && !localExists)
                        {
                            _remoteDiffRows.Add(remoteRow);
                            _remoteNewRowOriginalIndices.Add(remoteRow);
                            for (int c = 0; c < _remoteData.Columns.Count; c++)
                                _remoteDiffCells.Add((remoteRow << 16) | c);
                        }
                        else if (!remoteExists && localExists)
                        {
                            _localDiffRows.Add(localRow);
                            _localDeletedRowOriginalIndices.Add(localRow);
                            for (int c = 0; c < _localData.Columns.Count; c++)
                                _localDiffCells.Add((localRow << 16) | c);
                        }
                        else
                        {
                            bool rowHasDiff = false;
                            var remoteDiffCols = new HashSet<int>();
                            var localDiffCols = new HashSet<int>();

                            for (int col = 0; col < maxCols; col++)
                            {
                                string? remoteVal = (col < _remoteData.Columns.Count)
                                    ? _remoteData.Rows[remoteRow][col]?.ToString() : null;
                                string? localVal = (col < _localData.Columns.Count)
                                    ? _localData.Rows[localRow][col]?.ToString() : null;

                                if (remoteVal != localVal)
                                {
                                    rowHasDiff = true;
                                    remoteDiffCols.Add(col);
                                    localDiffCols.Add(col);
                                }
                            }

                            if (rowHasDiff)
                            {
                                _remoteDiffRows.Add(remoteRow);
                                foreach (var c in remoteDiffCols)
                                    _remoteDiffCells.Add((remoteRow << 16) | c);
                                _localDiffRows.Add(localRow);
                                foreach (var c in localDiffCols)
                                    _localDiffCells.Add((localRow << 16) | c);

                                var commonCols = remoteDiffCols.Intersect(localDiffCols);
                                foreach (var c in commonCols)
                                    _conflictCells.Add((remoteRow << 16) | c);
                            }
                        }
                    }
                }
                else
                {
                    int maxRows = Math.Max(_remoteData.Rows.Count, _localData.Rows.Count);

                    for (int row = dataStartRow; row < maxRows; row++)
                    {
                        bool rowHasDiff = false;
                        HashSet<int> remoteDiffCols = new();
                        HashSet<int> localDiffCols = new();

                        for (int col = 0; col < maxCols; col++)
                        {
                            string? remoteVal = (row < _remoteData.Rows.Count && col < _remoteData.Columns.Count)
                                ? _remoteData.Rows[row][col]?.ToString() : null;
                            string? localVal = (row < _localData.Rows.Count && col < _localData.Columns.Count)
                                ? _localData.Rows[row][col]?.ToString() : null;

                            if (remoteVal != localVal)
                            {
                                rowHasDiff = true;
                                if (row < _remoteData.Rows.Count) remoteDiffCols.Add(col);
                                if (row < _localData.Rows.Count) localDiffCols.Add(col);
                            }
                        }

                        if (rowHasDiff)
                        {
                            if (row < _remoteData.Rows.Count)
                            {
                                _remoteDiffRows.Add(row);
                                foreach (var c in remoteDiffCols)
                                    _remoteDiffCells.Add((row << 16) | c);
                            }
                            if (row < _localData.Rows.Count)
                            {
                                _localDiffRows.Add(row);
                                foreach (var c in localDiffCols)
                                    _localDiffCells.Add((row << 16) | c);
                            }
                        }
                    }
                }
            });
        }

        // ==================== 构建对齐后的 FileTableData（与 VersionCompareForm 完全一致） ====================
        /// <summary>
        /// 用双指针归并算法构建对齐后的远程/本地 FileTableData。
        /// 固定行分别复制，数据行以本地（右表）ID 顺序为主，空白行插入到正确位置。
        /// </summary>
        private (FileTableData remoteAligned, FileTableData localAligned) BuildAlignedTables()
        {
            int dataStartRow = _frozenRows;
            int idCol = _idColumnIndex;

            // ★ 优先用文件解析出来的 FileTableData（带源文件物理行号），拿不到才按行索引临时构建
            var tempRemoteFd = ResolveVersionFileData(_remoteData, _remoteFileData);
            var tempLocalFd = ResolveVersionFileData(_localData, _localFileData);

            // 构建 ID → FileTableData 行索引 的映射
            var remoteIdMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var localIdMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < tempRemoteFd.Rows.Count; i++)
            {
                var row = tempRemoteFd.Rows[i];
                if (row.OriginalRowIndex < dataStartRow) continue;
                string? id = idCol < row.Cells.Count ? row.Cells[idCol]?.Trim() : null;
                if (!string.IsNullOrEmpty(id) && !remoteIdMap.ContainsKey(id))
                    remoteIdMap[id] = i;
            }
            for (int i = 0; i < tempLocalFd.Rows.Count; i++)
            {
                var row = tempLocalFd.Rows[i];
                if (row.OriginalRowIndex < dataStartRow) continue;
                string? id = idCol < row.Cells.Count ? row.Cells[idCol]?.Trim() : null;
                if (!string.IsNullOrEmpty(id) && !localIdMap.ContainsKey(id))
                    localIdMap[id] = i;
            }

            var remoteAligned = new FileTableData();
            var localAligned = new FileTableData();

            // 复制列名
            for (int c = 0; c < _remoteData.Columns.Count; c++)
                remoteAligned.ColumnNames.Add(_remoteData.Columns[c].ColumnName);
            for (int c = 0; c < _localData.Columns.Count; c++)
                localAligned.ColumnNames.Add(_localData.Columns[c].ColumnName);

            // ★ 先复制固定行/表头行（原样保留，不对齐）
            for (int i = 0; i < tempRemoteFd.Rows.Count; i++)
            {
                if (tempRemoteFd.Rows[i].OriginalRowIndex < dataStartRow)
                    remoteAligned.Rows.Add(CloneRow(tempRemoteFd.Rows[i], tempRemoteFd.ColumnNames.Count));
            }
            for (int i = 0; i < tempLocalFd.Rows.Count; i++)
            {
                if (tempLocalFd.Rows[i].OriginalRowIndex < dataStartRow)
                    localAligned.Rows.Add(CloneRow(tempLocalFd.Rows[i], tempLocalFd.ColumnNames.Count));
            }

            // ★ 合并 remoteIdsOrdered 和 localIdsOrdered 为有序去重列表，统一遍历
            var remoteIdsOrdered = SortIds(remoteIdMap.Keys);
            var localIdsOrdered = SortIds(localIdMap.Keys);
            var allIdsOrdered = new List<string>();
            int rp = 0, lp = 0;
            while (rp < remoteIdsOrdered.Count || lp < localIdsOrdered.Count)
            {
                if (rp >= remoteIdsOrdered.Count)
                    allIdsOrdered.Add(localIdsOrdered[lp++]);
                else if (lp >= localIdsOrdered.Count)
                    allIdsOrdered.Add(remoteIdsOrdered[rp++]);
                else
                {
                    int cmp = string.Compare(remoteIdsOrdered[rp], localIdsOrdered[lp], StringComparison.OrdinalIgnoreCase);
                    if (cmp == 0) { allIdsOrdered.Add(remoteIdsOrdered[rp]); rp++; lp++; }
                    else if (cmp < 0) allIdsOrdered.Add(remoteIdsOrdered[rp++]);
                    else allIdsOrdered.Add(localIdsOrdered[lp++]);
                }
            }

            // ★ 按有序 ID 遍历，统一处理三态（匹配/新增/删除）
            int colCount = Math.Max(_remoteData.Columns.Count, _localData.Columns.Count);

            foreach (var id in allIdsOrdered)
            {
                bool remoteHas = remoteIdMap.ContainsKey(id);
                bool localHas = localIdMap.ContainsKey(id);

                if (remoteHas && localHas)
                {
                    // 两表都有 → 匹配行
                    int remoteFdIdx = remoteIdMap[id];
                    int localFdIdx = localIdMap[id];
                    remoteAligned.Rows.Add(CloneRow(tempRemoteFd.Rows[remoteFdIdx], colCount));
                    localAligned.Rows.Add(CloneRow(tempLocalFd.Rows[localFdIdx], colCount));
                }
                else if (remoteHas)
                {
                    // 远程有、本地无 → 新增行（远程数据行 + 本地空白行）
                    int remoteFdIdx = remoteIdMap[id];
                    var remoteRow = CloneRow(tempRemoteFd.Rows[remoteFdIdx], colCount);
                    remoteRow.DiffType = DiffRowType.Added;
                    remoteRow.IsDiffRow = true;
                    remoteAligned.Rows.Add(remoteRow);
                    localAligned.Rows.Add(CreateBlankRow(tempRemoteFd.Rows[remoteFdIdx].OriginalRowIndex, colCount, DiffRowType.Added));
                }
                else
                {
                    // 本地有、远程无 → 删除行（远程空白行 + 本地数据行）
                    int localFdIdx = localIdMap[id];
                    remoteAligned.Rows.Add(CreateBlankRow(tempLocalFd.Rows[localFdIdx].OriginalRowIndex, colCount, DiffRowType.Deleted));
                    var localRow = CloneRow(tempLocalFd.Rows[localFdIdx], colCount);
                    localRow.DiffType = DiffRowType.Deleted;
                    localRow.IsDiffRow = true;
                    localAligned.Rows.Add(localRow);
                }
            }

            return (remoteAligned, localAligned);
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

        /// <summary>取某版本的 FileTableData：文件解析结果可用就用它（保留物理行号），否则按 DataTable 行索引临时构建</summary>
        private static FileTableData ResolveVersionFileData(DataTable dt, FileTableData? fileData)
            => fileData != null && fileData.Rows.Count == dt.Rows.Count ? fileData : BuildTempFileDataFromDataTable(dt);

        /// <summary>从 DataTable 构建临时 FileTableData（行索引作为 OriginalRowIndex，无物理行号）</summary>
        private static FileTableData BuildTempFileDataFromDataTable(DataTable dt)
        {
            var fd = new FileTableData();
            foreach (DataColumn col in dt.Columns)
                fd.ColumnNames.Add(col.ColumnName);
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                var row = new TableRow { OriginalRowIndex = r };
                for (int c = 0; c < dt.Columns.Count; c++)
                    row.Cells.Add(dt.Rows[r][c]?.ToString() ?? "");
                fd.Rows.Add(row);
            }
            return fd;
        }

        /// <summary>克隆 TableRow</summary>
        private static TableRow CloneRow(TableRow src, int colCount)
        {
            var row = new TableRow
            {
                OriginalRowIndex = src.OriginalRowIndex,
                PhysicalStartLine = src.PhysicalStartLine,
                PhysicalEndLine = src.PhysicalEndLine,
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

        /// <summary>构建DGV行索引→原始行索引映射</summary>
        private List<int> BuildRowIndexMap(FileTableData fd, out List<(DiffRowType diffType, bool isBlankRow)> rowProps)
        {
            var map = new List<int>();
            rowProps = new List<(DiffRowType, bool)>();
            // 跳过前 _frozenRows 行（固定行在 DGV 中不显示，包含空白行也一并跳过）
            for (int i = _frozenRows; i < fd.Rows.Count; i++)
            {
                int origIdx = fd.Rows[i].OriginalRowIndex;
                map.Add(origIdx);
                rowProps.Add((fd.Rows[i].DiffType, fd.Rows[i].IsBlankRow));
            }
            return map;
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

        /// <summary>从FileTableData的某一行获取该行在原始DataTable中的行索引</summary>
        private int GetOriginalRowIndex(int alignedRowIndex, bool isRemote)
        {
            var map = isRemote ? _remoteRowToOriginal : _localRowToOriginal;
            if (alignedRowIndex < map.Count)
                return map[alignedRowIndex];
            return -1;
        }

        private bool DetectIdColumn()
        {
            // ★ 优先使用配置的唯一列索引
            if (_uniqueColumnIndex >= 0 && _uniqueColumnIndex < Math.Min(_remoteData.Columns.Count, _localData.Columns.Count))
            {
                _idColumnIndex = _uniqueColumnIndex;
                return true;
            }

            // ★ 当 _uniqueColumnIndex=-1（配置未设置）时，默认使用第0列作为ID列
            //    （与 TableBrowserControl 中 _uniqueColumnIndex 默认值 0 保持一致）
            if (_remoteData.Columns.Count > 0 && _localData.Columns.Count > 0)
            {
                _idColumnIndex = 0;
                return true;
            }

            string[] idKeywords = { "id", "编号", "序号", "key", "code", "编码" };
            int maxCols = Math.Min(_remoteData.Columns.Count, _localData.Columns.Count);

            for (int c = 0; c < maxCols; c++)
            {
                string remoteName = (_headerRowIndex >= 0 && _headerRowIndex < _remoteData.Rows.Count)
                    ? _remoteData.Rows[_headerRowIndex][c]?.ToString()?.Trim()?.ToLowerInvariant() ?? ""
                    : "";
                string localName = (_headerRowIndex >= 0 && _headerRowIndex < _localData.Rows.Count)
                    ? _localData.Rows[_headerRowIndex][c]?.ToString()?.Trim()?.ToLowerInvariant() ?? ""
                    : "";

                if (string.IsNullOrEmpty(remoteName) || remoteName != localName) continue;
                if (idKeywords.Any(k => remoteName.Contains(k)))
                {
                    _idColumnIndex = c;
                    return true;
                }
            }

            _idColumnIndex = -1;
            return false;
        }

        // ==================== 差异高亮绑定 ====================
        private void ApplyDiffHighlight()
        {
            if (_remoteView != null)
            {
                _remoteView.OnExternalCellFormatting -= OnDiffCellFormatting;
                _remoteView.OnExternalCellFormatting += OnDiffCellFormatting;
            }
            if (_localView != null)
            {
                _localView.OnExternalCellFormatting -= OnDiffCellFormatting;
                _localView.OnExternalCellFormatting += OnDiffCellFormatting;
            }

            if (_remoteView != null)
            {
                var dgv = _remoteView.GetDataGridView();
                if (dgv != null)
                {
                    dgv.RowHeaderMouseClick -= OnRemoteRowHeaderClick;
                    dgv.RowHeaderMouseClick += OnRemoteRowHeaderClick;
                    dgv.CellMouseClick -= OnRemoteCellClick;
                    dgv.CellMouseClick += OnRemoteCellClick;
                }
                // 通过 OnExternalCellMouseDown 统一处理右键菜单（ContextMenuMode.Conflict 模式）
                _remoteView.OnExternalCellMouseDown -= OnConflictGridCellMouseDown;
                _remoteView.OnExternalCellMouseDown += OnConflictGridCellMouseDown;
            }
            if (_localView != null)
            {
                var dgv = _localView.GetDataGridView();
                if (dgv != null)
                {
                    dgv.RowHeaderMouseClick -= OnLocalRowHeaderClick;
                    dgv.RowHeaderMouseClick += OnLocalRowHeaderClick;
                    dgv.CellMouseClick -= OnLocalCellClick;
                    dgv.CellMouseClick += OnLocalCellClick;
                }
                // 通过 OnExternalCellMouseDown 统一处理右键菜单（ContextMenuMode.Conflict 模式）
                _localView.OnExternalCellMouseDown -= OnConflictGridCellMouseDown;
                _localView.OnExternalCellMouseDown += OnConflictGridCellMouseDown;
            }

            // 绑定合并视图右键菜单（通过 OnExternalCellMouseDown）
            if (_mergedView != null)
            {
                _mergedView.OnExternalCellMouseDown -= OnConflictGridCellMouseDown;
                _mergedView.OnExternalCellMouseDown += OnConflictGridCellMouseDown;
            }

            // 绑定滚动同步
            SetupSyncScroll();
        }

        // ==================== 滚动同步（三个视图联动） ====================
        private void SetupSyncScroll()
        {
            var remoteDgv = _remoteView?.GetDataGridView();
            var localDgv = _localView?.GetDataGridView();
            var mergedDgv = _mergedView?.GetDataGridView();

            // 先移除再注册，防止重复注册
            if (remoteDgv != null)
            {
                remoteDgv.Scroll -= OnRemoteGridScroll;
                remoteDgv.Scroll += OnRemoteGridScroll;
            }
            if (localDgv != null)
            {
                localDgv.Scroll -= OnLocalGridScroll;
                localDgv.Scroll += OnLocalGridScroll;
            }
            if (mergedDgv != null)
            {
                mergedDgv.Scroll -= OnMergedGridScroll;
                mergedDgv.Scroll += OnMergedGridScroll;
            }
        }

        /// <summary>将 sourceDgv 的滚动状态同步到其他两个 DataGridView</summary>
        private void SyncScrollToOthers(DataGridView sourceDgv, ScrollEventArgs e)
        {
            var remoteDgv = _remoteView?.GetDataGridView();
            var localDgv = _localView?.GetDataGridView();
            var mergedDgv = _mergedView?.GetDataGridView();

            var targets = new[]
            {
                (dgv: remoteDgv, isRemote: true),
                (dgv: localDgv, isRemote: false),
                (dgv: mergedDgv, isRemote: false)
            };

            _syncingScroll = true;
            try
            {
                foreach (var (dgv, _) in targets)
                {
                    if (dgv == null || dgv == sourceDgv) continue;

                    if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
                    {
                        // 垂直滚动：仅远程和本地视图行数对齐（通过 BuildAlignedTables），同步有意义
                        // 合并视图行数可能不同，不同步垂直滚动
                        if (dgv == remoteDgv || dgv == localDgv)
                        {
                            int idx = sourceDgv.FirstDisplayedScrollingRowIndex;
                            if (dgv.RowCount > 0)
                                dgv.FirstDisplayedScrollingRowIndex = Math.Min(idx, dgv.RowCount - 1);
                        }
                    }
                    else if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                    {
                        // 水平滚动：三个表列结构基本一致，全部同步
                        dgv.HorizontalScrollingOffset = sourceDgv.HorizontalScrollingOffset;
                    }
                }
            }
            finally
            {
                _syncingScroll = false;
            }
        }

        private void OnRemoteGridScroll(object? sender, ScrollEventArgs e)
        {
            if (_syncingScroll) return;
            var dgv = _remoteView?.GetDataGridView();
            if (dgv == null) return;
            SyncScrollToOthers(dgv, e);
        }

        private void OnLocalGridScroll(object? sender, ScrollEventArgs e)
        {
            if (_syncingScroll) return;
            var dgv = _localView?.GetDataGridView();
            if (dgv == null) return;
            SyncScrollToOthers(dgv, e);
        }

        private void OnMergedGridScroll(object? sender, ScrollEventArgs e)
        {
            if (_syncingScroll) return;
            var dgv = _mergedView?.GetDataGridView();
            if (dgv == null) return;
            SyncScrollToOthers(dgv, e);
        }

        // ==================== 右键菜单（冲突解决模式，统一处理远程/本地/合并视图） ====================
        private void OnConflictGridCellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right || e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (sender is not DataGridView dgv) return;

            bool isRemote = dgv == _remoteView?.GetDataGridView();
            bool isLocal = dgv == _localView?.GetDataGridView();
            bool isMerged = dgv == _mergedView?.GetDataGridView();

            if (isRemote || isLocal)
            {
                // === 远程/本地视图右键菜单 ===
                var rowProps = isRemote ? _remoteRowProps : _localRowProps;
                var rowMap = isRemote ? _remoteRowToOriginal : _localRowToOriginal;

                if (e.RowIndex >= rowProps.Count) return;
                if (e.RowIndex >= rowMap.Count) return;
                if (rowProps[e.RowIndex].isBlankRow) return;

                int originalRow = rowMap[e.RowIndex];
                if (originalRow < _frozenRows && originalRow >= 0) return;

                // 选中该行/单元格
                if (isRemote)
                    OnRemoteCellClick(sender, e);
                else
                    OnLocalCellClick(sender, e);

                ShowConflictContextMenu(dgv, isRemote, originalRow, e.ColumnIndex);
            }
            else if (isMerged)
            {
                // === 合并视图右键菜单（编辑操作） ===
                ShowMergedContextMenu(dgv, e.RowIndex, e.ColumnIndex);
            }
        }

        private void ShowMergedContextMenu(DataGridView dgv, int rowIndex, int colIndex)
        {
            var cellValue = dgv.Rows[rowIndex].Cells[colIndex].Value?.ToString() ?? "";
            var cms = new ContextMenuStrip();
            cms.Font = new Font("Microsoft YaHei UI", 9f);

            var copyItem = new ToolStripMenuItem("复制");
            copyItem.Click += (s, ev) => { if (!string.IsNullOrEmpty(cellValue)) Clipboard.SetText(cellValue); };
            cms.Items.Add(copyItem);

            var editCellItem = new ToolStripMenuItem("编辑单元格");
            editCellItem.Click += (s, ev) =>
            {
                dgv.CurrentCell = dgv.Rows[rowIndex].Cells[colIndex];
                dgv.BeginEdit(true);
            };
            cms.Items.Add(editCellItem);

            cms.Items.Add(new ToolStripSeparator());

            var deleteRowItem = new ToolStripMenuItem("删除行");
            deleteRowItem.ForeColor = Color.FromArgb(200, 50, 50);
            deleteRowItem.Click += (s, ev) =>
            {
                if (rowIndex >= 0 && rowIndex < _mergedTable.Rows.Count)
                {
                    _mergedTable.Rows.RemoveAt(rowIndex);
                    RefreshMergedView();
                    _lblInfo.Text = $"已从合并结果中删除行 {rowIndex + 1}";
                }
            };
            cms.Items.Add(deleteRowItem);

            cms.Show(dgv, dgv.PointToClient(Cursor.Position));
        }

        private void ShowConflictContextMenu(DataGridView dgv, bool isRemote, int originalRow, int colIndex)
        {
            var cms = new ContextMenuStrip();
            cms.Font = new Font("Microsoft YaHei UI", 9f);

            var useCellItem = new ToolStripMenuItem("使用此单元格");
            useCellItem.Click += (s, ev) =>
            {
                if (isRemote)
                {
                    _selectedRemoteCell = (originalRow, colIndex);
                    _selectedRemoteRows.Clear();
                    UseRemoteSelected();
                }
                else
                {
                    _selectedLocalCell = (originalRow, colIndex);
                    _selectedLocalRows.Clear();
                    UseLocalSelected();
                }
            };
            cms.Items.Add(useCellItem);

            var useRowItem = new ToolStripMenuItem("使用此行");
            useRowItem.Click += (s, ev) =>
            {
                if (isRemote)
                {
                    _selectedRemoteRows.Clear();
                    _selectedRemoteRows.Add(originalRow);
                    _selectedRemoteCell = null;
                    UseRemoteSelected();
                }
                else
                {
                    _selectedLocalRows.Clear();
                    _selectedLocalRows.Add(originalRow);
                    _selectedLocalCell = null;
                    UseLocalSelected();
                }
            };
            cms.Items.Add(useRowItem);

            cms.Items.Add(new ToolStripSeparator());

            var deleteRowItem = new ToolStripMenuItem("删除此行");
            deleteRowItem.ForeColor = Color.FromArgb(200, 50, 50);
            deleteRowItem.Click += (s, ev) =>
            {
                DeleteMergedRow(originalRow, isRemote);
            };
            cms.Items.Add(deleteRowItem);

            cms.Show(dgv, dgv.PointToClient(Cursor.Position));
        }

        /// <summary>从合并结果中删除指定行</summary>
        private void DeleteMergedRow(int originalRow, bool isRemote)
        {
            var sourceTable = isRemote ? _remoteData : _localData;
            int mergedRow = FindMergedRowById(originalRow, sourceTable);

            if (mergedRow < 0)
            {
                // 无ID匹配 → 用行索引回退
                if (originalRow < _mergedTable.Rows.Count)
                    mergedRow = originalRow;
                else
                {
                    AlertForm.Show(ParentForm, "无法在合并结果中找到对应的行。", "提示", AlertType.Warning);
                    return;
                }
            }

            if (mergedRow >= 0 && mergedRow < _mergedTable.Rows.Count)
            {
                _mergedTable.Rows.RemoveAt(mergedRow);
                RefreshMergedView();
                _lblInfo.Text = $"已从合并结果中删除行 (原始行 {originalRow})";
            }
        }

        private void OnDiffCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (sender is not DataGridView dgv) return;
            if (e.RowIndex < 0) return;

            bool isRemote = dgv == _remoteView?.GetDataGridView();
            var rowProps = isRemote ? _remoteRowProps : _localRowProps;
            var rowMap = isRemote ? _remoteRowToOriginal : _localRowToOriginal;
            var diffCells = isRemote ? _remoteDiffCells : _localDiffCells;
            var diffRows = isRemote ? _remoteDiffRows : _localDiffRows;

            if (e.RowIndex >= rowProps.Count) return;
            var (diffType, isBlankRow) = rowProps[e.RowIndex];

            // ★ 空白行 → 根据对面行类型选择背景色
            if (isBlankRow)
            {
                if (diffType == DiffRowType.Added)
                {
                    // 本地空白行（对面远程是新增行）
                    e.CellStyle.BackColor = _colorBlankRowForAdded;
                    e.CellStyle.SelectionBackColor = Color.FromArgb(210, 240, 225);
                }
                else if (diffType == DiffRowType.Deleted)
                {
                    // 远程空白行（对面本地是删除行）
                    e.CellStyle.BackColor = _colorBlankRowForDeleted;
                    e.CellStyle.SelectionBackColor = Color.FromArgb(245, 225, 220);
                }
                else
                {
                    e.CellStyle.BackColor = Color.FromArgb(245, 245, 245);
                    e.CellStyle.SelectionBackColor = Color.FromArgb(230, 230, 230);
                }
                e.FormattingApplied = true;
                return;
            }

            // ★ 新增行（只在远程存在）
            if (diffType == DiffRowType.Added)
            {
                e.CellStyle.BackColor = _colorAddedRowBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(190, 230, 200);
                e.FormattingApplied = true;
                return;
            }

            // ★ 删除行（只在本地存在）
            if (diffType == DiffRowType.Deleted)
            {
                e.CellStyle.BackColor = _colorDeletedRowBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(240, 200, 190);
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
            bool isConflict = _conflictCells.Contains(cellKey);

            if (isConflict)
            {
                e.CellStyle.BackColor = _colorConflictBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(255, 140, 140);
            }
            else if (isDiffCell)
            {
                e.CellStyle.BackColor = _colorDiffCellBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(255, 170, 170);
            }
            else
            {
                e.CellStyle.BackColor = _colorDiffRowBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(240, 230, 180);
            }

            e.FormattingApplied = true;
        }

        // ==================== 行选择处理（使用对齐后的行索引） ====================
        private void OnRemoteRowHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (e.RowIndex >= _remoteRowToOriginal.Count) return;
            int originalRow = _remoteRowToOriginal[e.RowIndex];
            if (originalRow < _frozenRows && originalRow >= 0) return;

            // 检查是否是空白行（对面新增，本地空白）
            if (e.RowIndex < _remoteRowProps.Count && _remoteRowProps[e.RowIndex].isBlankRow) return;

            if (ModifierKeys == Keys.Control)
            {
                if (!_selectedRemoteRows.Remove(originalRow))
                    _selectedRemoteRows.Add(originalRow);
            }
            else if (ModifierKeys == Keys.Shift)
            {
                if (_selectedRemoteRows.Count > 0)
                {
                    int last = _selectedRemoteRows.Last();
                    int min = Math.Min(last, originalRow);
                    int max = Math.Max(last, originalRow);
                    for (int r = min; r <= max; r++)
                        if (r >= 0) _selectedRemoteRows.Add(r);
                }
                else
                    _selectedRemoteRows.Add(originalRow);
            }
            else
            {
                _selectedRemoteRows.Clear();
                _selectedRemoteRows.Add(originalRow);
            }

            _selectionFromRemote = true;
            _selectedRemoteCell = null;
            HighlightSelectedRows();
            UpdateInfoLabel();
        }

        private void OnLocalRowHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (e.RowIndex >= _localRowToOriginal.Count) return;
            int originalRow = _localRowToOriginal[e.RowIndex];
            if (originalRow < _frozenRows && originalRow >= 0) return;

            // 检查是否是空白行（对面删除，远程空白）
            if (e.RowIndex < _localRowProps.Count && _localRowProps[e.RowIndex].isBlankRow) return;

            if (ModifierKeys == Keys.Control)
            {
                if (!_selectedLocalRows.Remove(originalRow))
                    _selectedLocalRows.Add(originalRow);
            }
            else if (ModifierKeys == Keys.Shift)
            {
                if (_selectedLocalRows.Count > 0)
                {
                    int last = _selectedLocalRows.Last();
                    int min = Math.Min(last, originalRow);
                    int max = Math.Max(last, originalRow);
                    for (int r = min; r <= max; r++)
                        if (r >= 0) _selectedLocalRows.Add(r);
                }
                else
                    _selectedLocalRows.Add(originalRow);
            }
            else
            {
                _selectedLocalRows.Clear();
                _selectedLocalRows.Add(originalRow);
            }

            _selectionFromRemote = false;
            _selectedLocalCell = null;
            HighlightSelectedRows();
            UpdateInfoLabel();
        }

        private void OnRemoteCellClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (e.RowIndex >= _remoteRowToOriginal.Count) return;
            if (e.RowIndex < _remoteRowProps.Count && _remoteRowProps[e.RowIndex].isBlankRow) return;

            int originalRow = _remoteRowToOriginal[e.RowIndex];
            if (originalRow < _frozenRows && originalRow >= 0) return;

            _selectedRemoteRows.Clear();
            _selectedRemoteRows.Add(originalRow);
            _selectedRemoteCell = (originalRow, e.ColumnIndex);
            _selectionFromRemote = true;
            _selectedLocalCell = null;
            _selectedLocalRows.Clear();

            HighlightSelectedRows();
            UpdateInfoLabel();
        }

        private void OnLocalCellClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (e.RowIndex >= _localRowToOriginal.Count) return;
            if (e.RowIndex < _localRowProps.Count && _localRowProps[e.RowIndex].isBlankRow) return;

            int originalRow = _localRowToOriginal[e.RowIndex];
            if (originalRow < _frozenRows && originalRow >= 0) return;

            _selectedLocalRows.Clear();
            _selectedLocalRows.Add(originalRow);
            _selectedLocalCell = (originalRow, e.ColumnIndex);
            _selectionFromRemote = false;
            _selectedRemoteCell = null;
            _selectedRemoteRows.Clear();

            HighlightSelectedRows();
            UpdateInfoLabel();
        }

        private void HighlightSelectedRows()
        {
            if (_remoteView != null)
            {
                var dgv = _remoteView.GetDataGridView();
                if (dgv != null) dgv.Invalidate();
            }
            if (_localView != null)
            {
                var dgv = _localView.GetDataGridView();
                if (dgv != null) dgv.Invalidate();
            }
        }

        private void UpdateInfoLabel()
        {
            if (_selectionFromRemote)
            {
                if (_selectedRemoteCell.HasValue)
                    _lblInfo.Text = $"已选择远程单元格 [行{_selectedRemoteCell.Value.row}, 列{_selectedRemoteCell.Value.col}]";
                else if (_selectedRemoteRows.Count > 0)
                    _lblInfo.Text = $"已选择远程 {_selectedRemoteRows.Count} 行";
            }
            else
            {
                if (_selectedLocalCell.HasValue)
                    _lblInfo.Text = $"已选择本地单元格 [行{_selectedLocalCell.Value.row}, 列{_selectedLocalCell.Value.col}]";
                else if (_selectedLocalRows.Count > 0)
                    _lblInfo.Text = $"已选择本地 {_selectedLocalRows.Count} 行";
            }
        }

        // ==================== 合并操作 ====================
        /// <summary>在合并表中查找与原始表中指定行具有相同ID的行索引</summary>
        private int FindMergedRowById(int originalRow, DataTable originalTable)
        {
            if (_idColumnIndex < 0) return -1;
            if (_idColumnIndex >= originalTable.Columns.Count) return -1;
            if (_idColumnIndex >= _mergedTable.Columns.Count) return -1;

            string id = originalTable.Rows[originalRow][_idColumnIndex]?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(id)) return -1;

            for (int mr = 0; mr < _mergedTable.Rows.Count; mr++)
            {
                string mergedId = _mergedTable.Rows[mr][_idColumnIndex]?.ToString()?.Trim() ?? "";
                if (string.Equals(id, mergedId, StringComparison.OrdinalIgnoreCase))
                    return mr;
            }
            return -1;
        }

        private void UseRemoteSelected()
        {
            if (_selectedRemoteCell.HasValue)
            {
                var (row, col) = _selectedRemoteCell.Value;
                int mergedRow = FindMergedRowById(row, _remoteData);
                if (mergedRow < 0) mergedRow = row; // 找不到则用行索引回退
                if (row >= 0 && row < _remoteData.Rows.Count && col < _remoteData.Columns.Count &&
                    mergedRow >= 0 && mergedRow < _mergedTable.Rows.Count && col < _mergedTable.Columns.Count)
                {
                    _mergedTable.Rows[mergedRow][col] = _remoteData.Rows[row][col]?.ToString() ?? "";
                }
                _selectedRemoteCell = null;
            }
            else if (_selectedRemoteRows.Count > 0)
            {
                foreach (var row in _selectedRemoteRows.OrderBy(r => r))
                {
                    if (row < 0 || row >= _remoteData.Rows.Count) continue;

                    // ★ 通过 ID 在合并表中查找对应行（不依赖 _useIdAlign）
                    int mergedRow = FindMergedRowById(row, _remoteData);

                    if (mergedRow >= 0)
                    {
                        // 找到匹配行 → 覆盖
                        for (int c = 0; c < Math.Min(_remoteData.Columns.Count, _mergedTable.Columns.Count); c++)
                            _mergedTable.Rows[mergedRow][c] = _remoteData.Rows[row][c]?.ToString() ?? "";
                    }
                    else if (row < _mergedTable.Rows.Count)
                    {
                        // 无ID匹配，但行索引在范围内 → 用行索引覆盖
                        for (int c = 0; c < Math.Min(_remoteData.Columns.Count, _mergedTable.Columns.Count); c++)
                            _mergedTable.Rows[row][c] = _remoteData.Rows[row][c]?.ToString() ?? "";
                    }
                    else
                    {
                        // 超出合并表行数 → 追加新行
                        var newRow = _mergedTable.NewRow();
                        for (int c = 0; c < Math.Min(_remoteData.Columns.Count, _mergedTable.Columns.Count); c++)
                            newRow[c] = _remoteData.Rows[row][c]?.ToString() ?? "";
                        _mergedTable.Rows.Add(newRow);
                    }
                }
                _selectedRemoteRows.Clear();
            }

            RefreshMergedView();
            _lblInfo.Text = "已应用远程版本选择内容到合并结果";
        }

        private void UseLocalSelected()
        {
            if (_selectedLocalCell.HasValue)
            {
                var (row, col) = _selectedLocalCell.Value;
                int mergedRow = FindMergedRowById(row, _localData);
                if (mergedRow < 0) mergedRow = row;
                if (row >= 0 && row < _localData.Rows.Count && col < _localData.Columns.Count &&
                    mergedRow >= 0 && mergedRow < _mergedTable.Rows.Count && col < _mergedTable.Columns.Count)
                {
                    _mergedTable.Rows[mergedRow][col] = _localData.Rows[row][col]?.ToString() ?? "";
                }
                _selectedLocalCell = null;
            }
            else if (_selectedLocalRows.Count > 0)
            {
                foreach (var row in _selectedLocalRows.OrderBy(r => r))
                {
                    if (row < 0 || row >= _localData.Rows.Count) continue;

                    int mergedRow = FindMergedRowById(row, _localData);

                    if (mergedRow >= 0)
                    {
                        for (int c = 0; c < Math.Min(_localData.Columns.Count, _mergedTable.Columns.Count); c++)
                            _mergedTable.Rows[mergedRow][c] = _localData.Rows[row][c]?.ToString() ?? "";
                    }
                    else if (row < _mergedTable.Rows.Count)
                    {
                        for (int c = 0; c < Math.Min(_localData.Columns.Count, _mergedTable.Columns.Count); c++)
                            _mergedTable.Rows[row][c] = _localData.Rows[row][c]?.ToString() ?? "";
                    }
                    else
                    {
                        var newRow = _mergedTable.NewRow();
                        for (int c = 0; c < Math.Min(_localData.Columns.Count, _mergedTable.Columns.Count); c++)
                            newRow[c] = _localData.Rows[row][c]?.ToString() ?? "";
                        _mergedTable.Rows.Add(newRow);
                    }
                }
                _selectedLocalRows.Clear();
            }

            RefreshMergedView();
            _lblInfo.Text = "已应用本地版本选择内容到合并结果";
        }

        private void UseAllRemote()
        {
            // ★ 直接重建合并表为远程版本副本（"全部使用远程" = 合并结果等于远程版本）
            _mergedTable = _remoteData.Copy();
            RefreshMergedView();
            _lblInfo.Text = $"已全部使用远程版本 ({_remoteData.Rows.Count} 行)";
        }

        private void UseAllLocal()
        {
            // ★ 直接重建合并表为本地版本副本（"全部使用本地" = 合并结果等于本地版本）
            _mergedTable = _localData.Copy();
            RefreshMergedView();
            _lblInfo.Text = $"已全部使用本地版本 ({_localData.Rows.Count} 行)";
        }

        private void SelectAllRemoteRows()
        {
            _selectedRemoteRows.Clear();
            int dataStartRow = _frozenRows;
            foreach (int r in _remoteDiffRows)
            {
                if (r >= dataStartRow)
                    _selectedRemoteRows.Add(r);
            }
            _selectionFromRemote = true;
            _selectedRemoteCell = null;
            HighlightSelectedRows();
            _lblInfo.Text = $"已全选远程 {_selectedRemoteRows.Count} 差异行";
        }

        private void SelectAllLocalRows()
        {
            _selectedLocalRows.Clear();
            int dataStartRow = _frozenRows;
            foreach (int r in _localDiffRows)
            {
                if (r >= dataStartRow)
                    _selectedLocalRows.Add(r);
            }
            _selectionFromRemote = false;
            _selectedLocalCell = null;
            HighlightSelectedRows();
            _lblInfo.Text = $"已全选本地 {_selectedLocalRows.Count} 差异行";
        }

        private void RefreshMergedView()
        {
            _mergedView.SetDataTable(_mergedTable, $"合并结果 | 行数: {_mergedTable.Rows.Count}", BuildMergedFileData());
            _mergedView.SetEmbeddedMode(false);
        }

        /// <summary>
        /// 为合并表构建带源文件物理行号的 FileTableData。
        /// 合并表会被追加/排序打乱行序，所以按 ID 回查本地版本拿行号：
        /// 本地存在的行显示本地物理行号，远程独有的行没有本地行号（画 "—"）。
        /// 两个版本都没有文件解析结果时返回 null，退回视图自身的序号口径。
        /// </summary>
        private FileTableData? BuildMergedFileData()
        {
            if (_mergedTable == null) return null;
            if (_localFileData == null && _remoteFileData == null) return null;

            // ID → 本地物理行号
            var localLineById = new Dictionary<string, (int Start, int End)>(StringComparer.OrdinalIgnoreCase);
            int idCol = _idColumnIndex;
            if (_localFileData != null && idCol >= 0)
            {
                for (int i = _frozenRows; i < _localFileData.Rows.Count; i++)
                {
                    var src = _localFileData.Rows[i];
                    if (src.PhysicalStartLine <= 0) continue;
                    string? id = idCol < src.Cells.Count ? src.Cells[idCol]?.Trim() : null;
                    if (!string.IsNullOrEmpty(id) && !localLineById.ContainsKey(id))
                        localLineById[id] = (src.PhysicalStartLine, src.PhysicalEndLine);
                }
            }

            // 固定行（表头/描述行）不参与对齐，两个版本里都排在文件开头，按索引取行号即可
            var frozenSource = _localFileData ?? _remoteFileData;

            var fd = new FileTableData();
            foreach (DataColumn col in _mergedTable.Columns)
                fd.ColumnNames.Add(col.ColumnName);

            for (int r = 0; r < _mergedTable.Rows.Count; r++)
            {
                var row = new TableRow { OriginalRowIndex = r };
                for (int c = 0; c < _mergedTable.Columns.Count; c++)
                    row.Cells.Add(_mergedTable.Rows[r][c]?.ToString() ?? "");

                if (r < _frozenRows)
                {
                    if (frozenSource != null && r < frozenSource.Rows.Count)
                    {
                        row.PhysicalStartLine = frozenSource.Rows[r].PhysicalStartLine;
                        row.PhysicalEndLine = frozenSource.Rows[r].PhysicalEndLine;
                    }
                }
                else if (localLineById.Count > 0)
                {
                    string id = _mergedTable.Rows[r][idCol]?.ToString()?.Trim() ?? "";
                    if (localLineById.TryGetValue(id, out var line))
                    {
                        row.PhysicalStartLine = line.Start;
                        row.PhysicalEndLine = line.End;
                    }
                    else
                    {
                        // 远程独有行：本地没有对应行号，标记成占位行让行号列画 "—"
                        row.IsBlankRow = true;
                    }
                }
                fd.Rows.Add(row);
            }
            return fd;
        }

        // ==================== 合并结果表格初始化（包含所有行） ====================
        private void InitializeMergedTable()
        {
            if (_localData == null && _remoteData == null) return;

            // ★ 合并表以本地为基础
            _mergedTable = _localData != null ? _localData.Copy() : _remoteData!.Copy();

            if (_remoteData == null) { RefreshMergedView(); return; }

            // ★ 追加远程新增行（本地无对应ID的行）
            if (_useIdAlign && _idColumnIndex >= 0)
            {
                foreach (int r in _remoteNewRowOriginalIndices.OrderBy(r => r))
                {
                    if (r >= 0 && r < _remoteData.Rows.Count)
                    {
                        var newRow = _mergedTable.NewRow();
                        for (int c = 0; c < Math.Min(_remoteData.Columns.Count, _mergedTable.Columns.Count); c++)
                            newRow[c] = _remoteData.Rows[r][c]?.ToString() ?? "";
                        _mergedTable.Rows.Add(newRow);
                    }
                }
            }
            else if (_idColumnIndex >= 0)
            {
                // 无ID对齐模式但ID列可用：按ID匹配追加远程独有行
                var localIdSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int idCol = _idColumnIndex;
                for (int r = 0; r < _mergedTable.Rows.Count; r++)
                {
                    string id = _mergedTable.Rows[r][idCol]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(id)) localIdSet.Add(id);
                }
                for (int r = 0; r < _remoteData.Rows.Count; r++)
                {
                    string id = _remoteData.Rows[r][idCol]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(id) && !localIdSet.Contains(id))
                    {
                        var newRow = _mergedTable.NewRow();
                        for (int c = 0; c < Math.Min(_remoteData.Columns.Count, _mergedTable.Columns.Count); c++)
                            newRow[c] = _remoteData.Rows[r][c]?.ToString() ?? "";
                        _mergedTable.Rows.Add(newRow);
                        localIdSet.Add(id);
                    }
                }
            }
            else
            {
                // 无ID列：按行号追加远程比本地多的行
                if (_localData != null && _remoteData.Rows.Count > _localData.Rows.Count)
                {
                    for (int r = _localData.Rows.Count; r < _remoteData.Rows.Count; r++)
                    {
                        var newRow = _mergedTable.NewRow();
                        for (int c = 0; c < Math.Min(_remoteData.Columns.Count, _mergedTable.Columns.Count); c++)
                            newRow[c] = _remoteData.Rows[r][c]?.ToString() ?? "";
                        _mergedTable.Rows.Add(newRow);
                    }
                }
            }

            // ★ 对数据行按 ID 排序，确保合并表行顺序与 ID 排序一致（跳过固定行）
            if (_idColumnIndex >= 0 && _idColumnIndex < _mergedTable.Columns.Count && _mergedTable.Rows.Count > 1)
            {
                try
                {
                    int dataStartRow = _frozenRows;
                    if (dataStartRow > 0 && dataStartRow < _mergedTable.Rows.Count)
                    {
                        var fixedRows = _mergedTable.AsEnumerable().Take(dataStartRow).ToList();
                        var dataRows = _mergedTable.AsEnumerable().Skip(dataStartRow)
                            .OrderBy(r =>
                            {
                                var idStr = r[_idColumnIndex]?.ToString()?.Trim() ?? "";
                                return int.TryParse(idStr, out int id) ? id : int.MaxValue;
                            })
                            .ThenBy(r =>
                            {
                                var idStr = r[_idColumnIndex]?.ToString()?.Trim() ?? "";
                                return idStr;
                            })
                            .ToList();

                        if (dataRows.Count > 0)
                        {
                            _mergedTable = fixedRows.Concat(dataRows).CopyToDataTable();
                        }
                    }
                }
                catch
                {
                    // 排序失败不阻塞，使用未排序的 _mergedTable
                }
            }

            RefreshMergedView();
        }

        // ==================== 合并结果编辑 ====================
        private void OnMergedCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (_remoteData == null || _localData == null) return;

            int originalRow = e.RowIndex + _frozenRows;
            int col = e.ColumnIndex;

            int cellKey = (originalRow << 16) | col;
            if (_conflictCells.Contains(cellKey))
            {
                e.CellStyle.BackColor = _colorConflictBg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(255, 140, 140);
                e.FormattingApplied = true;
                return;
            }

            bool isRemoteDiff = _remoteDiffRows.Contains(originalRow);
            bool isLocalDiff = _localDiffRows.Contains(originalRow);

            if (isRemoteDiff || isLocalDiff)
            {
                bool isDiffCell = _remoteDiffCells.Contains(cellKey) || _localDiffCells.Contains(cellKey);
                if (isDiffCell)
                {
                    e.CellStyle.BackColor = Color.FromArgb(255, 240, 200);
                    e.CellStyle.SelectionBackColor = Color.FromArgb(255, 220, 180);
                }
                else
                {
                    e.CellStyle.BackColor = Color.FromArgb(255, 250, 235);
                }
                e.FormattingApplied = true;
            }
        }

        private void OnMergedCellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            int originalRow = e.RowIndex + _frozenRows;
            int cellKey = (originalRow << 16) | e.ColumnIndex;
            if (_conflictCells.Contains(cellKey))
                _conflictCells.Remove(cellKey);
            var dgv = _mergedView.GetDataGridView();
            dgv.Invalidate();
            _lblInfo.Text = "合并结果已手动编辑";
        }

        private void OnMergedGridKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.S)
            {
                SaveMergedResult();
                e.Handled = true;
            }
        }

        // ==================== 保存 ====================
        private async void SaveMergedResult()
        {
            try
            {
                if (_mergedTable == null || _mergedTable.Rows.Count == 0)
                {
                    AlertForm.Show(ParentForm, "合并结果为空，无法保存。", "提示", AlertType.Warning);
                    return;
                }

                // 备份到 %temp%/wxd-bak/ 目录，带版本号
                string backupPath = "";
                try
                {
                    if (File.Exists(_filePath))
                    {
                        var bakDir = Path.Combine(Path.GetTempPath(), "wxd-bak");
                        Directory.CreateDirectory(bakDir);
                        var fileName = Path.GetFileNameWithoutExtension(_filePath);
                        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                        // 获取 SVN/Git 版本号
                        string vcsRevision = "";
                        try
                        {
                            var rootDir = Path.GetDirectoryName(_filePath) ?? "";
                            if (_vcsType == VcsType.Svn)
                            {
                                var psi = new ProcessStartInfo
                                {
                                    FileName = "svn",
                                    Arguments = $"info --show-item revision \"{rootDir}\"",
                                    RedirectStandardOutput = true,
                                    RedirectStandardError = true,
                                    UseShellExecute = false,
                                    CreateNoWindow = true
                                };
                                using var proc = Process.Start(psi);
                                if (proc != null)
                                {
                                    vcsRevision = (await proc.StandardOutput.ReadToEndAsync()).Trim();
                                    await proc.WaitForExitAsync();
                                }
                            }
                            else if (_vcsType == VcsType.Git)
                            {
                                var psi = new ProcessStartInfo
                                {
                                    FileName = "git",
                                    Arguments = $"rev-parse --short HEAD",
                                    WorkingDirectory = rootDir,
                                    RedirectStandardOutput = true,
                                    RedirectStandardError = true,
                                    UseShellExecute = false,
                                    CreateNoWindow = true
                                };
                                using var proc = Process.Start(psi);
                                if (proc != null)
                                {
                                    vcsRevision = (await proc.StandardOutput.ReadToEndAsync()).Trim();
                                    await proc.WaitForExitAsync();
                                }
                            }
                        }
                        catch { }

                        var revisionTag = string.IsNullOrEmpty(vcsRevision) ? "" : $"_{vcsRevision}";

                        backupPath = Path.Combine(bakDir, $"{fileName}_backup{revisionTag}_{timestamp}.bak");
                        File.Copy(_filePath, backupPath, true);
                    }
                }
                catch { }

                // 删除原文件（备份已完成），确保新文件可以覆盖写入
                try { if (File.Exists(_filePath)) File.Delete(_filePath); } catch { }

                // 根据文件扩展名选择保存格式
                FileDataCache.SaveFileDirect(_filePath, _mergedTable);

                // 清除 VCS 冲突标记
                try
                {
                    VersionControlHelper.MarkConflictResolved(_filePath, _vcsType);
                }
                catch { }

                var backupInfo = string.IsNullOrEmpty(backupPath) ? "" : $" (备份: {Path.GetFileName(backupPath)})";
                _lblInfo.Text = $"已保存到 {Path.GetFileName(_filePath)}{backupInfo}";
                AlertForm.Show(ParentForm, "合并结果已保存。" + GlobalConst.NewLine + "原文件已备份到 wxd-bak 临时文件夹。", "保存成功", AlertType.Information);

                // 触发保存完成事件
                SaveCompleted?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                AlertForm.Show(ParentForm, $"保存失败: {ex.Message}", "错误", AlertType.Error, exception: ex);
            }
        }

        private void SaveMergedResultAs()
        {
            using var dialog = new SaveFileDialog
            {
                Title = "另存为合并结果",
                Filter = "CSV 文件 (*.csv)|*.csv|Excel 文件 (*.xlsx)|*.xlsx",
                FileName = Path.GetFileNameWithoutExtension(_filePath) + "_merged.csv"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var delimiter = ',';
                    var ext = Path.GetExtension(dialog.FileName).ToLowerInvariant();
                    if (ext == ".csv")
                    {
                        if (!FileDataCache.SaveFileDirect(dialog.FileName, _mergedTable, delimiter))
                            throw new IOException("写入CSV文件失败");
                    }
                    else
                    {
                        if (!FileDataCache.SaveFileDirect(dialog.FileName, _mergedTable))
                            throw new IOException("写入Excel文件失败");
                    }

                    _lblInfo.Text = $"已另存为: {dialog.FileName}";
                    AlertForm.Show(ParentForm, "合并结果已另存。", "保存成功", AlertType.Information);
                }
                catch (Exception ex)
                {
                    AlertForm.Show(ParentForm, $"保存失败: {ex.Message}", "错误", AlertType.Error, exception: ex);
                }
            }
        }

        /// <summary>加载配置文件，同时返回带源文件物理行号的解析结果（冲突解决入口用）</summary>
        public static (DataTable? Data, FileTableData? FullData) LoadDataTableFromFile(string filePath)
        {
            if (!File.Exists(filePath)) return (null, null);

            try
            {
                var (dataTable, fileData) = FileDataCache.LoadFileDirect(filePath);
                return (dataTable, fileData);
            }
            catch { return (null, null); }
        }
    }
}