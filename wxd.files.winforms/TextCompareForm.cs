using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 文本 / 表格比对：左右各选一个文件，按行做差分。
    ///
    /// 两种模式自动切换（看扩展名）：
    ///   文本（txt/xml/yaml/json/ini/log…）：按行差分，网格是"行号 + 整行内容"；
    ///   表格（csv/xlsx/xls）：按行号一一对齐、逐单元格比较，网格是"行号 + A/B/C… 列"。
    ///     表格模式不认表头行、不做冻结行，也不复用阅读器的任何界面逻辑。
    ///
    /// 两边网格绑定同一份"按行对齐"的结果（<see cref="TextDiffEngine"/> / <see cref="TableDiffEngine"/>），
    /// 所以行数天然一致、滚动同步不会错位：
    ///   修改 = 黄色（表格模式精确到单元格），左侧独有 = 红色，右侧独有 = 绿色，另一侧的空位 = 灰色。
    /// 勾选"只保留差异行"后只显示有差异的行，行号仍是原文件行号。
    /// </summary>
    internal sealed class TextCompareForm : Form
    {
        private readonly TableLayoutPanel _top = new();
        private readonly FlowLayoutPanel _actionRow = new();
        private readonly ThemeField _fldLeft = new();
        private readonly ThemeField _fldRight = new();
        private readonly ThemeButton _btnBrowseLeft = new();
        private readonly ThemeButton _btnBrowseRight = new();
        private readonly ThemeButton _btnCompare = new();
        private readonly ThemeCheckBox _chkOnlyDiff = new();
        private readonly ThemeCheckBox _chkSortKeys = new();
        private readonly Label _lblLeftCaption = new();
        private readonly Label _lblRightCaption = new();
        private readonly Label _lblStatus = new();
        private readonly Label _lblLeftTitle = new();
        private readonly Label _lblRightTitle = new();
        private readonly SplitContainer _split = new();
        private readonly DataGridView _gridLeft = new();
        private readonly DataGridView _gridRight = new();

        private List<TextDiffLine> _diff = new();
        private List<TextDiffLine> _display = new();
        private List<TableDiffRow> _tableDiff = new();
        private List<TableDiffRow> _tableDisplay = new();
        private bool _tableMode;
        private int _tableColumnCount;
        private int _compareVersion;
        private string _compareNote = "";
        private bool _syncingScroll;
        private bool _busy;
        private int _leftLineCount;
        private int _rightLineCount;

        internal TextCompareForm()
        {
            InitializeComponent();
            BuildLayout();
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;

            Shown += (_, _) =>
            {
                // 先定分隔位置再定最小宽度：SplitContainer 默认距离 50，
                // 直接设 MinSize 会触发"距离必须介于 Panel1MinSize 与 Width-Panel2MinSize 之间"
                try
                {
                    _split.SplitterDistance = Math.Max(240, _split.Width / 2);
                    _split.Panel1MinSize = 240;
                    _split.Panel2MinSize = 240;
                    // SplitContainer 在句柄创建前算过一次子控件宽度，这里按最终面板尺寸再钳一次；
                    // 否则左网格会停在初始宽度、右网格溢出到面板外
                    NormalizePaneLayout();
                }
                catch (InvalidOperationException) { }
                _fldLeft.Edit.Focus();
            };
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            ClientSize = new Size(1280, 780);
            MinimumSize = new Size(900, 560);
            Text = "文本比对";
            StartPosition = FormStartPosition.CenterScreen;
            Font = Theme.Body;
            Icon = WindowIcons.Vcmp;
            ResumeLayout(false);
        }

        // ==================== 布局 ====================

        private void BuildLayout()
        {
            SuspendLayout();

            _top.Dock = DockStyle.Top;
            _top.Height = 122;
            _top.ColumnCount = 3;
            _top.RowCount = 3;
            _top.Padding = new Padding(ThemeMetrics.Margin, ThemeMetrics.RowPad, ThemeMetrics.Margin, 0);
            _top.BackColor = Color.Transparent;
            _top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            _top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92f));
            _top.RowStyles.Add(new RowStyle(SizeType.Absolute, 34f));
            _top.RowStyles.Add(new RowStyle(SizeType.Absolute, 34f));
            _top.RowStyles.Add(new RowStyle(SizeType.Absolute, 46f));

            ConfigureCaption(_lblLeftCaption, "左文件");
            ConfigureCaption(_lblRightCaption, "右文件");

            _fldLeft.Dock = DockStyle.Fill;
            _fldLeft.Margin = new Padding(0, 4, ThemeMetrics.Gap, 4);
            _fldLeft.Placeholder = "选择左侧文件，或把文件拖进来";
            _fldRight.Dock = DockStyle.Fill;
            _fldRight.Margin = new Padding(0, 4, ThemeMetrics.Gap, 4);
            _fldRight.Placeholder = "选择右侧文件，或把文件拖进来";

            ConfigureBrowseButton(_btnBrowseLeft, "浏览…", (_, _) => BrowseInto(left: true));
            ConfigureBrowseButton(_btnBrowseRight, "浏览…", (_, _) => BrowseInto(left: false));

            _btnCompare.Text = "比对";
            _btnCompare.Primary = true;
            _btnCompare.Font = Theme.Body;
            _btnCompare.Margin = new Padding(0, 0, ThemeMetrics.Group, 0);
            _btnCompare.Click += OnCompareClick;

            _chkOnlyDiff.Text = "只保留差异行";
            _chkOnlyDiff.AutoSize = true;
            _chkOnlyDiff.Font = Theme.Body;
            _chkOnlyDiff.Margin = new Padding(0, 7, ThemeMetrics.Group, 0);
            _chkOnlyDiff.CheckedChanged += OnOnlyDifferencesChanged;

            _chkSortKeys.Text = "对比前按 key 排序";
            _chkSortKeys.AutoSize = true;
            _chkSortKeys.Font = Theme.Body;
            _chkSortKeys.Margin = new Padding(0, 7, ThemeMetrics.Group, 0);
            _chkSortKeys.Visible = false;
            _chkSortKeys.CheckedChanged += OnSortKeysChanged;

            _lblStatus.AutoSize = true;
            _lblStatus.Font = Theme.Caption;
            _lblStatus.Margin = new Padding(0, 9, 0, 0);
            _lblStatus.Text = "选择左右两个文件后点「比对」";

            _actionRow.Dock = DockStyle.Fill;
            _actionRow.FlowDirection = FlowDirection.LeftToRight;
            _actionRow.WrapContents = false;
            _actionRow.BackColor = Color.Transparent;
            _actionRow.Margin = new Padding(0, 4, 0, 0);
            _actionRow.Controls.AddRange(new Control[] { _btnCompare, _chkOnlyDiff, _chkSortKeys, _lblStatus });

            _top.Controls.Add(_lblLeftCaption, 0, 0);
            _top.Controls.Add(_fldLeft, 1, 0);
            _top.Controls.Add(_btnBrowseLeft, 2, 0);
            _top.Controls.Add(_lblRightCaption, 0, 1);
            _top.Controls.Add(_fldRight, 1, 1);
            _top.Controls.Add(_btnBrowseRight, 2, 1);
            _top.Controls.Add(_actionRow, 0, 2);
            _top.SetColumnSpan(_actionRow, 3);

            // ---------- 主体：左右两个网格 ----------
            _split.Dock = DockStyle.Fill;
            _split.Orientation = Orientation.Vertical;
            _split.SplitterWidth = 6;

            ConfigureGrid(_gridLeft);
            ConfigureGrid(_gridRight);
            ConfigurePaneTitle(_lblLeftTitle, "左：未选择文件");
            ConfigurePaneTitle(_lblRightTitle, "右：未选择文件");

            _split.Panel1.Controls.Add(_gridLeft);
            _split.Panel1.Controls.Add(_lblLeftTitle);
            _split.Panel2.Controls.Add(_gridRight);
            _split.Panel2.Controls.Add(_lblRightTitle);

            Controls.Add(_split);
            Controls.Add(_top);

            // 拖拽两个文件进来：先落左、再落右；只有一个文件时补空着的一侧
            AllowDrop = true;
            DragEnter += OnDragEnter;
            DragDrop += OnFormDragDrop;
            foreach (var grid in new[] { _gridLeft, _gridRight })
            {
                grid.AllowDrop = true;
                grid.DragEnter += OnDragEnter;
                grid.DragDrop += (s, e) => OnGridDragDrop(grid, e);
            }

            _fldLeft.Edit.KeyDown += OnPathKeyDown;
            _fldRight.Edit.KeyDown += OnPathKeyDown;

            ResumeLayout(false);
            PerformLayout();
        }

        private static void ConfigureCaption(Label label, string text)
        {
            label.Text = text;
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Font = Theme.Body;
            label.AutoSize = false;
            label.Margin = new Padding(0, 0, ThemeMetrics.Gap, 0);
        }

        private void NormalizePaneLayout()
        {
            if (_gridLeft.Parent != null)
                _gridLeft.Width = _split.Panel1.ClientSize.Width;
            if (_gridRight.Parent != null)
                _gridRight.Width = _split.Panel2.ClientSize.Width;
        }

        private static void ConfigurePaneTitle(Label label, string text)
        {
            label.Text = text;
            label.Dock = DockStyle.Top;
            label.Height = 26;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Padding = new Padding(ThemeMetrics.Margin, 0, 0, 0);
            label.Font = Theme.Caption;
            label.AutoEllipsis = true;
        }

        private void ConfigureBrowseButton(ThemeButton button, string text, EventHandler onClick)
        {
            button.Text = text;
            button.Font = Theme.Body;
            button.Dock = DockStyle.Fill;
            button.Margin = new Padding(0, 4, 0, 4);
            button.Click += onClick;
        }

        private void ConfigureGrid(DataGridView grid)
        {
            grid.Dock = DockStyle.Fill;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.AllowUserToOrderColumns = false;
            grid.RowHeadersVisible = false;
            grid.AutoGenerateColumns = false;
            grid.MultiSelect = true;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.EditMode = DataGridViewEditMode.EditProgrammatically;
            grid.StandardTab = true;
            grid.ShowCellToolTips = true;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // 列在每次绑定前按模式重建（文本 = 行号+内容；表格 = 行号+A/B/C…）
            grid.CellFormatting += OnCellFormatting;
            grid.Scroll += OnGridScroll;
        }

        /// <summary>按模式重建两侧网格的列。表格模式列数取自两边最大列数。</summary>
        private void ConfigureColumns(bool table, int dataColumnCount)
        {
            foreach (var grid in new[] { _gridLeft, _gridRight })
            {
                grid.DataSource = null;
                grid.Columns.Clear();

                var lineNumberColumn = new DataGridViewTextBoxColumn
                {
                    Name = "ln",
                    DataPropertyName = "ln",
                    HeaderText = "行号",
                    Width = 64,
                    MinimumWidth = 48,
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    Resizable = DataGridViewTriState.False
                };
                lineNumberColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                lineNumberColumn.DefaultCellStyle.Font = Theme.Mono;
                lineNumberColumn.DefaultCellStyle.ForeColor = Theme.P.Muted;
                grid.Columns.Add(lineNumberColumn);

                if (!table)
                {
                    var textColumn = new DataGridViewTextBoxColumn
                    {
                        Name = "text",
                        DataPropertyName = "text",
                        HeaderText = "内容",
                        AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                        MinimumWidth = 120,
                        SortMode = DataGridViewColumnSortMode.NotSortable
                    };
                    textColumn.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
                    textColumn.DefaultCellStyle.Font = Theme.Mono;
                    grid.Columns.Add(textColumn);
                    continue;
                }

                for (int i = 0; i < dataColumnCount; i++)
                {
                    var cellColumn = new DataGridViewTextBoxColumn
                    {
                        Name = "c" + i,
                        DataPropertyName = "c" + i,
                        HeaderText = ColumnName(i),
                        Width = 96,
                        MinimumWidth = 60,
                        SortMode = DataGridViewColumnSortMode.NotSortable
                    };
                    cellColumn.DefaultCellStyle.Font = Theme.Mono;
                    cellColumn.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
                    if (i == dataColumnCount - 1)
                    {
                        cellColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        cellColumn.MinimumWidth = 96;
                    }
                    grid.Columns.Add(cellColumn);
                }
            }
        }

        /// <summary>列头字母：A、B…Z、AA…（和 Excel 列口径一致）。</summary>
        private static string ColumnName(int index)
        {
            var sb = new StringBuilder();
            int value = index;
            do
            {
                sb.Insert(0, (char)('A' + value % 26));
                value = value / 26 - 1;
            } while (value >= 0);
            return sb.ToString();
        }

        // ==================== 主题 ====================

        private void ApplyTheme()
        {
            if (IsDisposed) return;
            var p = Theme.P;

            Font = Theme.Body;
            BackColor = p.Window;
            _top.BackColor = p.Chrome;
            _actionRow.BackColor = p.Chrome;

            _lblLeftCaption.ForeColor = p.Text;
            _lblRightCaption.ForeColor = p.Text;
            _lblLeftCaption.BackColor = Color.Transparent;
            _lblRightCaption.BackColor = Color.Transparent;

            _fldLeft.ApplyTheme();
            _fldRight.ApplyTheme();

            foreach (var button in new[] { _btnBrowseLeft, _btnBrowseRight, _btnCompare })
            {
                button.Font = Theme.Body;
                button.Height = ThemeMetrics.RowH;
                button.Invalidate();
            }
            _btnCompare.Size = new Size(ThemeButton.FitWidth(_btnCompare.Text), ThemeMetrics.RowH);

            _lblStatus.Font = Theme.Caption;
            _lblStatus.ForeColor = p.Muted;
            _chkSortKeys.BackColor = p.Chrome;
            _chkSortKeys.ForeColor = p.Text;
            _chkSortKeys.Invalidate();

            _split.BackColor = p.Divider;
            _split.Panel1.BackColor = p.Surface;
            _split.Panel2.BackColor = p.Surface;

            foreach (var title in new[] { _lblLeftTitle, _lblRightTitle })
            {
                title.BackColor = p.Surface;
                title.ForeColor = p.Muted;
                title.Font = Theme.Caption;
            }

            ApplyGridTheme(_gridLeft);
            ApplyGridTheme(_gridRight);
            Invalidate(true);
        }

        private static void ApplyGridTheme(DataGridView grid)
        {
            ThemeGrid.Apply(grid);
            grid.GridColor = Theme.P.GridLine;
            grid.BackgroundColor = Theme.P.Surface;
            grid.RowTemplate.Height = Theme.P.RowHeight;

            foreach (DataGridViewColumn column in grid.Columns)
            {
                column.DefaultCellStyle.Font = Theme.Mono;
                if (column.Name == "ln")
                    column.DefaultCellStyle.ForeColor = Theme.P.Muted;
            }
            grid.Invalidate();
        }

        // ==================== 选择文件 / 比对 ====================

        private void BrowseInto(bool left)
        {
            using var dialog = new OpenFileDialog
            {
                Title = left ? "选择左侧文件" : "选择右侧文件",
                CheckFileExists = true,
                Filter = "文本文件 (*.txt;*.csv;*.json;*.xml;*.log;*.ini;*.cfg;*.yaml;*.yml)|*.txt;*.csv;*.json;*.xml;*.log;*.ini;*.cfg;*.yaml;*.yml|所有文件 (*.*)|*.*"
            };

            var field = left ? _fldLeft : _fldRight;
            var current = field.Edit.Text.Trim();
            if (File.Exists(current))
            {
                dialog.InitialDirectory = Path.GetDirectoryName(current) ?? "";
                dialog.FileName = Path.GetFileName(current);
            }
            else
            {
                var lastDir = CsvSettings.TextCompareLastDirectory;
                if (!string.IsNullOrEmpty(lastDir) && Directory.Exists(lastDir))
                    dialog.InitialDirectory = lastDir;
            }

            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            SetFile(left, dialog.FileName);
        }

        private void SetFile(bool left, string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            if (left) _fldLeft.Edit.Text = path;
            else _fldRight.Edit.Text = path;

            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir)) CsvSettings.TextCompareLastDirectory = dir;

            UpdateSortKeysVisibility();

            if (File.Exists(_fldLeft.Edit.Text.Trim()) && File.Exists(_fldRight.Edit.Text.Trim())
                && _diff.Count == 0)
                BeginCompare();   // 两个文件都齐了就先比一次，省一次点击
        }

        private void UpdateSortKeysVisibility()
        {
            var leftPath = _fldLeft.Edit.Text.Trim();
            var rightPath = _fldRight.Edit.Text.Trim();
            _chkSortKeys.Visible = File.Exists(leftPath) && File.Exists(rightPath)
                                   && IsKeySortable(leftPath) && IsKeySortable(rightPath);
        }

        private void OnSortKeysChanged(object? sender, EventArgs e)
        {
            var leftPath = _fldLeft.Edit.Text.Trim();
            var rightPath = _fldRight.Edit.Text.Trim();
            if (File.Exists(leftPath) && File.Exists(rightPath))
                BeginCompare();
        }

        private static bool IsKeySortable(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext is ".json" or ".yaml" or ".yml";
        }

        /// <summary>比对前把 key 顺序归一（JSON 递归排序；YAML 保守排序，不支持就返回原文）。</summary>
        private static (string[] Lines, string? Error) SortLinesForCompare(string path, string[] lines)
        {
            if (!IsKeySortable(path)) return (lines, null);

            string text = string.Join("\n", lines);
            try
            {
                if (string.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase))
                {
                    var parsed = JsonToolEngine.Parse(text);
                    if (!parsed.Success) return (lines, parsed.Error ?? "JSON 解析失败");
                    text = JsonToolEngine.Format(parsed.Root, sortKeys: true);
                }
                else
                {
                    if (!YamlKeySorter.TrySort(text, out string sorted))
                        return (lines, "YAML 结构暂不支持排序（含锚点/根序列/多文档等）");
                    text = sorted;
                }

                return (text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'), null);
            }
            catch (Exception ex)
            {
                return (lines, ex.Message);
            }
        }

        private void OnCompareClick(object? sender, EventArgs e) => BeginCompare();

        private void OnPathKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;
            e.Handled = true;
            e.SuppressKeyPress = true;
            BeginCompare();
        }

        private async void BeginCompare()
        {
            if (_busy) return;

            var leftPath = _fldLeft.Edit.Text.Trim();
            var rightPath = _fldRight.Edit.Text.Trim();

            if (!File.Exists(leftPath))
            {
                AlertForm.Show(this, "左文件不存在，请先选择。", "文本比对", AlertType.Warning);
                return;
            }
            if (!File.Exists(rightPath))
            {
                AlertForm.Show(this, "右文件不存在，请先选择。", "文本比对", AlertType.Warning);
                return;
            }

            try
            {
                if (!ConfirmLargeFiles(leftPath, rightPath)) return;
            }
            catch (Exception ex)
            {
                AlertForm.Show(this, "读取文件信息失败：" + ex.Message, "文本比对", AlertType.Error);
                return;
            }

            bool leftIsTable = SimpleTableReader.IsTableFile(leftPath);
            bool rightIsTable = SimpleTableReader.IsTableFile(rightPath);
            if (leftIsTable != rightIsTable)
            {
                AlertForm.Show(this,
                    "两边文件类型不一致：表格（csv/xlsx/xls）请对表格，文本请对文本。",
                    "文本比对", AlertType.Warning);
                return;
            }

            _chkSortKeys.Visible = !leftIsTable && IsKeySortable(leftPath) && IsKeySortable(rightPath);

            int version = ++_compareVersion;
            _compareNote = "";
            var stopwatch = Stopwatch.StartNew();
            SetBusy(true, leftIsTable ? "正在读取并比对表格…" : "正在读取并比对…");

            try
            {
                if (leftIsTable)
                {
                    var leftRows = await Task.Run(() => SimpleTableReader.ReadRows(leftPath));
                    var rightRows = await Task.Run(() => SimpleTableReader.ReadRows(rightPath));
                    var tableDiff = await Task.Run(() => TableDiffEngine.Compare(leftRows, rightRows));

                    if (version != _compareVersion || IsDisposed) return;

                    _tableMode = true;
                    _tableDiff = tableDiff;
                    _leftLineCount = leftRows.Count;
                    _rightLineCount = rightRows.Count;
                    _tableColumnCount = TableDiffEngine.MaxColumnCount(tableDiff);
                    UpdateTitles(leftPath, rightPath);
                    _compareNote = $" · 用时 {stopwatch.ElapsedMilliseconds} ms";
                    BindGrids();
                }
                else
                {
                    var leftLines = await Task.Run(() => TextFileReader.ReadAllLines(leftPath));
                    var rightLines = await Task.Run(() => TextFileReader.ReadAllLines(rightPath));

                    string? sortError = null;
                    if (_chkSortKeys.Checked)
                    {
                        var leftSorted = await Task.Run(() => SortLinesForCompare(leftPath, leftLines));
                        var rightSorted = await Task.Run(() => SortLinesForCompare(rightPath, rightLines));
                        leftLines = leftSorted.Lines;
                        rightLines = rightSorted.Lines;
                        sortError = leftSorted.Error ?? rightSorted.Error;
                    }

                    var diff = await Task.Run(() => TextDiffEngine.Compare(leftLines, rightLines));

                    if (version != _compareVersion || IsDisposed) return;

                    _tableMode = false;
                    _tableDiff = new List<TableDiffRow>();
                    _leftLineCount = leftLines.Length;
                    _rightLineCount = rightLines.Length;
                    _diff = diff;
                    UpdateTitles(leftPath, rightPath);
                    _compareNote = $" · 用时 {stopwatch.ElapsedMilliseconds} ms";
                    if (_chkSortKeys.Checked)
                        _compareNote += sortError == null
                            ? " · 已按 key 排序"
                            : " · key 排序失败：" + sortError;
                    BindGrids();
                }
            }
            catch (Exception ex)
            {
                if (version == _compareVersion)
                    AlertForm.Show(this, "比对失败：" + ex.Message, "文本比对", AlertType.Error);
            }
            finally
            {
                if (version == _compareVersion)
                {
                    SetBusy(false);
                    UpdateStatus();
                }
            }
        }

        /// <summary>超过 20MB 先问一句，避免用户以为界面卡死。</summary>
        private bool ConfirmLargeFiles(string leftPath, string rightPath)
        {
            const long limit = 20L * 1024 * 1024;
            long size = Math.Max(new FileInfo(leftPath).Length, new FileInfo(rightPath).Length);
            if (size <= limit) return true;

            return AlertForm.ShowConfirm(this,
                $"文件较大（最大 {size / 1024 / 1024} MB），比对可能需要较长时间，是否继续？",
                "文本比对", AlertType.Warning) == DialogResult.Yes;
        }

        private void UpdateTitles(string leftPath, string rightPath)
        {
            string suffix = _tableMode ? "（表格）" : "";
            _lblLeftTitle.Text = "左：" + Path.GetFileName(leftPath) + suffix;
            _lblRightTitle.Text = "右：" + Path.GetFileName(rightPath) + suffix;
        }

        private void SetBusy(bool busy, string? message = null)
        {
            _busy = busy;
            _btnCompare.Enabled = !busy;
            _btnBrowseLeft.Enabled = !busy;
            _btnBrowseRight.Enabled = !busy;
            _chkOnlyDiff.Enabled = !busy;
            _chkSortKeys.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
            if (message != null) _lblStatus.Text = message;
        }

        // ==================== 绑定与着色 ====================

        private void OnOnlyDifferencesChanged(object? sender, EventArgs e)
        {
            if (_tableMode ? _tableDiff.Count == 0 : _diff.Count == 0) return;
            BindGrids();
        }

        private void BindGrids()
        {
            if (_tableMode)
            {
                _tableDisplay = _chkOnlyDiff.Checked ? TableDiffEngine.OnlyDifferences(_tableDiff) : _tableDiff;
                ConfigureColumns(table: true, _tableColumnCount);
                _gridLeft.DataSource = BuildTableGrid(left: true);
                _gridRight.DataSource = BuildTableGrid(left: false);
            }
            else
            {
                _display = _chkOnlyDiff.Checked ? TextDiffEngine.OnlyDifferences(_diff) : _diff;
                ConfigureColumns(table: false, 0);
                _gridLeft.DataSource = BuildTextTable(left: true);
                _gridRight.DataSource = BuildTextTable(left: false);
            }

            try
            {
                if (_gridLeft.Rows.Count > 0)
                {
                    _gridLeft.FirstDisplayedScrollingRowIndex = 0;
                    _gridRight.FirstDisplayedScrollingRowIndex = 0;
                }
            }
            catch (InvalidOperationException) { }

            UpdateStatus();
        }

        private DataTable BuildTextTable(bool left)
        {
            var table = new DataTable();
            table.Columns.Add("ln", typeof(string));
            table.Columns.Add("text", typeof(string));

            foreach (var line in _display)
            {
                var number = left ? line.LeftLineNumber : line.RightLineNumber;
                table.Rows.Add(number?.ToString() ?? "", left ? line.LeftText : line.RightText);
            }
            return table;
        }

        private DataTable BuildTableGrid(bool left)
        {
            var table = new DataTable();
            table.Columns.Add("ln", typeof(string));
            for (int i = 0; i < _tableColumnCount; i++)
                table.Columns.Add("c" + i, typeof(string));

            foreach (var row in _tableDisplay)
            {
                var values = new object[_tableColumnCount + 1];
                values[0] = (left ? row.LeftRowNumber : row.RightRowNumber)?.ToString() ?? "";
                var cells = left ? row.LeftCells : row.RightCells;
                for (int i = 0; i < _tableColumnCount; i++)
                    values[i + 1] = i < cells.Length ? cells[i] : "";
                table.Rows.Add(values);
            }
            return table;
        }

        private void UpdateStatus()
        {
            if (_tableMode)
            {
                UpdateTableStatus();
                return;
            }

            if (_diff.Count == 0)
            {
                _lblStatus.Text = _compareNote.Length > 0
                    ? "两文件内容一致" + _compareNote
                    : "选择左右两个文件后点「比对」";
                return;
            }

            int modified = 0, added = 0, deleted = 0;
            foreach (var line in _diff)
            {
                switch (line.Kind)
                {
                    case TextDiffKind.Modified: modified++; break;
                    case TextDiffKind.Added: added++; break;
                    case TextDiffKind.Deleted: deleted++; break;
                }
            }

            var parts = new List<string>
            {
                $"左 {_leftLineCount} 行",
                $"右 {_rightLineCount} 行"
            };

            if (modified + added + deleted == 0)
            {
                parts.Add("两文件内容一致");
            }
            else
            {
                parts.Add($"修改 {modified}");
                parts.Add($"新增 {added}");
                parts.Add($"删除 {deleted}");
                if (_chkOnlyDiff.Checked)
                    parts.Add($"仅显示差异 {_display.Count} 行");
            }

            _lblStatus.Text = string.Join(" · ", parts) + _compareNote;
        }

        private void UpdateTableStatus()
        {
            if (_tableDiff.Count == 0)
            {
                _lblStatus.Text = "两边都是空表" + _compareNote;
                return;
            }

            int modified = 0, added = 0, deleted = 0, diffCells = 0;
            foreach (var row in _tableDiff)
            {
                switch (row.Kind)
                {
                    case TableRowDiffKind.Modified:
                        modified++;
                        diffCells += row.DiffColumns.Count;
                        break;
                    case TableRowDiffKind.Added: added++; break;
                    case TableRowDiffKind.Deleted: deleted++; break;
                }
            }

            var parts = new List<string>
            {
                "表格比对",
                $"左 {_leftLineCount} 行 × {_tableColumnCount} 列",
                $"右 {_rightLineCount} 行 × {_tableColumnCount} 列"
            };

            if (modified + added + deleted == 0)
            {
                parts.Add("两边内容一致");
            }
            else
            {
                parts.Add($"修改 {modified} 行（{diffCells} 处单元格）");
                parts.Add($"新增 {added} 行");
                parts.Add($"删除 {deleted} 行");
                if (_chkOnlyDiff.Checked)
                    parts.Add($"仅显示差异 {_tableDisplay.Count} 行");
            }

            _lblStatus.Text = string.Join(" · ", parts) + _compareNote;
        }

        private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (_tableMode)
            {
                FormatTableCell(sender, e);
                return;
            }

            if (e.RowIndex < 0 || e.RowIndex >= _display.Count) return;

            bool left = ReferenceEquals(sender, _gridLeft);
            var kind = _display[e.RowIndex].Kind;

            Color? background = kind switch
            {
                TextDiffKind.Modified => Theme.CellTint(Color.FromArgb(255, 238, 170)),
                TextDiffKind.Deleted => left ? Theme.CellTint(Color.FromArgb(255, 200, 200))
                                             : Theme.CellTint(Color.FromArgb(242, 242, 242)),
                TextDiffKind.Added => left ? Theme.CellTint(Color.FromArgb(242, 242, 242))
                                           : Theme.CellTint(Color.FromArgb(200, 238, 200)),
                _ => null
            };

            if (background.HasValue)
            {
                e.CellStyle.BackColor = background.Value;
                e.CellStyle.SelectionBackColor = Theme.CellTint(Color.FromArgb(210, 220, 235));
            }

            // 对侧没有这一行（新增/删除）时，行号列留空、内容用灰字提示"这里没有内容"
            bool hasLine = left
                ? _display[e.RowIndex].LeftLineNumber != null
                : _display[e.RowIndex].RightLineNumber != null;
            if (!hasLine && e.ColumnIndex == 1)
                e.CellStyle.ForeColor = Theme.P.Muted;
        }

        /// <summary>表格模式着色：修改行只染发生变化的单元格，新增/删除整行染色。</summary>
        private void FormatTableCell(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _tableDisplay.Count) return;
            if (e.ColumnIndex <= 0) return;   // 行号列不染色

            bool left = ReferenceEquals(sender, _gridLeft);
            var row = _tableDisplay[e.RowIndex];
            int dataColumn = e.ColumnIndex - 1;

            Color? background = row.Kind switch
            {
                TableRowDiffKind.Modified => row.DiffColumns.Contains(dataColumn)
                    ? Theme.CellTint(Color.FromArgb(255, 200, 200))
                    : null,
                TableRowDiffKind.Deleted => left
                    ? Theme.CellTint(Color.FromArgb(255, 200, 200))
                    : Theme.CellTint(Color.FromArgb(242, 242, 242)),
                TableRowDiffKind.Added => left
                    ? Theme.CellTint(Color.FromArgb(242, 242, 242))
                    : Theme.CellTint(Color.FromArgb(200, 238, 200)),
                _ => null
            };

            if (background.HasValue)
            {
                e.CellStyle.BackColor = background.Value;
                e.CellStyle.SelectionBackColor = Theme.CellTint(Color.FromArgb(210, 220, 235));
            }

            bool hasRow = left ? row.LeftRowNumber != null : row.RightRowNumber != null;
            if (!hasRow)
                e.CellStyle.ForeColor = Theme.P.Muted;
        }

        // ==================== 滚动同步 / 拖拽 ====================

        private void OnGridScroll(object? sender, ScrollEventArgs e)
        {
            if (_syncingScroll || sender is not DataGridView source) return;
            var target = ReferenceEquals(source, _gridLeft) ? _gridRight : _gridLeft;
            if (target.Rows.Count == 0 || source.Rows.Count == 0) return;

            _syncingScroll = true;
            try
            {
                int row = source.FirstDisplayedScrollingRowIndex;
                if (row >= 0 && row < target.Rows.Count && target.FirstDisplayedScrollingRowIndex != row)
                    target.FirstDisplayedScrollingRowIndex = row;
            }
            catch (InvalidOperationException) { }
            finally
            {
                _syncingScroll = false;
            }
        }

        private static void OnDragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effect = DragDropEffects.Copy;
        }

        private void OnFormDragDrop(object? sender, DragEventArgs e)
        {
            var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;

            if (files.Length >= 2)
            {
                SetFile(left: true, files[0]);
                SetFile(left: false, files[1]);
                return;
            }

            bool leftEmpty = string.IsNullOrWhiteSpace(_fldLeft.Edit.Text);
            SetFile(leftEmpty, files[0]);
        }

        private void OnGridDragDrop(DataGridView grid, DragEventArgs e)
        {
            var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;
            SetFile(ReferenceEquals(grid, _gridLeft), files[0]);
        }
    }
}
