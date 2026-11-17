using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 文本比对：左右各选一个文件，按行做差分。
    ///
    /// 两边网格绑定同一份"按行对齐"的结果（<see cref="TextDiffEngine"/>），
    /// 所以行数天然一致、滚动同步不会错位：
    ///   修改 = 黄色，左侧独有 = 红色，右侧独有 = 绿色，另一侧的空位 = 灰色。
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
        private int _compareVersion;
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

            _lblStatus.AutoSize = true;
            _lblStatus.Font = Theme.Caption;
            _lblStatus.Margin = new Padding(0, 9, 0, 0);
            _lblStatus.Text = "选择左右两个文件后点「比对」";

            _actionRow.Dock = DockStyle.Fill;
            _actionRow.FlowDirection = FlowDirection.LeftToRight;
            _actionRow.WrapContents = false;
            _actionRow.BackColor = Color.Transparent;
            _actionRow.Margin = new Padding(0, 4, 0, 0);
            _actionRow.Controls.AddRange(new Control[] { _btnCompare, _chkOnlyDiff, _lblStatus });

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

            grid.Columns.AddRange(lineNumberColumn, textColumn);

            grid.CellFormatting += OnCellFormatting;
            grid.Scroll += OnGridScroll;
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

            if (grid.Columns["text"] is { } textColumn)
                textColumn.DefaultCellStyle.Font = Theme.Mono;
            if (grid.Columns["ln"] is { } lineNumberColumn)
            {
                lineNumberColumn.DefaultCellStyle.Font = Theme.Mono;
                lineNumberColumn.DefaultCellStyle.ForeColor = Theme.P.Muted;
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

            if (File.Exists(_fldLeft.Edit.Text.Trim()) && File.Exists(_fldRight.Edit.Text.Trim())
                && _diff.Count == 0)
                BeginCompare();   // 两个文件都齐了就先比一次，省一次点击
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

            int version = ++_compareVersion;
            SetBusy(true, "正在读取并比对…");

            try
            {
                var leftLines = await Task.Run(() => TextFileReader.ReadAllLines(leftPath));
                var rightLines = await Task.Run(() => TextFileReader.ReadAllLines(rightPath));
                var diff = await Task.Run(() => TextDiffEngine.Compare(leftLines, rightLines));

                if (version != _compareVersion || IsDisposed) return;

                _leftLineCount = leftLines.Length;
                _rightLineCount = rightLines.Length;
                _diff = diff;
                _lblLeftTitle.Text = "左：" + Path.GetFileName(leftPath);
                _lblRightTitle.Text = "右：" + Path.GetFileName(rightPath);

                BindGrids();
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

        private void SetBusy(bool busy, string? message = null)
        {
            _busy = busy;
            _btnCompare.Enabled = !busy;
            _btnBrowseLeft.Enabled = !busy;
            _btnBrowseRight.Enabled = !busy;
            _chkOnlyDiff.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
            if (message != null) _lblStatus.Text = message;
        }

        // ==================== 绑定与着色 ====================

        private void OnOnlyDifferencesChanged(object? sender, EventArgs e)
        {
            if (_diff.Count == 0) return;
            BindGrids();
        }

        private void BindGrids()
        {
            _display = _chkOnlyDiff.Checked ? TextDiffEngine.OnlyDifferences(_diff) : _diff;

            _gridLeft.DataSource = BuildTable(left: true);
            _gridRight.DataSource = BuildTable(left: false);

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

        private DataTable BuildTable(bool left)
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

        private void UpdateStatus()
        {
            if (_diff.Count == 0)
            {
                _lblStatus.Text = "选择左右两个文件后点「比对」";
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

            _lblStatus.Text = string.Join(" · ", parts);
        }

        private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
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
