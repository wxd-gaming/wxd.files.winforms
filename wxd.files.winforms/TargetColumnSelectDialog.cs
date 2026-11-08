using System.Data;

namespace wxd.files.winforms
{
    /// <summary>目标列选择对话框 — 带固定行数据的单选面板（类似列选择面板，但使用单选框，支持表头行显示）</summary>
    public class TargetColumnSelectDialog : Form
    {
        private ListView _lvColumns = new();
        private ThemeButton _btnOK = new();
        private ThemeButton _btnCancel = new();
        private ThemeField _fldSearch = new();
        private Label _lblHint = new();
        private Panel _topBar = new();
        private Panel _topDivider = new();
        private Panel _bar = new();
        private int _selectedIndex = -1;
        private string _selectedColumnName = "";        // 选中的列的表头名称
        private int _headerRowIndex = -1;              // 表头行索引
        private bool _hasHeaderRow = false;
        private string _customTitle = "选择目标列";
        private DataTable _frozenData;                  // 保存原始数据用于过滤重建
        private int _frozenRows;                        // 保存冻结行数

        public int SelectedColumnIndex => _selectedIndex;
        public string SelectedColumnName => _selectedColumnName ?? "";

        public TargetColumnSelectDialog(DataTable frozenData, int frozenRows, int initialSelection, int headerRowIndex = -1, string title = "选择目标列")
        {
            _selectedIndex = initialSelection;
            _headerRowIndex = headerRowIndex;
            _customTitle = title;
            _frozenData = frozenData;
            _frozenRows = frozenRows;
            InitializeComponent();
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
            PopulateList(frozenData, frozenRows);
            if (_selectedIndex >= 0 && _selectedIndex < _lvColumns.Items.Count)
            {
                _lvColumns.Items[_selectedIndex].Selected = true;
                _lvColumns.EnsureVisible(_selectedIndex);
            }
        }

        private void PopulateList(DataTable frozenData, int frozenRows, string filter = "")
        {
            _lvColumns.BeginUpdate();
            _lvColumns.Items.Clear();
            _lvColumns.Columns.Clear();

            int frozenCount = Math.Min(frozenRows > 0 ? frozenRows : 2, frozenData.Rows.Count);

            // 表头行标志：如果有设置表头行，用于显示和获取列名
            _hasHeaderRow = (_headerRowIndex >= 0 && _headerRowIndex < frozenData.Rows.Count && _headerRowIndex < frozenCount);

            // 第一列：选择
            _lvColumns.Columns.Add("选择", 50, System.Windows.Forms.HorizontalAlignment.Center);

            // 表头名列：如果有设置表头行，优先显示字段名
            if (_hasHeaderRow)
                _lvColumns.Columns.Add($"字段名(行{_headerRowIndex + 1})", 160, System.Windows.Forms.HorizontalAlignment.Left);

            // 其余冻结行
            for (int r = 0; r < frozenCount; r++)
            {
                if (r == _headerRowIndex) continue; // 已作为字段名列显示
                _lvColumns.Columns.Add($"行{r + 1}", 140, System.Windows.Forms.HorizontalAlignment.Left);
            }

            for (int col = 0; col < frozenData.Columns.Count; col++)
            {
                var item = new ListViewItem("")
                {
                    Tag = col
                };

                // 先添加字段名（如果设置了表头行）
                if (_hasHeaderRow)
                {
                    var headerVal = frozenData.Rows[_headerRowIndex][col]?.ToString() ?? "";
                    item.SubItems.Add(headerVal);
                }

                // 添加其余冻结行的值
                for (int r = 0; r < frozenCount; r++)
                {
                    if (r == _headerRowIndex) continue;
                    var val = frozenData.Rows[r][col]?.ToString() ?? "";
                    item.SubItems.Add(val);
                }

                // 过滤：检查所有 SubItems 是否包含搜索词
                if (!string.IsNullOrEmpty(filter))
                {
                    bool match = false;
                    for (int si = 0; si < item.SubItems.Count; si++)
                    {
                        if (item.SubItems[si].Text.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            match = true;
                            break;
                        }
                    }
                    if (!match) continue;
                }

                _lvColumns.Items.Add(item);
            }

            _lvColumns.EndUpdate();

            // 恢复选中状态（搜索后尝试重新选中之前的列）
            if (_selectedIndex >= 0)
            {
                foreach (ListViewItem item in _lvColumns.Items)
                {
                    if (item.Tag is int colIndex && colIndex == _selectedIndex)
                    {
                        item.Selected = true;
                        item.EnsureVisible();
                        break;
                    }
                }
            }
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = _customTitle;
            Icon = WindowIcons.Colmgr;
            ClientSize = new Size(780, 520);
            MinimumSize = new Size(560, 400);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false; MaximizeBox = false;
            AutoScaleMode = AutoScaleMode.Font;

            _topBar = new Panel { Dock = DockStyle.Top, Height = ThemeMetrics.BarHeight };
            _topDivider = new Panel { Dock = DockStyle.Bottom, Height = 1 };
            _topBar.Controls.Add(_topDivider);

            _lblHint = new Label { Text = "请点击行选择目标关联列（单选）:", AutoSize = true };
            _fldSearch = new ThemeField { Placeholder = "搜索列..." };
            _fldSearch.Edit.TextChanged += (s, e) =>
                PopulateList(_frozenData, _frozenRows, _fldSearch.Edit.Text.Trim());
            _topBar.Controls.AddRange(new Control[] { _lblHint, _fldSearch });
            LayoutTopBar();
            _topBar.Resize += (_, _) => LayoutTopBar();

            _lvColumns = new ListView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                GridLines = true,
                CheckBoxes = false,
                MultiSelect = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            _lvColumns.SelectedIndexChanged += (s, e) =>
            {
                if (_lvColumns.SelectedItems.Count > 0)
                {
                    var selectedItem = _lvColumns.SelectedItems[0];
                    _selectedIndex = (int)selectedItem.Tag;
                    // 获取选中的列的显示名称：有表头行用表头名称，否则用第一行数据作为列名
                    if (selectedItem.SubItems.Count > 1)
                        _selectedColumnName = selectedItem.SubItems[1]?.Text ?? "";
                    else
                        _selectedColumnName = "";
                }
            };

            _bar = DialogButtonBar.Create();
            _btnCancel = DialogButtonBar.Make("取消");
            _btnCancel.DialogResult = DialogResult.Cancel;
            _btnOK = DialogButtonBar.Make("确定", null, primary: true);
            _btnOK.DialogResult = DialogResult.OK;
            _bar.Controls.AddRange(new Control[] { _btnCancel, _btnOK });
            LayoutBar();
            _bar.Resize += (_, _) => LayoutBar();

            // 停靠顺序：后加入的靠外，所以列表先加、页眉页脚后加
            Controls.Add(_lvColumns);
            Controls.Add(_topBar);
            Controls.Add(_bar);
            AcceptButton = _btnOK; CancelButton = _btnCancel;
            ResumeLayout(false);
        }

        private void LayoutTopBar()
        {
            int rowH = ThemeMetrics.RowH;
            _lblHint.Location = new Point(ThemeMetrics.Margin,
                (ThemeMetrics.BarHeight - _lblHint.Height) / 2);
            // 搜索框贴右边，和提示文字之间留一组间距
            int w = 220;
            _fldSearch.SetBounds(_topBar.Width - ThemeMetrics.Margin - w, ThemeMetrics.RowPad, w, rowH);
        }

        private void LayoutBar()
            => DialogButtonBar.LayoutRight(_bar, _btnCancel, _btnOK);

        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _topBar.BackColor = p.Chrome;
            _bar.BackColor = p.Chrome;
            _topDivider.BackColor = p.Divider;

            _lblHint.ForeColor = p.Muted;
            _lvColumns.BackColor = p.Surface;
            _lvColumns.ForeColor = p.Text;
            _lvColumns.Font = Theme.List;

            _fldSearch.ApplyTheme();
            _btnCancel.BackColor = p.Chrome;
            _btnOK.BackColor = p.Chrome;

            LayoutTopBar();
            LayoutBar();
            Invalidate(true);
        }
    }
}
