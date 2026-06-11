using System.Data;

namespace wxd.files.winforms
{
    /// <summary>目标列选择对话框 — 带固定行数据的单选面板（类似列选择面板，但使用单选框，支持表头行显示）</summary>
    public class TargetColumnSelectDialog : Form
    {
        private ListView _lvColumns = new();
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private TextBox _txtSearch = new();
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
            SuspendLayout();
            Text = _customTitle;
            ClientSize = new Size(780, 520);
            MinimumSize = new Size(560, 400);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false; MaximizeBox = false;
            AutoScaleMode = AutoScaleMode.Font;

            Controls.Add(new Label { Text = "请点击行选择目标关联列（单选）:", Location = new Point(10, 10), Size = new Size(240, 20),
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(31, 41, 55) });

            // 搜索框（靠右对齐，不被标签遮挡）
            _txtSearch = new TextBox
            {
                Location = new Point(260, 8),
                Size = new Size(200, 24),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                PlaceholderText = "搜索列...",
                ForeColor = Color.FromArgb(20, 60, 100),
                BackColor = Color.White
            };
            _txtSearch.TextChanged += (s, e) =>
            {
                PopulateList(_frozenData, _frozenRows, _txtSearch.Text.Trim());
            };
            Controls.Add(_txtSearch);

            _lvColumns = new ListView
            {
                Location = new Point(10, 38),
                Size = new Size(760, 430),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(20, 60, 100),
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

            _btnOK = new Button { Text = "确定", Location = new Point(590, 478),
                Size = new Size(80, 28), Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White, FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                DialogResult = DialogResult.OK, Cursor = Cursors.Hand };
            Controls.Add(_btnOK);

            _btnCancel = new Button { Text = "取消", Location = new Point(680, 478),
                Size = new Size(70, 28), Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat, DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(235, 248, 255), ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) } };
            Controls.Add(_btnCancel);

            Controls.Add(_lvColumns);
            AcceptButton = _btnOK; CancelButton = _btnCancel;
            ResumeLayout(false);
        }
    }
}
