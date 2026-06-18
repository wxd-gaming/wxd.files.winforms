using System.Data;

namespace wxd.files.winforms
{
    /// <summary>列管理对话框 — ListView列显隐 + 关联设置</summary>
    public class ColumnManagerDialog : Form
    {
        private ListView _lvColumns = new();
        private Button _btnSelectAll = new();
        private Button _btnDeselectAll = new();
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private Button _btnLinkManage = new();
        private TextBox _txtSearch = new();
        private int _lockedColumnIndex = -1;
        private bool[] _checkedStates;
        private string _rootDir = "";
        private string _currentFilePath = "";
        private int _frozenRows = 0;
        private int _headerRowIndex = -1;
        private int _mappingRowIndex = -1;
        private List<(string display, string fullPath)> _fileEntries = new();
        private List<ColumnLinkConfig> _columnLinks = new();
        private DataTable _dataTable;

        public ColumnManagerDialog(DataTable dataTable, bool[] checkedStates, int frozenRows, int lockedColumnIndex,
            string rootDir = "", string currentFilePath = "", List<(string display, string fullPath)>? fileEntries = null,
            List<ColumnLinkConfig>? columnLinks = null, int headerRowIndex = -1,
            int mappingRowIndex = -1)
        {
            _dataTable = dataTable;
            _lockedColumnIndex = lockedColumnIndex;
            _checkedStates = (bool[])checkedStates.Clone();
            _rootDir = rootDir;
            _currentFilePath = currentFilePath;
            _frozenRows = frozenRows;
            _headerRowIndex = headerRowIndex;
            _mappingRowIndex = mappingRowIndex;
            _fileEntries = fileEntries ?? new();
            _columnLinks = columnLinks != null ? new List<ColumnLinkConfig>(columnLinks) : new();
            InitializeComponent();
            PopulateList(dataTable, frozenRows);
        }

        public bool[] GetCheckedStates() => _checkedStates;
        public List<ColumnLinkConfig> GetColumnLinks() => _columnLinks;

        private string GetRelativePath(string fullPath)
        {
            if (string.IsNullOrEmpty(_rootDir) || string.IsNullOrEmpty(fullPath))
                return Path.GetFileName(fullPath);
            if (fullPath.StartsWith(_rootDir, StringComparison.OrdinalIgnoreCase))
                return fullPath.Substring(_rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(fullPath);
        }

        private void PopulateList(DataTable dataTable, int frozenRows, string filter = "")
        {
            _lvColumns.BeginUpdate();
            _lvColumns.Items.Clear();
            _lvColumns.Columns.Clear();

            int frozenCount = Math.Min(frozenRows, dataTable.Rows.Count);

            // 第一列：选择
            _lvColumns.Columns.Add("选择", 50, System.Windows.Forms.HorizontalAlignment.Center);

            // 字段名列
            if (_headerRowIndex >= 0 && _headerRowIndex < dataTable.Rows.Count && _headerRowIndex < frozenCount)
                _lvColumns.Columns.Add($"字段名(行{_headerRowIndex + 1})", 140, System.Windows.Forms.HorizontalAlignment.Left);

            // 映射名列
            if (_mappingRowIndex >= 0 && _mappingRowIndex < dataTable.Rows.Count && _mappingRowIndex < frozenCount
                && _mappingRowIndex != _headerRowIndex)
                _lvColumns.Columns.Add($"映射名(行{_mappingRowIndex + 1})", 140, System.Windows.Forms.HorizontalAlignment.Left);

            // 其余冻结行
            for (int r = 0; r < frozenCount; r++)
            {
                if (r == _headerRowIndex || r == _mappingRowIndex) continue;
                _lvColumns.Columns.Add($"行{r + 1}", 130, System.Windows.Forms.HorizontalAlignment.Left);
            }

            // 关联列
            _lvColumns.Columns.Add("关联", 100, System.Windows.Forms.HorizontalAlignment.Center);

            string currentRelPath = GetRelativePath(_currentFilePath);

            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                var item = new ListViewItem("")
                {
                    Tag = col,
                    Checked = _checkedStates[col]
                };

                // 字段名
                if (_headerRowIndex >= 0 && _headerRowIndex < dataTable.Rows.Count && _headerRowIndex < frozenCount)
                {
                    var headerVal = dataTable.Rows[_headerRowIndex][col]?.ToString() ?? "";
                    item.SubItems.Add(headerVal);
                }

                // 映射名
                if (_mappingRowIndex >= 0 && _mappingRowIndex < dataTable.Rows.Count && _mappingRowIndex < frozenCount
                    && _mappingRowIndex != _headerRowIndex)
                {
                    var mappingVal = dataTable.Rows[_mappingRowIndex][col]?.ToString() ?? "";
                    item.SubItems.Add(mappingVal);
                }

                // 其余冻结行的值
                for (int r = 0; r < frozenCount; r++)
                {
                    if (r == _headerRowIndex || r == _mappingRowIndex) continue;
                    var val = dataTable.Rows[r][col]?.ToString() ?? "";
                    item.SubItems.Add(val);
                }

                var link = _columnLinks.Find(l =>
                    l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                    l.SourceColumnIndex == col);
                string displayText = "";
                if (link != null)
                {
                    int linkedFields = link.FieldLinks?.Count(f => !string.IsNullOrEmpty(f.TargetFile)) ?? 0;
                    if (linkedFields > 0)
                        displayText = $"⚡{linkedFields}字段";
                    else
                        displayText = "📋格式";
                }
                item.SubItems.Add(displayText);

                // 过滤
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
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "列管理";
            ClientSize = new Size(980, 520);
            MinimumSize = new Size(600, 300);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false;
            MaximizeBox = false;

            _btnSelectAll = new Button
            {
                Text = "全选", Location = new Point(10, 8), Size = new Size(70, 26),
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) },
                BackColor = Color.FromArgb(235, 245, 255), ForeColor = Color.FromArgb(20, 60, 100)
            };
            _btnSelectAll.Click += (s, e) => SetAllChecked(true);

            _btnDeselectAll = new Button
            {
                Text = "取消全选", Location = new Point(88, 8), Size = new Size(80, 26),
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) },
                BackColor = Color.FromArgb(235, 245, 255), ForeColor = Color.FromArgb(20, 60, 100)
            };
            _btnDeselectAll.Click += (s, e) => SetAllChecked(false);

            _txtSearch = new TextBox
            {
                Location = new Point(176, 9), Size = new Size(160, 24),
                Font = new Font("Microsoft YaHei UI", 9f), BorderStyle = BorderStyle.FixedSingle,
                PlaceholderText = "搜索列...", ForeColor = Color.FromArgb(20, 60, 100), BackColor = Color.White
            };
            _txtSearch.TextChanged += (s, e) => PopulateList(_dataTable, _frozenRows, _txtSearch.Text.Trim());

            _lvColumns = new ListView
            {
                Location = new Point(10, 42), Size = new Size(960, 420),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BorderStyle = BorderStyle.FixedSingle, Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.White, ForeColor = Color.FromArgb(20, 60, 100),
                View = View.Details, FullRowSelect = true, HideSelection = false,
                GridLines = true, CheckBoxes = true, HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            _lvColumns.ItemCheck += (s, e) =>
            {
                if (_lvColumns.Items[e.Index].Tag is int colIndex &&
                    colIndex == _lockedColumnIndex && e.NewValue == CheckState.Unchecked)
                    e.NewValue = CheckState.Checked;
            };
            _lvColumns.ItemChecked += (s, e) =>
            {
                if (e.Item.Tag is int colIndex && colIndex < _checkedStates.Length)
                    _checkedStates[colIndex] = e.Item.Checked;
            };

            _btnLinkManage = new Button
            {
                Text = "关联设置", Location = new Point(600, 468), Size = new Size(80, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right, FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(22, 163, 74), ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(21, 128, 61), MouseOverBackColor = Color.FromArgb(21, 128, 61) },
                Cursor = Cursors.Hand
            };
            _btnLinkManage.Click += OnLinkManageClick;

            _btnOK = new Button
            {
                Text = "确定", Location = new Point(690, 468), Size = new Size(80, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right, FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK, BackColor = Color.FromArgb(66, 165, 245), ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
            };

            _btnCancel = new Button
            {
                Text = "取消", Location = new Point(780, 468), Size = new Size(70, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right, FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel, BackColor = Color.FromArgb(235, 248, 255), ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };

            Controls.AddRange(new Control[] { _btnSelectAll, _btnDeselectAll, _txtSearch, _lvColumns, _btnLinkManage, _btnOK, _btnCancel });
            AcceptButton = _btnOK;
            CancelButton = _btnCancel;
            ResumeLayout(false);
        }

        private void OnLinkManageClick(object? sender, EventArgs e)
        {
            if (_lvColumns.SelectedItems.Count == 0)
            {
                AlertForm.Show(this, "请先在列表中选择要设置关联的列（点击行选中）", "提示", AlertType.Information);
                return;
            }
            var selectedItem = _lvColumns.SelectedItems[0];
            int colIndex = (int)selectedItem.Tag;
            string currentRelPath = GetRelativePath(_currentFilePath);

            using var setupDialog = new ColumnLinkSetupDialog(
                _fileEntries, currentRelPath, colIndex, _rootDir, _frozenRows,
                _columnLinks.FindAll(l =>
                    l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                    l.SourceColumnIndex == colIndex),
                _headerRowIndex);
            if (setupDialog.ShowDialog(this) == DialogResult.Yes)
            {
                _columnLinks.RemoveAll(l =>
                    l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                    l.SourceColumnIndex == colIndex);
                var newLink = setupDialog.GetResult();
                if (newLink != null)
                    _columnLinks.Add(newLink);
                UpdateLinkColumnInListView();
            }
        }

        private void UpdateLinkColumnInListView()
        {
            string currentRelPath = GetRelativePath(_currentFilePath);
            foreach (ListViewItem item in _lvColumns.Items)
            {
                if (item.Tag is int col)
                {
                    var link = _columnLinks.Find(l =>
                        l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                        l.SourceColumnIndex == col);
                    var linkText = item.SubItems[item.SubItems.Count - 1];
                    string displayText = "";
                    if (link != null)
                    {
                        int linkedFields = link.FieldLinks?.Count(f => !string.IsNullOrEmpty(f.TargetFile)) ?? 0;
                        if (linkedFields > 0)
                            displayText = $"⚡{linkedFields}字段";
                        else
                            displayText = "📋格式";
                    }
                    linkText.Text = displayText;
                }
            }
        }

        private void SetAllChecked(bool check)
        {
            for (int i = 0; i < _lvColumns.Items.Count; i++)
            {
                var item = _lvColumns.Items[i];
                if (item.Tag is int colIndex)
                {
                    if (colIndex == _lockedColumnIndex)
                    {
                        item.Checked = true;
                        _checkedStates[colIndex] = true;
                    }
                    else
                    {
                        item.Checked = check;
                        _checkedStates[colIndex] = check;
                    }
                }
            }
        }
    }
}
