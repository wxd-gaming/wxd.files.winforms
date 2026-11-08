using System.Data;

namespace wxd.files.winforms
{
    /// <summary>列管理对话框 — ListView列显隐 + 关联设置</summary>
    public class ColumnManagerDialog : Form
    {
        private ListView _lvColumns = new();
        private ThemeButton _btnSelectAll = new();
        private ThemeButton _btnDeselectAll = new();
        private ThemeButton _btnOK = new();
        private ThemeButton _btnCancel = new();
        private ThemeButton _btnLinkManage = new();
        private ThemeButton _btnTemplateRef = new();
        private ThemeButton _btnTemplateManage = new();
        private ThemeField _fldSearch = new();
        private Panel _bar = new();
        private Panel _topBar = new();
        private Panel _topDivider = new();
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
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
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
                    if (!string.IsNullOrEmpty(link.TemplateUid))
                    {
                        var template = LinkTemplateManager.GetTemplate(link.TemplateUid);
                        if (template != null)
                            displayText = $"📋{template.Name}";
                        else
                            displayText = "📋(模板已删除)";
                    }
                    else
                    {
                        int linkedFields = link.FieldLinks?.Count(f => !string.IsNullOrEmpty(f.TargetFile)) ?? 0;
                        if (linkedFields > 0)
                            displayText = $"⚡{linkedFields}字段";
                        else
                            displayText = "📋格式";
                    }
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
            MinimizeBox = false;
            MaximizeBox = false;

            // ==================== 顶部工具区 ====================
            _topBar = new Panel { Dock = DockStyle.Top, Height = ThemeMetrics.BarHeight };
            _topDivider = new Panel { Dock = DockStyle.Bottom, Height = 1 };
            _topBar.Controls.Add(_topDivider);

            _btnSelectAll = DialogButtonBar.Make("全选", (s, e) => SetAllChecked(true));
            _btnDeselectAll = DialogButtonBar.Make("取消全选", (s, e) => SetAllChecked(false));

            _fldSearch = new ThemeField { Placeholder = "搜索列..." };
            _fldSearch.Edit.TextChanged += (s, e) => PopulateList(_dataTable, _frozenRows, _fldSearch.Edit.Text.Trim());

            _topBar.Controls.AddRange(new Control[] { _btnSelectAll, _btnDeselectAll, _fldSearch });
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
                CheckBoxes = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
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

            // ==================== 底部按钮条 ====================
            // 一个对话框只留一颗实心（确定），其余动作一律描边，别再各涂各的品牌蓝和绿
            _bar = DialogButtonBar.Create();
            _btnTemplateManage = DialogButtonBar.Make("管理模板", OnTemplateManageClick);
            _btnTemplateRef = DialogButtonBar.Make("引用模板", OnTemplateRefClick);
            _btnLinkManage = DialogButtonBar.Make("关联设置", OnLinkManageClick);
            _btnCancel = DialogButtonBar.Make("取消");
            _btnCancel.DialogResult = DialogResult.Cancel;
            _btnOK = DialogButtonBar.Make("确定", null, primary: true);
            _btnOK.DialogResult = DialogResult.OK;
            _bar.Controls.AddRange(new Control[]
            {
                _btnTemplateManage, _btnTemplateRef, _btnLinkManage, _btnCancel, _btnOK
            });
            LayoutBar();
            _bar.Resize += (_, _) => LayoutBar();

            // 停靠顺序：后加入的靠外，所以列表先加、页眉页脚后加
            Controls.Add(_lvColumns);
            Controls.Add(_topBar);
            Controls.Add(_bar);
            AcceptButton = _btnOK;
            CancelButton = _btnCancel;
            ResumeLayout(false);
        }

        private void LayoutTopBar()
        {
            int y = ThemeMetrics.RowPad, x = ThemeMetrics.Margin;
            int rowH = ThemeMetrics.RowH;

            _btnSelectAll.SetBounds(x, y, ThemeButton.FitWidth(_btnSelectAll.Text), rowH);
            x += _btnSelectAll.Width + ThemeMetrics.Gap;
            _btnDeselectAll.SetBounds(x, y, ThemeButton.FitWidth(_btnDeselectAll.Text), rowH);
            x += _btnDeselectAll.Width + ThemeMetrics.Group;
            _fldSearch.SetBounds(x, y, 220, rowH);
        }

        private void LayoutBar()
            => DialogButtonBar.LayoutRight(_bar, _btnTemplateManage, _btnTemplateRef, _btnLinkManage, _btnCancel, _btnOK);

        /// <summary>取色全部走当前皮肤令牌；换皮肤或翻明暗时由 Theme.Changed 重新走一遍。</summary>
        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _topBar.BackColor = p.Chrome;
            _bar.BackColor = p.Chrome;
            _topDivider.BackColor = p.Divider;

            _lvColumns.BackColor = p.Surface;
            _lvColumns.ForeColor = p.Text;
            _lvColumns.Font = Theme.List;

            _fldSearch.ApplyTheme();
            foreach (var b in new[] { _btnSelectAll, _btnDeselectAll, _btnTemplateManage, _btnTemplateRef,
                                      _btnLinkManage, _btnCancel, _btnOK })
            {
                b.BackColor = p.Chrome;
                b.Font = Theme.Body;
            }

            LayoutTopBar();
            LayoutBar();
            Invalidate(true);
        }

        private void OnTemplateManageClick(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_rootDir))
            {
                AlertForm.Show(this, "未设置目录路径，无法管理模板", "提示", AlertType.Information);
                return;
            }
            using var dialog = new LinkTemplateDialog(mode: false, _fileEntries, _frozenRows, _headerRowIndex, _rootDir);
            dialog.ShowDialog(this);
            // 如果模板有变更，刷新关联列的显示
            if (dialog.TemplatesChanged)
            {
                UpdateLinkColumnInListView();
            }
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

            // 查找已有配置，传入深拷贝副本
            var existingLink = _columnLinks.Find(l =>
                l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == colIndex);
            var linkCopy = existingLink?.DeepClone();

            using var setupDialog = new ColumnLinkSetupDialog(
                _fileEntries, currentRelPath, colIndex, _rootDir, _frozenRows,
                linkCopy, _headerRowIndex);
            if (setupDialog.ShowDialog(this) == DialogResult.Yes)
            {
                var newLink = setupDialog.GetResult();
                if (newLink != null)
                {
                    // 删除旧配置，用返回的对象替换
                    _columnLinks.RemoveAll(l =>
                        l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                        l.SourceColumnIndex == colIndex);
                    _columnLinks.Add(newLink);
                }
                else
                {
                    // 返回 null 表示删除了配置
                    _columnLinks.RemoveAll(l =>
                        l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                        l.SourceColumnIndex == colIndex);
                }
                UpdateLinkColumnInListView();
            }
        }

        private void OnTemplateRefClick(object? sender, EventArgs e)
        {
            if (_lvColumns.SelectedItems.Count == 0)
            {
                AlertForm.Show(this, "请先在列表中选择要引用模板的列（点击行选中）", "提示", AlertType.Information);
                return;
            }
            if (string.IsNullOrEmpty(_rootDir))
            {
                AlertForm.Show(this, "未设置目录路径，无法管理模板", "提示", AlertType.Information);
                return;
            }

            using var dialog = new LinkTemplateDialog(mode: true, _fileEntries, _frozenRows, _headerRowIndex, _rootDir);
            if (dialog.ShowDialog(this) == DialogResult.OK && dialog.IsTemplateSelected)
            {
                var selectedItem = _lvColumns.SelectedItems[0];
                int colIndex = (int)selectedItem.Tag;
                string currentRelPath = GetRelativePath(_currentFilePath);

                var template = LinkTemplateManager.GetTemplate(dialog.SelectedTemplateUid!);
                if (template == null) return;

                // 删除当前列已有配置
                _columnLinks.RemoveAll(l =>
                    l.SourceFile.Equals(currentRelPath, StringComparison.OrdinalIgnoreCase) &&
                    l.SourceColumnIndex == colIndex);

                // 创建引用模板的配置 — 只记 TemplateUid，不复制配置，运行时从模板实时读取
                var newLink = new ColumnLinkConfig
                {
                    SourceFile = currentRelPath,
                    SourceColumnIndex = colIndex,
                    TemplateUid = template.Uid,
                    // 不再调用 ApplyTemplate，配置从模板实时读取
                };
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
                        if (!string.IsNullOrEmpty(link.TemplateUid))
                        {
                            var template = LinkTemplateManager.GetTemplate(link.TemplateUid);
                            if (template != null)
                                displayText = $"📋{template.Name}";
                            else
                                displayText = "📋(模板已删除)";
                        }
                        else
                        {
                            int linkedFields = link.FieldLinks?.Count(f => !string.IsNullOrEmpty(f.TargetFile)) ?? 0;
                            if (linkedFields > 0)
                                displayText = $"⚡{linkedFields}字段";
                            else
                                displayText = "📋格式";
                        }
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
