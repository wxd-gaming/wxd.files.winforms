using System.Data;
using System.Text;
using ExcelDataReader;

namespace wxd.files.winforms
{
    /// <summary>列关联设置对话框 — 双击文件后弹出带固定行数据的单选框面板</summary>
    public class ColumnLinkSetupDialog : Form
    {
        private ListView _lstFiles = new();
        private TextBox _txtFilter = new();
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private Button _btnDelete = new();
        private Label _lblSelection = new();
        private List<(string display, string fullPath)> _fileEntries = new();
        private string _sourceFile = "";
        private int _sourceColumnIndex;
        private string _rootDir = "";
        private int _frozenRows = 0;
        private int _targetHeaderRowIndex = -1;       // 目标文件的表头行索引
        private ColumnLinkConfig? _existingLink;
        private string _selectedTargetPath = "";
        private int _selectedTargetColIndex = -1;
        private string _selectedTargetColumnName = "";  // 目标列的表头名称（用于显示）
        private int _selectedDisplayColIndex = -1;      // 显示列的索引（-1 表示不使用）
        private string _selectedDisplayColumnName = ""; // 显示列的表头名称
        private Button _btnSetDisplayCol = new();       // 设置显示列的按钮
        public int SelectedColumnIndex => _selectedTargetColIndex;
        public string SelectedColumnName => _selectedTargetColumnName ?? "";
        // 值提取规则控件（多层切割）
        private CheckBox _chkExtractEnabled = new();
        private ListBox _lstSteps = new();           // 切割步骤列表
        private Button _btnAddStep = new();          // 添加步骤
        private Button _btnDeleteStep = new();       // 删除步骤
        private Label _lblStepDelimiter = new();     // 分隔符标签
        private TextBox _txtStepDelimiter = new();   // 分隔符输入
        private Label _lblStepPartIndex = new();     // 取段标签
        private ComboBox _cboStepPartIndex = new();  // 取段下拉
        private Button _btnUpdateStep = new();       // 更新选中步骤
        private Label _lblExtractPreview = new();    // 预览标签
        private bool _hasUnsavedChanges = false;
        private bool _isLoadingStep = false;

        public ColumnLinkSetupDialog(
            List<(string display, string fullPath)> fileEntries,
            string sourceFile, int sourceColumnIndex,
            string rootDir, int frozenRows,
            List<ColumnLinkConfig> existingLinks,
            int targetHeaderRowIndex = -1)
        {
            _fileEntries = fileEntries;
            _sourceFile = sourceFile;
            _sourceColumnIndex = sourceColumnIndex;
            _rootDir = rootDir;
            _frozenRows = frozenRows;
            _targetHeaderRowIndex = targetHeaderRowIndex;
            _existingLink = existingLinks.FirstOrDefault();

            if (_existingLink != null)
            {
                // 尝试解析为全路径
                string resolved = _existingLink.TargetFile;
                if (!string.IsNullOrEmpty(rootDir))
                {
                    string candidate = Path.Combine(rootDir, _existingLink.TargetFile);
                    if (File.Exists(candidate))
                    {
                        resolved = candidate;
                    }
                    else
                    {
                        // 尝试在 fileEntries 中按文件名匹配
                        var match = _fileEntries.FirstOrDefault(f =>
                            Path.GetFileName(f.fullPath).Equals(Path.GetFileName(_existingLink.TargetFile), StringComparison.OrdinalIgnoreCase));
                        if (match.fullPath != null)
                            resolved = match.fullPath;
                    }
                }
                _selectedTargetPath = resolved;
                _selectedTargetColIndex = _existingLink.TargetColumnIndex;

                // 恢复目标列的表头名称（如果有保存）
                if (!string.IsNullOrEmpty(_existingLink.TargetColumnName))
                    _selectedTargetColumnName = _existingLink.TargetColumnName;
                // 恢复显示列设置
                _selectedDisplayColIndex = _existingLink.TargetDisplayColumnIndex;
                _selectedDisplayColumnName = _existingLink.TargetDisplayColumnName;
            }

            InitializeComponent();
            PopulateFiles("");
            UpdateSelectionLabel();

            // 加载已有配置（必须在 InitializeComponent 之后，否则控件会被重新创建覆盖）
            if (_existingLink != null)
            {
                _chkExtractEnabled.Checked = _existingLink.ValueExtractionEnabled;
                _existingLink.MigrateFromOldFields();
                // 填充 ListBox
                RefreshStepList();
                if (_lstSteps.Items.Count > 0)
                    _lstSteps.SelectedIndex = 0;
            }

            if (_existingLink != null)
            {
                string targetRel = _existingLink.TargetFile.Replace("\\", "/");
                foreach (ListViewItem item in _lstFiles.Items)
                {
                    if (item.Tag is string fp)
                    {
                        string fpNorm = fp.Replace("\\", "/");
                        if (fpNorm.EndsWith(targetRel, StringComparison.OrdinalIgnoreCase) ||
                            Path.GetFileName(fp).Equals(Path.GetFileName(_existingLink.TargetFile), StringComparison.OrdinalIgnoreCase))
                        {
                            item.Selected = true;
                            _lstFiles.EnsureVisible(item.Index);
                            break;
                        }
                    }
                }
            }
        }

        public ColumnLinkConfig? GetResult()
        {
            if (string.IsNullOrEmpty(_selectedTargetPath) || _selectedTargetColIndex < 0)
                return null;
            string relPath;
            if (!string.IsNullOrEmpty(_rootDir) && _selectedTargetPath.StartsWith(_rootDir, StringComparison.OrdinalIgnoreCase))
                relPath = _selectedTargetPath.Substring(_rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            else
                relPath = Path.GetFileName(_selectedTargetPath);
            return new ColumnLinkConfig
            {
                SourceFile = _sourceFile,
                SourceColumnIndex = _sourceColumnIndex,
                TargetFile = relPath,
                TargetColumnIndex = _selectedTargetColIndex,
                TargetColumnName = _selectedTargetColumnName,  // 保存表头名称
                TargetDisplayColumnIndex = _selectedDisplayColIndex,
                TargetDisplayColumnName = _selectedDisplayColumnName,
                // 值提取规则
                ValueExtractionEnabled = _chkExtractEnabled.Checked,
                ExtractSteps = GetExtractStepsFromUI()
            };
        }

        private void UpdateSelectionLabel()
        {
            if (!string.IsNullOrEmpty(_selectedTargetPath) && _selectedTargetColIndex >= 0)
            {
                string displayPath = _selectedTargetPath;
                if (!string.IsNullOrEmpty(_rootDir) && _selectedTargetPath.StartsWith(_rootDir, StringComparison.OrdinalIgnoreCase))
                    displayPath = _selectedTargetPath.Substring(_rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
                else
                    displayPath = Path.GetFileName(_selectedTargetPath);
                string colInfo = !string.IsNullOrEmpty(_selectedTargetColumnName)
                    ? $"[{_selectedTargetColumnName}]"
                    : $"列{_selectedTargetColIndex}";
                string displayInfo = _selectedDisplayColIndex >= 0
                    ? (!string.IsNullOrEmpty(_selectedDisplayColumnName) ? $" (显示: {_selectedDisplayColumnName})" : $" (显示: 列{_selectedDisplayColIndex})")
                    : "";
                _lblSelection.Text = $"已选: {displayPath} → {colInfo}{displayInfo}";
            }
            else
                _lblSelection.Text = "未选择关联列（双击目标文件设置）";
        }

        private void PopulateFiles(string filter)
        {
            _lstFiles.BeginUpdate();
            _lstFiles.Items.Clear();
            var query = string.IsNullOrEmpty(filter)
                ? _fileEntries
                : _fileEntries.Where(f => f.display.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
            foreach (var entry in query.OrderBy(f => f.display))
            {
                _lstFiles.Items.Add(new ListViewItem(entry.display) { Tag = entry.fullPath, ToolTipText = entry.fullPath });
            }
            _lstFiles.EndUpdate();
        }

        /// <summary>双击文件，弹出列选择面板（带固定行数据 + 单选框）</summary>
        private void OnFileDoubleClick(object? sender, EventArgs e)
        {
            if (_lstFiles.SelectedItems.Count == 0) return;
            var fullPath = _lstFiles.SelectedItems[0].Tag as string ?? "";
            if (!File.Exists(fullPath))
            {
                AlertForm.Show(this, "文件不存在", "错误", AlertType.Error);
                return;
            }

            // 读取目标文件固定行数据用于展示
            DataTable targetFrozenTable;
            try
            {
                var ext = Path.GetExtension(fullPath).ToLowerInvariant();
                if (ext == ".csv")
                {
                    // ★ 使用 FileDataCache 统一读取（支持 FileShare.ReadWrite）
                    var readOnlyData = FileDataCache.GetDataTableReadOnly(fullPath);

                    var dt = new DataTable();
                    if (readOnlyData.Columns.Count > 0)
                    {
                        for (int i = 0; i < readOnlyData.Columns.Count; i++)
                            dt.Columns.Add($"Col{i}", typeof(string));
                        int maxRows = Math.Min(_frozenRows > 0 ? _frozenRows : 2, readOnlyData.Rows.Count);
                        for (int r = 0; r < maxRows; r++)
                        {
                            var row = dt.NewRow();
                            for (int i = 0; i < Math.Min(readOnlyData.Columns.Count, dt.Columns.Count); i++)
                                row[i] = readOnlyData.Rows[r][i]?.ToString() ?? "";
                            dt.Rows.Add(row);
                        }
                    }
                    targetFrozenTable = dt;
                }
                else
                {
                    try
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using var fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        using var excelReader = ExcelReaderFactory.CreateReader(fs);
                        var dataSet = excelReader.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
                        });
                        if (dataSet.Tables.Count > 0)
                        {
                            var sheet = dataSet.Tables[0];
                            int maxCols = sheet.Columns.Count;
                            var dt = new DataTable();
                            for (int i = 0; i < maxCols; i++)
                                dt.Columns.Add($"Col{i}", typeof(string));
                            int maxRows = Math.Min(_frozenRows > 0 ? _frozenRows : 2, sheet.Rows.Count);
                            for (int r = 0; r < maxRows; r++)
                            {
                                var dtRow = dt.NewRow();
                                for (int i = 0; i < maxCols; i++)
                                {
                                    dtRow[i] = sheet.Rows[r][i]?.ToString() ?? "";
                                }
                                dt.Rows.Add(dtRow);
                            }
                            targetFrozenTable = dt;
                        }
                        else targetFrozenTable = new DataTable();
                    }
                    catch
                    {
                        targetFrozenTable = new DataTable();
                    }
                }
            }
            catch
            {
                AlertForm.Show(this, "无法读取目标文件数据", "错误", AlertType.Error);
                return;
            }

            int initialSelection = -1;
            if (_existingLink != null)
            {
                string targetRel = _existingLink.TargetFile.Replace("\\", "/");
                string fullNorm = fullPath.Replace("\\", "/");
                if (fullNorm.EndsWith(targetRel, StringComparison.OrdinalIgnoreCase) ||
                    Path.GetFileName(fullPath).Equals(Path.GetFileName(_existingLink.TargetFile), StringComparison.OrdinalIgnoreCase))
                {
                    initialSelection = _existingLink.TargetColumnIndex;
                }
            }

            using var colDlg = new TargetColumnSelectDialog(targetFrozenTable, _frozenRows, initialSelection,
                _targetHeaderRowIndex >= 0 ? _targetHeaderRowIndex : GetTargetHeaderRowIndexStatic(fullPath));
            if (colDlg.ShowDialog(this) == DialogResult.OK && colDlg.SelectedColumnIndex >= 0)
            {
                _selectedTargetPath = fullPath;
                _selectedTargetColIndex = colDlg.SelectedColumnIndex;
                _selectedTargetColumnName = colDlg.SelectedColumnName;  // 保存表头名称
                UpdateSelectionLabel();
            }
        }

        private void OnDeleteClick(object? sender, EventArgs e)
        {
            _selectedTargetPath = "";
            _selectedTargetColIndex = -1;
            _selectedDisplayColIndex = -1;
            _selectedDisplayColumnName = "";
            _existingLink = null;
            _btnDelete.Visible = false;
            UpdateSelectionLabel();
        }

        /// <summary>设置显示列 — 在目标文件中再选一列作为显示名</summary>
        private void OnSetDisplayColClick(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedTargetPath) || _selectedTargetColIndex < 0)
            {
                AlertForm.Show(this, "请先选择关联列（双击目标文件）", "提示", AlertType.Information);
                return;
            }
            if (!File.Exists(_selectedTargetPath))
            {
                AlertForm.Show(this, "目标文件不存在", "错误", AlertType.Error);
                return;
            }

            // 读取目标文件固定行数据用于展示
            DataTable targetFrozenTable;
            try
            {
                var ext = Path.GetExtension(_selectedTargetPath).ToLowerInvariant();
                if (ext == ".csv")
                {
                    // ★ 使用 FileDataCache 统一读取
                    var readOnlyData = FileDataCache.GetDataTableReadOnly(_selectedTargetPath);

                    var dt = new DataTable();
                    if (readOnlyData.Columns.Count > 0)
                    {
                        for (int i = 0; i < readOnlyData.Columns.Count; i++)
                            dt.Columns.Add($"Col{i}", typeof(string));
                        int maxRows = Math.Min(_frozenRows > 0 ? _frozenRows : 2, readOnlyData.Rows.Count);
                        for (int r = 0; r < maxRows; r++)
                        {
                            var row = dt.NewRow();
                            for (int i = 0; i < Math.Min(readOnlyData.Columns.Count, dt.Columns.Count); i++)
                                row[i] = readOnlyData.Rows[r][i]?.ToString() ?? "";
                            dt.Rows.Add(row);
                        }
                    }
                    targetFrozenTable = dt;
                }
                else
                {
                    try
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using var fs = new FileStream(_selectedTargetPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        using var excelReader = ExcelReaderFactory.CreateReader(fs);
                        var dataSet = excelReader.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
                        });
                        if (dataSet.Tables.Count > 0)
                        {
                            var sheet = dataSet.Tables[0];
                            int maxCols = sheet.Columns.Count;
                            var dt = new DataTable();
                            for (int i = 0; i < maxCols; i++)
                                dt.Columns.Add($"Col{i}", typeof(string));
                            int maxRows = Math.Min(_frozenRows > 0 ? _frozenRows : 2, sheet.Rows.Count);
                            for (int r = 0; r < maxRows; r++)
                            {
                                var dtRow = dt.NewRow();
                                for (int i = 0; i < maxCols; i++)
                                    dtRow[i] = sheet.Rows[r][i]?.ToString() ?? "";
                                dt.Rows.Add(dtRow);
                            }
                            targetFrozenTable = dt;
                        }
                        else targetFrozenTable = new DataTable();
                    }
                    catch { targetFrozenTable = new DataTable(); }
                }
            }
            catch
            {
                AlertForm.Show(this, "无法读取目标文件数据", "错误", AlertType.Error);
                return;
            }

            using var colDlg = new TargetColumnSelectDialog(targetFrozenTable, _frozenRows, _selectedDisplayColIndex,
                _targetHeaderRowIndex >= 0 ? _targetHeaderRowIndex : GetTargetHeaderRowIndexStatic(_selectedTargetPath),
                title: "选择显示列 — 用于提示中显示中文名称");
            if (colDlg.ShowDialog(this) == DialogResult.OK && colDlg.SelectedColumnIndex >= 0)
            {
                _selectedDisplayColIndex = colDlg.SelectedColumnIndex;
                _selectedDisplayColumnName = colDlg.SelectedColumnName;
                _btnSetDisplayCol.Text = !string.IsNullOrEmpty(_selectedDisplayColumnName)
                    ? $"显示列: {_selectedDisplayColumnName}"
                    : $"显示列: 列{_selectedDisplayColIndex}";
                UpdateSelectionLabel();
            }
        }

        // ==================== 切割步骤管理 ====================
        /// <summary>确保 _existingLink 不为 null（用于首次添加/更新切割步骤时自动创建）</summary>
        private void EnsureExistingLink()
        {
            if (_existingLink == null)
            {
                _existingLink = new ColumnLinkConfig
                {
                    SourceFile = _sourceFile,
                    SourceColumnIndex = _sourceColumnIndex,
                    ExtractSteps = new List<ExtractStep>()
                };
            }
        }

        private void RefreshStepList()
        {
            _lstSteps.BeginUpdate();
            _lstSteps.Items.Clear();
            var steps = _existingLink?.ExtractSteps ?? new List<ExtractStep>();
            for (int i = 0; i < steps.Count; i++)
            {
                var step = steps[i];
                string desc = step.PartIndex >= 0
                    ? $"{i + 1}. 用 \"{step.Delimiter}\" 切割 → 取第 {step.PartIndex} 段"
                    : $"{i + 1}. 用 \"{step.Delimiter}\" 切割 → 保留全部";
                _lstSteps.Items.Add(desc);
            }
            _lstSteps.EndUpdate();
        }

        private void MarkUnsaved()
        {
            if (!_isLoadingStep && _lstSteps.SelectedIndex >= 0)
                _hasUnsavedChanges = true;
        }

        private void OnStepSelected(object? sender, EventArgs e)
        {
            _isLoadingStep = true;
            if (_lstSteps.SelectedIndex < 0) { _isLoadingStep = false; return; }
            var steps = _existingLink?.ExtractSteps;
            if (steps == null || _lstSteps.SelectedIndex >= steps.Count) { _isLoadingStep = false; return; }
            var step = steps[_lstSteps.SelectedIndex];

            _txtStepDelimiter.Text = step.Delimiter;
            // 映射 PartIndex 到 ComboBox 索引：-1→0, 0→1, 1→2, ...
            int cbIdx = step.PartIndex + 1;
            if (cbIdx >= 0 && cbIdx < _cboStepPartIndex.Items.Count)
                _cboStepPartIndex.SelectedIndex = cbIdx;
            _isLoadingStep = false;
        }

        private void AddStep()
        {
            EnsureExistingLink();
            _existingLink!.MigrateFromOldFields();
            _existingLink.ExtractSteps.Add(new ExtractStep { Delimiter = "#", PartIndex = 0 });
            RefreshStepList();
            _lstSteps.SelectedIndex = _lstSteps.Items.Count - 1;
            UpdatePreview();
            _hasUnsavedChanges = false;
        }

        private void DeleteStep()
        {
            if (_existingLink == null || _lstSteps.SelectedIndex < 0) return;
            _existingLink.ExtractSteps.RemoveAt(_lstSteps.SelectedIndex);
            RefreshStepList();
            if (_lstSteps.Items.Count > 0)
                _lstSteps.SelectedIndex = 0;
            UpdatePreview();
            _hasUnsavedChanges = false;
        }

        private void UpdateStep()
        {
            EnsureExistingLink();
            if (_lstSteps.SelectedIndex < 0) return;
            var steps = _existingLink!.ExtractSteps;
            if (_lstSteps.SelectedIndex >= steps.Count) return;

            var step = steps[_lstSteps.SelectedIndex];
            step.Delimiter = _txtStepDelimiter.Text.Trim();
            if (string.IsNullOrEmpty(step.Delimiter)) step.Delimiter = "#";

            int cbIdx = _cboStepPartIndex.SelectedIndex;
            step.PartIndex = cbIdx - 1; // 0→-1, 1→0, 2→1, ...

            RefreshStepList();
            _lstSteps.SelectedIndex = Math.Min(_lstSteps.SelectedIndex, _lstSteps.Items.Count - 1);
            UpdatePreview();
            _hasUnsavedChanges = false;
        }

        private List<ExtractStep> GetExtractStepsFromUI()
        {
            if (_existingLink != null)
            {
                _existingLink.MigrateFromOldFields();
                return _existingLink.ExtractSteps;
            }
            // 默认
            return new List<ExtractStep>
            {
                new ExtractStep { Delimiter = "|", PartIndex = -1 },
                new ExtractStep { Delimiter = "#", PartIndex = 0 }
            };
        }

        private void OnOKClick(object? sender, EventArgs e)
        {
            if (_hasUnsavedChanges)
            {
                var result = AlertForm.Show(this,
                    "切割步骤有未保存的修改（未点击\"更新\"按钮），是否放弃修改？" + GlobalConst.NewLine + GlobalConst.NewLine + "点击\"是\"将放弃修改并关闭对话框。" + GlobalConst.NewLine + "点击\"否\"返回继续编辑。",
                    "未保存的修改", AlertType.Warning);
                if (result == DialogResult.No)
                    return;
            }
            DialogResult = DialogResult.OK;
            Close();
        }

        private void UpdatePreview()
        {
            var steps = GetExtractStepsFromUI();
            if (!_chkExtractEnabled.Checked)
            {
                _lblExtractPreview.Text = "预览: 未启用提取";
                return;
            }
            if (steps.Count == 0)
            {
                _lblExtractPreview.Text = "预览: 已启用提取，但无切割步骤（请添加）";
                return;
            }

            // 构建步骤描述
            var stepDescs = new List<string>();
            foreach (var s in steps)
            {
                stepDescs.Add(s.PartIndex >= 0 ? $"\"{s.Delimiter}\"切取{s.PartIndex}段" : $"\"{s.Delimiter}\"切全部");
            }
            string flow = string.Join(" → ", stepDescs);

            // 模拟提取
            var tempConfig = new ColumnLinkConfig
            {
                ValueExtractionEnabled = true,
                ExtractSteps = steps
            };
            string sample1 = "2#1|3#1";
            string sample2 = "1:2#1|2:3#1";
            var result1 = tempConfig.ExtractValues(sample1);
            var result2 = tempConfig.ExtractValues(sample2);
            string r1 = $"[{string.Join(",", result1.Select(v => $"\"{v}\""))}]";
            string r2 = $"[{string.Join(",", result2.Select(v => $"\"{v}\""))}]";

            _lblExtractPreview.Text = $"预览: {flow}  |  例 \"{sample1}\"→{r1}  \"{sample2}\"→{r2}";
        }
        // ==================== 结束切割步骤管理 ====================

        private void InitializeComponent()
        {
            SuspendLayout();
            Text = "列关联设置 - 列 " + _sourceColumnIndex;
            ClientSize = new Size(500, 630);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false; MaximizeBox = false;

            Controls.Add(new Label { Text = "选择关联的目标文件（双击设置列）:", Location = new Point(10, 10), Size = new Size(480, 20),
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(31, 41, 55) });

            _txtFilter = new TextBox { Location = new Point(10, 32), Size = new Size(480, 23),
                Font = new Font("Microsoft YaHei UI", 9f), BorderStyle = BorderStyle.FixedSingle, PlaceholderText = "输入关键字筛选文件..." };
            _txtFilter.TextChanged += (s, e) => PopulateFiles(_txtFilter.Text.Trim());
            Controls.Add(_txtFilter);

            _lstFiles = new ListView { Location = new Point(10, 62), Size = new Size(480, 200), BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f), View = View.Details, FullRowSelect = true, HideSelection = false,
                HeaderStyle = ColumnHeaderStyle.None, MultiSelect = false };
            _lstFiles.Columns.Add("文件", 460);
            _lstFiles.DoubleClick += OnFileDoubleClick;
            Controls.Add(_lstFiles);

            _lblSelection = new Label { Text = "未选择关联列（双击目标文件设置）", Location = new Point(10, 272),
                Size = new Size(480, 40), Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = Color.FromArgb(31, 41, 55), BackColor = Color.FromArgb(235, 248, 255),
                BorderStyle = BorderStyle.FixedSingle, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(4, 0, 0, 0) };
            Controls.Add(_lblSelection);

            // 设置显示列按钮
            _btnSetDisplayCol = new Button
            {
                Text = _selectedDisplayColIndex >= 0
                    ? (!string.IsNullOrEmpty(_selectedDisplayColumnName)
                        ? $"显示列: {_selectedDisplayColumnName}"
                        : $"显示列: 列{_selectedDisplayColIndex}")
                    : "设置显示列（可选）",
                Location = new Point(10, 320), Size = new Size(200, 28),
                Font = new Font("Microsoft YaHei UI", 8f),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) },
                Cursor = Cursors.Hand
            };
            _btnSetDisplayCol.Click += OnSetDisplayColClick;
            Controls.Add(_btnSetDisplayCol);

            Controls.Add(new Label { Text = "提示: 关联基于列的位置索引。双击目标文件后在弹出面板中单选目标列。",
                Location = new Point(220, 322), Size = new Size(270, 30), Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(107, 114, 128) });

            // ===== 值提取规则区域 =====
            var grpExtract = new GroupBox
            {
                Text = "值提取规则（可选）— 多层切割：从复合格式中提取部分内容匹配",
                Location = new Point(10, 355), Size = new Size(480, 175),
                Font = new Font("Microsoft YaHei UI", 8f), ForeColor = Color.FromArgb(55, 65, 81)
            };

            _chkExtractEnabled = new CheckBox
            {
                Text = "启用提取（如 \"4002#1\" → 提取 \"4002\" 匹配）",
                Location = new Point(10, 20), Size = new Size(300, 22),
                Font = new Font("Microsoft YaHei UI", 8f), Checked = false,
                ForeColor = Color.FromArgb(31, 41, 55)
            };
            _chkExtractEnabled.CheckedChanged += (s, e) =>
            {
                bool en = _chkExtractEnabled.Checked;
                _lstSteps.Enabled = en;
                _btnAddStep.Enabled = en;
                _btnDeleteStep.Enabled = en;
                _txtStepDelimiter.Enabled = en;
                _cboStepPartIndex.Enabled = en;
                _btnUpdateStep.Enabled = en;
                _lblExtractPreview.Visible = en;
                if (en) UpdatePreview();
            };
            grpExtract.Controls.Add(_chkExtractEnabled);

            // ===== 切割步骤列表 =====
            _lstSteps = new ListBox
            {
                Location = new Point(10, 42), Size = new Size(350, 70),
                Font = new Font("Microsoft YaHei UI", 8f), Enabled = false,
                BorderStyle = BorderStyle.FixedSingle, IntegralHeight = false,
                BackColor = Color.FromArgb(250, 252, 255)
            };
            _lstSteps.SelectedIndexChanged += OnStepSelected;
            grpExtract.Controls.Add(_lstSteps);

            // 添加/删除按钮
            _btnAddStep = new Button
            {
                Text = "＋添加", Location = new Point(368, 42), Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 7.5f), Enabled = false,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(235, 248, 255), ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };
            _btnAddStep.Click += (s, e) => { AddStep(); };
            grpExtract.Controls.Add(_btnAddStep);

            _btnDeleteStep = new Button
            {
                Text = "✕删除", Location = new Point(422, 42), Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 7.5f), Enabled = false,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(255, 240, 240), ForeColor = Color.FromArgb(180, 40, 40),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(240, 200, 200) }
            };
            _btnDeleteStep.Click += (s, e) => { DeleteStep(); };
            grpExtract.Controls.Add(_btnDeleteStep);

            // ===== 编辑选中步骤 =====
            _lblStepDelimiter = new Label
            {
                Text = "分隔符:", Location = new Point(10, 118), Size = new Size(50, 18),
                Font = new Font("Microsoft YaHei UI", 8f), Enabled = false,
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            grpExtract.Controls.Add(_lblStepDelimiter);

            _txtStepDelimiter = new TextBox
            {
                Text = "", Location = new Point(60, 116), Size = new Size(40, 21),
                Font = new Font("Microsoft YaHei UI", 8f), BorderStyle = BorderStyle.FixedSingle,
                Enabled = false, PlaceholderText = "# | :"
            };
            _txtStepDelimiter.TextChanged += (s, e) => MarkUnsaved();
            grpExtract.Controls.Add(_txtStepDelimiter);

            _lblStepPartIndex = new Label
            {
                Text = "取段:", Location = new Point(108, 118), Size = new Size(35, 18),
                Font = new Font("Microsoft YaHei UI", 8f), Enabled = false,
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            grpExtract.Controls.Add(_lblStepPartIndex);

            _cboStepPartIndex = new ComboBox
            {
                Location = new Point(143, 116), Size = new Size(100, 21),
                Font = new Font("Microsoft YaHei UI", 8f), Enabled = false,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat
            };
            _cboStepPartIndex.Items.AddRange(new object[]
            {
                "-1 = 保留全部",
                "0 = 取第 0 段",
                "1 = 取第 1 段",
                "2 = 取第 2 段",
                "3 = 取第 3 段",
                "4 = 取第 4 段",
                "5 = 取第 5 段",
                "6 = 取第 6 段",
                "7 = 取第 7 段",
                "8 = 取第 8 段",
                "9 = 取第 9 段"
            });
            _cboStepPartIndex.SelectedIndex = 1; // 默认取第0段
            _cboStepPartIndex.SelectedIndexChanged += (s, e) => MarkUnsaved();
            grpExtract.Controls.Add(_cboStepPartIndex);

            _btnUpdateStep = new Button
            {
                Text = "更新", Location = new Point(250, 114), Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 7.5f), Enabled = false,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(66, 165, 245), ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
            };
            _btnUpdateStep.Click += (s, e) => { UpdateStep(); };
            grpExtract.Controls.Add(_btnUpdateStep);

            // ===== 预览 =====
            _lblExtractPreview = new Label
            {
                Text = "预览: 输入示例 \"2#1|3#1\"  →  结果: [\"2\",\"3\"]",
                Location = new Point(10, 142), Size = new Size(460, 18),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = Color.FromArgb(140, 145, 150), Visible = false
            };
            grpExtract.Controls.Add(_lblExtractPreview);
            // ===== 结束提取规则区域 =====

            Controls.Add(grpExtract);

            _btnDelete = new Button { Text = "删除关联", Location = new Point(10, 540), Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(239, 68, 68), ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(220, 50, 50) }, Cursor = Cursors.Hand,
                Visible = _existingLink != null || !string.IsNullOrEmpty(_selectedTargetPath) };
            _btnDelete.Click += OnDeleteClick;
            Controls.Add(_btnDelete);

            _btnOK = new Button { Text = "确定", Location = new Point(330, 540), Size = new Size(80, 28), FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(66, 165, 245), ForeColor = Color.White, FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                Cursor = Cursors.Hand };
            _btnOK.Click += OnOKClick;
            Controls.Add(_btnOK);

            _btnCancel = new Button { Text = "取消", Location = new Point(420, 540), Size = new Size(70, 28), FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel, BackColor = Color.FromArgb(235, 248, 255), ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) } };
            Controls.Add(_btnCancel);

            AcceptButton = _btnOK; CancelButton = _btnCancel;
            ResumeLayout(false);
        }

        // ★ SplitCsvLinesStatic 和 ParseCsvLineStatic 已删除，统一使用 DirectoryCache.SplitCsvLines / ParseCsvLine

        /// <summary>根据目标文件路径获取其表头行索引（从设置中读取）</summary>
        private static int GetTargetHeaderRowIndexStatic(string targetFullPath)
        {
            if (string.IsNullOrEmpty(targetFullPath)) return -1;
            string targetDir = Path.GetDirectoryName(targetFullPath) ?? "";
            // 如果在同一目录下，使用该目录的设置
            var ext = Path.GetExtension(targetFullPath).ToLowerInvariant();
            if (ext == ".csv")
                return CsvSettings.GetCsvHeaderRowIndexByDirectory(targetDir);
            else if (ext == ".xls" || ext == ".xlsx")
                return CsvSettings.GetExcelHeaderRowIndexByDirectory(targetDir);
            return -1;
        }
    }
}
