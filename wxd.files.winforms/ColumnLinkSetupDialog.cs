using System.Data;
using System.Text;
using ExcelDataReader;

namespace wxd.files.winforms
{
    /// <summary>Tips 显示设置对话框 — 先设置切割方案，再设置表头，然后各段关联关系</summary>
    public class ColumnLinkSetupDialog : Form
    {
        // ===== 表头输入 =====
        private TextBox _txtDisplayHeaders = new();
        // ===== 值提取规则控件 =====
        private CheckBox _chkExtractEnabled = new();
        private ListBox _lstSteps = new();
        private Button _btnAddStep = new();
        private Button _btnDeleteStep = new();
        private Label _lblStepDelimiter = new();
        private TextBox _txtStepDelimiter = new();
        private Label _lblStepPartIndex = new();
        private ComboBox _cboStepPartIndex = new();
        private Button _btnUpdateStep = new();
        private Label _lblExtractPreview = new();
        private bool _hasUnsavedChanges = false;
        private bool _isLoadingStep = false;
        // ===== 多字段关联列表 =====
        private ListBox _lstFieldLinks = new();
        private Button _btnAddFieldLink = new();   // 添加新段
        private Button _btnDeleteFieldLink = new(); // 删除选中段配置
        private Button _btnUpdateFieldLink = new(); // 更新选中段配置
        private Label _lblFieldLinkInfo = new();
        private Button _btnSelectFieldFile = new();
        private Button _btnSelectFieldCol = new();
        private Button _btnSelectFieldDisplayCol = new();
        private int _editingFieldIndex = -1; // 当前编辑的段索引
        // ===== 按钮 =====
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private Button _btnDelete = new();
        // ===== 状态 =====
        private List<(string display, string fullPath)> _fileEntries = new();
        private string _sourceFile = "";
        private int _sourceColumnIndex;
        private string _rootDir = "";
        private int _frozenRows = 0;
        private int _targetHeaderRowIndex = -1;
        private ColumnLinkConfig? _existingLink;
        // 字段关联编辑暂存（用于选择文件/列后更新）
        private string _fieldEditTargetPath = "";
        private int _fieldEditTargetColIndex = -1;
        private string _fieldEditTargetColName = "";
        private int _fieldEditDisplayColIndex = -1;
        private string _fieldEditDisplayColName = "";
        private bool _isClosingConfirmed = false;

        /// <summary>
        /// 构造关联设置对话框。
        /// </summary>
        /// <param name="existingLink">传入的对象必须是**深拷贝副本**，对话框会直接修改此对象。确认时返回此对象，取消时返回 null。</param>
        public ColumnLinkSetupDialog(
            List<(string display, string fullPath)> fileEntries,
            string sourceFile, int sourceColumnIndex,
            string rootDir, int frozenRows,
            ColumnLinkConfig? existingLink,
            int targetHeaderRowIndex = -1)
        {
            _fileEntries = fileEntries;
            _sourceFile = sourceFile;
            _sourceColumnIndex = sourceColumnIndex;
            _rootDir = rootDir;
            _frozenRows = frozenRows;
            _targetHeaderRowIndex = targetHeaderRowIndex;
            // 传入的 existingLink 必须是深拷贝副本，直接操作它
            _existingLink = existingLink;

            InitializeComponent();

            // 加载已有配置
            if (_existingLink != null)
            {
                _chkExtractEnabled.Checked = _existingLink.GetEffectiveValueExtractionEnabled();
                _txtDisplayHeaders.Text = _existingLink.GetEffectiveDisplayColumnHeaders();
                RefreshStepList();
                if (_lstSteps.Items.Count > 0)
                    _lstSteps.SelectedIndex = 0;
                RefreshFieldLinkList();
            }

            // 在标题显示模板信息
            if (_existingLink != null && !string.IsNullOrEmpty(_existingLink.TemplateUid))
            {
                var template = LinkTemplateManager.GetTemplate(_existingLink.TemplateUid);
                if (template != null)
                    Text = $"Tips 显示设置 - 列{_sourceColumnIndex} [模板: {template.Name}]";
                else
                    Text = $"Tips 显示设置 - 列{_sourceColumnIndex} [模板已删除]";
            }

            // 拦截所有关闭操作（取消按钮、ESC、点 X）
            _btnCancel.Click += OnCancelClick;
            FormClosing += OnFormClosing;
        }

        public ColumnLinkConfig? GetResult()
        {
            // 确认时，将 UI 值同步到 _existingLink，然后返回
            if (_existingLink == null) return null;

            // 同步 UI 到 _existingLink
            _existingLink.ValueExtractionEnabled = _chkExtractEnabled.Checked;
            _existingLink.DisplayColumnHeaders = _txtDisplayHeaders.Text.Trim();
            // ExtractSteps 已在 AddStep/DeleteStep/UpdateStep 中直接修改 _existingLink.ExtractSteps
            // FieldLinks 已在字段操作中直接修改 _existingLink.FieldLinks
            // TemplateUid 保持传入时的值（如果是模板引用且用户编辑了，它仍然保留，运行时通过 GetEffective* 决定实际配置）

            bool hasConfig = _chkExtractEnabled.Checked ||
                             _txtDisplayHeaders.Text.Trim().Length > 0 ||
                             _existingLink.FieldLinks.Count > 0;
            return hasConfig ? _existingLink : null;
        }

        // ==================== 切割步骤管理 ====================
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
            var steps = _existingLink != null
                ? _existingLink.GetEffectiveExtractSteps()
                : new List<ExtractStep>();
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
            var steps = _existingLink?.GetEffectiveExtractSteps();
            if (steps == null || _lstSteps.SelectedIndex >= steps.Count) { _isLoadingStep = false; return; }
            var step = steps[_lstSteps.SelectedIndex];
            _txtStepDelimiter.Text = step.Delimiter;
            int cbIdx = step.PartIndex + 1;
            if (cbIdx >= 0 && cbIdx < _cboStepPartIndex.Items.Count)
                _cboStepPartIndex.SelectedIndex = cbIdx;
            _isLoadingStep = false;
        }

        private void AddStep()
        {
            EnsureExistingLink();
            _existingLink!.ExtractSteps.Add(new ExtractStep { Delimiter = "#", PartIndex = 0 });
            RefreshStepList();
            _lstSteps.SelectedIndex = _lstSteps.Items.Count - 1;
            UpdatePreview();
            RefreshFieldLinkList(); // 段数可能变化，刷新关联列表
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
            RefreshFieldLinkList(); // 段数可能变化，刷新关联列表
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
            step.PartIndex = cbIdx - 1;
            RefreshStepList();
            _lstSteps.SelectedIndex = Math.Min(_lstSteps.SelectedIndex, _lstSteps.Items.Count - 1);
            UpdatePreview();
            RefreshFieldLinkList(); // 段数可能变化，刷新关联列表
            _hasUnsavedChanges = false;
        }

        private List<ExtractStep> GetExtractStepsFromUI()
        {
            if (_existingLink != null)
            {
                return _existingLink.ExtractSteps;
            }
            return new List<ExtractStep>
            {
                new ExtractStep { Delimiter = "|", PartIndex = -1 },
                new ExtractStep { Delimiter = "#", PartIndex = 0 }
            };
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
            var stepDescs = new List<string>();
            foreach (var s in steps)
            {
                stepDescs.Add(s.PartIndex >= 0 ? $"\"{s.Delimiter}\"切取{s.PartIndex}段" : $"\"{s.Delimiter}\"切全部");
            }
            string flow = string.Join(" → ", stepDescs);
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

        // ==================== 多字段关联管理 ====================

        /// <summary>根据切割方案获取段数。最后一步 PartIndex=-1（保留全部）时由示例值切割决定段数，默认至少 5 段。</summary>
        private int GetFieldCount()
        {
            var steps = GetExtractStepsFromUI();
            if (steps.Count == 0) return 1;
            var lastStep = steps[steps.Count - 1];
            if (lastStep.PartIndex >= 0)
            {
                // 取特定段：根据已配置的 FieldLinks 动态决定段数
                int maxConfigured = _existingLink?.FieldLinks?.Count > 0 ? _existingLink.FieldLinks.Max(f => f.FieldIndex) + 1 : 0;
                return Math.Max(1, maxConfigured);
            }
            // 保留全部：取已配置的段数 + 默认额外段数
            int maxConfigured2 = _existingLink?.FieldLinks?.Count > 0 ? _existingLink.FieldLinks.Max(f => f.FieldIndex) + 1 : 0;
            return Math.Max(5, maxConfigured2);
        }

        /// <summary>获取当前段数，供外部（如列关联跳转时）使用。</summary>
        public int GetEffectiveFieldCount()
        {
            return GetFieldCount();
        }

        private void RefreshFieldLinkList()
        {
            _lstFieldLinks.BeginUpdate();
            _lstFieldLinks.Items.Clear();
            int fieldCount = GetFieldCount();
            for (int i = 0; i < fieldCount; i++)
            {
                var existing = _existingLink?.FieldLinks?.FirstOrDefault(f => f.FieldIndex == i);
                if (existing != null && !string.IsNullOrEmpty(existing.TargetFile))
                {
                    string colInfo = !string.IsNullOrEmpty(existing.TargetColumnName)
                        ? $"[{existing.TargetColumnName}]"
                        : $"[列{existing.TargetColumnIndex}]";
                    string displayInfo = existing.TargetDisplayColumnIndex >= 0
                        ? (!string.IsNullOrEmpty(existing.TargetDisplayColumnName)
                            ? $", 显示:{existing.TargetDisplayColumnName}"
                            : $", 显示:列{existing.TargetDisplayColumnIndex}")
                        : "";
                    string fn = Path.GetFileNameWithoutExtension(existing.TargetFile);
                    _lstFieldLinks.Items.Add($"第{i}段 - {fn} {colInfo}{displayInfo}");
                }
                else
                {
                    _lstFieldLinks.Items.Add($"第{i}段 - (未配置)");
                }
            }
            _lstFieldLinks.EndUpdate();
            // 刷新后恢复选中状态，由 SelectedIndexChanged 自动触发 UpdateFieldLinkInfo
            if (_editingFieldIndex >= 0 && _editingFieldIndex < _lstFieldLinks.Items.Count)
                _lstFieldLinks.SelectedIndex = _editingFieldIndex;
            else
                UpdateFieldLinkInfo();
        }

        /// <summary>检查当前选中段的暂存值相对于已保存值是否有未保存的修改</summary>
        private bool HasPendingFieldEditForCurrent()
        {
            if (_editingFieldIndex < 0) return false;
            if (string.IsNullOrEmpty(_fieldEditTargetPath)) return false;
            var existing = _existingLink?.FieldLinks?.FirstOrDefault(f => f.FieldIndex == _editingFieldIndex);
            if (existing == null) return true; // 新段，有暂存文件 = 有未保存
            return existing.TargetFile != _fieldEditTargetPath ||
                   existing.TargetColumnIndex != _fieldEditTargetColIndex ||
                   existing.TargetDisplayColumnIndex != _fieldEditDisplayColIndex;
        }

        private void UpdateFieldLinkInfo()
        {
            if (_lstFieldLinks.SelectedIndex < 0)
            {
                _lblFieldLinkInfo.Text = "请选择一段进行关联配置";
                _btnSelectFieldFile.Enabled = false;
                _btnSelectFieldCol.Enabled = false;
                _btnSelectFieldDisplayCol.Enabled = false;
                _editingFieldIndex = -1;
                return;
            }
            _editingFieldIndex = _lstFieldLinks.SelectedIndex;

            // ★ 每次切换段时先清空暂存值，然后从已保存值加载（如果已配置）
            _fieldEditTargetPath = "";
            _fieldEditTargetColIndex = -1;
            _fieldEditTargetColName = "";
            _fieldEditDisplayColIndex = -1;
            _fieldEditDisplayColName = "";
            var existing = _existingLink?.FieldLinks?.FirstOrDefault(f => f.FieldIndex == _editingFieldIndex);
            if (existing != null && !string.IsNullOrEmpty(existing.TargetFile))
            {
                _fieldEditTargetPath = existing.TargetFile;
                _fieldEditTargetColIndex = existing.TargetColumnIndex;
                _fieldEditTargetColName = existing.TargetColumnName;
                _fieldEditDisplayColIndex = existing.TargetDisplayColumnIndex;
                _fieldEditDisplayColName = existing.TargetDisplayColumnName;
            }

            // 构建信息文本：显示暂存值
            string filePart = "";
            string colPart = "";
            string displayPart = "";

            if (!string.IsNullOrEmpty(_fieldEditTargetPath))
                filePart = Path.GetFileName(_fieldEditTargetPath);
            else
                filePart = "(未选择文件)";

            if (_fieldEditTargetColIndex >= 0)
                colPart = !string.IsNullOrEmpty(_fieldEditTargetColName) ? $"[{_fieldEditTargetColName}]" : $"[列{_fieldEditTargetColIndex}]";

            if (_fieldEditDisplayColIndex >= 0)
                displayPart = !string.IsNullOrEmpty(_fieldEditDisplayColName) ? $" 显示:{_fieldEditDisplayColName}" : $" 显示:列{_fieldEditDisplayColIndex}";

            string info = $"第{_editingFieldIndex}段: {filePart}";
            if (!string.IsNullOrEmpty(colPart))
                info += $" {colPart}";
            if (!string.IsNullOrEmpty(displayPart))
                info += $",{displayPart}";

            // 标记未保存的修改
            if (HasPendingFieldEditForCurrent())
                info += " (未保存，请点击 [更新])";

            _lblFieldLinkInfo.Text = info;
            _btnSelectFieldFile.Enabled = true;
            _btnSelectFieldCol.Enabled = true;
            _btnSelectFieldDisplayCol.Enabled = true;
        }

        private void OnAddFieldLinkClick(object? sender, EventArgs e)
        {
            // 计算新段的索引：取当前最大索引 + 1
            EnsureExistingLink();
            var fieldLinks = _existingLink!.FieldLinks;
            int newIndex = fieldLinks.Count > 0 ? fieldLinks.Max(f => f.FieldIndex) + 1 : 0;

            // 直接往 FieldLinks 中添加一个新的空项（未配置的）
            fieldLinks.Add(new FieldLinkConfig
            {
                FieldIndex = newIndex,
                TargetFile = "",
                TargetColumnIndex = -1,
                TargetColumnName = "",
                TargetDisplayColumnIndex = -1,
                TargetDisplayColumnName = "",
            });

            RefreshFieldLinkList();
            UpdateFieldLinkInfo();
        }

        private void SaveCurrentFieldEdit()
        {
            if (_editingFieldIndex < 0) return;
            if (string.IsNullOrEmpty(_fieldEditTargetPath)) return;

            if (_existingLink == null) return;
            _existingLink.FieldLinks.RemoveAll(f => f.FieldIndex == _editingFieldIndex);
            _existingLink.FieldLinks.Add(new FieldLinkConfig
            {
                FieldIndex = _editingFieldIndex,
                TargetFile = _fieldEditTargetPath,
                TargetColumnIndex = _fieldEditTargetColIndex,
                TargetColumnName = _fieldEditTargetColName,
                TargetDisplayColumnIndex = _fieldEditDisplayColIndex,
                TargetDisplayColumnName = _fieldEditDisplayColName,
            });
            RefreshFieldLinkList();
            _lstFieldLinks.SelectedIndex = _editingFieldIndex;
        }

        private void OnDeleteFieldLinkClick(object? sender, EventArgs e)
        {
            if (_lstFieldLinks.SelectedIndex < 0) return;
            int fieldIdx = _lstFieldLinks.SelectedIndex;
            if (_existingLink == null) return;
            _existingLink.FieldLinks.RemoveAll(f => f.FieldIndex == fieldIdx);
            RefreshFieldLinkList();
            if (_lstFieldLinks.Items.Count > 0)
                _lstFieldLinks.SelectedIndex = Math.Min(fieldIdx, _lstFieldLinks.Items.Count - 1);
        }

        private void OnUpdateFieldLinkClick(object? sender, EventArgs e)
        {
            if (_editingFieldIndex < 0)
            {
                AlertForm.Show(this, "请先在列表中选择一段", "提示", AlertType.Information);
                return;
            }
            if (string.IsNullOrEmpty(_fieldEditTargetPath))
            {
                AlertForm.Show(this, "请先选择目标文件", "提示", AlertType.Information);
                return;
            }
            SaveCurrentFieldEdit();
        }

        private void OnSelectFieldFileClick(object? sender, EventArgs e)
        {
            if (_editingFieldIndex < 0) return;
            using var fileDlg = new LinkTableSelectorDialog(_fileEntries);
            if (fileDlg.ShowDialog(this) != DialogResult.OK) return;

            string selectedPath = fileDlg.SelectedFilePath;
            if (string.IsNullOrEmpty(selectedPath) || !File.Exists(selectedPath)) return;

            // 只选择文件，暂存路径（关联列和显示列后续单独选择）
            _fieldEditTargetPath = GetRelativePath(selectedPath);
            // 选择新文件后重置列选择
            _fieldEditTargetColIndex = -1;
            _fieldEditTargetColName = "";
            _fieldEditDisplayColIndex = -1;
            _fieldEditDisplayColName = "";

            // 只更新信息显示，不刷新 ListBox（避免破坏暂存值）
            _lblFieldLinkInfo.Text = $"第{_editingFieldIndex}段: {Path.GetFileName(_fieldEditTargetPath)} (未保存，请点击 [更新])";
        }

        private void OnSelectFieldColClick(object? sender, EventArgs e)
        {
            if (_editingFieldIndex < 0) return;
            // 使用暂存的路径或已保存的路径
            string targetPath = _fieldEditTargetPath;
            if (string.IsNullOrEmpty(targetPath))
            {
                var existing = _existingLink?.FieldLinks?.FirstOrDefault(f => f.FieldIndex == _editingFieldIndex);
                if (existing != null && !string.IsNullOrEmpty(existing.TargetFile))
                    targetPath = existing.TargetFile;
            }
            if (string.IsNullOrEmpty(targetPath))
            {
                AlertForm.Show(this, "请先选择目标文件", "提示", AlertType.Information);
                return;
            }
            string fullPath = ResolveTargetPath(targetPath);
            if (!File.Exists(fullPath))
            {
                AlertForm.Show(this, "目标文件不存在", "错误", AlertType.Error);
                return;
            }
            DataTable targetFrozenTable = ReadTargetFileForColumnSelect(fullPath);
            if (targetFrozenTable == null) return;

            int currentCol = _fieldEditTargetColIndex >= 0 ? _fieldEditTargetColIndex : -1;
            using var colDlg = new TargetColumnSelectDialog(targetFrozenTable, _frozenRows, currentCol,
                _targetHeaderRowIndex >= 0 ? _targetHeaderRowIndex : GetTargetHeaderRowIndexStatic(fullPath));
            if (colDlg.ShowDialog(this) == DialogResult.OK && colDlg.SelectedColumnIndex >= 0)
            {
                _fieldEditTargetColIndex = colDlg.SelectedColumnIndex;
                _fieldEditTargetColName = colDlg.SelectedColumnName;
                // 同时确保路径同步
                if (string.IsNullOrEmpty(_fieldEditTargetPath))
                    _fieldEditTargetPath = targetPath;
                // 直接更新信息显示，不调用 UpdateFieldLinkInfo（后者会清空暂存值）
                string colInfo = !string.IsNullOrEmpty(_fieldEditTargetColName)
                    ? $"[{_fieldEditTargetColName}]"
                    : $"[列{_fieldEditTargetColIndex}]";
                string displayInfo = _fieldEditDisplayColIndex >= 0
                    ? (!string.IsNullOrEmpty(_fieldEditDisplayColName)
                        ? $", 显示:{_fieldEditDisplayColName}"
                        : $", 显示:列{_fieldEditDisplayColIndex}")
                    : "";
                _lblFieldLinkInfo.Text = $"第{_editingFieldIndex}段: {Path.GetFileName(_fieldEditTargetPath)} {colInfo}{displayInfo} (未保存，请点击 [更新])";
            }
        }

        private void OnSelectFieldDisplayColClick(object? sender, EventArgs e)
        {
            if (_editingFieldIndex < 0) return;
            string targetPath = _fieldEditTargetPath;
            if (string.IsNullOrEmpty(targetPath))
            {
                var existing = _existingLink?.FieldLinks?.FirstOrDefault(f => f.FieldIndex == _editingFieldIndex);
                if (existing != null && !string.IsNullOrEmpty(existing.TargetFile))
                    targetPath = existing.TargetFile;
            }
            if (string.IsNullOrEmpty(targetPath))
            {
                AlertForm.Show(this, "请先选择目标文件", "提示", AlertType.Information);
                return;
            }
            string fullPath = ResolveTargetPath(targetPath);
            if (!File.Exists(fullPath))
            {
                AlertForm.Show(this, "目标文件不存在", "错误", AlertType.Error);
                return;
            }
            DataTable targetFrozenTable = ReadTargetFileForColumnSelect(fullPath);
            if (targetFrozenTable == null) return;

            int currentCol = _fieldEditDisplayColIndex >= 0 ? _fieldEditDisplayColIndex : -1;
            using var colDlg = new TargetColumnSelectDialog(targetFrozenTable, _frozenRows, currentCol,
                _targetHeaderRowIndex >= 0 ? _targetHeaderRowIndex : GetTargetHeaderRowIndexStatic(fullPath),
                title: "选择显示列 — 用于提示中显示中文名称");
            if (colDlg.ShowDialog(this) == DialogResult.OK && colDlg.SelectedColumnIndex >= 0)
            {
                _fieldEditDisplayColIndex = colDlg.SelectedColumnIndex;
                _fieldEditDisplayColName = colDlg.SelectedColumnName;
                if (string.IsNullOrEmpty(_fieldEditTargetPath))
                    _fieldEditTargetPath = targetPath;
                // 直接更新信息显示，不调用 UpdateFieldLinkInfo（后者会清空暂存值）
                string colInfo = !string.IsNullOrEmpty(_fieldEditTargetColName)
                    ? $"[{_fieldEditTargetColName}]"
                    : $"[列{_fieldEditTargetColIndex}]";
                string displayInfo = !string.IsNullOrEmpty(_fieldEditDisplayColName)
                    ? $" 显示:{_fieldEditDisplayColName}"
                    : $" 显示:列{_fieldEditDisplayColIndex}";
                _lblFieldLinkInfo.Text = $"第{_editingFieldIndex}段: {Path.GetFileName(_fieldEditTargetPath)} {colInfo},{displayInfo} (未保存，请点击 [更新])";
            }
        }

        // ==================== 辅助方法 ====================
        private string GetRelativePath(string fullPath)
        {
            if (!string.IsNullOrEmpty(_rootDir) && fullPath.StartsWith(_rootDir, StringComparison.OrdinalIgnoreCase))
                return fullPath.Substring(_rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(fullPath);
        }

        private string ResolveTargetPath(string targetRel)
        {
            string candidate = Path.Combine(_rootDir, targetRel);
            if (File.Exists(candidate)) return candidate;
            var match = _fileEntries.FirstOrDefault(f =>
                Path.GetFileName(f.fullPath).Equals(Path.GetFileName(targetRel), StringComparison.OrdinalIgnoreCase));
            return match.fullPath ?? candidate;
        }

        private DataTable? ReadTargetFileForColumnSelect(string fullPath)
        {
            try
            {
                var ext = Path.GetExtension(fullPath).ToLowerInvariant();
                DataTable readOnlyData;
                if (ext == ".csv")
                {
                    // 只做列选择预览，取前几列即可：直接解析、不进阅读器窗口的缓存，
                    // 否则配置向导扫过的每个目标表都会常驻一份完整 DataTable。
                    readOnlyData = FileDataCache.LoadFileDirect(fullPath).Item1;
                }
                else
                {
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                    using var fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var excelReader = ExcelReaderFactory.CreateReader(fs);
                    var dataSet = excelReader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
                    });
                    if (dataSet.Tables.Count > 0)
                        readOnlyData = dataSet.Tables[0];
                    else
                        return new DataTable();
                }

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
                return dt;
            }
            catch
            {
                return null;
            }
        }

        private void OnOKClick(object? sender, EventArgs e)
        {
            // 检查是否有未保存的字段编辑暂存
            bool hasPendingFieldEdit = HasPendingFieldEditForCurrent() && _editingFieldIndex >= 0;
            bool hasPendingStepEdit = _hasUnsavedChanges;

            if (hasPendingFieldEdit || hasPendingStepEdit)
            {
                var sb = new StringBuilder();
                if (hasPendingFieldEdit)
                    sb.AppendLine($"第{_editingFieldIndex}段关联配置有未保存的修改。");
                if (hasPendingStepEdit)
                    sb.AppendLine("切割步骤有未保存的修改。");
                sb.AppendLine();
                sb.Append("「保存并关闭」— 保存修改并关闭对话框" + GlobalConst.NewLine);
                sb.Append("「放弃并关闭」— 放弃修改并关闭对话框" + GlobalConst.NewLine);
                sb.Append("「继续编辑」— 返回对话框继续编辑");

                var result = AlertForm.ShowConfirmThreeButton(this, sb.ToString(), "未保存的修改",
                    yesText: "保存并关闭", noText: "放弃并关闭", cancelText: "继续编辑");
                if (result == DialogResult.Yes)
                {
                    if (hasPendingFieldEdit) SaveCurrentFieldEdit();
                    if (hasPendingStepEdit) UpdateStep();
                    _isClosingConfirmed = true;
                    DialogResult = DialogResult.Yes;
                    Close();
                }
                else if (result == DialogResult.No)
                {
                    _isClosingConfirmed = true;
                    DialogResult = DialogResult.Yes;
                    Close();
                }
                // DialogResult.Cancel: 返回继续编辑
                return;
            }
            _isClosingConfirmed = true;
            DialogResult = DialogResult.Yes;
            Close();
        }

        private void OnFieldLinkSelectionChanged(object? sender, EventArgs e)
        {
            int newSelectedIndex = _lstFieldLinks.SelectedIndex;
            if (newSelectedIndex < 0) { UpdateFieldLinkInfo(); return; }

            // 如果当前段有未保存的修改，且确实在切换段
            if (HasPendingFieldEditForCurrent() && _editingFieldIndex >= 0 && newSelectedIndex != _editingFieldIndex)
            {
                string msg = $"第{_editingFieldIndex}段有未保存的修改。{GlobalConst.NewLine}{GlobalConst.NewLine}" +
                             $"「保存并切换」— 保存当前段的编辑并切换到第{newSelectedIndex}段{GlobalConst.NewLine}" +
                             $"「放弃并切换」— 放弃当前段的修改并切换到第{newSelectedIndex}段";
                var result = AlertForm.ShowConfirmThreeButton(this, msg, "未保存的修改",
                    yesText: "保存并切换", noText: "放弃并切换", cancelText: "继续编辑");
                if (result == DialogResult.Yes)
                {
                    SaveCurrentFieldEdit();
                }
                else if (result == DialogResult.Cancel)
                {
                    // 取消切换：恢复选中到原段
                    _lstFieldLinks.SelectedIndexChanged -= OnFieldLinkSelectionChanged;
                    _lstFieldLinks.SelectedIndex = _editingFieldIndex;
                    _lstFieldLinks.SelectedIndexChanged += OnFieldLinkSelectionChanged;
                    return;
                }
                // DialogResult.No: 放弃修改，继续切换
            }
            UpdateFieldLinkInfo();
        }

        private void OnCancelClick(object? sender, EventArgs e)
        {
            if (TryCloseWithPendingCheck())
            {
                _isClosingConfirmed = true;
                Close();
            }
        }

        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            if (_isClosingConfirmed) return;
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                if (TryCloseWithPendingCheck())
                {
                    _isClosingConfirmed = true;
                    Close();
                }
            }
        }

        private bool TryCloseWithPendingCheck()
        {
            bool hasPendingFieldEdit = HasPendingFieldEditForCurrent() && _editingFieldIndex >= 0;
            bool hasPendingStepEdit = _hasUnsavedChanges;

            if (hasPendingFieldEdit || hasPendingStepEdit)
            {
                var sb = new StringBuilder();
                if (hasPendingFieldEdit)
                    sb.AppendLine($"第{_editingFieldIndex}段关联配置有未保存的修改。");
                if (hasPendingStepEdit)
                    sb.AppendLine("切割步骤有未保存的修改。");
                sb.AppendLine();
                sb.Append("「保存并关闭」— 保存修改并关闭对话框" + GlobalConst.NewLine);
                sb.Append("「放弃并关闭」— 放弃修改并关闭对话框" + GlobalConst.NewLine);
                sb.Append("「继续编辑」— 返回对话框继续编辑");

                var result = AlertForm.ShowConfirmThreeButton(this, sb.ToString(), "未保存的修改",
                    yesText: "保存并关闭", noText: "放弃并关闭", cancelText: "继续编辑");
                if (result == DialogResult.Yes)
                {
                    if (hasPendingFieldEdit) SaveCurrentFieldEdit();
                    if (hasPendingStepEdit) UpdateStep();
                    return true;
                }
                else if (result == DialogResult.No)
                {
                    return true;
                }
                else // DialogResult.Cancel — 继续编辑
                {
                    return false;
                }
            }
            return true; // 无未保存修改，允许关闭
        }

        private void OnDeleteClick(object? sender, EventArgs e)
        {
            _existingLink = null;
            RefreshFieldLinkList();
            RefreshStepList();
            _txtDisplayHeaders.Text = "";
            _chkExtractEnabled.Checked = false;
        }

        // ==================== UI 初始化 ====================
        private void InitializeComponent()
        {
            SuspendLayout();
            Text = "Tips 显示设置 - 列 " + _sourceColumnIndex;
            ClientSize = new Size(540, 540);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false; MaximizeBox = false;

            int y = 10;
            int labelH = 18;
            int gap = 4;

            // ===== 标题 =====
            Controls.Add(new Label
            {
                Text = $"设置列 {_sourceColumnIndex} 的 Tips 显示",
                Location = new Point(10, y),
                Size = new Size(520, 22),
                Font = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55)
            });
            y += 28;

            // ===== 第一步：切割方案 =====
            var grpExtract = new GroupBox
            {
                Text = "第一步：切割方案 — 从单元格值中提取多个显示值",
                Location = new Point(10, y),
                Size = new Size(520, 175),
                Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            y += 180;

            _chkExtractEnabled = new CheckBox
            {
                Text = "启用切割（如 \"4002#1|3001#2\" → 按规则提取）",
                Location = new Point(10, 20),
                Size = new Size(320, 22),
                Font = new Font("Microsoft YaHei UI", 8f),
                Checked = false,
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

            _lstSteps = new ListBox
            {
                Location = new Point(10, 42),
                Size = new Size(350, 70),
                Font = new Font("Microsoft YaHei UI", 8f),
                Enabled = false,
                BorderStyle = BorderStyle.FixedSingle,
                IntegralHeight = false,
                BackColor = Color.FromArgb(250, 252, 255)
            };
            _lstSteps.SelectedIndexChanged += OnStepSelected;
            grpExtract.Controls.Add(_lstSteps);

            _btnAddStep = new Button
            {
                Text = "＋添加",
                Location = new Point(368, 42),
                Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                Enabled = false,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };
            _btnAddStep.Click += (s, e) => { AddStep(); };
            grpExtract.Controls.Add(_btnAddStep);

            _btnDeleteStep = new Button
            {
                Text = "✕删除",
                Location = new Point(422, 42),
                Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                Enabled = false,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(255, 240, 240),
                ForeColor = Color.FromArgb(180, 40, 40),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(240, 200, 200) }
            };
            _btnDeleteStep.Click += (s, e) => { DeleteStep(); };
            grpExtract.Controls.Add(_btnDeleteStep);

            _lblStepDelimiter = new Label
            {
                Text = "分隔符:",
                Location = new Point(10, 118),
                Size = new Size(50, 18),
                Font = new Font("Microsoft YaHei UI", 8f),
                Enabled = false,
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            grpExtract.Controls.Add(_lblStepDelimiter);

            _txtStepDelimiter = new TextBox
            {
                Text = "",
                Location = new Point(60, 116),
                Size = new Size(40, 21),
                Font = new Font("Microsoft YaHei UI", 8f),
                BorderStyle = BorderStyle.FixedSingle,
                Enabled = false,
                PlaceholderText = "# | :"
            };
            _txtStepDelimiter.TextChanged += (s, e) => MarkUnsaved();
            grpExtract.Controls.Add(_txtStepDelimiter);

            _lblStepPartIndex = new Label
            {
                Text = "取段:",
                Location = new Point(108, 118),
                Size = new Size(35, 18),
                Font = new Font("Microsoft YaHei UI", 8f),
                Enabled = false,
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            grpExtract.Controls.Add(_lblStepPartIndex);

            _cboStepPartIndex = new ComboBox
            {
                Location = new Point(143, 116),
                Size = new Size(100, 21),
                Font = new Font("Microsoft YaHei UI", 8f),
                Enabled = false,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat
            };
            _cboStepPartIndex.Items.AddRange(new object[]
            {
                "-1 = 保留全部", "0 = 取第 0 段", "1 = 取第 1 段", "2 = 取第 2 段",
                "3 = 取第 3 段", "4 = 取第 4 段", "5 = 取第 5 段", "6 = 取第 6 段",
                "7 = 取第 7 段", "8 = 取第 8 段", "9 = 取第 9 段"
            });
            _cboStepPartIndex.SelectedIndex = 1;
            _cboStepPartIndex.SelectedIndexChanged += (s, e) => MarkUnsaved();
            grpExtract.Controls.Add(_cboStepPartIndex);

            _btnUpdateStep = new Button
            {
                Text = "更新",
                Location = new Point(250, 114),
                Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                Enabled = false,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
            };
            _btnUpdateStep.Click += (s, e) => { UpdateStep(); };
            grpExtract.Controls.Add(_btnUpdateStep);

            _lblExtractPreview = new Label
            {
                Text = "预览: 输入示例 \"2#1|3#1\"  →  结果: [\"2\",\"3\"]",
                Location = new Point(10, 142),
                Size = new Size(500, 18),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = Color.FromArgb(140, 145, 150),
                Visible = false
            };
            grpExtract.Controls.Add(_lblExtractPreview);

            Controls.Add(grpExtract);

            // ===== 第二步：表头输入 =====
            Controls.Add(new Label
            {
                Text = "第二步：表头（可选，用分隔符分隔，如 \"道具#数量#权重\"，留空不显示表头）:",
                Location = new Point(10, y),
                Size = new Size(520, labelH),
                Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(55, 65, 81)
            });
            y += labelH + gap;

            _txtDisplayHeaders = new TextBox
            {
                Text = "",
                Location = new Point(10, y),
                Size = new Size(520, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                PlaceholderText = "道具#数量#权重"
            };
            Controls.Add(_txtDisplayHeaders);
            y += 30;

            // ===== 第三步：多字段关联配置 =====
            var grpFieldLinks = new GroupBox
            {
                Text = "第三步：各段关联配置 — 每段可以关联不同的表",
                Location = new Point(10, y),
                Size = new Size(520, 215),
                Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            y += 220;

            _lstFieldLinks = new ListBox
            {
                Location = new Point(10, 20),
                Size = new Size(340, 80),
                Font = new Font("Microsoft YaHei UI", 8f),
                BorderStyle = BorderStyle.FixedSingle,
                IntegralHeight = false,
                BackColor = Color.FromArgb(250, 252, 255)
            };
            _lstFieldLinks.SelectedIndexChanged += OnFieldLinkSelectionChanged;
            grpFieldLinks.Controls.Add(_lstFieldLinks);

            // 添加按钮
            _btnAddFieldLink = new Button
            {
                Text = "添加",
                Location = new Point(358, 20),
                Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };
            _btnAddFieldLink.Click += OnAddFieldLinkClick;
            grpFieldLinks.Controls.Add(_btnAddFieldLink);

            // 删除按钮
            _btnDeleteFieldLink = new Button
            {
                Text = "删除",
                Location = new Point(358, 48),
                Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(255, 240, 240),
                ForeColor = Color.FromArgb(180, 40, 40),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(240, 200, 200) }
            };
            _btnDeleteFieldLink.Click += OnDeleteFieldLinkClick;
            grpFieldLinks.Controls.Add(_btnDeleteFieldLink);

            // 更新按钮
            _btnUpdateFieldLink = new Button
            {
                Text = "更新",
                Location = new Point(358, 76),
                Size = new Size(50, 24),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
            };
            _btnUpdateFieldLink.Click += OnUpdateFieldLinkClick;
            grpFieldLinks.Controls.Add(_btnUpdateFieldLink);

            // 选中段的详情
            _lblFieldLinkInfo = new Label
            {
                Text = "请选择一段进行关联配置",
                Location = new Point(10, 106),
                Size = new Size(500, 18),
                Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(31, 41, 55)
            };
            grpFieldLinks.Controls.Add(_lblFieldLinkInfo);

            // 编辑提示
            grpFieldLinks.Controls.Add(new Label
            {
                Text = "编辑当前选中段：",
                Location = new Point(10, 128),
                Size = new Size(500, 16),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = Color.FromArgb(100, 120, 140)
            });

            // 编辑按钮（全部放在 GroupBox 内部）
            _btnSelectFieldFile = new Button
            {
                Text = "选择目标文件",
                Location = new Point(10, 148),
                Size = new Size(130, 26),
                Font = new Font("Microsoft YaHei UI", 8f),
                Enabled = false,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };
            _btnSelectFieldFile.Click += OnSelectFieldFileClick;
            grpFieldLinks.Controls.Add(_btnSelectFieldFile);

            _btnSelectFieldCol = new Button
            {
                Text = "选择关联列",
                Location = new Point(150, 148),
                Size = new Size(100, 26),
                Font = new Font("Microsoft YaHei UI", 8f),
                Enabled = false,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };
            _btnSelectFieldCol.Click += OnSelectFieldColClick;
            grpFieldLinks.Controls.Add(_btnSelectFieldCol);

            _btnSelectFieldDisplayCol = new Button
            {
                Text = "显示列（可选）",
                Location = new Point(260, 148),
                Size = new Size(120, 26),
                Font = new Font("Microsoft YaHei UI", 8f),
                Enabled = false,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };
            _btnSelectFieldDisplayCol.Click += OnSelectFieldDisplayColClick;
            grpFieldLinks.Controls.Add(_btnSelectFieldDisplayCol);

            // 操作说明
            grpFieldLinks.Controls.Add(new Label
            {
                Text = "操作：选中一段 → 编辑（文件/关联列/显示列）→ 点击 [更新] 保存",
                Location = new Point(10, 180),
                Size = new Size(500, 16),
                Font = new Font("Microsoft YaHei UI", 7.5f),
                ForeColor = Color.FromArgb(140, 145, 150)
            });

            Controls.Add(grpFieldLinks);

            // ===== 按钮区域 =====
            _btnDelete = new Button
            {
                Text = "清除设置",
                Location = new Point(10, y),
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(239, 68, 68),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(220, 50, 50) },
                Cursor = Cursors.Hand,
                Visible = _existingLink != null
            };
            _btnDelete.Click += OnDeleteClick;
            Controls.Add(_btnDelete);

            _btnOK = new Button
            {
                Text = "确定",
                Location = new Point(370, y),
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                Cursor = Cursors.Hand
            };
            _btnOK.Click += OnOKClick;
            Controls.Add(_btnOK);

            _btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(460, y),
                Size = new Size(70, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };
            Controls.Add(_btnCancel);

            AcceptButton = _btnOK;
            ResumeLayout(false);
        }

        private static int GetTargetHeaderRowIndexStatic(string targetFullPath)
        {
            if (string.IsNullOrEmpty(targetFullPath)) return -1;
            string targetDir = Path.GetDirectoryName(targetFullPath) ?? "";
            var ext = Path.GetExtension(targetFullPath).ToLowerInvariant();
            if (ext == ".csv")
                return CsvSettings.GetCsvHeaderRowIndexByDirectory(targetDir);
            else if (ext == ".xls" || ext == ".xlsx")
                return CsvSettings.GetExcelHeaderRowIndexByDirectory(targetDir);
            return -1;
        }
    }
}
