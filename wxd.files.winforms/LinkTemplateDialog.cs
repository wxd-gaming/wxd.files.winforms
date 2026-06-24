using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>关联配置模板管理对话框 — 新增、编辑、删除模板</summary>
    public class LinkTemplateDialog : Form
    {
        private ListBox _lstTemplates = new();
        private TextBox _txtName = new();
        private Button _btnNew = new();
        private Button _btnSave = new();
        private Button _btnDelete = new();
        private Button _btnEditConfig = new();
        private Button _btnClose = new();
        private Button? _btnSelect = null;
        private Label _lblPreview = new();

        private string _directoryPath = "";
        private List<LinkTemplate> _templates = new();
        private LinkTemplate? _editingTemplate = null;
        private List<(string display, string fullPath)> _fileEntries = new();
        private int _frozenRows = 0;
        private int _headerRowIndex = -1;

        /// <summary>标记是否需要重新加载（供外部刷新用）</summary>
        public bool TemplatesChanged { get; private set; } = false;

        /// <summary>选中的模板 Uid（点"引用此模板"按钮时设置）</summary>
        public string? SelectedTemplateUid { get; private set; } = null;

        /// <summary>是否选择了模板（点"引用此模板"）</summary>
        public bool IsTemplateSelected { get; private set; } = false;

        /// <param name="mode">true=选择模式（显示"引用此模板"按钮），false=管理模式（仅管理）</param>
        /// <param name="fileEntries">可选的文件条目列表，供编辑配置时选择关联文件使用</param>
        /// <param name="frozenRows">当前文件的冻结行数，用于编辑配置时传递给关联设置对话框</param>
        /// <param name="headerRowIndex">当前文件的表头行索引，用于编辑配置时传递给关联设置对话框</param>
        /// <param name="directoryPath">关联配置的目录路径，用于解除引用和扫描文件降级</param>
        public LinkTemplateDialog(bool mode = false,
            List<(string display, string fullPath)>? fileEntries = null,
            int frozenRows = 0, int headerRowIndex = -1,
            string directoryPath = "")
        {
            _directoryPath = directoryPath;
            _fileEntries = fileEntries ?? new();
            _frozenRows = frozenRows;
            _headerRowIndex = headerRowIndex;
            ReloadTemplates();
            InitializeComponent(mode);
            RefreshTemplateList();
        }

        /// <summary>从全局存储重新加载模板列表</summary>
        private void ReloadTemplates()
        {
            _templates = LinkTemplateManager.GetTemplates();
        }

        private void InitializeComponent(bool selectMode)
        {
            SuspendLayout();
            Text = selectMode ? "选择关联配置模板" : "管理关联配置模板";
            ClientSize = new Size(560, 400);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false;
            MaximizeBox = false;

            // 左侧：模板列表
            var lblList = new Label
            {
                Text = "模板列表：",
                Location = new Point(10, 10),
                Size = new Size(200, 18),
                Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            Controls.Add(lblList);

            _lstTemplates = new ListBox
            {
                Location = new Point(10, 30),
                Size = new Size(220, 280),
                Font = new Font("Microsoft YaHei UI", 8f),
                BorderStyle = BorderStyle.FixedSingle,
                IntegralHeight = false,
                BackColor = Color.FromArgb(250, 252, 255)
            };
            _lstTemplates.SelectedIndexChanged += OnTemplateSelected;
            Controls.Add(_lstTemplates);

            // 右侧：编辑区
            var lblName = new Label
            {
                Text = "模板名称：",
                Location = new Point(250, 30),
                Size = new Size(100, 18),
                Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            Controls.Add(lblName);

            _txtName = new TextBox
            {
                Location = new Point(250, 50),
                Size = new Size(290, 24),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                PlaceholderText = "输入模板名称..."
            };
            Controls.Add(_txtName);

            // 预览区
            var lblPrev = new Label
            {
                Text = "配置预览：",
                Location = new Point(250, 82),
                Size = new Size(100, 18),
                Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(55, 65, 81)
            };
            Controls.Add(lblPrev);

            _lblPreview = new Label
            {
                Text = "（选择模板后显示预览）",
                Location = new Point(250, 102),
                Size = new Size(290, 120),
                Font = new Font("Microsoft YaHei UI", 8f),
                ForeColor = Color.FromArgb(100, 120, 140),
                BackColor = Color.FromArgb(250, 252, 255),
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(4)
            };
            Controls.Add(_lblPreview);

            // 管理按钮
            _btnNew = new Button
            {
                Text = "新增",
                Location = new Point(250, 235),
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                Cursor = Cursors.Hand
            };
            _btnNew.Click += OnNewClick;
            Controls.Add(_btnNew);

            _btnSave = new Button
            {
                Text = "保存",
                Location = new Point(340, 235),
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(22, 163, 74),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(21, 128, 61) },
                Cursor = Cursors.Hand,
                Enabled = false
            };
            _btnSave.Click += OnSaveClick;
            Controls.Add(_btnSave);

            _btnDelete = new Button
            {
                Text = "删除",
                Location = new Point(430, 235),
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(239, 68, 68),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(220, 50, 50) },
                Cursor = Cursors.Hand,
                Enabled = false
            };
            _btnDelete.Click += OnDeleteClick;
            Controls.Add(_btnDelete);

            // 编辑配置按钮
            _btnEditConfig = new Button
            {
                Text = "编辑配置",
                Location = new Point(250, 270),
                Size = new Size(130, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                Cursor = Cursors.Hand,
                Enabled = false
            };
            _btnEditConfig.Click += OnEditConfigClick;
            Controls.Add(_btnEditConfig);

            // 底部按钮
            int y = 340;

            if (selectMode)
            {
                _btnSelect = new Button
                {
                    Text = "引用此模板",
                    Location = new Point(10, y),
                    Size = new Size(110, 28),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(22, 163, 74),
                    ForeColor = Color.White,
                    FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(21, 128, 61) },
                    Cursor = Cursors.Hand,
                    Enabled = false
                };
                _btnSelect.Click += (s, e) =>
                {
                    if (_editingTemplate != null)
                    {
                        SelectedTemplateUid = _editingTemplate.Uid;
                        IsTemplateSelected = true;
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                };
                Controls.Add(_btnSelect);

                // 选择模式也允许直接"取消"
                _btnClose = new Button
                {
                    Text = "取消",
                    Location = new Point(470, y),
                    Size = new Size(70, 28),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(235, 248, 255),
                    ForeColor = Color.FromArgb(20, 60, 100),
                    FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
                };
                _btnClose.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
                Controls.Add(_btnClose);
            }
            else
            {
                _btnClose = new Button
                {
                    Text = "关闭",
                    Location = new Point(470, y),
                    Size = new Size(70, 28),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(235, 248, 255),
                    ForeColor = Color.FromArgb(20, 60, 100),
                    FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
                };
                _btnClose.Click += (s, e) => Close();
                Controls.Add(_btnClose);
            }

            ResumeLayout(false);
        }

        private void RefreshTemplateList()
        {
            _lstTemplates.BeginUpdate();
            _lstTemplates.Items.Clear();
            foreach (var t in _templates)
            {
                string summary = LinkTemplateManager.GetTemplateSummary(t);
                _lstTemplates.Items.Add($"{t.Name}  —  {summary}");
            }
            _lstTemplates.EndUpdate();
        }

        private void OnTemplateSelected(object? sender, EventArgs e)
        {
            if (_lstTemplates.SelectedIndex < 0) return;
            _editingTemplate = _templates[_lstTemplates.SelectedIndex];
            _txtName.Text = _editingTemplate.Name;
            _btnSave.Enabled = true;
            _btnDelete.Enabled = true;
            _btnEditConfig.Enabled = true;
            if (_btnSelect != null) _btnSelect.Enabled = true;
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (_editingTemplate == null)
            {
                _lblPreview.Text = "（选择模板后显示预览）";
                return;
            }
            var lines = new List<string>();
            if (_editingTemplate.ValueExtractionEnabled && _editingTemplate.ExtractSteps.Count > 0)
            {
                var descs = _editingTemplate.ExtractSteps.Select((s, i) =>
                    s.PartIndex >= 0
                        ? $"{i + 1}. \"{s.Delimiter}\" → 取第{s.PartIndex}段"
                        : $"{i + 1}. \"{s.Delimiter}\" → 保留全部");
                lines.Add("【切割方案】");
                lines.AddRange(descs);
            }
            if (!string.IsNullOrEmpty(_editingTemplate.DisplayColumnHeaders))
            {
                lines.Add($"【表头】{_editingTemplate.DisplayColumnHeaders}");
            }
            int linkedFields = _editingTemplate.FieldLinks?.Count(f => !string.IsNullOrEmpty(f.TargetFile)) ?? 0;
            if (linkedFields > 0)
            {
                lines.Add($"【关联】{linkedFields}个字段已配置");
                foreach (var f in _editingTemplate.FieldLinks.Where(f => !string.IsNullOrEmpty(f.TargetFile)))
                {
                    string colInfo = !string.IsNullOrEmpty(f.TargetColumnName)
                        ? $"[{f.TargetColumnName}]" : $"[列{f.TargetColumnIndex}]";
                    lines.Add($"  第{f.FieldIndex}段 → {Path.GetFileNameWithoutExtension(f.TargetFile)} {colInfo}");
                }
            }
            _lblPreview.Text = lines.Count > 0 ? string.Join(GlobalConst.NewLine, lines) : "（空模板）";
        }

        private void OnNewClick(object? sender, EventArgs e)
        {
            var template = new LinkTemplate
            {
                Name = "新建模板",
                ValueExtractionEnabled = false,
                ExtractSteps = new List<ExtractStep>
                {
                    new ExtractStep { Delimiter = "|", PartIndex = -1 },
                    new ExtractStep { Delimiter = "#", PartIndex = 0 }
                },
                FieldLinks = new()
            };
            // 立即持久化到文件，确保后续 UpdateTemplate 能找到
            LinkTemplateManager.AddTemplate(template);
            TemplatesChanged = true;
            ReloadTemplates();
            RefreshTemplateList();
            // 选中新建的模板并强制刷新预览
            _lstTemplates.SelectedIndex = _templates.Count - 1;
            OnTemplateSelected(null, EventArgs.Empty);
        }

        private void OnSaveClick(object? sender, EventArgs e)
        {
            if (_editingTemplate == null) return;
            string name = _txtName.Text.Trim();
            if (string.IsNullOrEmpty(name))
            {
                AlertForm.Show(this, "请输入模板名称", "提示", AlertType.Information);
                return;
            }
            _editingTemplate.Name = name;
            LinkTemplateManager.UpdateTemplate(_editingTemplate);
            TemplatesChanged = true;
            ReloadTemplates();
            RefreshTemplateList();
            // 重新定位到编辑的模板并强制刷新预览
            int idx = _templates.FindIndex(t => t.Uid == _editingTemplate.Uid);
            if (idx >= 0)
            {
                _lstTemplates.SelectedIndex = idx;
                OnTemplateSelected(null, EventArgs.Empty);
            }
            AlertForm.Show(this, "模板已保存", "提示", AlertType.Information);
        }

        private void OnDeleteClick(object? sender, EventArgs e)
        {
            if (_editingTemplate == null) return;
            string msg = $"确定要删除模板「{_editingTemplate.Name}」吗？";
            // 检查是否有引用
            var refLinks = LinkTemplateManager.GetReferencingLinks(_directoryPath, _editingTemplate.Uid);
            // 也查询全局是否有其他引用
            if (refLinks.Count > 0)
            {
                // 列出引用的详细信息
                var refDetails = refLinks
                    .GroupBy(l => l.SourceFile)
                    .Select(g => $"  文件: {g.Key} → 列 {string.Join(", ", g.Select(l => l.SourceColumnIndex))}")
                    .ToList();
                msg += $"{GlobalConst.NewLine}{GlobalConst.NewLine}⚠ 以下 {refLinks.Count} 个列关联配置引用了此模板，删除后将无法从模板获取配置：{GlobalConst.NewLine}{string.Join(GlobalConst.NewLine, refDetails)}";
            }

            var result = AlertForm.ShowConfirm(this, msg, "确认删除");
            if (result != DialogResult.Yes) return;

            // 解除所有引用
            LinkTemplateManager.UnlinkAllReferences(_directoryPath, _editingTemplate.Uid);
            LinkTemplateManager.DeleteTemplate(_editingTemplate.Uid);
            TemplatesChanged = true;
            ReloadTemplates();
            _editingTemplate = null;
            _txtName.Text = "";
            _lblPreview.Text = "（选择模板后显示预览）";
            _btnSave.Enabled = false;
            _btnDelete.Enabled = false;
            _btnEditConfig.Enabled = false;
            RefreshTemplateList();
        }

        private void OnEditConfigClick(object? sender, EventArgs e)
        {
            if (_editingTemplate == null) return;

            // 使用外部传入的文件列表，而不是自己扫描目录
            // 这样能覆盖跨窗格/跨目录的文件
            var fileEntries = _fileEntries;

            // 如果外部没有传入文件列表，则降级为扫描目录
            if (fileEntries.Count == 0 && Directory.Exists(_directoryPath))
            {
                bool showHidden = CsvSettings.ShowHiddenFiles;
                foreach (var ext in new[] { "*.csv", "*.xls", "*.xlsx" })
                {
                    foreach (var file in Directory.GetFiles(_directoryPath, ext))
                    {
                        if (!showHidden && (File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden)
                            continue;
                        string display = Path.GetFileName(file);
                        fileEntries.Add((display, file));
                    }
                }
            }

            // 构建临时 ColumnLinkConfig（深拷贝模板配置）
            var tempLink = new ColumnLinkConfig();
            tempLink.TemplateUid = _editingTemplate.Uid;
            tempLink.RefreshFromTemplate();

            // 传入深拷贝副本，对话框直接修改它
            using var dialog = new ColumnLinkSetupDialog(
                fileEntries, "", 0, _directoryPath, _frozenRows,
                tempLink, _headerRowIndex);
            dialog.Text = $"编辑模板配置 - {_editingTemplate.Name}";

            if (dialog.ShowDialog(this) == DialogResult.Yes)
            {
                var result = dialog.GetResult();
                if (result != null)
                {
                    // 将结果写回模板
                    _editingTemplate.ValueExtractionEnabled = result.ValueExtractionEnabled;
                    _editingTemplate.ExtractSteps = result.ExtractSteps;
                    _editingTemplate.DisplayColumnHeaders = result.DisplayColumnHeaders;
                    _editingTemplate.FieldLinks = result.FieldLinks;

                    LinkTemplateManager.UpdateTemplate(_editingTemplate);
                    TemplatesChanged = true;
                    ReloadTemplates();
                    // 重新定位到编辑的模板并强制刷新预览和列表
                    int idx = _templates.FindIndex(t => t.Uid == _editingTemplate.Uid);
                    if (idx >= 0)
                    {
                        RefreshTemplateList();
                        _lstTemplates.SelectedIndex = idx;
                        OnTemplateSelected(null, EventArgs.Empty);
                    }
                    else
                    {
                        RefreshTemplateList();
                        UpdatePreview();
                    }
                    AlertForm.Show(this, "模板配置已更新", "提示", AlertType.Information);
                }
            }
        }
    }
}
