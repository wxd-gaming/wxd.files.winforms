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
        private ThemeField _fldName = new();
        private TextBox _txtName;
        private ThemeButton _btnNew = new();
        private ThemeButton _btnSave = new();
        private ThemeButton _btnDelete = new();
        private ThemeButton _btnEditConfig = new();
        private ThemeButton _btnClose = new();
        private ThemeButton? _btnSelect = null;
        private Label _lblPreview = new();
        private Label _lblList = new();
        private Label _lblName = new();
        private Label _lblPrev = new();
        private Panel _bar = new();

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
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
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
            Icon = WindowIcons.Link;
            ClientSize = new Size(560, 400);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;

            // 左侧：模板列表
            _lblList = new Label { Text = "模板列表：", AutoSize = true };

            _lstTemplates = new ListBox
            {
                Font = Theme.Caption,
                BorderStyle = BorderStyle.None,
                IntegralHeight = false
            };
            _lstTemplates.SelectedIndexChanged += OnTemplateSelected;

            // 右侧：编辑区
            _lblName = new Label { Text = "模板名称：", AutoSize = true };
            _fldName = new ThemeField { Placeholder = "输入模板名称..." };
            _txtName = _fldName.Edit;

            _lblPrev = new Label { Text = "配置预览：", AutoSize = true };
            _lblPreview = new Label
            {
                Text = "（选择模板后显示预览）",
                Font = Theme.Caption,
                BorderStyle = BorderStyle.None,
                Padding = new Padding(4),
                AutoSize = false
            };

            _btnNew = DialogButtonBar.Make("新增", OnNewClick);
            _btnSave = DialogButtonBar.Make("保存", OnSaveClick);
            _btnSave.Enabled = false;
            _btnDelete = DialogButtonBar.Make("删除", OnDeleteClick);
            _btnDelete.Enabled = false;
            _btnEditConfig = DialogButtonBar.Make("编辑配置", OnEditConfigClick);
            _btnEditConfig.Enabled = false;

            // 底部：选择模式把"引用此模板"当主操作，管理模式只留关闭
            _bar = DialogButtonBar.Create();
            if (selectMode)
            {
                _btnSelect = DialogButtonBar.Make("引用此模板", (s, e) =>
                {
                    if (_editingTemplate == null) return;
                    SelectedTemplateUid = _editingTemplate.Uid;
                    IsTemplateSelected = true;
                    DialogResult = DialogResult.OK;
                }, primary: true);
                _btnSelect.Enabled = false;
                _btnClose = DialogButtonBar.Make("取消");
            }
            else
            {
                _btnClose = DialogButtonBar.Make("关闭");
            }
            _btnClose.Click += (s, e) => Close();
            _bar.Controls.Add(_btnClose);
            if (_btnSelect != null) _bar.Controls.Add(_btnSelect);

            Controls.AddRange(new Control[]
            {
                _lblList, new ThemeBordered(_lstTemplates), _lblName, _fldName,
                _lblPrev, new ThemeBordered(_lblPreview), _btnNew, _btnSave, _btnDelete, _btnEditConfig, _bar
            });
            LayoutBody();
            Resize += (_, _) => LayoutBody();
            CancelButton = _btnClose;
            if (_btnSelect != null) AcceptButton = _btnSelect;

            ResumeLayout(false);
        }

        /// <summary>固定尺寸对话框，但按钮宽度和行高跟着皮肤变，所以排版仍要现算。</summary>
        private void LayoutBody()
        {
            int m = ThemeMetrics.Margin, gap = ThemeMetrics.Gap, group = ThemeMetrics.Group;
            int rowH = ThemeMetrics.RowH;
            int colW = 220;
            int rightX = m + colW + group;
            int rightW = ClientSize.Width - rightX - m;
            int bottom = ClientSize.Height - _bar.Height - m;

            _lblList.Location = new Point(m, ThemeMetrics.RowPad);
            var listHost = (Control)_lstTemplates.Parent!;
            listHost.SetBounds(m, _lblList.Bottom + gap, colW, bottom - _lblList.Bottom - gap);

            _lblName.Location = new Point(rightX, ThemeMetrics.RowPad);
            _fldName.SetBounds(rightX, _lblName.Bottom + gap, rightW, rowH);

            _lblPrev.Location = new Point(rightX, _fldName.Bottom + group);
            int prevTop = _lblPrev.Bottom + gap;
            var prevHost = (Control)_lblPreview.Parent!;
            prevHost.SetBounds(rightX, prevTop, rightW, 120);

            int y = prevHost.Bottom + group;
            int x = rightX;
            foreach (var b in new[] { _btnNew, _btnSave, _btnDelete })
            {
                b.SetBounds(x, y, ThemeButton.FitWidth(b.Text), rowH);
                x += b.Width + gap;
            }
            _btnEditConfig.SetBounds(rightX, y + rowH + gap, ThemeButton.FitWidth(_btnEditConfig.Text), rowH);

            if (_btnSelect != null)
                DialogButtonBar.LayoutRight(_bar, _btnClose, _btnSelect);
            else
                DialogButtonBar.LayoutRight(_bar, _btnClose);
        }

        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _bar.BackColor = p.Chrome;

            foreach (var lbl in new[] { _lblList, _lblName, _lblPrev })
                lbl.ForeColor = p.Muted;

            _lstTemplates.BackColor = p.Surface;
            _lstTemplates.ForeColor = p.Text;

            // 预览块靠 ThemeBordered 的描边和正文区拉开，底色直接用 Surface；
            // 套 CellTint 的淡色只适合小格子，摊到大面板上深色皮肤里就成一块脏灰
            _lblPreview.BackColor = p.Surface;
            _lblPreview.ForeColor = p.Text;

            _fldName.ApplyTheme();
            foreach (var b in new[] { _btnNew, _btnSave, _btnDelete, _btnEditConfig, _btnClose, _btnSelect })
                if (b != null) b.BackColor = p.Chrome;

            LayoutBody();
            Invalidate(true);
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
