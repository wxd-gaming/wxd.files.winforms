using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace wxd.files.winforms;

/// <summary>批量重命名窗口：规则 + 两栏实时预览 + 执行/撤销。</summary>
internal sealed class BatchRenameForm : Form
{
    private readonly List<string> _paths;
    private readonly List<RenameJournalEntry> _journal = new();

    private readonly TextBox _txtFind = new();
    private readonly TextBox _txtReplace = new();
    private readonly TextBox _txtTemplate = new();
    private readonly CheckBox _chkRegex = new();
    private readonly CheckBox _chkCaseSensitive = new();
    private readonly CheckBox _chkIncludeExtension = new();
    private readonly NumericUpDown _numStart = new();
    private readonly NumericUpDown _numStep = new();
    private readonly NumericUpDown _numDigits = new();
    private readonly ComboBox _cboCase = new();

    private readonly ListView _list = new();
    private readonly Label _lblStatus = new();
    private readonly ThemeButton _btnApply = new();
    private readonly ThemeButton _btnUndo = new();
    private readonly ThemeButton _btnClose = new();

    private RenamePlan? _plan;

    public BatchRenameForm(IReadOnlyList<string> paths)
    {
        _paths = paths.ToList();

        Text = "批量重命名";
        ClientSize = new Size(940, 600);
        MinimumSize = new Size(760, 420);
        StartPosition = FormStartPosition.CenterParent;
        ShowIcon = false;
        Font = Theme.Body;

        BuildUi();
        ApplyTheme();
        Theme.Changed += ApplyTheme;
        Disposed += (_, _) => Theme.Changed -= ApplyTheme;

        _txtTemplate.Text = "{name}";
        RefreshPreview();
    }

    private void BuildUi()
    {
        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 110,
            ColumnCount = 6,
            Padding = new Padding(ThemeMetrics.Margin, ThemeMetrics.Gap, ThemeMetrics.Margin, 0)
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 76));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));

        _chkRegex.Text = "正则";
        _chkRegex.AutoSize = true;
        _chkRegex.Anchor = AnchorStyles.Left;
        _chkCaseSensitive.Text = "区分大小写";
        _chkCaseSensitive.AutoSize = true;
        _chkCaseSensitive.Anchor = AnchorStyles.Left;
        _chkIncludeExtension.Text = "扩展名参与查找替换";
        _chkIncludeExtension.AutoSize = true;
        _chkIncludeExtension.Anchor = AnchorStyles.Left;

        _cboCase.DropDownStyle = ComboBoxStyle.DropDownList;
        _cboCase.Items.AddRange(new object[] { "保持原样", "全部大写", "全部小写", "首字母大写" });
        _cboCase.SelectedIndex = 0;
        _cboCase.Width = 110;

        _numStart.Minimum = 0;
        _numStart.Maximum = 999999;
        _numStart.Value = 1;
        _numStart.Width = 64;
        _numStep.Minimum = 1;
        _numStep.Maximum = 9999;
        _numStep.Value = 1;
        _numStep.Width = 56;
        _numDigits.Minimum = 1;
        _numDigits.Maximum = 12;
        _numDigits.Value = 2;
        _numDigits.Width = 50;

        var counterFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            WrapContents = false,
            AutoSize = false,
            Padding = new Padding(0)
        };
        counterFlow.Controls.AddRange(new Control[]
        {
            new Label { Text = "起始", AutoSize = true, Padding = new Padding(0, 6, 2, 0) },
            _numStart,
            new Label { Text = "步长", AutoSize = true, Padding = new Padding(8, 6, 2, 0) },
            _numStep,
            new Label { Text = "位数", AutoSize = true, Padding = new Padding(8, 6, 2, 0) },
            _numDigits
        });

        AddRow(grid, 0, "查找", _txtFind, "替换", _txtReplace, _chkRegex, _chkCaseSensitive);
        AddRow(grid, 1, "模板", _txtTemplate, "大小写", _cboCase, null, null);
        AddRow(grid, 2, "序号", counterFlow, "", null, _chkIncludeExtension, null);

        grid.SetColumnSpan(counterFlow, 2);
        grid.SetColumnSpan(_chkIncludeExtension, 2);

        _list.Dock = DockStyle.Fill;
        _list.View = View.Details;
        _list.FullRowSelect = true;
        _list.HideSelection = false;
        _list.MultiSelect = false;
        _list.Columns.Add("原名称", 240);
        _list.Columns.Add("新名称", 240);
        _list.Columns.Add("状态", 90);
        _list.Columns.Add("说明", 300);

        var bottom = new Panel { Dock = DockStyle.Bottom, Height = ThemeMetrics.BarHeight };
        _lblStatus.Dock = DockStyle.Fill;
        _lblStatus.TextAlign = ContentAlignment.MiddleLeft;
        _lblStatus.Padding = new Padding(ThemeMetrics.Margin, 0, 0, 0);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            Width = 300,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(0, ThemeMetrics.RowPad, ThemeMetrics.Margin, 0)
        };
        _btnApply.Text = "执行";
        _btnApply.Primary = true;
        _btnApply.Size = new Size(82, ThemeMetrics.RowH);
        _btnUndo.Text = "撤销上次";
        _btnUndo.Size = new Size(96, ThemeMetrics.RowH);
        _btnUndo.Enabled = false;
        _btnClose.Text = "关闭";
        _btnClose.Size = new Size(82, ThemeMetrics.RowH);

        _btnApply.Click += (_, _) => ApplyPlan();
        _btnUndo.Click += (_, _) => UndoLast();
        _btnClose.Click += (_, _) => Close();

        buttons.Controls.AddRange(new Control[] { _btnClose, _btnUndo, _btnApply });
        bottom.Controls.Add(_lblStatus);
        bottom.Controls.Add(buttons);

        Controls.Add(_list);
        Controls.Add(grid);
        Controls.Add(bottom);

        CancelButton = _btnClose;

        _txtFind.TextChanged += (_, _) => RefreshPreview();
        _txtReplace.TextChanged += (_, _) => RefreshPreview();
        _txtTemplate.TextChanged += (_, _) => RefreshPreview();
        _chkRegex.CheckedChanged += (_, _) => RefreshPreview();
        _chkCaseSensitive.CheckedChanged += (_, _) => RefreshPreview();
        _chkIncludeExtension.CheckedChanged += (_, _) => RefreshPreview();
        _numStart.ValueChanged += (_, _) => RefreshPreview();
        _numStep.ValueChanged += (_, _) => RefreshPreview();
        _numDigits.ValueChanged += (_, _) => RefreshPreview();
        _cboCase.SelectedIndexChanged += (_, _) => RefreshPreview();
    }

    private static void AddRow(
        TableLayoutPanel grid,
        int row,
        string label1,
        Control field1,
        string label2,
        Control? field2,
        Control? trailing1,
        Control? trailing2)
    {
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

        grid.Controls.Add(NewLabel(label1), 0, row);
        field1.Dock = DockStyle.Fill;
        grid.Controls.Add(field1, 1, row);

        if (label2.Length > 0) grid.Controls.Add(NewLabel(label2), 2, row);
        if (field2 != null)
        {
            field2.Dock = DockStyle.Fill;
            grid.Controls.Add(field2, 3, row);
        }
        if (trailing1 != null) grid.Controls.Add(trailing1, 4, row);
        if (trailing2 != null) grid.Controls.Add(trailing2, 5, row);
    }

    private static Label NewLabel(string text) => new()
    {
        Text = text,
        Dock = DockStyle.Fill,
        TextAlign = ContentAlignment.MiddleLeft,
        AutoSize = false
    };

    private RenameRule BuildRule() => new()
    {
        Find = _txtFind.Text,
        Replace = _txtReplace.Text,
        UseRegex = _chkRegex.Checked,
        CaseSensitive = _chkCaseSensitive.Checked,
        IncludeExtensionInFind = _chkIncludeExtension.Checked,
        Template = _txtTemplate.Text,
        CounterStart = (int)_numStart.Value,
        CounterStep = (int)_numStep.Value,
        CounterDigits = (int)_numDigits.Value,
        CaseMode = _cboCase.SelectedIndex switch
        {
            1 => RenameCaseMode.Upper,
            2 => RenameCaseMode.Lower,
            3 => RenameCaseMode.Title,
            _ => RenameCaseMode.None
        }
    };

    private void RefreshPreview()
    {
        try
        {
            _plan = BatchRenameEngine.BuildPlan(_paths, BuildRule());
        }
        catch (Exception ex)
        {
            _plan = null;
            _lblStatus.Text = "规则无效: " + ex.Message;
            _btnApply.Enabled = false;
            _list.Items.Clear();
            return;
        }

        _list.BeginUpdate();
        try
        {
            _list.Items.Clear();
            foreach (var item in _plan.Items)
            {
                var row = new ListViewItem(new[]
                {
                    item.SourceName,
                    item.TargetName,
                    StatusText(item.Status),
                    item.Message ?? ""
                });
                row.ForeColor = item.Status switch
                {
                    RenamePlanStatus.Ready => Theme.P.Text,
                    RenamePlanStatus.Unchanged => Theme.P.Muted,
                    _ => Theme.Signal(Color.FromArgb(200, 60, 60))
                };
                _list.Items.Add(row);
            }
        }
        finally
        {
            _list.EndUpdate();
        }

        int conflicts = _plan.Items.Count(i => i.Status == RenamePlanStatus.Conflict);
        int invalid = _plan.Items.Count(i => i.Status == RenamePlanStatus.Invalid);
        int duplicates = _plan.Items.Count(i => i.Status == RenamePlanStatus.Duplicate);

        _lblStatus.Text = "共 " + _plan.Items.Count + " 项，待改 " + _plan.ChangeCount
            + " 项" + (conflicts > 0 ? "，冲突 " + conflicts : "")
            + (duplicates > 0 ? "，重名 " + duplicates : "")
            + (invalid > 0 ? "，无效 " + invalid : "");

        _btnApply.Enabled = !_plan.HasErrors && _plan.ChangeCount > 0;
    }

    private static string StatusText(RenamePlanStatus status) => status switch
    {
        RenamePlanStatus.Ready => "待改",
        RenamePlanStatus.Unchanged => "未变化",
        RenamePlanStatus.Conflict => "冲突",
        RenamePlanStatus.Duplicate => "重名",
        _ => "无效"
    };

    private void ApplyPlan()
    {
        if (_plan == null) return;

        var result = BatchRenameEngine.Apply(_plan);
        if (!result.Success)
        {
            AlertForm.Show(this, "重命名失败: " + string.Join(GlobalConst.NewLine, result.Errors),
                "批量重命名", AlertType.Error);
            return;
        }

        _journal.Clear();
        _journal.AddRange(result.Journal);

        _paths.Clear();
        foreach (var item in _plan.Items)
            _paths.Add(item.Status == RenamePlanStatus.Ready ? item.TargetPath : item.SourcePath);

        _btnUndo.Enabled = _journal.Count > 0;
        RefreshPreview();
        _lblStatus.Text = "已重命名 " + _journal.Count + " 项，可撤销";
    }

    private void UndoLast()
    {
        if (_journal.Count == 0) return;

        if (!BatchRenameEngine.Undo(_journal, out var error))
        {
            AlertForm.Show(this, "撤销失败: " + error, "批量重命名", AlertType.Error);
            return;
        }

        _paths.Clear();
        _paths.AddRange(_journal.Select(j => j.Source));
        _journal.Clear();
        _btnUndo.Enabled = false;
        RefreshPreview();
        _lblStatus.Text = "已撤销";
    }

    private void ApplyTheme()
    {
        var p = Theme.P;
        BackColor = p.Window;
        ForeColor = p.Text;
        _list.BackColor = p.Surface;
        _list.ForeColor = p.Text;
        _lblStatus.ForeColor = p.Muted;

        foreach (Control control in AllControls(this))
        {
            switch (control)
            {
                case TextBox text:
                    text.BackColor = p.Surface;
                    text.ForeColor = p.Text;
                    break;
                case CheckBox check:
                    check.ForeColor = p.Text;
                    check.BackColor = Color.Transparent;
                    break;
                case ComboBox combo:
                    combo.BackColor = p.Surface;
                    combo.ForeColor = p.Text;
                    break;
                case NumericUpDown number:
                    number.BackColor = p.Surface;
                    number.ForeColor = p.Text;
                    break;
                case ThemeButton button:
                    button.BackColor = p.Chrome;
                    button.ForeColor = p.Text;
                    break;
                case Label label:
                    label.ForeColor = p.Text;
                    break;
                case Panel panel:
                    panel.BackColor = p.Window;
                    break;
            }
        }

        _lblStatus.ForeColor = p.Muted;
    }

    private static IEnumerable<Control> AllControls(Control root)
    {
        foreach (Control child in root.Controls)
        {
            yield return child;
            foreach (var nested in AllControls(child)) yield return nested;
        }
    }
}
