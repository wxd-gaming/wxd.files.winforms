using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace wxd.files.winforms;

/// <summary>同名冲突询问框：覆盖 / 跳过 / 重命名 / 取消，可对全部同类冲突应用。</summary>
internal sealed class ConflictDialog : Form
{
    private readonly ConflictInfo _info;
    private readonly CheckBox _chkApplyAll = new();
    private readonly ThemeButton _btnOverwrite = new();
    private readonly ThemeButton _btnSkip = new();
    private readonly ThemeButton _btnRename = new();
    private readonly ThemeButton _btnCancel = new();

    /// <summary>用户的选择；关闭窗口视为取消。</summary>
    public ConflictAnswer Answer { get; private set; } = new(ConflictDecision.Cancel);

    private ConflictDialog(ConflictInfo info)
    {
        _info = info;
        Text = "文件已存在";
        ClientSize = new Size(600, 330);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowIcon = false;
        BackColor = Theme.P.Window;
        ForeColor = Theme.P.Text;
        Font = Theme.Body;

        BuildUi();
        ApplyTheme();
        Theme.Changed += ApplyTheme;
        Disposed += (_, _) => Theme.Changed -= ApplyTheme;
    }

    /// <summary>在 UI 线程弹框；没有 owner 时居中屏幕。</summary>
    public static ConflictAnswer Ask(IWin32Window? owner, ConflictInfo info)
    {
        using var dialog = new ConflictDialog(info);
        if (owner != null) dialog.ShowDialog(owner);
        else dialog.ShowDialog();
        return dialog.Answer;
    }

    private void BuildUi()
    {
        var body = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(ThemeMetrics.Margin, ThemeMetrics.Gap,
                ThemeMetrics.Margin, ThemeMetrics.RowPad)
        };

        var title = new Label
        {
            Text = _info.IsDirectory ? "目标位置已存在同名文件夹，继续合并吗？" : "目标位置已存在同名文件",
            Font = Theme.TitleBold,
            AutoSize = false,
            Height = 30,
            Dock = DockStyle.Top,
            TextAlign = ContentAlignment.MiddleLeft
        };

        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 5,
            Padding = new Padding(0, ThemeMetrics.Gap, 0, 0)
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        AddRow(grid, 0, "来源", _info.SourcePath);
        AddRow(grid, 1, "目标", _info.TargetPath);
        AddRow(grid, 2, "来源信息", Describe(_info.SourceSize, _info.SourceModifiedUtc));
        AddRow(grid, 3, "目标信息", Describe(_info.TargetSize, _info.TargetModifiedUtc));

        _chkApplyAll.Text = "对全部同类冲突应用此选择";
        _chkApplyAll.AutoSize = true;
        _chkApplyAll.Margin = new Padding(0, ThemeMetrics.Gap, 0, 0);
        grid.Controls.Add(_chkApplyAll, 1, 4);

        body.Controls.Add(grid);
        body.Controls.Add(title);

        var bar = DialogButtonBar.Create(out _);
        _btnOverwrite.Text = "覆盖";
        _btnSkip.Text = "跳过";
        _btnRename.Text = "重命名";
        _btnCancel.Text = "取消";
        _btnRename.Primary = true;

        _btnOverwrite.Click += (_, _) => Choose(ConflictDecision.Overwrite);
        _btnSkip.Click += (_, _) => Choose(ConflictDecision.Skip);
        _btnRename.Click += (_, _) => Choose(ConflictDecision.Rename);
        _btnCancel.Click += (_, _) => Choose(ConflictDecision.Cancel);

        bar.Controls.AddRange(new Control[] { _btnOverwrite, _btnSkip, _btnRename, _btnCancel });
        DialogButtonBar.LayoutRight(bar, _btnOverwrite, _btnSkip, _btnRename, _btnCancel);
        bar.Resize += (_, _) => DialogButtonBar.LayoutRight(bar, _btnOverwrite, _btnSkip, _btnRename, _btnCancel);

        Controls.Add(body);
        Controls.Add(bar);
        CancelButton = _btnCancel;
        AcceptButton = _btnRename;
    }

    private void Choose(ConflictDecision decision)
    {
        ConflictPolicy? applyAll = null;
        if (_chkApplyAll.Checked)
        {
            applyAll = decision switch
            {
                ConflictDecision.Overwrite => ConflictPolicy.Overwrite,
                ConflictDecision.Skip => ConflictPolicy.Skip,
                ConflictDecision.Rename => ConflictPolicy.Rename,
                _ => null
            };
        }

        Answer = new ConflictAnswer(decision, applyAll);
        DialogResult = decision == ConflictDecision.Cancel ? DialogResult.Cancel : DialogResult.OK;
        Close();
    }

    private static void AddRow(TableLayoutPanel grid, int row, string caption, string value)
    {
        var label = new Label
        {
            Text = caption,
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Theme.P.Muted
        };
        var text = new Label
        {
            Text = value,
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            AutoEllipsis = true,
            ForeColor = Theme.P.Text
        };
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        grid.Controls.Add(label, 0, row);
        grid.Controls.Add(text, 1, row);
    }

    private static string Describe(long size, DateTime modifiedUtc)
    {
        var when = modifiedUtc == default ? "未知" : modifiedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
        return FormatSize(size) + "  ·  " + when;
    }

    private static string FormatSize(long bytes)
    {
        if (bytes < 1024) return bytes + " B";
        string[] units = { "KB", "MB", "GB", "TB" };
        double value = bytes;
        int unit = -1;
        while (value >= 1024 && unit < units.Length - 1)
        {
            value /= 1024;
            unit++;
        }
        return value.ToString("0.##") + " " + units[unit];
    }

    private void ApplyTheme()
    {
        var p = Theme.P;
        BackColor = p.Window;
        ForeColor = p.Text;
        _chkApplyAll.ForeColor = p.Text;
        _chkApplyAll.BackColor = p.Window;
        foreach (Control control in Controls) ApplyTree(control, p);
    }

    private static void ApplyTree(Control root, UiTheme p)
    {
        root.BackColor = p.Chrome;
        foreach (Control child in root.Controls)
        {
            if (child is Label label)
            {
                label.BackColor = Color.Transparent;
                if (label.Font == Theme.TitleBold) label.ForeColor = p.Text;
            }
            else if (child is Panel panel)
            {
                panel.BackColor = p.Chrome;
            }
            ApplyTree(child, p);
        }
    }
}
