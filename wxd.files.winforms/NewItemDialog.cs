using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace wxd.files.winforms;

/// <summary>新建文件/文件夹的轻量输入框，带名称校验与错误提示。</summary>
internal sealed class NewItemDialog : Form
{
    private readonly TextBox _txtName = new();
    private readonly Label _lblError = new();
    private readonly ThemeButton _btnOk = new();
    private readonly ThemeButton _btnCancel = new();
    private readonly Func<string, string?>? _validate;

    public string Value => _txtName.Text.Trim();

    public NewItemDialog(string title, string prompt, string initialName, Func<string, string?>? validate = null)
    {
        _validate = validate;

        Text = title;
        ClientSize = new Size(480, 190);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowIcon = false;
        Font = Theme.Body;

        BuildUi(title, prompt, initialName);
        ApplyTheme();
        Theme.Changed += ApplyTheme;
        Disposed += (_, _) => Theme.Changed -= ApplyTheme;
    }

    private void BuildUi(string title, string prompt, string initialName)
    {
        var body = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(ThemeMetrics.Margin, ThemeMetrics.Gap,
                ThemeMetrics.Margin, ThemeMetrics.RowPad)
        };

        var lblTitle = new Label
        {
            Text = title,
            Dock = DockStyle.Top,
            Height = 28,
            TextAlign = ContentAlignment.MiddleLeft,
            Font = Theme.TitleBold
        };

        var lblPrompt = new Label
        {
            Text = prompt,
            Dock = DockStyle.Top,
            Height = 22,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Theme.P.Muted
        };

        _txtName.Dock = DockStyle.Top;
        _txtName.Text = initialName;
        _txtName.Font = Theme.List;

        _lblError.Dock = DockStyle.Top;
        _lblError.Height = 38;
        _lblError.TextAlign = ContentAlignment.MiddleLeft;
        _lblError.ForeColor = Theme.Signal(Color.FromArgb(200, 60, 60));
        _lblError.AutoEllipsis = true;

        body.Controls.Add(_lblError);
        body.Controls.Add(_txtName);
        body.Controls.Add(lblPrompt);
        body.Controls.Add(lblTitle);

        var bar = DialogButtonBar.Create(out _);
        _btnOk.Text = "确定";
        _btnOk.Primary = true;
        _btnCancel.Text = "取消";
        _btnOk.Click += (_, _) => Confirm();
        _btnCancel.Click += (_, _) =>
        {
            DialogResult = DialogResult.Cancel;
            Close();
        };

        bar.Controls.AddRange(new Control[] { _btnOk, _btnCancel });
        DialogButtonBar.LayoutRight(bar, _btnOk, _btnCancel);
        bar.Resize += (_, _) => DialogButtonBar.LayoutRight(bar, _btnOk, _btnCancel);

        Controls.Add(body);
        Controls.Add(bar);

        AcceptButton = _btnOk;
        CancelButton = _btnCancel;

        Shown += (_, _) =>
        {
            var stem = Path.GetFileNameWithoutExtension(initialName);
            _txtName.Focus();
            _txtName.Select(0, Math.Max(0, stem.Length));
        };
    }

    private void Confirm()
    {
        var error = _validate?.Invoke(Value);
        if (!string.IsNullOrEmpty(error))
        {
            _lblError.Text = error;
            _txtName.Focus();
            return;
        }

        DialogResult = DialogResult.OK;
        Close();
    }

    private void ApplyTheme()
    {
        var p = Theme.P;
        BackColor = p.Window;
        ForeColor = p.Text;
        _txtName.BackColor = p.Surface;
        _txtName.ForeColor = p.Text;

        foreach (Control control in AllControls(this))
        {
            if (control is Label label)
            {
                label.BackColor = Color.Transparent;
                if (label.Font == Theme.TitleBold) label.ForeColor = p.Text;
            }
            else if (control is Panel panel)
            {
                panel.BackColor = p.Window;
            }
            else if (control is ThemeButton button)
            {
                button.BackColor = p.Chrome;
                button.ForeColor = p.Text;
            }
        }

        _lblError.ForeColor = Theme.Signal(Color.FromArgb(200, 60, 60));
    }

    private static System.Collections.Generic.IEnumerable<Control> AllControls(Control root)
    {
        foreach (Control child in root.Controls)
        {
            yield return child;
            foreach (var nested in AllControls(child)) yield return nested;
        }
    }
}
