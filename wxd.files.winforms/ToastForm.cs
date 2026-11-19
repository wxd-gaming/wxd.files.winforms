using System;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms;

/// <summary>右下角飘窗提示：无边框、不抢焦点、自动消失。</summary>
internal sealed class ToastForm : Form
{
    private readonly System.Windows.Forms.Timer _timer;
    private static ToastForm? _current;

    private ToastForm(string message, bool warning, int durationMs)
    {
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        ShowInTaskbar = false;
        TopMost = true;
        Font = Theme.Body;
        BackColor = Theme.P.Surface;
        ForeColor = Theme.P.Text;
        Padding = new Padding(0);

        var accent = new Panel
        {
            Dock = DockStyle.Left,
            Width = 4,
            BackColor = warning ? Color.FromArgb(210, 120, 30) : Theme.P.Accent
        };

        var label = new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = false,
            Text = message,
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Theme.P.Text,
            BackColor = Theme.P.Surface,
            Padding = new Padding(12, 9, 12, 9),
            Font = Theme.Body
        };

        Controls.Add(label);
        Controls.Add(accent);

        var measured = TextRenderer.MeasureText(message, Theme.Body,
            new Size(340, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
        int width = Math.Min(400, Math.Max(260, measured.Width + 56));
        int height = Math.Max(52, measured.Height + 28);
        ClientSize = new Size(width, height);

        var area = Screen.FromPoint(Cursor.Position).WorkingArea;
        Location = new Point(area.Right - Width - 16, area.Bottom - Height - 16);

        Click += (_, _) => Close();
        label.Click += (_, _) => Close();

        _timer = new System.Windows.Forms.Timer { Interval = Math.Max(1500, durationMs) };
        _timer.Tick += (_, _) =>
        {
            _timer.Stop();
            Close();
        };
    }

    public static void ShowToast(string message, bool warning = true, int durationMs = 3500)
    {
        try
        {
            if (_current is { IsDisposed: false })
                _current.Close();

            var toast = new ToastForm(message, warning, durationMs);
            _current = toast;
            toast.FormClosed += (_, _) =>
            {
                if (ReferenceEquals(_current, toast)) _current = null;
            };
            toast.Show();
        }
        catch
        {
            // 提示失败不影响主流程
        }
    }

    protected override bool ShowWithoutActivation => true;

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE：不抢焦点
            cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW：不显示任务栏按钮
            return cp;
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) _timer?.Dispose();
        base.Dispose(disposing);
    }
}
