using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms;

internal enum ToastKind { Info, Warning, Error }

/// <summary>右下角飘窗提示：多实例向上叠加、默认 2 秒后淡出，不抢焦点。</summary>
internal sealed class ToastForm : Form
{
    private const int FadeDurationMs = 260;
    private const int StackGap = 8;
    private const int ScreenMargin = 16;
    private const int MaxVisible = 6;

    private static readonly List<ToastForm> Active = new();

    private readonly System.Windows.Forms.Timer _lifeTimer;
    private readonly System.Windows.Forms.Timer _fadeTimer;
    private DateTime _fadeStartedAt;
    private bool _closing;

    private ToastForm(string message, ToastKind kind, int durationMs)
    {
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        ShowInTaskbar = false;
        TopMost = true;
        Font = Theme.Body;
        BackColor = Theme.P.Surface;
        ForeColor = Theme.P.Text;
        Opacity = 1;

        var accent = new Panel
        {
            Dock = DockStyle.Left,
            Width = 4,
            BackColor = kind switch
            {
                ToastKind.Error => Color.FromArgb(196, 43, 43),
                ToastKind.Warning => Color.FromArgb(210, 120, 30),
                _ => Theme.P.Accent
            }
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

        Click += (_, _) => BeginFade();
        label.Click += (_, _) => BeginFade();

        _lifeTimer = new System.Windows.Forms.Timer { Interval = Math.Max(500, durationMs) };
        _lifeTimer.Tick += (_, _) => BeginFade();

        _fadeTimer = new System.Windows.Forms.Timer { Interval = 15 };
        _fadeTimer.Tick += (_, _) => StepFade();

        FormClosed += (_, _) =>
        {
            _lifeTimer.Stop();
            _fadeTimer.Stop();
            Active.Remove(this);
            RepositionAll();
        };
    }

    internal static string Shorten(string? text, int max = 160)
    {
        var line = (text ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault() ?? "";
        return line.Length <= max ? line : line[..max] + "…";
    }

    public static void ShowToast(string message, ToastKind kind = ToastKind.Warning, int durationMs = 2000)
    {
        IntPtr focus = IntPtr.Zero;
        IntPtr foreground = IntPtr.Zero;
        try
        {
            focus = GetFocus();
            foreground = GetForegroundWindow();

            while (Active.Count >= MaxVisible)
            {
                var oldest = Active[0];
                Active.RemoveAt(0);
                oldest.Close();
            }

            var toast = new ToastForm(message, kind, durationMs);
            Active.Add(toast);
            RepositionAll();
            toast.Show();
            toast._lifeTimer.Start();
        }
        catch
        {
            // 提示失败不影响主流程
        }
        finally
        {
            // 某些主题/输入法下 Show() 仍会把焦点带走，这里强制还给原控件
            if (focus != IntPtr.Zero && GetFocus() != focus)
            {
                try { SetFocus(focus); } catch { }
            }

            if (foreground != IntPtr.Zero && GetForegroundWindow() != foreground)
            {
                try { SetForegroundWindow(foreground); } catch { }
            }
        }
    }

    private static void RepositionAll()
    {
        if (Active.Count == 0) return;

        var area = (Screen.PrimaryScreen ?? Screen.FromPoint(Cursor.Position)).WorkingArea;
        int y = area.Bottom - ScreenMargin;

        // 最新的在最下面，旧的依次往上叠
        for (int i = Active.Count - 1; i >= 0; i--)
        {
            var toast = Active[i];
            if (toast.IsDisposed) continue;

            y -= toast.Height;
            int x = area.Right - toast.Width - ScreenMargin;
            if (y < area.Top)
                y = area.Top;

            toast.Location = new Point(x, y);
            y -= StackGap;
        }
    }

    private void BeginFade()
    {
        if (_closing || IsDisposed) return;
        _closing = true;
        _lifeTimer.Stop();
        _fadeStartedAt = DateTime.UtcNow;
        _fadeTimer.Start();
    }

    private void StepFade()
    {
        if (IsDisposed)
        {
            _fadeTimer.Stop();
            return;
        }

        double elapsed = (DateTime.UtcNow - _fadeStartedAt).TotalMilliseconds;
        double progress = Math.Min(1, elapsed / FadeDurationMs);
        Opacity = 1 - progress;

        if (progress >= 1)
        {
            _fadeTimer.Stop();
            Close();
        }
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr GetFocus();

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr SetFocus(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

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
        if (disposing)
        {
            _lifeTimer?.Dispose();
            _fadeTimer?.Dispose();
        }
        base.Dispose(disposing);
    }
}
