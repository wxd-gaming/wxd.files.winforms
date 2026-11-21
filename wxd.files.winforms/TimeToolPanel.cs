using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 时间戳页签：输入秒/毫秒/微秒时间戳或日期文本，输出 UTC、本地时间与三种时间戳。
    /// 顶部的"当前时间"每秒刷新，只在页签可见时跑计时器。
    /// </summary>
    internal sealed class TimeToolPanel : ToolPanelBase
    {
        private readonly ThemeField _field;
        private readonly TextBox _output = MakeEditor(readOnly: true);
        private readonly ThemeBordered _border;
        private readonly Label _lblNow = new();
        private readonly System.Windows.Forms.Timer _clock = new();

        internal TimeToolPanel()
        {
            _field = AddField("Unix 时间戳（秒/毫秒/微秒）或 2024-01-01 12:00:00", 420);
            _field.Edit.KeyDown += (_, e) =>
            {
                if (e.KeyCode != Keys.Enter) return;
                e.Handled = true;
                e.SuppressKeyPress = true;
                ConvertInput();
            };

            AddButton("现在", (_, _) =>
            {
                _field.Edit.Text = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString(CultureInfo.InvariantCulture);
                ConvertInput();
            }, primary: true);
            AddButton("转换", (_, _) => ConvertInput());
            AddButton("复制结果", (_, _) => CopyToClipboard(_output.Text));

            _lblNow.Dock = DockStyle.Top;
            _lblNow.Height = 28;
            _lblNow.TextAlign = ContentAlignment.MiddleLeft;
            _lblNow.Padding = new Padding(ThemeMetrics.Margin, 0, 0, 0);
            _lblNow.Font = Theme.Mono;
            _lblNow.BackColor = Theme.P.Window;
            _lblNow.ForeColor = Theme.P.Muted;

            _border = new ThemeBordered(_output) { Dock = DockStyle.Fill };

            var content = new Panel { Dock = DockStyle.Fill, BackColor = Theme.P.Window };
            content.Controls.Add(_border);
            content.Controls.Add(_lblNow);

            FinishChrome(content);

            _clock.Interval = 1000;
            _clock.Tick += (_, _) => UpdateNowLabel();
            VisibleChanged += (_, _) => UpdateClock();
            Disposed += (_, _) => _clock.Dispose();

            _output.Text = "输入时间戳或日期文本后点「转换」。" + Environment.NewLine
                         + "时间戳按数值大小自动识别：秒 / 毫秒 / 微秒。";
            SetStatus("支持 1700000000、1700000000123、2024-01-01 12:00:00、ISO 8601");
            UpdateNowLabel();
        }

        protected override void ApplyThemeCore()
        {
            _border.ApplyTheme();
            ApplyEditorTheme(_output);
            _lblNow.BackColor = Theme.P.Window;
            _lblNow.ForeColor = Theme.P.Muted;
            _lblNow.Font = Theme.Mono;
        }

        private void UpdateClock()
        {
            if (Visible && !_clock.Enabled) _clock.Start();
            else if (!Visible && _clock.Enabled) _clock.Stop();
            if (Visible) UpdateNowLabel();
        }

        private void UpdateNowLabel()
        {
            var now = DateTimeOffset.Now;
            long ms = now.ToUnixTimeMilliseconds();
            _lblNow.Text = "当前："
                + now.ToUnixTimeSeconds().ToString(CultureInfo.InvariantCulture) + " 秒 / "
                + ms.ToString(CultureInfo.InvariantCulture) + " 毫秒    UTC "
                + now.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private void ConvertInput()
        {
            var result = TimeToolEngine.Convert(_field.Edit.Text);
            if (!result.Success)
            {
                SetStatus(result.Error ?? "无法识别该时间格式", error: true);
                return;
            }

            _output.Text = result.Report;
            _output.SelectionStart = 0;
            _output.SelectionLength = 0;
            SetStatus("已换算 · " + result.InputKind);
        }
    }
}
