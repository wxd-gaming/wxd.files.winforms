using System;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// Cron 页签：解析 5 段 / 6 段表达式，给出逐字段中文说明和接下来 10 次触发时间。
    /// 输入变化后防抖计算，回车立即计算。
    /// </summary>
    internal sealed class CronToolPanel : ToolPanelBase
    {
        private const int NextRunCount = 10;

        private readonly ThemeField _field;
        private readonly TextBox _output = MakeEditor(readOnly: true);
        private readonly ThemeBordered _border;
        private readonly System.Windows.Forms.Timer _debounce = new();

        internal CronToolPanel()
        {
            _field = AddField("例如 */5 * * * * 或 0 0/5 * * * ?", 420);

            AddButton("解析", (_, _) => ParseNow(), primary: true);
            AddButton("示例", (_, _) =>
            {
                _field.Edit.Text = "*/5 9-18 * * 1-5";
                _field.Edit.SelectionStart = _field.Edit.TextLength;
                ParseNow();
            });
            AddButton("复制结果", (_, _) => CopyToClipboard(_output.Text));

            _field.Edit.TextChanged += (_, _) =>
            {
                _debounce.Stop();
                if (_field.Edit.TextLength > 0) _debounce.Start();
            };
            _field.Edit.KeyDown += (_, e) =>
            {
                if (e.KeyCode != Keys.Enter) return;
                e.Handled = true;
                e.SuppressKeyPress = true;
                _debounce.Stop();
                ParseNow();
            };

            _border = new ThemeBordered(_output) { Dock = DockStyle.Fill };
            FinishChrome(_border);

            _debounce.Interval = 350;
            _debounce.Tick += (_, _) =>
            {
                _debounce.Stop();
                ParseNow();
            };
            Disposed += (_, _) => _debounce.Dispose();

            _output.Text = "输入 cron 表达式后自动解析。" + Environment.NewLine
                         + "支持 5 段（分 时 日 月 周）和 6 段（秒 分 时 日 月 周）。";
            SetStatus("支持 * , - / 与英文月份/星期名；? 当作 *，不支持 L / W / #");
        }

        protected override void ApplyThemeCore()
        {
            _border.ApplyTheme();
            ApplyEditorTheme(_output);
        }

        private void ParseNow()
        {
            string expression = _field.Edit.Text.Trim();
            if (expression.Length == 0)
            {
                SetStatus("请输入 cron 表达式", error: true);
                return;
            }

            var schedule = CronToolEngine.Parse(expression);
            if (!schedule.Success)
            {
                SetStatus(schedule.Error ?? "cron 解析失败", error: true);
                return;
            }

            var runs = CronToolEngine.NextRuns(schedule, DateTime.Now, NextRunCount);
            var sb = new StringBuilder();
            sb.AppendLine("表达式 : " + schedule.Expression);
            sb.AppendLine("类型   : " + (schedule.WithSeconds ? "6 段（秒 分 时 日 月 周）" : "5 段（分 时 日 月 周）"));
            sb.AppendLine("说明   : " + schedule.Summary);
            sb.AppendLine();
            sb.AppendLine("字段明细");

            string[] labels = schedule.WithSeconds
                ? new[] { "秒", "分", "时", "日", "月", "周" }
                : new[] { "分", "时", "日", "月", "周" };
            for (int i = 0; i < schedule.FieldNotes.Length; i++)
                sb.AppendLine("  " + labels[i] + "    " + schedule.FieldNotes[i]);

            sb.AppendLine();
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);
            string offsetText = (offset < TimeSpan.Zero ? "-" : "+")
                + $"{Math.Abs(offset.Hours):00}:{Math.Abs(offset.Minutes):00}";
            sb.AppendLine($"接下来 {NextRunCount} 次（本地时间 UTC{offsetText}）");
            if (runs.Count == 0)
            {
                sb.AppendLine("  8 年内没有可匹配的时间，请检查表达式");
            }
            else
            {
                int index = 1;
                foreach (var run in runs)
                {
                    sb.AppendLine($"  {index++,2}. {run.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)}  {CronToolEngine.WeekdayName(run)}");
                    if (index > NextRunCount) break;
                }
            }

            _output.Text = sb.ToString();
            _output.SelectionStart = 0;
            _output.SelectionLength = 0;
            SetStatus($"解析成功 · 找到 {runs.Count} 次触发时间");
        }
    }
}
