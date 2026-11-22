using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// Cron 页签：可选标准 cron / Quartz / Spring 三套规范，上方解析、下方可视化生成。
    /// 生成器只负责把结果填回表达式输入框，手工改表达式不会被生成器覆盖。
    /// </summary>
    internal sealed class CronToolPanel : ToolPanelBase
    {
        private const int NextRunCount = 10;

        private static readonly string[] QuartzExamples =
        {
            "0 0/5 9-18 ? * MON-FRI",
            "0 0 12 L * ?",
            "0 0 9 ? * 6#3",
            "0 0 0 1 1 ? 2027",
            "0/30 * * * * ?",
            "0 15 10 ? * MON-FRI",
            "0 0 2 ? * SAT"
        };

        private static readonly string[] StandardExamples =
        {
            "0/5 9-18 * * MON-FRI",
            "0 3 * * *",
            "30 8 * * MON",
            "0 0 1 * *",
            "15 10 * * MON-FRI",
            "0 0/2 * * *",
            "0 0 * * SUN"
        };

        private static readonly string[] SpringExamples =
        {
            "0 0/5 9-18 * * MON-FRI",
            "0 0 12 L * *",
            "0 0 9 * * 5#3",
            "0 0 0 1 * *",
            "0 0/15 * * * *",
            "0 30 8 * * 1-5"
        };

        private readonly ThemeComboBox _dialect = new();
        private readonly ThemeField _field;
        private readonly CronBuilderControl _builder = new();
        private readonly SafeSplitContainer _split = new();
        private readonly TextBox _output = MakeEditor(readOnly: true);
        private readonly ThemeBordered _border;
        private readonly System.Windows.Forms.Timer _debounce = new();
        private bool _splitMinsSet;
        private bool _splitPinned;
        private string? _lastExample;

        internal CronToolPanel()
        {
            _dialect.Font = Theme.Body;
            _dialect.DropDownStyle = ComboBoxStyle.DropDownList;
            _dialect.BackColor = Theme.P.Surface;
            _dialect.ForeColor = Theme.P.Text;
            _dialect.Size = new Size(140, ThemeMetrics.RowH);
            _dialect.Margin = new Padding(0, 0, ThemeMetrics.Gap, 0);
            _dialect.Items.AddRange(new object[] { "标准 cron", "Quartz", "Spring" });
            _dialect.SelectedIndex = 1;
            _dialect.SelectedIndexChanged += (_, _) => OnDialectChanged();
            Top.Controls.Add(_dialect);

            _field = AddField("例如 0 0/5 9-18 ? * MON-FRI", 440);

            AddButton("解析", (_, _) => ParseNow(), primary: true);
            AddButton("示例", (_, _) => LoadExample());
            AddButton("语法参考", (_, _) => ShowReference());
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

            _split.Dock = DockStyle.Fill;
            _split.Orientation = Orientation.Horizontal;
            _split.SplitterWidth = 6;
            _split.BackColor = Theme.P.Border;
            _split.SplitterMoved += (_, _) => _splitPinned = true;
            _split.SizeChanged += (_, _) => InitSplit();
            _split.Panel1.BackColor = Theme.P.Window;
            _split.Panel2.BackColor = Theme.P.Window;

            _builder.Dock = DockStyle.Fill;
            _builder.ExpressionChanged += (_, _) => ApplyBuilderExpression();
            _split.Panel1.Controls.Add(_builder);

            // 结果是一份报告，长行换行比横向滚动条更好读
            _output.WordWrap = true;
            _output.ScrollBars = ScrollBars.Vertical;

            _border = new ThemeBordered(_output) { Dock = DockStyle.Fill, Margin = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Margin, 0) };
            _split.Panel2.Controls.Add(_border);
            FinishChrome(_split);

            _debounce.Interval = 350;
            _debounce.Tick += (_, _) =>
            {
                _debounce.Stop();
                ParseNow();
            };
            Disposed += (_, _) => _debounce.Dispose();

            _field.Edit.Text = _builder.Expression;
            _field.Edit.SelectionStart = _field.Edit.TextLength;
            _output.Text = "输入或生成表达式后自动解析。" + Environment.NewLine
                         + "上方可选 Quartz / Spring；下方的可视化生成器会实时把结果填回表达式。" + Environment.NewLine
                         + "点“语法参考”查看两套规范的完整差异与特殊字符说明。";
            SetStatus("Quartz：日/周必须有一列写 ?；周字段 1=周日。下方生成器可实时拼表达式。");
            ParseNow();
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            InitSplit();
        }

        private void InitSplit()
        {
            // 构造期 SplitContainer 还是默认的 150×100，此时设最小栏高会因当前 SplitterDistance 越界直接抛；
            // 等窗口真正有高度后再设。用户拖过分割条后不再自动改高度。
            if (_split.Height < 460) return;
            try
            {
                if (!_splitMinsSet)
                {
                    _split.Panel1MinSize = 250;
                    _split.Panel2MinSize = 160;
                    _splitMinsSet = true;
                }
                if (!_splitPinned)
                    _split.SplitterDistance = Math.Clamp(_split.Height - 210, 280, 340);
            }
            catch (InvalidOperationException)
            {
            }
        }

        protected override void ApplyThemeCore()
        {
            _dialect.Font = Theme.Body;
            _dialect.BackColor = Theme.P.Surface;
            _dialect.ForeColor = Theme.P.Text;
            _field.ApplyTheme();
            _builder.ApplyTheme();
            _split.BackColor = Theme.P.Border;
            _split.Panel1.BackColor = Theme.P.Window;
            _split.Panel2.BackColor = Theme.P.Window;
            _border.ApplyTheme();
            ApplyEditorTheme(_output);
        }

        private CronDialect CurrentDialect => _dialect.SelectedIndex switch
        {
            0 => CronDialect.Standard,
            2 => CronDialect.Spring,
            _ => CronDialect.Quartz
        };

        private void OnDialectChanged()
        {
            var dialect = CurrentDialect;
            _builder.SetDialect(dialect);
            _field.Placeholder = dialect switch
            {
                CronDialect.Standard => "例如 */5 9-18 * * 1-5",
                CronDialect.Spring => "例如 0 0/5 9-18 * * MON-FRI 或 @daily",
                _ => "例如 0 0/5 9-18 ? * MON-FRI"
            };

            SetStatus(dialect switch
            {
                CronDialect.Standard => "标准 cron：5 段（分 时 日 月 周），日/周同时限定时取“或”；不支持 L / W / #。",
                CronDialect.Spring => "Spring：6 段（秒 分 时 日 月 周），无年字段；周字段 0/7=周日；日/周同时限定时需同时满足；支持宏 @daily 等。",
                _ => "Quartz：6 段或 7 段（末尾年），日/周必须有一列写 ?；周字段 1=周日。"
            });
            ParseNow();
        }

        private void ApplyBuilderExpression()
        {
            string expression = _builder.Expression;
            if (_field.Edit.Text == expression) return;
            _field.Edit.Text = expression;
            _field.Edit.SelectionStart = _field.Edit.TextLength;
            _debounce.Stop();
            ParseNow();
        }

        private void LoadExample()
        {
            var pool = CurrentDialect switch
            {
                CronDialect.Standard => StandardExamples,
                CronDialect.Spring => SpringExamples,
                _ => QuartzExamples
            };
            string expression = PickExample(pool);
            _lastExample = expression;

            // 只填输入框的话下方生成器不会动，这里把同一份表达式灌回生成器，再以生成结果为准回填
            bool loaded = _builder.LoadExpression(expression, CurrentDialect);
            _field.Edit.Text = loaded ? _builder.Expression : expression;
            _field.Edit.SelectionStart = _field.Edit.TextLength;
            _debounce.Stop();
            ParseNow();
        }

        /// <summary>随机挑一个和上次不同、且不等于当前输入框的示例，保证连续点击能看到不同场景。</summary>
        private string PickExample(string[] pool)
        {
            if (pool.Length == 1) return pool[0];
            string current = _field.Edit.Text.Trim();
            var candidates = pool.Where(e => e != _lastExample && e != current).ToArray();
            var source = candidates.Length > 0 ? candidates : pool;
            return source[Random.Shared.Next(source.Length)];
        }

        private void ShowReference()
        {
            _debounce.Stop();
            // 源文件是 LF 时裸 \n 在 TextBox 里不换行，参考文本会挤成几行
            _output.Text = ReferenceText().Replace("\r\n", "\n").Replace("\n", Environment.NewLine);
            ResetOutputScroll();
            SetStatus("已显示 Quartz / Spring 语法参考，可点“复制结果”带走");
        }

        /// <summary>只设 SelectionStart 不能保证多行 TextBox 滚回顶行，必须把光标位置滚进视野。</summary>
        private void ResetOutputScroll()
        {
            _output.SelectionStart = 0;
            _output.SelectionLength = 0;
            _output.ScrollToCaret();
        }

        private void ParseNow()
        {
            string expression = _field.Edit.Text.Trim();
            if (expression.Length == 0)
            {
                SetStatus("请输入 cron 表达式", error: true);
                return;
            }

            var dialect = CurrentDialect;
            var schedule = CronToolEngine.Parse(expression, dialect);
            if (!schedule.Success)
            {
                _output.Text = "解析失败（" + CronToolEngine.DialectName(dialect) + "）" + Environment.NewLine
                             + Environment.NewLine + "  " + (schedule.Error ?? "cron 解析失败") + Environment.NewLine
                             + Environment.NewLine + "检查字段段数、星期编号和特殊字符；" + Environment.NewLine
                             + "需要对照说明时点上方“语法参考”。";
                ResetOutputScroll();
                SetStatus(schedule.Error ?? "cron 解析失败", error: true);
                return;
            }

            var runs = CronToolEngine.NextRuns(schedule, DateTime.Now, NextRunCount);
            var sb = new StringBuilder();
            sb.AppendLine("规范   : " + schedule.DialectName);
            sb.AppendLine("表达式 : " + schedule.Expression);
            if (!string.IsNullOrEmpty(schedule.Macro))
                sb.AppendLine("宏     : " + schedule.Macro + "  →  " + schedule.Normalized);
            if (!string.IsNullOrEmpty(schedule.Warning))
                sb.AppendLine("提示   : " + schedule.Warning);
            sb.AppendLine("说明   : " + schedule.Summary);
            sb.AppendLine();
            sb.AppendLine("字段明细");
            for (int i = 0; i < schedule.FieldNotes.Length; i++)
                sb.AppendLine("  " + schedule.FieldLabels[i].PadRight(2) + "    " + schedule.FieldNotes[i]);

            sb.AppendLine();
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);
            string offsetText = (offset < TimeSpan.Zero ? "-" : "+")
                + $"{Math.Abs(offset.Hours):00}:{Math.Abs(offset.Minutes):00}";
            sb.AppendLine($"接下来 {NextRunCount} 次（本地时间 UTC{offsetText}）");
            if (runs.Count == 0)
            {
                sb.AppendLine(schedule.Dialect == CronDialect.Quartz
                    ? "  指定年份范围内没有可匹配的时间，请检查表达式"
                    : "  8 年内没有可匹配的时间，请检查表达式");
            }
            else
            {
                int index = 1;
                foreach (var run in runs)
                {
                    sb.AppendLine($"  {index++,2}. {run.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)}  {CronToolEngine.WeekdayName(run)}");
                }
            }

            _output.Text = sb.ToString();
            ResetOutputScroll();
            SetStatus($"{schedule.DialectName} · 解析成功 · 找到 {runs.Count} 次触发时间");
        }

        private static string ReferenceText() => """
Cron 语法参考（Quartz / Spring）
================================================================

字段对照
----------------------------------------------------------------
字段      Quartz（秒字段必填）        Spring（秒字段必填）
秒        0-59                        0-59
分        0-59                        0-59
时        0-23                        0-23
日        1-31，可用 ? L W            1-31，可用 ? L W
月        1-12 或 JAN-DEC             1-12 或 JAN-DEC
周        1-7 或 SUN-SAT，1=周日      0-7 或 SUN-SAT，0/7=周日
年        可选 1970-2099              不支持
段数      6 段，或 7 段（末尾加年）   固定 6 段

两个规范都支持：*   ,   -   /   以及英文月份/星期缩写。

Quartz 规则
----------------------------------------------------------------
1. 日（DOM）和周（DOW）必须有一列写 ?，不能两列同时指定具体值。
   例：0 15 10 15 * ?      每月 15 号 10:15
       0 15 10 ? * MON-FRI 周一到周五 10:15
2. 星期编号 1=周日、2=周一 … 7=周六。
   所以“周一到周五”写 MON-FRI（等价 2-6），不是 1-5。
3. L：日字段单独写 = 当月最后一天；
   周字段单独写 = 7 / SAT（注意：不是“最后一个周六”）；
   6L = 当月最后一个周五。
4. L-n：倒数第 n 天，L-2 = 倒数第 2 天。
5. W：最接近指定日期的工作日，15W 表示最接近 15 号的工作日；
   1W 遇到周六会顺延到下周一，不跨月；W 只能用于单个日期。
6. LW：当月最后一个工作日。
7. #：6#3 = 当月第 3 个周五；该月没有第 5 个时整月不触发。
8. 年字段可选：2010、2020-2030、2025/2（从 2025 起每 2 年）。
9. L / W / # 只能单独使用，不能和列表、范围混写。
   例：0 0 12 L * ?        每月最后一天 12:00
       0 0 12 L-2 * ?      每月倒数第 2 天 12:00
       0 0 12 ? * 6L       每月最后一个周五 12:00
       0 0 12 ? * 6#3      每月第 3 个周五 12:00
       0 0 9 1W * ?        每月最接近 1 号的工作日 09:00

Spring 规则
----------------------------------------------------------------
1. 固定 6 段：秒 分 时 日 月 周；没有年字段，写 7 段会报错。
2. 星期编号 0/7=周日、1=周一 … 6=周六；MON-FRI 这类缩写同样可用。
3. 支持 ?、L、L-n、LW、nW、dL、d#n（Spring 5.3 起）。
   例：0 0 12 L * ?        每月最后一天 12:00
       0 0 12 L-2 * ?      每月倒数第 2 天 12:00
       0 0 12 ? * 5L       每月最后一个周五 12:00
       0 0 12 ? * 5#3      每月第 3 个周五 12:00
4. 日和周可以同时限定，但 Spring 要求两个条件都满足（“与”语义）；
   ? 等同 *，所以只限定日字段时周字段写 ? 或 * 都可以。
   例：0 0 12 15 * MON  只在“15 号且是周一”时 12:00 触发。
   注意这与标准 cron 的“或”语义相反。
5. 支持宏：
     @yearly / @annually  → 0 0 0 1 1 *
     @monthly             → 0 0 0 1 * *
     @weekly              → 0 0 0 * * 0
     @daily / @midnight   → 0 0 0 * * *
     @hourly              → 0 0 * * * *

常见坑
----------------------------------------------------------------
· Quartz 的 1 是周日不是周一；写 MON-FRI 最不容易错。
· Quartz 日/周两列都写具体值会直接报错，必须有一列写 ?。
· Spring 日/周同时限定是“与”，不是标准 cron 的“或”；需要“或”请切标准 cron。
· Spring 没有年字段；需要“只在某几年触发”请选 Quartz。
· L / W / # 不能和列表或范围混用，例如 1,15,L 是非法的。
· 标准 5 段模式不支持 L / W / #；需要这些规则请切到 Quartz 或 Spring。
· 工具按本机时区计算触发时间；服务器时区不同时结果可能不一致。
· 夏令时切换的跳变小时会被跳过，实际触发以服务器为准。

本工具的操作方式
----------------------------------------------------------------
· 上方选择规范并输入表达式；手工输入的表达式不会被下方生成器改动。
· 下方生成器：先选字段页签，再选“每 / 周期 / 间隔 / 指定”；
  日、周字段多一个“特殊”，用于 L / W / # 规则（Quartz / Spring 才显示）。
· 生成器改动会立即写回上方表达式并重新解析。
· Quartz 下日、周只有一列能具体指定，生成器会自动把另一列写成 ?。
""";
    }
}