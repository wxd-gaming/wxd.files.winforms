using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 时间戳页：当前时间三种单位；秒/毫秒/纳秒分栏换算；年月日时分秒全部可选的时间转时间戳。
    /// 顶部当前时间每秒刷新，只在页签可见时跑计时器。
    /// </summary>
    internal sealed class TimeToolPanel : ToolPanelBase
    {
        private const int HeaderRowH = 24;
        private const int DataRowH = 32;
        private const int LiveSectionH = 64;
        private const int TimestampSectionH = 132;
        private const int DateTimeSectionH = 100;

        private readonly ThemeField _secInput = MakeInput("如 1700000000");
        private readonly ThemeField _msInput = MakeInput("如 1700000000123");
        private readonly ThemeField _nsInput = MakeInput("如 1700000000123456789");
        private readonly ThemeField _secResult = MakeReadOnlyField();
        private readonly ThemeField _msResult = MakeReadOnlyField();
        private readonly ThemeField _nsResult = MakeReadOnlyField();

        private readonly ThemeComboBox _year = MakeCombo(90);
        private readonly ThemeComboBox _month = MakeCombo(60);
        private readonly ThemeComboBox _day = MakeCombo(60);
        private readonly ThemeComboBox _hour = MakeTimeCombo();
        private readonly ThemeComboBox _minute = MakeTimeCombo();
        private readonly ThemeComboBox _second = MakeTimeCombo();
        private readonly ThemeComboBox _dateUnit = MakeCombo(90);
        private readonly ThemeField _dateResult = MakeReadOnlyField();
        private readonly ThemeButton _dateConvert;

        private readonly Label _lblNow = new();
        private readonly ThemeField _liveSec = MakeLiveField(120);
        private readonly ThemeField _liveMs = MakeLiveField(150);
        private readonly ThemeField _liveNs = MakeLiveField(200);

        private readonly TextBox _output = MakeEditor(readOnly: true);
        private readonly ThemeBordered _border;
        private readonly TableLayoutPanel _content;
        private readonly System.Windows.Forms.Timer _clock = new();
        private bool _dateUpdating;

        internal TimeToolPanel()
        {
            _dateConvert = MakeContentButton("转换", (_, _) => ConvertDate(reportErrors: true));
            _border = new ThemeBordered(_output) { Dock = DockStyle.Fill, Margin = new Padding(3, 0, 3, 0) };

            AddButton("填入当前", (_, _) => FillNow(), primary: true);
            AddButton("复制详细结果", (_, _) => CopyToClipboard(_output.Text));

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Theme.P.Window,
                ColumnCount = 1,
                RowCount = 4,
                Margin = new Padding(0),
                Padding = new Padding(ThemeMetrics.Margin, ThemeMetrics.RowPad, ThemeMetrics.Margin, ThemeMetrics.RowPad)
            };
            _content = root;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, LiveSectionH));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, TimestampSectionH));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, DateTimeSectionH));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.Controls.Add(BuildNowSection(), 0, 0);
            root.Controls.Add(BuildTimestampSection(), 0, 1);
            root.Controls.Add(BuildDateTimeSection(), 0, 2);
            root.Controls.Add(BuildReportSection(), 0, 3);
            FinishChrome(root);

            InitDateInputs();

            _clock.Interval = 1000;
            _clock.Tick += (_, _) => UpdateNowLabel();
            VisibleChanged += (_, _) => UpdateClock();
            Disposed += (_, _) => _clock.Dispose();

            _output.Text = "上方按单位分栏换算时间戳；下方可选择年月日和时间，再转成秒 / 毫秒 / 纳秒时间戳。"
                         + Environment.NewLine + "当前时间每秒刷新；纳秒按 .NET tick 精度显示到 100ns。";
            SetStatus("秒 / 毫秒 / 纳秒分栏换算；具体时间可选择年月日。");
            UpdateNowLabel();
        }

        // ==================== 布局 ====================

        private Control BuildNowSection()
        {
            var section = NewGrid(1, 2);
            section.RowStyles.Add(new RowStyle(SizeType.Absolute, HeaderRowH));
            section.RowStyles.Add(new RowStyle(SizeType.Absolute, DataRowH));

            _lblNow.AutoSize = false;
            _lblNow.Dock = DockStyle.Fill;
            _lblNow.TextAlign = ContentAlignment.MiddleLeft;
            _lblNow.Font = Theme.Mono;
            _lblNow.ForeColor = Theme.P.Muted;
            _lblNow.BackColor = Theme.P.Window;
            _lblNow.Margin = new Padding(3, 0, 3, 0);
            _lblNow.Tag = "now";
            section.Controls.Add(_lblNow, 0, 0);

            var units = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoScroll = true,
                BackColor = Theme.P.Window,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            units.Controls.Add(MakeLiveLabel("秒"));
            units.Controls.Add(_liveSec);
            units.Controls.Add(MakeLiveLabel("毫秒"));
            units.Controls.Add(_liveMs);
            units.Controls.Add(MakeLiveLabel("纳秒"));
            units.Controls.Add(_liveNs);
            section.Controls.Add(units, 0, 1);
            return section;
        }

        private Control BuildTimestampSection()
        {
            var table = NewGrid(4, 4);
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, HeaderRowH));
            for (int i = 0; i < 3; i++)
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, DataRowH));

            var header = MakeLabel("时间戳 → 时间", "header");
            table.Controls.Add(header, 0, 0);
            table.SetColumnSpan(header, 4);

            AddTimestampRow(table, 1, "秒", _secInput, _secResult, TimeUnit.Seconds);
            AddTimestampRow(table, 2, "毫秒", _msInput, _msResult, TimeUnit.Milliseconds);
            AddTimestampRow(table, 3, "纳秒", _nsInput, _nsResult, TimeUnit.Nanoseconds);
            return table;
        }

        private void AddTimestampRow(TableLayoutPanel table, int row, string label,
            ThemeField input, ThemeField result, TimeUnit unit)
        {
            table.Controls.Add(MakeLabel(label, "body"), 0, row);

            input.Dock = DockStyle.Fill;
            input.Margin = new Padding(3);
            table.Controls.Add(input, 1, row);

            var button = MakeContentButton("转换", (_, _) => ConvertTimestamp(unit, input, result));
            table.Controls.Add(button, 2, row);

            result.Dock = DockStyle.Fill;
            result.Margin = new Padding(3);
            table.Controls.Add(result, 3, row);
        }

        private Control BuildDateTimeSection()
        {
            var section = NewGrid(1, 3);
            section.RowStyles.Add(new RowStyle(SizeType.Absolute, HeaderRowH));
            section.RowStyles.Add(new RowStyle(SizeType.Absolute, DataRowH));
            section.RowStyles.Add(new RowStyle(SizeType.Absolute, DataRowH));

            var header = MakeLabel("具体时间 → 时间戳", "header");
            section.Controls.Add(header, 0, 0);

            var dateRow = NewDateTimeGrid();
            dateRow.Controls.Add(MakeLabel("时间", "body"), 0, 0);
            dateRow.Controls.Add(_year, 1, 0);
            dateRow.Controls.Add(MakeTinyLabel("年"), 2, 0);
            dateRow.Controls.Add(_month, 3, 0);
            dateRow.Controls.Add(MakeTinyLabel("月"), 4, 0);
            dateRow.Controls.Add(_day, 5, 0);
            dateRow.Controls.Add(MakeTinyLabel("日"), 6, 0);
            dateRow.Controls.Add(_hour, 7, 0);
            dateRow.Controls.Add(MakeTinyLabel("时"), 8, 0);
            dateRow.Controls.Add(_minute, 9, 0);
            dateRow.Controls.Add(MakeTinyLabel("分"), 10, 0);
            dateRow.Controls.Add(_second, 11, 0);
            dateRow.Controls.Add(MakeTinyLabel("秒"), 12, 0);
            section.Controls.Add(dateRow, 0, 1);

            var outputRow = NewDateTimeGrid();
            outputRow.Controls.Add(MakeLabel("输出", "body"), 0, 0);
            outputRow.Controls.Add(_dateUnit, 1, 0);
            outputRow.Controls.Add(_dateConvert, 3, 0);
            outputRow.SetColumnSpan(_dateConvert, 2);
            _dateResult.Dock = DockStyle.Fill;
            _dateResult.Margin = new Padding(3);
            outputRow.Controls.Add(_dateResult, 5, 0);
            outputRow.SetColumnSpan(_dateResult, 9);
            section.Controls.Add(outputRow, 0, 2);
            return section;
        }

        private static TableLayoutPanel NewDateTimeGrid()
        {
            var grid = NewGrid(14, 1);
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, DataRowH));
            int[] widths = { 40, 90, 16, 60, 16, 60, 16, 60, 16, 60, 16, 60, 16 };
            foreach (int width in widths)
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, width));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            return grid;
        }

        private Control BuildReportSection()
        {
            var section = NewGrid(1, 2);
            section.RowStyles.Add(new RowStyle(SizeType.Absolute, HeaderRowH));
            section.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            section.Controls.Add(MakeLabel("详细结果（UTC / 本地 / 秒 · 毫秒 · 纳秒）", "header"), 0, 0);
            section.Controls.Add(_border, 0, 1);
            return section;
        }

        // ==================== 初始化 ====================

        private void InitDateInputs()
        {
            for (int year = 1900; year <= 2100; year++)
                _year.Items.Add(year);
            for (int month = 1; month <= 12; month++)
                _month.Items.Add(month.ToString("D2", CultureInfo.InvariantCulture));
            for (int hour = 0; hour <= 23; hour++)
                _hour.Items.Add(hour.ToString("D2", CultureInfo.InvariantCulture));
            for (int value = 0; value <= 59; value++)
            {
                string text = value.ToString("D2", CultureInfo.InvariantCulture);
                _minute.Items.Add(text);
                _second.Items.Add(text);
            }
            _dateUnit.Items.AddRange(new object[] { "秒", "毫秒", "纳秒" });
            _dateUnit.SelectedIndex = 0;
            _year.DropDownHeight = 260;
            _hour.DropDownHeight = _minute.DropDownHeight = _second.DropDownHeight = 240;

            FillDateSelectors(DateTimeOffset.Now);

            _year.SelectedIndexChanged += (_, _) =>
            {
                UpdateDayOptions();
                ConvertDate(reportErrors: false);
            };
            _month.SelectedIndexChanged += (_, _) =>
            {
                UpdateDayOptions();
                ConvertDate(reportErrors: false);
            };
            _day.SelectedIndexChanged += (_, _) => ConvertDate(reportErrors: false);
            _hour.SelectedIndexChanged += (_, _) => ConvertDate(reportErrors: false);
            _minute.SelectedIndexChanged += (_, _) => ConvertDate(reportErrors: false);
            _second.SelectedIndexChanged += (_, _) => ConvertDate(reportErrors: false);
            _dateUnit.SelectedIndexChanged += (_, _) => ConvertDate(reportErrors: false);

            WireTimestamp(_secInput, _secResult, TimeUnit.Seconds);
            WireTimestamp(_msInput, _msResult, TimeUnit.Milliseconds);
            WireTimestamp(_nsInput, _nsResult, TimeUnit.Nanoseconds);
        }

        private void WireTimestamp(ThemeField input, ThemeField result, TimeUnit unit)
        {
            input.Edit.TextChanged += (_, _) =>
            {
                input.Edit.ForeColor = Theme.P.Text;
                result.Edit.ForeColor = Theme.P.Text;
                result.Edit.Text = "";
            };
            input.Edit.KeyDown += (_, e) =>
            {
                if (e.KeyCode != Keys.Enter) return;
                e.Handled = true;
                e.SuppressKeyPress = true;
                ConvertTimestamp(unit, input, result);
            };
        }

        private void FillDateSelectors(DateTimeOffset now)
        {
            bool previous = _dateUpdating;
            _dateUpdating = true;
            try
            {
                int year = Math.Clamp(now.Year, 1900, 2100);
                int index = _year.Items.IndexOf(year);
                _year.SelectedIndex = index >= 0 ? index : 0;
                _month.SelectedIndex = now.Month - 1;
                UpdateDayOptions();
                _day.SelectedIndex = Math.Clamp(now.Day, 1, _day.Items.Count) - 1;
                _hour.SelectedIndex = now.Hour;
                _minute.SelectedIndex = now.Minute;
                _second.SelectedIndex = now.Second;
            }
            finally
            {
                _dateUpdating = previous;
            }
        }

        private void UpdateDayOptions()
        {
            int year = SelectedYear();
            int month = Math.Clamp(_month.SelectedIndex + 1, 1, 12);
            int maxDay = DateTime.DaysInMonth(year, month);
            int selected = _day.SelectedIndex >= 0 ? _day.SelectedIndex + 1 : DateTime.Now.Day;

            bool previous = _dateUpdating;
            _dateUpdating = true;
            try
            {
                _day.BeginUpdate();
                _day.Items.Clear();
                for (int day = 1; day <= maxDay; day++)
                    _day.Items.Add(day.ToString("D2", CultureInfo.InvariantCulture));
                _day.SelectedIndex = Math.Clamp(selected, 1, maxDay) - 1;
                _day.EndUpdate();
            }
            finally
            {
                _dateUpdating = previous;
            }
        }

        // ==================== 主题 ====================

        protected override void ApplyThemeCore()
        {
            _content.BackColor = Theme.P.Window;
            _border.ApplyTheme();
            ApplyEditorTheme(_output);
            StyleTree(_content);
            UpdateNowLabel();
        }

        private static void StyleTree(Control parent)
        {
            foreach (Control child in parent.Controls)
            {
                switch (child)
                {
                    case ThemeBordered:
                        break;
                    case ThemeField field:
                        field.ApplyTheme();
                        break;
                    case ThemeComboBox combo:
                        combo.Font = Theme.Body;
                        combo.BackColor = Theme.P.Surface;
                        combo.ForeColor = Theme.P.Text;
                        break;
                    case ThemeButton button:
                        button.Font = Theme.Body;
                        button.BackColor = Theme.P.Window;
                        button.Size = new Size(ThemeButton.FitWidth(button.Text), ThemeMetrics.RowH);
                        button.Invalidate();
                        break;
                    case Label label:
                        switch (label.Tag as string)
                        {
                            case "header":
                                label.Font = Theme.BodyBold;
                                label.ForeColor = Theme.P.Text;
                                break;
                            case "muted":
                                label.Font = Theme.Body;
                                label.ForeColor = Theme.P.Muted;
                                break;
                            case "now":
                                label.Font = Theme.Mono;
                                label.ForeColor = Theme.P.Muted;
                                break;
                            default:
                                label.Font = Theme.Body;
                                label.ForeColor = Theme.P.Text;
                                break;
                        }
                        label.BackColor = Theme.P.Window;
                        break;
                    case Panel panel:
                        panel.BackColor = Theme.P.Window;
                        break;
                }
                StyleTree(child);
            }
        }

        // ==================== 当前时间 ====================

        private void UpdateClock()
        {
            if (Visible && !_clock.Enabled) _clock.Start();
            else if (!Visible && _clock.Enabled) _clock.Stop();
            if (Visible) UpdateNowLabel();
        }

        private void UpdateNowLabel()
        {
            var now = DateTimeOffset.Now;
            _lblNow.Text = "当前本地时间  "
                + now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                + "  (UTC" + FormatOffset(now.Offset) + ")";

            _liveSec.Edit.Text = now.ToUnixTimeSeconds().ToString(CultureInfo.InvariantCulture);
            _liveMs.Edit.Text = now.ToUnixTimeMilliseconds().ToString(CultureInfo.InvariantCulture);
            _liveNs.Edit.Text = TimeToolEngine.TryToUnixTime(now, TimeUnit.Nanoseconds, out long ns, out _)
                ? ns.ToString(CultureInfo.InvariantCulture)
                : "超出范围";
        }

        private void FillNow()
        {
            var now = DateTimeOffset.Now;
            _secInput.Edit.Text = now.ToUnixTimeSeconds().ToString(CultureInfo.InvariantCulture);
            _msInput.Edit.Text = now.ToUnixTimeMilliseconds().ToString(CultureInfo.InvariantCulture);
            _nsInput.Edit.Text = TimeToolEngine.TryToUnixTime(now, TimeUnit.Nanoseconds, out long ns, out _)
                ? ns.ToString(CultureInfo.InvariantCulture)
                : "";

            FillDateSelectors(now);
            _dateResult.Edit.Text = "";
            _dateResult.Edit.ForeColor = Theme.P.Text;
            _output.Text = "已填入当前时间。时间戳各栏和下方的时间可以分别点击「转换」。";
            UpdateNowLabel();
            SetStatus("已填入当前时间；秒 / 毫秒 / 纳秒和年月日时间都可单独转换。");
        }

        // ==================== 换算 ====================

        private void ConvertTimestamp(TimeUnit unit, ThemeField input, ThemeField result)
        {
            var converted = TimeToolEngine.Convert(input.Edit.Text, unit);
            if (!converted.Success)
            {
                input.Edit.ForeColor = Theme.P.Danger;
                result.Edit.Text = "转换失败";
                result.Edit.ForeColor = Theme.P.Danger;
                SetStatus(converted.Error ?? "时间戳转换失败", error: true);
                return;
            }

            input.Edit.ForeColor = Theme.P.Text;
            result.Edit.ForeColor = Theme.P.Text;
            var local = converted.Value.ToLocalTime();
            string format = unit == TimeUnit.Nanoseconds
                ? "yyyy-MM-dd HH:mm:ss.fffffff"
                : "yyyy-MM-dd HH:mm:ss.fff";
            result.Edit.Text = local.ToString(format, CultureInfo.InvariantCulture)
                             + "  (UTC" + FormatOffset(local.Offset) + ")";
            _output.Text = converted.Report;
            SetStatus("已换算 · " + converted.InputKind);
        }

        private void ConvertDate(bool reportErrors)
        {
            if (_dateUpdating) return;

            if (!TryGetDateTime(out DateTime local, out string error))
            {
                _dateResult.Edit.Text = "";
                _dateResult.Edit.ForeColor = Theme.P.Text;
                if (reportErrors) SetStatus(error, error: true);
                return;
            }

            var value = new DateTimeOffset(local);
            TimeUnit unit = SelectedDateUnit();
            if (!TimeToolEngine.TryToUnixTime(value, unit, out long timestamp, out string? unitError))
            {
                _dateResult.Edit.Text = "超出范围";
                _dateResult.Edit.ForeColor = Theme.P.Danger;
                if (reportErrors) SetStatus(unitError ?? "时间戳超出范围", error: true);
                return;
            }

            _dateResult.Edit.ForeColor = Theme.P.Text;
            _dateResult.Edit.Text = timestamp.ToString(CultureInfo.InvariantCulture);
            _output.Text = TimeToolEngine.Describe(value, "指定时间（本地）").Report;
            SetStatus("已换算 · 指定时间 → " + TimeToolEngine.UnitName(unit) + "时间戳 " + timestamp);
        }

        private bool TryGetDateTime(out DateTime value, out string error)
        {
            value = default;
            if (_year.SelectedItem is not int year)
            {
                error = "请选择年份";
                return false;
            }

            int month = _month.SelectedIndex + 1;
            int day = _day.SelectedIndex + 1;
            if (month < 1 || month > 12 || day < 1 || day > DateTime.DaysInMonth(year, month))
            {
                error = "请选择有效的年月日";
                return false;
            }

            if (_hour.SelectedIndex < 0 || _minute.SelectedIndex < 0 || _second.SelectedIndex < 0)
            {
                error = "请选择时、分、秒";
                return false;
            }

            value = new DateTime(year, month, day,
                _hour.SelectedIndex, _minute.SelectedIndex, _second.SelectedIndex, DateTimeKind.Local);
            error = "";
            return true;
        }

        private int SelectedYear() => _year.SelectedItem is int year ? year : DateTime.Now.Year;

        private TimeUnit SelectedDateUnit() => _dateUnit.SelectedIndex switch
        {
            1 => TimeUnit.Milliseconds,
            2 => TimeUnit.Nanoseconds,
            _ => TimeUnit.Seconds
        };

        private static string FormatOffset(TimeSpan offset) =>
            (offset < TimeSpan.Zero ? "-" : "+")
            + Math.Abs(offset.Hours).ToString("D2", CultureInfo.InvariantCulture)
            + ":" + Math.Abs(offset.Minutes).ToString("D2", CultureInfo.InvariantCulture);

        // ==================== 控件工厂 ====================

        private static TableLayoutPanel NewGrid(int columns, int rows)
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Theme.P.Window,
                ColumnCount = columns,
                RowCount = rows,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            return grid;
        }

        private static ThemeField MakeInput(string placeholder)
        {
            return new ThemeField
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(3),
                Placeholder = placeholder
            };
        }

        private static ThemeField MakeReadOnlyField()
        {
            var field = MakeInput("");
            field.Edit.ReadOnly = true;
            field.Edit.TabStop = false;
            return field;
        }

        private static ThemeComboBox MakeTimeCombo()
        {
            var combo = MakeCombo(60);
            combo.DropDownHeight = 240;
            return combo;
        }

        private static ThemeField MakeLiveField(int width)
        {
            var field = new ThemeField
            {
                Size = new Size(width, ThemeMetrics.RowH),
                Margin = new Padding(0, 3, 0, 3)
            };
            field.Edit.ReadOnly = true;
            field.Edit.TabStop = false;
            field.Edit.TextAlign = HorizontalAlignment.Center;
            return field;
        }

        private static ThemeComboBox MakeCombo(int width)
        {
            return new ThemeComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = Theme.Body,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
                Size = new Size(width, ThemeMetrics.RowH),
                Dock = DockStyle.Fill,
                Margin = new Padding(3)
            };
        }

        private static ThemeButton MakeContentButton(string text, EventHandler onClick)
        {
            var button = new ThemeButton
            {
                Text = text,
                Font = Theme.Body,
                Size = new Size(ThemeButton.FitWidth(text), ThemeMetrics.RowH),
                Margin = new Padding(3),
                Anchor = AnchorStyles.Left,
                BackColor = Theme.P.Window
            };
            button.Click += onClick;
            return button;
        }

        private static Label MakeLabel(string text, string tag)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Tag = tag,
                Font = tag == "header" ? Theme.BodyBold : Theme.Body,
                ForeColor = tag == "muted" ? Theme.P.Muted : Theme.P.Text,
                BackColor = Theme.P.Window,
                Margin = new Padding(3, tag == "header" ? 3 : 7, 3, 3),
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static Label MakeTinyLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Tag = "muted",
                Font = Theme.Body,
                ForeColor = Theme.P.Muted,
                BackColor = Theme.P.Window,
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0)
            };
        }

        private static Label MakeLiveLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Tag = "muted",
                Font = Theme.Body,
                ForeColor = Theme.P.Muted,
                BackColor = Theme.P.Window,
                Margin = new Padding(12, 7, 4, 3),
                TextAlign = ContentAlignment.MiddleLeft
            };
        }
    }
}