using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// cron 可视化生成器：按字段页签选择“每 / 周期 / 间隔 / 指定”，日、周另有 L / W / # 规则。
    /// 生成结果通过 <see cref="ExpressionChanged"/> 交给调用方填回表达式输入框；
    /// 切换规范只改变字段布局和编号规则，不会覆盖已经写好的表达式。
    /// </summary>
    internal sealed class CronBuilderControl : UserControl
    {
        private const int FieldSecond = 0;
        private const int FieldMinute = 1;
        private const int FieldHour = 2;
        private const int FieldDom = 3;
        private const int FieldMonth = 4;
        private const int FieldDow = 5;
        private const int FieldYear = 6;

        private const int ModeAny = 0;
        private const int ModeRange = 1;
        private const int ModeStep = 2;
        private const int ModeList = 3;
        private const int ModeSpecial = 4;

        private static readonly string[] FieldNames = { "秒", "分", "时", "日", "月", "周", "年" };
        private static readonly string[] DayNames = { "SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT" };
        private static readonly string[] DayLabels = { "周日", "周一", "周二", "周三", "周四", "周五", "周六" };

        /// <summary>周字段按“周一在前”展示，值仍用内部编号 0=周日。</summary>
        private static readonly int[] WeekOrder = { 1, 2, 3, 4, 5, 6, 0 };

        private readonly ThemeSegmented _tabs = new();
        private readonly FlowLayoutPanel _head = new();
        private readonly FlowLayoutPanel _modesRow = new();
        private readonly FlowLayoutPanel _options = new();
        private readonly Dictionary<int, FlowLayoutPanel> _fieldGrids = new();
        private readonly Dictionary<int, List<ThemeCheckBox>> _fieldChecks = new();
        private readonly Panel _gridHost = new();
        private readonly Label _gridHint = new();
        private readonly Label _preview = new();

        private readonly ThemeRadioButton[] _modeRadios = new ThemeRadioButton[5];
        private readonly int[] _modes = new int[7];
        private readonly HashSet<int>[] _selected = new HashSet<int>[7];
        private readonly int[] _rangeStart = new int[7];
        private readonly int[] _rangeEnd = new int[7];
        private readonly int[] _stepStart = new int[7];
        private readonly int[] _stepStep = new int[7];

        private int _domSpecialKind;
        private int _domSpecialN = 1;
        private int _dowSpecialKind;
        private int _dowSpecialN = 1;
        private int _specialWeekday = 1;
        private int _activeField = FieldSecond;
        private bool _updating;

        public event EventHandler? ExpressionChanged;

        public CronDialect Dialect { get; private set; } = CronDialect.Quartz;

        public string Expression => BuildExpression();

        public CronBuilderControl()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            BackColor = Theme.P.Window;
            Font = Theme.Body;
            Padding = new Padding(ThemeMetrics.Margin, 4, ThemeMetrics.Margin, 2);

            InitState();
            BuildLayout();
            RebuildTabs();
            RefreshFromState();
        }

        // ==================== 初始化与布局 ====================

        private void InitState()
        {
            for (int i = 0; i < 7; i++)
            {
                _selected[i] = new HashSet<int>();
                _modes[i] = ModeList;
                _rangeStart[i] = MinOf(i);
                _rangeEnd[i] = MaxOf(i);
                _stepStart[i] = MinOf(i);
                _stepStep[i] = 1;
            }

            // 默认每天 0 点：秒/分/时为指定 0，其余字段空选即不限
            _selected[FieldSecond].Add(0);
            _selected[FieldMinute].Add(0);
            _selected[FieldHour].Add(0);
            _rangeStart[FieldSecond] = 0;
            _rangeEnd[FieldSecond] = 59;
            _rangeStart[FieldMinute] = 0;
            _rangeEnd[FieldMinute] = 59;
            _rangeStart[FieldHour] = 0;
            _rangeEnd[FieldHour] = 23;

            _rangeStart[FieldDom] = 1;
            _rangeEnd[FieldDom] = 31;
            _rangeStart[FieldMonth] = 1;
            _rangeEnd[FieldMonth] = 12;
            _rangeStart[FieldDow] = 1;
            _rangeEnd[FieldDow] = 5;
            _rangeStart[FieldYear] = 1970;
            _rangeEnd[FieldYear] = 2099;
            _stepStep[FieldSecond] = 5;
            _stepStep[FieldMinute] = 5;
            _stepStep[FieldHour] = 1;
            _stepStep[FieldDom] = 2;
            _stepStep[FieldMonth] = 1;
            _stepStep[FieldDow] = 1;
            _stepStep[FieldYear] = 1;
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                BackColor = Theme.P.Window,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 33));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));

            _head.Dock = DockStyle.Fill;
            _head.FlowDirection = FlowDirection.LeftToRight;
            _head.WrapContents = false;
            _head.BackColor = Theme.P.Window;
            _head.Margin = new Padding(0);
            _head.Padding = new Padding(0, 5, 0, 0);
            var fieldLabel = new Label
            {
                Text = "字段",
                AutoSize = true,
                Font = Theme.Body,
                ForeColor = Theme.P.Muted,
                BackColor = Theme.P.Window,
                Margin = new Padding(0, 9, 10, 0)
            };
            _tabs.Margin = new Padding(0);
            _tabs.SegmentClicked += OnTabClicked;
            _head.Controls.Add(fieldLabel);
            _head.Controls.Add(_tabs);

            _modesRow.Dock = DockStyle.Fill;
            _modesRow.FlowDirection = FlowDirection.LeftToRight;
            _modesRow.WrapContents = false;
            _modesRow.BackColor = Theme.P.Window;
            _modesRow.Margin = new Padding(0);
            _modesRow.Padding = new Padding(0, 2, 0, 0);
            var modeLabel = new Label
            {
                Text = "模式",
                AutoSize = true,
                Font = Theme.Body,
                ForeColor = Theme.P.Muted,
                BackColor = Theme.P.Window,
                Margin = new Padding(0, 7, 10, 0)
            };
            _modesRow.Controls.Add(modeLabel);

            // “每 (*)”和“指定”合并：网格常显，空选就是 *；这里只保留周期/间隔/特殊三种切换
            (int Mode, string Text)[] modeOptions =
            {
                (ModeList, "指定"),
                (ModeRange, "周期"),
                (ModeStep, "间隔"),
                (ModeSpecial, "特殊")
            };
            foreach (var (mode, text) in modeOptions)
            {
                var radio = new ThemeRadioButton
                {
                    Text = text,
                    AutoSize = true,
                    Font = Theme.Body,
                    ForeColor = Theme.P.Text,
                    BackColor = Theme.P.Window,
                    Margin = new Padding(0, 4, 14, 0)
                };
                int captured = mode;
                radio.Click += (_, _) => SetMode(captured);
                _modeRadios[mode] = radio;
                _modesRow.Controls.Add(radio);
            }

            _options.Dock = DockStyle.Fill;
            _options.FlowDirection = FlowDirection.LeftToRight;
            _options.WrapContents = false;
            _options.AutoScroll = true;
            _options.BackColor = Theme.P.Window;
            _options.Margin = new Padding(0);
            _options.Padding = new Padding(0, 5, 0, 0);

            _gridHost.Dock = DockStyle.Fill;
            _gridHost.BackColor = Theme.P.Window;
            _gridHost.Margin = new Padding(0);

            _gridHint.Dock = DockStyle.Fill;
            _gridHint.AutoSize = false;
            _gridHint.TextAlign = ContentAlignment.MiddleLeft;
            _gridHint.Font = Theme.Body;
            _gridHint.ForeColor = Theme.P.Muted;
            _gridHint.BackColor = Theme.P.Window;
            _gridHint.Padding = new Padding(4, 0, 0, 0);

            _gridHost.Controls.Add(_gridHint);

            _preview.Dock = DockStyle.Fill;
            _preview.AutoSize = false;
            _preview.TextAlign = ContentAlignment.MiddleLeft;
            _preview.Font = Theme.Caption;
            _preview.ForeColor = Theme.P.Muted;
            _preview.BackColor = Theme.P.Window;
            _preview.AutoEllipsis = true;
            _preview.Margin = new Padding(0);

            root.Controls.Add(_head, 0, 0);
            root.Controls.Add(_modesRow, 0, 1);
            root.Controls.Add(_options, 0, 2);
            root.Controls.Add(_gridHost, 0, 3);
            root.Controls.Add(_preview, 0, 4);
            Controls.Add(root);
        }

        // ==================== 规范切换 ====================

        public void SetDialect(CronDialect dialect)
        {
            if (dialect == Dialect) return;
            Dialect = dialect;
            if (dialect == CronDialect.Quartz && BothDayFieldsSpecific())
                MakeUnrestricted(FieldDow);
            if (dialect == CronDialect.Standard)
            {
                // 标准 cron 没有 L / W / #，藏在不可见的“特殊”模式里会继续生成非法表达式
                if (_modes[FieldDom] == ModeSpecial) MakeUnrestricted(FieldDom);
                if (_modes[FieldDow] == ModeSpecial) MakeUnrestricted(FieldDow);
            }

            var visible = VisibleFields();
            if (!visible.Contains(_activeField)) _activeField = visible[0];
            RebuildTabs();
            RefreshFromState();
        }

        private bool BothDayFieldsSpecific() =>
            !IsUnrestricted(FieldDom) && !IsUnrestricted(FieldDow);

        /// <summary>示例按钮用：把表达式解析结果灌回生成器，让下方页签、模式和勾选项跟着示例走。</summary>
        public bool LoadExpression(string expression, CronDialect dialect)
        {
            var schedule = CronToolEngine.Parse(expression, dialect);
            if (!schedule.Success || schedule.Tokens.Length != 7) return false;

            // 示例按当前规范筛选，这里直接换规范再整体覆盖状态，不走 SetDialect 的旧状态修补
            Dialect = dialect;
            int[][] values =
            {
                schedule.Seconds,
                schedule.Minutes,
                schedule.Hours,
                schedule.DaysOfMonth,
                schedule.Months,
                schedule.DaysOfWeek,
                schedule.Years
            };

            _updating = true;
            try
            {
                for (int field = 0; field < 7; field++)
                {
                    _selected[field].Clear();
                    ApplyToken(field, schedule.Tokens[field], values[field], schedule);
                }
            }
            finally
            {
                _updating = false;
            }

            _activeField = ChooseDemoField(VisibleFields());
            RebuildTabs();
            RefreshFromState();
            return true;
        }

        /// <summary>把单个字段的原始 token 映射回生成器能表达的模式，表达不了就退回完整的“指定”勾选。</summary>
        private void ApplyToken(int field, string token, int[] values, CronSchedule schedule)
        {
            string t = (token ?? "").Trim();
            _modes[field] = ModeList;
            _selected[field].Clear();
            if (t.Length == 0 || t == "*" || t == "?") return;

            if (field == FieldDom && ApplyDomSpecial(schedule)) return;
            if (field == FieldDow && ApplyDowSpecial(schedule)) return;

            int slash = t.IndexOf('/');
            if (slash > 0 && slash == t.LastIndexOf('/'))
            {
                string startText = t.Substring(0, slash).Trim();
                string stepText = t.Substring(slash + 1).Trim();
                bool startOk;
                int start;
                if (startText == "*")
                {
                    start = MinOf(field);
                    startOk = true;
                }
                else
                {
                    startOk = int.TryParse(startText, out start);
                }

                if (startOk && int.TryParse(stepText, out int step) && step > 0)
                {
                    _modes[field] = ModeStep;
                    _stepStart[field] = Math.Clamp(start, MinOf(field), MaxOf(field));
                    _stepStep[field] = step;
                    return;
                }
            }

            // 只有 a-b 这种连续简单范围能还原成“周期”；列表、混合写法一律退回勾选
            if (t.IndexOf('-') > 0 && t.IndexOf(',') < 0 && slash < 0 && values.Length > 1)
            {
                int start = values.Min();
                int end = values.Max();
                if (values.Length == end - start + 1)
                {
                    _modes[field] = ModeRange;
                    _rangeStart[field] = start;
                    _rangeEnd[field] = end;
                    return;
                }
            }

            foreach (int value in values) _selected[field].Add(value);
        }

        private bool ApplyDomSpecial(CronSchedule schedule)
        {
            if (schedule.DomIsLast)
            {
                _modes[FieldDom] = ModeSpecial;
                _domSpecialKind = 0;
                return true;
            }
            if (schedule.DomLastOffset >= 0)
            {
                _modes[FieldDom] = ModeSpecial;
                _domSpecialKind = 1;
                _domSpecialN = schedule.DomLastOffset;
                return true;
            }
            if (schedule.DomIsLastWeekday)
            {
                _modes[FieldDom] = ModeSpecial;
                _domSpecialKind = 2;
                return true;
            }
            if (schedule.DomNearestWeekday >= 0)
            {
                _modes[FieldDom] = ModeSpecial;
                _domSpecialKind = 3;
                _domSpecialN = schedule.DomNearestWeekday;
                return true;
            }
            return false;
        }

        private bool ApplyDowSpecial(CronSchedule schedule)
        {
            if (schedule.DowLastWeekday >= 0)
            {
                _modes[FieldDow] = ModeSpecial;
                _dowSpecialKind = 0;
                _specialWeekday = schedule.DowLastWeekday;
                return true;
            }
            if (schedule.DowNthWeekday >= 0)
            {
                _modes[FieldDow] = ModeSpecial;
                _dowSpecialKind = 1;
                _specialWeekday = schedule.DowNthWeekday;
                _dowSpecialN = schedule.DowNth;
                return true;
            }
            return false;
        }

        /// <summary>优先停在周期/间隔/特殊字段，其次是最后一个有勾选的字段，让示例的重点一眼可见。</summary>
        private int ChooseDemoField(int[] visible)
        {
            foreach (int field in visible)
                if (_modes[field] == ModeRange || _modes[field] == ModeStep || _modes[field] == ModeSpecial)
                    return field;

            int lastSelected = -1;
            foreach (int field in visible)
                if (_modes[field] == ModeList && _selected[field].Count > 0)
                    lastSelected = field;
            return lastSelected >= 0 ? lastSelected : visible[0];
        }

        private int[] VisibleFields()
        {
            if (Dialect == CronDialect.Standard)
                return new[] { FieldMinute, FieldHour, FieldDom, FieldMonth, FieldDow };
            if (Dialect == CronDialect.Quartz)
                return new[] { FieldSecond, FieldMinute, FieldHour, FieldDom, FieldMonth, FieldDow, FieldYear };
            return new[] { FieldSecond, FieldMinute, FieldHour, FieldDom, FieldMonth, FieldDow };
        }

        private void RebuildTabs()
        {
            var visible = VisibleFields();
            _updating = true;
            _tabs.Items = visible.Select(f => FieldNames[f]).ToArray();
            _tabs.ActiveIndex = Math.Max(0, Array.IndexOf(visible, _activeField));
            _updating = false;
        }

        private void OnTabClicked(int index)
        {
            var visible = VisibleFields();
            if (index < 0 || index >= visible.Length) return;
            if (_activeField == visible[index])
            {
                _tabs.ActiveIndex = index;
                return;
            }
            _activeField = visible[index];
            _tabs.ActiveIndex = index;
            RefreshFromState();
        }

        // ==================== 模式与选项 ====================

        private void SetMode(int mode)
        {
            if (_updating) return;
            _modes[_activeField] = mode;

            // Quartz 的日/周必须有一列是 ?，两个都指定时保留当前字段、另一列退回不限
            if (Dialect == CronDialect.Quartz && IsDayField(_activeField) && !IsUnrestricted(_activeField))
                MakeUnrestricted(_activeField == FieldDom ? FieldDow : FieldDom);

            RefreshModeRadios();
            RefreshOptions();
            RefreshGrid();
            UpdatePreview();
            RaiseChanged();
        }

        private void RefreshFromState()
        {
            RefreshModeRadios();
            RefreshOptions();
            RefreshGrid();
            UpdatePreview();
        }

        private void RefreshModeRadios()
        {
            _updating = true;
            bool dayField = IsDayField(_activeField);
            _modeRadios[ModeSpecial].Visible = dayField && Dialect != CronDialect.Standard;

            int shown = _modes[_activeField] == ModeAny ? ModeList : _modes[_activeField];
            for (int i = 1; i < _modeRadios.Length; i++)
                if (_modeRadios[i] != null) _modeRadios[i].Checked = i == shown;

            _modesRow.PerformLayout();
            _updating = false;
        }

        private void RefreshOptions()
        {
            _options.SuspendLayout();
            _updating = true;
            ClearControls(_options);

            switch (_modes[_activeField])
            {
                case ModeRange:
                    AddOptionLabel("从");
                    AddNumber(_rangeStart[_activeField], MinOf(_activeField), MaxOf(_activeField), v =>
                    {
                        _rangeStart[_activeField] = v;
                        UpdatePreview();
                        RaiseChanged();
                    });
                    AddOptionLabel("到");
                    AddNumber(_rangeEnd[_activeField], MinOf(_activeField), MaxOf(_activeField), v =>
                    {
                        _rangeEnd[_activeField] = v;
                        UpdatePreview();
                        RaiseChanged();
                    });
                    AddOptionLabel(IsDayField(_activeField) && _activeField == FieldDow
                        ? "（星期用英文缩写表示）"
                        : "");
                    break;

                case ModeStep:
                    AddOptionLabel("从");
                    AddNumber(_stepStart[_activeField], MinOf(_activeField), MaxOf(_activeField), v =>
                    {
                        _stepStart[_activeField] = v;
                        UpdatePreview();
                        RaiseChanged();
                    });
                    AddOptionLabel("每");
                    AddNumber(_stepStep[_activeField], 1, Math.Max(1, MaxOf(_activeField)), v =>
                    {
                        _stepStep[_activeField] = v;
                        UpdatePreview();
                        RaiseChanged();
                    });
                    AddOptionLabel(_activeField == FieldYear ? "年执行一次" : "执行一次");
                    break;

                case ModeAny:
                case ModeList:
                    AddOptionLabel(GridHintText());
                    AddGridAction("全选", true);
                    AddGridAction("取消全选", false);
                    break;

                case ModeSpecial:
                    if (_activeField == FieldDom) BuildDomSpecial();
                    else BuildDowSpecial();
                    break;

                default:
                    AddOptionLabel(Dialect == CronDialect.Quartz && IsDayField(_activeField)
                        ? "这一列不限制，Quartz 会写 ?。"
                        : "这一字段不限制。");
                    break;
            }

            _updating = false;
            _options.ResumeLayout(true);
            _options.PerformLayout();
        }

        private void BuildDomSpecial()
        {
            var kind = MakeCombo(new[]
            {
                "最后一天 (L)",
                "倒数第 n 天 (L-n)",
                "最后一个工作日 (LW)",
                "最接近 n 号的工作日 (nW)"
            }, 190);
            kind.SelectedIndex = _domSpecialKind;
            _options.Controls.Add(kind);
            var nLabel = AddOptionLabel("n =");
            var nField = AddNumber(_domSpecialN, 1, 31, v =>
            {
                _domSpecialN = v;
                UpdatePreview();
                RaiseChanged();
            });
            bool needN = _domSpecialKind == 1 || _domSpecialKind == 3;
            nLabel.Visible = needN;
            nField.Visible = needN;
            nField.Placeholder = _domSpecialKind == 3 ? "号" : "天";
            kind.SelectedIndexChanged += (_, _) =>
            {
                if (_updating) return;
                _domSpecialKind = Math.Max(0, kind.SelectedIndex);
                bool show = _domSpecialKind == 1 || _domSpecialKind == 3;
                nLabel.Visible = show;
                nField.Visible = show;
                nField.Placeholder = _domSpecialKind == 3 ? "号" : "天";
                UpdatePreview();
                RaiseChanged();
            };
        }

        private void BuildDowSpecial()
        {
            var weekday = MakeCombo(WeekOrder.Select(v => DayLabels[v]).ToArray(), 90);
            weekday.SelectedIndex = Math.Max(0, Array.IndexOf(WeekOrder, _specialWeekday));
            var kind = MakeCombo(new[] { "最后一个 (dL)", "第 n 个 (d#n)" }, 150);
            kind.SelectedIndex = _dowSpecialKind;
            _options.Controls.Add(weekday);
            _options.Controls.Add(kind);
            var nLabel = AddOptionLabel("n =");
            var nField = AddNumber(_dowSpecialN, 1, 5, v =>
            {
                _dowSpecialN = v;
                UpdatePreview();
                RaiseChanged();
            });
            bool needN = _dowSpecialKind == 1;
            nLabel.Visible = needN;
            nField.Visible = needN;

            weekday.SelectedIndexChanged += (_, _) =>
            {
                if (_updating || weekday.SelectedIndex < 0) return;
                _specialWeekday = WeekOrder[weekday.SelectedIndex];
                UpdatePreview();
                RaiseChanged();
            };
            kind.SelectedIndexChanged += (_, _) =>
            {
                if (_updating) return;
                _dowSpecialKind = Math.Max(0, kind.SelectedIndex);
                nLabel.Visible = _dowSpecialKind == 1;
                nField.Visible = _dowSpecialKind == 1;
                UpdatePreview();
                RaiseChanged();
            };
        }

        private Label AddOptionLabel(string text)
        {
            var label = new Label
            {
                Text = text,
                AutoSize = true,
                Font = Theme.Body,
                ForeColor = Theme.P.Text,
                BackColor = Theme.P.Window,
                Margin = new Padding(0, 8, 8, 0)
            };
            _options.Controls.Add(label);
            return label;
        }

        private void AddGridAction(string text, bool selectAll)
        {
            var button = new ThemeButton
            {
                Text = text,
                Font = Theme.Body,
                BackColor = Theme.P.Window,
                Size = new Size(ThemeButton.FitWidth(text), ThemeMetrics.RowH),
                Margin = new Padding(0, 2, ThemeMetrics.Gap, 0)
            };
            button.Click += (_, _) => SetAllGridValues(selectAll);
            _options.Controls.Add(button);
        }

        private string GridHintText() =>
            Dialect == CronDialect.Quartz && IsDayField(_activeField)
                ? "在下方勾选允许的取值；不勾选表示不限，日/周按 Quartz 规则使用 ?。"
                : "在下方勾选允许的取值；一个都不勾选时等同不限 (*)。";

        /// <summary>网格模式就是“指定”；空集合等同不限，表达式按 * 处理。</summary>
        private bool IsListMode(int field) =>
            _modes[field] == ModeList || _modes[field] == ModeAny;

        /// <summary>Quartz 日/周没有勾选值时视为不限，由表达式生成逻辑决定哪一侧写 ?。</summary>
        private bool IsUnrestricted(int field) =>
            _modes[field] == ModeAny
            || (_modes[field] == ModeList && _selected[field].Count == 0);

        /// <summary>把字段退回“空勾选”状态，同时清掉旧勾选，避免切回网格时又变成指定。</summary>
        private void MakeUnrestricted(int field)
        {
            _modes[field] = ModeList;
            _selected[field].Clear();
            if (!_fieldChecks.TryGetValue(field, out var checks)) return;
            bool wasUpdating = _updating;
            _updating = true;
            try
            {
                foreach (var check in checks) check.Checked = false;
            }
            finally
            {
                _updating = wasUpdating;
            }
        }

        /// <summary>Quartz 下日/周互斥：当前字段勾上值后，另一侧退回不限（输出 ?）。</summary>
        private void EnforceQuartzDayExclusive(int field)
        {
            if (Dialect != CronDialect.Quartz || !IsDayField(field) || IsUnrestricted(field)) return;
            int other = field == FieldDom ? FieldDow : FieldDom;
            if (!IsUnrestricted(other)) MakeUnrestricted(other);
        }

        /// <summary>整列勾选/清空只触发一次表达式更新，避免几十个复选框各发一遍 Changed。</summary>
        private void SetAllGridValues(bool on)
        {
            if (!IsListMode(_activeField)) return;
            if (!_fieldChecks.TryGetValue(_activeField, out var checks)) return;
            _updating = true;
            try
            {
                if (on) _selected[_activeField].UnionWith(ValuesOf(_activeField));
                else _selected[_activeField].Clear();
                foreach (var check in checks) check.Checked = on;
            }
            finally
            {
                _updating = false;
            }
            EnforceQuartzDayExclusive(_activeField);
            UpdatePreview();
            RaiseChanged();
        }

        private ThemeField AddNumber(int value, int min, int max, Action<int> apply)
        {
            var field = new ThemeField
            {
                Size = new Size(64, ThemeMetrics.RowH),
                Margin = new Padding(0, 0, 10, 0),
                Placeholder = "数字"
            };
            field.Edit.Text = value.ToString();
            field.Edit.TextChanged += (_, _) =>
            {
                if (_updating) return;
                if (!int.TryParse(field.Edit.Text.Trim(), out int parsed)) return;
                parsed = Math.Clamp(parsed, min, max);
                apply(parsed);
            };
            field.Edit.Leave += (_, _) =>
            {
                if (!int.TryParse(field.Edit.Text.Trim(), out int parsed)) parsed = value;
                parsed = Math.Clamp(parsed, min, max);
                if (field.Edit.Text != parsed.ToString()) field.Edit.Text = parsed.ToString();
                apply(parsed);
                UpdatePreview();
                RaiseChanged();
            };
            _options.Controls.Add(field);
            return field;
        }

        private static ThemeComboBox MakeCombo(string[] items, int width)
        {
            var combo = new ThemeComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = Theme.Body,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
                Size = new Size(width, ThemeMetrics.RowH),
                Margin = new Padding(0, 0, 10, 0)
            };
            combo.Items.AddRange(items);
            return combo;
        }

        // ==================== 取值网格 ====================

        private void RefreshGrid()
        {
            bool listMode = IsListMode(_activeField);
            _updating = true;
            try
            {
                foreach (var pair in _fieldChecks)
                {
                    bool on = listMode && pair.Key == _activeField;
                    foreach (var check in pair.Value)
                        check.TabStop = on;
                }

                if (!listMode)
                {
                    _gridHint.Text = _modes[_activeField] == ModeSpecial
                        ? "特殊规则由上方选项生成，不使用取值网格。"
                        : "当前模式不使用取值网格。";
                    _gridHint.Visible = true;
                    _gridHint.BringToFront();
                    return;
                }

                _gridHint.Visible = false;
                var shown = EnsureFieldGrid(_activeField);
                shown.BringToFront();
                var checks = _fieldChecks[_activeField];
                int i = 0;
                foreach (int value in ValuesOf(_activeField))
                {
                    checks[i].Checked = _selected[_activeField].Contains(value);
                    i++;
                }
                shown.AutoScrollPosition = Point.Empty;
            }
            finally
            {
                _updating = false;
            }
        }

        /// <summary>数字网格固定每行 10 个，靠 FlowBreak 强制换行，宽度再宽也不会挤成一行。</summary>
        private const int GridColumns = 10;

        /// <summary>每个字段保留自己的数字网格，切页签只改 z 序，避免几十个复选框反复改字、改宽和重建。</summary>
        private FlowLayoutPanel EnsureFieldGrid(int field)
        {
            if (_fieldGrids.TryGetValue(field, out var existing)) return existing;

            var grid = new GridFlowPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                AutoScroll = true,
                BackColor = Theme.P.Window,
                Padding = new Padding(2, 2, 2, 2)
            };
            grid.SuspendLayout();

            var checks = new List<ThemeCheckBox>();
            int index = 0;
            foreach (int value in ValuesOf(field))
            {
                int captured = value;
                var check = new ThemeCheckBox
                {
                    Text = ValueLabel(field, value),
                    AutoSize = true,
                    Font = Theme.Body,
                    ForeColor = Theme.P.Text,
                    BackColor = Theme.P.Window,
                    Margin = new Padding(2, 3, 6, 3),
                    Tag = value,
                    Checked = _selected[field].Contains(value)
                };
                check.CheckedChanged += (_, _) => OnGridCheckChanged(field, check, captured);
                checks.Add(check);
                grid.Controls.Add(check);
                if (++index % GridColumns == 0)
                    grid.SetFlowBreak(check, true);
            }

            NormalizeGridCells(checks);
            grid.ResumeLayout(true);
            _fieldChecks[field] = checks;
            _fieldGrids[field] = grid;
            _gridHost.Controls.Add(grid);
            return grid;
        }

        /// <summary>AutoSize 会按文字各排各的，1 和 30 混在一起就成锯齿；统一成该字段最宽的格子，列才对得齐。</summary>
        private static void NormalizeGridCells(List<ThemeCheckBox> checks)
        {
            if (checks.Count == 0) return;
            foreach (var check in checks) check.AutoSize = true;
            int cellWidth = 0;
            int cellHeight = 0;
            foreach (var check in checks)
            {
                var pref = check.GetPreferredSize(Size.Empty);
                cellWidth = Math.Max(cellWidth, pref.Width);
                cellHeight = Math.Max(cellHeight, pref.Height);
            }
            foreach (var check in checks)
            {
                check.AutoSize = false;
                check.Size = new Size(cellWidth, cellHeight);
            }
        }

        private void OnGridCheckChanged(int field, ThemeCheckBox check, int value)
        {
            if (_updating) return;
            if (check.Checked) _selected[field].Add(value);
            else _selected[field].Remove(value);
            EnforceQuartzDayExclusive(field);
            UpdatePreview();
            RaiseChanged();
        }

        /// <summary>网格置顶切换时避免父级背景先擦后画造成闪一下。</summary>
        private sealed class GridFlowPanel : FlowLayoutPanel
        {
            public GridFlowPanel() => DoubleBuffered = true;
        }

        private static IEnumerable<int> ValuesOf(int field)
        {
            if (field == FieldDow) return WeekOrder;
            if (field == FieldYear) return Enumerable.Range(1970, 130);
            return Enumerable.Range(MinOf(field), MaxOf(field) - MinOf(field) + 1);
        }

        private static string ValueLabel(int field, int value)
        {
            if (field == FieldDow) return DayLabels[value];
            if (field == FieldMonth) return value + " 月";
            if (field == FieldSecond || field == FieldMinute) return value.ToString("00");
            return value.ToString();
        }

        // ==================== 表达式生成 ====================

        private void UpdatePreview()
        {
            _preview.Text = "生成：" + BuildExpression() + "　（" + CronToolEngine.DialectName(Dialect) + "）";
        }

        private void RaiseChanged() => ExpressionChanged?.Invoke(this, EventArgs.Empty);

        private string BuildExpression()
        {
            var fields = VisibleFields();
            var parts = new List<string>();
            foreach (int field in fields)
            {
                // 年字段是可选的，不限制时省略，默认生成常见的 6 段 Quartz 表达式
                if (field == FieldYear && IsUnrestricted(field)) continue;
                if (field == FieldDom && Dialect == CronDialect.Quartz)
                {
                    parts.Add(QuartzDomToken());
                    continue;
                }
                if (field == FieldDow && Dialect == CronDialect.Quartz)
                {
                    parts.Add(QuartzDowToken());
                    continue;
                }
                parts.Add(BuildField(field));
            }
            return string.Join(" ", parts);
        }

        private string QuartzDomToken()
        {
            bool domFree = IsUnrestricted(FieldDom);
            bool dowFree = IsUnrestricted(FieldDow);
            if (domFree && dowFree) return "*";
            return domFree ? "?" : BuildField(FieldDom);
        }

        private string QuartzDowToken()
        {
            bool domFree = IsUnrestricted(FieldDom);
            bool dowFree = IsUnrestricted(FieldDow);
            if (domFree && dowFree) return "?";
            return dowFree ? "?" : BuildField(FieldDow);
        }

        private string BuildField(int field)
        {
            switch (_modes[field])
            {
                case ModeRange:
                    int rangeStart = _rangeStart[field];
                    int rangeEnd = _rangeEnd[field];
                    if (rangeStart > rangeEnd) (rangeStart, rangeEnd) = (rangeEnd, rangeStart);
                    string startToken = ValueToken(field, rangeStart);
                    string endToken = ValueToken(field, rangeEnd);
                    return startToken == endToken ? startToken : startToken + "-" + endToken;

                case ModeStep:
                    int step = Math.Max(1, _stepStep[field]);
                    return ValueToken(field, _stepStart[field]) + "/" + step;

                case ModeList:
                    var values = _selected[field].OrderBy(v => v).ToArray();
                    if (values.Length == 0) return "*";
                    return FormatValueList(field, values);

                case ModeSpecial:
                    if (field == FieldDom) return DomSpecialToken();
                    if (field == FieldDow) return DowSpecialToken();
                    return "*";

                default:
                    return "*";
            }
        }

        private string DomSpecialToken()
        {
            switch (_domSpecialKind)
            {
                case 1: return "L-" + Math.Clamp(_domSpecialN, 1, 30);
                case 2: return "LW";
                case 3: return Math.Clamp(_domSpecialN, 1, 31) + "W";
                default: return "L";
            }
        }

        private string DowSpecialToken()
        {
            int raw = Dialect == CronDialect.Quartz ? _specialWeekday + 1 : _specialWeekday;
            if (_dowSpecialKind == 1)
                return raw + "#" + Math.Clamp(_dowSpecialN, 1, 5);
            return raw + "L";
        }

        private string FormatValueList(int field, int[] values)
        {
            var sb = new System.Text.StringBuilder();
            int i = 0;
            while (i < values.Length)
            {
                int start = i;
                while (i + 1 < values.Length && values[i + 1] == values[i] + 1) i++;
                int length = i - start + 1;
                if (sb.Length > 0) sb.Append(',');
                if (length >= 3)
                    sb.Append(ValueToken(field, values[start])).Append('-').Append(ValueToken(field, values[i]));
                else
                {
                    for (int j = start; j <= i; j++)
                    {
                        if (j > start) sb.Append(',');
                        sb.Append(ValueToken(field, values[j]));
                    }
                }
                i++;
            }
            return sb.ToString();
        }

        private static string ValueToken(int field, int value) =>
            field == FieldDow ? DayNames[value] : value.ToString();

        private static bool IsDayField(int field) => field == FieldDom || field == FieldDow;

        private static int MinOf(int field) => field switch
        {
            FieldMonth => 1,
            FieldDom => 1,
            FieldYear => 1970,
            _ => 0
        };

        private static int MaxOf(int field) => field switch
        {
            FieldSecond or FieldMinute => 59,
            FieldHour => 23,
            FieldDom => 31,
            FieldMonth => 12,
            FieldDow => 6,
            FieldYear => 2099,
            _ => 59
        };

        // ==================== 主题 ====================

        public void ApplyTheme()
        {
            if (IsDisposed) return;
            BackColor = Theme.P.Window;
            Font = Theme.Body;

            _head.BackColor = Theme.P.Window;
            _modesRow.BackColor = Theme.P.Window;
            _options.BackColor = Theme.P.Window;
            _gridHost.BackColor = Theme.P.Window;
            _gridHint.BackColor = Theme.P.Window;
            _gridHint.ForeColor = Theme.P.Muted;
            _gridHint.Font = Theme.Body;
            _preview.BackColor = Theme.P.Window;
            _preview.ForeColor = Theme.P.Muted;
            _preview.Font = Theme.Caption;

            _tabs.BackColor = Theme.P.Window;
            _tabs.RefreshTheme();

            StyleTree(this);
            // 皮肤字号可能不同，格子宽度按新字体重新量一遍，避免文字被截断或列错位
            foreach (var pair in _fieldGrids)
            {
                if (!_fieldChecks.TryGetValue(pair.Key, out var checks)) continue;
                pair.Value.SuspendLayout();
                NormalizeGridCells(checks);
                pair.Value.ResumeLayout(true);
            }
            _gridHint.ForeColor = Theme.P.Muted;
            _gridHint.Font = Theme.Body;
            _preview.ForeColor = Theme.P.Muted;
            _preview.Font = Theme.Caption;

            RefreshModeRadios();
            RefreshGrid();
            UpdatePreview();
        }

        /// <summary>子控件颜色是创建时取的令牌值，切皮肤后逐棵重刷；ThemeField / 下拉框各自有更细的绘制逻辑。</summary>
        private static void StyleTree(Control parent)
        {
            foreach (Control child in parent.Controls)
            {
                switch (child)
                {
                    case ThemeField field:
                        field.ApplyTheme();
                        break;
                    case ThemeComboBox combo:
                        combo.Font = Theme.Body;
                        combo.BackColor = Theme.P.Surface;
                        combo.ForeColor = Theme.P.Text;
                        break;
                    case ThemeRadioButton radio:
                        radio.Font = Theme.Body;
                        radio.ForeColor = Theme.P.Text;
                        radio.BackColor = Theme.P.Window;
                        radio.Invalidate();
                        break;
                    case ThemeCheckBox check:
                        check.Font = Theme.Body;
                        check.ForeColor = Theme.P.Text;
                        check.BackColor = Theme.P.Window;
                        check.Invalidate();
                        break;
                    case Label label:
                        label.Font = Theme.Body;
                        label.ForeColor = Theme.P.Muted;
                        label.BackColor = Theme.P.Window;
                        break;
                    case FlowLayoutPanel flow:
                        flow.BackColor = Theme.P.Window;
                        break;
                    case TableLayoutPanel table:
                        table.BackColor = Theme.P.Window;
                        break;
                    case Panel panel when child is not ThemeSegmented:
                        panel.BackColor = Theme.P.Window;
                        break;
                }

                StyleTree(child);
            }
        }

        // ==================== 工具 ====================

        private static void ClearControls(Control parent)
        {
            while (parent.Controls.Count > 0)
            {
                var child = parent.Controls[0];
                parent.Controls.RemoveAt(0);
                child.Dispose();
            }
        }
    }
}