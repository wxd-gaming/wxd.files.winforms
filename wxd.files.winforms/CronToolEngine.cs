using System;
using System.Collections.Generic;
using System.Linq;

namespace wxd.files.winforms
{
    /// <summary>cron 表达式所属规范。</summary>
    internal enum CronDialect
    {
        /// <summary>标准 cron：5 段，或沿用旧版的 6 段（含秒），日/周按“或”匹配。</summary>
        Standard,
        /// <summary>Quartz：6 段必填 + 可选年，星期 1-7=周日-周六，日/周必须有一列写 ?。</summary>
        Quartz,
        /// <summary>Spring：6 段，星期 0-7=周日-周六，日/周同时限定时需同时满足，支持 @daily 等宏。</summary>
        Spring
    }

    /// <summary>解析后的 cron 表达式：各字段取值 + 中文说明。</summary>
    internal sealed class CronSchedule
    {
        public bool Success { get; init; }
        public string? Error { get; init; }
        public string? Warning { get; init; }
        public CronDialect Dialect { get; init; }
        public string Expression { get; init; } = "";
        public string? Macro { get; init; }
        public string? Normalized { get; init; }
        public bool WithSeconds { get; init; }
        public bool WithYear { get; init; }

        public int[] Seconds { get; init; } = Array.Empty<int>();
        public int[] Minutes { get; init; } = Array.Empty<int>();
        public int[] Hours { get; init; } = Array.Empty<int>();
        public int[] DaysOfMonth { get; init; } = Array.Empty<int>();
        public int[] Months { get; init; } = Array.Empty<int>();
        /// <summary>内部编号：0=周日 … 6=周六。</summary>
        public int[] DaysOfWeek { get; init; } = Array.Empty<int>();
        public int[] Years { get; init; } = Array.Empty<int>();

        public bool DomWildcard { get; init; }
        public bool DowWildcard { get; init; }
        public bool YearWildcard { get; init; } = true;

        /// <summary>日字段的 L：当月最后一天。</summary>
        public bool DomIsLast { get; init; }
        /// <summary>日字段的 L-n：倒数第 n 天（-1 表示未用）。</summary>
        public int DomLastOffset { get; init; } = -1;
        /// <summary>日字段的 LW：当月最后一个工作日。</summary>
        public bool DomIsLastWeekday { get; init; }
        /// <summary>日字段的 nW：最接近 n 号的工作日（-1 表示未用）。</summary>
        public int DomNearestWeekday { get; init; } = -1;

        /// <summary>周字段的 dL：当月最后一个星期 d（内部编号，-1 表示未用）。</summary>
        public int DowLastWeekday { get; init; } = -1;
        /// <summary>周字段的 d#n：当月第 n 个星期 d（内部编号，-1 表示未用）。</summary>
        public int DowNthWeekday { get; init; } = -1;
        public int DowNth { get; init; }

        /// <summary>与 FieldNotes 一一对应的字段名。</summary>
        public string[] FieldLabels { get; init; } = Array.Empty<string>();

        /// <summary>按 秒/分/时/日/月/周/年 顺序的原始字段 token，缺省字段补默认值；示例回填生成器要用。</summary>
        public string[] Tokens { get; init; } = Array.Empty<string>();

        /// <summary>逐字段中文说明，顺序与 FieldLabels 一致。</summary>
        public string[] FieldNotes { get; init; } = Array.Empty<string>();

        /// <summary>一句话概括，用于结果首行。</summary>
        public string Summary { get; init; } = "";

        public string DialectName => CronToolEngine.DialectName(Dialect);
    }

    /// <summary>
    /// cron 表达式解析与触发时间推算：
    /// 标准 5/6 段、Quartz 6/7 段、Spring 6 段（含宏）。
    /// Quartz / Spring 支持 L / W / #，星期编号两套规范不同，解析时按所选规范处理。
    /// </summary>
    internal static class CronToolEngine
    {
        private static readonly Dictionary<string, int> MonthNames = new(StringComparer.OrdinalIgnoreCase)
        {
            ["JAN"] = 1, ["FEB"] = 2, ["MAR"] = 3, ["APR"] = 4, ["MAY"] = 5, ["JUN"] = 6,
            ["JUL"] = 7, ["AUG"] = 8, ["SEP"] = 9, ["OCT"] = 10, ["NOV"] = 11, ["DEC"] = 12
        };

        /// <summary>标准 / Spring：0=周日 … 6=周六。</summary>
        private static readonly Dictionary<string, int> DayNamesCron = new(StringComparer.OrdinalIgnoreCase)
        {
            ["SUN"] = 0, ["MON"] = 1, ["TUE"] = 2, ["WED"] = 3, ["THU"] = 4, ["FRI"] = 5, ["SAT"] = 6
        };

        /// <summary>Quartz：1=周日 … 7=周六。</summary>
        private static readonly Dictionary<string, int> DayNamesQuartz = new(StringComparer.OrdinalIgnoreCase)
        {
            ["SUN"] = 1, ["MON"] = 2, ["TUE"] = 3, ["WED"] = 4, ["THU"] = 5, ["FRI"] = 6, ["SAT"] = 7
        };

        private static readonly Dictionary<string, string> SpringMacros = new(StringComparer.OrdinalIgnoreCase)
        {
            ["@yearly"] = "0 0 0 1 1 *",
            ["@annually"] = "0 0 0 1 1 *",
            ["@monthly"] = "0 0 0 1 * *",
            ["@weekly"] = "0 0 0 * * 0",
            ["@daily"] = "0 0 0 * * *",
            ["@midnight"] = "0 0 0 * * *",
            ["@hourly"] = "0 0 * * * *"
        };

        private static readonly string[] WeekdayNames = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };

        private sealed class DomSpec
        {
            public int[] Values = Array.Empty<int>();
            public bool Wildcard;
            public bool IsLast;
            public int LastOffset = -1;
            public bool IsLastWeekday;
            public int NearestWeekday = -1;
        }

        private sealed class DowSpec
        {
            public int[] Values = Array.Empty<int>();
            public bool Wildcard;
            public int LastWeekday = -1;
            public int NthWeekday = -1;
            public int Nth;
        }

        internal static string DialectName(CronDialect dialect) => dialect switch
        {
            CronDialect.Quartz => "Quartz",
            CronDialect.Spring => "Spring",
            _ => "标准 cron"
        };

        /// <summary>旧入口：按标准 cron 解析（5 段，或沿用旧版的 6 段）。</summary>
        internal static CronSchedule Parse(string expression) => Parse(expression, CronDialect.Standard);

        internal static CronSchedule Parse(string expression, CronDialect dialect)
        {
            string expr = (expression ?? "").Trim();
            if (expr.Length == 0) return Fail(dialect, "请输入 cron 表达式");

            string working = expr;
            string? macro = null;
            string? warning = null;

            if (expr.StartsWith("@", StringComparison.Ordinal))
            {
                if (dialect != CronDialect.Spring)
                    return Fail(dialect, "只有 Spring 支持 @yearly / @daily 这类宏");
                if (!SpringMacros.TryGetValue(expr, out string? expanded))
                    return Fail(dialect,
                        $"不支持的宏“{expr}”；Spring 支持 @yearly / @annually / @monthly / @weekly / @daily / @midnight / @hourly");
                working = expanded;
                macro = expr;
            }

            string[] parts = working.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

            // 5 段一律按标准 cron 理解：Quartz/Spring 都需要秒字段，但用户粘贴标准表达式时给个提示直接算
            if (parts.Length == 5 && dialect != CronDialect.Standard)
            {
                warning = $"5 段表达式按标准 cron 解析（{DialectName(dialect)} 需要 6 段起步）";
                dialect = CronDialect.Standard;
            }

            if (dialect == CronDialect.Quartz)
            {
                if (parts.Length != 6 && parts.Length != 7)
                    return Fail(dialect, "Quartz 需要 6 段（秒 分 时 日 月 周）或 7 段（末尾加年）");
            }
            else if (dialect == CronDialect.Spring)
            {
                if (parts.Length != 6)
                    return Fail(dialect, "Spring 需要 6 段（秒 分 时 日 月 周），可选年字段请改用 Quartz");
            }
            else if (parts.Length != 5 && parts.Length != 6)
            {
                return Fail(dialect, "需要 5 段（分 时 日 月 周）或 6 段（秒 分 时 日 月 周）");
            }

            bool withSeconds = parts.Length >= 6;
            bool withYear = dialect == CronDialect.Quartz && parts.Length == 7;
            int index = 0;

            string secondToken = withSeconds ? parts[index++] : "0";
            string minuteToken = parts[index++];
            string hourToken = parts[index++];
            string domToken = parts[index++];
            string monthToken = parts[index++];
            string dowToken = parts[index++];
            string yearToken = withYear ? parts[index] : "*";

            bool allowQuestionEverywhere = dialect == CronDialect.Standard;

            // 标准模式沿用旧行为：任意字段的 ? 都当作 *
            if (allowQuestionEverywhere)
            {
                secondToken = NormalizeQuestion(secondToken);
                minuteToken = NormalizeQuestion(minuteToken);
                hourToken = NormalizeQuestion(hourToken);
                monthToken = NormalizeQuestion(monthToken);
            }

            if (!allowQuestionEverywhere && secondToken.Trim() == "?")
                return Fail(dialect, "? 只能用在日或周字段");
            if (!allowQuestionEverywhere && minuteToken.Trim() == "?")
                return Fail(dialect, "? 只能用在日或周字段");
            if (!allowQuestionEverywhere && hourToken.Trim() == "?")
                return Fail(dialect, "? 只能用在日或周字段");
            if (!allowQuestionEverywhere && monthToken.Trim() == "?")
                return Fail(dialect, "? 只能用在日或周字段");
            if (withYear && yearToken.Trim() == "?")
                return Fail(dialect, "? 只能用在日或周字段");

            if (!TryParseList(secondToken, 0, 59, null, out int[] seconds, out _, out string error))
                return Fail(dialect, "秒字段：" + error);
            if (!TryParseList(minuteToken, 0, 59, null, out int[] minutes, out _, out error))
                return Fail(dialect, "分钟字段：" + error);
            if (!TryParseList(hourToken, 0, 23, null, out int[] hours, out _, out error))
                return Fail(dialect, "小时字段：" + error);
            if (!TryParseList(monthToken, 1, 12, MonthNames, out int[] months, out _, out error))
                return Fail(dialect, "月字段：" + error);

            if (!TryParseDom(domToken, dialect, out DomSpec dom, out error))
                return Fail(dialect, "日字段：" + error);
            if (!TryParseDow(dowToken, dialect, out DowSpec dow, out error))
                return Fail(dialect, "星期字段：" + error);

            if (dialect == CronDialect.Quartz)
            {
                // Quartz 不支持日/周同时具体：必须有一列写 ?
                bool domQuestion = domToken.Trim() == "?";
                bool dowQuestion = dowToken.Trim() == "?";
                if (!domQuestion && !dowQuestion)
                    return Fail(dialect, "Quartz 要求日/周其中一列写 ?（两列不能同时指定具体值）");
            }

            int[] years = Array.Empty<int>();
            bool yearWildcard = true;
            if (withYear)
            {
                if (!TryParseList(yearToken, 1970, 2099, null, out years, out yearWildcard, out error))
                    return Fail(dialect, "年字段：" + error);
            }

            var notes = new List<string>();
            var labels = new List<string>();
            if (withSeconds)
            {
                labels.Add("秒");
                notes.Add(Describe("秒", seconds, secondToken));
            }
            labels.Add("分");
            notes.Add(Describe("分", minutes, minuteToken));
            labels.Add("时");
            notes.Add(Describe("点", hours, hourToken));
            labels.Add("日");
            notes.Add(DescribeDom(dom));
            labels.Add("月");
            notes.Add(Describe("月", months, monthToken));
            labels.Add("周");
            notes.Add(DescribeDow(dow));
            if (!dom.Wildcard && !dow.Wildcard)
            {
                if (dialect == CronDialect.Spring)
                    notes[^1] += "（Spring 日/周同时限定时需两者都满足）";
                else if (dialect == CronDialect.Standard)
                    notes[^1] += "（日/周同时限定时满足任一即可）";
            }
            if (withYear)
            {
                labels.Add("年");
                notes.Add(yearWildcard ? "每年" : DescribeYear(years));
            }

            string segment = withYear ? "7 段（含秒、年）" : withSeconds ? "6 段（含秒）" : "5 段（标准）";
            return new CronSchedule
            {
                Success = true,
                Dialect = dialect,
                Expression = expr,
                Macro = macro,
                Normalized = macro == null ? null : working,
                WithSeconds = withSeconds,
                WithYear = withYear,
                Seconds = seconds,
                Minutes = minutes,
                Hours = hours,
                DaysOfMonth = dom.Values,
                Months = months,
                DaysOfWeek = dow.Values,
                Years = years,
                DomWildcard = dom.Wildcard,
                DowWildcard = dow.Wildcard,
                YearWildcard = yearWildcard,
                DomIsLast = dom.IsLast,
                DomLastOffset = dom.LastOffset,
                DomIsLastWeekday = dom.IsLastWeekday,
                DomNearestWeekday = dom.NearestWeekday,
                DowLastWeekday = dow.LastWeekday,
                DowNthWeekday = dow.NthWeekday,
                DowNth = dow.Nth,
                FieldLabels = labels.ToArray(),
                FieldNotes = notes.ToArray(),
                Tokens = new[] { secondToken, minuteToken, hourToken, domToken, monthToken, dowToken, yearToken },
                Summary = segment + " · " + string.Join("；", notes),
                Warning = warning
            };
        }

        // ==================== 字段解析 ====================

        private static bool TryParseDom(string token, CronDialect dialect, out DomSpec spec, out string error)
        {
            spec = new DomSpec();
            error = "";
            string t = token.Trim();
            if (t.Length == 0)
            {
                error = "字段为空";
                return false;
            }

            string up = t.ToUpperInvariant();
            if (up == "?" || up == "*")
            {
                spec.Wildcard = true;
                return true;
            }

            bool allowSpecials = dialect != CronDialect.Standard;
            if (allowSpecials)
            {
                if (up == "L")
                {
                    spec.IsLast = true;
                    return true;
                }
                if (up == "LW")
                {
                    spec.IsLastWeekday = true;
                    return true;
                }
                if (up.StartsWith("L-", StringComparison.Ordinal))
                {
                    if (int.TryParse(up[2..], out int offset) && offset is >= 1 and <= 30)
                    {
                        spec.LastOffset = offset;
                        return true;
                    }
                    error = $"“{t}”不合法；L-n 的 n 应在 1-30";
                    return false;
                }
                if (up.EndsWith("W", StringComparison.Ordinal) && up.Length > 1)
                {
                    if (int.TryParse(up[..^1], out int day) && day is >= 1 and <= 31)
                    {
                        spec.NearestWeekday = day;
                        return true;
                    }
                    error = $"“{t}”不合法；W 只能写成 15W 这种单日形式";
                    return false;
                }
            }

            if (up.Contains('L') || up.Contains('W') || up.Contains('#'))
            {
                error = allowSpecials
                    ? $"“{t}”不合法；L / W / # 不能和列表或范围混用"
                    : $"标准 cron 不支持 L / W / #（“{t}”）";
                return false;
            }

            if (!TryParseList(t, 1, 31, null, out int[] values, out bool wildcard, out error))
                return false;
            spec.Values = values;
            spec.Wildcard = wildcard;
            return true;
        }

        private static bool TryParseDow(string token, CronDialect dialect, out DowSpec spec, out string error)
        {
            spec = new DowSpec();
            error = "";
            string t = token.Trim();
            if (t.Length == 0)
            {
                error = "字段为空";
                return false;
            }

            string up = t.ToUpperInvariant();
            if (up == "?" || up == "*")
            {
                spec.Wildcard = true;
                return true;
            }

            bool allowSpecials = dialect != CronDialect.Standard;
            if (allowSpecials && up == "L")
            {
                if (dialect == CronDialect.Quartz)
                {
                    // Quartz 里 L 单独用等价 7 / SAT，不是“最后一个周六”
                    spec.Values = new[] { 6 };
                    return true;
                }
                error = "Spring 的 L 需要带星期前缀，例如 5L / FRIL";
                return false;
            }

            if (allowSpecials && up.EndsWith("L", StringComparison.Ordinal) && up.Length > 1)
            {
                string prefix = up[..^1];
                if (!TryDayValue(prefix, dialect, out int raw))
                {
                    error = $"“{t}”不合法；dL 的 d 应是星期值或 SUN-SAT";
                    return false;
                }
                spec.LastWeekday = ToInternalWeekday(raw, dialect);
                return true;
            }

            int hash = up.IndexOf('#');
            if (hash >= 0)
            {
                if (!allowSpecials)
                {
                    error = $"标准 cron 不支持 #（“{t}”）";
                    return false;
                }
                string prefix = up[..hash];
                string nthText = up[(hash + 1)..];
                if (!TryDayValue(prefix, dialect, out int raw)
                    || !int.TryParse(nthText, out int nth)
                    || nth is < 1 or > 5)
                {
                    error = $"“{t}”不合法；# 应写成 6#3（当月第 3 个周五）";
                    return false;
                }
                spec.NthWeekday = ToInternalWeekday(raw, dialect);
                spec.Nth = nth;
                return true;
            }

            var names = dialect == CronDialect.Quartz ? DayNamesQuartz : DayNamesCron;
            int min = dialect == CronDialect.Quartz ? 1 : 0;
            if (!TryParseList(t, min, 7, names, out int[] rawValues, out bool wildcard, out error))
            {
                if (up.Contains('L') || up.Contains('W') || up.Contains('#'))
                {
                    error = allowSpecials
                        ? $"“{t}”不合法；L / # 不能和列表或范围混用"
                        : $"标准 cron 不支持 L / W / #（“{t}”）";
                }
                return false;
            }

            spec.Values = rawValues
                .Select(v => ToInternalWeekday(v, dialect))
                .Distinct()
                .OrderBy(v => v)
                .ToArray();
            spec.Wildcard = wildcard;
            return true;
        }

        private static string NormalizeQuestion(string token) =>
            token.Trim() == "?" ? "*" : token;

        private static bool TryDayValue(string text, CronDialect dialect, out int value)
        {
            var names = dialect == CronDialect.Quartz ? DayNamesQuartz : DayNamesCron;
            if (names.TryGetValue(text, out value)) return true;
            if (!int.TryParse(text, out value)) return false;
            int min = dialect == CronDialect.Quartz ? 1 : 0;
            return value >= min && value <= 7;
        }

        private static int ToInternalWeekday(int raw, CronDialect dialect)
        {
            // Quartz：1=周日 … 7=周六；标准 / Spring：0/7=周日 … 6=周六
            if (dialect == CronDialect.Quartz) return (raw - 1) % 7;
            return raw % 7;
        }

        private static bool TryParseList(
            string token, int min, int max, Dictionary<string, int>? names,
            out int[] values, out bool wildcard, out string error)
        {
            values = Array.Empty<int>();
            error = "";
            string trimmed = token.Trim();
            wildcard = trimmed == "*";
            if (trimmed.Length == 0)
            {
                error = "字段为空";
                return false;
            }

            var set = new SortedSet<int>();
            foreach (string rawItem in trimmed.Split(','))
            {
                string item = rawItem.Trim();
                if (item.Length == 0)
                {
                    error = "取值列表里有空项";
                    return false;
                }

                int step = 1;
                string rangePart = item;
                int slash = item.IndexOf('/');
                if (slash >= 0)
                {
                    rangePart = item[..slash];
                    if (!int.TryParse(item[(slash + 1)..], out step) || step is <= 0 or > 10000)
                    {
                        error = $"步长“{item}”不合法";
                        return false;
                    }
                    if (rangePart.Length == 0)
                    {
                        error = $"“{item}”缺少起始值";
                        return false;
                    }
                }

                int start, end;
                if (rangePart == "*")
                {
                    start = min;
                    end = max;
                }
                else if (rangePart.Contains('-'))
                {
                    string[] pair = rangePart.Split('-');
                    if (pair.Length != 2
                        || !TryValue(pair[0], names, out start)
                        || !TryValue(pair[1], names, out end))
                    {
                        error = $"范围“{rangePart}”不合法";
                        return false;
                    }
                    if (start > end)
                    {
                        error = $"范围“{rangePart}”起点大于终点";
                        return false;
                    }
                }
                else if (!TryValue(rangePart, names, out start))
                {
                    error = $"取值“{rangePart}”不合法";
                    return false;
                }
                else
                {
                    end = slash >= 0 ? max : start;
                }

                if (start < min || end > max)
                {
                    error = $"取值“{item}”超出范围 {min}-{max}";
                    return false;
                }

                for (int v = start; v <= end; v += step)
                    set.Add(v);
            }

            if (set.Count == 0)
            {
                error = "字段没有产生任何取值";
                return false;
            }

            values = set.ToArray();
            return true;
        }

        private static bool TryValue(string text, Dictionary<string, int>? names, out int value)
        {
            text = text.Trim();
            if (names != null && names.TryGetValue(text, out value)) return true;
            return int.TryParse(text, out value);
        }

        // ==================== 中文说明 ====================

        private static string Describe(string unit, int[] values, string token)
        {
            if (token.Trim() == "*" || token.Trim() == "?")
            {
                return unit switch
                {
                    "月" => "每月",
                    "点" => "每小时",
                    "分" => "每分钟",
                    "秒" => "每秒",
                    _ => "每" + unit
                };
            }

            int slash = token.IndexOf('/');
            if (slash > 0 && int.TryParse(token[(slash + 1)..], out int step))
                return $"每 {step} {unit}";

            if (values.Length == 1)
                return FormatValue(values[0].ToString(), unit);

            string joined = string.Join(", ", values.Take(12).Select(v => FormatValue(v.ToString(), unit)));
            if (values.Length > 12) joined += " …";
            return joined;
        }

        private static string DescribeDom(DomSpec spec)
        {
            if (spec.Wildcard) return "每天";
            if (spec.IsLast) return "每月最后一天";
            if (spec.IsLastWeekday) return "每月最后一个工作日";
            if (spec.LastOffset >= 0) return $"每月倒数第 {spec.LastOffset} 天";
            if (spec.NearestWeekday > 0) return $"每月最接近 {spec.NearestWeekday} 号的工作日";
            if (spec.Values.Length == 1) return $"每月 {spec.Values[0]} 号";

            string joined = string.Join(", ", spec.Values.Take(12).Select(v => v + " 号"));
            if (spec.Values.Length > 12) joined += " …";
            return joined;
        }

        private static string DescribeDow(DowSpec spec)
        {
            if (spec.Wildcard) return "每天";
            if (spec.LastWeekday >= 0) return $"每月最后一个{WeekdayNames[spec.LastWeekday]}";
            if (spec.NthWeekday >= 0) return $"每月第 {spec.Nth} 个{WeekdayNames[spec.NthWeekday]}";
            if (spec.Values.Length == 1) return WeekdayNames[spec.Values[0]];

            string joined = string.Join(", ", spec.Values.Select(v => WeekdayNames[v]));
            return joined;
        }

        private static string DescribeYear(int[] years)
        {
            if (years.Length == 0) return "每年";
            if (years.Length == 1) return years[0] + " 年";
            string joined = string.Join(", ", years.Take(12).Select(v => v + " 年"));
            if (years.Length > 12) joined += " …";
            return joined;
        }

        private static string FormatValue(string raw, string unit)
        {
            if (int.TryParse(raw, out int v))
            {
                if (unit == "月") return v + " 月";
                if (unit == "点") return v + " 点";
                if (unit == "分") return "第 " + v + " 分钟";
                if (unit == "秒") return "第 " + v + " 秒";
            }
            return raw;
        }

        // ==================== 触发时间 ====================

        /// <summary>从 from 之后开始算，返回接下来 count 次触发时间（本地时间）。</summary>
        internal static List<DateTime> NextRuns(CronSchedule schedule, DateTime from, int count)
        {
            var results = new List<DateTime>();
            if (!schedule.Success || count <= 0) return results;

            DateTime day = from.Date;
            DateTime limit;
            if (schedule.Dialect == CronDialect.Quartz && !schedule.YearWildcard && schedule.Years.Length > 0)
            {
                int maxYear = schedule.Years.Max();
                if (maxYear < from.Year) return results;
                limit = new DateTime(maxYear, 12, 31);
            }
            else
            {
                limit = day.AddYears(8);
            }

            while (day <= limit && results.Count < count)
            {
                if (DayMatches(schedule, day))
                {
                    foreach (int h in schedule.Hours)
                    {
                        foreach (int m in schedule.Minutes)
                        {
                            foreach (int s in schedule.Seconds)
                            {
                                var candidate = day.AddHours(h).AddMinutes(m).AddSeconds(s);
                                if (candidate <= from) continue;
                                if (TimeZoneInfo.Local.IsInvalidTime(candidate)) continue;

                                results.Add(candidate);
                                if (results.Count >= count) return results;
                            }
                        }
                    }
                }
                day = day.AddDays(1);
            }

            return results;
        }

        private static bool DayMatches(CronSchedule schedule, DateTime day)
        {
            if (schedule.Dialect == CronDialect.Quartz
                && !schedule.YearWildcard
                && !schedule.Years.Contains(day.Year))
                return false;

            if (!schedule.Months.Contains(day.Month)) return false;

            bool domMatch = DomMatches(schedule, day);
            bool dowMatch = DowMatches(schedule, day);

            if (schedule.DomWildcard && schedule.DowWildcard) return true;
            if (schedule.DomWildcard) return dowMatch;
            if (schedule.DowWildcard) return domMatch;

            // Spring 的 ? 等同 *，日、周两个约束都会参与判断（官方实现为“与”）
            if (schedule.Dialect == CronDialect.Spring)
                return domMatch && dowMatch;

            // 标准 cron：日和周都被限定时满足任一即可
            return domMatch || dowMatch;
        }

        private static bool DomMatches(CronSchedule schedule, DateTime day)
        {
            if (schedule.DomWildcard) return true;
            int daysInMonth = DateTime.DaysInMonth(day.Year, day.Month);

            if (schedule.DomIsLast) return day.Day == daysInMonth;
            // L-n 表示倒数第 n 天：L-1 是倒数第 1 天，L-3 是倒数第 3 天
            if (schedule.DomLastOffset >= 0) return day.Day == daysInMonth - schedule.DomLastOffset + 1;
            if (schedule.DomIsLastWeekday) return day.Day == LastWeekdayOfMonth(day.Year, day.Month);
            if (schedule.DomNearestWeekday > 0)
                return day.Day == NearestWeekday(day.Year, day.Month, schedule.DomNearestWeekday);

            return schedule.DaysOfMonth.Contains(day.Day);
        }

        private static bool DowMatches(CronSchedule schedule, DateTime day)
        {
            if (schedule.DowWildcard) return true;
            int weekday = (int)day.DayOfWeek;

            if (schedule.DowLastWeekday >= 0)
                return weekday == schedule.DowLastWeekday && day.AddDays(7).Month != day.Month;

            if (schedule.DowNthWeekday >= 0)
                return weekday == schedule.DowNthWeekday && (day.Day - 1) / 7 + 1 == schedule.DowNth;

            return schedule.DaysOfWeek.Contains(weekday);
        }

        private static int LastWeekdayOfMonth(int year, int month)
        {
            int day = DateTime.DaysInMonth(year, month);
            while (day > 1)
            {
                var date = new DateTime(year, month, day);
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                    return day;
                day--;
            }
            return day;
        }

        /// <summary>离 n 号最近的工作日，且不跨出当月；n 超出当月天数时返回 -1。</summary>
        private static int NearestWeekday(int year, int month, int n)
        {
            int daysInMonth = DateTime.DaysInMonth(year, month);
            if (n < 1 || n > daysInMonth) return -1;

            var date = new DateTime(year, month, n);
            if (date.DayOfWeek == DayOfWeek.Saturday)
                date = n == 1 ? date.AddDays(2) : date.AddDays(-1);
            else if (date.DayOfWeek == DayOfWeek.Sunday)
                date = n == daysInMonth ? date.AddDays(-2) : date.AddDays(1);
            return date.Day;
        }

        internal static string WeekdayName(DateTime value) => WeekdayNames[(int)value.DayOfWeek];

        private static CronSchedule Fail(CronDialect dialect, string error) =>
            new() { Success = false, Dialect = dialect, Error = error };
    }
}