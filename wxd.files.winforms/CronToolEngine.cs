using System;
using System.Collections.Generic;
using System.Linq;

namespace wxd.files.winforms
{
    /// <summary>解析后的 cron 表达式：各字段取值 + 中文说明。</summary>
    internal sealed class CronSchedule
    {
        public bool Success { get; init; }
        public string? Error { get; init; }
        public string Expression { get; init; } = "";
        public bool WithSeconds { get; init; }
        public int[] Seconds { get; init; } = Array.Empty<int>();
        public int[] Minutes { get; init; } = Array.Empty<int>();
        public int[] Hours { get; init; } = Array.Empty<int>();
        public int[] DaysOfMonth { get; init; } = Array.Empty<int>();
        public int[] Months { get; init; } = Array.Empty<int>();
        public int[] DaysOfWeek { get; init; } = Array.Empty<int>();
        public bool DomWildcard { get; init; }
        public bool DowWildcard { get; init; }

        /// <summary>逐字段中文说明，按 秒/分/时/日/月/周 顺序。</summary>
        public string[] FieldNotes { get; init; } = Array.Empty<string>();

        /// <summary>一句话概括，用于结果首行。</summary>
        public string Summary { get; init; } = "";
    }

    /// <summary>
    /// cron 表达式解析：支持标准 5 段（分 时 日 月 周）和 6 段（秒 分 时 日 月 周），
    /// 支持 * , - / 与英文月份/星期名，? 当作 *；不支持 L / W / # 扩展。
    /// </summary>
    internal static class CronToolEngine
    {
        private static readonly Dictionary<string, int> MonthNames = new(StringComparer.OrdinalIgnoreCase)
        {
            ["JAN"] = 1, ["FEB"] = 2, ["MAR"] = 3, ["APR"] = 4, ["MAY"] = 5, ["JUN"] = 6,
            ["JUL"] = 7, ["AUG"] = 8, ["SEP"] = 9, ["OCT"] = 10, ["NOV"] = 11, ["DEC"] = 12
        };

        private static readonly Dictionary<string, int> DayNames = new(StringComparer.OrdinalIgnoreCase)
        {
            ["SUN"] = 0, ["MON"] = 1, ["TUE"] = 2, ["WED"] = 3, ["THU"] = 4, ["FRI"] = 5, ["SAT"] = 6
        };

        private static readonly string[] WeekdayNames = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };

        internal static CronSchedule Parse(string expression)
        {
            string expr = (expression ?? "").Trim();
            if (expr.Length == 0) return Fail("请输入 cron 表达式");

            string[] parts = expr.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 5 && parts.Length != 6)
                return Fail("需要 5 段（分 时 日 月 周）或 6 段（秒 分 时 日 月 周）");

            bool withSeconds = parts.Length == 6;
            int index = 0;

            if (!TryParseField(withSeconds ? parts[index++] : "0", 0, 59, null,
                    out int[] seconds, out _, out string error))
                return Fail("秒字段：" + error);

            if (!TryParseField(parts[index++], 0, 59, null,
                    out int[] minutes, out _, out error))
                return Fail("分钟字段：" + error);

            if (!TryParseField(parts[index++], 0, 23, null,
                    out int[] hours, out _, out error))
                return Fail("小时字段：" + error);

            if (!TryParseField(parts[index++], 1, 31, null,
                    out int[] daysOfMonth, out bool domWildcard, out error))
                return Fail("日字段：" + error);

            if (!TryParseField(parts[index++], 1, 12, MonthNames,
                    out int[] months, out _, out error))
                return Fail("月字段：" + error);

            string dowToken = parts[index];
            if (!TryParseField(dowToken, 0, 7, DayNames,
                    out int[] daysOfWeek, out bool dowWildcard, out error))
                return Fail("星期字段：" + error);

            // cron 里 0 和 7 都是星期日
            daysOfWeek = daysOfWeek.Select(v => v == 7 ? 0 : v).Distinct().OrderBy(v => v).ToArray();

            var notes = new List<string>();
            if (withSeconds)
                notes.Add(Describe("秒", "秒", parts[0], seconds, false));
            notes.Add(Describe("分", "分钟", parts[withSeconds ? 1 : 0], minutes, false));
            notes.Add(Describe("时", "点", parts[withSeconds ? 2 : 1], hours, false));
            notes.Add(Describe("日", "号", parts[withSeconds ? 3 : 2], daysOfMonth, domWildcard));
            notes.Add(Describe("月", "月", parts[withSeconds ? 4 : 3], months, false));
            notes.Add(Describe("周", "星期", dowToken, daysOfWeek, dowWildcard));

            var schedule = new CronSchedule
            {
                Success = true,
                Expression = expr,
                WithSeconds = withSeconds,
                Seconds = seconds,
                Minutes = minutes,
                Hours = hours,
                DaysOfMonth = daysOfMonth,
                Months = months,
                DaysOfWeek = daysOfWeek,
                DomWildcard = domWildcard,
                DowWildcard = dowWildcard,
                FieldNotes = notes.ToArray(),
                Summary = (withSeconds ? "6 段（含秒）" : "5 段（标准）") + " · " + string.Join("；", notes)
            };
            return schedule;
        }

        private static bool TryParseField(
            string token, int min, int max, Dictionary<string, int>? names,
            out int[] values, out bool wildcard, out string error)
        {
            values = Array.Empty<int>();
            error = "";
            string trimmed = token.Trim();
            wildcard = trimmed == "*" || trimmed == "?";
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
                    if (!int.TryParse(item[(slash + 1)..], out step) || step <= 0 || step > 10000)
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
                if (rangePart == "*" || rangePart == "?")
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
                    error = $"取值“{rangePart}”不合法（不支持 L / W / #）";
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

        private static string Describe(string label, string unit, string token, int[] values, bool wildcard)
        {
            if (wildcard)
            {
                return unit switch
                {
                    "星期" => "每天",
                    "号" => "每天",
                    "月" => "每月",
                    "点" => "每小时",
                    "分钟" => "每分钟",
                    "秒" => "每秒",
                    _ => "每" + unit
                };
            }

            int slash = token.IndexOf('/');
            if (slash > 0 && int.TryParse(token[(slash + 1)..], out int step))
                return $"每 {step} {unit}";

            if (token.Contains('-'))
            {
                string[] pair = token.Split('-');
                return $"{FormatValue(pair[0], unit)} 到 {FormatValue(pair[1], unit)}";
            }

            if (values.Length == 1)
                return FormatValue(values[0].ToString(), unit);

            string joined = string.Join(", ", values.Take(12).Select(v => FormatValue(v.ToString(), unit)));
            if (values.Length > 12) joined += " …";
            return joined;
        }

        private static string FormatValue(string raw, string unit)
        {
            if (int.TryParse(raw, out int v))
            {
                if (unit == "星期") return WeekdayNames[v is >= 0 and <= 6 ? v : 0];
                if (unit == "月") return v + " 月";
                if (unit == "号") return v + " 号";
                if (unit == "点") return v + " 点";
                if (unit == "分钟") return "第 " + v + " 分钟";
                if (unit == "秒") return "第 " + v + " 秒";
            }
            return raw;
        }

        /// <summary>从 from 之后开始算，返回接下来 count 次触发时间（本地时间）。</summary>
        internal static List<DateTime> NextRuns(CronSchedule schedule, DateTime from, int count)
        {
            var results = new List<DateTime>();
            if (!schedule.Success || count <= 0) return results;

            DateTime day = from.Date;
            DateTime limit = day.AddYears(8);

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
            if (!schedule.Months.Contains(day.Month)) return false;

            bool domMatch = schedule.DaysOfMonth.Contains(day.Day);
            bool dowMatch = schedule.DaysOfWeek.Contains((int)day.DayOfWeek);

            if (schedule.DomWildcard && schedule.DowWildcard) return true;
            if (schedule.DomWildcard) return dowMatch;
            if (schedule.DowWildcard) return domMatch;
            // 标准 cron：日和周都被限定时，满足任一即可
            return domMatch || dowMatch;
        }

        internal static string WeekdayName(DateTime value) => WeekdayNames[(int)value.DayOfWeek];

        private static CronSchedule Fail(string error) => new() { Success = false, Error = error };
    }
}
