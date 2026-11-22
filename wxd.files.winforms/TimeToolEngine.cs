using System;
using System.Globalization;

namespace wxd.files.winforms
{
    /// <summary>时间戳单位；纳秒只保留 100ns 精度（.NET DateTime 的 tick）。</summary>
    internal enum TimeUnit
    {
        Seconds,
        Milliseconds,
        Nanoseconds
    }

    /// <summary>一次时间换算的结果：识别到的类型 + 格式化后的多行报告。</summary>
    internal readonly struct TimeToolResult
    {
        public bool Success { get; init; }
        public DateTimeOffset Value { get; init; }
        public string InputKind { get; init; }
        public string Report { get; init; }
        public string? Error { get; init; }
    }

    /// <summary>
    /// 时间戳 / 日期文本换算。自动识别秒、毫秒、微秒、纳秒时间戳（含小数），
    /// 也接受 ISO 8601 和常见日期写法；显式单位转换由 TimeToolPanel 的分栏调用。
    /// </summary>
    internal static class TimeToolEngine
    {
        private const long UnixEpochTicks = 621355968000000000L;
        private static readonly string[] Weekdays = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };

        private static readonly string[] DateFormats =
        {
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-dd HH:mm:ss.fff",
            "yyyy/MM/dd HH:mm:ss",
            "yyyy/MM/dd HH:mm:ss.fff",
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyyMMddHHmmss"
        };

        internal static TimeToolResult Convert(string input, DateTimeOffset? now = null)
        {
            string text = (input ?? "").Trim();
            if (text.Length == 0)
                return Fail("请输入时间戳或日期时间");

            if (!TryParse(text, out var value, out string kind, out string? error))
                return Fail(error ?? "无法识别该时间格式");

            return Describe(value, kind, now);
        }

        /// <summary>按指定单位转换，不再按数值大小猜；纳秒输入只有 100ns 精度。</summary>
        internal static TimeToolResult Convert(string input, TimeUnit unit, DateTimeOffset? now = null)
        {
            string text = (input ?? "").Trim();
            if (text.Length == 0)
                return Fail("请输入" + UnitName(unit) + "时间戳");

            if (!decimal.TryParse(text, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture, out decimal number))
                return Fail(UnitName(unit) + "时间戳必须是数字");

            if (!TryFromUnit(number, unit, out var value, out string? error))
                return Fail(error ?? "时间戳超出可表示范围");

            string kind = UnitName(unit) + "级时间戳" + (text.Contains('.') ? "（带小数）" : "");
            return Describe(value, kind, now);
        }

        /// <summary>把已经确定的时间包装成报告，供“时间 → 时间戳”使用。</summary>
        internal static TimeToolResult Describe(DateTimeOffset value, string kind, DateTimeOffset? now = null)
        {
            return new TimeToolResult
            {
                Success = true,
                Value = value,
                InputKind = kind,
                Report = BuildReport(value, kind, now ?? DateTimeOffset.Now)
            };
        }

        internal static string UnitName(TimeUnit unit) => unit switch
        {
            TimeUnit.Milliseconds => "毫秒",
            TimeUnit.Nanoseconds => "纳秒",
            _ => "秒"
        };

        /// <summary>把时间转成指定单位的时间戳；纳秒超出 long 范围时返回 false。</summary>
        internal static bool TryToUnixTime(DateTimeOffset value, TimeUnit unit, out long timestamp, out string? error)
        {
            error = null;
            try
            {
                switch (unit)
                {
                    case TimeUnit.Seconds:
                        timestamp = value.ToUnixTimeSeconds();
                        return true;
                    case TimeUnit.Milliseconds:
                        timestamp = value.ToUnixTimeMilliseconds();
                        return true;
                    default:
                        timestamp = checked((value.UtcDateTime.Ticks - UnixEpochTicks) * 100L);
                        return true;
                }
            }
            catch (OverflowException)
            {
                timestamp = 0;
                error = "纳秒时间戳超出 long 范围（约 1677-09-21 至 2262-04-11）";
                return false;
            }
        }

        private static bool TryParse(string text, out DateTimeOffset value, out string kind, out string? error)
        {
            value = default;
            kind = "";
            error = null;

            if (long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long raw))
                return TryFromNumber(raw, out value, out kind, out error);

            // 带小数的秒时间戳，例如 1700000000.123
            if (decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal dec))
                return TryFromDecimal(dec, out value, out kind, out error);

            if (DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal, out var parsed))
            {
                value = parsed;
                kind = HasExplicitOffset(text) ? "日期时间（带时区）" : "日期时间（本地）";
                return true;
            }

            foreach (string format in DateFormats)
            {
                if (DateTime.TryParseExact(text, format, CultureInfo.InvariantCulture,
                        DateTimeStyles.AssumeLocal, out var local))
                {
                    value = new DateTimeOffset(local);
                    kind = "日期时间（本地）";
                    return true;
                }
            }

            error = "无法识别：支持秒 / 毫秒 / 微秒 / 纳秒时间戳，或 2024-01-01 12:00:00、ISO 8601 等写法";
            return false;
        }

        private static bool TryFromUnit(decimal number, TimeUnit unit, out DateTimeOffset value, out string? error)
        {
            value = default;
            error = null;
            decimal ticks = unit switch
            {
                TimeUnit.Seconds => number * TimeSpan.TicksPerSecond,
                TimeUnit.Milliseconds => number * TimeSpan.TicksPerMillisecond,
                _ => number / 100m
            };

            try
            {
                long rounded = decimal.ToInt64(decimal.Round(ticks, 0, MidpointRounding.AwayFromZero));
                return TryFromUnixTicks(rounded, out value, out error);
            }
            catch (OverflowException)
            {
                error = "时间戳超出可表示范围";
                return false;
            }
        }

        private static bool TryFromNumber(long raw, out DateTimeOffset value, out string kind, out string? error)
        {
            long abs = raw == long.MinValue ? long.MaxValue : Math.Abs(raw);
            if (abs < 100_000_000_000L)
            {
                kind = "秒级时间戳";
                return TryFromUnixTicks(checked(raw * TimeSpan.TicksPerSecond), out value, out error);
            }
            if (abs < 100_000_000_000_000L)
            {
                kind = "毫秒级时间戳";
                return TryFromUnixTicks(checked(raw * TimeSpan.TicksPerMillisecond), out value, out error);
            }
            if (abs < 100_000_000_000_000_000L)
            {
                kind = "微秒级时间戳";
                return TryFromUnixTicks(checked(raw * 10L), out value, out error);
            }

            // long 能表示的时间戳里，剩余全部按纳秒处理；上限约到 2262 年
            kind = "纳秒级时间戳";
            long ticks = decimal.ToInt64(decimal.Round(raw / 100m, 0, MidpointRounding.AwayFromZero));
            return TryFromUnixTicks(ticks, out value, out error);
        }

        private static bool TryFromDecimal(decimal dec, out DateTimeOffset value, out string kind, out string? error)
        {
            decimal abs = Math.Abs(dec);
            decimal ticks;
            if (abs < 100_000_000_000m)
            {
                kind = "秒级时间戳（带小数）";
                ticks = dec * TimeSpan.TicksPerSecond;
            }
            else if (abs < 100_000_000_000_000m)
            {
                kind = "毫秒级时间戳（带小数）";
                ticks = dec * TimeSpan.TicksPerMillisecond;
            }
            else if (abs < 100_000_000_000_000_000m)
            {
                kind = "微秒级时间戳（带小数）";
                ticks = dec * 10m;
            }
            else if (abs < 10_000_000_000_000_000_000m)
            {
                kind = "纳秒级时间戳（带小数）";
                ticks = dec / 100m;
            }
            else
            {
                value = default;
                kind = "";
                error = "时间戳位数过多，无法识别（支持秒 / 毫秒 / 微秒 / 纳秒）";
                return false;
            }

            try
            {
                return TryFromUnixTicks(decimal.ToInt64(decimal.Round(ticks, 0, MidpointRounding.AwayFromZero)), out value, out error);
            }
            catch (OverflowException)
            {
                value = default;
                error = "时间戳超出可表示范围";
                return false;
            }
        }

        private static bool TryFromUnixTicks(long unixTicks, out DateTimeOffset value, out string? error)
        {
            value = default;
            error = null;
            try
            {
                long ticks = checked(UnixEpochTicks + unixTicks);
                if (ticks < DateTime.MinValue.Ticks || ticks > DateTime.MaxValue.Ticks)
                {
                    error = "时间戳超出可表示范围";
                    return false;
                }
                value = new DateTimeOffset(ticks, TimeSpan.Zero);
                return true;
            }
            catch (OverflowException)
            {
                error = "时间戳超出可表示范围";
                return false;
            }
        }

        private static bool HasExplicitOffset(string text)
        {
            string t = text.Trim();
            if (t.EndsWith("Z", StringComparison.OrdinalIgnoreCase)) return true;

            // 只在时间部分里找 +08:00 / -0500，日期里的 - 不算时区
            int timeStart = Math.Max(t.IndexOf('T'), t.IndexOf(' '));
            if (timeStart < 0) return false;

            for (int i = t.Length - 1; i > timeStart; i--)
                if (t[i] == '+' || t[i] == '-') return true;
            return false;
        }

        private static string BuildReport(DateTimeOffset value, string kind, DateTimeOffset now)
        {
            var utc = value.ToUniversalTime();
            var local = value.ToLocalTime();
            long ms = value.ToUnixTimeMilliseconds();
            long sec = value.ToUnixTimeSeconds();
            string ns = TryToUnixTime(value, TimeUnit.Nanoseconds, out long nano, out _)
                ? nano.ToString(CultureInfo.InvariantCulture)
                : "超出 long 范围（约 1677-09-21 至 2262-04-11）";
            string offsetText = $"UTC{(local.Offset >= TimeSpan.Zero ? "+" : "-")}{Math.Abs(local.Offset.Hours):D2}:{Math.Abs(local.Offset.Minutes):D2}";

            return string.Join(Environment.NewLine, new[]
            {
                "识别为      : " + kind,
                "UTC         : " + utc.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                "UTC 高精度  : " + utc.ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture),
                "UTC ISO     : " + utc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture),
                "本地时间    : " + local.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "  (" + offsetText + ")",
                "本地高精度  : " + local.ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture),
                "本地 ISO    : " + local.ToString("yyyy-MM-dd'T'HH:mm:sszzz", CultureInfo.InvariantCulture),
                "秒时间戳    : " + sec.ToString(CultureInfo.InvariantCulture),
                "毫秒时间戳  : " + ms.ToString(CultureInfo.InvariantCulture),
                "纳秒时间戳  : " + ns,
                "星期        : " + Weekdays[(int)local.DayOfWeek],
                "相对现在    : " + DescribeRelative(value - now)
            });
        }

        private static string DescribeRelative(TimeSpan delta)
        {
            bool future = delta >= TimeSpan.Zero;
            double seconds = Math.Abs(delta.TotalSeconds);
            string suffix = future ? "后" : "前";
            if (seconds < 10) return "刚刚";
            if (seconds < 60) return $"{(int)seconds} 秒{suffix}";
            if (seconds < 3600) return $"{(int)(seconds / 60)} 分钟{suffix}";
            if (seconds < 86400) return $"{(int)(seconds / 3600)} 小时{suffix}";
            return $"{(int)(seconds / 86400)} 天{suffix}";
        }

        private static TimeToolResult Fail(string error)
            => new() { Success = false, Error = error, Report = "" };
    }
}
