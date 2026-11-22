using System;
using System.Linq;
using System.Text.Json.Nodes;
using wxd.files.winforms;

namespace wxd.files.winforms.RegressionTests
{
    internal static partial class Program
    {
        private static void RunToolboxTests()
        {
            // 时间戳
            Run("时间戳 秒级换算与报告", Test_TimeSeconds);
            Run("时间戳 毫秒/微秒保留精度", Test_TimeMillisAndMicros);
            Run("时间戳 显式单位换算", Test_TimeExplicitUnits);
            Run("时间戳 纳秒自动识别", Test_TimeNanoseconds);
            Run("时间戳 时间转指定单位", Test_TimeToUnits);
            Run("时间戳 非法输入", Test_TimeInvalid);

            // 编码
            Run("编码 Unicode 转义/还原", Test_UnicodeEscapeRoundTrip);
            Run("编码 UTF-8 Hex 往返", Test_Utf8HexRoundTrip);
            Run("编码 Base64 往返", Test_Base64RoundTrip);
            Run("编码 URL / HTML 往返", Test_UrlAndHtmlRoundTrip);

            // 哈希
            Run("哈希 MD5/SHA256 标准向量", Test_HashVectors);
            Run("哈希 文件流式 MD5", Test_FileMd5);
            Run("HTTP 参数组装与解码", Test_HttpAssembly);
            Run("HTTP 编码与请求体", Test_HttpEncoding);
            Run("HTTP 保存请求序列化", Test_HttpSavedRequests);

            // Cron
            Run("Cron 工作日每5分钟", Test_CronWeekdayStep);
            Run("Cron 6段含秒", Test_CronWithSeconds);
            Run("Cron 非法表达式", Test_CronInvalid);
            Run("Cron 日/周 OR 语义", Test_CronDayOrWeekday);

            // Quartz / Spring
            Run("Quartz 工作日与 ? 规则", Test_QuartzWeekdayAndQuestion);
            Run("Quartz 星期编号 1=周日", Test_QuartzSundayIsOne);
            Run("Quartz L / L-n / 6L / 6#3", Test_QuartzLastAndNth);
            Run("Quartz 15W 最近工作日", Test_QuartzNearestWeekday);
            Run("Quartz 年字段限制", Test_QuartzYear);
            Run("Quartz 5段回退标准", Test_QuartzFiveFieldFallback);
            Run("Spring 0/7=周日", Test_SpringSunday);
            Run("含 W 的星期名", Test_WeekdayNamesContainingW);
            Run("Spring 日/周 AND 语义", Test_SpringDayAndWeek);
            Run("Spring L / L-n / 5L", Test_SpringLastRules);
            Run("Spring 宏", Test_SpringMacros);
            Run("Spring 段数限制", Test_SpringFieldCount);

            // YAML
            Run("YAML 转 JSON 类型推断", Test_YamlToJsonTypes);
            Run("YAML JSON 往返", Test_JsonYamlRoundTrip);
            Run("YAML 多文档/语法错误", Test_YamlErrors);
        }

        // ==================== 时间戳 ====================

        private static void Test_TimeSeconds()
        {
            var result = TimeToolEngine.Convert("1700000000");
            Assert(result.Success, "合法秒时间戳应成功");
            AssertEq("秒级时间戳", result.InputKind, "识别类型");
            Assert(result.Report.Contains("2023-11-14 22:13:20"), "UTC 时间应正确");
            Assert(result.Report.Contains("1700000000"), "应输出秒时间戳");
            Assert(result.Report.Contains("1700000000000"), "应输出毫秒时间戳");
        }

        private static void Test_TimeMillisAndMicros()
        {
            var millis = TimeToolEngine.Convert("1700000000123");
            Assert(millis.Success, "毫秒时间戳应成功");
            AssertEq("毫秒级时间戳", millis.InputKind, "识别毫秒");
            Assert(millis.Report.Contains("22:13:20.123"), "毫秒应保留");

            var micros = TimeToolEngine.Convert("1700000000123456");
            Assert(micros.Success, "微秒时间戳应成功");
            AssertEq("微秒级时间戳", micros.InputKind, "识别微秒");
            Assert(micros.Report.Contains("22:13:20.123"), "微秒换算后仍应保留毫秒");
        }

        private static void Test_TimeInvalid()
        {
            Assert(!TimeToolEngine.Convert("").Success, "空输入应失败");
            Assert(!TimeToolEngine.Convert("not-a-time").Success, "乱文本应失败");
            Assert(!TimeToolEngine.Convert("99999999999999999999").Success, "超长数字应失败");
        }

        private static void Test_TimeExplicitUnits()
        {
            var seconds = TimeToolEngine.Convert("1700000000", TimeUnit.Seconds);
            Assert(seconds.Success, "显式秒时间戳应成功");
            AssertEq("秒级时间戳", seconds.InputKind, "显式秒类型");
            Assert(seconds.Report.Contains("2023-11-14 22:13:20"), "显式秒 UTC 时间应正确");

            var millis = TimeToolEngine.Convert("1700000000123", TimeUnit.Milliseconds);
            Assert(millis.Success, "显式毫秒时间戳应成功");
            AssertEq("毫秒级时间戳", millis.InputKind, "显式毫秒类型");
            Assert(millis.Report.Contains("22:13:20.1230000"), "显式毫秒应保留精度");

            var nanos = TimeToolEngine.Convert("1700000000123456789", TimeUnit.Nanoseconds);
            Assert(nanos.Success, "显式纳秒时间戳应成功");
            AssertEq("纳秒级时间戳", nanos.InputKind, "显式纳秒类型");
            Assert(nanos.Report.Contains("22:13:20.1234568"), "纳秒应四舍五入到 100ns");
            Assert(nanos.Report.Contains("1700000000123456800"), "报告应按 100ns 精度回显纳秒");

            Assert(!TimeToolEngine.Convert("abc", TimeUnit.Seconds).Success, "非数字应失败");
            Assert(!TimeToolEngine.Convert("", TimeUnit.Nanoseconds).Success, "空纳秒应失败");
        }

        private static void Test_TimeNanoseconds()
        {
            var result = TimeToolEngine.Convert("1700000000123456789");
            Assert(result.Success, "纳秒时间戳应能自动识别");
            AssertEq("纳秒级时间戳", result.InputKind, "自动识别纳秒");
            Assert(result.Report.Contains("22:13:20.1234568"), "自动纳秒应四舍五入到 100ns");
            Assert(result.Report.Contains("纳秒时间戳"), "报告应列出纳秒时间戳");
        }

        private static void Test_TimeToUnits()
        {
            var value = new DateTimeOffset(2023, 11, 14, 22, 13, 20, TimeSpan.Zero);

            Assert(TimeToolEngine.TryToUnixTime(value, TimeUnit.Seconds, out long seconds, out _), "时间转秒应成功");
            AssertEq(1700000000L, seconds, "时间转秒值");
            Assert(TimeToolEngine.TryToUnixTime(value, TimeUnit.Milliseconds, out long millis, out _), "时间转毫秒应成功");
            AssertEq(1700000000000L, millis, "时间转毫秒值");
            Assert(TimeToolEngine.TryToUnixTime(value, TimeUnit.Nanoseconds, out long nanos, out _), "时间转纳秒应成功");
            AssertEq(1700000000000000000L, nanos, "时间转纳秒值");

            var report = TimeToolEngine.Describe(value, "指定时间（本地）");
            Assert(report.Success, "指定时间报告应成功");
            Assert(report.Report.Contains("2023-11-14 22:13:20"), "报告应包含指定时间");

            var overflow = new DateTimeOffset(2500, 1, 1, 0, 0, 0, TimeSpan.Zero);
            Assert(!TimeToolEngine.TryToUnixTime(overflow, TimeUnit.Nanoseconds, out _, out string? error),
                "超出 long 的纳秒应失败");
            Assert(error != null && error.Contains("纳秒"), "超范围错误应说明纳秒");
        }

        // ==================== 编码 ====================

        private static void Test_UnicodeEscapeRoundTrip()
        {
            AssertEq("\\u53F2\\u83B1\\u59C6",
                EncodingToolEngine.EscapeUnicode("史莱姆", asciiOnly: false), "全量转义");
            AssertEq("a\\u53F2",
                EncodingToolEngine.EscapeUnicode("a史", asciiOnly: true), "仅非 ASCII");

            var back = EncodingToolEngine.UnescapeUnicode("a\\u53F2\\u83B1\\u59C6");
            Assert(back.Success, "还原应成功");
            AssertEq("a史莱姆", back.Text, "Unicode 往返");

            var emoji = EncodingToolEngine.UnescapeUnicode("\\uD83D\\uDE00");
            Assert(emoji.Success && emoji.Text == "😀", "代理对还原");

            var bad = EncodingToolEngine.UnescapeUnicode("\\u12");
            Assert(!bad.Success, "不完整 \\u 应报错");
        }

        private static void Test_Utf8HexRoundTrip()
        {
            AssertEq("E5 8F B2", EncodingToolEngine.ToUtf8Hex("史"), "UTF-8 Hex");
            var back = EncodingToolEngine.FromUtf8Hex("0xE5, 0x8F, 0xB2");
            Assert(back.Success, "Hex 还原应成功");
            AssertEq("史", back.Text, "Hex 往返");

            Assert(!EncodingToolEngine.FromUtf8Hex("E5 8F").Success, "奇数位应报错");
            Assert(!EncodingToolEngine.FromUtf8Hex("FF").Success, "非法 UTF-8 应报错");
        }

        private static void Test_Base64RoundTrip()
        {
            AssertEq("5Y+y", EncodingToolEngine.ToBase64("史"), "Base64 编码");
            var back = EncodingToolEngine.FromBase64("5Y+y");
            Assert(back.Success, "Base64 解码应成功");
            AssertEq("史", back.Text, "Base64 往返");

            var unpadded = EncodingToolEngine.FromBase64("5Y+y");
            Assert(unpadded.Success, "带填充的 Base64 应成功");
            Assert(!EncodingToolEngine.FromBase64("!!!").Success, "非法 Base64 应报错");
        }

        private static void Test_UrlAndHtmlRoundTrip()
        {
            AssertEq("a%20b%26c", EncodingToolEngine.UrlEncode("a b&c"), "URL 编码");
            AssertEq("a b&c", EncodingToolEngine.UrlDecode("a%20b%26c"), "URL 解码");

            string encoded = EncodingToolEngine.HtmlEncode("<a href=\"x\">&</a>");
            Assert(encoded.Contains("&lt;") && encoded.Contains("&amp;"), "HTML 编码");
            AssertEq("<a href=\"x\">&</a>", EncodingToolEngine.HtmlDecode(encoded), "HTML 往返");
        }

        // ==================== 哈希 ====================

        private static void Test_HashVectors()
        {
            AssertEq("d41d8cd98f00b204e9800998ecf8427e", HashToolEngine.Md5Hex(""), "MD5 空串");
            AssertEq("900150983cd24fb0d6963f7d28e17f72", HashToolEngine.Md5Hex("abc"), "MD5 abc");
            AssertEq("D41D8CD98F00B204E9800998ECF8427E", HashToolEngine.Md5Hex("", upper: true), "MD5 大写");
            AssertEq("e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
                HashToolEngine.Sha256Hex(""), "SHA256 空串");
        }

        private static void Test_FileMd5()
        {
            // %TEMP% 对仓库跑出来的低完整性进程不可写，落在本程序输出目录（同为 Low 标签，可写）
            string path = Path.Combine(AppContext.BaseDirectory, "wxd-filemd5-" + Guid.NewGuid().ToString("N") + ".bin");
            try
            {
                File.WriteAllText(path, "abc");
                AssertEq("900150983cd24fb0d6963f7d28e17f72", HashToolEngine.FileMd5Hex(path), "文件 MD5 标准向量");

                // 1MB 数据跨多个流式缓冲块，验证增量计算和整读结果一致
                var data = new byte[1024 * 1024];
                new Random(20261121).NextBytes(data);
                File.WriteAllBytes(path, data);
                string expected = Convert.ToHexString(System.Security.Cryptography.MD5.HashData(data)).ToLowerInvariant();
                AssertEq(expected, HashToolEngine.FileMd5Hex(path), "流式与整读一致");
                AssertEq(expected.ToUpperInvariant(), HashToolEngine.FileMd5Hex(path, upper: true), "文件 MD5 大写");

                File.WriteAllBytes(path, Array.Empty<byte>());
                AssertEq("d41d8cd98f00b204e9800998ecf8427e", HashToolEngine.FileMd5Hex(path), "空文件 MD5");

                bool threw = false;
                try { HashToolEngine.FileMd5Hex(path + ".不存在"); }
                catch (FileNotFoundException) { threw = true; }
                Assert(threw, "文件不存在应抛 FileNotFoundException");
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static void Test_HttpAssembly()
        {
            var pairs = new List<KeyValuePair<string, string>>
            {
                new("q", "史莱姆"), new("tag", "a&b c")
            };
            AssertEq("https://x/api?q=%E5%8F%B2%E8%8E%B1%E5%A7%86&tag=a%26b%20c",
                HttpToolEngine.BuildUrl("https://x/api", pairs), "URL 追加查询串");
            AssertEq("https://x/api?p=1&size=9",
                HttpToolEngine.BuildUrl("https://x/api?p=1", new List<KeyValuePair<string, string>> { new("size", "9") }),
                "URL 已带 ? 时用 & 续接");
            AssertEq("https://x/api", HttpToolEngine.BuildUrl("https://x/api", pairs: new List<KeyValuePair<string, string>>()),
                "无参数时 URL 原样");

            AssertEq("{\"a\":\"1\",\"\\u4E2D\\u6587\":\"\\u003Cx\\u003E\"}",
                HttpToolEngine.BuildJsonBody(new List<KeyValuePair<string, string>> { new("a", "1"), new("中文", "<x>") }),
                "JSON 保持键顺序并对非 ASCII/HTML 字符转义");

            AssertEq("k=1%262&n=%E5%8F%B2%E8%8E%B1%E5%A7%86",
                HttpToolEngine.BuildFormBody(new List<KeyValuePair<string, string>> { new("k", "1&2"), new("n", "史莱姆") },
                    System.Text.Encoding.UTF8),
                "Form 转义保留同名键");

            AssertEq("", HttpToolEngine.DecodeBody(Array.Empty<byte>(), "application/json", System.Text.Encoding.UTF8), "空响应");
            AssertEq("中文", HttpToolEngine.DecodeBody(
                System.Text.Encoding.UTF8.GetBytes("中文"), "application/json; charset=utf-8", System.Text.Encoding.UTF8), "UTF-8 解码");
            AssertEq("中文", HttpToolEngine.DecodeBody(
                new byte[] { 0xD6, 0xD0, 0xCE, 0xC4 }, "text/html; charset=gbk", System.Text.Encoding.UTF8), "GBK 解码");
            Assert(HttpToolEngine.DecodeBody(new byte[] { 0xFF, 0xD8 }, "image/png", System.Text.Encoding.UTF8) == null, "图片返回 null");
        }

        private static void Test_HttpEncoding()
        {
            // GBK 表单体：按所选编码转义，字节能还原回中文
            string body = HttpToolEngine.BuildFormBody(
                new List<KeyValuePair<string, string>> { new("n", "史莱姆") },
                System.Text.Encoding.GetEncoding("GBK"));
            Assert(body.StartsWith("n=%", StringComparison.Ordinal), "GBK 模式按 GBK 字节做 %XX 转义");
            var bytes = new List<byte>();
            for (int i = 2; i < body.Length; i += 3)
                bytes.Add(Convert.ToByte(body.Substring(i + 1, 2), 16));
            AssertEq("史莱姆", System.Text.Encoding.GetEncoding("GBK").GetString(bytes.ToArray()), "GBK 字节还原中文");

            // 自定义 Content-Type 的兜底：原样拼接不转义
            AssertEq("cmd=login&ts=123",
                HttpToolEngine.BuildRawBody(new List<KeyValuePair<string, string>> { new("cmd", "login"), new("ts", "123") }),
                "Raw 拼接不转义");

            AssertEq("utf-8", HttpToolEngine.ResolveCharset("text/html", System.Text.Encoding.UTF8), "无 charset 回退面板编码");
            AssertEq("gb2312", HttpToolEngine.ResolveCharset("text/html; charset=GBK", System.Text.Encoding.UTF8), "charset 命中返回真实编码名");
            AssertEq("utf-8", HttpToolEngine.ResolveCharset("text/html; charset=x-nope", System.Text.Encoding.UTF8), "未知 charset 回退面板编码");
            AssertEq("中文", HttpToolEngine.DecodeBody(
                new byte[] { 0xD6, 0xD0, 0xCE, 0xC4 }, "text/plain", System.Text.Encoding.GetEncoding("GBK")),
                "无 charset 时按面板所选编码解码");
            AssertEq("utf-8", HttpToolEngine.GetEncoding(" UTF-8 ").WebName, "编码名容忍首尾空格");
            Assert(ReferenceEquals(HttpToolEngine.GetEncoding("no-such-enc"), System.Text.Encoding.UTF8), "未知编码回退 UTF-8");
        }

        private static void Test_HttpSavedRequests()
        {
            var saved = new List<SavedHttpRequest>
            {
                new()
                {
                    Name = "登录", Method = "POST", Url = "https://api.example.com/login?src=测",
                    ContentType = "multipart/form-data",
                    BodyFormat = "Form",
                    Params = new[] { new[] { "src", "test" } },
                    BodyParams = new[] { new[] { "user", "史莱姆" }, new[] { "level", "99" } },
                    Headers = new[] { new[] { "X-Game", "wxd" }, new[] { "Authorization", "Bearer t" } },
                    Count = 5, Threads = 2
                },
                new()
                {
                    Name = "查询", Method = "GET", Url = "https://api.example.com/query",
                    Params = Array.Empty<string[]>(), Headers = Array.Empty<string[]>()
                },
                new()
                {
                    // 新档 none：ContentType 是空串（区别于旧档的 null），保存后重开仍是 none
                    Name = "无体", Method = "POST", Url = "https://api.example.com/ping",
                    ContentType = "", BodyFormat = "JSON",
                    Params = Array.Empty<string[]>(), Headers = Array.Empty<string[]>()
                }
            };

            string json = HttpToolEngine.SerializeRequests(saved);
            var back = HttpToolEngine.DeserializeRequests(json);
            Assert(back.Count == 3, "往返后条数一致");
            AssertEq("登录", back[0].Name, "名称含中文往返一致");
            AssertEq("POST", back[0].Method, "方法往返一致");
            AssertEq("https://api.example.com/login?src=测", back[0].Url, "URL 往返一致");
            AssertEq("Form", back[0].BodyFormat, "提交格式往返一致");
            AssertEq("multipart/form-data", back[0].ContentType!, "Content-Type 往返一致");
            Assert(back[1].ContentType == null, "旧档无 ContentType 保持 null");
            Assert(back[1].BodyParams == null, "旧档无 BodyParams 保持 null");
            AssertEq("", back[2].ContentType!, "新档 none 存空串往返一致");
            Assert(back[0].Params.Length == 1 && back[0].Params[0][0] == "src", "URL 参数往返一致");
            Assert(back[0].BodyParams != null && back[0].BodyParams.Length == 2
                && back[0].BodyParams[0][0] == "user" && back[0].BodyParams[0][1] == "史莱姆"
                && back[0].BodyParams[1][1] == "99", "请求体键值往返一致");
            Assert(back[0].Headers.Length == 2 && back[0].Headers[0][0] == "X-Game"
                && back[0].Headers[1][0] == "Authorization", "请求头往返一致");
            AssertEq(5, back[0].Count, "次数往返一致");
            AssertEq(2, back[0].Threads, "线程往返一致");
            Assert(back[1].Params.Length == 0 && back[1].Headers.Length == 0, "空网格往返一致");
            Assert(json.Contains("\"user\"") && json.Contains("史莱姆"), "序列化中文保持原样");

            bool threw = false;
            try { HttpToolEngine.DeserializeRequests("{ 这不是 JSON"); }
            catch (System.Text.Json.JsonException) { threw = true; }
            Assert(threw, "损坏文本应抛 JsonException");
        }

        // ==================== Cron ====================

        private static void Test_CronWeekdayStep()
        {
            var schedule = CronToolEngine.Parse("*/5 9-18 * * 1-5");
            Assert(schedule.Success, "表达式应解析成功");
            Assert(!schedule.WithSeconds, "应为 5 段");

            var from = new DateTime(2026, 11, 23, 8, 58, 0); // 星期一
            var runs = CronToolEngine.NextRuns(schedule, from, 10);
            AssertEq(10, runs.Count, "应找到 10 次");
            AssertEq(new DateTime(2026, 11, 23, 9, 0, 0), runs[0], "首次应为周一 9:00");
            Assert(runs.All(r => r.Minute % 5 == 0), "分钟应为 5 的倍数");
            Assert(runs.All(r => r.DayOfWeek >= DayOfWeek.Monday && r.DayOfWeek <= DayOfWeek.Friday),
                "只应落在工作日");
            Assert(runs.All(r => r.Hour >= 9 && r.Hour <= 18), "小时应在 9-18");
        }

        private static void Test_CronWithSeconds()
        {
            var schedule = CronToolEngine.Parse("0 0/5 * * * ?");
            Assert(schedule.Success, "6 段表达式应成功");
            Assert(schedule.WithSeconds, "应识别含秒");

            var runs = CronToolEngine.NextRuns(schedule, new DateTime(2026, 11, 23, 10, 1, 30), 3);
            AssertEq(new DateTime(2026, 11, 23, 10, 5, 0), runs[0], "首次应为 10:05:00");
            AssertEq(new DateTime(2026, 11, 23, 10, 10, 0), runs[1], "第二次应为 10:10:00");
        }

        private static void Test_CronInvalid()
        {
            Assert(!CronToolEngine.Parse("").Success, "空表达式应失败");
            Assert(!CronToolEngine.Parse("61 * * * *").Success, "分钟越界应失败");
            Assert(!CronToolEngine.Parse("* * *").Success, "段数不足应失败");
            Assert(!CronToolEngine.Parse("0 0 L * *").Success, "不支持的 L 应失败");
        }

        private static void Test_CronDayOrWeekday()
        {
            var schedule = CronToolEngine.Parse("0 0 1 * 1");
            Assert(schedule.Success, "日/周限定表达式应成功");
            var runs = CronToolEngine.NextRuns(schedule, new DateTime(2026, 11, 20, 0, 0, 0), 10);
            Assert(runs.Count > 0, "应找到触发时间");
            Assert(runs.All(r => r.Day == 1 || r.DayOfWeek == DayOfWeek.Monday),
                "日或周任一匹配都应触发");
        }

        // ==================== Quartz / Spring ====================

        private static void Test_QuartzWeekdayAndQuestion()
        {
            var schedule = CronToolEngine.Parse("0 0/5 9-18 ? * MON-FRI", CronDialect.Quartz);
            Assert(schedule.Success, "Quartz 工作日表达式应成功: " + schedule.Error);
            AssertEq(CronDialect.Quartz, schedule.Dialect, "规范");
            Assert(schedule.WithSeconds && !schedule.WithYear, "6 段含秒且无年");

            var runs = CronToolEngine.NextRuns(schedule, new DateTime(2026, 11, 23, 8, 58, 0), 10);
            AssertEq(new DateTime(2026, 11, 23, 9, 0, 0), runs[0], "首次应为周一 9:00");
            Assert(runs.All(r => r.DayOfWeek >= DayOfWeek.Monday && r.DayOfWeek <= DayOfWeek.Friday),
                "只应落在工作日");
            Assert(runs.All(r => r.Minute % 5 == 0 && r.Hour >= 9 && r.Hour <= 18),
                "分钟/小时应正确");

            Assert(!CronToolEngine.Parse("0 0 0 1 * MON", CronDialect.Quartz).Success,
                "日/周都指定具体值应失败");
            Assert(!CronToolEngine.Parse("0 0 ? 1 * MON", CronDialect.Quartz).Success,
                "? 不能用在秒字段");
            Assert(CronToolEngine.Parse("0 0 0 ? * ?", CronDialect.Quartz).Success,
                "两列都写 ? 应可解析");
        }

        private static void Test_QuartzSundayIsOne()
        {
            var schedule = CronToolEngine.Parse("0 0 12 ? * 1", CronDialect.Quartz);
            Assert(schedule.Success, "表达式应成功: " + schedule.Error);

            var runs = CronToolEngine.NextRuns(schedule, new DateTime(2026, 11, 21, 13, 0, 0), 1);
            AssertEq(new DateTime(2026, 11, 22, 12, 0, 0), runs[0], "Quartz 的 1 应为周日");
            AssertEq(DayOfWeek.Sunday, runs[0].DayOfWeek, "触发日应为周日");

            var lastSaturday = CronToolEngine.Parse("0 0 12 ? * L", CronDialect.Quartz);
            Assert(lastSaturday.Success, "单独的 L 应可解析: " + lastSaturday.Error);
            var saturday = CronToolEngine.NextRuns(lastSaturday, new DateTime(2026, 11, 20, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 11, 21, 12, 0, 0), saturday[0], "单独的 L 应等价周六");
        }

        private static void Test_QuartzLastAndNth()
        {
            var lastDay = CronToolEngine.Parse("0 0 12 L * ?", CronDialect.Quartz);
            Assert(lastDay.Success, "L 应成功: " + lastDay.Error);
            var lastDayRuns = CronToolEngine.NextRuns(lastDay, new DateTime(2026, 2, 1, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 2, 28, 12, 0, 0), lastDayRuns[0], "2 月最后一天");

            var lastSecond = CronToolEngine.Parse("0 0 12 L-2 * ?", CronDialect.Quartz);
            Assert(lastSecond.Success, "L-2 应成功: " + lastSecond.Error);
            var lastSecondRuns = CronToolEngine.NextRuns(lastSecond, new DateTime(2026, 11, 1, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 11, 29, 12, 0, 0), lastSecondRuns[0], "11 月倒数第 2 天");

            var lastFriday = CronToolEngine.Parse("0 0 12 ? * 6L", CronDialect.Quartz);
            Assert(lastFriday.Success, "6L 应成功: " + lastFriday.Error);
            var lastFridayRuns = CronToolEngine.NextRuns(lastFriday, new DateTime(2026, 11, 1, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 11, 27, 12, 0, 0), lastFridayRuns[0], "最后一个周五");

            var thirdFriday = CronToolEngine.Parse("0 0 12 ? * 6#3", CronDialect.Quartz);
            Assert(thirdFriday.Success, "6#3 应成功: " + thirdFriday.Error);
            var thirdFridayRuns = CronToolEngine.NextRuns(thirdFriday, new DateTime(2026, 11, 1, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 11, 20, 12, 0, 0), thirdFridayRuns[0], "第 3 个周五");
        }

        private static void Test_QuartzNearestWeekday()
        {
            var schedule = CronToolEngine.Parse("0 0 9 15W * ?", CronDialect.Quartz);
            Assert(schedule.Success, "15W 应成功: " + schedule.Error);

            // 2026-08-15 是周六，应落到前一个工作日周五 08-14
            var runs = CronToolEngine.NextRuns(schedule, new DateTime(2026, 8, 1, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 8, 14, 9, 0, 0), runs[0], "15 号周六应提前到周五");

            var firstW = CronToolEngine.Parse("0 0 9 1W * ?", CronDialect.Quartz);
            Assert(firstW.Success, "1W 应成功: " + firstW.Error);
            var firstRuns = CronToolEngine.NextRuns(firstW, new DateTime(2026, 8, 1, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 8, 3, 9, 0, 0), firstRuns[0], "1 号周六应顺延到周一");
        }

        private static void Test_QuartzYear()
        {
            var schedule = CronToolEngine.Parse("0 0 0 1 1 ? 2030", CronDialect.Quartz);
            Assert(schedule.Success, "带年应成功: " + schedule.Error);
            Assert(schedule.WithYear, "应识别年字段");
            AssertEq(2030, schedule.Years[0], "年取值");

            var runs = CronToolEngine.NextRuns(schedule, new DateTime(2026, 6, 1, 0, 0, 0), 5);
            AssertEq(1, runs.Count, "只应命中一次");
            AssertEq(new DateTime(2030, 1, 1, 0, 0, 0), runs[0], "2030-01-01");

            Assert(!CronToolEngine.Parse("0 0 0 1 1 ? 2100", CronDialect.Quartz).Success,
                "年份越界应失败");
        }

        private static void Test_QuartzFiveFieldFallback()
        {
            var schedule = CronToolEngine.Parse("*/5 9-18 * * 1-5", CronDialect.Quartz);
            Assert(schedule.Success, "5 段应回退标准解析: " + schedule.Error);
            AssertEq(CronDialect.Standard, schedule.Dialect, "回退后规范为标准");
            Assert(!string.IsNullOrEmpty(schedule.Warning), "应给出回退提示");
        }

        private static void Test_SpringSunday()
        {
            var zero = CronToolEngine.Parse("0 0 12 * * 0", CronDialect.Spring);
            Assert(zero.Success, "0 应可解析: " + zero.Error);
            var zeroRuns = CronToolEngine.NextRuns(zero, new DateTime(2026, 11, 21, 13, 0, 0), 1);
            AssertEq(DayOfWeek.Sunday, zeroRuns[0].DayOfWeek, "Spring 的 0 应为周日");

            var seven = CronToolEngine.Parse("0 0 12 * * 7", CronDialect.Spring);
            Assert(seven.Success, "7 应可解析: " + seven.Error);
            var sevenRuns = CronToolEngine.NextRuns(seven, new DateTime(2026, 11, 21, 13, 0, 0), 1);
            AssertEq(DayOfWeek.Sunday, sevenRuns[0].DayOfWeek, "Spring 的 7 也应为周日");
        }

        private static void Test_WeekdayNamesContainingW()
        {
            var spring = CronToolEngine.Parse("0 0 12 ? * WED", CronDialect.Spring);
            Assert(spring.Success, "Spring 的 WED 应可解析: " + spring.Error);
            var springRun = CronToolEngine.NextRuns(spring, new DateTime(2026, 11, 22, 0, 0, 0), 1);
            AssertEq(DayOfWeek.Wednesday, springRun[0].DayOfWeek, "WED 应为周三");

            var list = CronToolEngine.Parse("0 0 12 ? * MON,WED,FRI", CronDialect.Quartz);
            Assert(list.Success, "Quartz 含 WED 的列表应可解析: " + list.Error);
            var listRuns = CronToolEngine.NextRuns(list, new DateTime(2026, 11, 22, 0, 0, 0), 3);
            AssertEq(DayOfWeek.Monday, listRuns[0].DayOfWeek, "列表首个应为周一");
            AssertEq(DayOfWeek.Wednesday, listRuns[1].DayOfWeek, "列表第二个应为周三");
            AssertEq(DayOfWeek.Friday, listRuns[2].DayOfWeek, "列表第三个应为周五");

            Assert(!CronToolEngine.Parse("0 0 12 ? * 1W", CronDialect.Quartz).Success,
                "W 特殊语法只属于日字段");
        }

        private static void Test_SpringDayAndWeek()
        {
            var both = CronToolEngine.Parse("0 0 12 15 * MON", CronDialect.Spring);
            Assert(both.Success, "日/周同时限定应成功: " + both.Error);
            var bothRuns = CronToolEngine.NextRuns(both, new DateTime(2026, 1, 1, 0, 0, 0), 4);
            Assert(bothRuns.All(r => r.Day == 15 && r.DayOfWeek == DayOfWeek.Monday),
                "Spring 日/周同时限定时应两者都满足");
            AssertEq(new DateTime(2026, 6, 15, 12, 0, 0), bothRuns[0], "首个“15 号且周一”");

            var domOnly = CronToolEngine.Parse("0 0 12 15 * ?", CronDialect.Spring);
            Assert(domOnly.Success, "日字段限定 + ? 应成功: " + domOnly.Error);
            var domRuns = CronToolEngine.NextRuns(domOnly, new DateTime(2026, 11, 1, 0, 0, 0), 2);
            AssertEq(new DateTime(2026, 11, 15, 12, 0, 0), domRuns[0], "Spring 的 ? 等同 *，只按日字段");
        }

        private static void Test_SpringLastRules()
        {
            var lastDay = CronToolEngine.Parse("0 0 12 L * ?", CronDialect.Spring);
            Assert(lastDay.Success, "Spring L 应成功: " + lastDay.Error);
            var lastDayRuns = CronToolEngine.NextRuns(lastDay, new DateTime(2026, 2, 1, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 2, 28, 12, 0, 0), lastDayRuns[0], "Spring 最后一天");

            var lastSecond = CronToolEngine.Parse("0 0 12 L-2 * ?", CronDialect.Spring);
            Assert(lastSecond.Success, "Spring L-2 应成功: " + lastSecond.Error);
            var lastSecondRuns = CronToolEngine.NextRuns(lastSecond, new DateTime(2026, 11, 1, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 11, 29, 12, 0, 0), lastSecondRuns[0], "Spring 倒数第 2 天");

            var lastFriday = CronToolEngine.Parse("0 0 12 ? * 5L", CronDialect.Spring);
            Assert(lastFriday.Success, "Spring 5L 应成功: " + lastFriday.Error);
            var lastFridayRuns = CronToolEngine.NextRuns(lastFriday, new DateTime(2026, 11, 1, 0, 0, 0), 1);
            AssertEq(new DateTime(2026, 11, 27, 12, 0, 0), lastFridayRuns[0], "Spring 最后一个周五");

            Assert(!CronToolEngine.Parse("0 0 12 ? * L", CronDialect.Spring).Success,
                "Spring 单独的 L 应报错并提示带星期前缀");
        }

        private static void Test_SpringMacros()
        {
            var daily = CronToolEngine.Parse("@daily", CronDialect.Spring);
            Assert(daily.Success, "@daily 应成功: " + daily.Error);
            AssertEq("@daily", daily.Macro, "应记录宏");
            AssertEq("0 0 0 * * *", daily.Normalized, "宏展开结果");

            var runs = CronToolEngine.NextRuns(daily, new DateTime(2026, 11, 21, 13, 0, 0), 1);
            AssertEq(new DateTime(2026, 11, 22, 0, 0, 0), runs[0], "@daily 次日零点");

            Assert(CronToolEngine.Parse("@hourly", CronDialect.Spring).Success, "@hourly 应成功");
            Assert(!CronToolEngine.Parse("@daily", CronDialect.Quartz).Success, "Quartz 不支持宏");
            Assert(!CronToolEngine.Parse("@nope", CronDialect.Spring).Success, "未知宏应失败");
        }

        private static void Test_SpringFieldCount()
        {
            Assert(!CronToolEngine.Parse("0 0 12 * * * 2026", CronDialect.Spring).Success,
                "Spring 7 段应失败");

            var fallback = CronToolEngine.Parse("0 12 * * *", CronDialect.Spring);
            Assert(fallback.Success, "Spring 5 段应回退标准而不是直接报错: " + fallback.Error);
            AssertEq(CronDialect.Standard, fallback.Dialect, "回退后规范为标准");
            Assert(!string.IsNullOrEmpty(fallback.Warning), "应给出回退提示");
        }

        // ==================== YAML ====================

        private static void Test_YamlToJsonTypes()
        {
            string yaml = "name: 史莱姆\nage: 3\nboss: true\nempty:\nquoted: \"001\"\nskills:\n  - 撞击\n  - 分裂\n";
            var result = YamlToolEngine.YamlToJson(yaml);
            Assert(result.Success, "YAML 转 JSON 应成功: " + result.Error);

            var root = JsonNode.Parse(result.Text) as JsonObject;
            Assert(root != null, "输出应为 JSON 对象");
            AssertEq("史莱姆", root!["name"]?.GetValue<string>(), "字符串");
            AssertEq(3L, root["age"]?.GetValue<long>(), "整数");
            AssertEq(true, root["boss"]?.GetValue<bool>(), "布尔");
            Assert(root.ContainsKey("empty") && root["empty"] == null, "空值应为 null");
            AssertEq("001", root["quoted"]?.GetValue<string>(), "引号数字应保持字符串");
            AssertEq(2, (root["skills"] as JsonArray)?.Count ?? 0, "序列长度");
        }

        private static void Test_JsonYamlRoundTrip()
        {
            string json = """{"name":"史莱姆","ids":[1,2],"boss":true,"note":"on"}""";
            var toYaml = YamlToolEngine.JsonToYaml(json);
            Assert(toYaml.Success, "JSON 转 YAML 应成功: " + toYaml.Error);
            Assert(toYaml.Text.Contains("name:"), "应输出映射");

            var back = YamlToolEngine.YamlToJson(toYaml.Text);
            Assert(back.Success, "YAML 回转 JSON 应成功: " + back.Error);

            var root = JsonNode.Parse(back.Text) as JsonObject;
            Assert(root != null, "回转结果应为对象");
            AssertEq("史莱姆", root!["name"]?.GetValue<string>(), "字符串往返");
            AssertEq(2, (root["ids"] as JsonArray)?.Count ?? 0, "数组往返");
            AssertEq(true, root["boss"]?.GetValue<bool>(), "布尔往返");
            AssertEq("on", root["note"]?.GetValue<string>(), "普通字符串不应被误判为布尔");
        }

        private static void Test_YamlErrors()
        {
            Assert(!YamlToolEngine.YamlToJson("a: [1, 2").Success, "不闭合的序列应报错");
            Assert(!YamlToolEngine.YamlToJson("---\na: 1\n---\nb: 2\n").Success, "多文档应明确拒绝");
            Assert(!YamlToolEngine.YamlToJson("   ").Success, "空内容应失败");
        }
    }
}
