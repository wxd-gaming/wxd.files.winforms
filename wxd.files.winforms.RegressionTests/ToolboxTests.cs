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
            Run("时间戳 非法输入", Test_TimeInvalid);

            // 编码
            Run("编码 Unicode 转义/还原", Test_UnicodeEscapeRoundTrip);
            Run("编码 UTF-8 Hex 往返", Test_Utf8HexRoundTrip);
            Run("编码 Base64 往返", Test_Base64RoundTrip);
            Run("编码 URL / HTML 往返", Test_UrlAndHtmlRoundTrip);

            // 哈希
            Run("哈希 MD5/SHA256 标准向量", Test_HashVectors);

            // Cron
            Run("Cron 工作日每5分钟", Test_CronWeekdayStep);
            Run("Cron 6段含秒", Test_CronWithSeconds);
            Run("Cron 非法表达式", Test_CronInvalid);
            Run("Cron 日/周 OR 语义", Test_CronDayOrWeekday);

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
