using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using YamlDotNet.Core;
using YamlDotNet.RepresentationModel;
using YamlDotNet.Serialization;

namespace wxd.files.winforms
{
    /// <summary>
    /// YAML ↔ JSON 转换。YAML 解析走 YamlDotNet 的表示模型，
    /// 普通标量按 null / bool / 整数 / 小数 / 字符串推断，引号或块状标量一律按字符串。
    /// </summary>
    internal static class YamlToolEngine
    {
        internal static TextToolResult YamlToJson(string yaml)
        {
            if (string.IsNullOrWhiteSpace(yaml))
                return TextToolResult.Fail("内容为空");

            try
            {
                var stream = new YamlStream();
                stream.Load(new StringReader(yaml));

                if (stream.Documents.Count == 0)
                    return TextToolResult.Fail("没有 YAML 文档");
                if (stream.Documents.Count > 1)
                    return TextToolResult.Fail("暂不支持多文档 YAML（--- 分隔）");

                var stack = new HashSet<YamlNode>();
                JsonNode? root = ConvertNode(stream.Documents[0].RootNode, stack);
                return TextToolResult.Ok(JsonToolEngine.Format(root));
            }
            catch (YamlException ex)
            {
                long line = ex.Start.Line;
                long column = ex.Start.Column;
                return TextToolResult.Fail($"YAML 语法错误（第 {line} 行，第 {column} 列）：{ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                return TextToolResult.Fail(ex.Message);
            }
        }

        internal static TextToolResult Validate(string yaml)
        {
            var converted = YamlToJson(yaml);
            if (!converted.Success) return converted;

            var parsed = JsonToolEngine.Parse(converted.Text);
            int count = parsed.Success ? JsonToolEngine.CountNodes(parsed.Root) : 0;
            return TextToolResult.Ok($"YAML 合法 · {count} 个节点");
        }

        internal static TextToolResult JsonToYaml(string json)
        {
            var parsed = JsonToolEngine.Parse(json);
            if (!parsed.Success)
                return TextToolResult.Fail(parsed.Error ?? "JSON 解析失败");

            object? graph = ToPlainObject(parsed.Root);
            try
            {
                var serializer = new SerializerBuilder()
                    .WithIndentedSequences()
                    .Build();
                return TextToolResult.Ok(serializer.Serialize(graph));
            }
            catch (YamlException ex)
            {
                return TextToolResult.Fail("YAML 生成失败：" + ex.Message);
            }
        }

        private static JsonNode? ConvertNode(YamlNode node, HashSet<YamlNode> stack)
        {
            if (!stack.Add(node))
                throw new InvalidOperationException("YAML 中存在循环引用（锚点 / 别名），无法转换为 JSON");
            try
            {
                switch (node)
                {
                    case YamlMappingNode mapping:
                    {
                        var obj = new JsonObject();
                        foreach (var pair in mapping.Children)
                        {
                            if (pair.Key is not YamlScalarNode key)
                                throw new InvalidOperationException("YAML 映射的 key 必须是标量，无法转换为 JSON");
                            obj[key.Value ?? ""] = ConvertNode(pair.Value, stack);
                        }
                        return obj;
                    }
                    case YamlSequenceNode sequence:
                    {
                        var array = new JsonArray();
                        foreach (var item in sequence.Children)
                            array.Add(ConvertNode(item, stack));
                        return array;
                    }
                    case YamlScalarNode scalar:
                        return ScalarToJson(scalar);
                    default:
                        return JsonValue.Create(node.ToString());
                }
            }
            finally
            {
                stack.Remove(node);
            }
        }

        private static JsonNode? ScalarToJson(YamlScalarNode scalar)
        {
            string? raw = scalar.Value;
            if (scalar.Style != ScalarStyle.Plain)
                return JsonValue.Create(raw ?? "");

            if (string.IsNullOrEmpty(raw))
                return null;

            string text = raw.Trim();
            switch (text.ToLowerInvariant())
            {
                case "null":
                case "~":
                    return null;
                case "true":
                    return JsonValue.Create(true);
                case "false":
                    return JsonValue.Create(false);
            }

            if (long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long integer))
                return JsonValue.Create(integer);

            if ((text.Contains('.') || text.Contains('e') || text.Contains('E'))
                && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                && double.IsFinite(number))
                return JsonValue.Create(number);

            return JsonValue.Create(raw);
        }

        private static object? ToPlainObject(JsonNode? node)
        {
            switch (node)
            {
                case null:
                    return null;
                case JsonObject obj:
                {
                    var map = new Dictionary<string, object?>();
                    foreach (var property in obj)
                        map[property.Key] = ToPlainObject(property.Value);
                    return map;
                }
                case JsonArray array:
                {
                    var list = new List<object?>();
                    foreach (var item in array)
                        list.Add(ToPlainObject(item));
                    return list;
                }
                case JsonValue value:
                {
                    if (value.TryGetValue<bool>(out bool b)) return b;
                    if (value.TryGetValue<long>(out long l)) return l;
                    if (value.TryGetValue<decimal>(out decimal dec)) return dec;
                    if (value.TryGetValue<double>(out double d)) return d;
                    if (value.TryGetValue<string>(out string? s)) return s;
                    return value.ToJsonString();
                }
                default:
                    return node.ToJsonString();
            }
        }
    }
}
