using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace wxd.files.winforms
{
    /// <summary>JSON 解析结果：成功时 Root 可能为 null（顶层是字面量 null）。</summary>
    internal readonly struct JsonParseResult
    {
        public bool Success { get; init; }
        public JsonNode? Root { get; init; }
        public string? Error { get; init; }

        /// <summary>严格解析失败、靠给非字符串 key 补引号才解析成功。</summary>
        public bool NormalizedKeys { get; init; }
    }

    /// <summary>
    /// JSON 面板的纯逻辑：解析、格式化、压缩、key 排序、节点计数。
    /// 一律用 System.Text.Json，保留数字原始写法；格式化/压缩都不转义中文。
    /// </summary>
    internal static class JsonToolEngine
    {
        private static readonly JsonDocumentOptions ParseOptions = new()
        {
            AllowTrailingCommas = true,
            CommentHandling = JsonCommentHandling.Skip
        };

        private static readonly JsonSerializerOptions FormatOptions = new()
        {
            WriteIndented = true,
            IndentSize = 2,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        private static readonly JsonSerializerOptions MinifyOptions = new()
        {
            WriteIndented = false,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        internal static JsonParseResult Parse(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new JsonParseResult { Success = false, Error = "内容为空" };

            try
            {
                var root = JsonNode.Parse(text, new JsonNodeOptions(), ParseOptions);
                return new JsonParseResult { Success = true, Root = root };
            }
            catch (JsonException strictError)
            {
                // fastjson 序列化 Map<Integer,?> 默认输出 {1:"a"} 这类非标准 key，
                // 严格解析必然失败；补上引号后重试，避免用户看到"语法错误"就以为文件坏了。
                string normalized = QuoteNonStringKeys(text, out bool changed);
                if (!changed) return Failure(strictError);

                try
                {
                    var root = JsonNode.Parse(normalized, new JsonNodeOptions(), ParseOptions);
                    return new JsonParseResult { Success = true, Root = root, NormalizedKeys = true };
                }
                catch (JsonException)
                {
                    // 补引号后仍失败：报原始文本的错误，行列号才对得上用户看到的源码
                    return Failure(strictError);
                }
            }
        }

        private static JsonParseResult Failure(JsonException ex)
        {
            long line = (ex.LineNumber ?? 0) + 1;
            long column = (ex.BytePositionInLine ?? 0) + 1;
            return new JsonParseResult
            {
                Success = false,
                Error = $"JSON 语法错误（第 {line} 行，第 {column} 列）：{ex.Message}"
            };
        }

        /// <summary>
        /// 给对象里不带引号的 key 补上引号（数字、true/false 等），其余字符原样保留。
        /// 按"当前是否在 key 位置"扫描，字符串内容、数组元素、注释不会被误改。
        /// 返回 changed=false 表示没有可补的 key，调用方应报告原始解析错误。
        /// </summary>
        internal static string QuoteNonStringKeys(string text, out bool changed)
        {
            changed = false;
            var output = new StringBuilder(text.Length + 16);
            var openFrames = new List<bool>();   // 每层容器是否为对象
            bool inObject = false;
            int state = StateValueOrEnd;
            int i = 0;

            while (i < text.Length)
            {
                char c = text[i];

                if (c == '/' && i + 1 < text.Length && (text[i + 1] == '/' || text[i + 1] == '*'))
                {
                    i = CopyComment(text, i, output);
                    continue;
                }

                if (inObject && state == StateKeyOrEnd)
                {
                    if (char.IsWhiteSpace(c) || c == ',' || c == '}')
                    {
                        output.Append(c);
                        i++;
                        if (c == '}') PopFrame(openFrames, ref inObject, ref state);
                        continue;
                    }
                    if (c == '"')
                    {
                        i = CopyString(text, i, output);
                        state = StateAfterKey;
                        continue;
                    }

                    int start = i;
                    while (i < text.Length && text[i] != ':' && text[i] != '}' && text[i] != ','
                           && !char.IsWhiteSpace(text[i]))
                        i++;
                    if (i == start)
                    {
                        output.Append(c);
                        i++;
                        continue;
                    }

                    output.Append('"');
                    for (int k = start; k < i; k++)
                    {
                        if (text[k] == '"' || text[k] == '\\') output.Append('\\');
                        output.Append(text[k]);
                    }
                    output.Append('"');
                    changed = true;
                    state = StateAfterKey;
                    continue;
                }

                switch (state)
                {
                    case StateAfterKey:
                        output.Append(c);
                        i++;
                        if (c == ':') state = StateValueOrEnd;
                        continue;

                    case StateValueOrEnd:
                        if (char.IsWhiteSpace(c))
                        {
                            output.Append(c);
                            i++;
                            continue;
                        }
                        if (c == '{')
                        {
                            openFrames.Add(inObject);
                            inObject = true;
                            state = StateKeyOrEnd;
                            output.Append(c);
                            i++;
                            continue;
                        }
                        if (c == '[')
                        {
                            openFrames.Add(inObject);
                            inObject = false;
                            state = StateValueOrEnd;
                            output.Append(c);
                            i++;
                            continue;
                        }
                        if (c == ']' || c == '}')
                        {
                            output.Append(c);
                            i++;
                            PopFrame(openFrames, ref inObject, ref state);
                            continue;
                        }
                        if (c == '"')
                        {
                            i = CopyString(text, i, output);
                            state = StateAfterValue;
                            continue;
                        }

                        // 字面量 / 数字：原样复制到分隔符
                        int literalStart = i;
                        while (i < text.Length && text[i] != ':' && text[i] != ',' && text[i] != ']'
                               && text[i] != '}' && !char.IsWhiteSpace(text[i]))
                            i++;
                        if (i == literalStart)
                        {
                            output.Append(c);
                            i++;
                            continue;
                        }
                        output.Append(text, literalStart, i - literalStart);
                        state = StateAfterValue;
                        continue;

                    default: // StateAfterValue
                        if (char.IsWhiteSpace(c))
                        {
                            output.Append(c);
                            i++;
                            continue;
                        }
                        if (c == ',')
                        {
                            output.Append(c);
                            i++;
                            state = inObject ? StateKeyOrEnd : StateValueOrEnd;
                            continue;
                        }
                        if (c == '}' || c == ']')
                        {
                            output.Append(c);
                            i++;
                            PopFrame(openFrames, ref inObject, ref state);
                            continue;
                        }
                        output.Append(c);
                        i++;
                        continue;
                }
            }

            return output.ToString();
        }

        private const int StateKeyOrEnd = 0;   // 对象内：期待 key 或 '}'
        private const int StateAfterKey = 1;   // 已读 key：期待 ':'
        private const int StateValueOrEnd = 2; // 期待值（数组或顶层还允许 ']' / 结束）
        private const int StateAfterValue = 3; // 已读值：期待 ',' 或容器结束

        private static void PopFrame(List<bool> openFrames, ref bool inObject, ref int state)
        {
            state = StateAfterValue;
            if (openFrames.Count == 0)
            {
                inObject = false;
                return;
            }
            inObject = openFrames[openFrames.Count - 1];
            openFrames.RemoveAt(openFrames.Count - 1);
        }

        private static int CopyString(string text, int start, StringBuilder output)
        {
            int i = start + 1;
            while (i < text.Length)
            {
                char c = text[i];
                i++;
                if (c == '\\' && i < text.Length) { i++; continue; }
                if (c == '"') break;
            }
            output.Append(text, start, i - start);
            return i;
        }

        private static int CopyComment(string text, int start, StringBuilder output)
        {
            int end;
            if (text[start + 1] == '/')
            {
                end = text.IndexOf('\n', start);
                if (end < 0) end = text.Length;
            }
            else
            {
                int close = text.IndexOf("*/", start + 2, StringComparison.Ordinal);
                end = close < 0 ? text.Length : close + 2;
            }
            output.Append(text, start, end - start);
            return end;
        }

        /// <summary>缩进 2 空格输出；sortKeys=true 时先递归排序 key。</summary>
        internal static string Format(JsonNode? node, bool sortKeys = false)
        {
            if (node == null) return "null";
            var target = sortKeys ? SortKeys(node) : node.DeepClone();
            return target?.ToJsonString(FormatOptions) ?? "null";
        }

        /// <summary>压缩成一行。</summary>
        internal static string Minify(JsonNode? node)
            => node == null ? "null" : node.ToJsonString(MinifyOptions);

        /// <summary>递归按 key 的序数（区分大小写）升序重排对象；数组顺序保持不变。</summary>
        internal static JsonNode? SortKeys(JsonNode? node)
        {
            switch (node)
            {
                case JsonObject obj:
                {
                    var sorted = new JsonObject();
                    foreach (var property in obj.OrderBy(p => p.Key, StringComparer.Ordinal))
                        sorted[property.Key] = property.Value == null ? null : SortKeys(property.Value);
                    return sorted;
                }
                case JsonArray array:
                {
                    var copy = new JsonArray();
                    foreach (var item in array)
                        copy.Add(item == null ? null : SortKeys(item));
                    return copy;
                }
                default:
                    return node?.DeepClone();
            }
        }

        /// <summary>
        /// 检测文本里是否有 // 或 /* 注释（字符串内部的 // 不算）。
        /// JsonNode 序列化会丢掉注释，所以在"自动格式化/格式化/排序"之前必须先问一句。
        /// </summary>
        internal static bool ContainsComments(string? text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            bool inString = false;
            bool escaped = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (inString)
                {
                    if (escaped) escaped = false;
                    else if (c == '\\') escaped = true;
                    else if (c == '"') inString = false;
                    continue;
                }

                if (c == '"') { inString = true; continue; }
                if (c == '/' && i + 1 < text.Length && (text[i + 1] == '/' || text[i + 1] == '*'))
                    return true;
            }
            return false;
        }

        internal static int CountNodes(JsonNode? node)
        {
            switch (node)
            {
                case JsonObject obj:
                {
                    int count = 1;
                    foreach (var property in obj)
                        count += CountNodes(property.Value);
                    return count;
                }
                case JsonArray array:
                {
                    int count = 1;
                    foreach (var item in array)
                        count += CountNodes(item);
                    return count;
                }
                default:
                    return 1;
            }
        }
    }
}
