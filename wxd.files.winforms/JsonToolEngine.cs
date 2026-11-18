using System;
using System.Collections.Generic;
using System.Linq;
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
            catch (JsonException ex)
            {
                long line = (ex.LineNumber ?? 0) + 1;
                long column = (ex.BytePositionInLine ?? 0) + 1;
                return new JsonParseResult
                {
                    Success = false,
                    Error = $"JSON 语法错误（第 {line} 行，第 {column} 列）：{ex.Message}"
                };
            }
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
