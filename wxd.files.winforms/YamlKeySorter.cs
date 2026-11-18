using System;
using System.Collections.Generic;
using System.Linq;

namespace wxd.files.winforms
{
    /// <summary>
    /// 保守的 YAML key 排序：只处理"块状映射"（key: value / 嵌套映射），
    /// 序列顺序、块标量（| / &gt;）内容保持原样；遇到锚点/别名/多文档/无法确定的结构直接放弃。
    ///
    /// 用途是"比对前把两边 key 顺序归一"，不写回文件，所以宁可放弃也不猜。
    /// </summary>
    internal static class YamlKeySorter
    {
        internal static bool TrySort(string text, out string sorted)
        {
            sorted = text;
            if (string.IsNullOrWhiteSpace(text)) return false;

            var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = normalized.Split('\n');

            foreach (var line in lines)
            {
                var trimmed = line.TrimStart();
                if (trimmed == "---" || trimmed == "..." || trimmed.StartsWith("--- ", StringComparison.Ordinal))
                    return false;   // 多文档/文档标记：不排
                if (line.Length > 0 && line[0] == '\t')
                    return false;   // YAML 缩进不允许 tab，结构不可信
                if (ContainsAnchorOrAlias(line))
                    return false;   // 锚点/别名重排可能改变语义
            }

            int probe = 0;
            while (probe < lines.Length && IsBlankOrComment(lines[probe])) probe++;
            if (probe >= lines.Length) return false;
            int rootIndent = IndentOf(lines[probe]);

            int index = 0;
            var body = ParseBlock(lines, ref index, rootIndent);
            if (body == null) return false;
            if (index < lines.Length)
            {
                // 根层没消费完：结构超出支持范围
                return false;
            }

            var output = new List<string>(body.Count);
            output.AddRange(body);
            sorted = string.Join("\n", output);
            if (normalized.EndsWith("\n", StringComparison.Ordinal))
                sorted += "\n";
            return true;
        }

        private static List<string>? ParseBlock(string[] lines, ref int index, int indent)
        {
            var entries = new List<(string Key, List<string> Lines)>();
            var pending = new List<string>();

            while (index < lines.Length)
            {
                var line = lines[index];
                if (IsBlankOrComment(line))
                {
                    pending.Add(line);
                    index++;
                    continue;
                }

                int lineIndent = IndentOf(line);
                if (lineIndent < indent) break;
                if (lineIndent > indent) return null;

                var trimmed = line.TrimStart();
                if (trimmed == "-" || trimmed.StartsWith("- ", StringComparison.Ordinal))
                    return null;    // 根/映射层是序列：不排

                if (!TrySplitKey(trimmed, out string key, out string rest))
                    return null;

                var entryLines = new List<string>();
                entryLines.AddRange(pending);
                pending.Clear();
                entryLines.Add(line);
                index++;

                if (rest.Length == 0 || rest.StartsWith("#", StringComparison.Ordinal))
                {
                    // 值在下一层：探测子块缩进
                    int probe = index;
                    while (probe < lines.Length && IsBlankOrComment(lines[probe])) probe++;

                    if (probe < lines.Length && IndentOf(lines[probe]) > indent)
                    {
                        int childIndent = IndentOf(lines[probe]);
                        var childTrimmed = lines[probe].TrimStart();

                        if (childTrimmed == "-" || childTrimmed.StartsWith("- ", StringComparison.Ordinal))
                        {
                            // 序列：原样保留（顺序有语义）
                            while (index < lines.Length
                                   && (IsBlankOrComment(lines[index]) || IndentOf(lines[index]) >= childIndent))
                                entryLines.Add(lines[index++]);
                        }
                        else
                        {
                            var child = ParseBlock(lines, ref index, childIndent);
                            if (child == null) return null;
                            entryLines.AddRange(child);
                        }
                    }
                }
                else
                {
                    bool blockScalar = rest.StartsWith("|", StringComparison.Ordinal)
                                    || rest.StartsWith(">", StringComparison.Ordinal);

                    if (blockScalar)
                    {
                        // 块标量：缩进更深的行（包括空行、# 开头的行）都属于字符串内容
                        while (index < lines.Length
                               && (lines[index].Trim().Length == 0 || IndentOf(lines[index]) > indent))
                            entryLines.Add(lines[index++]);
                    }
                    else
                    {
                        // 普通标量/流式集合：只带缩进更深的续行，同级的注释/空行留给下一个 key
                        while (index < lines.Length
                               && !IsBlankOrComment(lines[index])
                               && IndentOf(lines[index]) > indent)
                            entryLines.Add(lines[index++]);
                    }
                }

                entries.Add((key, entryLines));
            }

            if (entries.Count == 0)
            {
                // 只有注释/空行
                return new List<string>(pending);
            }

            var ordered = entries
                .OrderBy(e => e.Key, StringComparer.Ordinal)
                .ToList();

            var output = new List<string>();
            foreach (var entry in ordered)
                output.AddRange(entry.Lines);
            output.AddRange(pending);
            return output;
        }

        private static bool TrySplitKey(string line, out string key, out string rest)
        {
            key = "";
            rest = "";

            if (line.Length == 0) return false;

            int colon;
            if (line[0] == '"' || line[0] == '\'')
            {
                char quote = line[0];
                int end = line.IndexOf(quote, 1);
                if (end <= 1) return false;
                colon = line.IndexOf(':', end + 1);
                if (colon < 0) return false;
                key = line.Substring(1, end - 1);
            }
            else
            {
                colon = line.IndexOf(':');
                if (colon <= 0) return false;
                key = line.Substring(0, colon).Trim();
            }

            // YAML 的 "key:" 后面必须是空白/行尾/注释，否则整行是标量
            if (colon + 1 < line.Length)
            {
                char after = line[colon + 1];
                if (!char.IsWhiteSpace(after) && after != '#') return false;
            }

            rest = line[(colon + 1)..].Trim();
            return key.Length > 0;
        }

        private static int IndentOf(string line)
        {
            int indent = 0;
            while (indent < line.Length && line[indent] == ' ') indent++;
            return indent;
        }

        private static bool IsBlankOrComment(string line)
        {
            var trimmed = line.TrimStart();
            return trimmed.Length == 0 || trimmed[0] == '#';
        }

        private static bool ContainsAnchorOrAlias(string line)
        {
            for (int i = 0; i < line.Length; i++)
            {
                if ((line[i] == '&' || line[i] == '*') && (i == 0 || char.IsWhiteSpace(line[i - 1])))
                    return true;
            }
            return false;
        }
    }
}
