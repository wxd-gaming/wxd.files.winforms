using System;

namespace wxd.files.winforms
{
    /// <summary>
    /// 内容过滤词的预编译结果。
    /// 原实现对每个单元格都做一次 "||"/"&&" 替换 + Split，3 万行 × 50 列就是上百万次
    /// 数组分配；现在按"一次输入、一个过滤词表"解析，逐单元格只做 IndexOf。
    /// 语义与旧的 MatchesFilter 保持一致：
    ///   含 "|" → 任一词命中（OR）；否则含 "&amp;" → 所有词命中（AND）；否则整串包含匹配。
    ///   大小写不敏感，均用 OrdinalIgnoreCase。
    /// </summary>
    internal readonly struct ContentFilterMatcher
    {
        private readonly string[]? _terms;
        private readonly char _separator; // '|' = OR，'&' = AND，'\0' = 单串包含

        private ContentFilterMatcher(string[] terms, char separator)
        {
            _terms = terms;
            _separator = separator;
        }

        /// <summary>解析过滤词；空串返回"全部命中"的空匹配器。</summary>
        internal static ContentFilterMatcher Parse(string? filter)
        {
            if (string.IsNullOrEmpty(filter)) return default;

            string processed = filter.Replace("||", " | ").Replace("&&", " & ");
            if (processed.Contains('|'))
            {
                var terms = processed.Split('|',
                    StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                return new ContentFilterMatcher(terms, '|');
            }
            if (processed.Contains('&'))
            {
                var terms = processed.Split('&',
                    StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                return new ContentFilterMatcher(terms, '&');
            }
            return new ContentFilterMatcher(new[] { processed }, '\0');
        }

        /// <summary>单元格文本是否命中该过滤词。</summary>
        internal bool Matches(string? text)
        {
            var terms = _terms;
            if (terms == null) return true;          // 空过滤 = 全部命中
            if (string.IsNullOrEmpty(text)) return false;

            switch (_separator)
            {
                case '|':
                    foreach (var term in terms)
                        if (text.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0) return true;
                    return false;

                case '&':
                    foreach (var term in terms)
                        if (text.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0) return false;
                    return true;

                default:
                    return terms.Length > 0
                        && text.IndexOf(terms[0], StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }
    }
}
