using System;
using System.Collections.Generic;

namespace wxd.files.winforms
{
    /// <summary>文本比对后一行的差异类型。</summary>
    internal enum TextDiffKind
    {
        Same,
        Modified,
        Added,
        Deleted
    }

    /// <summary>按行对齐后的一行：左右各挂一条文本，缺失一侧为 null（新增/删除）。</summary>
    internal sealed class TextDiffLine
    {
        /// <summary>左侧行号（1-based）；左侧没有这一行时为 null。</summary>
        public int? LeftLineNumber { get; init; }

        public string LeftText { get; init; } = "";

        /// <summary>右侧行号（1-based）；右侧没有这一行时为 null。</summary>
        public int? RightLineNumber { get; init; }

        public string RightText { get; init; } = "";

        public TextDiffKind Kind { get; init; }
    }

    /// <summary>
    /// 纯逻辑的按行文本比对。
    ///
    /// 做法：先裁掉公共前缀/后缀（配置表/日志通常只有中间一小块变化，裁完 Myers 只需处理差异段），
    /// 中间段跑 Myers O((N+M)·D)；差异步数超过 <see cref="MaxEditDistance"/> 时
    /// （两个完全不相干的大文件）退化为"整块替换"，保证不会长时间占着后台线程。
    ///
    /// 输出是"按行对齐"的列表，左右网格用同一份列表，行数天然一致，滚动/筛选不会错位。
    /// </summary>
    internal static class TextDiffEngine
    {
        /// <summary>Myers 允许的最大编辑步数；超过就把差异段当成一整块替换。</summary>
        internal const int MaxEditDistance = 1000;

        /// <summary>比对两份按行切好的文本，返回按行对齐的结果。</summary>
        internal static List<TextDiffLine> Compare(IReadOnlyList<string> left, IReadOnlyList<string> right)
        {
            if (left == null) throw new ArgumentNullException(nameof(left));
            if (right == null) throw new ArgumentNullException(nameof(right));

            int n = left.Count;
            int m = right.Count;
            var result = new List<TextDiffLine>(Math.Max(n, m));

            int prefix = 0;
            while (prefix < n && prefix < m && string.Equals(left[prefix], right[prefix], StringComparison.Ordinal))
                prefix++;

            int suffix = 0;
            while (suffix < n - prefix && suffix < m - prefix
                   && string.Equals(left[n - 1 - suffix], right[m - 1 - suffix], StringComparison.Ordinal))
                suffix++;

            for (int i = 0; i < prefix; i++)
                result.Add(Same(i + 1, left[i], i + 1, right[i]));

            int leftEnd = n - suffix;   // 左中段 [prefix, leftEnd)
            int rightEnd = m - suffix;  // 右中段 [prefix, rightEnd)

            if (leftEnd > prefix || rightEnd > prefix)
            {
                var matches = leftEnd > prefix && rightEnd > prefix
                    ? MyersMatches(left, prefix, leftEnd, right, prefix, rightEnd)
                    : new List<(int A, int B)>();

                if (matches == null)
                {
                    // 差异太大：整段按"左删 + 右增"处理（前 min 条配对成修改行）
                    AppendChangedBlock(result, left, prefix, leftEnd, right, prefix, rightEnd);
                }
                else
                {
                    int ai = prefix;
                    int bi = prefix;
                    foreach (var (matchA, matchB) in matches)
                    {
                        if (matchA > ai || matchB > bi)
                            AppendChangedBlock(result, left, ai, matchA, right, bi, matchB);

                        result.Add(Same(matchA + 1, left[matchA], matchB + 1, right[matchB]));
                        ai = matchA + 1;
                        bi = matchB + 1;
                    }
                    if (ai < leftEnd || bi < rightEnd)
                        AppendChangedBlock(result, left, ai, leftEnd, right, bi, rightEnd);
                }
            }

            for (int i = 0; i < suffix; i++)
            {
                int li = leftEnd + i;
                int ri = rightEnd + i;
                result.Add(Same(li + 1, left[li], ri + 1, right[ri]));
            }

            return result;
        }

        /// <summary>只保留有差异的行（"只保留差异行"开关用）。</summary>
        internal static List<TextDiffLine> OnlyDifferences(IReadOnlyList<TextDiffLine> lines)
        {
            var result = new List<TextDiffLine>();
            foreach (var line in lines)
                if (line.Kind != TextDiffKind.Same) result.Add(line);
            return result;
        }

        private static TextDiffLine Same(int leftNumber, string leftText, int rightNumber, string rightText)
            => new()
            {
                LeftLineNumber = leftNumber,
                LeftText = leftText,
                RightLineNumber = rightNumber,
                RightText = rightText,
                Kind = TextDiffKind.Same
            };

        /// <summary>
        /// 把一段"左侧被替换成右侧"的行配对：前 min(删,增) 条算修改，多出来的算删除/新增。
        /// 配对从段首开始，和文本编辑器逐行替换的直觉一致。
        /// </summary>
        private static void AppendChangedBlock(
            List<TextDiffLine> result,
            IReadOnlyList<string> left, int leftStart, int leftEnd,
            IReadOnlyList<string> right, int rightStart, int rightEnd)
        {
            int deleteCount = leftEnd - leftStart;
            int addCount = rightEnd - rightStart;
            int pairCount = Math.Min(deleteCount, addCount);

            for (int i = 0; i < pairCount; i++)
            {
                result.Add(new TextDiffLine
                {
                    LeftLineNumber = leftStart + i + 1,
                    LeftText = left[leftStart + i],
                    RightLineNumber = rightStart + i + 1,
                    RightText = right[rightStart + i],
                    Kind = TextDiffKind.Modified
                });
            }

            for (int i = pairCount; i < deleteCount; i++)
            {
                result.Add(new TextDiffLine
                {
                    LeftLineNumber = leftStart + i + 1,
                    LeftText = left[leftStart + i],
                    Kind = TextDiffKind.Deleted
                });
            }

            for (int i = pairCount; i < addCount; i++)
            {
                result.Add(new TextDiffLine
                {
                    RightLineNumber = rightStart + i + 1,
                    RightText = right[rightStart + i],
                    Kind = TextDiffKind.Added
                });
            }
        }

        /// <summary>
        /// Myers 差分，返回按顺序排列的匹配对（左索引, 右索引）。
        /// 返回 null 表示编辑步数超过 <see cref="MaxEditDistance"/>，调用方退化处理。
        /// 内存：只按 2·MaxEditDistance 分配 V，每步留一份快照，最坏约 8MB。
        /// </summary>
        private static List<(int A, int B)>? MyersMatches(
            IReadOnlyList<string> a, int aStart, int aEnd,
            IReadOnlyList<string> b, int bStart, int bEnd)
        {
            int n = aEnd - aStart;
            int m = bEnd - bStart;
            if (n == 0 || m == 0) return new List<(int A, int B)>();

            int maxD = Math.Min(n + m, MaxEditDistance);
            int offset = maxD + 1;
            var v = new int[2 * maxD + 3];
            var trace = new List<int[]>(maxD + 1);

            for (int d = 0; d <= maxD; d++)
            {
                trace.Add((int[])v.Clone());
                for (int k = -d; k <= d; k += 2)
                {
                    int index = offset + k;
                    int x = (k == -d || (k != d && v[index - 1] < v[index + 1]))
                        ? v[index + 1]
                        : v[index - 1] + 1;
                    int y = x - k;

                    while (x < n && y < m
                           && string.Equals(a[aStart + x], b[bStart + y], StringComparison.Ordinal))
                    {
                        x++;
                        y++;
                    }

                    v[index] = x;
                    if (x >= n && y >= m)
                        return Backtrack(trace, aStart, bStart, n, m, d, offset);
                }
            }
            return null;
        }

        /// <summary>按 Myers 的 trace 反向还原匹配对（升序）。</summary>
        private static List<(int A, int B)> Backtrack(
            List<int[]> trace, int aStart, int bStart, int n, int m, int endD, int offset)
        {
            var matches = new List<(int A, int B)>();
            int x = n;
            int y = m;

            for (int d = endD; d >= 0; d--)
            {
                var v = trace[d];
                int k = x - y;
                int prevK = (k == -d || (k != d && v[offset + k - 1] < v[offset + k + 1]))
                    ? k + 1
                    : k - 1;

                int prevX = v[offset + prevK];
                int prevY = prevX - prevK;

                while (x > prevX && y > prevY)
                {
                    x--;
                    y--;
                    matches.Add((aStart + x, bStart + y));
                }

                if (d > 0)
                {
                    x = prevX;
                    y = prevY;
                }
            }

            matches.Reverse();
            return matches;
        }
    }
}
