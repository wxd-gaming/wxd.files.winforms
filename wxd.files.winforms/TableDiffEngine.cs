using System;
using System.Collections.Generic;

namespace wxd.files.winforms
{
    /// <summary>表格比对一行的差异类型。</summary>
    internal enum TableRowDiffKind
    {
        Same,
        Modified,
        Added,
        Deleted
    }

    /// <summary>按行号对齐后的一行：左右各挂一条原始行，缺一侧表示新增/删除。</summary>
    internal sealed class TableDiffRow
    {
        /// <summary>左侧行号（1-based）；左侧没有这一行时为 null。</summary>
        public int? LeftRowNumber { get; init; }

        public string[] LeftCells { get; init; } = Array.Empty<string>();

        /// <summary>右侧行号（1-based）；右侧没有这一行时为 null。</summary>
        public int? RightRowNumber { get; init; }

        public string[] RightCells { get; init; } = Array.Empty<string>();

        public TableRowDiffKind Kind { get; init; }

        /// <summary>同一行内发生变化的列索引（仅 Modified 有意义）。</summary>
        public HashSet<int> DiffColumns { get; init; } = new();
    }

    /// <summary>
    /// 简单的表格比对：按行号一一对齐，逐单元格比较。
    ///
    /// 刻意不做"按唯一列对齐"——那是阅读器/版本对比的语义；
    /// 这里就是"第 N 行对第 N 行"，中间插入一行会让后面整片显示为修改，
    /// 与 Excel 里两列并排肉眼比对的直觉一致。
    ///
    /// 行数不一致时，多出来的一侧记为新增/删除，另一侧用空行占位，保证左右网格行数相同。
    /// </summary>
    internal static class TableDiffEngine
    {
        internal static List<TableDiffRow> Compare(IReadOnlyList<string[]> left, IReadOnlyList<string[]> right)
        {
            if (left == null) throw new ArgumentNullException(nameof(left));
            if (right == null) throw new ArgumentNullException(nameof(right));

            int count = Math.Max(left.Count, right.Count);
            var result = new List<TableDiffRow>(count);

            for (int i = 0; i < count; i++)
            {
                bool hasLeft = i < left.Count;
                bool hasRight = i < right.Count;

                if (!hasLeft)
                {
                    result.Add(new TableDiffRow
                    {
                        RightRowNumber = i + 1,
                        RightCells = right[i],
                        Kind = TableRowDiffKind.Added
                    });
                    continue;
                }

                if (!hasRight)
                {
                    result.Add(new TableDiffRow
                    {
                        LeftRowNumber = i + 1,
                        LeftCells = left[i],
                        Kind = TableRowDiffKind.Deleted
                    });
                    continue;
                }

                var leftCells = left[i];
                var rightCells = right[i];
                int columnCount = Math.Max(leftCells.Length, rightCells.Length);
                HashSet<int>? diffColumns = null;

                for (int c = 0; c < columnCount; c++)
                {
                    string leftValue = c < leftCells.Length ? leftCells[c] : "";
                    string rightValue = c < rightCells.Length ? rightCells[c] : "";
                    if (!string.Equals(leftValue, rightValue, StringComparison.Ordinal))
                        (diffColumns ??= new HashSet<int>()).Add(c);
                }

                result.Add(new TableDiffRow
                {
                    LeftRowNumber = i + 1,
                    LeftCells = leftCells,
                    RightRowNumber = i + 1,
                    RightCells = rightCells,
                    Kind = diffColumns == null ? TableRowDiffKind.Same : TableRowDiffKind.Modified,
                    DiffColumns = diffColumns ?? new HashSet<int>()
                });
            }

            return result;
        }

        /// <summary>只保留有差异的行（"只保留差异行"开关用）。</summary>
        internal static List<TableDiffRow> OnlyDifferences(IReadOnlyList<TableDiffRow> rows)
        {
            var result = new List<TableDiffRow>();
            foreach (var row in rows)
                if (row.Kind != TableRowDiffKind.Same) result.Add(row);
            return result;
        }

        /// <summary>两边所有行的最大列数，用来决定网格要建多少列。</summary>
        internal static int MaxColumnCount(IReadOnlyList<TableDiffRow> rows)
        {
            int max = 0;
            foreach (var row in rows)
            {
                if (row.LeftCells.Length > max) max = row.LeftCells.Length;
                if (row.RightCells.Length > max) max = row.RightCells.Length;
            }
            return max;
        }
    }
}
