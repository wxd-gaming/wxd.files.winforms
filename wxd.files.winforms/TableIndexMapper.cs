using System;
using System.Collections.Generic;

namespace wxd.files.winforms
{
    /// <summary>
    /// 「原始索引 → 显示索引」反查表。
    /// 过滤/排序完成后需要把匹配单元格的原始行列翻译成显示行列；
    /// 逐条调用 List.IndexOf 是 O(n²)，大表上一次过滤就是上亿次比较，
    /// 这里统一改成"一次建表、O(1) 反查"。
    /// </summary>
    internal static class TableIndexMapper
    {
        /// <summary>
        /// 把显示顺序列表反转成映射表：map[原始索引] = 显示索引。
        /// 未出现在列表里的原始索引取值为 -1，调用方据此跳过。
        /// </summary>
        internal static int[] BuildReverseMap(IReadOnlyList<int>? indices, int capacity)
        {
            var map = new int[Math.Max(0, capacity)];
            Array.Fill(map, -1);
            if (indices == null) return map;

            for (int i = 0; i < indices.Count; i++)
            {
                int original = indices[i];
                if ((uint)original < (uint)map.Length)
                    map[original] = i;
            }
            return map;
        }

        /// <summary>
        /// 把匹配单元格（原始行列索引）翻译成显示行列索引，分成冻结行区与内容区两本字典。
        /// </summary>
        /// <param name="visibleRowIndices">
        /// 内容行的显示顺序；传 null 表示"原始行号即显示行号"（无过滤、无排序、无冻结行）。
        /// </param>
        /// <param name="frozenRows">前 N 行固定在独立网格里，行号即原始行号。</param>
        /// <param name="columnCount">源表列数（反查表容量）。</param>
        /// <param name="rowCount">源表行数（反查表容量）。</param>
        internal static (Dictionary<int, HashSet<int>> HeaderCells, Dictionary<int, HashSet<int>> BodyCells)
            MapMatchCellsToDisplay(
                IReadOnlyDictionary<int, HashSet<int>>? merged,
                IReadOnlyList<int> visibleColumnIndices,
                IReadOnlyList<int>? visibleRowIndices,
                int frozenRows,
                int columnCount,
                int rowCount)
        {
            var header = new Dictionary<int, HashSet<int>>();
            var body = new Dictionary<int, HashSet<int>>();
            if (merged == null || merged.Count == 0) return (header, body);

            var columnMap = BuildReverseMap(visibleColumnIndices, columnCount);
            var rowMap = visibleRowIndices == null ? null : BuildReverseMap(visibleRowIndices, rowCount);

            foreach (var pair in merged)
            {
                HashSet<int>? displayCols = null;
                foreach (var originalCol in pair.Value)
                {
                    if ((uint)originalCol >= (uint)columnMap.Length) continue;
                    int displayCol = columnMap[originalCol];
                    if (displayCol < 0) continue;
                    (displayCols ??= new HashSet<int>()).Add(displayCol);
                }
                if (displayCols == null) continue;

                int originalRow = pair.Key;
                if (originalRow < frozenRows)
                {
                    header[originalRow] = displayCols;
                    continue;
                }

                int displayRow;
                if (rowMap == null)
                {
                    if ((uint)originalRow >= (uint)rowCount) continue;
                    displayRow = originalRow;
                }
                else
                {
                    if ((uint)originalRow >= (uint)rowMap.Length) continue;
                    displayRow = rowMap[originalRow];
                    if (displayRow < 0) continue;
                }
                body[displayRow] = displayCols;
            }

            return (header, body);
        }
    }
}
