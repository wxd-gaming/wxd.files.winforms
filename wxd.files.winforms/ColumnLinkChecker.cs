using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelDataReader;

namespace wxd.files.winforms
{
    /// <summary>列关联颜色检查器 — 根据跨文件列关联规则为DataGridView单元格着色（支持多文件独立缓存）</summary>
    public static class ColumnLinkChecker
    {
        /// <summary>缓存 key: "{sourceRelativePath}|{sourceColumnIndex}->{targetRelativePath}:{targetColumnIndex}"</summary>
        private static Dictionary<string, HashSet<string>> _linkLookups = new();
        /// <summary>已加载的源文件路径集合（用于判断是否需要重建）</summary>
        private static HashSet<string> _loadedSources = new(StringComparer.OrdinalIgnoreCase);

        public static void Reset()
        {
            _linkLookups.Clear();
            _loadedSources.Clear();
        }

        /// <summary>构建关联查找缓存，读取目标文件指定列的所有值（不清除其他文件的缓存）</summary>
        public static void BuildLinkLookups(List<ColumnLinkConfig> columnLinks, string rootDir, string filePath)
        {
            if (columnLinks.Count == 0 || string.IsNullOrEmpty(rootDir)) return;

            // 生成当前文件的相对路径，用作缓存 key 前缀
            string sourceRel;
            if (filePath.StartsWith(rootDir, StringComparison.OrdinalIgnoreCase))
                sourceRel = filePath.Substring(rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            else
                sourceRel = Path.GetFileName(filePath);

            // 先清除当前文件已有的旧缓存条目
            string prefix = sourceRel + "|";
            var keysToRemove = _linkLookups.Keys.Where(k => k.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (var k in keysToRemove)
                _linkLookups.Remove(k);

            _loadedSources.Add(sourceRel);

            foreach (var link in columnLinks)
            {
                if (!link.SourceFile.EndsWith(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase))
                    continue;

                string targetFullPath = Path.Combine(rootDir, link.TargetFile);
                if (!File.Exists(targetFullPath)) continue;

                try
                {
                    var ext = Path.GetExtension(targetFullPath).ToLowerInvariant();
                    var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    if (ext == ".csv")
                    {
                        string contentData = File.ReadAllText(targetFullPath, Encoding.UTF8);
                        var lines = SplitCsvLines(contentData);
                        for (int i = 0; i < lines.Count; i++)
                        {
                            var parts = ParseCsvLine(lines[i], ',');
                            if (link.TargetColumnIndex >= 0 && link.TargetColumnIndex < parts.Length)
                            {
                                var val = parts[link.TargetColumnIndex].Trim();
                                if (!string.IsNullOrEmpty(val))
                                    values.Add(val);
                            }
                        }
                    }
                    else if (ext == ".xls" || ext == ".xlsx")
                    {
                        using var fs = new FileStream(targetFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        using var reader = ExcelReaderFactory.CreateReader(fs);
                        var result = reader.AsDataSet();
                        if (result.Tables.Count > 0)
                        {
                            var table = result.Tables[0];
                            foreach (DataRow row in table.Rows)
                            {
                                if (link.TargetColumnIndex >= 0 && link.TargetColumnIndex < table.Columns.Count)
                                {
                                    var val = row[link.TargetColumnIndex]?.ToString()?.Trim() ?? "";
                                    if (!string.IsNullOrEmpty(val))
                                        values.Add(val);
                                }
                            }
                        }
                    }

                    // key 格式: "{sourceRel}|{sourceColIndex}->{targetRel}:{targetColIndex}"
                    string key = string.Format("{0}|{1}->{2}:{3}", sourceRel, link.SourceColumnIndex, link.TargetFile, link.TargetColumnIndex);
                    _linkLookups[key] = values;
                }
                catch { }
            }
        }

        /// <summary>检查指定值在目标表中是否存在（用于右键菜单置灰判断）</summary>
        /// <returns>true=值有效可打开，false=值无效或不存在</returns>
        public static bool IsLinkValueValid(string filePath, string rootDir, int originalColumnIndex, string cellValue)
        {
            if (string.IsNullOrEmpty(cellValue) || _loadedSources.Count == 0) return false;
            
            string currentRel;
            if (!string.IsNullOrEmpty(rootDir) && filePath.StartsWith(rootDir, StringComparison.OrdinalIgnoreCase))
                currentRel = filePath.Substring(rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            else
                currentRel = Path.GetFileName(filePath);

            string prefix = currentRel + "|" + originalColumnIndex + "->";
            foreach (var kvp in _linkLookups)
            {
                if (kvp.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return kvp.Value.Contains(cellValue);
                }
            }
            return false;
        }

        /// <summary>检查单元格颜色 — 限匹配当前文件的源关联，不存在则标红，存在则恢复默认；跳过冻结行</summary>
        public static void CheckCellColor(DataGridViewCellFormattingEventArgs e,
            List<int> visibleColumnIndices, DataGridView dataGrid, string filePath, string rootDir, int frozenRows)
        {
            // 冻结行不标红
            if (e.RowIndex >= 0 && e.RowIndex < frozenRows) return;
            if (_loadedSources.Count == 0) return;

            // 计算当前文件的相对路径，用于过滤缓存条目
            string currentRel;
            if (!string.IsNullOrEmpty(rootDir) && filePath.StartsWith(rootDir, StringComparison.OrdinalIgnoreCase))
                currentRel = filePath.Substring(rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            else
                currentRel = Path.GetFileName(filePath);

            int originalColIndex = e.ColumnIndex;
            if (visibleColumnIndices.Count > 0 && e.ColumnIndex < dataGrid.Columns.Count)
            {
                int adjustedCol = e.ColumnIndex;
                if (dataGrid.Columns.Count > 0 && dataGrid.Columns[0].Name == "__RowNum__")
                    adjustedCol = e.ColumnIndex - 1;
                if (adjustedCol >= 0 && adjustedCol < visibleColumnIndices.Count)
                    originalColIndex = visibleColumnIndices[adjustedCol];
            }

            if (_linkLookups.Count == 0) return;

            foreach (var kvp in _linkLookups)
            {
                // key 格式: "{sourceRel}|{srcCol}->{targetRel}:{targetCol}"
                var pipeIdx = kvp.Key.IndexOf('|');
                if (pipeIdx < 0) continue;
                // 只匹配当前文件的关联，排除其他文件的缓存
                var keySourceRel = kvp.Key.Substring(0, pipeIdx);
                if (!keySourceRel.Equals(currentRel, StringComparison.OrdinalIgnoreCase)) continue;
                var afterPipe = kvp.Key.Substring(pipeIdx + 1); // "{srcCol}->{targetRel}:{targetCol}"
                var arrowIdx = afterPipe.IndexOf("->", StringComparison.Ordinal);
                if (arrowIdx < 0) continue;
                var colStr = afterPipe.Substring(0, arrowIdx);
                if (int.TryParse(colStr, out int srcCol) && srcCol == originalColIndex)
                {
                    var cellValue = e.Value?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(cellValue) && !kvp.Value.Contains(cellValue))
                    {
                        e.CellStyle.ForeColor = Color.Red;
                        e.CellStyle.SelectionForeColor = Color.Red;
                    }
                    else
                    {
                        e.CellStyle.ForeColor = dataGrid.DefaultCellStyle.ForeColor;
                        e.CellStyle.SelectionForeColor = dataGrid.DefaultCellStyle.SelectionForeColor;
                    }
                    break;
                }
            }
        }

        private static List<string> SplitCsvLines(string data)
        {
            var lines = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < data.Length; i++)
            {
                char c = data[i];
                if (c == '"') inQuotes = !inQuotes;
                else if (!inQuotes && (c == '\n' || c == '\r'))
                {
                    if (c == '\r' && i + 1 < data.Length && data[i + 1] == '\n') i++;
                    lines.Add(sb.ToString());
                    sb.Clear();
                }
                else sb.Append(c);
            }
            if (sb.Length > 0) lines.Add(sb.ToString());
            return lines;
        }

        private static string[] ParseCsvLine(string line, char delimiter)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                    else inQuotes = !inQuotes;
                }
                else if (c == delimiter && !inQuotes) { result.Add(sb.ToString().Trim()); sb.Clear(); }
                else sb.Append(c);
            }
            result.Add(sb.ToString().Trim());
            return result.ToArray();
        }
    }
}
