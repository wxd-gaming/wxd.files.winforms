using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>关联查找条目（包含ID集合和ID→显示名映射）</summary>
    public class LinkLookupEntry
    {
        /// <summary>关联列的值集合（用于快速查找）</summary>
        public HashSet<string> Values { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>ID → 显示名映射（当配置了显示列时填充）</summary>
        public Dictionary<string, string> DisplayNames { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>目标文件的相对路径</summary>
        public string TargetFile { get; set; } = "";
        /// <summary>目标列的表头名称</summary>
        public string TargetColumnName { get; set; } = "";
        /// <summary>显示列的表头名称</summary>
        public string DisplayColumnName { get; set; } = "";
    }

    /// <summary>
    /// 列关联颜色检查器 — 根据跨文件列关联规则为DataGridView单元格着色。
    /// 所有缓存操作通过 FileDataCache 静态方法。
    /// </summary>
    public static class ColumnLinkChecker
    {
        /// <summary>从绝对路径获取相对路径（相对于 rootDir）</summary>
        private static string GetRelativePath(string fullPath, string rootDir)
        {
            if (fullPath.StartsWith(rootDir, StringComparison.OrdinalIgnoreCase))
                return fullPath.Substring(rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(fullPath);
        }

        /// <summary>检查指定值在目标表中是否存在</summary>
        public static bool IsLinkValueValid(string filePath, string rootDir, int originalColumnIndex,
            string cellValue, List<ColumnLinkConfig>? columnLinks = null)
        {
            if (string.IsNullOrEmpty(cellValue) || string.IsNullOrEmpty(rootDir)) return false;

            string currentRel = GetRelativePath(filePath, rootDir);

            var linkConfig = columnLinks?.FirstOrDefault(l =>
                l.SourceFile.EndsWith(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == originalColumnIndex);

            var matchValues = (linkConfig != null && linkConfig.ValueExtractionEnabled)
                ? linkConfig.ExtractValues(cellValue) : new List<string> { cellValue };

            string prefix = currentRel + "|" + originalColumnIndex + "->";
            var lookups = FileDataCache.GetLookups();
            foreach (var kvp in lookups)
            {
                if (kvp.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return matchValues.Any(v => kvp.Value.Values.Contains(v));
                }
            }
            return false;
        }

        /// <summary>获取关联字段无效的提示信息</summary>
        public static string? GetInvalidLinkToolTip(string filePath, string rootDir,
            List<int> visibleColumnIndices, DataGridView dataGrid, int rowIndex, int columnIndex,
            List<ColumnLinkConfig>? columnLinks = null,
            Dictionary<string, string>? fileRemarks = null)
        {
            if (string.IsNullOrEmpty(rootDir)) return null;

            string currentRel = GetRelativePath(filePath, rootDir);

            int originalColIndex = columnIndex;
            if (visibleColumnIndices.Count > 0 && columnIndex < dataGrid.Columns.Count)
            {
                int adjustedCol = columnIndex;
                if (adjustedCol >= 0 && adjustedCol < visibleColumnIndices.Count)
                    originalColIndex = visibleColumnIndices[adjustedCol];
            }

            var cellVal = dataGrid.Rows[rowIndex]?.Cells[columnIndex]?.Value?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(cellVal)) return null;

            var linkConfig = columnLinks?.FirstOrDefault(l =>
                l.SourceFile.EndsWith(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == originalColIndex);

            var matchValues = (linkConfig != null && linkConfig.ValueExtractionEnabled)
                ? linkConfig.ExtractValues(cellVal) : new List<string> { cellVal };

            string prefix = currentRel + "|" + originalColIndex + "->";
            var lookups = FileDataCache.GetLookups();
            foreach (var kvp in lookups)
            {
                if (!kvp.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                var keyAfterArrow = kvp.Key.Substring(prefix.Length);
                int extractIdx = keyAfterArrow.IndexOf("|extract:", StringComparison.Ordinal);
                if (extractIdx >= 0)
                    keyAfterArrow = keyAfterArrow.Substring(0, extractIdx);

                if (matchValues.Count > 1 && matchValues.All(v => !kvp.Value.Values.Contains(v)))
                {
                    var colonIdx = keyAfterArrow.LastIndexOf(':');
                    string targetFile = colonIdx > 0 ? keyAfterArrow.Substring(0, colonIdx) : keyAfterArrow;
                    string targetCol = colonIdx > 0 ? keyAfterArrow.Substring(colonIdx + 1) : "";
                    string displayTableName = Path.GetFileName(targetFile);
                    if (fileRemarks != null && fileRemarks.TryGetValue(targetFile, out var remark) && !string.IsNullOrEmpty(remark))
                        displayTableName = $"{remark} ({Path.GetFileName(targetFile)})";
                    else if (!string.IsNullOrEmpty(targetCol))
                        displayTableName = $"{Path.GetFileName(targetFile)} [{targetCol}]";
                    string displayValue = linkConfig?.ValueExtractionEnabled == true
                        ? $"{cellVal} (提取: {string.Join(", ", matchValues)})" : cellVal;
                    return $"关联表 [{displayTableName}] 的关联列中未找到该数据: \"{displayValue}\"";
                }
                break;
            }
            return null;
        }

        /// <summary>获取关联字段的显示提示信息</summary>
        public static string? GetLinkInfoToolTip(string filePath, string rootDir,
            List<int> visibleColumnIndices, DataGridView dataGrid, int rowIndex, int columnIndex,
            List<ColumnLinkConfig>? columnLinks = null,
            Dictionary<string, string>? fileRemarks = null)
        {
            if (string.IsNullOrEmpty(rootDir)) return null;

            string currentRel = GetRelativePath(filePath, rootDir);

            int originalColIndex = columnIndex;
            if (visibleColumnIndices.Count > 0 && columnIndex < dataGrid.Columns.Count)
            {
                int adjustedCol = columnIndex;
                if (adjustedCol >= 0 && adjustedCol < visibleColumnIndices.Count)
                    originalColIndex = visibleColumnIndices[adjustedCol];
            }

            var cellVal = dataGrid.Rows[rowIndex]?.Cells[columnIndex]?.Value?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(cellVal)) return null;

            var linkConfig = columnLinks?.FirstOrDefault(l =>
                l.SourceFile.EndsWith(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == originalColIndex);

            var matchValues = (linkConfig != null && linkConfig.ValueExtractionEnabled)
                ? linkConfig.ExtractValues(cellVal) : new List<string> { cellVal };

            string prefix = currentRel + "|" + originalColIndex + "->";
            var lookups = FileDataCache.GetLookups();
            foreach (var kvp in lookups)
            {
                if (!kvp.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;

                string targetFileName = Path.GetFileName(kvp.Value.TargetFile);
                string displayTableName = targetFileName;
                if (fileRemarks != null && fileRemarks.TryGetValue(kvp.Value.TargetFile, out var remark) && !string.IsNullOrEmpty(remark))
                    displayTableName = $"{remark} ({targetFileName})";
                else if (!string.IsNullOrEmpty(kvp.Value.TargetColumnName))
                    displayTableName = $"{targetFileName} [{kvp.Value.TargetColumnName}]";

                var displayLines = new List<string>();
                foreach (var mv in matchValues)
                {
                    if (kvp.Value.Values.Contains(mv))
                    {
                        if (kvp.Value.DisplayNames.TryGetValue(mv, out var displayName) && !string.IsNullOrEmpty(displayName))
                            displayLines.Add($"{mv}: {displayName}");
                        else
                            displayLines.Add($"{mv}: (无显示名)");
                    }
                    else
                    {
                        displayLines.Add($"{mv}: 关联表未找到");
                    }
                }

                if (displayLines.Count == 0) return null;

                var orderedLines = displayLines
                    .Where(l => l.Contains("关联表未找到"))
                    .Concat(displayLines.Where(l => !l.Contains("关联表未找到")))
                    .ToList();

                var matchedCount = displayLines.Count(l => !l.Contains("关联表未找到"));
                var sb = new StringBuilder();
                sb.AppendLine(displayTableName);
                sb.AppendLine($"已匹配: {matchedCount}/{displayLines.Count}");
                int maxLines = 15;
                foreach (var line in orderedLines.Take(maxLines))
                    sb.AppendLine(line);
                if (orderedLines.Count > maxLines)
                    sb.AppendLine($"... 还有 {orderedLines.Count - maxLines} 条未显示");
                return sb.ToString().TrimEnd('\r', '\n');
            }
            return null;
        }

        /// <summary>检查单元格颜色</summary>
        public static void CheckCellColor(DataGridViewCellFormattingEventArgs e,
            List<int> visibleColumnIndices, DataGridView dataGrid, string filePath, string rootDir,
            List<ColumnLinkConfig>? columnLinks = null)
        {
            if (string.IsNullOrEmpty(rootDir)) return;

            string currentRel = GetRelativePath(filePath, rootDir);

            int originalColIndex = e.ColumnIndex;
            if (visibleColumnIndices.Count > 0 && e.ColumnIndex < dataGrid.Columns.Count)
            {
                int adjustedCol = e.ColumnIndex;
                if (adjustedCol >= 0 && adjustedCol < visibleColumnIndices.Count)
                    originalColIndex = visibleColumnIndices[adjustedCol];
            }

            var linkConfig = columnLinks?.FirstOrDefault(l =>
                l.SourceFile.EndsWith(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == originalColIndex);

            var lookups = FileDataCache.GetLookups();
            foreach (var kvp in lookups)
            {
                var pipeIdx = kvp.Key.IndexOf('|');
                if (pipeIdx < 0) continue;
                var keySourceRel = kvp.Key.Substring(0, pipeIdx);
                if (!keySourceRel.Equals(currentRel, StringComparison.OrdinalIgnoreCase)) continue;
                var afterPipe = kvp.Key.Substring(pipeIdx + 1);
                var arrowIdx = afterPipe.IndexOf("->", StringComparison.Ordinal);
                if (arrowIdx < 0) continue;
                var colStr = afterPipe.Substring(0, arrowIdx);
                if (int.TryParse(colStr, out int srcCol) && srcCol == originalColIndex)
                {
                    var cellValue = e.Value?.ToString()?.Trim() ?? "";

                    var matchValues = (linkConfig != null && linkConfig.ValueExtractionEnabled)
                        ? linkConfig.ExtractValues(cellValue) : new List<string> { cellValue };

                    if (!string.IsNullOrEmpty(cellValue) && matchValues.Count > 0)
                    {
                        bool hasMissing = matchValues.Any(v => !kvp.Value.Values.Contains(v));
                        if (hasMissing)
                        {
                            e.CellStyle.BackColor = Color.FromArgb(255, 180, 180);
                            e.CellStyle.SelectionBackColor = Color.FromArgb(255, 150, 150);
                        }
                        else
                        {
                            e.CellStyle.BackColor = dataGrid.DefaultCellStyle.BackColor;
                            e.CellStyle.SelectionBackColor = dataGrid.DefaultCellStyle.SelectionBackColor;
                        }
                    }
                    else
                    {
                        e.CellStyle.BackColor = dataGrid.DefaultCellStyle.BackColor;
                        e.CellStyle.SelectionBackColor = dataGrid.DefaultCellStyle.SelectionBackColor;
                    }
                    break;
                }
            }
        }
    }
}
