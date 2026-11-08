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

        /// <summary>ID → 完整行数据（表模式使用：ID → 该行所有单元格值列表）</summary>
        public Dictionary<string, List<string>> RowData { get; set; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>目标文件的列名列表（表模式使用：用于自动生成列头）</summary>
        public List<string> ColumnNames { get; set; } = new();

        /// <summary>ID → 行原始数据字符串（未切割的整行）</summary>
        public Dictionary<string, string> RowRawStrings { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>列关联颜色检查器 — 根据跨文件列关联规则为DataGridView单元格着色。</summary>
    public static class ColumnLinkChecker
    {
        private static string GetRelativePath(string fullPath, string rootDir)
        {
            if (fullPath.StartsWith(rootDir, StringComparison.OrdinalIgnoreCase))
                return fullPath.Substring(rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(fullPath);
        }

        /// <summary>检查指定值在目标表中是否存在（检查所有有关联的字段）</summary>
        /// <param name="cache">关联 lookup 所在的缓存实例（由调用方的阅读器窗口持有）</param>
        public static bool IsLinkValueValid(FileDataCache cache, string filePath, string rootDir, int originalColumnIndex,
            string cellValue, List<ColumnLinkConfig>? columnLinks = null)
        {
            if (string.IsNullOrEmpty(cellValue) || string.IsNullOrEmpty(rootDir)) return false;

            string currentRel = GetRelativePath(filePath, rootDir);

            var linkConfig = columnLinks?.FirstOrDefault(l =>
                LinkSourceMatcher.Matches(l.SourceFile, currentRel, Path.GetFileName(filePath)) &&
                l.SourceColumnIndex == originalColumnIndex);

            if (linkConfig == null || linkConfig.FieldLinks == null || linkConfig.FieldLinks.Count == 0)
                return false;

            // ★ 切割为"行×字段"，与 ToolTip 显示逻辑一致
            var fieldRows = SplitCellToFieldRows(linkConfig, cellValue);

            foreach (var fl in linkConfig.FieldLinks)
            {
                if (string.IsNullOrEmpty(fl.TargetFile)) continue;

                string fieldPrefix = $"{currentRel}|{originalColumnIndex}->field{fl.FieldIndex}->";
                foreach (var entry in cache.GetLookupsByPrefix(fieldPrefix))
                {
                    // 检查每行该字段的值是否在 lookup 中存在
                    foreach (var fieldValues in fieldRows)
                    {
                        string fieldVal = fl.FieldIndex < fieldValues.Count ? fieldValues[fl.FieldIndex] : "";
                        if (!string.IsNullOrEmpty(fieldVal) && entry.Values.Contains(fieldVal))
                            return true;
                    }
                }
            }
            return false;
        }

        /// <summary>获取关联字段无效的提示信息</summary>
        public static string? GetInvalidLinkToolTip(FileDataCache cache, string filePath, string rootDir,
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
                LinkSourceMatcher.Matches(l.SourceFile, currentRel, Path.GetFileName(filePath)) &&
                l.SourceColumnIndex == originalColIndex);

            if (linkConfig == null || linkConfig.FieldLinks == null || linkConfig.FieldLinks.Count == 0)
                return null;

            // ★ 切割为"行×字段"，与 ToolTip 显示逻辑一致
            var fieldRows = SplitCellToFieldRows(linkConfig, cellVal);

            foreach (var fl in linkConfig.FieldLinks)
            {
                if (string.IsNullOrEmpty(fl.TargetFile)) continue;

                string fieldPrefix = $"{currentRel}|{originalColIndex}->field{fl.FieldIndex}->";
                foreach (var entry in cache.GetLookupsByPrefix(fieldPrefix))
                {
                    bool allFound = true;
                    var notFoundValues = new List<string>();

                    foreach (var fieldValues in fieldRows)
                    {
                        string fieldVal = fl.FieldIndex < fieldValues.Count ? fieldValues[fl.FieldIndex] : "";
                        if (!string.IsNullOrEmpty(fieldVal) && !entry.Values.Contains(fieldVal))
                        {
                            allFound = false;
                            notFoundValues.Add(fieldVal);
                        }
                    }

                    if (!allFound)
                    {
                        string displayTableName = Path.GetFileName(fl.TargetFile);
                        if (fileRemarks != null && fileRemarks.TryGetValue(fl.TargetFile, out var remark) && !string.IsNullOrEmpty(remark))
                            displayTableName = $"{remark} ({Path.GetFileName(fl.TargetFile)})";
                        else if (!string.IsNullOrEmpty(fl.TargetColumnName))
                            displayTableName = $"{Path.GetFileName(fl.TargetFile)} [{fl.TargetColumnName}]";
                        return $"关联表 [{displayTableName}] 第{fl.FieldIndex}段未找到: \"{string.Join(", ", notFoundValues)}\"";
                    }
                    break;
                }
            }
            return null;
        }

        /// <summary>获取关联字段的显示提示信息</summary>
        public static string? GetLinkInfoToolTip(FileDataCache cache, string filePath, string rootDir,
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
                LinkSourceMatcher.Matches(l.SourceFile, currentRel, Path.GetFileName(filePath)) &&
                l.SourceColumnIndex == originalColIndex);

            if (linkConfig == null || linkConfig.FieldLinks == null || linkConfig.FieldLinks.Count == 0)
                return null;

            // ★ 切割为"行×字段"，与 ToolTip 显示逻辑一致
            var fieldRows = SplitCellToFieldRows(linkConfig, cellVal);

            var allFieldLines = new List<string>();
            int totalMatchable = 0;
            int totalMatched = 0;

            foreach (var fl in linkConfig.FieldLinks)
            {
                if (string.IsNullOrEmpty(fl.TargetFile)) continue;

                string fieldPrefix = $"{currentRel}|{originalColIndex}->field{fl.FieldIndex}->";
                foreach (var entry in cache.GetLookupsByPrefix(fieldPrefix))
                {
                    string targetFileName = Path.GetFileName(entry.TargetFile);
                    string displayTableName = targetFileName;
                    if (fileRemarks != null && fileRemarks.TryGetValue(entry.TargetFile, out var remark) && !string.IsNullOrEmpty(remark))
                        displayTableName = $"{remark} ({targetFileName})";
                    else if (!string.IsNullOrEmpty(entry.TargetColumnName))
                        displayTableName = $"{targetFileName} [{entry.TargetColumnName}]";

                    allFieldLines.Add($"--- 第{fl.FieldIndex}段: {displayTableName} ---");

                    foreach (var fieldValues in fieldRows)
                    {
                        string fieldVal = fl.FieldIndex < fieldValues.Count ? fieldValues[fl.FieldIndex] : "";
                        if (string.IsNullOrEmpty(fieldVal)) continue;

                        totalMatchable++;
                        if (entry.Values.Contains(fieldVal))
                        {
                            totalMatched++;
                            if (entry.DisplayNames.TryGetValue(fieldVal, out var displayName) && !string.IsNullOrEmpty(displayName))
                                allFieldLines.Add($"  {fieldVal}: {displayName}");
                            else
                                allFieldLines.Add($"  {fieldVal}: (无显示名)");
                        }
                        else
                        {
                            allFieldLines.Add($"  {fieldVal}: ❌ 关联表未找到");
                        }
                    }
                    break;
                }
            }

            if (allFieldLines.Count == 0) return null;

            var sb = new StringBuilder();
            sb.AppendLine($"已匹配: {totalMatched}/{totalMatchable}");
            int maxLines = 20;
            foreach (var line in allFieldLines.Take(maxLines))
                sb.AppendLine(line);
            if (allFieldLines.Count > maxLines)
                sb.AppendLine($"... 还有 {allFieldLines.Count - maxLines} 条未显示");
            return sb.ToString().TrimEnd('\r', '\n');
        }

        /// <summary>检查单元格颜色（检查所有有关联的字段）</summary>
        public static void CheckCellColor(FileDataCache cache, DataGridViewCellFormattingEventArgs e,
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
                LinkSourceMatcher.Matches(l.SourceFile, currentRel, Path.GetFileName(filePath)) &&
                l.SourceColumnIndex == originalColIndex);

            if (linkConfig == null || linkConfig.FieldLinks == null || linkConfig.FieldLinks.Count == 0)
                return;

            var cellValue = e.Value?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(cellValue)) return;

            // ★ 切割为"行×字段"，与 ToolTip 显示逻辑一致
            var fieldRows = SplitCellToFieldRows(linkConfig, cellValue);

            bool hasMissing = false;

            foreach (var fl in linkConfig.FieldLinks)
            {
                if (string.IsNullOrEmpty(fl.TargetFile)) continue;

                string fieldPrefix = $"{currentRel}|{originalColIndex}->field{fl.FieldIndex}->";
                foreach (var entry in cache.GetLookupsByPrefix(fieldPrefix))
                {
                    foreach (var fieldValues in fieldRows)
                    {
                        string fieldVal = fl.FieldIndex < fieldValues.Count ? fieldValues[fl.FieldIndex] : "";
                        if (!string.IsNullOrEmpty(fieldVal) && !entry.Values.Contains(fieldVal))
                        {
                            hasMissing = true;
                            break;
                        }
                    }
                    if (hasMissing) break;
                }
                if (hasMissing) break;
            }

            if (hasMissing)
            {
                e.CellStyle.BackColor = Theme.CellTint(Color.FromArgb(255, 180, 180));
                e.CellStyle.SelectionBackColor = Theme.CellTint(Color.FromArgb(255, 150, 150));
            }
            else
            {
                e.CellStyle.BackColor = dataGrid.DefaultCellStyle.BackColor;
                e.CellStyle.SelectionBackColor = dataGrid.DefaultCellStyle.SelectionBackColor;
            }
        }

        private static List<string> SplitByDelimiter(string value, string delimiter)
        {
            if (string.IsNullOrEmpty(delimiter) || !value.Contains(delimiter))
                return new List<string> { value };
            return value.Split(new[] { delimiter }, StringSplitOptions.None)
                .Select(s => s.Trim())
                .ToList();
        }

        /// <summary>
        /// 将单元格内容按 ExtractSteps 切割为"行×字段"的二维结构。
        /// 切割逻辑与 ToolTip 显示完全一致：
        ///   前 N-1 步按行拆分，最后一步按列拆分。
        /// 使用 ColumnLinkConfig.ExtractRows 方法。
        /// </summary>
        private static List<List<string>> SplitCellToFieldRows(ColumnLinkConfig linkConfig, string cellValue)
        {
            if (string.IsNullOrEmpty(cellValue)) return new List<List<string>>();
            return linkConfig.ExtractRows(cellValue);
        }
    }
}
