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
        public static bool IsLinkValueValid(string filePath, string rootDir, int originalColumnIndex,
            string cellValue, List<ColumnLinkConfig>? columnLinks = null)
        {
            if (string.IsNullOrEmpty(cellValue) || string.IsNullOrEmpty(rootDir)) return false;

            string currentRel = GetRelativePath(filePath, rootDir);

            var linkConfig = columnLinks?.FirstOrDefault(l =>
                l.SourceFile.EndsWith(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == originalColumnIndex);

            if (linkConfig == null || linkConfig.FieldLinks == null || linkConfig.FieldLinks.Count == 0)
                return false;

            // ★ 切割为"行×字段"，与 ToolTip 显示逻辑一致
            var fieldRows = SplitCellToFieldRows(linkConfig, cellValue);

            var lookups = FileDataCache.GetLookups();

            foreach (var fl in linkConfig.FieldLinks)
            {
                if (string.IsNullOrEmpty(fl.TargetFile)) continue;

                string fieldPrefix = $"{currentRel}|{originalColumnIndex}->field{fl.FieldIndex}->";
                foreach (var kvp in lookups)
                {
                    if (kvp.Key.StartsWith(fieldPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        // 检查每行该字段的值是否在 lookup 中存在
                        foreach (var fieldValues in fieldRows)
                        {
                            string fieldVal = fl.FieldIndex < fieldValues.Count ? fieldValues[fl.FieldIndex] : "";
                            if (!string.IsNullOrEmpty(fieldVal) && kvp.Value.Values.Contains(fieldVal))
                                return true;
                        }
                    }
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

            if (linkConfig == null || linkConfig.FieldLinks == null || linkConfig.FieldLinks.Count == 0)
                return null;

            // ★ 切割为"行×字段"，与 ToolTip 显示逻辑一致
            var fieldRows = SplitCellToFieldRows(linkConfig, cellVal);

            var lookups = FileDataCache.GetLookups();

            foreach (var fl in linkConfig.FieldLinks)
            {
                if (string.IsNullOrEmpty(fl.TargetFile)) continue;

                string fieldPrefix = $"{currentRel}|{originalColIndex}->field{fl.FieldIndex}->";
                foreach (var kvp in lookups)
                {
                    if (!kvp.Key.StartsWith(fieldPrefix, StringComparison.OrdinalIgnoreCase)) continue;

                    bool allFound = true;
                    var notFoundValues = new List<string>();

                    foreach (var fieldValues in fieldRows)
                    {
                        string fieldVal = fl.FieldIndex < fieldValues.Count ? fieldValues[fl.FieldIndex] : "";
                        if (!string.IsNullOrEmpty(fieldVal) && !kvp.Value.Values.Contains(fieldVal))
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

            if (linkConfig == null || linkConfig.FieldLinks == null || linkConfig.FieldLinks.Count == 0)
                return null;

            // ★ 切割为"行×字段"，与 ToolTip 显示逻辑一致
            var fieldRows = SplitCellToFieldRows(linkConfig, cellVal);

            var lookups = FileDataCache.GetLookups();

            var allFieldLines = new List<string>();
            int totalMatchable = 0;
            int totalMatched = 0;

            foreach (var fl in linkConfig.FieldLinks)
            {
                if (string.IsNullOrEmpty(fl.TargetFile)) continue;

                string fieldPrefix = $"{currentRel}|{originalColIndex}->field{fl.FieldIndex}->";
                foreach (var kvp in lookups)
                {
                    if (!kvp.Key.StartsWith(fieldPrefix, StringComparison.OrdinalIgnoreCase)) continue;

                    string targetFileName = Path.GetFileName(kvp.Value.TargetFile);
                    string displayTableName = targetFileName;
                    if (fileRemarks != null && fileRemarks.TryGetValue(kvp.Value.TargetFile, out var remark) && !string.IsNullOrEmpty(remark))
                        displayTableName = $"{remark} ({targetFileName})";
                    else if (!string.IsNullOrEmpty(kvp.Value.TargetColumnName))
                        displayTableName = $"{targetFileName} [{kvp.Value.TargetColumnName}]";

                    allFieldLines.Add($"--- 第{fl.FieldIndex}段: {displayTableName} ---");

                    foreach (var fieldValues in fieldRows)
                    {
                        string fieldVal = fl.FieldIndex < fieldValues.Count ? fieldValues[fl.FieldIndex] : "";
                        if (string.IsNullOrEmpty(fieldVal)) continue;

                        totalMatchable++;
                        if (kvp.Value.Values.Contains(fieldVal))
                        {
                            totalMatched++;
                            if (kvp.Value.DisplayNames.TryGetValue(fieldVal, out var displayName) && !string.IsNullOrEmpty(displayName))
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

            if (linkConfig == null || linkConfig.FieldLinks == null || linkConfig.FieldLinks.Count == 0)
                return;

            var cellValue = e.Value?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(cellValue)) return;

            // ★ 切割为"行×字段"，与 ToolTip 显示逻辑一致
            var fieldRows = SplitCellToFieldRows(linkConfig, cellValue);

            var lookups = FileDataCache.GetLookups();

            bool hasMissing = false;

            foreach (var fl in linkConfig.FieldLinks)
            {
                if (string.IsNullOrEmpty(fl.TargetFile)) continue;

                string fieldPrefix = $"{currentRel}|{originalColIndex}->field{fl.FieldIndex}->";
                foreach (var kvp in lookups)
                {
                    if (!kvp.Key.StartsWith(fieldPrefix, StringComparison.OrdinalIgnoreCase)) continue;

                    foreach (var fieldValues in fieldRows)
                    {
                        string fieldVal = fl.FieldIndex < fieldValues.Count ? fieldValues[fl.FieldIndex] : "";
                        if (!string.IsNullOrEmpty(fieldVal) && !kvp.Value.Values.Contains(fieldVal))
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
                e.CellStyle.BackColor = Color.FromArgb(255, 180, 180);
                e.CellStyle.SelectionBackColor = Color.FromArgb(255, 150, 150);
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
        /// 切割逻辑与 BuildFieldLinkedToolTipContent 完全一致：
        ///   1. 找多值分隔符（第一个 PartIndex==-1 的步骤的 Delimiter）作为行分隔符
        ///   2. 用 EffectiveDisplayDelimiter（最后一步的 Delimiter）作为字段分隔符
        /// 注意：关联配置只用于查询显示，不应改变切割/列数规则。
        /// </summary>
        private static List<List<string>> SplitCellToFieldRows(ColumnLinkConfig linkConfig, string cellValue)
        {
            var result = new List<List<string>>();
            if (string.IsNullOrEmpty(cellValue)) return result;

            // 步骤1：找多值分隔符（第一个 PartIndex==-1 的步骤）
            string multiDelimiter = "|";
            foreach (var step in linkConfig.ExtractSteps)
            {
                if (step.PartIndex == -1 && !string.IsNullOrEmpty(step.Delimiter))
                {
                    multiDelimiter = step.Delimiter;
                    break;
                }
            }

            // 步骤2：用 multiDelimiter 切行
            string cellDelimiter = linkConfig.EffectiveDisplayDelimiter;
            var rawLines = cellValue.Split(new[] { multiDelimiter }, StringSplitOptions.None);

            // 步骤3：每行用 cellDelimiter 切字段
            foreach (var rawLine in rawLines)
            {
                var trimmed = rawLine.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;

                var fieldValues = new List<string>();
                if (!string.IsNullOrEmpty(cellDelimiter) && trimmed.Contains(cellDelimiter))
                {
                    foreach (var p in trimmed.Split(new[] { cellDelimiter }, StringSplitOptions.None))
                        fieldValues.Add(p.Trim());
                }
                else
                {
                    fieldValues.Add(trimmed);
                }
                result.Add(fieldValues);
            }
            return result;
        }
    }
}
