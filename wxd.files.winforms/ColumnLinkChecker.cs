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
                        try
                        {
                            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                            using var fs = new FileStream(targetFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            using var excelReader = ExcelReaderFactory.CreateReader(fs);
                            var dataSet = excelReader.AsDataSet(new ExcelDataSetConfiguration
                            {
                                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
                            });
                            if (dataSet.Tables.Count > 0)
                            {
                                var sheet = dataSet.Tables[0];
                                for (int r = 0; r < sheet.Rows.Count; r++)
                                {
                                    var val = sheet.Rows[r][link.TargetColumnIndex]?.ToString()?.Trim() ?? "";
                                    if (!string.IsNullOrEmpty(val))
                                        values.Add(val);
                                }
                            }
                        }
                        catch { }
                    }

                    // key 格式: "{sourceRel}|{sourceColIndex}->{targetRel}:{targetColIndex}"
                    // 如果启用了值提取，在 key 中附加提取规则标识
                    string key = string.Format("{0}|{1}->{2}:{3}", sourceRel, link.SourceColumnIndex, link.TargetFile, link.TargetColumnIndex);
                    if (link.ValueExtractionEnabled)
                    {
                        key += string.Format("|extract:{0}:{1}",
                            string.IsNullOrEmpty(link.ExtractDelimiter) ? "_pos_" : link.ExtractDelimiter,
                            link.ExtractPartIndex);
                    }
                    _linkLookups[key] = values;
                }
                catch { }
            }
        }

        /// <summary>检查指定值在目标表中是否存在（用于右键菜单置灰判断）</summary>
        /// <param name="columnLinks">关联配置列表（用于获取提取规则）</param>
        /// <returns>true=值有效可打开，false=值无效或不存在</returns>
        public static bool IsLinkValueValid(string filePath, string rootDir, int originalColumnIndex,
            string cellValue, List<ColumnLinkConfig>? columnLinks = null)
        {
            if (string.IsNullOrEmpty(cellValue) || _loadedSources.Count == 0) return false;
            
            string currentRel;
            if (!string.IsNullOrEmpty(rootDir) && filePath.StartsWith(rootDir, StringComparison.OrdinalIgnoreCase))
                currentRel = filePath.Substring(rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            else
                currentRel = Path.GetFileName(filePath);

            // 查找当前列的提取规则
            var linkConfig = columnLinks?.FirstOrDefault(l =>
                l.SourceFile.EndsWith(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == originalColumnIndex);
            
            // 如果有提取规则，先提取实际匹配值（支持多值）
            var matchValues = (linkConfig != null && linkConfig.ValueExtractionEnabled)
                ? linkConfig.ExtractValues(cellValue) : new List<string> { cellValue };

            string prefix = currentRel + "|" + originalColumnIndex + "->";
            foreach (var kvp in _linkLookups)
            {
                if (kvp.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return matchValues.Any(v => kvp.Value.Contains(v));
                }
            }
            return false;
        }

        /// <summary>获取关联字段无效的提示信息（用于 CellToolTipTextNeeded）</summary>
        /// <param name="columnLinks">关联配置列表（用于获取提取规则）</param>
        /// <returns>提示文本，若单元格无关联或值有效则返回 null</returns>
        public static string? GetInvalidLinkToolTip(string filePath, string rootDir,
            List<int> visibleColumnIndices, DataGridView dataGrid, int rowIndex, int columnIndex, int frozenRows,
            List<ColumnLinkConfig>? columnLinks = null)
        {
            // 冻结行不提示
            if (rowIndex >= 0 && rowIndex < frozenRows) return null;
            if (_loadedSources.Count == 0 || _linkLookups.Count == 0) return null;

            // 获取当前文件的相对路径
            string currentRel;
            if (!string.IsNullOrEmpty(rootDir) && filePath.StartsWith(rootDir, StringComparison.OrdinalIgnoreCase))
                currentRel = filePath.Substring(rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            else
                currentRel = Path.GetFileName(filePath);

            // 计算原始列索引
            int originalColIndex = columnIndex;
            if (visibleColumnIndices.Count > 0 && columnIndex < dataGrid.Columns.Count)
            {
                int adjustedCol = columnIndex;
                if (dataGrid.Columns.Count > 0 && dataGrid.Columns[0].Name == "__RowNum__")
                    adjustedCol = columnIndex - 1;
                if (adjustedCol >= 0 && adjustedCol < visibleColumnIndices.Count)
                    originalColIndex = visibleColumnIndices[adjustedCol];
            }

            // 获取单元格值
            var cellVal = dataGrid.Rows[rowIndex]?.Cells[columnIndex]?.Value?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(cellVal)) return null;

            // 查找当前列的提取规则
            var linkConfig = columnLinks?.FirstOrDefault(l =>
                l.SourceFile.EndsWith(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == originalColIndex);

            // 如果有提取规则，提取实际匹配值（支持多值）
            var matchValues = (linkConfig != null && linkConfig.ValueExtractionEnabled)
                ? linkConfig.ExtractValues(cellVal) : new List<string> { cellVal };

            // 查找匹配的关联配置
            string prefix = currentRel + "|" + originalColIndex + "->";
            foreach (var kvp in _linkLookups)
            {
                if (!kvp.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                var keyAfterArrow = kvp.Key.Substring(prefix.Length); // "targetRel:targetCol" 或带 extract 后缀
                // 去掉可能的 extract 后缀（"|extract:..."）
                int extractIdx = keyAfterArrow.IndexOf("|extract:", StringComparison.Ordinal);
                if (extractIdx >= 0)
                    keyAfterArrow = keyAfterArrow.Substring(0, extractIdx);
                
                if (!matchValues.Any(v => kvp.Value.Contains(v)))
                {
                    var colonIdx = keyAfterArrow.LastIndexOf(':');
                    string targetFile = colonIdx > 0 ? keyAfterArrow.Substring(0, colonIdx) : keyAfterArrow;
                    string targetCol = colonIdx > 0 ? keyAfterArrow.Substring(colonIdx + 1) : "";
                    string displayValue = linkConfig?.ValueExtractionEnabled == true
                        ? $"{cellVal} (提取: {string.Join(", ", matchValues)})" : cellVal;
                    return $"关联表 [{Path.GetFileName(targetFile)}] 的关联列中未找到该数据: \"{displayValue}\"";
                }
                break; // 值存在，不需要提示
            }
            return null;
        }

        /// <summary>检查单元格颜色 — 限匹配当前文件的源关联，不存在则标红，存在则恢复默认；跳过冻结行</summary>
        /// <param name="columnLinks">关联配置列表（用于获取提取规则）</param>
        public static void CheckCellColor(DataGridViewCellFormattingEventArgs e,
            List<int> visibleColumnIndices, DataGridView dataGrid, string filePath, string rootDir, int frozenRows,
            List<ColumnLinkConfig>? columnLinks = null)
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

            // 查找当前列的提取规则
            var linkConfig = columnLinks?.FirstOrDefault(l =>
                l.SourceFile.EndsWith(Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase) &&
                l.SourceColumnIndex == originalColIndex);

            foreach (var kvp in _linkLookups)
            {
                // key 格式: "{sourceRel}|{srcCol}->{targetRel}:{targetCol}" 或带 "|extract:..." 后缀
                var pipeIdx = kvp.Key.IndexOf('|');
                if (pipeIdx < 0) continue;
                // 只匹配当前文件的关联，排除其他文件的缓存
                var keySourceRel = kvp.Key.Substring(0, pipeIdx);
                if (!keySourceRel.Equals(currentRel, StringComparison.OrdinalIgnoreCase)) continue;
                var afterPipe = kvp.Key.Substring(pipeIdx + 1); // "{srcCol}->{targetRel}:{targetCol}..."
                var arrowIdx = afterPipe.IndexOf("->", StringComparison.Ordinal);
                if (arrowIdx < 0) continue;
                var colStr = afterPipe.Substring(0, arrowIdx);
                if (int.TryParse(colStr, out int srcCol) && srcCol == originalColIndex)
                {
                    var cellValue = e.Value?.ToString()?.Trim() ?? "";
                    
                    // 如果有提取规则，提取实际匹配值（支持多值）
                    var matchValues = (linkConfig != null && linkConfig.ValueExtractionEnabled)
                        ? linkConfig.ExtractValues(cellValue) : new List<string> { cellValue };
                    
                    if (!string.IsNullOrEmpty(cellValue) && !matchValues.Any(v => kvp.Value.Contains(v)))
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
