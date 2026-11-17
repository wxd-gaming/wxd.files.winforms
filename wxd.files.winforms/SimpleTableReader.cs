using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;

namespace wxd.files.winforms
{
    /// <summary>
    /// 简单的表格读取：CSV / Excel 只取"原始单元格"，不区分表头行、不做冻结行。
    /// 与阅读器（TableBrowserControl）完全独立：不解析表头语义、不建 FileTableData、不管列显隐。
    /// </summary>
    internal static class SimpleTableReader
    {
        private static bool _providerRegistered;

        /// <summary>扩展名是否按表格比对（csv / xlsx / xls）。</summary>
        internal static bool IsTableFile(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext is ".csv" or ".xlsx" or ".xls";
        }

        /// <summary>读取表格所有行；行内单元格统一 Trim，缺列不补齐（比对时按空串处理）。</summary>
        internal static List<string[]> ReadRows(string path)
        {
            return IsCsv(path) ? ReadCsvRows(path) : ReadExcelRows(path);
        }

        private static bool IsCsv(string path)
            => string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase);

        // ==================== CSV ====================

        private static List<string[]> ReadCsvRows(string path)
        {
            var text = TextFileReader.ReadAllText(path);
            char delimiter = DetectCsvDelimiter(text);
            return SplitCsvRecords(text, delimiter);
        }

        /// <summary>只看第一行，统计引号外的候选分隔符出现次数，取最多的那个。</summary>
        internal static char DetectCsvDelimiter(string text)
        {
            var counts = new Dictionary<char, int>
            {
                [','] = 0,
                [GlobalConst.TabChar] = 0,
                [';'] = 0,
                ['|'] = 0
            };

            bool inQuotes = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    continue;
                }
                if (!inQuotes && (c == GlobalConst.LfChar || c == GlobalConst.CrChar)) break;
                if (!inQuotes && counts.ContainsKey(c)) counts[c]++;
            }

            return counts.OrderByDescending(kv => kv.Value).First().Key;
        }

        /// <summary>按 CSV 规则切记录：引号内换行不切行、双引号转义，单元格 Trim。</summary>
        internal static List<string[]> SplitCsvRecords(string text, char delimiter)
        {
            var rows = new List<string[]>();
            var cells = new List<string>();
            var cell = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < text.Length && text[i + 1] == '"')
                    {
                        cell.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                    continue;
                }

                if (!inQuotes && c == delimiter)
                {
                    cells.Add(cell.ToString().Trim());
                    cell.Clear();
                    continue;
                }

                if (!inQuotes && (c == GlobalConst.LfChar || c == GlobalConst.CrChar))
                {
                    if (c == GlobalConst.CrChar && i + 1 < text.Length && text[i + 1] == GlobalConst.LfChar)
                        i++;
                    cells.Add(cell.ToString().Trim());
                    cell.Clear();
                    rows.Add(cells.ToArray());
                    cells.Clear();
                    continue;
                }

                cell.Append(c);
            }

            // 最后一行没有换行结尾
            if (cell.Length > 0 || cells.Count > 0)
            {
                cells.Add(cell.ToString().Trim());
                rows.Add(cells.ToArray());
            }

            return rows;
        }

        // ==================== Excel ====================

        private static List<string[]> ReadExcelRows(string path)
        {
            var rows = new List<string[]>();

            if (!_providerRegistered)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                _providerRegistered = true;
            }

            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(fs);
            do
            {
                while (reader.Read())
                {
                    int count = reader.FieldCount;
                    var cells = new string[count];
                    for (int i = 0; i < count; i++)
                        cells[i] = Convert.ToString(reader.GetValue(i), CultureInfo.InvariantCulture)?.Trim() ?? "";
                    rows.Add(cells);
                }
                break; // 只看第一个工作表，和阅读器保持一致
            } while (reader.NextResult());

            return rows;
        }
    }
}
