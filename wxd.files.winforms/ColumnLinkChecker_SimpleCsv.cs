using System.Text;

namespace wxd.files.winforms
{
    internal static class ColumnLinkChecker_SimpleCsv
    {
        public static string[] ParseLine(string line)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\u0022')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '\u0022') { sb.Append('\u0022'); i++; }
                    else inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes) { result.Add(sb.ToString().Trim()); sb.Clear(); }
                else sb.Append(c);
            }
            result.Add(sb.ToString());
            return result.ToArray();
        }

        /// <summary>根据目标文件路径获取其表头行索引（从设置中读取）</summary>
        private static int GetTargetHeaderRowIndexStatic(string targetFullPath)
        {
            if (string.IsNullOrEmpty(targetFullPath)) return -1;
            string targetDir = Path.GetDirectoryName(targetFullPath) ?? "";
            // 如果在同一目录下，使用该目录的设置
            var ext = Path.GetExtension(targetFullPath).ToLowerInvariant();
            if (ext == ".csv")
                return CsvSettings.GetCsvHeaderRowIndexByDirectory(targetDir);
            else if (ext == ".xls" || ext == ".xlsx")
                return CsvSettings.GetExcelHeaderRowIndexByDirectory(targetDir);
            return -1;
        }
    }
}
