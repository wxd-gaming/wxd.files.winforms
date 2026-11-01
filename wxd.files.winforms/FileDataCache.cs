using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDataReader;

namespace wxd.files.winforms
{
    /// <summary>
    /// 表格数据缓存（实例级，由目录阅读器窗体持有）。
    /// key = 全小写绝对路径（如 "c:\work\a.csv"）。
    ///
    /// 生命周期：一个 CsvReaderForm（一个目录的阅读器窗口）持有一个实例，
    /// 窗口关闭时 Dispose，缓存里的 DataTable / FileTableData / LinkLookup 一起失去引用被回收。
    /// A 目录窗口和 B 目录窗口各持一份，互不干扰；同一文件被两个窗口打开时各自解析一份，
    /// 另一窗口写入磁盘后，本窗口下次访问会因 LastWriteTime/FileSize 不符自动重建。
    ///
    /// 释放后的行为：不抛异常，读请求退化为"直接解析、不进字典"，写请求只落盘，
    /// 这样窗口关闭时仍在跑的后台任务（关联 lookup 构建、内容搜索）只会白跑一次，不会崩溃也不会留内存。
    ///
    /// 并发设计（同一实例内）：
    /// - 不同文件使用独立锁（_fileLocks），可并行加载/访问
    /// - 同一文件的并发请求通过 per-file lock + double-check 保证只加载一次
    /// - _entries 使用 ConcurrentDictionary，字典操作本身线程安全
    /// - LinkLookup 相关操作使用独立的 _linkLock，与文件缓存互不干扰
    ///
    /// 无状态部分（解析/写盘/克隆）保持静态：调用方只需要"读一份磁盘文件"时不必扯上缓存归属。
    /// </summary>
    public class FileDataCache : IDisposable
    {
        /// <summary>缓存条目（key = NormalizePath(fullPath)）</summary>
        private readonly ConcurrentDictionary<string, CacheEntry> _entries =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>关联查找缓存（key 沿用现有格式："sourceRel|srcCol->targetRel:targetCol"）</summary>
        private readonly Dictionary<string, LinkLookupEntry> _linkLookups =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>LinkLookup 专用锁（跨文件关联，但作用域只到本实例）</summary>
        private readonly object _linkLock = new();

        /// <summary>每个文件的独立锁对象（保证同一文件的 load 不重入）</summary>
        private readonly ConcurrentDictionary<string, object> _fileLocks =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>正在后台构建关联 lookup 的源文件相对路径集合（防重入）</summary>
        private readonly HashSet<string> _buildingLinkLookups = new(StringComparer.OrdinalIgnoreCase);

        private bool _disposed;

        /// <summary>当前缓存是否已释放（调用方可据此放弃后台工作）</summary>
        public bool IsDisposed => _disposed;

        /// <summary>将绝对路径标准化为 key</summary>
        private static string NormalizePath(string path)
            => Path.GetFullPath(path).ToLowerInvariant();

        /// <summary>从绝对路径获取相对路径</summary>
        private static string GetRelativePath(string fullPath, string rootDir)
        {
            if (fullPath.StartsWith(rootDir, StringComparison.OrdinalIgnoreCase))
                return fullPath.Substring(rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(fullPath);
        }

        // ============ 公共 API：读取（走本实例缓存） ============

        /// <summary>读取文件（走缓存）。自动判断 CSV/Excel。返回深拷贝副本。</summary>
        public (DataTable, FileTableData) LoadFile(string filePath)
        {
            var entry = GetOrLoadEntry(NormalizePath(filePath), filePath);
            return (entry.MasterDataTable.Copy(), CloneFileData(entry.MasterFileData));
        }

        /// <summary>获取只读 DataTable（走缓存，返回引用，调用方不应修改）</summary>
        public DataTable GetDataTableReadOnly(string filePath)
        {
            return GetOrLoadEntry(NormalizePath(filePath), filePath).MasterDataTable;
        }

        /// <summary>确保文件已加载到缓存（预加载用）</summary>
        public void EnsureLoaded(string filePath)
        {
            GetOrLoadEntry(NormalizePath(filePath), filePath);
        }

        /// <summary>清除指定文件的缓存（本实例内）</summary>
        public void InvalidateFile(string filePath)
        {
            if (_disposed) return;
            var key = NormalizePath(filePath);
            _entries.TryRemove(key, out _);
            _fileLocks.TryRemove(key, out _);
            ClearLinkLookupsForFile(key);
        }

        // ============ 公共 API：写入 ============

        /// <summary>
        /// 写入文件并更新本实例缓存。自动判断 CSV/Excel。
        /// ★ 缓存条目用"重新解析磁盘文件"的结果，不从 DataTable 反推：
        ///   只有磁盘解析才能拿到每条记录的物理起止行号，行号区才不会退化。
        /// </summary>
        public bool SaveFile(string filePath, DataTable dataTable, char delimiter = ',')
        {
            if (!WriteToDisk(filePath, dataTable, delimiter)) return false;

            if (!_disposed)
            {
                var key = NormalizePath(filePath);
                var fileLock = _fileLocks.GetOrAdd(key, _ => new object());
                lock (fileLock)
                {
                    _entries[key] = LoadAndParseFile(key, filePath);
                    // 清除该文件相关的旧 linkLookup（下次访问时懒重建）
                    ClearLinkLookupsForFile(key);
                }
            }
            return true;
        }

        /// <summary>写入文件（不碰任何缓存）。自动判断 CSV/Excel。</summary>
        public static bool SaveFileDirect(string filePath, DataTable dataTable, char delimiter = ',')
            => WriteToDisk(filePath, dataTable, delimiter);

        /// <summary>统一落盘逻辑（无状态）</summary>
        private static bool WriteToDisk(string fullPath, DataTable dataTable, char delimiter)
        {
            var ext = Path.GetExtension(fullPath).ToLowerInvariant();
            if (ext == ".csv")
                return WriteCsvToDisk(fullPath, dataTable, delimiter);
            return WriteExcelToDisk(fullPath, dataTable);
        }

        // ============ 关联查找（LinkLookup，key 保持相对路径格式） ============

        /// <summary>获取所有已缓存的关联查找条目（本实例）</summary>
        public Dictionary<string, LinkLookupEntry> GetLookups()
        {
            lock (_linkLock)
            {
                return new Dictionary<string, LinkLookupEntry>(_linkLookups, StringComparer.OrdinalIgnoreCase);
            }
        }

        /// <summary>异步构建所有关联配置的 lookup</summary>
        public async Task BuildLinkLookupsAsync(
            string rootDir, List<ColumnLinkConfig> columnLinks, string sourceFilePath,
            CancellationToken cancellationToken = default)
        {
            if (_disposed) return;

            string sourceRel = GetRelativePath(sourceFilePath, rootDir);

            // ★ 防重入检查
            lock (_linkLock)
            {
                if (_buildingLinkLookups.Contains(sourceRel))
                    return;

                _buildingLinkLookups.Add(sourceRel);

                string prefix = sourceRel + "|";
                var keysToRemove = _linkLookups.Keys
                    .Where(k => k.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                foreach (var k in keysToRemove)
                    _linkLookups.Remove(k);
            }

            try
            {
                var relevantLinks = columnLinks
                    .Where(l => l.SourceFile.EndsWith(Path.GetFileName(sourceFilePath),
                        StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var link in relevantLinks)
                {
                    if (_disposed) return;

                    if (link.FieldLinks == null || link.FieldLinks.Count == 0)
                        continue;

                    foreach (var fl in link.FieldLinks)
                    {
                        if (_disposed) return;

                        if (string.IsNullOrEmpty(fl.TargetFile) || fl.TargetColumnIndex < 0)
                            continue;

                        cancellationToken.ThrowIfCancellationRequested();

                        string targetFullPath = Path.Combine(rootDir, fl.TargetFile);
                        if (!File.Exists(targetFullPath)) continue;

                        var targetKey = NormalizePath(targetFullPath);
                        CacheEntry entry = await Task.Run(() =>
                        {
                            return GetOrLoadEntry(targetKey, targetFullPath);
                        }, cancellationToken);

                        string key = BuildLinkKey(sourceRel, link.SourceColumnIndex,
                            fl.TargetFile, fl.TargetColumnIndex, link, fl.FieldIndex);

                        var lookup = await Task.Run(() =>
                            BuildLinkLookupFromMaster(entry.MasterFileData,
                                fl.TargetColumnIndex, fl), cancellationToken);

                        lock (_linkLock)
                        {
                            if (!_disposed && !_linkLookups.ContainsKey(key))
                                _linkLookups[key] = lookup;
                        }
                    }
                }
            }
            finally
            {
                lock (_linkLock) { _buildingLinkLookups.Remove(sourceRel); }
            }
        }

        // ============ 释放 ============

        /// <summary>
        /// 释放本实例缓存：清空所有字典，让缓存的 DataTable / FileTableData / LinkLookup 被回收。
        /// 只断引用、不 DataTable.Dispose()——可能还有后台线程正在读同一份引用，
        /// 提前 Dispose 会让它抛 ObjectDisposedException；断引用后 GC 自然回收。
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            _entries.Clear();
            _fileLocks.Clear();
            lock (_linkLock)
            {
                _linkLookups.Clear();
                _buildingLinkLookups.Clear();
            }
        }

        // ============ 内部实现 ============

        /// <summary>
        /// 获取或加载条目（并发安全）。
        /// 快速路径：缓存命中且文件未变 → 直接返回（无锁）。
        /// 慢路径：获取该文件独立锁 → double-check → 加载解析。
        /// 已释放的实例：每次现场解析一份临时条目，不进字典，不留数据。
        /// </summary>
        private CacheEntry GetOrLoadEntry(string key, string fullPath)
        {
            if (_disposed)
            {
                var (tempDt, tempFd) = ParseFile(fullPath);
                return new CacheEntry { MasterDataTable = tempDt, MasterFileData = tempFd };
            }

            // 1. 快速路径：缓存已存在，直接返回（ConcurrentDictionary.TryGetValue 无锁）
            if (_entries.TryGetValue(key, out var entry))
            {
                var fi = new FileInfo(fullPath);
                if (fi.Exists &&
                    (fi.LastWriteTimeUtc != entry.LastWriteTime || fi.Length != entry.FileSize))
                {
                    // 文件被外部修改，需失效重建 → 进入慢路径
                }
                else
                {
                    return entry;
                }
            }

            // 2. 慢路径：获取该文件的独立锁，保证同一文件的 load 不重入
            var fileLock = _fileLocks.GetOrAdd(key, _ => new object());
            lock (fileLock)
            {
                // double-check：在等待锁期间，其他线程可能已加载该文件
                if (_entries.TryGetValue(key, out entry))
                {
                    var fi2 = new FileInfo(fullPath);
                    if (fi2.Exists &&
                        (fi2.LastWriteTimeUtc != entry.LastWriteTime || fi2.Length != entry.FileSize))
                    {
                        // 文件仍已变更，需重建
                        _entries.TryRemove(key, out _);
                        ClearLinkLookupsForFile(key);
                    }
                    else
                    {
                        return entry;
                    }
                }

                // 确实需要加载
                entry = LoadAndParseFile(key, fullPath);
                _entries[key] = entry;
                return entry;
            }
        }

        /// <summary>解析文件并构建 CacheEntry</summary>
        private static CacheEntry LoadAndParseFile(string key, string fullPath)
        {
            var (dt, fd) = ParseFile(fullPath);
            var fi = new FileInfo(fullPath);
            return new CacheEntry
            {
                MasterDataTable = dt,
                MasterFileData = fd,
                LastWriteTime = fi.LastWriteTimeUtc,
                FileSize = fi.Length
            };
        }

        // ============ 无状态：读取 / 解析 ============

        /// <summary>读取文件（不走缓存）。自动判断 CSV/Excel。返回深拷贝副本。</summary>
        public static (DataTable, FileTableData) LoadFileDirect(string filePath)
        {
            return ParseFile(filePath);
        }

        /// <summary>统一的文件解析（不涉及缓存），根据扩展名自动路由</summary>
        private static (DataTable, FileTableData) ParseFile(string fullPath)
        {
            var ext = Path.GetExtension(fullPath).ToLowerInvariant();
            if (ext == ".csv")
                return ParseCsvFile(fullPath);
            else
                return ParseExcelFile(fullPath);
        }

        /// <summary>统一的 CSV 解析（不涉及缓存）</summary>
        internal static (DataTable, FileTableData) ParseCsvFile(string fullPath)
        {
            string content;
            using (var fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var sr = new StreamReader(fs, Encoding.UTF8, true))
            {
                content = sr.ReadToEnd();
            }

            var delimiter = DetectDelimiter(content);
            var records = SplitCsvRecords(content);
            var dataTable = new DataTable();
            var fileData = new FileTableData
            {
                FilePath = fullPath,
                Type = TableFileType.Csv,
                Delimiter = delimiter
            };

            if (records.Count == 0)
                return (dataTable, fileData);

            var headers = ParseCsvLine(records[0].Text, delimiter);
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var header in headers)
            {
                var columnName = header;
                int suffix = 1;
                while (usedNames.Contains(columnName)) { columnName = $"{header}_{suffix}"; suffix++; }
                usedNames.Add(columnName);
                dataTable.Columns.Add(columnName, typeof(string));
            }

            // ★ ColumnNames 只需赋值一次（修复循环内 5 万次重复赋值的 bug）
            fileData.ColumnNames = new List<string>(headers);

            for (int i = 0; i < records.Count; i++)
            {
                var values = ParseCsvLine(records[i].Text, delimiter);
                var row = dataTable.NewRow();
                for (int j = 0; j < Math.Min(values.Length, headers.Length); j++)
                    row[j] = values[j];
                dataTable.Rows.Add(row);

                var tableRow = new TableRow
                {
                    Cells = new List<string>(values),
                    OriginalRowIndex = i,
                    PhysicalStartLine = records[i].StartLine,
                    PhysicalEndLine = records[i].EndLine
                };
                fileData.Rows.Add(tableRow);
            }

            return (dataTable, fileData);
        }

        /// <summary>统一的 Excel 解析（不涉及缓存），UseHeaderRow=false</summary>
        internal static (DataTable, FileTableData) ParseExcelFile(string fullPath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using var fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(fs);
            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
            });

            var dataTable = new DataTable();
            var fileData = new FileTableData
            {
                FilePath = fullPath,
                Type = TableFileType.Excel
            };

            if (dataSet.Tables.Count == 0)
                return (dataTable, fileData);

            var sheet = dataSet.Tables[0];
            int maxCols = sheet.Columns.Count;

            for (int i = 0; i < maxCols; i++)
            {
                var colName = GetExcelColumnName(i);
                dataTable.Columns.Add(colName, typeof(string));
                // 版本对比/冲突解决按 ColumnNames.Count 克隆冻结行，留空会让表头行/类型行整行变空白
                fileData.ColumnNames.Add(colName);
            }

            for (int r = 0; r < sheet.Rows.Count; r++)
            {
                var newRow = dataTable.NewRow();
                var cells = new List<string>();
                for (int c = 0; c < maxCols; c++)
                {
                    var val = sheet.Rows[r][c]?.ToString() ?? "";
                    newRow[c] = val;
                    cells.Add(val);
                }
                dataTable.Rows.Add(newRow);

                var tableRow = new TableRow
                {
                    Cells = cells,
                    OriginalRowIndex = r,
                    // ExcelDataReader 不暴露源表行的真实行号，只能给到"数据区第 r+1 行"；
                    // 配置表数据区基本都从 A1 起，等高于物理行号。
                    PhysicalStartLine = r + 1,
                    PhysicalEndLine = r + 1
                };
                fileData.Rows.Add(tableRow);
            }

            return (dataTable, fileData);
        }

        // ============ 无状态：写盘 ============

        /// <summary>写入 CSV 到磁盘（不写列名，有重试）</summary>
        private static bool WriteCsvToDisk(string fullPath, DataTable dataTable, char delimiter)
        {
            int retries = 3;
            while (retries > 0)
            {
                try
                {
                    var content = BuildCsvContent(dataTable, delimiter);
                    var bytes = Encoding.UTF8.GetBytes(content);
                    using (var fs = new FileStream(fullPath, FileMode.Create,
                                   FileAccess.Write, FileShare.ReadWrite))
                    {
                        fs.Write(bytes, 0, bytes.Length);
                        fs.Flush();
                    }
                    return true;
                }
                catch (IOException)
                {
                    retries--;
                    if (retries > 0)
                        System.Threading.Thread.Sleep(300);
                }
            }
            return false;
        }

        /// <summary>写入 Excel 到磁盘（不写列名，有备份还原+重试）</summary>
        private static bool WriteExcelToDisk(string fullPath, DataTable dataTable)
        {
            var backupPath = fullPath + "." + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".bak";
            int retries = 3;
            while (retries > 0)
            {
                try
                {
                    if (File.Exists(fullPath))
                    {
                        File.Copy(fullPath, backupPath, overwrite: true);
                        File.Delete(fullPath);
                    }

                    MiniExcelLibs.MiniExcel.SaveAs(fullPath, dataTable, printHeader: false);

                    if (File.Exists(backupPath))
                        File.Delete(backupPath);

                    return true;
                }
                catch (IOException)
                {
                    retries--;
                    if (retries > 0)
                        System.Threading.Thread.Sleep(300);
                }
                catch
                {
                    if (File.Exists(backupPath))
                    {
                        try
                        {
                            if (File.Exists(fullPath))
                                File.Delete(fullPath);
                            File.Move(backupPath, fullPath);
                        }
                        catch { }
                    }
                    return false;
                }
            }
            return false;
        }

        // ============ 辅助方法 ============

        private static char DetectDelimiter(string content)
        {
            var records = SplitCsvRecords(content);
            if (records.Count == 0) return ',';
            var firstLine = records[0].Text;
            var counts = new Dictionary<char, int>
            {
                { ',', firstLine.Count(c => c == ',') },
                { GlobalConst.TabChar, firstLine.Count(c => c == GlobalConst.TabChar) },
                { ';', firstLine.Count(c => c == ';') },
                { '|', firstLine.Count(c => c == '|') }
            };
            return counts.OrderByDescending(kv => kv.Value).First().Key;
        }

        internal static List<(string Text, int StartLine, int EndLine)> SplitCsvRecords(string content)
        {
            var records = new List<(string Text, int StartLine, int EndLine)>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            int lineNo = 1;       // 正在扫描的物理行号（1-based）
            int recordStart = 1;  // 当前记录的起始物理行号
            for (int i = 0; i < content.Length; i++)
            {
                char c = content[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < content.Length && content[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                        continue;
                    }
                    inQuotes = !inQuotes;
                    sb.Append(c);
                }
                else if (!inQuotes && (c == GlobalConst.LfChar || c == GlobalConst.CrChar))
                {
                    if (c == GlobalConst.CrChar && i + 1 < content.Length && content[i + 1] == GlobalConst.LfChar) i++;
                    records.Add((sb.ToString(), recordStart, lineNo));
                    sb.Clear();
                    lineNo++;
                    recordStart = lineNo;
                }
                else
                {
                    sb.Append(c);
                    // 引号内的换行属于字段内容，不做分割，但它是真实存在的一行物理文本
                    if (inQuotes && (c == GlobalConst.LfChar ||
                        (c == GlobalConst.CrChar && (i + 1 >= content.Length || content[i + 1] != GlobalConst.LfChar))))
                    {
                        lineNo++;
                    }
                }
            }
            if (sb.Length > 0) records.Add((sb.ToString(), recordStart, lineNo));
            return records;
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
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else inQuotes = !inQuotes;
                }
                else if (c == delimiter && !inQuotes)
                {
                    result.Add(sb.ToString().Trim());
                    sb.Clear();
                }
                else sb.Append(c);
            }
            result.Add(sb.ToString().Trim());
            return result.ToArray();
        }

        private static string GetExcelColumnName(int columnIndex)
        {
            var result = new StringBuilder();
            int index = columnIndex;
            do { result.Insert(0, (char)('A' + (index % 26))); index = index / 26 - 1; } while (index >= 0);
            return result.ToString();
        }

        private static string BuildCsvContent(DataTable dataTable, char delimiter)
        {
            var sb = new StringBuilder();
            foreach (DataRow row in dataTable.Rows)
            {
                var values = new List<string>();
                foreach (var item in row.ItemArray)
                {
                    var value = item?.ToString() ?? "";
                    if (value.Contains(delimiter) || value.Contains('"') || value.Contains(GlobalConst.LfChar) || value.Contains(GlobalConst.CrChar))
                        value = "\"" + value.Replace("\"", "\"\"") + "\"";
                    values.Add(value);
                }
                sb.AppendLine(string.Join(delimiter.ToString(), values));
            }
            return sb.ToString();
        }

        private static LinkLookupEntry BuildLinkLookupFromMaster(
            FileTableData fileData, int targetColIndex, FieldLinkConfig? fieldLink)
        {
            var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var displayNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            bool hasDisplayCol = fieldLink != null
                && fieldLink.TargetDisplayColumnIndex >= 0
                && fieldLink.TargetDisplayColumnIndex != targetColIndex;

            var columnNames = new List<string>();
            if (fileData.ColumnNames.Count > 0)
                columnNames = fileData.ColumnNames;

            // ★ 只构建红色高亮必需的 Values + DisplayNames
            //    RowData / RowRawStrings 不再立即构建（实测 tooltip 不读取这两个字段，
            //    且大表场景下会拷贝整行导致内存浪费 ~60MB）
            //    保留空字典以保持 LinkLookupEntry 字段兼容性
            foreach (var row in fileData.Rows)
            {
                if (targetColIndex >= 0 && targetColIndex < row.Cells.Count)
                {
                    var idVal = row.Cells[targetColIndex].Trim();
                    if (!string.IsNullOrEmpty(idVal))
                    {
                        values.Add(idVal);
                        if (hasDisplayCol && fieldLink!.TargetDisplayColumnIndex < row.Cells.Count)
                        {
                            var displayVal = row.Cells[fieldLink.TargetDisplayColumnIndex].Trim();
                            if (!string.IsNullOrEmpty(displayVal) && !displayNames.ContainsKey(idVal))
                                displayNames[idVal] = displayVal;
                        }
                    }
                }
            }

            return new LinkLookupEntry
            {
                Values = values,
                DisplayNames = displayNames,
                TargetFile = fieldLink?.TargetFile ?? "",
                TargetColumnName = fieldLink?.TargetColumnName ?? "",
                DisplayColumnName = fieldLink?.TargetDisplayColumnName ?? "",
                RowData = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase),
                RowRawStrings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                ColumnNames = columnNames
            };
        }

        private static string BuildLinkKey(
            string sourceRel, int sourceColIndex,
            string targetRel, int targetColIndex,
            ColumnLinkConfig? linkConfig, int fieldIndex = -1)
        {
            string key = $"{sourceRel}|{sourceColIndex}->field{fieldIndex}->{targetRel}:{targetColIndex}";
            if (linkConfig != null && linkConfig.ValueExtractionEnabled)
            {
                key += "|extract:";
                if (linkConfig.ExtractSteps.Count > 0)
                {
                    key += string.Join("_", linkConfig.ExtractSteps
                        .Select(s => $"{s.Delimiter}:{s.PartIndex}"));
                }
                else
                {
                    key += "_default";
                }
            }
            return key;
        }

        private void ClearLinkLookupsForFile(string fileKey)
        {
            var fileName = Path.GetFileName(fileKey);
            string marker = "->" + fileName + ":";
            var keysToRemove = _linkLookups.Keys
                .Where(k => k.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
            foreach (var k in keysToRemove)
                _linkLookups.Remove(k);
        }

        /// <summary>深拷贝 FileTableData（版本对比用：交给视图前克隆，避免差异标记写回污染缓存）</summary>
        internal static FileTableData CloneFileData(FileTableData source)
        {
            var clone = new FileTableData
            {
                FilePath = source.FilePath,
                Type = source.Type,
                Delimiter = source.Delimiter,
                ColumnNames = new List<string>(source.ColumnNames)
            };
            foreach (var row in source.Rows)
            {
                clone.Rows.Add(new TableRow
                {
                    Cells = new List<string>(row.Cells),
                    OriginalRowIndex = row.OriginalRowIndex,
                    PhysicalStartLine = row.PhysicalStartLine,
                    PhysicalEndLine = row.PhysicalEndLine,
                    IsDiffRow = row.IsDiffRow,
                    DiffColumns = row.DiffColumns != null ? new HashSet<int>(row.DiffColumns) : null,
                    DiffType = row.DiffType,
                    IsBlankRow = row.IsBlankRow
                });
            }
            return clone;
        }

        // ============ 内部数据类 ============

        private class CacheEntry
        {
            public DataTable MasterDataTable { get; set; } = null!;
            public FileTableData MasterFileData { get; set; } = null!;
            public DateTime LastWriteTime { get; set; }
            public long FileSize { get; set; }
        }
    }
}
