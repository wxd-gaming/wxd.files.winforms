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
    /// 文件数据缓存（AppDomain 级别）。
    /// key = 全小写绝对路径（如 "c:\work\a.csv"）。
    /// 30分钟无访问自动过期，过期返回旧数据 + 后台异步刷新。
    /// 关联查找（LinkLookup）key 保持相对路径格式，与 ColumnLinkConfig 一致。
    /// 
    /// 并发设计：
    /// - 不同文件使用独立锁（_fileLocks），可并行加载/访问
    /// - 同一文件的并发请求通过 per-file lock + double-check 保证只加载一次
    /// - _entries 使用 ConcurrentDictionary，字典操作本身线程安全
    /// - LinkLookup 相关操作使用独立的 _linkLock，与文件缓存互不干扰
    /// </summary>
    public static class FileDataCache
    {
        /// <summary>缓存条目（key = NormalizePath(fullPath)）</summary>
        private static readonly ConcurrentDictionary<string, CacheEntry> _entries =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>关联查找缓存（key 沿用现有格式："sourceRel|srcCol->targetRel:targetCol"）</summary>
        private static readonly Dictionary<string, LinkLookupEntry> _linkLookups =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>LinkLookup 专用锁（跨文件关联，需全局视角）</summary>
        private static readonly object _linkLock = new();

        /// <summary>每个文件的独立锁对象（保证同一文件的 load 不重入）</summary>
        private static readonly ConcurrentDictionary<string, object> _fileLocks =
            new(StringComparer.OrdinalIgnoreCase);

        private static readonly TimeSpan ExpireDuration = TimeSpan.FromMinutes(30);

        /// <summary>正在后台构建关联 lookup 的源文件相对路径集合（防重入）</summary>
        private static readonly HashSet<string> _buildingLinkLookups = new(StringComparer.OrdinalIgnoreCase);

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

        // ============ 公共 API：读取（走缓存） ============

        /// <summary>读取文件（走缓存）。自动判断 CSV/Excel。返回深拷贝副本。</summary>
        public static (DataTable, FileTableData) LoadFile(string filePath)
        {
            var key = NormalizePath(filePath);
            var entry = GetOrLoadEntry(key, filePath);
            entry.Touch();
            return (entry.MasterDataTable.Copy(), CloneFileData(entry.MasterFileData));
        }

        /// <summary>获取 DataTable 副本（走缓存）</summary>
        public static DataTable GetDataTableCopy(string filePath)
        {
            var key = NormalizePath(filePath);
            var entry = GetOrLoadEntry(key, filePath);
            entry.Touch();
            return entry.MasterDataTable.Copy();
        }

        /// <summary>获取 FileTableData 副本（走缓存）</summary>
        public static FileTableData GetFileDataCopy(string filePath)
        {
            var key = NormalizePath(filePath);
            var entry = GetOrLoadEntry(key, filePath);
            entry.Touch();
            return CloneFileData(entry.MasterFileData);
        }

        /// <summary>获取只读 DataTable（走缓存，返回引用，调用方不应修改）</summary>
        public static DataTable GetDataTableReadOnly(string filePath)
        {
            var key = NormalizePath(filePath);
            var entry = GetOrLoadEntry(key, filePath);
            entry.Touch();
            return entry.MasterDataTable;
        }

        /// <summary>获取只读 FileTableData（走缓存，返回引用）</summary>
        public static FileTableData GetFileDataReadOnly(string filePath)
        {
            var key = NormalizePath(filePath);
            var entry = GetOrLoadEntry(key, filePath);
            entry.Touch();
            return entry.MasterFileData;
        }

        /// <summary>确保文件已加载到缓存（预加载用）</summary>
        public static void EnsureLoaded(string filePath)
        {
            var key = NormalizePath(filePath);
            GetOrLoadEntry(key, filePath);
        }

        // ============ 公共 API：读取（不走缓存） ============

        /// <summary>读取文件（不走缓存）。自动判断 CSV/Excel。返回深拷贝副本。</summary>
        public static (DataTable, FileTableData) LoadFileDirect(string filePath)
        {
            return ParseFile(filePath);
        }

        // ============ 公共 API：写入 ============

        /// <summary>写入文件（走缓存：写入后自动更新内存缓存）。自动判断 CSV/Excel。</summary>
        public static bool SaveFile(string filePath, DataTable dataTable, char delimiter = ',')
        {
            return SaveFileInternal(filePath, dataTable, delimiter, updateCache: true);
        }

        /// <summary>写入文件（不走缓存）。自动判断 CSV/Excel。</summary>
        public static bool SaveFileDirect(string filePath, DataTable dataTable, char delimiter = ',')
        {
            return SaveFileInternal(filePath, dataTable, delimiter, updateCache: false);
        }

        /// <summary>统一写入逻辑</summary>
        private static bool SaveFileInternal(string fullPath, DataTable dataTable, char delimiter, bool updateCache)
        {
            var ext = Path.GetExtension(fullPath).ToLowerInvariant();
            bool success;
            if (ext == ".csv")
                success = WriteCsvToDisk(fullPath, dataTable, delimiter);
            else
                success = WriteExcelToDisk(fullPath, dataTable);

            if (success && updateCache)
            {
                var key = NormalizePath(fullPath);
                var fileLock = _fileLocks.GetOrAdd(key, _ => new object());
                lock (fileLock)
                {
                    var fi = new FileInfo(fullPath);
                    var fileData = BuildFileDataFromDataTable(dataTable, fullPath);
                    _entries[key] = new CacheEntry
                    {
                        MasterDataTable = dataTable.Copy(),
                        MasterFileData = fileData,
                        LastWriteTime = fi.Exists ? fi.LastWriteTimeUtc : DateTime.UtcNow,
                        FileSize = fi.Exists ? fi.Length : 0,
                        LastAccessTime = DateTime.UtcNow
                    };
                    // 清除该文件相关的旧 linkLookup（下次访问时懒重建）
                    ClearLinkLookupsForFile(key, Path.GetDirectoryName(fullPath)!);
                }
            }
            return success;
        }

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

        // ============ 缓存管理 ============

        /// <summary>清除指定文件的缓存</summary>
        public static void InvalidateFile(string filePath)
        {
            var key = NormalizePath(filePath);
            _entries.TryRemove(key, out _);
            _fileLocks.TryRemove(key, out _);
            ClearLinkLookupsForFile(key, Path.GetDirectoryName(filePath)!);
        }

        /// <summary>清除所有缓存</summary>
        public static void ClearAll()
        {
            _entries.Clear();
            _fileLocks.Clear();
            lock (_linkLock)
            {
                _linkLookups.Clear();
                _buildingLinkLookups.Clear();
            }
        }

        // ============ 关联查找（LinkLookup，key 保持相对路径格式） ============

        /// <summary>获取关联查找条目（只读，懒构建），按字段索引查找</summary>
        public static LinkLookupEntry? GetLinkLookup(
            string rootDir, string sourceRel, int sourceColIndex,
            string targetRel, int targetColIndex,
            FieldLinkConfig? fieldLink, int fieldIndex)
        {
            string key = BuildLinkKey(sourceRel, sourceColIndex, targetRel, targetColIndex, null, fieldIndex);

            LinkLookupEntry? lookup;
            lock (_linkLock)
            {
                if (_linkLookups.TryGetValue(key, out lookup))
                    return lookup;
            }

            var targetFullPath = Path.Combine(rootDir, targetRel);
            var targetKey = NormalizePath(targetFullPath);
            var entry = GetOrLoadEntry(targetKey, targetFullPath);
            lookup = BuildLinkLookupFromMaster(entry.MasterFileData,
                targetColIndex, fieldLink);
            lock (_linkLock)
            {
                _linkLookups[key] = lookup;
            }
            return lookup;
        }

        /// <summary>获取所有已缓存的关联查找条目</summary>
        public static Dictionary<string, LinkLookupEntry> GetLookups()
        {
            lock (_linkLock)
            {
                return new Dictionary<string, LinkLookupEntry>(_linkLookups, StringComparer.OrdinalIgnoreCase);
            }
        }

        /// <summary>检查指定源文件的关联 lookup 是否正在后台构建中</summary>
        public static bool IsLinkLookupBuilding(string sourceRel)
        {
            lock (_linkLock) { return _buildingLinkLookups.Contains(sourceRel); }
        }

        /// <summary>异步构建所有关联配置的 lookup</summary>
        public static async Task BuildLinkLookupsAsync(
            string rootDir, List<ColumnLinkConfig> columnLinks, string sourceFilePath,
            CancellationToken cancellationToken = default)
        {
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
                    cancellationToken.ThrowIfCancellationRequested();

                    if (link.FieldLinks == null || link.FieldLinks.Count == 0)
                        continue;

                    foreach (var fl in link.FieldLinks)
                    {
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
                            if (!_linkLookups.ContainsKey(key))
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

        // ============ 内部实现 ============

        /// <summary>
        /// 获取或加载条目（并发安全）。
        /// 快速路径：缓存命中且文件未变 → 直接返回（无锁）。
        /// 慢路径：获取该文件独立锁 → double-check → 加载解析。
        /// </summary>
        private static CacheEntry GetOrLoadEntry(string key, string fullPath)
        {
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
                        ClearLinkLookupsForFile(key, Path.GetDirectoryName(fullPath)!);
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
                FileSize = fi.Length,
                LastAccessTime = DateTime.UtcNow
            };
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
            var lines = SplitCsvLines(content);
            var dataTable = new DataTable();
            var fileData = new FileTableData
            {
                FilePath = fullPath,
                Type = TableFileType.Csv,
                Delimiter = delimiter
            };

            if (lines.Count == 0)
                return (dataTable, fileData);

            var headers = ParseCsvLine(lines[0], delimiter);
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

            for (int i = 0; i < lines.Count; i++)
            {
                var values = ParseCsvLine(lines[i], delimiter);
                var row = dataTable.NewRow();
                for (int j = 0; j < Math.Min(values.Length, headers.Length); j++)
                    row[j] = values[j];
                dataTable.Rows.Add(row);

                var tableRow = new TableRow
                {
                    Cells = new List<string>(values),
                    OriginalRowIndex = i
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
                dataTable.Columns.Add(GetExcelColumnName(i), typeof(string));

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
                    OriginalRowIndex = r
                };
                fileData.Rows.Add(tableRow);
            }

            return (dataTable, fileData);
        }

        /// <summary>从 DataTable 构建 FileTableData</summary>
        private static FileTableData BuildFileDataFromDataTable(DataTable dataTable, string filePath)
        {
            var fileData = new FileTableData
            {
                FilePath = filePath,
                Type = TableFileType.Csv,
                Delimiter = ','
            };
            for (int c = 0; c < dataTable.Columns.Count; c++)
                fileData.ColumnNames.Add(dataTable.Columns[c].ColumnName);
            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                var cells = new List<string>();
                for (int c = 0; c < dataTable.Columns.Count; c++)
                    cells.Add(dataTable.Rows[r][c]?.ToString() ?? "");
                fileData.Rows.Add(new TableRow
                {
                    Cells = cells,
                    OriginalRowIndex = r
                });
            }
            return fileData;
        }

        // ============ 过期/刷新 ============

        /// <summary>后台异步刷新所有已缓存的条目（由外部定时器触发）</summary>
        internal static void RefreshAllInBackground()
        {
            try
            {
                var now = DateTime.UtcNow;
                var expiredKeys = _entries
                    .Where(kvp => (now - kvp.Value.LastAccessTime) > ExpireDuration)
                    .Select(kvp => kvp.Key)
                    .ToList();

                foreach (var key in expiredKeys)
                {
                    try
                    {
                        var fileLock = _fileLocks.GetOrAdd(key, _ => new object());
                        lock (fileLock)
                        {
                            if (!_entries.TryGetValue(key, out var oldEntry))
                                continue;

                            var (dt, fd) = ParseFile(key);
                            _entries[key] = new CacheEntry
                            {
                                MasterDataTable = dt,
                                MasterFileData = fd,
                                LastWriteTime = oldEntry.LastWriteTime,
                                FileSize = oldEntry.FileSize,
                                LastAccessTime = oldEntry.LastAccessTime
                            };
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        // ============ 辅助方法 ============

        private static char DetectDelimiter(string content)
        {
            var lines = SplitCsvLines(content);
            if (lines.Count == 0) return ',';
            var firstLine = lines[0];
            var counts = new Dictionary<char, int>
            {
                { ',', firstLine.Count(c => c == ',') },
                { GlobalConst.TabChar, firstLine.Count(c => c == GlobalConst.TabChar) },
                { ';', firstLine.Count(c => c == ';') },
                { '|', firstLine.Count(c => c == '|') }
            };
            return counts.OrderByDescending(kv => kv.Value).First().Key;
        }

        private static List<string> SplitCsvLines(string content)
        {
            var lines = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
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
                    lines.Add(sb.ToString());
                    sb.Clear();
                }
                else { sb.Append(c); }
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

        private static void ClearLinkLookupsForFile(string fileKey, string dirPath)
        {
            var fileName = Path.GetFileName(fileKey);
            var keysToRemove = _linkLookups.Keys
                .Where(k => k.IndexOf("->" + fileName + ":", StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
            foreach (var k in keysToRemove)
                _linkLookups.Remove(k);
        }

        private static FileTableData CloneFileData(FileTableData source)
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
            public DateTime LastAccessTime { get; set; }

            public bool IsExpired => (DateTime.UtcNow - LastAccessTime) > ExpireDuration;
            internal void Touch() => LastAccessTime = DateTime.UtcNow;
        }
    }
}
