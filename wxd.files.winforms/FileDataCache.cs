using System;
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
    /// 文件数据缓存总入口（static，AppDomain 级别）。
    /// 按根目录隔离，每个目录一个独立的 DirectoryCache 实例。
    /// 30分钟无访问自动过期，过期返回旧数据 + 后台异步刷新。
    /// </summary>
    public static class FileDataCache
    {
        /// <summary>key = 根目录绝对路径（如 "c:\work1"）</summary>
        private static readonly Dictionary<string, DirectoryCache> _caches =
            new(StringComparer.OrdinalIgnoreCase);

        private static readonly object _lock = new();

        /// <summary>获取或创建指定根目录的缓存容器</summary>
        public static DirectoryCache GetDirectoryCache(string rootDir)
        {
            lock (_lock)
            {
                if (_caches.TryGetValue(rootDir, out var cache))
                {
                    cache.Touch();

                    // 过期不删除，不阻塞，后台异步刷新
                    if (cache.IsExpired && !cache.IsRefreshing)
                    {
                        cache.IsRefreshing = true;
                        var capturedCache = cache;
                        Task.Run(() => capturedCache.RefreshInBackground());
                    }

                    return cache; // 立即返回旧实例，UI 不卡
                }

                var newCache = new DirectoryCache(rootDir);
                _caches[rootDir] = newCache;
                return newCache;
            }
        }

        /// <summary>清除指定目录的缓存</summary>
        public static void InvalidateDirectory(string rootDir)
        {
            lock (_lock) { _caches.Remove(rootDir); }
        }

        /// <summary>清除所有缓存</summary>
        public static void ClearAll()
        {
            lock (_lock) { _caches.Clear(); }
        }
    }

    /// <summary>
    /// 单个根目录的文件数据缓存容器。
    /// 内部使用相对路径，与 ColumnLinkConfig 保持一致。
    /// 30分钟无访问自动过期。
    /// </summary>
    public class DirectoryCache
    {
        private readonly string _rootDir;
        private readonly object _lock = new();

        /// <summary>文件数据缓存（key = 相对路径，如 "a.csv"）</summary>
        private readonly Dictionary<string, CacheEntry> _fileEntries =
            new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 关联查找缓存（key 沿用现有格式："sourceRel|srcCol->targetRel:targetCol"）
        /// </summary>
        private readonly Dictionary<string, LinkLookupEntry> _linkLookups =
            new(StringComparer.OrdinalIgnoreCase);

        private DateTime _lastAccessTime = DateTime.UtcNow;
        private static readonly TimeSpan ExpireDuration = TimeSpan.FromMinutes(30);

        /// <summary>是否正在后台刷新（防重入）</summary>
        public volatile bool IsRefreshing = false;

        /// <summary>正在后台构建关联 lookup 的源文件集合（防重入）</summary>
        private readonly HashSet<string> _buildingLinkLookups = new(StringComparer.OrdinalIgnoreCase);

        public DirectoryCache(string rootDir)
        {
            _rootDir = rootDir ?? throw new ArgumentNullException(nameof(rootDir));
        }

        public bool IsExpired => (DateTime.UtcNow - _lastAccessTime) > ExpireDuration;

        internal void Touch() => _lastAccessTime = DateTime.UtcNow;

        // ============ 公共 API ============

        /// <summary>获取 DataTable 副本（给 TableBrowserControl 用）</summary>
        public DataTable GetDataTableCopy(string relativePath)
        {
            Touch();
            var entry = GetOrLoadEntry(relativePath);
            return entry.MasterDataTable.Copy();
        }

        /// <summary>获取 FileTableData 副本</summary>
        public FileTableData GetFileDataCopy(string relativePath)
        {
            Touch();
            var entry = GetOrLoadEntry(relativePath);
            return CloneFileData(entry.MasterFileData);
        }

        /// <summary>
        /// 获取文件的只读 DataTable（用于 ToolTip / 颜色检查 / 内容搜索，不需要拷贝）
        /// </summary>
        public DataTable GetDataTableReadOnly(string relativePath)
        {
            Touch();
            var entry = GetOrLoadEntry(relativePath);
            return entry.MasterDataTable;
        }

        /// <summary>
        /// 获取文件的只读 FileTableData（用于 ToolTip / 内容搜索，不需要拷贝）
        /// </summary>
        public FileTableData GetFileDataReadOnly(string relativePath)
        {
            Touch();
            var entry = GetOrLoadEntry(relativePath);
            return entry.MasterFileData;
        }

        /// <summary>确保文件已加载到缓存（预加载用）</summary>
        public void EnsureLoaded(string relativePath)
        {
            Touch();
            GetOrLoadEntry(relativePath);
        }

        /// <summary>
        /// 获取关联查找条目（ToolTip / 颜色检查用，只读）
        /// 如果缓存中不存在，从已缓存的 FileData 构建。
        /// </summary>
        public LinkLookupEntry? GetLinkLookup(
            string sourceRel, int sourceColIndex,
            string targetRel, int targetColIndex,
            ColumnLinkConfig? linkConfig)
        {
            Touch();

            string key = BuildLinkKey(sourceRel, sourceColIndex, targetRel, targetColIndex, linkConfig);

            LinkLookupEntry? lookup;
            lock (_lock)
            {
                if (_linkLookups.TryGetValue(key, out lookup))
                    return lookup;
            }

            // 懒构建：目标文件可能尚未加载到缓存
            var targetEntry = GetOrLoadEntry(targetRel);
            lookup = BuildLinkLookupFromMaster(targetEntry.MasterFileData,
                targetColIndex, linkConfig);
            lock (_lock)
            {
                _linkLookups[key] = lookup;
            }
            return lookup;
        }

        /// <summary>获取所有已缓存的关联查找条目（供外部遍历使用）</summary>
        public Dictionary<string, LinkLookupEntry> GetLookups()
        {
            Touch();
            lock (_lock)
            {
                return new Dictionary<string, LinkLookupEntry>(_linkLookups, StringComparer.OrdinalIgnoreCase);
            }
        }

        /// <summary>检查指定源文件的关联 lookup 是否正在后台构建中</summary>
        public bool IsLinkLookupBuilding(string sourceRel)
        {
            lock (_lock) { return _buildingLinkLookups.Contains(sourceRel); }
        }

        /// <summary>异步构建所有关联配置的 lookup（不阻塞调用方，内部防重入）</summary>
        public async Task BuildLinkLookupsAsync(
            List<ColumnLinkConfig> columnLinks, string sourceFilePath,
            CancellationToken cancellationToken = default)
        {
            Touch();

            string sourceRel = ToRelative(sourceFilePath);

            // ★ 防重入检查
            lock (_lock)
            {
                if (_buildingLinkLookups.Contains(sourceRel))
                    return; // 已有后台任务在构建

                _buildingLinkLookups.Add(sourceRel);

                // 清除当前源文件的旧缓存
                string prefix = sourceRel + "|";
                var keysToRemove = _linkLookups.Keys
                    .Where(k => k.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                foreach (var k in keysToRemove)
                    _linkLookups.Remove(k);
            }

            try
            {
                // 筛选属于当前源文件的关联配置
                var relevantLinks = columnLinks
                    .Where(l => l.SourceFile.EndsWith(Path.GetFileName(sourceFilePath),
                        StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var link in relevantLinks)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    string targetFullPath = Path.Combine(_rootDir, link.TargetFile);
                    if (!File.Exists(targetFullPath)) continue;

                    // ★ 使用 Task.Run 将文件 IO 放到线程池，不阻塞调用方
                    CacheEntry entry = await Task.Run(() =>
                    {
                        lock (_lock) { return GetOrLoadEntry(link.TargetFile); }
                    }, cancellationToken);

                    string key = BuildLinkKey(sourceRel, link.SourceColumnIndex,
                        link.TargetFile, link.TargetColumnIndex, link);

                    // ★ 将 lookup 构建也放到线程池
                    var lookup = await Task.Run(() =>
                        BuildLinkLookupFromMaster(entry.MasterFileData,
                            link.TargetColumnIndex, link), cancellationToken);

                    lock (_lock)
                    {
                        if (!_linkLookups.ContainsKey(key))
                            _linkLookups[key] = lookup;
                    }
                }
            }
            finally
            {
                lock (_lock) { _buildingLinkLookups.Remove(sourceRel); }
            }
        }

        /// <summary>
        /// 用内存中的 DataTable 更新缓存主副本（写入磁盘成功后调用）
        /// </summary>
        public void UpdateCacheFromMemory(string relativePath, DataTable updatedTable, FileTableData updatedFileData)
        {
            lock (_lock)
            {
                var fullPath = Path.Combine(_rootDir, relativePath);
                var fi = new FileInfo(fullPath);

                var entry = new CacheEntry
                {
                    MasterDataTable = updatedTable.Copy(),
                    MasterFileData = CloneFileData(updatedFileData),
                    LastWriteTime = fi.Exists ? fi.LastWriteTimeUtc : DateTime.UtcNow,
                    FileSize = fi.Exists ? fi.Length : 0
                };

                _fileEntries[relativePath] = entry;

                // 清除该文件相关的旧 linkLookup（下次访问时懒重建）
                ClearLinkLookupsForFile(relativePath);
            }
        }

        /// <summary>
        /// 尝试回写 CSV 到磁盘，失败时返回 false
        /// 使用 FileShare.ReadWrite + 重试机制
        /// </summary>
        public bool TrySaveCsvToDisk(string relativePath, DataTable dataTable, char delimiter)
        {
            var fullPath = Path.Combine(_rootDir, relativePath);

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

        /// <summary>
        /// 尝试回写 Excel 到磁盘，失败时返回 false
        /// </summary>
        public bool TrySaveExcelToDisk(string relativePath, DataTable dataTable)
        {
            var fullPath = Path.Combine(_rootDir, relativePath);
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
                    // 写入失败，尝试还原
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

        /// <summary>将相对路径转为完整路径</summary>
        public string ToFullPath(string relativePath)
            => Path.Combine(_rootDir, relativePath);

        /// <summary>将绝对路径转为相对路径</summary>
        public string ToRelative(string absolutePath)
        {
            if (absolutePath.StartsWith(_rootDir, StringComparison.OrdinalIgnoreCase))
                return absolutePath.Substring(_rootDir.Length).TrimStart(Path.DirectorySeparatorChar);
            return Path.GetFileName(absolutePath);
        }

        /// <summary>后台异步刷新：重新加载所有已缓存的文件</summary>
        internal void RefreshInBackground()
        {
            try
            {
                var newEntries = new Dictionary<string, CacheEntry>(StringComparer.OrdinalIgnoreCase);
                var newLookups = new Dictionary<string, LinkLookupEntry>(StringComparer.OrdinalIgnoreCase);

                lock (_lock)
                {
                    // 重新加载所有已缓存的文件
                    foreach (var kvp in _fileEntries)
                    {
                        try
                        {
                            var entry = LoadFile(kvp.Key);
                            newEntries[kvp.Key] = entry;
                        }
                        catch { } // 单个文件加载失败不影响其他
                    }

                    // 保留旧的 lookups（不做重建，因为依赖 sourceRel 映射关系可能已变）
                    // 只需更新 _lastAccessTime 续期
                }

                lock (_lock)
                {
                    foreach (var kvp in newEntries)
                        _fileEntries[kvp.Key] = kvp.Value;

                    _lastAccessTime = DateTime.UtcNow;
                }
            }
            finally
            {
                IsRefreshing = false;
            }
        }

        // ============ 内部实现 ============

        private CacheEntry GetOrLoadEntry(string relativePath)
        {
            lock (_lock)
            {
                if (_fileEntries.TryGetValue(relativePath, out var entry))
                {
                    string fullPath = Path.Combine(_rootDir, relativePath);
                    var fi = new FileInfo(fullPath);
                    if (fi.Exists &&
                        (fi.LastWriteTimeUtc != entry.LastWriteTime || fi.Length != entry.FileSize))
                    {
                        // 文件被外部修改，失效重建
                        _fileEntries.Remove(relativePath);
                        ClearLinkLookupsForFile(relativePath);
                    }
                    else
                    {
                        return entry;
                    }
                }

                entry = LoadFile(relativePath);
                _fileEntries[relativePath] = entry;
                return entry;
            }
        }

        private CacheEntry LoadFile(string relativePath)
        {
            var fullPath = Path.Combine(_rootDir, relativePath);
            var ext = Path.GetExtension(fullPath).ToLowerInvariant();
            var fi = new FileInfo(fullPath);

            DataTable dt;
            FileTableData fd;

            if (ext == ".csv")
            {
                (dt, fd) = LoadCsvFile(fullPath);
            }
            else // .xls / .xlsx
            {
                (dt, fd) = LoadExcelFile(fullPath);
            }

            return new CacheEntry
            {
                MasterDataTable = dt,
                MasterFileData = fd,
                LastWriteTime = fi.LastWriteTimeUtc,
                FileSize = fi.Length
            };
        }

        private (DataTable, FileTableData) LoadCsvFile(string fullPath)
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

            for (int i = 0; i < lines.Count; i++)
            {
                var values = ParseCsvLine(lines[i], delimiter);
                var row = dataTable.NewRow();
                for (int j = 0; j < Math.Min(values.Length, headers.Length); j++)
                    row[j] = values[j];
                dataTable.Rows.Add(row);

                // 构建 FileTableData
                var tableRow = new TableRow
                {
                    Cells = new List<string>(values),
                    OriginalRowIndex = i
                };
                fileData.ColumnNames = new List<string>(headers);
                fileData.Rows.Add(tableRow);
            }

            return (dataTable, fileData);
        }

        private (DataTable, FileTableData) LoadExcelFile(string fullPath)
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

        // ============ CSV 解析工具方法 ============

        private static char DetectDelimiter(string content)
        {
            var lines = SplitCsvLines(content);
            if (lines.Count == 0) return ',';
            var firstLine = lines[0];
            var counts = new Dictionary<char, int>
            {
                { ',', firstLine.Count(c => c == ',') },
                { '\t', firstLine.Count(c => c == '\t') },
                { ';', firstLine.Count(c => c == ';') },
                { '|', firstLine.Count(c => c == '|') }
            };
            return counts.OrderByDescending(kv => kv.Value).First().Key;
        }

        internal static List<string> SplitCsvLines(string content)
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
                else if (!inQuotes && (c == '\n' || c == '\r'))
                {
                    if (c == '\r' && i + 1 < content.Length && content[i + 1] == '\n') i++;
                    lines.Add(sb.ToString());
                    sb.Clear();
                }
                else { sb.Append(c); }
            }
            if (sb.Length > 0) lines.Add(sb.ToString());
            return lines;
        }

        internal static string[] ParseCsvLine(string line, char delimiter)
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

        // ============ 辅助方法 ============

        private static string GetExcelColumnName(int columnIndex)
        {
            var result = new StringBuilder();
            int index = columnIndex;
            do { result.Insert(0, (char)('A' + (index % 26))); index = index / 26 - 1; } while (index >= 0);
            return result.ToString();
        }

        private static LinkLookupEntry BuildLinkLookupFromMaster(
            FileTableData fileData, int targetColIndex, ColumnLinkConfig? linkConfig)
        {
            var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var displayNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            bool hasDisplayCol = linkConfig != null
                && linkConfig.TargetDisplayColumnIndex >= 0
                && linkConfig.TargetDisplayColumnIndex != targetColIndex;

            foreach (var row in fileData.Rows)
            {
                if (targetColIndex >= 0 && targetColIndex < row.Cells.Count)
                {
                    var idVal = row.Cells[targetColIndex].Trim();
                    if (!string.IsNullOrEmpty(idVal))
                    {
                        values.Add(idVal);
                        if (hasDisplayCol && linkConfig!.TargetDisplayColumnIndex < row.Cells.Count)
                        {
                            var displayVal = row.Cells[linkConfig.TargetDisplayColumnIndex].Trim();
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
                TargetFile = linkConfig?.TargetFile ?? "",
                TargetColumnName = linkConfig?.TargetColumnName ?? "",
                DisplayColumnName = linkConfig?.TargetDisplayColumnName ?? ""
            };
        }

        private static string BuildLinkKey(
            string sourceRel, int sourceColIndex,
            string targetRel, int targetColIndex,
            ColumnLinkConfig? linkConfig)
        {
            string key = $"{sourceRel}|{sourceColIndex}->{targetRel}:{targetColIndex}";
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

        private void ClearLinkLookupsForFile(string relativePath)
        {
            var keysToRemove = _linkLookups.Keys
                .Where(k => k.IndexOf("->" + relativePath + ":", StringComparison.OrdinalIgnoreCase) >= 0)
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

        private static string BuildCsvContent(DataTable dataTable, char delimiter)
        {
            var sb = new StringBuilder();
            foreach (DataRow row in dataTable.Rows)
            {
                var values = new List<string>();
                foreach (var item in row.ItemArray)
                {
                    var value = item?.ToString() ?? "";
                    if (value.Contains(delimiter) || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
                        value = "\"" + value.Replace("\"", "\"\"") + "\"";
                    values.Add(value);
                }
                sb.AppendLine(string.Join(delimiter.ToString(), values));
            }
            return sb.ToString();
        }

        // ============ 内部数据类 ============

        private class CacheEntry
        {
            /// <summary>原始 DataTable（只读，不对外暴露引用）</summary>
            public DataTable MasterDataTable { get; set; } = null!;

            /// <summary>原始 FileTableData（只读）</summary>
            public FileTableData MasterFileData { get; set; } = null!;

            /// <summary>文件最后写入时间（UTC，用于失效判断）</summary>
            public DateTime LastWriteTime { get; set; }

            /// <summary>文件大小（辅助失效判断）</summary>
            public long FileSize { get; set; }
        }
    }
}
