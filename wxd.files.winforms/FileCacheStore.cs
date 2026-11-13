using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace wxd.files.winforms
{
    /// <summary>
    /// 搜索结果缓存：内存 LRU + 磁盘持久化双层
    /// 
    /// 内存层：按文件路径隔离，每个文件上限 500MB（LRU 淘汰最久未访问）
    /// 磁盘层：%Temp%\wxd_files_winforms\search_cache\{文件路径MD5前16位}\{cacheKey}.cache
    /// 
    /// 所有磁盘 IO 在后台线程执行，不阻塞 UI。
    /// 关闭页签时只清内存，磁盘保留供下次打开。
    /// </summary>
    public static class FileCacheStore
    {
        // ===== 参数 =====
        /// <summary>内存 LRU 每文件上限 500MB</summary>
        private const long MaxMemoryPerFile = 500L * 1024 * 1024;
        /// <summary>磁盘每个文件最多 30 个关键词缓存</summary>
        private const int MaxDiskCacheEntries = 30;
        /// <summary>单条搜索结果超 5000 万行不缓存（防极端 OOM）</summary>
        private const int MaxSearchResultRowsToCache = 50_000_000;
        /// <summary>磁盘读每次搬这么多行（8 字节/行 ≈ 8MB），避免千万级结果一次性分配连续大块</summary>
        private const int ReadSegmentRows = 1_048_576;

        // 二进制文件头
        private static readonly byte[] MagicSearch = { (byte)'W', (byte)'S', (byte)'C', (byte)'C' }; // Wxd Search Cache Chunk
        // ★ CacheVersion=2：行号从 int(4字节) 升级为 long(8字节)，支持超过 21 亿行的超大文件
        private const int CacheVersion = 2;

        // ===== 内存 LRU =====
        /// <summary>每个文件的内存缓存条目链表（头部=最近访问，尾部=待淘汰）</summary>
        private class MemEntry
        {
            public string Key;
            public SearchResultSet Results;
            public long Bytes;
            public MemEntry(string k, SearchResultSet r) { Key = k; Results = r; Bytes = 56L + r.EstimatedBytes; }
        }
        private static readonly object _memLock = new();
        // filePath → (链表, 总字节数)
        private static readonly Dictionary<string, LinkedList<MemEntry>> _memCache = new();
        private static readonly Dictionary<string, long> _memBytes = new();

        // ===== 路径工具 =====
        private static string FileHashDir(string filePath)
        {
            using var md5 = MD5.Create();
            var bytes = md5.ComputeHash(Encoding.UTF8.GetBytes(filePath.ToLowerInvariant()));
            var sb = new StringBuilder();
            for (int i = 0; i < 8; i++) sb.Append(bytes[i].ToString("x2"));
            return sb.ToString(); // 前 16 位 hex
        }

        private static string BaseDir => Path.Combine(Path.GetTempPath(), "wxd_files_winforms", "search_cache");

        private static string FileCacheDir(string filePath) => Path.Combine(BaseDir, FileHashDir(filePath));

        private static string MetaPath(string filePath) => Path.Combine(FileCacheDir(filePath), "meta.json");

        private static string CachePath(string filePath, string cacheKey) =>
            Path.Combine(FileCacheDir(filePath), SanitizeFileName(cacheKey) + ".cache");

        private static string SanitizeFileName(string key)
        {
            var sb = new StringBuilder(key.Length);
            foreach (char c in key)
            {
                if (c == '|' || c == '\\' || c == '/' || c == ':' || c == '*' || c == '?' || c == '"' || c == '<' || c == '>' || c < 32)
                    sb.Append('_');
                else
                    sb.Append(c);
            }
            return sb.ToString();
        }

        // ===== 文件有效性检测 =====
        private class MetaInfo
        {
            public long Size { get; set; }
            public DateTime LastWrite { get; set; }
        }

        /// <summary>检查文件是否变化（大小+修改时间）。变化则清理该文件的磁盘+内存缓存。</summary>
        public static void CheckFileValid(string filePath, long size, DateTime lastWrite)
        {
            _ = Task.Run(() =>
            {
                try
                {
                    var dir = FileCacheDir(filePath);
                    var metaPath = MetaPath(filePath);
                    MetaInfo? meta = null;
                    if (File.Exists(metaPath))
                    {
                        var json = File.ReadAllText(metaPath);
                        meta = System.Text.Json.JsonSerializer.Deserialize<MetaInfo>(json);
                    }
                    if (meta != null && meta.Size == size && meta.LastWrite == lastWrite)
                    {
                        // 文件未变化，缓存有效
                        return;
                    }
                    // 文件变化（或首次），清理磁盘缓存
                    if (Directory.Exists(dir))
                    {
                        try { Directory.Delete(dir, true); } catch { }
                    }
                    // 写入新的 meta
                    Directory.CreateDirectory(dir);
                    var newMeta = new MetaInfo { Size = size, LastWrite = lastWrite };
                    File.WriteAllText(metaPath, System.Text.Json.JsonSerializer.Serialize(newMeta));
                }
                catch { }
            });
            // 内存层：文件路径未变就不清（切换文件后旧文件内存通过 UnloadFile 释放）
        }

        // ===== 读取 =====
        /// <summary>尝试获取搜索结果（先内存后磁盘）。命中返回结果，未命中返回 null。</summary>
        public static Task<SearchResultSet?> TryGetSearchResultAsync(string filePath, string cacheKey)
        {
            return Task.Run(() =>
            {
                // 1. 内存命中
                lock (_memLock)
                {
                    if (_memCache.TryGetValue(filePath, out var list))
                    {
                        var node = list.First;
                        while (node != null)
                        {
                            if (node.Value.Key == cacheKey)
                            {
                                // 移到头部（LRU）
                                list.Remove(node);
                                list.AddFirst(node);
                                // 段发布后不再改动，直接给引用即可；旧实现这里 ToList 拷一份，
                                // 千万级匹配每次命中都白拷几十 MB
                                return node.Value.Results;
                            }
                            node = node.Next;
                        }
                    }
                }
                // 2. 磁盘命中
                try
                {
                    var path = CachePath(filePath, cacheKey);
                    if (!File.Exists(path)) return null;
                    using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.SequentialScan);
                    var header = new byte[12]; // magic(4) + version(4) + count(4)
                    if (fs.Read(header, 0, 12) != 12) return null;
                    if (header[0] != MagicSearch[0] || header[1] != MagicSearch[1] || header[2] != MagicSearch[2] || header[3] != MagicSearch[3])
                        return null;
                    int version = BitConverter.ToInt32(header, 4);
                    // ★ 版本不匹配（旧版 int 布局）直接丢弃，不解析，下次搜索重建
                    if (version != CacheVersion) return null;
                    int count = BitConverter.ToInt32(header, 8);
                    if (count < 0 || count > MaxSearchResultRowsToCache) return null;
                    // ★ 分段读入：一次 count*8 的连续分配在千万级结果上就是几百 MB 尖峰
                    var set = new SearchResultSet();
                    var data = new byte[ReadSegmentRows * 8];
                    int done = 0, slot = 0;
                    while (done < count)
                    {
                        int take = Math.Min(ReadSegmentRows, count - done);
                        int need = take * 8;
                        int got = 0;
                        while (got < need)
                        {
                            int r = fs.Read(data, got, need - got);
                            if (r <= 0) break;
                            got += r;
                        }
                        if (got != need) return null;
                        var seg = new long[take];
                        Buffer.BlockCopy(data, 0, seg, 0, need);
                        set.Publish(slot++, seg);
                        done += take;
                    }
                    // 磁盘缓存只存完整扫描的产物，读回来即封口
                    set.Seal();
                    // 加载到内存 LRU
                    PutMemory(filePath, cacheKey, set);
                    return set;
                }
                catch { return null; }
            });
        }

        // ===== 写入 =====
        /// <summary>写入搜索结果（内存 + 磁盘，后台执行）</summary>
        public static Task PutSearchResultAsync(string filePath, string cacheKey, SearchResultSet results)
        {
            // 单条结果过大不缓存
            if (results.Count > MaxSearchResultRowsToCache) return Task.CompletedTask;

            // 内存层（同步，只存引用）
            PutMemory(filePath, cacheKey, results);

            // 磁盘层（后台）
            return Task.Run(() =>
            {
                try
                {
                    var dir = FileCacheDir(filePath);
                    Directory.CreateDirectory(dir);
                    var path = CachePath(filePath, cacheKey);
                    using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65536, FileOptions.SequentialScan);
                    // 头
                    fs.Write(MagicSearch, 0, 4);
                    fs.Write(BitConverter.GetBytes(CacheVersion), 0, 4);
                    fs.Write(BitConverter.GetBytes((int)results.Count), 0, 4);
                    // 数据：逐段直接落盘，不先把段拍成一个连续大数组
                    foreach (var seg in results.SnapshotSegments())
                        fs.Write(MemoryMarshal.AsBytes(seg.AsSpan()));

                    // 磁盘条数上限：超过则删最旧
                    EnforceDiskLimit(filePath);
                }
                catch { /* 缓存写入失败不影响功能 */ }
            });
        }

        /// <summary>磁盘条数上限保护：超过 MaxDiskCacheEntries 则按修改时间删最旧</summary>
        private static void EnforceDiskLimit(string filePath)
        {
            try
            {
                var dir = FileCacheDir(filePath);
                var files = Directory.GetFiles(dir, "*.cache")
                    .Select(p => new FileInfo(p))
                    .OrderBy(f => f.LastWriteTime)
                    .ToList();
                while (files.Count > MaxDiskCacheEntries)
                {
                    try { files[0].Delete(); } catch { }
                    files.RemoveAt(0);
                }
            }
            catch { }
        }

        // ===== 内存层操作 =====
        private static void PutMemory(string filePath, string cacheKey, SearchResultSet results)
        {
            lock (_memLock)
            {
                if (!_memCache.TryGetValue(filePath, out var list))
                {
                    list = new LinkedList<MemEntry>();
                    _memCache[filePath] = list;
                    _memBytes[filePath] = 0;
                }
                // 移除已有同 key 条目
                var node = list.First;
                while (node != null)
                {
                    var next = node.Next;
                    if (node.Value.Key == cacheKey)
                    {
                        _memBytes[filePath] -= node.Value.Bytes;
                        list.Remove(node);
                    }
                    node = next;
                }
                // 加入头部
                var entry = new MemEntry(cacheKey, results);
                list.AddFirst(entry);
                _memBytes[filePath] += entry.Bytes;

                // LRU 淘汰尾部
                while (_memBytes[filePath] > MaxMemoryPerFile && list.Count > 1)
                {
                    var last = list.Last!;
                    _memBytes[filePath] -= last.Value.Bytes;
                    list.RemoveLast();
                }
            }
        }

        // ===== 清理 =====
        /// <summary>清空指定文件的全部缓存（内存 + 磁盘）</summary>
        public static void ClearFile(string filePath)
        {
            // 内存
            lock (_memLock)
            {
                _memCache.Remove(filePath);
                _memBytes.Remove(filePath);
            }
            // 磁盘
            _ = Task.Run(() =>
            {
                try
                {
                    var dir = FileCacheDir(filePath);
                    if (Directory.Exists(dir)) Directory.Delete(dir, true);
                }
                catch { }
            });
        }

        /// <summary>关闭页签时调用：只清内存，磁盘保留</summary>
        public static void UnloadFile(string filePath)
        {
            lock (_memLock)
            {
                _memCache.Remove(filePath);
                _memBytes.Remove(filePath);
            }
        }

        /// <summary>获取指定文件的内存缓存占用字节数</summary>
        public static long GetMemorySize(string filePath)
        {
            lock (_memLock)
            {
                return _memBytes.TryGetValue(filePath, out var bytes) ? bytes : 0;
            }
        }
    }
}
