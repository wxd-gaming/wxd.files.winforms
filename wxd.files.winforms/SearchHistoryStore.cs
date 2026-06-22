using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace wxd.files.winforms
{
    /// <summary>搜索历史条目：包含关键字和配置状态</summary>
    public class SearchHistoryItem
    {
        public string Keyword { get; set; } = "";
        public bool CaseSensitive { get; set; }
        public bool UseRegex { get; set; }

        public override string ToString() => Keyword; // 下拉框显示文本
    }

    /// <summary>
    /// 全局搜索关键词历史管理（单例，跨文件通用）
    /// 持久化到 %Temp%\wxd_files_winforms\search_history.json
    /// 上限 50 条，最新在前
    /// </summary>
    public static class SearchHistoryStore
    {
        private const int MaxHistory = 50;
        private static readonly object _lock = new();
        private static List<SearchHistoryItem> _history = new();
        private static string? _savePath;
        private static bool _loaded;

        /// <summary>缓存文件路径：%Temp%\wxd_files_winforms\search_history.json</summary>
        private static string SavePath
        {
            get
            {
                if (string.IsNullOrEmpty(_savePath))
                {
                    var dir = Path.Combine(Path.GetTempPath(), "wxd_files_winforms");
                    try { Directory.CreateDirectory(dir); } catch { }
                    _savePath = Path.Combine(dir, "search_history.json");
                }
                return _savePath;
            }
        }

        /// <summary>获取历史列表（最新在前，返回副本）</summary>
        public static List<SearchHistoryItem> GetHistory()
        {
            EnsureLoaded();
            lock (_lock) return new List<SearchHistoryItem>(_history);
        }

        /// <summary>添加一条搜索记录（去重，移到最前），写入磁盘在后台执行</summary>
        public static void Add(string keyword, bool caseSensitive, bool useRegex)
        {
            if (string.IsNullOrWhiteSpace(keyword)) return;
            EnsureLoaded();
            lock (_lock)
            {
                // 去重：移除同关键字+同配置的旧条目
                _history.RemoveAll(h => h.Keyword == keyword && h.CaseSensitive == caseSensitive && h.UseRegex == useRegex);
                _history.Insert(0, new SearchHistoryItem { Keyword = keyword, CaseSensitive = caseSensitive, UseRegex = useRegex });
                if (_history.Count > MaxHistory)
                    _history.RemoveRange(MaxHistory, _history.Count - MaxHistory);
            }
            // 后台写盘，不阻塞 UI
            _ = Task.Run(Save);
        }

        /// <summary>清空全部历史</summary>
        public static void Clear()
        {
            lock (_lock) _history.Clear();
            _ = Task.Run(() =>
            {
                try { if (File.Exists(SavePath)) File.Delete(SavePath); } catch { }
            });
        }

        /// <summary>首次访问时从磁盘加载</summary>
        private static void EnsureLoaded()
        {
            if (_loaded) return;
            lock (_lock)
            {
                if (_loaded) return;
                try
                {
                    if (File.Exists(SavePath))
                    {
                        var json = File.ReadAllText(SavePath);
                        _history = JsonSerializer.Deserialize<List<SearchHistoryItem>>(json) ?? new List<SearchHistoryItem>();
                    }
                }
                catch { _history = new List<SearchHistoryItem>(); }
                _loaded = true;
            }
        }

        /// <summary>持久化到磁盘</summary>
        private static void Save()
        {
            try
            {
                List<SearchHistoryItem> snap;
                lock (_lock) snap = new List<SearchHistoryItem>(_history);
                var json = JsonSerializer.Serialize(snap);
                File.WriteAllText(SavePath, json);
            }
            catch { /* 忽略写入失败，不重要数据 */ }
        }
    }
}
