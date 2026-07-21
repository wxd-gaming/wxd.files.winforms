using System;
using System.Collections.Concurrent;
using System.Configuration;
using System.Text.Json;

namespace wxd.files.winforms
{
    /// <summary>
    /// 应用程序状态存储 — 基于多级 ConcurrentDictionary 的页签状态管理。
    /// 
    /// 数据结构：
    ///   ConcurrentDictionary&lt;paneIndex, ConcurrentDictionary&lt;tabIndex, TabPageState&gt;&gt;
    ///   
    /// paneIndex: 0~9 (窗格索引)
    /// tabIndex:  0~9 (窗格内页签索引)
    /// 
    /// 序列化为 JSON 存入 ConfigurationManager，退出时一次性持久化。
    /// </summary>
    public static class AppStateStore
    {
        private const string TabStatesKey = "TabStates";
        private const string PaneMetaKey = "PaneMeta";

        private const int MaxPaneCount = 10;
        private const int MaxTabCountPerPane = 10;

        // ==================== 内存数据结构 ====================

        /// <summary>多级存储：paneIndex → tabIndex → TabPageState</summary>
        private static readonly ConcurrentDictionary<int, ConcurrentDictionary<int, TabPageState>> _tabStates = new();

        /// <summary>窗格元数据：paneIndex → (tabCount, activeTabIndex)</summary>
        private static readonly ConcurrentDictionary<int, (int TabCount, int ActiveTabIndex)> _paneMeta = new();

        // ==================== 加载 / 保存 ====================

        /// <summary>从 ConfigurationManager 加载所有页签状态到内存</summary>
        public static void Load()
        {
            _tabStates.Clear();
            _paneMeta.Clear();

            try
            {
                // 加载页签状态
                var statesJson = GetConfigString(TabStatesKey, "{}");
                var statesDict = JsonSerializer.Deserialize<ConcurrentDictionary<int, ConcurrentDictionary<int, TabPageState>>>(statesJson);
                if (statesDict != null)
                {
                    foreach (var kv in statesDict)
                        _tabStates[kv.Key] = kv.Value;
                }

                // 加载窗格元数据
                var metaJson = GetConfigString(PaneMetaKey, "{}");
                var metaDict = JsonSerializer.Deserialize<ConcurrentDictionary<int, PaneMetaEntry>>(metaJson);
                if (metaDict != null)
                {
                    foreach (var kv in metaDict)
                        _paneMeta[kv.Key] = (kv.Value.TabCount, kv.Value.ActiveTabIndex);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AppStateStore.Load Error: {ex.Message}");
            }
        }

        /// <summary>将所有页签状态序列化并保存到 ConfigurationManager</summary>
        public static void Save()
        {
            try
            {
                // 序列化页签状态
                var statesJson = JsonSerializer.Serialize(_tabStates);
                SetConfigString(TabStatesKey, statesJson);

                // 序列化窗格元数据
                var metaDict = new ConcurrentDictionary<int, PaneMetaEntry>();
                foreach (var kv in _paneMeta)
                {
                    metaDict[kv.Key] = new PaneMetaEntry
                    {
                        TabCount = kv.Value.TabCount,
                        ActiveTabIndex = kv.Value.ActiveTabIndex
                    };
                }
                var metaJson = JsonSerializer.Serialize(metaDict);
                SetConfigString(PaneMetaKey, metaJson);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AppStateStore.Save Error: {ex.Message}");
            }
        }

        // ==================== 页签状态读写 ====================

        /// <summary>获取指定窗格中指定页签的状态，如果不存在则返回默认值</summary>
        public static TabPageState GetTabState(int paneIndex, int tabIndex)
        {
            if (_tabStates.TryGetValue(paneIndex, out var paneDict) &&
                paneDict.TryGetValue(tabIndex, out var state))
            {
                return state;
            }
            return new TabPageState();
        }

        /// <summary>设置指定窗格中指定页签的状态（仅更新内存，不持久化）</summary>
        public static void SetTabState(int paneIndex, int tabIndex, TabPageState state)
        {
            var paneDict = _tabStates.GetOrAdd(paneIndex, _ => new ConcurrentDictionary<int, TabPageState>());
            paneDict[tabIndex] = state;
        }

        // ==================== 窗格元数据 ====================

        /// <summary>获取指定窗格的页签数量（默认1）</summary>
        public static int GetPaneTabCount(int paneIndex)
        {
            if (_paneMeta.TryGetValue(paneIndex, out var meta))
                return Math.Max(1, Math.Min(MaxTabCountPerPane, meta.TabCount));
            return 1;
        }

        /// <summary>设置指定窗格的页签数量</summary>
        public static void SetPaneTabCount(int paneIndex, int count)
        {
            var meta = _paneMeta.GetOrAdd(paneIndex, _ => (1, 0));
            _paneMeta[paneIndex] = (Math.Max(1, Math.Min(MaxTabCountPerPane, count)), meta.ActiveTabIndex);
        }

        /// <summary>获取指定窗格的当前激活页签索引（默认0）</summary>
        public static int GetPaneActiveTab(int paneIndex)
        {
            if (_paneMeta.TryGetValue(paneIndex, out var meta))
                return Math.Max(0, Math.Min(MaxTabCountPerPane - 1, meta.ActiveTabIndex));
            return 0;
        }

        /// <summary>设置指定窗格的当前激活页签索引</summary>
        public static void SetPaneActiveTab(int paneIndex, int tabIndex)
        {
            var meta = _paneMeta.GetOrAdd(paneIndex, _ => (1, 0));
            _paneMeta[paneIndex] = (meta.TabCount, Math.Max(0, Math.Min(MaxTabCountPerPane - 1, tabIndex)));
        }

        // ==================== ConfigurationManager 辅助 ====================

        private static string GetConfigString(string key, string defaultValue)
        {
            try
            {
                var value = ConfigurationManager.AppSettings[key];
                return value ?? defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

        private static void SetConfigString(string key, string value)
        {
            try
            {
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var appSettings = config.AppSettings;

                if (appSettings.Settings[key] == null)
                    appSettings.Settings.Add(key, value);
                else
                    appSettings.Settings[key].Value = value;

                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AppStateStore.SetConfigString Error: {ex.Message}");
            }
        }

        // ==================== JSON 序列化辅助类 ====================

        private class PaneMetaEntry
        {
            public int TabCount { get; set; } = 1;
            public int ActiveTabIndex { get; set; } = 0;
        }
    }
}
