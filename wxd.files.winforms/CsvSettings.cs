using System;
using System.Configuration;
using System.IO;

namespace wxd.files.winforms
{
    /// <summary>
    /// 全局设置管理类 - 使用ConfigurationManager统一存储
    /// </summary>
    public static class CsvSettings
    {
        private const int DefaultValue = 2;

        #region CSV查看器设置

        private const string DefaultFrozenRowsKey = "DefaultFrozenRows";

        /// <summary>
        /// CSV查看器的默认冻结行数
        /// </summary>
        public static int DefaultFrozenRows
        {
            get => GetIntValue(DefaultFrozenRowsKey, DefaultValue, 0, 8);
            set => SetIntValue(DefaultFrozenRowsKey, Math.Max(0, Math.Min(8, value)));
        }

        #endregion

        #region Excel查看器设置

        private const string ExcelDefaultFrozenRowsKey = "ExcelDefaultFrozenRows";

        /// <summary>
        /// Excel查看器的默认冻结行数
        /// </summary>
        public static int ExcelDefaultFrozenRows
        {
            get => GetIntValue(ExcelDefaultFrozenRowsKey, DefaultValue, 0, 8);
            set => SetIntValue(ExcelDefaultFrozenRowsKey, Math.Max(0, Math.Min(8, value)));
        }

        #endregion

        #region 主题设置

        private const string ThemeModeKey = "ThemeMode";

        /// <summary>
        /// 主题模式: 0=深色, 1=浅色
        /// </summary>
        public static int ThemeMode
        {
            get => GetIntValue(ThemeModeKey, 1, 0, 1);
            set => SetIntValue(ThemeModeKey, Math.Max(0, Math.Min(1, value)));
        }

        #endregion

        #region 窗格设置

        private const string PaneCountKey = "PaneCount";
        private const string LayoutIndexKey = "LayoutIndex";

        /// <summary>
        /// 窗格数量
        /// </summary>
        public static int PaneCount
        {
            get => GetIntValue(PaneCountKey, 7, 1, 10);
            set => SetIntValue(PaneCountKey, Math.Max(1, Math.Min(10, value)));
        }

        /// <summary>
        /// 布局索引: 0=横向, 1=纵向, 2=流式
        /// </summary>
        public static int LayoutIndex
        {
            get => GetIntValue(LayoutIndexKey, 2, 0, 2);
            set => SetIntValue(LayoutIndexKey, Math.Max(0, Math.Min(2, value)));
        }

        #endregion

        #region 各窗格目录和过滤条件

        private const string Pane1DirectoryKey = "Pane1Directory";
        private const string Pane2DirectoryKey = "Pane2Directory";
        private const string Pane3DirectoryKey = "Pane3Directory";
        private const string Pane4DirectoryKey = "Pane4Directory";
        private const string Pane5DirectoryKey = "Pane5Directory";
        private const string Pane6DirectoryKey = "Pane6Directory";
        private const string Pane7DirectoryKey = "Pane7Directory";
        private const string Pane8DirectoryKey = "Pane8Directory";
        private const string Pane9DirectoryKey = "Pane9Directory";
        private const string Pane10DirectoryKey = "Pane10Directory";

        private const string Pane1FilterKey = "Pane1Filter";
        private const string Pane2FilterKey = "Pane2Filter";
        private const string Pane3FilterKey = "Pane3Filter";
        private const string Pane4FilterKey = "Pane4Filter";
        private const string Pane5FilterKey = "Pane5Filter";
        private const string Pane6FilterKey = "Pane6Filter";
        private const string Pane7FilterKey = "Pane7Filter";
        private const string Pane8FilterKey = "Pane8Filter";
        private const string Pane9FilterKey = "Pane9Filter";
        private const string Pane10FilterKey = "Pane10Filter";

        /// <summary>获取窗格目录</summary>
        public static string GetPaneDirectory(int paneIndex)
        {
            var key = $"Pane{paneIndex}Directory";
            return GetStringValue(key, "");
        }

        /// <summary>设置窗格目录</summary>
        public static void SetPaneDirectory(int paneIndex, string directory)
        {
            var key = $"Pane{paneIndex}Directory";
            SetStringValue(key, directory);
        }

        /// <summary>获取窗格过滤条件</summary>
        public static string GetPaneFilter(int paneIndex)
        {
            var key = $"Pane{paneIndex}Filter";
            return GetStringValue(key, "");
        }

        /// <summary>设置窗格过滤条件</summary>
        public static void SetPaneFilter(int paneIndex, string filter)
        {
            var key = $"Pane{paneIndex}Filter";
            SetStringValue(key, filter);
        }

        /// <summary>获取窗格视图模式</summary>
        public static int GetPaneViewMode(int paneIndex)
        {
            var key = $"Pane{paneIndex}ViewMode";
            return GetIntValue(key, 0, 0, 4);
        }

        /// <summary>设置窗格视图模式</summary>
        public static void SetPaneViewMode(int paneIndex, int mode)
        {
            var key = $"Pane{paneIndex}ViewMode";
            SetIntValue(key, Math.Max(0, Math.Min(4, mode)));
        }

        #endregion

        #region 显示隐藏文件夹设置

        private const string ShowHiddenFilesKey = "ShowHiddenFiles";

        /// <summary>
        /// 是否显示隐藏文件夹和文件，默认 false
        /// </summary>
        public static bool ShowHiddenFiles
        {
            get => GetStringValue(ShowHiddenFilesKey, "false") == "true";
            set => SetStringValue(ShowHiddenFilesKey, value ? "true" : "false");
        }

        #endregion

        #region 文件夹感知数据（.wxd-work/config.json 文件存储，可随文件夹移植）

        private const string WxdWorkDirName = ".wxd-work";
        private const string WxdConfigFileName = "config.json";
        private const int MaxDirConfigs = 50;

        /// <summary>.wxd-work 配置文件夹路径</summary>
        private static string GetWxdWorkDir(string directoryPath)
        {
            return Path.Combine(directoryPath, WxdWorkDirName);
        }

        /// <summary>.wxd-work/config.json 文件路径</summary>
        private static string GetWxdConfigFilePath(string directoryPath)
        {
            return Path.Combine(GetWxdWorkDir(directoryPath), WxdConfigFileName);
        }

        /// <summary>确保 .wxd-work 隐藏文件夹存在</summary>
        private static void EnsureWxdWorkDir(string directoryPath)
        {
            var dir = GetWxdWorkDir(directoryPath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
                // 设为隐藏文件夹
                File.SetAttributes(dir, File.GetAttributes(dir) | FileAttributes.Hidden);
            }
        }

        /// <summary>从 .wxd-work/config.json 读取配置</summary>
        private static DirectoryFrozenConfig? ReadDirConfigFromFile(string directoryPath)
        {
            var configPath = GetWxdConfigFilePath(directoryPath);
            if (!File.Exists(configPath)) return null;
            try
            {
                var json = File.ReadAllText(configPath);
                var cfg = System.Text.Json.JsonSerializer.Deserialize<DirectoryFrozenConfig>(json);
                if (cfg == null) return null;
                cfg.DirectoryPath = directoryPath;
                cfg.LastAccessed = DateTime.UtcNow;
                return cfg;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>将配置写入 .wxd-work/config.json（序列化前清除运行时字段，保持文件干净）</summary>
        private static void WriteDirConfigToFile(string directoryPath, DirectoryFrozenConfig cfg)
        {
            try
            {
                EnsureWxdWorkDir(directoryPath);
                var configPath = GetWxdConfigFilePath(directoryPath);
                // 复制一份用于序列化，清除运行时字段，避免绝对路径写入文件
                var fileCfg = new DirectoryFrozenConfig
                {
                    CsvFrozenRows = cfg.CsvFrozenRows,
                    ExcelFrozenRows = cfg.ExcelFrozenRows,
                    CsvHeaderRowIndex = cfg.CsvHeaderRowIndex,
                    ExcelHeaderRowIndex = cfg.ExcelHeaderRowIndex,
                    ColumnLinks = cfg.ColumnLinks,
                    FileRemarks = cfg.FileRemarks
                };
                var json = System.Text.Json.JsonSerializer.Serialize(fileCfg, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(configPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CsvSettings.WriteDirConfigToFile Error: {ex.Message}");
                System.Diagnostics.Trace.WriteLine($"CsvSettings.WriteDirConfigToFile Error: {ex.Message}");
            }
        }

        /// <summary>删除 .wxd-work/config.json（回退到默认值时调用）</summary>
        private static void DeleteDirConfigFile(string directoryPath)
        {
            try
            {
                var configPath = GetWxdConfigFilePath(directoryPath);
                if (File.Exists(configPath))
                    File.Delete(configPath);
            }
            catch { }
        }

        // ---- 以下方法向后兼容 ConfigurationManager（旧数据），但优先使用文件 ----

        /// <summary>读取所有目录配置列表（仅从 ConfigurationManager，作为回退）</summary>
        private static List<DirectoryFrozenConfig> GetDirConfigsFromConfigManager()
        {
            var json = GetStringValue("DirFrozenConfigs", "[]");
            try
            {
                return System.Text.Json.JsonSerializer.Deserialize<List<DirectoryFrozenConfig>>(json)
                       ?? new List<DirectoryFrozenConfig>();
            }
            catch
            {
                return new List<DirectoryFrozenConfig>();
            }
        }

        /// <summary>保存目录配置列表到 ConfigurationManager（作为回退/缓存）</summary>
        private static void SetDirConfigsToConfigManager(List<DirectoryFrozenConfig> configs)
        {
            configs.Sort((a, b) => b.LastAccessed.CompareTo(a.LastAccessed));
            if (configs.Count > MaxDirConfigs)
                configs.RemoveRange(MaxDirConfigs, configs.Count - MaxDirConfigs);
            var json = System.Text.Json.JsonSerializer.Serialize(configs);
            SetStringValue("DirFrozenConfigs", json);
        }

        /// <summary>
        /// 获取目录配置：优先从 .wxd-work/config.json 读取，
        /// 不存在则从 ConfigurationManager 读取（旧数据兼容），
        /// 仍不存在则返回默认值。
        /// </summary>
        private static DirectoryFrozenConfig? TryGetDirConfig(string directoryPath)
        {
            if (string.IsNullOrEmpty(directoryPath)) return null;

            // 1. 优先从文件读取
            var fileCfg = ReadDirConfigFromFile(directoryPath);
            if (fileCfg != null) return fileCfg;

            // 2. 回退到 ConfigurationManager（旧数据兼容）
            var configs = GetDirConfigsFromConfigManager();
            var cfg = configs.Find(c => c.DirectoryPath.Equals(directoryPath, StringComparison.OrdinalIgnoreCase));
            if (cfg != null)
            {
                cfg.LastAccessed = DateTime.UtcNow;
                // 同步写到文件，下次直接读文件
                WriteDirConfigToFile(directoryPath, cfg);
                SetDirConfigsToConfigManager(configs);
                return cfg;
            }

            return null;
        }

        /// <summary>
        /// 获取或创建目录配置：优先从文件读取，不存在则创建默认值并写回文件。
        /// </summary>
        private static DirectoryFrozenConfig GetOrCreateDirConfig(string directoryPath)
        {
            var cfg = TryGetDirConfig(directoryPath);
            if (cfg != null) return cfg;

            // 3. 创建默认配置
            cfg = new DirectoryFrozenConfig
            {
                DirectoryPath = directoryPath,
                CsvFrozenRows = DefaultValue,
                ExcelFrozenRows = DefaultValue,
                CsvHeaderRowIndex = -1,
                ExcelHeaderRowIndex = -1,
                ColumnLinks = new(),
                FileRemarks = new(),
                LastModified = DateTime.UtcNow,
                LastAccessed = DateTime.UtcNow
            };
            // 写入文件
            WriteDirConfigToFile(directoryPath, cfg);
            // 也写入 ConfigurationManager 缓存
            var configs = GetDirConfigsFromConfigManager();
            configs.Add(cfg);
            SetDirConfigsToConfigManager(configs);
            return cfg;
        }

        /// <summary>
        /// 保存目录配置：写入 .wxd-work/config.json 并同步到 ConfigurationManager 缓存。
        /// </summary>
        private static void SaveDirConfig(string directoryPath, Action<DirectoryFrozenConfig> updateAction)
        {
            if (string.IsNullOrEmpty(directoryPath)) return;
            var cfg = GetOrCreateDirConfig(directoryPath);
            updateAction(cfg);
            cfg.LastModified = DateTime.UtcNow;
            cfg.LastAccessed = DateTime.UtcNow;
            // 写入文件
            WriteDirConfigToFile(directoryPath, cfg);
            // 同步到 ConfigurationManager 缓存
            var configs = GetDirConfigsFromConfigManager();
            var idx = configs.FindIndex(c => c.DirectoryPath.Equals(directoryPath, StringComparison.OrdinalIgnoreCase));
            if (idx >= 0)
                configs[idx] = cfg;
            else
                configs.Add(cfg);
            SetDirConfigsToConfigManager(configs);
        }

        /// <summary>获取指定目录的CSV冻结行数</summary>
        public static int GetCsvFrozenRowsByDirectory(string directoryPath)
        {
            var cfg = TryGetDirConfig(directoryPath);
            return cfg?.CsvFrozenRows ?? DefaultValue;
        }

        /// <summary>设置指定目录的CSV冻结行数</summary>
        public static void SetCsvFrozenRowsByDirectory(string directoryPath, int value)
        {
            SaveDirConfig(directoryPath, cfg => cfg.CsvFrozenRows = Math.Max(0, Math.Min(8, value)));
        }

        /// <summary>获取指定目录的Excel冻结行数</summary>
        public static int GetExcelFrozenRowsByDirectory(string directoryPath)
        {
            var cfg = TryGetDirConfig(directoryPath);
            return cfg?.ExcelFrozenRows ?? DefaultValue;
        }

        /// <summary>设置指定目录的Excel冻结行数</summary>
        public static void SetExcelFrozenRowsByDirectory(string directoryPath, int value)
        {
            SaveDirConfig(directoryPath, cfg => cfg.ExcelFrozenRows = Math.Max(0, Math.Min(8, value)));
        }

        /// <summary>获取指定目录的列关联列表</summary>
        public static List<ColumnLinkConfig> GetColumnLinksByDirectory(string directoryPath)
        {
            var cfg = TryGetDirConfig(directoryPath);
            return cfg?.ColumnLinks ?? new();
        }

        /// <summary>设置指定目录的列关联列表</summary>
        public static void SetColumnLinksByDirectory(string directoryPath, List<ColumnLinkConfig> links)
        {
            SaveDirConfig(directoryPath, cfg => cfg.ColumnLinks = links ?? new());
        }

        /// <summary>获取指定目录的文件备注映射表</summary>
        public static Dictionary<string, string> GetFileRemarksByDirectory(string directoryPath)
        {
            var cfg = TryGetDirConfig(directoryPath);
            return cfg?.FileRemarks ?? new();
        }

        /// <summary>设置指定目录中某个文件的备注名</summary>
        public static void SetFileRemark(string directoryPath, string relativePath, string remark)
        {
            SaveDirConfig(directoryPath, cfg =>
            {
                if (string.IsNullOrEmpty(remark))
                    cfg.FileRemarks.Remove(relativePath);
                else
                    cfg.FileRemarks[relativePath] = remark;
            });
        }

        /// <summary>获取指定目录的CSV表头行索引</summary>
        public static int GetCsvHeaderRowIndexByDirectory(string directoryPath)
        {
            var cfg = TryGetDirConfig(directoryPath);
            return cfg?.CsvHeaderRowIndex ?? -1;
        }

        /// <summary>设置指定目录的CSV表头行索引</summary>
        public static void SetCsvHeaderRowIndexByDirectory(string directoryPath, int value)
        {
            SaveDirConfig(directoryPath, cfg => cfg.CsvHeaderRowIndex = value);
        }

        /// <summary>获取指定目录的Excel表头行索引</summary>
        public static int GetExcelHeaderRowIndexByDirectory(string directoryPath)
        {
            var cfg = TryGetDirConfig(directoryPath);
            return cfg?.ExcelHeaderRowIndex ?? -1;
        }

        /// <summary>设置指定目录的Excel表头行索引</summary>
        public static void SetExcelHeaderRowIndexByDirectory(string directoryPath, int value)
        {
            SaveDirConfig(directoryPath, cfg => cfg.ExcelHeaderRowIndex = value);
        }

        #endregion

        #region 基础配置方法

        private static int GetIntValue(string key, int defaultValue, int minValue, int maxValue)
        {
            try
            {
                var value = ConfigurationManager.AppSettings[key];
                if (int.TryParse(value, out int result))
                {
                    return Math.Max(minValue, Math.Min(maxValue, result));
                }
            }
            catch { }
            return defaultValue;
        }

        private static void SetIntValue(string key, int value)
        {
            try
            {
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var appSettings = config.AppSettings;

                if (appSettings.Settings[key] == null)
                    appSettings.Settings.Add(key, value.ToString());
                else
                    appSettings.Settings[key].Value = value.ToString();

                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch { }
        }

        private static string GetStringValue(string key, string defaultValue)
        {
            try
            {
                var value = ConfigurationManager.AppSettings[key];
                return value ?? defaultValue;
            }
            catch { }
            return defaultValue;
        }

        private static void SetStringValue(string key, string value)
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
                System.Diagnostics.Debug.WriteLine($"CsvSettings.SetStringValue Error: {ex.Message}");
                System.Diagnostics.Trace.WriteLine($"CsvSettings.SetStringValue Error: {ex.Message}");
            }
        }

        #endregion

        #region 阅读器记忆目录

        private const string CsvReaderDirKey = "CsvReaderLastDirectory";
        private const string LargeFileViewerDirKey = "LargeFileViewerLastDirectory";

        /// <summary>
        /// 阅读器（CSV/Excel）上次选择的目录
        /// </summary>
        public static string CsvReaderLastDirectory
        {
            get => GetStringValue(CsvReaderDirKey, "");
            set => SetStringValue(CsvReaderDirKey, value);
        }

        /// <summary>
        /// 大文件阅读器上次选择的目录
        /// </summary>
        public static string LargeFileViewerLastDirectory
        {
            get => GetStringValue(LargeFileViewerDirKey, "");
            set => SetStringValue(LargeFileViewerDirKey, value);
        }

        #endregion

        #region 自定义命令列表

        private const string CustomCommandsKey = "CustomCommands";

        /// <summary>获取所有自定义命令列表</summary>
        public static List<CustomCommandItem> GetCustomCommands()
        {
            var json = GetStringValue(CustomCommandsKey, "[]");
            try
            {
                return System.Text.Json.JsonSerializer.Deserialize<List<CustomCommandItem>>(json)
                       ?? new List<CustomCommandItem>();
            }
            catch
            {
                return new List<CustomCommandItem>();
            }
        }

        /// <summary>保存自定义命令列表</summary>
        public static void SetCustomCommands(List<CustomCommandItem> commands)
        {
            var json = System.Text.Json.JsonSerializer.Serialize(commands);
            SetStringValue(CustomCommandsKey, json);
        }

        #endregion

        #region 最近访问路径历史

        private const string RecentPathsKey = "RecentPaths";

        /// <summary>获取最近访问的路径列表（最多15条）</summary>
        public static List<string> GetRecentPaths()
        {
            try
            {
                var json = GetStringValue(RecentPathsKey, "[]");
                var list = System.Text.Json.JsonSerializer.Deserialize<List<string>>(json);
                return list ?? new List<string>();
            }
            catch
            {
                return new List<string>();
            }
        }

        /// <summary>添加一条路径到最近访问记录，保留最近15条</summary>
        public static void AddRecentPath(string path)
        {
            try
            {
                var list = GetRecentPaths();
                // 如果已存在则移除旧记录
                list.RemoveAll(p => p.Equals(path, StringComparison.OrdinalIgnoreCase));
                // 插入到最前面
                list.Insert(0, path);
                // 只保留15条
                if (list.Count > 15)
                    list.RemoveRange(15, list.Count - 15);
                var json = System.Text.Json.JsonSerializer.Serialize(list);
                SetStringValue(RecentPathsKey, json);
            }
            catch { }
        }

        #endregion

        /// <summary>
        /// 保存设置（兼容性方法）
        /// </summary>
        public static void Save()
        {
            // 设置已经在属性setter中保存，这里是为了兼容之前的调用
        }
    }
}
