using System;
using System.Configuration;

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

        #region 各窗格CSV/Excel固定行记录（独立存储）

        /// <summary>获取指定窗格的CSV固定行数</summary>
        public static int GetPaneCsvFrozenRows(int paneIndex)
        {
            var key = $"Pane{paneIndex}CsvFrozenRows";
            return GetIntValue(key, DefaultValue, 0, 8);
        }

        /// <summary>设置指定窗格的CSV固定行数</summary>
        public static void SetPaneCsvFrozenRows(int paneIndex, int value)
        {
            var key = $"Pane{paneIndex}CsvFrozenRows";
            SetIntValue(key, Math.Max(0, Math.Min(8, value)));
        }

        /// <summary>获取指定窗格的Excel固定行数</summary>
        public static int GetPaneExcelFrozenRows(int paneIndex)
        {
            var key = $"Pane{paneIndex}ExcelFrozenRows";
            return GetIntValue(key, DefaultValue, 0, 8);
        }

        /// <summary>设置指定窗格的Excel固定行数</summary>
        public static void SetPaneExcelFrozenRows(int paneIndex, int value)
        {
            var key = $"Pane{paneIndex}ExcelFrozenRows";
            SetIntValue(key, Math.Max(0, Math.Min(8, value)));
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
