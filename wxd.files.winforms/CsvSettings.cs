using System;
using System.Configuration;

namespace wxd.files.winforms
{
    /// <summary>
    /// CSV/Excel查看器设置管理类
    /// </summary>
    public static class CsvSettings
    {
        private const string DefaultFrozenRowsKey = "DefaultFrozenRows";
        private const string ExcelDefaultFrozenRowsKey = "ExcelDefaultFrozenRows";
        private const int DefaultFrozenRowsValue = 2; // 默认冻结2行

        /// <summary>
        /// CSV查看器的默认冻结行数
        /// </summary>
        public static int DefaultFrozenRows
        {
            get
            {
                try
                {
                    var value = ConfigurationManager.AppSettings[DefaultFrozenRowsKey];
                    if (int.TryParse(value, out int result))
                    {
                        return Math.Max(0, Math.Min(5, result)); // 限制在0-5之间
                    }
                }
                catch
                {
                    // 读取失败时使用默认值
                }
                return DefaultFrozenRowsValue;
            }
            set
            {
                try
                {
                    var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    var appSettings = config.AppSettings;
                    
                    if (appSettings.Settings[DefaultFrozenRowsKey] == null)
                    {
                        appSettings.Settings.Add(DefaultFrozenRowsKey, value.ToString());
                    }
                    else
                    {
                        appSettings.Settings[DefaultFrozenRowsKey].Value = value.ToString();
                    }
                    
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                }
                catch
                {
                    // 保存失败时忽略
                }
            }
        }

        /// <summary>
        /// Excel查看器的默认冻结行数
        /// </summary>
        public static int ExcelDefaultFrozenRows
        {
            get
            {
                try
                {
                    var value = ConfigurationManager.AppSettings[ExcelDefaultFrozenRowsKey];
                    if (int.TryParse(value, out int result))
                    {
                        return Math.Max(0, Math.Min(8, result)); // 限制在0-8之间
                    }
                }
                catch
                {
                    // 读取失败时使用默认值
                }
                return DefaultFrozenRowsValue;
            }
            set
            {
                try
                {
                    var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    var appSettings = config.AppSettings;

                    if (appSettings.Settings[ExcelDefaultFrozenRowsKey] == null)
                    {
                        appSettings.Settings.Add(ExcelDefaultFrozenRowsKey, value.ToString());
                    }
                    else
                    {
                        appSettings.Settings[ExcelDefaultFrozenRowsKey].Value = value.ToString();
                    }

                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                }
                catch
                {
                    // 保存失败时忽略
                }
            }
        }

        /// <summary>
        /// 保存设置
        /// </summary>
        public static void Save()
        {
            // 设置已经在属性setter中保存，这里是为了兼容之前的调用
        }
    }
}
