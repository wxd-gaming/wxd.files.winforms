using System;
using System.Collections.Generic;

namespace wxd.files.winforms
{
    /// <summary>列关联配置：基于列索引的跨文件关联（支持值提取）</summary>
    public class ColumnLinkConfig
    {
        /// <summary>源文件相对路径</summary>
        public string SourceFile { get; set; } = "";
        /// <summary>源文件中关联的列索引（index-based）</summary>
        public int SourceColumnIndex { get; set; }
        /// <summary>目标文件相对路径</summary>
        public string TargetFile { get; set; } = "";
        /// <summary>目标文件中被搜索的列索引（index-based）</summary>
        public int TargetColumnIndex { get; set; }

        /// <summary>目标文件中被搜索的列的表头名称（用于显示，可选）</summary>
        public string TargetColumnName { get; set; } = "";

        #region 值提取规则（用于处理如 "4002#1" -> 提取 "4002" 的场景）

        /// <summary>是否启用值提取（默认 false 表示完全匹配）</summary>
        public bool ValueExtractionEnabled { get; set; } = false;

        /// <summary>分隔符（如 "#"、"-"、"|" 等），空字符串表示按位置提取</summary>
        public string ExtractDelimiter { get; set; } = "";

        /// <summary>提取的分段索引（0-based），例如 "4002#1" 用 "#" 分隔后，取 index=0 得到 "4002"</summary>
        public int ExtractPartIndex { get; set; } = 0;

        /// <summary>
        /// 从原始值中提取用于关联匹配的实际值
        /// </summary>
        /// <param name="originalValue">原始单元格值</param>
        /// <returns>提取后的值（未启用提取或失败时返回原值）</returns>
        public string ExtractValue(string originalValue)
        {
            if (!ValueExtractionEnabled || string.IsNullOrEmpty(originalValue))
                return originalValue;

            try
            {
                if (string.IsNullOrEmpty(ExtractDelimiter))
                {
                    // 无分隔符模式：按字符位置截取
                    if (ExtractPartIndex >= 0 && ExtractPartIndex < originalValue.Length)
                        return originalValue.Substring(0, ExtractPartIndex + 1);
                    return originalValue;
                }
                else
                {
                    // 有分隔符模式：按分隔符拆分后取指定段
                    var parts = originalValue.Split(new[] { ExtractDelimiter }, StringSplitOptions.None);
                    if (ExtractPartIndex >= 0 && ExtractPartIndex < parts.Length)
                        return parts[ExtractPartIndex].Trim();
                    return originalValue;
                }
            }
            catch
            {
                return originalValue;
            }
        }

        #endregion
    }
}
