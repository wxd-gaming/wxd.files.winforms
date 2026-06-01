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

        /// <summary>目标文件中用于显示名称的列索引（-1 表示不使用，只显示ID）</summary>
        public int TargetDisplayColumnIndex { get; set; } = -1;

        /// <summary>目标文件中显示列的表头名称（用于提示显示）</summary>
        public string TargetDisplayColumnName { get; set; } = "";

        #region 值提取规则（用于处理如 "4002#1" -> 提取 "4002" 的场景）

        /// <summary>是否启用值提取（默认 false 表示完全匹配）</summary>
        public bool ValueExtractionEnabled { get; set; } = false;

        /// <summary>分隔符（如 "#"、"-"、"|" 等），空字符串表示按位置提取</summary>
        public string ExtractDelimiter { get; set; } = "";

        /// <summary>提取的分段索引（0-based），例如 "4002#1" 用 "#" 分隔后，取 index=0 得到 "4002"</summary>
        public int ExtractPartIndex { get; set; } = 0;

        /// <summary>多值分隔符（如 "|"），用于将整个值拆分为多组再各自提取。为空表示不启用多值模式</summary>
        public string MultiValueDelimiter { get; set; } = "";

        /// <summary>
        /// 从原始值中提取用于关联匹配的实际值（单个值，向后兼容）
        /// </summary>
        public string ExtractValue(string originalValue)
        {
            var values = ExtractValues(originalValue);
            return values.Count > 0 ? values[0] : originalValue;
        }

        /// <summary>
        /// 从原始值中提取用于关联匹配的所有实际值（支持多值切割）
        /// 流程：先按 MultiValueDelimiter 拆分为多组，每组再按 ExtractDelimiter 提取 ExtractPartIndex
        /// 例如 "100500201#1|100500202#1" 先按 "|" 拆分，每组按 "#" 取 index=0，得到 ["100500201", "100500202"]
        /// </summary>
        public List<string> ExtractValues(string originalValue)
        {
            var result = new List<string>();
            if (!ValueExtractionEnabled || string.IsNullOrEmpty(originalValue))
            {
                result.Add(originalValue);
                return result;
            }

            try
            {
                // 第一步：按多值分隔符拆分
                var groups = string.IsNullOrEmpty(MultiValueDelimiter)
                    ? new[] { originalValue }
                    : originalValue.Split(new[] { MultiValueDelimiter }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var group in groups)
                {
                    var trimmedGroup = group.Trim();
                    if (string.IsNullOrEmpty(trimmedGroup)) continue;

                    string extracted;
                    if (string.IsNullOrEmpty(ExtractDelimiter))
                    {
                        // 无分隔符模式：按字符位置截取
                        if (ExtractPartIndex >= 0 && ExtractPartIndex < trimmedGroup.Length)
                            extracted = trimmedGroup.Substring(0, ExtractPartIndex + 1);
                        else
                            extracted = trimmedGroup;
                    }
                    else
                    {
                        // 有分隔符模式：按分隔符拆分后取指定段
                        var parts = trimmedGroup.Split(new[] { ExtractDelimiter }, StringSplitOptions.None);
                        if (ExtractPartIndex >= 0 && ExtractPartIndex < parts.Length)
                            extracted = parts[ExtractPartIndex].Trim();
                        else
                            extracted = trimmedGroup;
                    }

                    if (!string.IsNullOrEmpty(extracted) && !result.Contains(extracted))
                        result.Add(extracted);
                }

                if (result.Count == 0)
                    result.Add(originalValue);
            }
            catch
            {
                result.Clear();
                result.Add(originalValue);
            }

            return result;
        }

        #endregion
    }
}
