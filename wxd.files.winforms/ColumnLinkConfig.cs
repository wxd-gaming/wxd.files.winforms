using System;
using System.Collections.Generic;

namespace wxd.files.winforms
{
    /// <summary>单层切割步骤</summary>
    public class ExtractStep
    {
        /// <summary>分隔符（如 "|"、"#"、":"、"-" 等）</summary>
        public string Delimiter { get; set; } = "";

        /// <summary>切割后取第几段（0-based），-1 表示保留全部片段不取段</summary>
        public int PartIndex { get; set; } = 0;
    }

    /// <summary>列关联配置：基于列索引的跨文件关联（支持多层值提取）</summary>
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

        #region 值提取规则（多层切割）

        /// <summary>是否启用值提取（默认 false 表示完全匹配）</summary>
        public bool ValueExtractionEnabled { get; set; } = false;

        /// <summary>多层切割步骤列表（替代旧的两层切割）。默认：| 切全部 → # 取第0段</summary>
        public List<ExtractStep> ExtractSteps { get; set; } = new()
        {
            new ExtractStep { Delimiter = "|", PartIndex = -1 },
            new ExtractStep { Delimiter = "#", PartIndex = 0 }
        };

        #region 旧字段（仅用于 JSON 反序列化兼容，新代码不要直接读写）
        /// <summary>[已废弃] 请使用 ExtractSteps</summary>
        [Obsolete("已迁移到 ExtractSteps")]
        public string ExtractDelimiter { get; set; } = "";

        /// <summary>[已废弃] 请使用 ExtractSteps</summary>
        [Obsolete("已迁移到 ExtractSteps")]
        public int ExtractPartIndex { get; set; } = 0;

        /// <summary>[已废弃] 请使用 ExtractSteps</summary>
        [Obsolete("已迁移到 ExtractSteps")]
        public string MultiValueDelimiter { get; set; } = "";

        /// <summary>从旧字段迁移到 ExtractSteps（应在反序列化后调用）</summary>
        public void MigrateFromOldFields()
        {
            // 如果已有步骤列表则跳过（说明已是新格式或已迁移）
            if (ExtractSteps.Count > 0) return;

            // 尝试从旧字段重建
            #pragma warning disable 612, 618
            bool hasOldMulti = !string.IsNullOrEmpty(MultiValueDelimiter);
            bool hasOldExtract = !string.IsNullOrEmpty(ExtractDelimiter) || ExtractPartIndex > 0;
            #pragma warning restore 612, 618

            if (hasOldMulti || hasOldExtract)
            {
                ExtractSteps.Clear();
                #pragma warning disable 612, 618
                if (hasOldMulti)
                {
                    ExtractSteps.Add(new ExtractStep { Delimiter = MultiValueDelimiter, PartIndex = -1 });
                }
                if (hasOldExtract)
                {
                    ExtractSteps.Add(new ExtractStep
                    {
                        Delimiter = string.IsNullOrEmpty(ExtractDelimiter) ? "#" : ExtractDelimiter,
                        PartIndex = ExtractPartIndex
                    });
                }
                #pragma warning restore 612, 618
            }
            else
            {
                // 无旧数据，设置默认值
                ExtractSteps = new List<ExtractStep>
                {
                    new ExtractStep { Delimiter = "|", PartIndex = -1 },
                    new ExtractStep { Delimiter = "#", PartIndex = 0 }
                };
            }
        }
        #endregion

        /// <summary>
        /// 从原始值中提取用于关联匹配的实际值（单个值，向后兼容）
        /// </summary>
        public string ExtractValue(string originalValue)
        {
            var values = ExtractValues(originalValue);
            return values.Count > 0 ? values[0] : originalValue;
        }

        /// <summary>
        /// 从原始值中提取用于关联匹配的所有实际值（支持多层切割）
        /// 遍历 ExtractSteps，每步按分隔符切割当前值列表：
        ///   PartIndex >= 0 → 切割后取指定段
        ///   PartIndex == -1 → 切割后保留全部片段
        /// 例如 "2#1|3#1" → [|:-1, #:0] → ["2","3"]
        /// 例如 "1:2#1|2:3#1" → [|:-1, :-1, #:0] → ["1","2"]
        /// </summary>
        public List<string> ExtractValues(string originalValue)
        {
            var result = new List<string>();
            if (!ValueExtractionEnabled || string.IsNullOrEmpty(originalValue))
            {
                result.Add(originalValue);
                return result;
            }

            // 确保旧数据已迁移
            MigrateFromOldFields();

            if (ExtractSteps.Count == 0)
            {
                result.Add(originalValue);
                return result;
            }

            try
            {
                // 初始值为整个字符串
                var current = new List<string> { originalValue };

                for (int stepIdx = 0; stepIdx < ExtractSteps.Count; stepIdx++)
                {
                    var step = ExtractSteps[stepIdx];
                    var next = new List<string>();

                    foreach (var part in current)
                    {
                        if (string.IsNullOrEmpty(step.Delimiter))
                        {
                            // 无分隔符：按字符位置截取（向后兼容）
                            if (step.PartIndex >= 0 && step.PartIndex < part.Length)
                                next.Add(part.Substring(0, step.PartIndex + 1));
                            else
                                next.Add(part);
                        }
                        else
                        {
                            var parts = part.Split(new[] { step.Delimiter }, StringSplitOptions.None);
                            if (step.PartIndex >= 0)
                            {
                                // 取指定段
                                if (step.PartIndex < parts.Length)
                                    next.Add(parts[step.PartIndex].Trim());
                            }
                            else
                            {
                                // PartIndex == -1：保留全部片段
                                foreach (var p in parts)
                                {
                                    var trimmed = p.Trim();
                                    if (!string.IsNullOrEmpty(trimmed))
                                        next.Add(trimmed);
                                }
                            }
                        }
                    }
                    current = next;
                    if (current.Count == 0) break;
                }

                foreach (var v in current)
                {
                    if (!string.IsNullOrEmpty(v) && !result.Contains(v))
                        result.Add(v);
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
