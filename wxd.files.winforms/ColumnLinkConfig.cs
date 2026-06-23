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

    /// <summary>每个切割段的关联配置（第0段关联哪个表、第1段关联哪个表...）</summary>
    public class FieldLinkConfig
    {
        /// <summary>切割后的第几段（0-based）</summary>
        public int FieldIndex { get; set; }

        /// <summary>关联目标文件相对路径</summary>
        public string TargetFile { get; set; } = "";

        /// <summary>目标文件中被搜索的列索引</summary>
        public int TargetColumnIndex { get; set; } = -1;

        /// <summary>目标列的表头名称（用于显示）</summary>
        public string TargetColumnName { get; set; } = "";

        /// <summary>目标文件中用于显示名称的列索引（-1 表示不使用）</summary>
        public int TargetDisplayColumnIndex { get; set; } = -1;

        /// <summary>显示列的表头名称</summary>
        public string TargetDisplayColumnName { get; set; } = "";
    }

    /// <summary>列关联配置：基于列索引的跨文件关联（支持多层值提取 + 多字段关联）</summary>
    public class ColumnLinkConfig
    {
        /// <summary>源文件相对路径</summary>
        public string SourceFile { get; set; } = "";
        /// <summary>源文件中关联的列索引（index-based）</summary>
        public int SourceColumnIndex { get; set; }

        #region 值提取规则（多层切割）

        /// <summary>是否启用值提取（默认 false 表示完全匹配）</summary>
        public bool ValueExtractionEnabled { get; set; } = false;

        /// <summary>多层切割步骤列表。默认：| 切全部 → # 取第0段</summary>
        public List<ExtractStep> ExtractSteps { get; set; } = new()
        {
            new ExtractStep { Delimiter = "|", PartIndex = -1 },
            new ExtractStep { Delimiter = "#", PartIndex = 0 }
        };

        /// <summary>表模式列头（可选），用 EffectiveDisplayDelimiter 分隔，如 "道具#数量#权重"。
        /// 为空时不显示表头行。</summary>
        public string DisplayColumnHeaders { get; set; } = "";

        /// <summary>从 ExtractSteps 最后一步获取有效的显示分隔符，用于表模式切割显示</summary>
        public string EffectiveDisplayDelimiter
        {
            get
            {
                if (ExtractSteps.Count > 0)
                {
                    var lastStep = ExtractSteps[ExtractSteps.Count - 1];
                    if (!string.IsNullOrEmpty(lastStep.Delimiter))
                        return lastStep.Delimiter;
                }
                return "#";
            }
        }

        /// <summary>多字段关联配置：每个切割段可以分别关联不同的表</summary>
        public List<FieldLinkConfig> FieldLinks { get; set; } = new();

        /// <summary>
        /// 从原始值中提取用于关联匹配的实际值（单个值）
        /// </summary>
        public string ExtractValue(string originalValue)
        {
            var values = ExtractValues(originalValue);
            return values.Count > 0 ? values[0] : originalValue;
        }

        /// <summary>
        /// 从原始值中提取用于关联匹配的所有实际值（支持多层切割）
        /// 前 ExtractSteps.Count - 1 步按行拆分，最后一步按列拆分，所有列值展平返回
        /// </summary>
        public List<string> ExtractValues(string originalValue)
        {
            var rows = ExtractRows(originalValue);
            return rows.SelectMany(r => r).ToList();
        }

        /// <summary>
        /// 按行列语义展开值：
        /// 前 ExtractSteps.Count - 1 步按行拆分（每次切割将每行拆成多行），
        /// 最后一步按列拆分（每行拆成多列）。
        /// 返回 List&lt;List&lt;string&gt;&gt;，每行是一个 List&lt;string&gt;（各列值）。
        /// </summary>
        public List<List<string>> ExtractRows(string originalValue)
        {
            var result = new List<List<string>>();
            if (!ValueExtractionEnabled || string.IsNullOrEmpty(originalValue))
            {
                result.Add(new List<string> { originalValue ?? "" });
                return result;
            }

            var steps = ExtractSteps;
            if (steps.Count == 0)
            {
                result.Add(new List<string> { originalValue });
                return result;
            }

            try
            {
                // 前 N-1 步：按行展开
                var rows = new List<string> { originalValue };
                for (int i = 0; i < steps.Count - 1; i++)
                {
                    var step = steps[i];
                    var nextRows = new List<string>();
                    foreach (var row in rows)
                    {
                        if (string.IsNullOrEmpty(step.Delimiter))
                        {
                            nextRows.Add(row);
                            continue;
                        }
                        var parts = row.Split(new[] { step.Delimiter }, StringSplitOptions.None);
                        if (step.PartIndex >= 0)
                        {
                            if (step.PartIndex < parts.Length)
                                nextRows.Add(parts[step.PartIndex].Trim());
                        }
                        else
                        {
                            // PartIndex == -1：保留全部片段
                            foreach (var p in parts)
                            {
                                var trimmed = p.Trim();
                                if (!string.IsNullOrEmpty(trimmed))
                                    nextRows.Add(trimmed);
                            }
                        }
                    }
                    rows = nextRows;
                    if (rows.Count == 0) break;
                }

                // 最后一步：按列展开
                var lastStep = steps[steps.Count - 1];
                foreach (var row in rows)
                {
                    var cols = new List<string>();
                    if (string.IsNullOrEmpty(lastStep.Delimiter))
                    {
                        cols.Add(row);
                    }
                    else
                    {
                        var parts = row.Split(new[] { lastStep.Delimiter }, StringSplitOptions.None);
                        if (lastStep.PartIndex >= 0)
                        {
                            if (lastStep.PartIndex < parts.Length)
                                cols.Add(parts[lastStep.PartIndex].Trim());
                        }
                        else
                        {
                            foreach (var p in parts)
                                cols.Add(p.Trim());
                        }
                    }
                    result.Add(cols);
                }

                if (result.Count == 0)
                    result.Add(new List<string> { originalValue });
            }
            catch
            {
                result.Clear();
                result.Add(new List<string> { originalValue });
            }

            return result;
        }

        #endregion
    }
}
