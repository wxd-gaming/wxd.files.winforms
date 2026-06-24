using System;
using System.Collections.Generic;

namespace wxd.files.winforms
{
    /// <summary>关联配置模板，包含完整的切割方案 + 各段关联配置 + 表头</summary>
    public class LinkTemplate
    {
        /// <summary>唯一标识符</summary>
        public string Uid { get; set; } = Guid.NewGuid().ToString("N");

        /// <summary>用户自定义名称</summary>
        public string Name { get; set; } = "";

        /// <summary>是否启用值提取</summary>
        public bool ValueExtractionEnabled { get; set; } = false;

        /// <summary>多层切割步骤</summary>
        public List<ExtractStep> ExtractSteps { get; set; } = new()
        {
            new ExtractStep { Delimiter = "|", PartIndex = -1 },
            new ExtractStep { Delimiter = "#", PartIndex = 0 }
        };

        /// <summary>表模式列头</summary>
        public string DisplayColumnHeaders { get; set; } = "";

        /// <summary>各段关联配置</summary>
        public List<FieldLinkConfig> FieldLinks { get; set; } = new();
    }

    /// <summary>将 LinkTemplate 的配置复制到 ColumnLinkConfig 上</summary>
    public static class LinkTemplateExtensions
    {
        /// <summary>将模板内容应用到 ColumnLinkConfig（深拷贝），不清除 SourceFile/SourceColumnIndex</summary>
        public static void ApplyTemplate(this ColumnLinkConfig config, LinkTemplate template)
        {
            if (template == null) return;
            config.ValueExtractionEnabled = template.ValueExtractionEnabled;
            config.ExtractSteps = template.ExtractSteps?.ConvertAll(s => new ExtractStep
            {
                Delimiter = s.Delimiter,
                PartIndex = s.PartIndex
            }) ?? new();
            config.DisplayColumnHeaders = template.DisplayColumnHeaders;
            config.FieldLinks = template.FieldLinks?.ConvertAll(f => new FieldLinkConfig
            {
                FieldIndex = f.FieldIndex,
                TargetFile = f.TargetFile,
                TargetColumnIndex = f.TargetColumnIndex,
                TargetColumnName = f.TargetColumnName,
                TargetDisplayColumnIndex = f.TargetDisplayColumnIndex,
                TargetDisplayColumnName = f.TargetDisplayColumnName,
            }) ?? new();
        }

        /// <summary>从 ColumnLinkConfig 创建 LinkTemplate（不含 SourceFile/SourceColumnIndex）</summary>
        public static LinkTemplate ToTemplate(this ColumnLinkConfig config, string name = "")
        {
            return new LinkTemplate
            {
                Uid = Guid.NewGuid().ToString("N"),
                Name = name,
                ValueExtractionEnabled = config.ValueExtractionEnabled,
                ExtractSteps = config.ExtractSteps?.ConvertAll(s => new ExtractStep
                {
                    Delimiter = s.Delimiter,
                    PartIndex = s.PartIndex
                }) ?? new(),
                DisplayColumnHeaders = config.DisplayColumnHeaders,
                FieldLinks = config.FieldLinks?.ConvertAll(f => new FieldLinkConfig
                {
                    FieldIndex = f.FieldIndex,
                    TargetFile = f.TargetFile,
                    TargetColumnIndex = f.TargetColumnIndex,
                    TargetColumnName = f.TargetColumnName,
                    TargetDisplayColumnIndex = f.TargetDisplayColumnIndex,
                    TargetDisplayColumnName = f.TargetDisplayColumnName,
                }) ?? new(),
            };
        }
    }
}
