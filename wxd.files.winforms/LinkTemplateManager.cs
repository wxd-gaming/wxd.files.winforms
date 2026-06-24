using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace wxd.files.winforms
{
    /// <summary>关联配置模板管理器（全局存储，通过 ConfigurationManager 持久化）</summary>
    public static class LinkTemplateManager
    {
        /// <summary>获取所有模板（全局）</summary>
        public static List<LinkTemplate> GetTemplates()
        {
            return CsvSettings.GetLinkTemplates();
        }

        /// <summary>获取指定 Uid 的模板</summary>
        public static LinkTemplate? GetTemplate(string uid)
        {
            var templates = GetTemplates();
            return templates.FirstOrDefault(t => t.Uid == uid);
        }

        /// <summary>添加模板（全局）</summary>
        public static void AddTemplate(LinkTemplate template)
        {
            if (template == null) return;
            var templates = GetTemplates();
            templates.Add(template);
            CsvSettings.SetLinkTemplates(templates);
        }

        /// <summary>更新模板（找不到则添加）</summary>
        public static void UpdateTemplate(LinkTemplate template)
        {
            if (template == null) return;
            var templates = GetTemplates();
            var idx = templates.FindIndex(t => t.Uid == template.Uid);
            if (idx >= 0)
                templates[idx] = template;
            else
                templates.Add(template);
            CsvSettings.SetLinkTemplates(templates);
        }

        /// <summary>删除模板</summary>
        public static void DeleteTemplate(string uid)
        {
            if (string.IsNullOrEmpty(uid)) return;
            var templates = GetTemplates();
            templates.RemoveAll(t => t.Uid == uid);
            CsvSettings.SetLinkTemplates(templates);
        }

        /// <summary>获取引用某模板的所有列关联配置（跨目录扫描）</summary>
        public static List<ColumnLinkConfig> GetReferencingLinks(string templateUid)
        {
            if (string.IsNullOrEmpty(templateUid))
                return new();
            // 由于模板是全局的，但列关联配置存储在目录级，这里只返回找到的引用
            // 删除时会在 LinkTemplateDialog.OnDeleteClick 中通过 _directoryPath 传入目录路径
            // 实际引用检查由调用方传入的 directoryPath 完成
            return new();
        }

        /// <summary>解除指定目录中所有列关联配置对指定模板的引用（TemplateUid 置 null）</summary>
        public static void UnlinkAllReferences(string directoryPath, string templateUid)
        {
            if (string.IsNullOrEmpty(directoryPath) || string.IsNullOrEmpty(templateUid)) return;
            var allLinks = CsvSettings.GetColumnLinksByDirectory(directoryPath);
            bool changed = false;
            foreach (var link in allLinks)
            {
                if (link.TemplateUid == templateUid)
                {
                    link.TemplateUid = null;
                    changed = true;
                }
            }
            if (changed)
                CsvSettings.SetColumnLinksByDirectory(directoryPath, allLinks);
        }

        /// <summary>获取指定目录中引用某模板的所有列关联配置</summary>
        public static List<ColumnLinkConfig> GetReferencingLinks(string directoryPath, string templateUid)
        {
            if (string.IsNullOrEmpty(directoryPath) || string.IsNullOrEmpty(templateUid))
                return new();
            var allLinks = CsvSettings.GetColumnLinksByDirectory(directoryPath);
            return allLinks.Where(l => l.TemplateUid == templateUid).ToList();
        }

        /// <summary>获取模板的摘要文本（用于列表显示）</summary>
        public static string GetTemplateSummary(LinkTemplate template)
        {
            if (template == null) return "";
            var parts = new List<string>();
            if (template.ValueExtractionEnabled && template.ExtractSteps.Count > 0)
            {
                var descs = template.ExtractSteps.ConvertAll(s =>
                    s.PartIndex >= 0 ? $"\"{s.Delimiter}\"取{s.PartIndex}段" : $"\"{s.Delimiter}\"全");
                parts.Add(string.Join("→", descs));
            }
            int linkedFields = template.FieldLinks?.Count(f => !string.IsNullOrEmpty(f.TargetFile)) ?? 0;
            if (linkedFields > 0)
                parts.Add($"{linkedFields}字段关联");
            return parts.Count > 0 ? string.Join(" | ", parts) : "（空模板）";
        }
    }
}
