using System;
using System.Collections.Generic;

namespace wxd.files.winforms;

/// <summary>目录冻结行配置（用于LRU淘汰存储）</summary>
public class DirectoryFrozenConfig
{
    /// <summary>目录路径</summary>
    public string DirectoryPath { get; set; } = "";

    /// <summary>CSV冻结行数</summary>
    public int CsvFrozenRows { get; set; } = 2;

    /// <summary>Excel冻结行数</summary>
    public int ExcelFrozenRows { get; set; } = 2;

    /// <summary>CSV表头行索引（0-based，-1表示未设置/默认第一行为表头），必须 < CsvFrozenRows</summary>
    public int CsvHeaderRowIndex { get; set; } = -1;

    /// <summary>Excel表头行索引（0-based，-1表示未设置/默认第一行为表头），必须 < ExcelFrozenRows</summary>
    public int ExcelHeaderRowIndex { get; set; } = -1;

    /// <summary>CSV唯一列索引（0-based，-1表示未设置）</summary>
    public int CsvUniqueColumnIndex { get; set; } = -1;

    /// <summary>Excel唯一列索引（0-based，-1表示未设置）</summary>
    public int ExcelUniqueColumnIndex { get; set; } = -1;

    /// <summary>CSV映射行索引（固定行中哪一行是中文映射，0-based，-1表示未设置），必须 < CsvFrozenRows</summary>
    public int CsvMappingRowIndex { get; set; } = -1;

    /// <summary>Excel映射行索引（固定行中哪一行是中文映射，0-based，-1表示未设置），必须 < ExcelFrozenRows</summary>
    public int ExcelMappingRowIndex { get; set; } = -1;

    /// <summary>最后修改时间</summary>
    public DateTime LastModified { get; set; } = DateTime.UtcNow;

    /// <summary>最后访问时间</summary>
    public DateTime LastAccessed { get; set; } = DateTime.UtcNow;

    // ========== 文件级配置（新）==========
    /// <summary>文件级配置（key: 文件相对路径, value: 文件配置）</summary>
    public Dictionary<string, FileConfig> FileSettings { get; set; } = new();

    /// <summary>关联配置模板列表</summary>
    public List<LinkTemplate>? LinkTemplates { get; set; }

    // ========== 已废弃（仅用于兼容旧数据，迁移后自动清空）==========
    [Obsolete("已迁移到 FileSettings.ColumnLinks")]
    public List<ColumnLinkConfig>? ColumnLinks { get; set; }

    [Obsolete("已迁移到 FileSettings.Remark")]
    public Dictionary<string, string>? FileRemarks { get; set; }

    [Obsolete("已迁移到 FileSettings.HiddenColumns")]
    public Dictionary<string, string>? ColumnVisibility { get; set; }
}

/// <summary>单个文件的配置（只记录与默认值不同的项，减少 JSON 体积）</summary>
public class FileConfig
{
    /// <summary>备注名（null = 无备注）</summary>
    public string? Remark { get; set; }

    /// <summary>唯一列索引（null = 使用目录级默认）</summary>
    public int? UniqueColumnIndex { get; set; }

    /// <summary>隐藏的列索引集合（null 或空 = 全部显示）。只记录被取消勾选的列</summary>
    public HashSet<int>? HiddenColumns { get; set; }

    /// <summary>列关联配置列表（按 SourceFile 分组存储）</summary>
    public List<ColumnLinkConfig>? ColumnLinks { get; set; }

    /// <summary>是否所有字段都是默认值（可用于清理空条目）</summary>
    public bool IsEmpty =>
        Remark == null && UniqueColumnIndex == null &&
        (HiddenColumns == null || HiddenColumns.Count == 0) &&
        (ColumnLinks == null || ColumnLinks.Count == 0);
}
