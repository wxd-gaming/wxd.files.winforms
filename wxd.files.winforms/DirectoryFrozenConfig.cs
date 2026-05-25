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

    /// <summary>最后修改时间</summary>
    public DateTime LastModified { get; set; } = DateTime.UtcNow;

    /// <summary>最后访问时间</summary>
    public DateTime LastAccessed { get; set; } = DateTime.UtcNow;

    /// <summary>列关联配置列表（按目录存储）</summary>
    public List<ColumnLinkConfig> ColumnLinks { get; set; } = new();
}
