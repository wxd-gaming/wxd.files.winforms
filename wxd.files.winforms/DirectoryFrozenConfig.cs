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

    /// <summary>最后修改时间</summary>
    public DateTime LastModified { get; set; } = DateTime.UtcNow;

    /// <summary>最后访问时间</summary>
    public DateTime LastAccessed { get; set; } = DateTime.UtcNow;

    /// <summary>列关联配置列表（按目录存储）</summary>
    public List<ColumnLinkConfig> ColumnLinks { get; set; } = new();

    /// <summary>文件备注名映射（key: 相对路径, value: 备注名）</summary>
    public Dictionary<string, string> FileRemarks { get; set; } = new();
}
