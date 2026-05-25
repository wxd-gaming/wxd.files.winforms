using System;
using System.Collections.Generic;

namespace wxd.files.winforms
{
    /// <summary>列关联配置：基于列索引的跨文件关联</summary>
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
    }
}
