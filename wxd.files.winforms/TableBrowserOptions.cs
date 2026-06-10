using System;

namespace wxd.files.winforms
{
    /// <summary>
    /// TableBrowserControl 统一配置对象（VO）。
    /// 所有参数集中管理，新增参数只需修改这一个类，避免级联修改构造函数。
    /// </summary>
    public class TableBrowserOptions
    {
        // ========== 必填 ==========
        /// <summary>文件完整路径</summary>
        public string FilePath { get; }

        // ========== 目录相关 ==========
        /// <summary>根目录（用于计算相对路径、读取配置），null 时自动从 FilePath 推断</summary>
        public string? RootDir { get; set; }

        // ========== 显示配置（-1 = 从目录配置自动读取）==========
        public int FrozenRows { get; set; } = -1;
        public int HeaderRowIndex { get; set; } = -1;
        public int MappingRowIndex { get; set; } = -1;
        public int UniqueColumnIndex { get; set; } = -1;

        // ========== 初始状态 ==========
        public string InitialContentFilter { get; set; } = "";
        public int PaneIndex { get; set; } = -1;
        public bool Lightweight { get; set; } = false;

        public TableBrowserOptions(string filePath)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
        }
    }

    /// <summary>关联跳转打开时的附加参数</summary>
    public class LinkOpenOptions
    {
        public int HighlightColumnIndex { get; set; } = -1;
        public string HighlightValue { get; set; } = "";
        public string HighlightColumnName { get; set; } = "";
        public ColumnLinkConfig? LinkConfig { get; set; }
        public string DisplayName { get; set; } = "";
        public bool IsLinkedFromCell { get; set; } = false;
    }
}
