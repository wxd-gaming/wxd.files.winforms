using System;

namespace wxd.files.winforms
{
    /// <summary>右键菜单模式</summary>
    public enum ContextMenuMode
    {
        /// <summary>默认完整右键菜单（编辑、插入行、删除行、关联等）</summary>
        Normal,
        /// <summary>禁用右键菜单（由外部事件自行处理）</summary>
        None
    }

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

        /// <summary>true=仅显示表格（固定行+内容行），隐藏所有工具栏和状态栏</summary>
        public bool ShowMinimalUI { get; set; } = false;

        /// <summary>右键菜单模式，默认 Normal（完整菜单）</summary>
        public ContextMenuMode ContextMenuMode { get; set; } = ContextMenuMode.Normal;

        /// <summary>true=跳过构造函数中的自动加载（由外部通过 SetDataTable 注入数据）</summary>
        public bool SkipAutoLoad { get; set; } = false;

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
