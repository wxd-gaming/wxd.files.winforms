using System;

namespace wxd.files.winforms
{
    /// <summary>
    /// 单个页签的完整状态快照：目录、过滤条件、视图模式、排序状态。
    /// 用于 AppStateStore 的多级存储。
    /// </summary>
    public class TabPageState
    {
        /// <summary>当前目录路径</summary>
        public string Directory { get; set; } = "";

        /// <summary>过滤关键词（空字符串表示无过滤）</summary>
        public string Filter { get; set; } = "";

        /// <summary>视图模式 (int)FileListViewMode</summary>
        public int ViewMode { get; set; } = 0;

        /// <summary>排序字段 (int)SortField</summary>
        public int SortField { get; set; } = 0;

        /// <summary>排序方向 (int)SortDirection</summary>
        public int SortDirection { get; set; } = 0;
    }
}
