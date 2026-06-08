using System;
using System.Collections.Generic;

namespace wxd.files.winforms
{
    /// <summary>版本控制的一条提交记录</summary>
    public class VersionHistoryEntry
    {
        /// <summary>SVN: revision号(如"12345"), Git: 短SHA(如"a1b2c3d")</summary>
        public string Revision { get; set; } = "";

        /// <summary>完整SHA/版本标识(用于命令行导出)</summary>
        public string FullRevision { get; set; } = "";

        /// <summary>提交时间</summary>
        public DateTime CommitDate { get; set; }

        /// <summary>提交人</summary>
        public string Author { get; set; } = "";

        /// <summary>提交信息</summary>
        public string Message { get; set; } = "";

        /// <summary>该版本下当前文件是否发生了变更</summary>
        public bool HasFileChange { get; set; } = true;

        /// <summary>在列表中的序号(0=最新)</summary>
        public int Index { get; set; }
    }

    /// <summary>版本控制操作的结果（带错误信息）</summary>
    public class VersionHistoryResult
    {
        /// <summary>历史记录列表（出错时为空）</summary>
        public List<VersionHistoryEntry> Entries { get; set; } = new();

        /// <summary>错误信息，为 null 表示成功</summary>
        public string? ErrorMessage { get; set; }

        /// <summary>是否成功</summary>
        public bool Success => ErrorMessage == null;
    }

    /// <summary>版本导出操作的结果（带错误信息）</summary>
    public class ExportRevisionResult
    {
        /// <summary>导出的临时文件路径（失败时为 null）</summary>
        public string? TempFile { get; set; }

        /// <summary>错误信息，为 null 表示成功</summary>
        public string? ErrorMessage { get; set; }

        /// <summary>是否成功</summary>
        public bool Success => ErrorMessage == null;
    }
}
