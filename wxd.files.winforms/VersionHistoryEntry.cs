using System;

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
}
