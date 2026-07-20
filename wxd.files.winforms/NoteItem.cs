using System;

namespace wxd.files.winforms
{
    /// <summary>笔记数据模型（含 RTF 富文本内容，支持内嵌图片）</summary>
    public class NoteItem
    {
        /// <summary>唯一标识</summary>
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        /// <summary>标题</summary>
        public string Title { get; set; } = "新建笔记";

        /// <summary>RTF 格式内容（图片以十六进制内嵌，无需单独管理图片文件）</summary>
        public string RtfContent { get; set; } = "";

        /// <summary>纯文本内容（用于搜索匹配，由 RTF 转换而来）</summary>
        public string PlainText { get; set; } = "";

        /// <summary>创建时间</summary>
        public DateTime CreatedAt { get; set; } = DateTime.Now;

        /// <summary>最后更新时间</summary>
        public DateTime UpdatedAt { get; set; } = DateTime.Now;
    }
}
