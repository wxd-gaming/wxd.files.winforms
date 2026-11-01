using System.Collections.Generic;

namespace wxd.files.winforms
{
    /// <summary>差异行类型（版本对比场景使用）</summary>
    public enum DiffRowType
    {
        None,       // 无差异
        Modified,   // 内容修改（两表都存在但内容不同）
        Added,      // 新增行（只在右表存在）
        Deleted     // 删除行（只在左表存在）
    }

    /// <summary>文件类型枚举</summary>
    public enum TableFileType
    {
        Csv,
        Excel
    }

    /// <summary>数据行（由单元格组成）</summary>
    public class TableRow
    {
        public List<string> Cells { get; set; } = new();

        /// <summary>原始行号（0-based），排序/筛选/列显隐后不变</summary>
        public int OriginalRowIndex { get; set; }

        /// <summary>
        /// 源文件物理起始行号（1-based）。CSV 引号内换行会让一条记录占多行，
        /// 所以它和 OriginalRowIndex+1 并不总是一致。-1 表示来源不是文件（如外部注入的 DataTable）。
        /// </summary>
        public int PhysicalStartLine { get; set; } = -1;

        /// <summary>源文件物理结束行号（1-based），与 PhysicalStartLine 成对；-1 表示未知</summary>
        public int PhysicalEndLine { get; set; } = -1;

        /// <summary>是否为差异行（版本对比场景使用）</summary>
        public bool IsDiffRow { get; set; }

        /// <summary>差异列索引集合（版本对比场景使用，null 表示无差异）</summary>
        public HashSet<int>? DiffColumns { get; set; }

        /// <summary>差异行类型（新增/删除/修改，版本对比场景使用）</summary>
        public DiffRowType DiffType { get; set; } = DiffRowType.None;

        /// <summary>是否为空白占位行（版本对比时，对面有数据而本侧无对应行时插入）</summary>
        public bool IsBlankRow { get; set; } = false;

        /// <summary>
        /// 行号的唯一取值口径：优先源文件物理行号；来源不是文件解析时退回记录序号+1；
        /// 空白占位行（版本对比里本侧不存在的行）返回 null，由调用方画 "—"。
        /// </summary>
        public int? DisplayLineNumber
        {
            get
            {
                if (IsBlankRow) return null;
                if (PhysicalStartLine > 0) return PhysicalStartLine;
                return OriginalRowIndex >= 0 ? OriginalRowIndex + 1 : null;
            }
        }
    }

    /// <summary>通用表格数据模型（CSV/Excel统一）</summary>
    public class FileTableData
    {
        /// <summary>文件完整路径</summary>
        public string FilePath { get; set; } = "";

        /// <summary>文件名</summary>
        public string FileName => System.IO.Path.GetFileName(FilePath);

        /// <summary>文件类型</summary>
        public TableFileType Type { get; set; }

        /// <summary>列名列表（CSV来自第一行，Excel为A/B/C...）</summary>
        public List<string> ColumnNames { get; set; } = new();

        /// <summary>数据行列表</summary>
        public List<TableRow> Rows { get; set; } = new();

        /// <summary>CSV分隔符（仅CSV有效）</summary>
        public char Delimiter { get; set; } = ',';
    }
}
