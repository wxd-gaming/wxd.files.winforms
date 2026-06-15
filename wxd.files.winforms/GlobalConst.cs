using System;

namespace wxd.files.winforms
{
    /// <summary>全局常量定义（换行符、分隔符等），避免代码中硬编码特殊字符</summary>
    public static class GlobalConst
    {
        /// <summary>Windows 换行符 \r\n</summary>
        public const string NewLine = "\r\n";

        /// <summary>Unix/Linux 换行符 \n</summary>
        public const string Lf = "\n";

        /// <summary>回车符 \r（旧版 Mac 使用）</summary>
        public const string Cr = "\r";

        /// <summary>显式的 CRLF 换行符 \r\n</summary>
        public const string CrLf = "\r\n";

        /// <summary>制表符 \t</summary>
        public const string Tab = "\t";

        /// <summary>空格</summary>
        public const string Space = " ";

        /// <summary>空字符串</summary>
        public const string Empty = "";
    }
}
