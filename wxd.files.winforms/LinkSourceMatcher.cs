using System;

namespace wxd.files.winforms
{
    /// <summary>
    /// 判断关联配置的 SourceFile 是否指向当前文件。
    /// SourceFile 既可能是相对路径（In/a.csv）也可能只是文件名（a.csv），
    /// 不能用 EndsWith 做后缀匹配，否则 ba.csv 会被 a.csv 的配置误命中。
    /// </summary>
    internal static class LinkSourceMatcher
    {
        /// <param name="sourceFile">配置中的源文件（相对路径或文件名）</param>
        /// <param name="currentRel">当前文件相对根目录的路径</param>
        /// <param name="fileName">当前文件名（currentRel 可能是纯文件名）</param>
        public static bool Matches(string? sourceFile, string currentRel, string fileName)
        {
            if (string.IsNullOrEmpty(sourceFile)) return false;

            var norm = Normalize(sourceFile);
            if (norm.Length == 0) return false;

            // 配置带目录：必须相对路径整段相等
            if (norm.IndexOf('/') >= 0)
                return string.Equals(norm, Normalize(currentRel), StringComparison.OrdinalIgnoreCase);

            // 配置只有文件名：与当前文件名精确比较（不同目录的同名文件仍视为同一张表）
            return string.Equals(norm, fileName, StringComparison.OrdinalIgnoreCase);
        }

        private static string Normalize(string? path)
        {
            if (string.IsNullOrEmpty(path)) return "";
            var p = path!.Replace('\\', '/').Trim('/');
            while (p.StartsWith("./", StringComparison.Ordinal))
                p = p.Substring(2);
            return p;
        }
    }
}
