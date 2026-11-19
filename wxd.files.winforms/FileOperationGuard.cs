using System;
using System.IO;

namespace wxd.files.winforms;

/// <summary>文件操作前的路径关系判断，避免把目录复制/移动进它自己。</summary>
internal static class FileOperationGuard
{
    /// <summary>目标目录等于源目录，或位于源目录内部时返回 true。</summary>
    public static bool WouldRecurse(string source, string targetDirectory)
    {
        var fullSource = Path.GetFullPath(source)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var fullTarget = Path.GetFullPath(targetDirectory)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (string.Equals(fullSource, fullTarget, StringComparison.OrdinalIgnoreCase))
            return true;

        return fullTarget.StartsWith(fullSource + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
    }
}
