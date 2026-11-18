using System;
using System.IO;

namespace wxd.files.winforms;

/// <summary>文件操作诊断日志：排查"操作没反应"这类没有弹窗的问题。</summary>
internal static class OperationLog
{
    private static readonly object Gate = new();

    public static void Write(string message)
    {
        try
        {
            var directory = Path.Combine(AppContext.BaseDirectory, ".wxd-work", "logs");
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, "operations.log");
            lock (Gate)
            {
                // 单文件超过 2MB 就重来，避免排查日志无限增长
                try
                {
                    var info = new FileInfo(path);
                    if (info.Exists && info.Length > 2 * 1024 * 1024) info.Delete();
                }
                catch { }

                File.AppendAllText(path,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + message + Environment.NewLine);
            }
        }
        catch
        {
            // 诊断日志失败不能影响文件操作
        }
    }
}
