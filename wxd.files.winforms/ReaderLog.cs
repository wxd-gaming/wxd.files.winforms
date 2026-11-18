using System;
using System.IO;

namespace wxd.files.winforms
{
    /// <summary>
    /// 阅读器诊断日志：写到 exe 同目录的 .wxd-work/logs/reader.log。
    /// 用于排查"打开某个目录必崩"这类只在用户机器上出现的问题；
    /// 只记步骤和异常，不记文件内容，超过 2MB 自动滚动。
    /// </summary>
    internal static class ReaderLog
    {
        private static readonly object Gate = new();

        internal static void Write(string message)
        {
            try
            {
                lock (Gate)
                {
                    var dir = Path.Combine(AppContext.BaseDirectory, ".wxd-work", "logs");
                    Directory.CreateDirectory(dir);
                    var file = Path.Combine(dir, "reader.log");

                    var info = new FileInfo(file);
                    if (info.Exists && info.Length > 2 * 1024 * 1024)
                        File.Delete(file);

                    File.AppendAllText(file,
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}");
                }
            }
            catch { }
        }

        internal static void Exception(string context, Exception ex)
            => Write($"{context} 异常：{ex}");
    }
}
