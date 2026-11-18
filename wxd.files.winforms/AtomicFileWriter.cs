using System;
using System.IO;

namespace wxd.files.winforms
{
    /// <summary>
    /// 小文件原子写：先写同目录临时文件，再替换目标文件。
    /// 避免进程在写盘中途退出或两次后台写交叉时留下半截 JSON。
    /// </summary>
    internal static class AtomicFileWriter
    {
        private static readonly object _lock = new();

        public static bool TryWriteAllText(string path, string content)
            => TryWriteAllText(path, content, out _);

        /// <summary>按指定编码原子写（JSON 保存用，保留原文件编码）。</summary>
        public static bool TryWriteAllText(string path, string content, System.Text.Encoding encoding, out string error)
        {
            error = "";
            var dir = Path.GetDirectoryName(path);
            if (string.IsNullOrEmpty(dir)) dir = ".";
            var tempPath = Path.Combine(dir, $"{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
            try
            {
                lock (_lock)
                {
                    File.WriteAllText(tempPath, content, encoding);
                    File.Move(tempPath, path, overwrite: true);
                }
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                return false;
            }
        }

        public static bool TryWriteAllText(string path, string content, out string error)
        {
            error = "";
            var dir = Path.GetDirectoryName(path);
            if (string.IsNullOrEmpty(dir)) dir = ".";
            var tempPath = Path.Combine(dir, $"{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
            try
            {
                lock (_lock)
                {
                    File.WriteAllText(tempPath, content);
                    File.Move(tempPath, path, overwrite: true);
                }
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                return false;
            }
        }
    }
}
