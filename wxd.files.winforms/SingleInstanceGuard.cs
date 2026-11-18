using System;
using System.IO;

namespace wxd.files.winforms
{
    /// <summary>
    /// 同一 exe 目录只允许一个实例：在 .wxd-work/app.lock 上开独占句柄。
    ///
    /// 这样做的目的不只是"防重复启动"：
    /// 全局设置写在 {exe}/.wxd-work/settings.json，是整文件读改写；
    /// 如果同目录能跑两个实例，后关闭的会把前一个实例刚改的键覆盖掉。
    /// 不同目录各自一份锁文件、各自一份 settings.json，所以不同安装目录可以同时运行。
    ///
    /// 目录不可写时放行（只读部署不能因为拿不到锁就起不来）。
    /// </summary>
    internal static class SingleInstanceGuard
    {
        private static FileStream? _lock;

        internal static bool TryAcquire()
        {
            try
            {
                var dir = Path.Combine(AppContext.BaseDirectory, ".wxd-work");
                Directory.CreateDirectory(dir);

                _lock = new FileStream(
                    Path.Combine(dir, "app.lock"),
                    FileMode.OpenOrCreate,
                    FileAccess.ReadWrite,
                    FileShare.None,
                    bufferSize: 1,
                    FileOptions.None);
                return true;
            }
            catch (IOException)
            {
                return false;   // 锁被占用：同目录已有实例
            }
            catch
            {
                return true;    // 权限/磁盘异常：不阻止启动
            }
        }
    }
}
