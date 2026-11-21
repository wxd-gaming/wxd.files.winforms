using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

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
    /// 同目录再次启动时不弹提示，而是通过按目录哈希命名的进程间事件通知已有实例显示窗口，
    /// 然后自己退出；不同目录的事件名不同，互不干扰。
    ///
    /// 目录不可写时放行（只读部署不能因为拿不到锁就起不来）。
    /// </summary>
    internal static class SingleInstanceGuard
    {
        private static FileStream? _lock;
        private static EventWaitHandle? _activateEvent;
        private static Thread? _listener;

        private const int AsfwAny = -1;

        [DllImport("user32.dll")]
        private static extern bool AllowSetForegroundWindow(int dwProcessId);

        // 事件名只由 exe 目录决定，同目录的每个进程算出来都一样
        private static readonly string ActivateEventName = BuildActivateEventName();

        internal static bool TryAcquire()
        {
            // 先建/打开事件再抢锁：第二个实例在锁失败后要用同一个句柄发信号，
            // 也避免"锁已抢到、事件还没建"时第二个实例把唤起请求丢掉
            EnsureActivateEvent();

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
            catch (IOException ex)
            {
                // 只有共享冲突/锁冲突才说明"已有实例"；磁盘满、路径被同名文件占用等不能误报
                int code = ex.HResult & 0xFFFF;
                return code != 32 && code != 33;
            }
            catch
            {
                return true;    // 权限/磁盘异常：不阻止启动
            }
        }

        /// <summary>同目录已有实例时调用：通知它把窗口显示出来。返回是否送达。</summary>
        internal static bool SignalRunningInstance()
        {
            try
            {
                EnsureActivateEvent();
                if (_activateEvent == null) return false;

                // 第二个实例是用户双击启起来的，通常握有前台权限；
                // 转交给已有实例，否则它的 SetForegroundWindow 会被前台锁挡住，只闪任务栏
                try { AllowSetForegroundWindow(AsfwAny); } catch { }

                _activateEvent.Set();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>UI 就绪后调用：后台等待唤起请求，并回到 UI 线程执行。</summary>
        internal static void StartListener(Action onActivate)
        {
            EnsureActivateEvent();
            if (_listener != null || _activateEvent == null) return;

            _listener = new Thread(() =>
            {
                while (true)
                {
                    try
                    {
                        _activateEvent.WaitOne();
                    }
                    catch (ObjectDisposedException)
                    {
                        return;
                    }

                    try
                    {
                        onActivate();
                    }
                    catch
                    {
                        // 唤起失败不能把监听线程带崩，否则之后每次双击都没反应
                    }
                }
            })
            {
                IsBackground = true,
                Name = "wxd-activate-listener"
            };
            _listener.Start();
        }

        private static void EnsureActivateEvent()
        {
            if (_activateEvent != null) return;
            try
            {
                // 已存在就打开，不存在就创建；同目录的另一个进程拿到的都是同一个内核事件
                _activateEvent = new EventWaitHandle(false, EventResetMode.AutoReset, ActivateEventName);
            }
            catch
            {
                _activateEvent = null;
            }
        }

        private static string BuildActivateEventName()
        {
            try
            {
                var dir = Path.TrimEndingDirectorySeparator(Path.GetFullPath(AppContext.BaseDirectory));
                var hash = SHA256.HashData(Encoding.UTF8.GetBytes(dir.ToUpperInvariant()));
                return "wxd.files.winforms.activate." + Convert.ToHexString(hash, 0, 8);
            }
            catch
            {
                // 连目录都拿不到时退回进程级名字：宁可不跨实例唤起，也不能把不同目录串在一起
                return "wxd.files.winforms.activate." + Environment.ProcessId;
            }
        }
    }
}
