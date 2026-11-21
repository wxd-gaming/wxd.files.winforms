using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    internal static class Program
    {
        /// <summary>防止同一异常在弹窗/记录过程中再次触发，形成递归</summary>
        private static int _handlingFatal;

        [DllImport("user32.dll")]
        private static extern uint GetGuiResources(IntPtr hProcess, uint uiFlags);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetCurrentProcess();

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            // 全局异常捕获：UI 线程异常
            // 提示用户后继续运行，不直接杀进程（弹窗里的异常信息已可复制反馈）
            Application.ThreadException += (sender, e) =>
            {
                LogFatal(e.Exception, "UI 线程异常");
                if (System.Threading.Interlocked.Exchange(ref _handlingFatal, 1) == 0)
                {
                    try
                    {
                        AlertForm.ShowException(null, e.Exception, "UI 线程异常");
                    }
                    catch (Exception alertEx)
                    {
                        // GDI/句柄耗尽时连错误框都可能建不出来，别让它覆盖原始异常
                        LogFatal(alertEx, "显示 UI 异常对话框失败");
                    }
                    finally
                    {
                        System.Threading.Interlocked.Exchange(ref _handlingFatal, 0);
                    }
                }
            };

            // 全局异常捕获：非 UI 线程异常（进程级，CLR 随后会终止进程，这里只记录 + 尽力提示）
            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                var ex = e.ExceptionObject as Exception;
                LogFatal(ex, "非 UI 线程未处理异常");

                // ★ 不要在这里创建 WinForms 对话框。
                // 非 UI 线程没有消息循环，而且走到这里时往往已经 GDI/句柄耗尽：
                // RichTextBox/Font 创建会再抛 ExternalException，覆盖掉真正的原始异常，
                // 之前的 crash 日志就是这么只剩一个“对象无法创建”的。
                // 只写日志，交给 CLR 终止；原始堆栈在 crash_*.log 里。
                // 不再显式 Environment.Exit：立即强退可能把正在写的配置文件/表格写坏。
            };

            // 全局异常捕获：Task 中未观察的异常
            TaskScheduler.UnobservedTaskException += (sender, e) =>
            {
                LogFatal(e.Exception, "Task 未处理异常");
                e.SetObserved();
            };

            
            ApplicationConfiguration.Initialize();

            // 同一 exe 目录只允许一个实例：settings.json 是整文件读改写，
            // 同目录双开会互相覆盖；不同目录各有一份锁和设置，可以同时运行。
            // 第二个实例不弹窗，通知已有实例显示窗口后自己退出。
            if (!SingleInstanceGuard.TryAcquire())
            {
                if (!SingleInstanceGuard.SignalRunningInstance())
                {
                    MessageBox.Show(
                        "同一目录已经有一个实例在运行，但没能唤起它的窗口。" + GlobalConst.NewLine + GlobalConst.NewLine +
                        "请从任务栏或右下角托盘手动打开；如需多开，请使用另一个目录的程序副本。",
                        "多窗格文件夹管理器",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                return;
            }

            // 按存档的皮肤+明暗激活令牌（含原生深色模式）
            Theme.Set((UiSkin)CsvSettings.ThemeSkin, CsvSettings.ThemeMode == 0);

            Application.Run(new Form1());
        }

        /// <summary>把未处理异常写到 .wxd-work/logs，目录不可写时退回临时目录</summary>
        private static void LogFatal(Exception? ex, string source)
        {
            try
            {
                var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                uint gdi = 0;
                try { gdi = GetGuiResources(GetCurrentProcess(), 0); } catch { }
                var content =
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {source} (v{Application.ProductVersion}, GDI={gdi}){GlobalConst.NewLine}" +
                    $"{ex}{GlobalConst.NewLine}{GlobalConst.NewLine}";

                var dirs = new[]
                {
                    Path.Combine(AppContext.BaseDirectory, ".wxd-work", "logs"),
                    Path.Combine(Path.GetTempPath(), "wxd.files.winforms-logs")
                };
                foreach (var dir in dirs)
                {
                    try
                    {
                        Directory.CreateDirectory(dir);
                        File.AppendAllText(Path.Combine(dir, $"crash_{stamp}.log"), content);
                        return;
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}
