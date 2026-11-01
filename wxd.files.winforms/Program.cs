using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    internal static class Program
    {
        /// <summary>防止同一异常在弹窗/记录过程中再次触发，形成递归</summary>
        private static int _handlingFatal;

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
                if (System.Threading.Interlocked.Exchange(ref _handlingFatal, 1) == 0)
                {
                    try
                    {
                        if (ex != null)
                            AlertForm.ShowException(null, ex, "非 UI 线程未处理异常");
                    }
                    catch { }
                    finally
                    {
                        System.Threading.Interlocked.Exchange(ref _handlingFatal, 0);
                    }
                }
                // 不再显式 Environment.Exit：立即强退可能把正在写的配置文件/表格写坏，
                // 交给 CLR 的默认终止流程
            };

            // 全局异常捕获：Task 中未观察的异常
            TaskScheduler.UnobservedTaskException += (sender, e) =>
            {
                LogFatal(e.Exception, "Task 未处理异常");
                e.SetObserved();
            };

            
            ApplicationConfiguration.Initialize();

            Application.Run(new Form1());
        }

        /// <summary>把未处理异常写到 .wxd-work/logs，目录不可写时退回临时目录</summary>
        private static void LogFatal(Exception? ex, string source)
        {
            try
            {
                var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                var content =
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {source}{GlobalConst.NewLine}" +
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
