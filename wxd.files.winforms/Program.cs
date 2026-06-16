using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // 全局异常捕获：UI 线程异常
            Application.ThreadException += (sender, e) =>
            {
                AlertForm.ShowException(null, e.Exception, "UI 线程异常");
                Environment.Exit(1);
            };

            // 全局异常捕获：非 UI 线程异常（进程级，无法阻止退出，但可弹窗提示）
            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                if (e.ExceptionObject is Exception ex)
                {
                    AlertForm.ShowException(null, ex, "非 UI 线程未处理异常");
                }
                Environment.Exit(1);
            };

            // 全局异常捕获：Task 中未观察的异常
            TaskScheduler.UnobservedTaskException += (sender, e) =>
            {
                AlertForm.ShowException(null, e.Exception, "Task 未处理异常");
                e.SetObserved();
            };

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            ApplicationConfiguration.Initialize();

            Application.Run(new Form1());
        }
    }
}
