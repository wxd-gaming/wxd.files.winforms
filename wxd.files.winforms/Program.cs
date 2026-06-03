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
            ApplicationConfiguration.Initialize();

            // 启动时异步检查更新（不阻塞主界面启动）
            _ = UpdateHelper.CheckSilentAsync();

            Application.Run(new Form1());
        }
    }
}
