using System.Reflection;

namespace wxd.files.winforms
{
    /// <summary>
    /// 标题栏图标的统一入口。图标全部是编译进程序的嵌入资源，按名字取、取过就缓存，
    /// 这样发布目录不用带 icons 文件夹，也避免「磁盘上的 app.ico 没拷过去，图标静默丢失」。
    /// </summary>
    internal static class WindowIcons
    {
        private static readonly Dictionary<string, Icon?> Cache = new(StringComparer.OrdinalIgnoreCase);

        public static Icon? App => Get("app");
        public static Icon? Reader => Get("reader");
        public static Icon? Logs => Get("logs");
        public static Icon? Notes => Get("notes");
        public static Icon? Vcmp => Get("vcmp");
        public static Icon? Cmdmgr => Get("cmdmgr");
        public static Icon? Dup => Get("dup");
        public static Icon? Colmgr => Get("colmgr");
        public static Icon? Celledit => Get("celledit");
        public static Icon? Link => Get("link");
        public static Icon? Jump => Get("jump");

        private static Icon? Get(string name)
        {
            lock (Cache)
            {
                if (Cache.TryGetValue(name, out var hit)) return hit;

                var asm = Assembly.GetExecutingAssembly();
                var full = asm.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith(name + ".ico", StringComparison.OrdinalIgnoreCase));
                Icon? icon = null;
                if (full != null)
                {
                    using var stream = asm.GetManifestResourceStream(full);
                    if (stream != null)
                    {
                        // 固定取 32px 这一档：16px 标题栏由系统降采样最不糊，任务栏放大也不明显。
                        icon = new Icon(stream, 32, 32);
                    }
                }

                Cache[name] = icon;
                return icon;
            }
        }
    }
}
