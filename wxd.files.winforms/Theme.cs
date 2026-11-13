using System.Reflection;

namespace wxd.files.winforms
{
    /// <summary>界面皮肤：天蓝提质 / 密集工具 / 北欧蓝灰(Nord) / 暖沙护眼(Solarized) / 复古暖棕(Gruvbox) / 德古拉紫(Dracula) / 薄荷绿(Mint) / 淑女粉(Rose)</summary>
    internal enum UiSkin { Sky = 0, Dense = 1, Nord = 2, Solarized = 3, Gruvbox = 4, Dracula = 5, Mint = 6, Rose = 7 }

    /// <summary>一套皮肤在一个明暗模式下的全部视觉参数。</summary>
    internal sealed class UiTheme
    {
        public required Color Window;
        public required Color Chrome;
        public required Color Surface;
        public required Color Border;
        public required Color Divider;
        public required Color Header;
        public required Color Text;
        public required Color Muted;
        public required Color Accent;
        public required Color AccentText;
        public required Color Selection;
        public required Color Hover;
        public required Color Zebra;
        public required Color GridLine;
        public required Color RowHead;
        public required Color Link;
        public required Color Mark;
        public required Color Danger;

        public int RowHeight = 24;
        public int ButtonHeight = 26;
        public int Radius = 4;
        public int CellPadX = 6;
        public float FontSize = 9f;
        public bool ShowGridLines = true;
        public bool ZebraOn = true;
        public bool LeftAccent = true;
        public bool HeaderFill = true;
    }

    /// <summary>全局主题入口：皮肤 + 明暗，供各控件取色取字体，取代散落的硬编码色值。</summary>
    internal static class Theme
    {
        public static UiSkin Skin { get; private set; } = UiSkin.Sky;
        public static bool Dark { get; private set; }

        public static UiTheme P { get; private set; } = Build(UiSkin.Sky, false);

        /// <summary>切换前的那一套配色，<see cref="ThemeRecolor"/> 靠它把已经建好的控件颜色平移过来。</summary>
        public static UiTheme? Previous { get; private set; }

        public static Font Body { get; private set; } = new("Microsoft YaHei UI", 9f);
        public static Font List { get; private set; } = new("Microsoft YaHei UI", 9.5f);
        public static Font Head { get; private set; } = new("Segoe UI Semibold", 9f);
        public static Font Mono { get; private set; } = new("Consolas", 8.5f);
        public static Font Caption { get; private set; } = new("Microsoft YaHei UI", 8f);

        /// <summary>顶栏标题这类「比正文大一档」的标题字，明暗/皮肤切换后仍然和正文成比例</summary>
        public static Font Title { get; private set; } = new("Microsoft YaHei UI", 11f);

        /// <summary>空态/加载提示用的强调标题</summary>
        public static Font TitleBold { get; private set; } = new("Microsoft YaHei UI", 11f, FontStyle.Bold);

        /// <summary>正文加粗（表头行、当前页签这类"和正文同字号但要强调"的地方）</summary>
        public static Font BodyBold { get; private set; } = new("Microsoft YaHei UI", 9f, FontStyle.Bold);

        /// <summary>列表字加粗，和 <see cref="List"/> 同字号</summary>
        public static Font ListBold { get; private set; } = new("Microsoft YaHei UI", 9.5f, FontStyle.Bold);

        /// <summary>
        /// 等宽正文（单元格值、日志行这类要对齐的文本）。
        /// 这几个加粗/等宽字体必须在这里建一次：ApplyTheme 里各自 new 的话，每切一次皮肤就漏一个 GDI 字体句柄。
        /// </summary>
        public static Font Code { get; private set; } = new("Consolas", 10f);

        /// <summary>
        /// 皮肤或深浅色变了。主窗自己会刷，独立窗口（日志分析、阅读器）靠这个事件重新取令牌，
        /// 订阅方记得在 Disposed 时退订，静态事件会攥住已关闭的窗体。
        /// </summary>
        public static event Action? Changed;

        public static void Set(UiSkin skin, bool dark)
        {
            if (skin == Skin && dark == Dark && P != null) return;
            Skin = skin;
            Dark = dark;
            Previous = P;
            P = Build(skin, dark);
            // 旧字体不回收：可能仍被现存控件引用，回收会让它们绘制失败
            Body = new Font("Microsoft YaHei UI", P.FontSize);
            List = new Font("Microsoft YaHei UI", P.FontSize + 0.5f);
            // 标题字不能再比正文小：列头、分区标题都用它，8.75pt 配 9pt 正文会读成「标题比内容还弱」
            Head = new Font("Segoe UI Semibold", P.FontSize);
            Mono = new Font("Consolas", Math.Max(8f, P.FontSize - 0.5f));
            Caption = new Font("Microsoft YaHei UI", Math.Max(7.5f, P.FontSize - 1f));
            Title = new Font("Microsoft YaHei UI", P.FontSize + 2f);
            TitleBold = new Font("Microsoft YaHei UI", P.FontSize + 2f, FontStyle.Bold);
            BodyBold = new Font("Microsoft YaHei UI", P.FontSize, FontStyle.Bold);
            ListBold = new Font("Microsoft YaHei UI", P.FontSize + 0.5f, FontStyle.Bold);
            Code = new Font("Consolas", P.FontSize + 1f);
            ApplyNativeMode(dark);
            ApplyCaptions();
            Broadcast();
        }

        /// <summary>
        /// 逐个通知订阅方重取令牌。多播委托是一次性调用的，某个窗体的 ApplyTheme 抛异常
        /// 会让排在它后面的窗体全部不执行（后半界面停在旧配色上），异常还会回抛到切肤按钮，
        /// 所以这里按订阅顺序单独兜住。
        /// </summary>
        private static void Broadcast()
        {
            var handlers = Changed;
            if (handlers == null) return;
            foreach (Action handler in handlers.GetInvocationList())
            {
                try { handler(); } catch { }
            }
        }

        [System.Runtime.InteropServices.DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;   // Win10 1809~1903 上是 19

        /// <summary>
        /// 把每个已打开窗口的标题栏明暗刷成当前主题。
        /// 标题栏是非客户区，WinForms 画不到，只能问 DWM；SetColorMode 只管新建的窗口，
        /// 已经开着的窗口切主题时不会回头改这个属性，所以必须自己扫一遍。
        /// </summary>
        private static void ApplyCaptions()
        {
            // 同步跑在 UI 线程上，DWM 调用不会泵消息，集合中途不会变，直接遍历即可
            foreach (Form f in Application.OpenForms)
                if (f.IsHandleCreated) SetCaption(f.Handle);
        }

        private static void SetCaption(IntPtr hwnd)
        {
            int on = Dark ? 1 : 0;
            // 老运行库不认 20，退回 19；两个都失败就当系统默认，不影响客户区配色
            if (DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref on, sizeof(int)) != 0)
                DwmSetWindowAttribute(hwnd, 19, ref on, sizeof(int));
        }

        /// <summary>
        /// 让原生滚动条、下拉列表等跟随明暗。用反射是因为低版本运行库没有 SetColorMode，
        /// 缺失时静默跳过，不影响自绘部分。
        /// </summary>
        private static void ApplyNativeMode(bool dark)
        {
            try
            {
                var app = typeof(Application);
                var method = app.GetMethod("SetColorMode");
                var enumType = app.Assembly.GetType("System.Windows.Forms.SystemColorMode");
                if (method == null || enumType == null) return;
                // 运行库的枚举成员不一定有 Light（当前只有 Classic/System/Dark），
                // 直接解析会抛异常被下面吞掉，结果浅色模式永远停在深色原生外观上。
                string name = dark ? "Dark" : "Light";
                if (Array.IndexOf(Enum.GetNames(enumType), name) < 0) name = "Classic";
                method.Invoke(null, new object[] { Enum.Parse(enumType, name) });
            }
            catch
            {
                // 运行库不支持就保持系统默认
            }
        }

        private static UiTheme Build(UiSkin skin, bool dark)
        {
            if (skin == UiSkin.Dense)
            {
                return dark
                    ? new UiTheme
                    {
                        Window = C("1D1F22"), Chrome = C("23262A"), Surface = C("212428"),
                        Border = C("35393F"), Divider = C("2B2F34"), Header = C("262A2F"),
                        Text = C("DCE0E5"), Muted = C("8B929B"), Accent = C("5B8DD0"),
                        AccentText = Color.White, Selection = C("2C3F5A"), Hover = C("2F343A"), Zebra = C("232629"),
                        GridLine = C("2E3238"), RowHead = C("232629"), Link = C("7FB0E8"), Mark = C("E08366"),
                        Danger = C("F0796B"),
                        RowHeight = 22, ButtonHeight = 24, Radius = 3, CellPadX = 5, FontSize = 9f,
                    }
                    : new UiTheme
                    {
                        Window = C("F2F3F5"), Chrome = C("E9EBEF"), Surface = C("FFFFFF"),
                        Border = C("D2D6DC"), Divider = C("E4E7EB"), Header = C("EEF0F3"),
                        Text = C("1B1D21"), Muted = C("5F666E"), Accent = C("3A6EA5"),
                        AccentText = Color.White, Selection = C("D9E6F5"), Hover = C("DDE1E6"), Zebra = C("F8F9FB"),
                        GridLine = C("E7E9ED"), RowHead = C("F2F3F5"), Link = C("2A5FA8"), Mark = C("B4442B"),
                        Danger = C("B03A2E"),
                        RowHeight = 22, ButtonHeight = 24, Radius = 3, CellPadX = 5, FontSize = 9f,
                    };
            }

            // Nord：官方 Polar Night（深）/ Snow Storm（浅）色板
            if (skin == UiSkin.Nord)
            {
                return dark
                    ? new UiTheme
                    {
                        Window = C("2E3440"), Chrome = C("3B4252"), Surface = C("434C5E"),
                        Border = C("4C566A"), Divider = C("404859"), Header = C("505A6E"),
                        Text = C("ECEFF4"), Muted = C("929CAD"), Accent = C("5E81AC"),
                        AccentText = Color.White, Selection = C("4A5B78"), Hover = C("4A5364"), Zebra = C("3E4553"),
                        GridLine = C("485162"), RowHead = C("404857"), Link = C("88C0D0"), Mark = C("D08770"),
                        Danger = C("BF616A"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    }
                    : new UiTheme
                    {
                        Window = C("ECEFF4"), Chrome = C("E5E9F0"), Surface = C("FFFFFF"),
                        Border = C("C8D0DD"), Divider = C("D8DEE9"), Header = C("DFE5EE"),
                        Text = C("2E3440"), Muted = C("616E82"), Accent = C("5E81AC"),
                        AccentText = Color.White, Selection = C("CFDAE8"), Hover = C("E0E6EF"), Zebra = C("F4F6FA"),
                        GridLine = C("E0E5EE"), RowHead = C("EFF2F7"), Link = C("4C74A0"), Mark = C("D08770"),
                        Danger = C("BF616A"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    };
            }

            // Solarized：官方 base03~base3 暖青/米黄色板，低对比护眼是它的设计本意
            if (skin == UiSkin.Solarized)
            {
                return dark
                    ? new UiTheme
                    {
                        Window = C("002B36"), Chrome = C("073642"), Surface = C("04313D"),
                        Border = C("274C56"), Divider = C("0A3E4A"), Header = C("0B4250"),
                        Text = C("93A1A1"), Muted = C("657B83"), Accent = C("268BD2"),
                        AccentText = Color.White, Selection = C("104F6D"), Hover = C("0A4050"), Zebra = C("03303C"),
                        GridLine = C("0D4452"), RowHead = C("053541"), Link = C("2AA198"), Mark = C("B58900"),
                        Danger = C("DC322F"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    }
                    : new UiTheme
                    {
                        Window = C("FDF6E3"), Chrome = C("EEE8D5"), Surface = C("FFFBF0"),
                        Border = C("D8D0B8"), Divider = C("E9E2CB"), Header = C("E7E0C8"),
                        Text = C("586E75"), Muted = C("657B83"), Accent = C("268BD2"),
                        AccentText = Color.White, Selection = C("E4DCC2"), Hover = C("EAE3CC"), Zebra = C("F7F0DC"),
                        GridLine = C("E6DFC8"), RowHead = C("F5EED8"), Link = C("1E78BE"), Mark = C("CB4B16"),
                        Danger = C("DC322F"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    };
            }

            // Gruvbox：官方 retro 暖棕/米黄色板
            if (skin == UiSkin.Gruvbox)
            {
                return dark
                    ? new UiTheme
                    {
                        Window = C("282828"), Chrome = C("32302F"), Surface = C("3C3836"),
                        Border = C("504945"), Divider = C("4A4440"), Header = C("45403D"),
                        Text = C("EBDBB2"), Muted = C("A89984"), Accent = C("D65D0E"),
                        AccentText = Color.White, Selection = C("4E463C"), Hover = C("423D3A"), Zebra = C("353230"),
                        GridLine = C("44403C"), RowHead = C("383432"), Link = C("83A598"), Mark = C("FABD2F"),
                        Danger = C("FB4934"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    }
                    : new UiTheme
                    {
                        Window = C("F9F5D7"), Chrome = C("F2E5BC"), Surface = C("FBF1C7"),
                        Border = C("D5C4A1"), Divider = C("E8DAB0"), Header = C("EBDBB2"),
                        Text = C("3C3836"), Muted = C("7C6F64"), Accent = C("AF3A03"),
                        AccentText = Color.White, Selection = C("E5D5A5"), Hover = C("F0E2B6"), Zebra = C("F5EDC0"),
                        GridLine = C("E6D9AE"), RowHead = C("F6EFC8"), Link = C("076678"), Mark = C("B57614"),
                        Danger = C("9D0006"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    };
            }

            // Dracula：官方深色色板 + Alucard 浅色变体；Accent 用加深的紫（官方 #BD93F9 太亮，白字读不出）
            if (skin == UiSkin.Dracula)
            {
                return dark
                    ? new UiTheme
                    {
                        Window = C("191A21"), Chrome = C("21222C"), Surface = C("282A36"),
                        Border = C("3E4152"), Divider = C("353849"), Header = C("343746"),
                        Text = C("F8F8F2"), Muted = C("818EB6"), Accent = C("8A63D2"),
                        AccentText = Color.White, Selection = C("44475A"), Hover = C("3A3D4E"), Zebra = C("2C2E3B"),
                        GridLine = C("383B4A"), RowHead = C("262834"), Link = C("8BE9FD"), Mark = C("FFB86C"),
                        Danger = C("FF5555"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    }
                    : new UiTheme
                    {
                        Window = C("F7F2E0"), Chrome = C("F0EAD4"), Surface = C("FFFBEB"),
                        Border = C("D8D2BC"), Divider = C("E5DFC9"), Header = C("ECE7D5"),
                        Text = C("1F1F1F"), Muted = C("635D97"), Accent = C("644AC9"),
                        AccentText = Color.White, Selection = C("CFCFDE"), Hover = C("EDE7D1"), Zebra = C("FAF6E6"),
                        GridLine = C("E8E2CC"), RowHead = C("F8F4E2"), Link = C("036A96"), Mark = C("A34D14"),
                        Danger = C("CB3A2A"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    };
            }

            // 薄荷绿：自定义清新青绿，浅色薄荷白底 / 深色深茶绿底
            if (skin == UiSkin.Mint)
            {
                return dark
                    ? new UiTheme
                    {
                        Window = C("14201C"), Chrome = C("1A2924"), Surface = C("1E302A"),
                        Border = C("2E4A40"), Divider = C("253C35"), Header = C("28423A"),
                        Text = C("DFF2EA"), Muted = C("8FB5A8"), Accent = C("35A882"),
                        AccentText = Color.White, Selection = C("2C5245"), Hover = C("264038"), Zebra = C("1B2C26"),
                        GridLine = C("2A453D"), RowHead = C("1D2F28"), Link = C("6FDDB4"), Mark = C("F0A868"),
                        Danger = C("FF7B6B"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    }
                    : new UiTheme
                    {
                        Window = C("EFF7F4"), Chrome = C("E2F0EA"), Surface = C("FFFFFF"),
                        Border = C("C2DED4"), Divider = C("DCEEE7"), Header = C("E6F3EE"),
                        Text = C("1E4D40"), Muted = C("5E8A7C"), Accent = C("2E9E7B"),
                        AccentText = Color.White, Selection = C("D5EEE4"), Hover = C("E0F1EA"), Zebra = C("F4FAF7"),
                        GridLine = C("DEEFE8"), RowHead = C("EDF7F3"), Link = C("1F7A5E"), Mark = C("E08A3C"),
                        Danger = C("C44536"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    };
            }

            // 淑女粉：自定义柔和玫瑰粉，浅色腮红底 / 深色深梅紫底
            if (skin == UiSkin.Rose)
            {
                return dark
                    ? new UiTheme
                    {
                        Window = C("201419"), Chrome = C("2A1B22"), Surface = C("302028"),
                        Border = C("4A333D"), Divider = C("38252E"), Header = C("402B35"),
                        Text = C("F5E3E9"), Muted = C("B98E9D"), Accent = C("CC6288"),
                        AccentText = Color.White, Selection = C("543247"), Hover = C("3A2731"), Zebra = C("2B1C24"),
                        GridLine = C("3E2A34"), RowHead = C("2D1E26"), Link = C("F0A0BC"), Mark = C("F0A868"),
                        Danger = C("FF7B8E"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    }
                    : new UiTheme
                    {
                        Window = C("FBF2F4"), Chrome = C("F6E4E8"), Surface = C("FFFFFF"),
                        Border = C("E8C9D1"), Divider = C("F2DDE2"), Header = C("F7E7EB"),
                        Text = C("5C2A3A"), Muted = C("A06B7C"), Accent = C("C2557E"),
                        AccentText = Color.White, Selection = C("F5DCE3"), Hover = C("F8E6EB"), Zebra = C("FCF5F7"),
                        GridLine = C("F0DCE1"), RowHead = C("F9EEF1"), Link = C("A84466"), Mark = C("D9822B"),
                        Danger = C("C43C50"),
                        RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                    };
            }

            return dark
                ? new UiTheme
                {
                    Window = C("161A20"), Chrome = C("1C222A"), Surface = C("1F2630"),
                    Border = C("2E3947"), Divider = C("263041"), Header = C("232C38"),
                    Text = C("E3EEF9"), Muted = C("92A6BC"), Accent = C("4E9BE8"),
                    AccentText = Color.White, Selection = C("23384F"), Hover = C("2A3543"), Zebra = C("1D232C"),
                    GridLine = C("28313C"), RowHead = C("1D232C"), Link = C("7FC0F5"), Mark = C("F0A070"),
                    Danger = C("FF6B6B"),
                    RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                }
                : new UiTheme
                {
                    Window = C("F1F7FC"), Chrome = C("E7F2FA"), Surface = C("FFFFFF"),
                    Border = C("C3DCEF"), Divider = C("DCEBF7"), Header = C("EAF4FC"),
                    Text = C("143C64"), Muted = C("5A82AA"), Accent = C("2E8BE0"),
                    AccentText = Color.White, Selection = C("D8ECFB"), Hover = C("E3EFF9"), Zebra = C("F7FBFE"),
                    GridLine = C("E2EFF8"), RowHead = C("EFF6FC"), Link = C("1565C0"), Mark = C("E07B39"),
                    Danger = C("C43C31"),
                    RowHeight = 24, ButtonHeight = 26, Radius = 4, CellPadX = 6, FontSize = 9f,
                };
        }

        private static Color C(string rgb) => Color.FromArgb(
            Convert.ToInt32(rgb.Substring(0, 2), 16),
            Convert.ToInt32(rgb.Substring(2, 2), 16),
            Convert.ToInt32(rgb.Substring(4, 2), 16));

        /// <summary>按比例混合两个颜色，t=0 全取 a，t=1 全取 b。</summary>
        public static Color Mix(Color a, Color b, float t)
        {
            t = Math.Clamp(t, 0f, 1f);
            return Color.FromArgb(
                (int)(a.R + (b.R - a.R) * t),
                (int)(a.G + (b.G - a.G) * t),
                (int)(a.B + (b.B - a.B) * t));
        }

        /// <summary>颜色是否偏暗，用来决定压在其上的文字该用白还是用墨。</summary>
        public static bool IsDark(Color c) => (c.R * 299 + c.G * 587 + c.B * 114) / 1000 < 145;

        /// <summary>
        /// 语义单元格底色（改过的浅红、重复的浅黄等）：
        /// 浅色模式直接用淡彩；深色模式把它压向表面色，否则浅底配浅字会读不出来。
        /// </summary>
        public static Color CellTint(Color pastel) => Dark ? Mix(pastel, P.Surface, 0.66f) : pastel;

        /// <summary>语义前景色（错误红/警告橙/成功绿）：深色模式提亮，保持和底色的对比。</summary>
        public static Color Signal(Color c) => Dark ? Mix(c, Color.White, 0.45f) : c;
    }

    /// <summary>
    /// 控件颜色基本是在创建时一次性赋成当时令牌的固定值，切皮肤或明暗后不会自己跟着变。
    /// 让每个窗口把子控件再逐个赋一遍注定会漏（TableBrowserControl 里上百个局部控件那种），
    /// 所以这里按「上一套配色的令牌 → 当前配色的同名令牌」把整棵控件树平移一次：
    /// 只动恰好等于旧令牌值的前景色和背景色，硬编码色与派生色原样保留，重复调用不改变结果。
    /// </summary>
    internal static class ThemeRecolor
    {
        private static readonly FieldInfo[] ColorTokens = typeof(UiTheme)
            .GetFields(BindingFlags.Public | BindingFlags.Instance)
            .Where(f => f.FieldType == typeof(Color))
            .ToArray();

        public static void Apply(Control root)
        {
            var prev = Theme.Previous;
            if (prev == null || ReferenceEquals(prev, Theme.P)) return;

            var map = new Dictionary<Color, Color>();
            foreach (var token in ColorTokens)
                map[(Color)token.GetValue(prev)!] = (Color)token.GetValue(Theme.P)!;

            Walk(root, map);
        }

        private static void Walk(Control c, Dictionary<Color, Color> map)
        {
            if (map.TryGetValue(c.BackColor, out var back) && back != c.BackColor) c.BackColor = back;
            if (map.TryGetValue(c.ForeColor, out var fore) && fore != c.ForeColor) c.ForeColor = fore;
            foreach (Control child in c.Controls) Walk(child, map);
        }
    }
}
