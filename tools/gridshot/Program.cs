using System.Reflection;

// 顶层语句生成的 Main 不带 STAThread，WinForms 里碰 OLE 的控件会直接抛异常，先把自己搬到 STA 线程
if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
{
    var sta = new Thread(() => Run(args));
    sta.SetApartmentState(ApartmentState.STA);
    sta.Start();
    sta.Join();
    return;
}
Run(args);

static void Run(string[] args)
{
    var binDir = Environment.GetEnvironmentVariable("THEMELAB_BIN")
                 ?? @"D:\work\wxd.files.winforms\wxd.files.winforms\bin\Debug\net10.0-windows7.0";
    var asm = Assembly.LoadFrom(Path.Combine(binDir, "wxd.files.winforms.dll"));

    int skin = args.Length > 0 ? int.Parse(args[0]) : 0;
    bool dark = args.Length < 2 || args[1] != "0";
    string file = args.Length > 2 ? args[2] : @"D:\策划配置\cfg_monster.csv";

    var theme = asm.GetType("wxd.files.winforms.Theme")!;
    var skinEnum = asm.GetType("wxd.files.winforms.UiSkin")!;
    theme.GetMethod("Set")!.Invoke(null, new object[] { Enum.GetValues(skinEnum).GetValue(skin), dark });

    var optionsType = asm.GetType("wxd.files.winforms.TableBrowserOptions")!;
    var controlType = asm.GetType("wxd.files.winforms.TableBrowserControl")!;

    Application.SetHighDpiMode(HighDpiMode.SystemAware);
    Application.EnableVisualStyles();
    Application.SetCompatibleTextRenderingDefault(false);

    // 运行库到底给了哪些明暗 API：Theme.ApplyNativeMode 靠反射猜枚举名，猜错就等于没调
    if (args.Length > 3 && args[3] == "colormode")
    {
        var sb = new System.Text.StringBuilder();
        var appType = typeof(Application);
        sb.AppendLine("SetColorMode=" + (appType.GetMethod("SetColorMode")?.ToString() ?? "<none>"));
        var cmMode = appType.Assembly.GetType("System.Windows.Forms.SystemColorMode");
        sb.AppendLine("SystemColorMode=" + (cmMode == null ? "<none>" : string.Join(",", Enum.GetNames(cmMode))));
        sb.AppendLine("Form.ColorMode=" + (typeof(Form).GetProperty("ColorMode")?.ToString() ?? "<none>"));
        sb.AppendLine("Application.ColorMode=" + (appType.GetProperty("ColorMode")?.ToString() ?? "<none>"));
        sb.AppendLine("ColorModeChanged event=" + (appType.GetEvent("ColorModeChanged")?.ToString() ?? "<none>"));
        File.WriteAllText(@"D:\work\wxd.files.winforms\tools\restyle-shots\colormode.txt", sb.ToString());
        return;
    }

    if (args.Length > 3 && args[3] == "main")
    {
        var settings = asm.GetType("wxd.files.winforms.CsvSettings")!;
        settings.GetProperty("ThemeMode")!.SetValue(null, dark ? 0 : 1);
        settings.GetProperty("ThemeSkin")!.SetValue(null, skin);

        var mainType = asm.GetType("wxd.files.winforms.Form1")!;
        var main = (Form)Activator.CreateInstance(mainType, true)!;

        if (args.Length > 4 && args[4] == "smoke")
            SmokeTest(mainType, main, args.Length > 5 ? args[5] : @"D:\work\wxd.files.winforms\tools\restyle-shots\smoke.txt");

        if (args.Length > 4 && args[4] == "bargeom")
        {
            bool hold = Array.IndexOf(args, "hold") >= 0;
            BarGeom(mainType, main,
                args.Length > 5 && args[5] != "hold" ? args[5]
                    : @"D:\work\wxd.files.winforms\tools\restyle-shots\bar-geom.txt",
                hold);
        }

        if (args.Length > 4 && args[4].StartsWith("switch"))
        {
            // 起来后隔 2.5s 在同一个实例里翻一次明暗，用来抓"运行时切主题后没刷新的控件"
            // "switchrec" 再多做一步：把可编辑下拉框的窗口句柄重建，看箭头块颜色会不会跟着新主题走
            // "switchrec"/"switchcm" 再多做一步：先按 args[5] 指定的枚举名重设原生明暗，再重建可编辑下拉框句柄，
            // 看箭头块颜色会不会跟着新主题走（args[5] 缺省试 Light）
            bool recreate = args[4] is "switchrec" or "switchcm";
            main.Shown += (s, e) =>
            {
                var timer = new System.Windows.Forms.Timer { Interval = 2500 };
                timer.Tick += (_, _) =>
                {
                    timer.Stop();
                    var modeType = mainType.GetNestedType("ThemeMode", BindingFlags.NonPublic)!;
                    var modes = Enum.GetValues(modeType);
                    mainType.GetField("_currentTheme", BindingFlags.Instance | BindingFlags.NonPublic)!
                          .SetValue(main, dark ? modes.GetValue(1) : modes.GetValue(0));
                    mainType.GetMethod("ApplyTheme", BindingFlags.Instance | BindingFlags.NonPublic)!
                          .Invoke(main, Array.Empty<object>());
                    if (recreate)
                    {
                        var log = "";
                        try
                        {
                            var cm = typeof(Application).GetProperty("ColorMode");
                            var setMode = typeof(Application).GetMethod("SetColorMode");
                            var cmType = typeof(Application).Assembly.GetType("System.Windows.Forms.SystemColorMode");
                            string want = args.Length > 5 ? args[5] : "Light";
                            log += $"enumVals={string.Join('/', Enum.GetNames(cmType!))} before={cm?.GetValue(null)} want={want}";
                            try
                            {
                                setMode!.Invoke(null, new object[] { Enum.Parse(cmType!, want) });
                                log += $" set=ok after={cm?.GetValue(null)}";
                            }
                            catch (Exception se)
                            {
                                log += $" setEX={se.InnerException?.GetType().Name ?? se.GetType().Name}:{se.InnerException?.Message ?? se.Message}";
                            }
                            // 原生件的颜色可能只在建句柄时定死，所以分两种：重建句柄 vs 不重建，看哪种能刷新
                            int n = 0;
                            bool noRec = Array.IndexOf(args, "norec") >= 0;
                            if (!noRec)
                                foreach (var c in Walk(main))
                            {
                                if (c is not System.Windows.Forms.ComboBox cbo) continue;
                                if (cbo.DropDownStyle != System.Windows.Forms.ComboBoxStyle.DropDown) continue;
                                cbo.GetType().GetMethod("RecreateHandle", BindingFlags.Instance | BindingFlags.NonPublic)!
                                  .Invoke(cbo, Array.Empty<object>());
                                n++;
                            }
                            log += $" recreated={n}";
                            // 表头同样是原生件：分别试"只让表头失效重绘"和"重建 ListView 句柄"两种代价
                            string lvAction = Array.IndexOf(args, "lvrec") >= 0 ? "rec"
                                            : Array.IndexOf(args, "lvinv") >= 0 ? "inv"
                                            : Array.IndexOf(args, "lvtheme") >= 0 ? "theme" : "";
                            // lvtheme- 表示传 null（回到经典主题），lvthemeXxx 表示传 "Xxx"
                            string themeName = "Explorer";
                            foreach (var a in args)
                            {
                                if (!a.StartsWith("lvtheme")) continue;
                                lvAction = "theme";
                                themeName = a == "lvtheme-" ? null : a.Substring("lvtheme".Length);
                            }
                            if (lvAction != "")
                            {
                                int k = 0;
                                foreach (var c in Walk(main))
                                {
                                    if (c is not System.Windows.Forms.ListView lv) continue;
                                    if (!lv.IsHandleCreated) continue;
                                    if (lvAction == "rec")
                                        lv.GetType().GetMethod("RecreateHandle", BindingFlags.Instance | BindingFlags.NonPublic)!
                                          .Invoke(lv, Array.Empty<object>());
                                    else
                                    {
                                        IntPtr hdr = Native.SendMessage(lv.Handle, Native.LVM_GETHEADER, IntPtr.Zero, IntPtr.Zero);
                                        if (hdr != IntPtr.Zero)
                                        {
                                            if (lvAction == "theme")
                                                Native.SetWindowTheme(hdr, themeName, null);
                                            Native.InvalidateRect(hdr, IntPtr.Zero, true);
                                            Native.UpdateWindow(hdr);
                                        }
                                    }
                                    k++;
                                }
                                log += $" lv[{lvAction}]={k}";
                            }
                            log += $" final={cm?.GetValue(null)}";
                        }
                        catch (Exception ex)
                        {
                            log += "EX:" + ex.GetType().Name + " " + ex.Message;
                        }
                        File.WriteAllText(@"D:\work\wxd.files.winforms\tools\restyle-shots\recreate-count.txt", log.ToString());
                    }
                };
                timer.Start();
            };
        }

        Application.Run(main);
        return;
    }

    if (args.Length > 3 && args[3] == "reader")
    {
        var readerType = asm.GetType("wxd.files.winforms.CsvReaderForm")!;
        var reader = (Form)Activator.CreateInstance(readerType, new object[] { file, -1 })!;
        if (args.Length > 4)
        {
            var fileType = asm.GetType("wxd.files.winforms.CsvReaderForm+FileType")!;
            reader.Shown += (s, e) =>
            {
                readerType.GetMethod("OpenFile", BindingFlags.Instance | BindingFlags.NonPublic)!
                      .Invoke(reader, new object[] { args[4], Enum.GetValues(fileType).GetValue(0) });

                // 构造时传入的是文件路径，左侧目录树不会自动扫描；这里补上真实使用时的加载入口
                var dirField = readerType.GetField("_currentDir", BindingFlags.Instance | BindingFlags.NonPublic)!;
                var parent = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(args[4])) ?? "";
                dirField.SetValue(reader, parent);
                readerType.GetMethod("LoadFilesAsync", BindingFlags.Instance | BindingFlags.NonPublic)!
                      .Invoke(reader, Array.Empty<object>());
            };
        }
        // args 里带 "switch"：表格里开了大表后隔 2.5s 运行时翻一次明暗，
        // TableBrowserControl 的配色大半是构造时一次性写进 DataGridViewCellStyle/BackColor 的局部值，
        // 只有这条路能证明 ThemeRecolor + Theme.Changed 订阅真的把它们平移过来了
        if (Array.IndexOf(args, "switch") >= 0)
        {
            reader.Shown += (s, e) =>
            {
                var t = new System.Windows.Forms.Timer { Interval = 2500 };
                t.Tick += (_, _) =>
                {
                    t.Stop();
                    theme.GetMethod("Set")!.Invoke(null, new object[] { Enum.GetValues(skinEnum).GetValue(skin), !dark });
                };
                t.Start();
            };
        }
        // 看门狗：100ms 的定时器只在 UI 线程空闲时才触发，两次心跳之间的大间隔就是「界面没响应」的实测时长
        var wdSb = new System.Text.StringBuilder();
        if (Array.IndexOf(args, "watchdog") >= 0)
        {
            reader.Shown += (s, e) =>
            {
                long start = Environment.TickCount64, last = 0;
                int ticks = 0;
                var wd = new System.Windows.Forms.Timer { Interval = 100 };
                wd.Tick += (_, _) =>
                {
                    long now = Environment.TickCount64;
                    if (last != 0 && now - last > 700)
                        wdSb.AppendLine($"UI 冻结 {now - last}ms @+{now - start}ms（第 {ticks} 次心跳）");
                    last = now; ticks++;
                };
                wd.Start();
            };
        }
        // args[5] == "rowmap"：加载完大表后用反射核对「显示行 ↔ 源文件物理行号」映射，
        // 行号口径是策划最依赖的东西，改成控件自己记账后必须拿数字证明没跑偏
        if (args.Length > 5 && args[5] == "rowmap")
        {
            reader.Shown += (s, e) =>
            {
                var t = new System.Windows.Forms.Timer { Interval = Array.IndexOf(args, "quick") >= 0 ? 6000 : 15000 };
                t.Tick += (_, _) =>
                {
                    t.Stop();
                    var outPath = args.Length > 6 ? args[6]
                        : @"D:\work\wxd.files.winforms\tools\restyle-shots\rowmap.txt";
                    string filterText = args.Length > 7 && args[7] != "quick" ? args[7] : "S0270D";
                    long[] navLines = args.Length > 8
                        ? args[8].Split(',', StringSplitOptions.RemoveEmptyEntries).Select(long.Parse).ToArray()
                        : new long[] { 1, 60000, 119999 };
                    var sb = new System.Text.StringBuilder();
                    try
                    {
                        foreach (var c in Walk(reader))
                        {
                            var ct = c.GetType();
                            if (ct.Name != "TableBrowserControl") continue;
                            const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic;
                            var map = ct.GetField("_rowTags", flags)!.GetValue(c) as Array;
                            var dgv = ct.GetField("_dataGrid", flags)!.GetValue(c) as DataGridView;
                            var fileData = ct.GetField("_fileData", flags)!.GetValue(c);
                            var rowsObj = fileData!.GetType().GetProperty("Rows")!.GetValue(fileData)!;
                            int dataRows = (int)rowsObj.GetType().GetProperty("Count")!.GetValue(rowsObj)!;
                            sb.AppendLine($"rowTags={map?.Length} gridRows={dgv?.Rows.Count} fileDataRows={dataRows}");

                            var getLine = ct.GetMethod("GetDisplayLineNumber", flags)!;
                            var findLine = ct.GetMethod("FindDisplayRowByLineNumber", flags)!;
                            var getOrig = ct.GetMethod("GetOriginalRowIndexFromDisplay", flags)!;
                            var findOrig = ct.GetMethod("FindDisplayRowByOrigIndex", flags)!;
                            int n = map?.Length ?? 1;
                            int[] samples = { 0, 1, 2, 3, 4, Math.Min(99, n - 1), n / 2, n - 2, n - 1 };
                            foreach (int i in samples.Distinct().Where(x => x >= 0))
                            {
                                object? line = getLine.Invoke(c, new object[] { i });
                                object? orig = getOrig.Invoke(c, new object[] { i });
                                object? backLine = line is int l
                                    ? findLine.Invoke(c, new object[] { (long)l }) : null;
                                object? backOrig = orig is int o
                                    ? findOrig.Invoke(c, new object[] { o }) : null;
                                sb.AppendLine($"display[{i,7}] line={line,-8} orig={orig,-8} " +
                                              $"roundTrip(line)={backLine,-8} roundTrip(orig)={backOrig}");
                            }

                            // 过滤态：高级过滤只留命中 filterText 的那几行，看映射是否跟着 VisibleRowIndices 走
                            var txt = ct.GetField("_txtAdvancedFilter", flags)!.GetValue(c) as TextBox;
                            var sw = System.Diagnostics.Stopwatch.StartNew();
                            long last = 0;
                            string Elapsed(string tag)
                            {
                                long now = sw.ElapsedMilliseconds;
                                string s2 = $"{tag} {now - last}ms";
                                last = now;
                                return s2;
                            }
                            txt!.Text = filterText;
                            ct.GetMethod("ApplyFilters", flags)!.Invoke(c, Array.Empty<object>());
                            sb.AppendLine($"高级过滤 \"{filterText}\" {Elapsed("用时")}");
                            object? fLine = getLine.Invoke(c, new object[] { 0 });
                            sb.AppendLine($"filtered rows={dgv?.Rows.Count} display[0] line={fLine} " +
                                          $"roundTrip={findLine.Invoke(c, new object[] { Convert.ToInt64(fLine ?? 0) })}");
                            for (int i = 1; i < Math.Min(dgv!.Rows.Count, 6); i++)
                                sb.AppendLine($"   filtered display[{i}] line={getLine.Invoke(c, new object[] { i })}");
                            txt.Text = "";
                            ct.GetMethod("ApplyFilters", flags)!.Invoke(c, Array.Empty<object>());
                            sb.AppendLine(Elapsed("还原全量"));

                            var nav = ct.GetMethod("NavigateToPhysicalLine",
                                flags | BindingFlags.Public)!;
                            foreach (long line in navLines)
                            {
                                object? ok = nav.Invoke(c, new object[] { line });
                                int cur = dgv!.CurrentRow?.Index ?? -1;
                                sb.AppendLine($"NavigateToPhysicalLine({line}) = {ok} {Elapsed("跳转")} " +
                                              $"current[{cur}] line={getLine.Invoke(c, new object[] { cur })}");
                            }
                            int fv = dgv!.FirstDisplayedScrollingRowIndex;
                            sb.AppendLine($"after-nav firstVisible={fv} itsLine={getLine.Invoke(c, new object[] { fv })} " +
                                          $"currentCellLine={getLine.Invoke(c, new object[] { dgv.CurrentRow?.Index ?? fv })}");

                            // 标记行 + 行编辑：这两条路径以前读 DataGridViewRow.Tag，现在都走行映射
                            string Dump(string fieldName)
                            {
                                var v = ct.GetField(fieldName, flags)!.GetValue(c) as System.Collections.IEnumerable;
                                return v == null ? "<null>" : string.Join(",", v.Cast<object>());
                            }
                            long firstMatched = Convert.ToInt64(fLine ?? 0);
                            int markRow = Convert.ToInt32(findLine.Invoke(c, new object[] { firstMatched }));
                            ct.GetMethod("ReplaceRowSelection", flags)!.Invoke(c, new object[] { markRow });
                            sb.AppendLine($"ReplaceRowSelection({markRow}) line={firstMatched} " +
                                          $"marked=[{Dump("_markedLineNumbers")}] selected=[{Dump("_selectedRows")}]");
                            txt.Text = filterText;
                            ct.GetMethod("ApplyFilters", flags)!.Invoke(c, Array.Empty<object>());
                            sb.AppendLine($"过滤后 marked=[{Dump("_markedLineNumbers")}] selected=[{Dump("_selectedRows")}]");
                            txt.Text = "";
                            ct.GetMethod("ApplyFilters", flags)!.Invoke(c, Array.Empty<object>());
                            sb.AppendLine($"还原后 marked=[{Dump("_markedLineNumbers")}] selected=[{Dump("_selectedRows")}]");

                            ct.GetMethod("EnterRowEditMode", flags)!.Invoke(c, new object[] { markRow, -1 });
                            sb.AppendLine($"EnterRowEditMode({markRow}) editingRow={ct.GetField("_editingRowIndex", flags)!.GetValue(c)} " +
                                          $"editingOrig={ct.GetField("_editingRowOrigIndex", flags)!.GetValue(c)} " +
                                          $"expectOrig={getOrig.Invoke(c, new object[] { markRow })}");
                            ct.GetMethod("ExitRowEditMode", flags)!.Invoke(c, new object[] { false });
                            sb.AppendLine($"ExitRowEditMode(false) editingRow={ct.GetField("_editingRowIndex", flags)!.GetValue(c)} " +
                                          $"selected=[{Dump("_selectedRows")}]");
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("EX " + ex);
                    }
                    sb.AppendLine("---- 加载期 UI 心跳（空=没有超过 700ms 的停顿） ----");
                    sb.Append(wdSb);
                    System.IO.File.WriteAllText(outPath, sb.ToString());
                };
                t.Start();
            };
        }
        Application.Run(reader);
        return;
    }

    if (args.Length > 3 && args[3] == "logs")
    {
        // 日志分析窗口：args[2] 当目录用，args[4]（可选）在启动后打开成标签页
        var viewerType = asm.GetType("wxd.files.winforms.LargeFileViewer")!;
        var viewer = (Form)Activator.CreateInstance(viewerType, new object[] { file })!;
        if (args.Length > 4)
        {
            string open = args[4];
            viewer.Shown += (s, e) =>
            {
                var timer = new System.Windows.Forms.Timer { Interval = 1200 };
                timer.Tick += (_, _) =>
                {
                    timer.Stop();
                    viewerType.GetMethod("OpenFile", BindingFlags.Instance | BindingFlags.NonPublic)!
                          .Invoke(viewer, new object[] { open });
                };
                timer.Start();
            };
        }
        // args[5] == "switch"：起来后隔 2.5s 翻一次明暗，看独立窗口（日志分析 + 里面的阅读器）会不会跟着主题事件重取令牌
        if (args.Length > 5 && args[5] == "switch")
        {
            viewer.Shown += (s, e) =>
            {
                var t = new System.Windows.Forms.Timer { Interval = 2500 };
                t.Tick += (_, _) =>
                {
                    t.Stop();
                    theme.GetMethod("Set")!.Invoke(null, new object[] { Enum.GetValues(skinEnum).GetValue(skin), !dark });
                };
                t.Start();
            };
        }
        // args[5] == "detail"：打开文件后再隔一会儿双击第 1 行，把行详情弹窗顶出来供截图
        // "detailraw" 再多走一步：点弹窗里的「原始」段，验分段切换的联动
        if (args.Length > 5 && (args[5] == "detail" || args[5] == "detailraw"))
        {
            viewer.Shown += (s, e) =>
            {
                var t = new System.Windows.Forms.Timer { Interval = 2200 };
                t.Tick += (_, _) =>
                {
                    t.Stop();
                    foreach (var c in Walk(viewer))
                    {
                        if (c.GetType().Name != "LargeFileReaderControl") continue;
                        // OnDoubleClick 有事件重载和行号重载两个同名方法，按参数挑行号那个
                        var mi = c.GetType().GetMethods(BindingFlags.Instance | BindingFlags.NonPublic)
                            .First(m => m.Name == "OnDoubleClick"
                                     && m.GetParameters().Length == 1
                                     && m.GetParameters()[0].ParameterType == typeof(long));
                        mi.Invoke(c, new object[] { 0L });
                        break;
                    }
                    if (args[5] != "detailraw") return;
                    var t2 = new System.Windows.Forms.Timer { Interval = 3000 };
                    t2.Tick += (_, _) =>
                    {
                        t2.Stop();
                        // 弹窗是 Show() 出来的独立窗体，只能从 OpenForms 里按标题找
                        var pop = Application.OpenForms.Cast<Form>().FirstOrDefault(f => f.Text.StartsWith("行 "));
                        var seg = pop == null ? null : Walk(pop).FirstOrDefault(c => c.GetType().Name == "ThemeSegmented");
                        if (seg == null) return;
                        // 合成鼠标点击在没到前台的窗口上打不准，直接走控件自己的鼠标路径
                        var pt = new System.Drawing.Point(seg.Width - 30, seg.Height / 2);
                        var me = new MouseEventArgs(MouseButtons.Left, 1, pt.X, pt.Y, 0);
                        var flags = BindingFlags.Instance | BindingFlags.NonPublic;
                        typeof(Control).GetMethod("OnMouseDown", flags)!.Invoke(seg, new object[] { me });
                        typeof(Control).GetMethod("OnMouseUp", flags)!.Invoke(seg, new object[] { me });
                    };
                    t2.Start();
                };
                t.Start();
            };
        }
        // args[5] == "geom"：把窗口压窄/压矮后导出工具区与状态栏的实际矩形，
        // 像素截图只能看个大概，折叠后状态栏有没有被挤掉要用数字说话
        if (args.Length > 5 && args[5] == "geom")
        {
            viewer.Shown += (s, e) =>
            {
                var t = new System.Windows.Forms.Timer { Interval = 1800 };
                // 先落一个占位文件，用来区分“Shown 没触发”和“tick 里出异常”
                System.IO.File.WriteAllText(args.Length > 6 ? args[6]
                    : @"D:\work\wxd.files.winforms\tools\restyle-shots\geom.txt", "pending\n");
                t.Tick += (_, _) =>
                {
                    t.Stop();
                    var sb = new System.Text.StringBuilder();
                    try
                    {
                        // 最大化状态下改 Size 是空操作，先落回普通窗口再量
                        viewer.WindowState = FormWindowState.Normal;
                        foreach (var size in new[] { new System.Drawing.Size(1000, 900), new System.Drawing.Size(1600, 900),
                                                     new System.Drawing.Size(1000, 900), new System.Drawing.Size(1000, 560),
                                                     new System.Drawing.Size(1600, 1200) })
                        {
                            viewer.Size = size;
                            // 同一 tick 里改完尺寸立刻量，布局还没走完，读到的是旧矩形；先跑几轮消息泵
                            for (int pump = 0; pump < 40; pump++)
                            {
                                Application.DoEvents();
                                System.Threading.Thread.Sleep(20);
                            }
                            viewer.Update();
                            foreach (var c in Walk(viewer))
                            {
                                if (c.GetType().Name != "LargeFileReaderControl") continue;
                                c.Update();
                                sb.AppendLine($"-- size={size.Width}x{size.Height} client={viewer.ClientSize.Width}x{viewer.ClientSize.Height}");
                                foreach (var name in new[] { "_toolPanel", "_toolDivider", "_statusBar", "_lstContent" })
                                {
                                    var f = c.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
                                    var cc = f?.GetValue(c) as Control;
                                    sb.AppendLine($"   {name} = {cc?.Bounds}");
                                }
                                sb.AppendLine($"   reader={c.Bounds} parent={c.Parent?.GetType().Name}:{c.Parent?.Bounds}");
                                var tp = c.GetType().GetField("_toolPanel", BindingFlags.Instance | BindingFlags.NonPublic)?.GetValue(c) as Control;
                                var sbr = c.GetType().GetField("_statusBar", BindingFlags.Instance | BindingFlags.NonPublic)?.GetValue(c) as Control;
                                if (tp != null && sbr != null)
                                    sb.AppendLine($"   overlap={tp.Bottom - sbr.Top} dock={tp.Dock}");
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("EX " + ex);
                    }
                    System.IO.File.WriteAllText(args.Length > 6 ? args[6]
                        : @"D:\work\wxd.files.winforms\tools\restyle-shots\geom.txt", sb.ToString());
                    viewer.Close();
                };
                t.Start();
            };
        }
        Application.Run(viewer);
        return;
    }

    if (args.Length > 3 && args[3] == "dlg")
    {
        // 单独把某个对话框顶出来截图：args[4] = alert|err|confirm3|celledit|colmgr
        string which = args.Length > 4 ? args[4] : "alert";
        var owner = new Form
        {
            Text = "dlg-owner",
            ClientSize = new Size(760, 460),
            StartPosition = FormStartPosition.CenterScreen
        };
        owner.FormClosed += (_, _) => Application.Exit();
        owner.Shown += (s, e) =>
        {
            var t = new System.Windows.Forms.Timer { Interval = 400 };
            t.Tick += (_, _) =>
            {
                t.Stop();
                try { ShowDialog(asm, owner, which); }
                catch (Exception ex)
                {
                    System.IO.File.WriteAllText(
                        @"D:\work\wxd.files.winforms\tools\restyle-shots\dlg-err.txt", ex.ToString());
                }
            };
            t.Start();
        };
        Application.Run(owner);
        return;
    }

    var options = Activator.CreateInstance(optionsType, new object[] { file })!;
    var ctor = controlType.GetConstructor(new[] { optionsType })
                 ?? throw new MissingMethodException("TableBrowserControl(TableBrowserOptions)");
    var control = (Control)ctor.Invoke(new[] { options });

    var form = new Form
    {
        Text = $"gridshot skin={skin} dark={dark}",
        StartPosition = FormStartPosition.CenterScreen,
        WindowState = FormWindowState.Maximized,
        BackColor = Color.White
    };
    control.Dock = DockStyle.Fill;
    form.Controls.Add(control);
    Application.Run(form);
}

/// <summary>列选择类对话框共用的假表头表：3 行冻结（字段名 / 说明 / 类型）+ 2 行数据</summary>
static System.Data.DataTable FakeTable()
{
    var table = new System.Data.DataTable();
    string[] heads = { "字段名", "说明", "类型", "默认值", "范围", "示例", "备注", "索引" };
    string[] labels = { "id", "name", "level", "atk", "def", "speed", "desc", "icon" };
    for (int c = 0; c < heads.Length; c++) table.Columns.Add(labels[c], typeof(string));
    table.Rows.Add(heads);
    table.Rows.Add("主键", "名称", "等级", "攻击", "防御", "速度", "描述", "图标");
    table.Rows.Add("int", "string", "int", "int", "int", "float", "string", "string");
    table.Rows.Add("1001", "野猪王", "12", "88", "40", "1.5", "森林里的头目", "monster_1001");
    table.Rows.Add("1002", "岩蜥", "18", "120", "96", "0.9", "峡谷刷怪", "monster_1002");
    return table;
}

static List<(string display, string fullPath)> FakeEntries() => new()
{
    ("cfg_monster.csv", @"D:\策划配置\cfg_monster.csv"),
    ("cfg_item.csv", @"D:\策划配置\cfg_item.csv"),
    ("cfg_drop.xlsx", @"D:\策划配置\cfg_drop.xlsx"),
};

/// <summary>
/// 单独把某个对话框顶出来截图。gridshot 对产品程序集没有编译期引用，只能反射调静态入口。
/// 这些都是模态弹窗，调用会阻塞在这里直到用户关掉——抓完图直接 taskkill 即可。
/// </summary>
static void ShowDialog(Assembly asm, Form owner, string which)
{
    var alertForm = asm.GetType("wxd.files.winforms.AlertForm")!;
    var alertType = asm.GetType("wxd.files.winforms.AlertType")!;
    object Info = Enum.GetValues(alertType).GetValue(0)!;
    object Warn = Enum.GetValues(alertType).GetValue(1)!;

    switch (which)
    {
        case "alert":
            alertForm.GetMethod("Show", new[] { typeof(IWin32Window), typeof(string), typeof(string),
                alertType, typeof(bool), typeof(Exception) })!
                .Invoke(null, new object?[] { owner,
                    "已找到 128 处匹配，是否全部替换？\n替换后原内容不再保留。", "提示", Info, false, null });
            break;

        case "err":
            alertForm.GetMethod("ShowError", new[] { typeof(IWin32Window), typeof(string), typeof(string), typeof(Exception) })!
                .Invoke(null, new object?[] { owner,
                    "读取 D:\\策划配置\\cfg_monster.csv 失败：文件正由另一个进程占用。", "错误", null });
            break;

        case "confirm3":
            alertForm.GetMethod("ShowConfirmThreeButton", new[] { typeof(IWin32Window), typeof(string), typeof(string),
                typeof(string), typeof(string), typeof(string), alertType })!
                .Invoke(null, new object?[] { owner,
                    "当前表格有未保存的修改。", "保存更改", "保存", "放弃", "继续编辑", Warn });
            break;

        case "celledit":
        {
            var t = asm.GetType("wxd.files.winforms.CellEditDialog")!;
            using var dlg = (Form)Activator.CreateInstance(t, new object?[] {
                "第一行内容\n第二行内容\t带制表符", 4, 3, "cfg_monster.csv", "desc",
                null, false, "", false, null })!;
            dlg.ShowDialog(owner);
            break;
        }

        case "colmgr":
        {
            var t = asm.GetType("wxd.files.winforms.ColumnManagerDialog")!;
            var checks = new bool[8];
            Array.Fill(checks, true);
            using var dlg = (Form)Activator.CreateInstance(t, new object?[] {
                FakeTable(), checks, 3, 0, @"D:\策划配置", @"D:\策划配置\cfg_monster.csv",
                FakeEntries(), null, 0, 2 })!;
            dlg.ShowDialog(owner);
            break;
        }

        case "tcolsel":
        {
            var t = asm.GetType("wxd.files.winforms.TargetColumnSelectDialog")!;
            using var dlg = (Form)Activator.CreateInstance(t, new object?[] {
                FakeTable(), 3, 2, 0, "选择目标关联列" })!;
            dlg.ShowDialog(owner);
            break;
        }

        case "linksel":
        {
            var t = asm.GetType("wxd.files.winforms.LinkTableSelectorDialog")!;
            using var dlg = (Form)Activator.CreateInstance(t, new object[] { FakeEntries() })!;
            dlg.ShowDialog(owner);
            break;
        }

        case "cmdmgr":
        {
            var t = asm.GetType("wxd.files.winforms.CustomCommandManagerForm")!;
            using var dlg = (Form)Activator.CreateInstance(t)!;
            dlg.ShowDialog(owner);
            break;
        }

        case "tmpl":
        case "tmpls":
        {
            var t = asm.GetType("wxd.files.winforms.LinkTemplateDialog")!;
            bool selectMode = which == "tmpls";
            using var dlg = (Form)Activator.CreateInstance(t, new object?[] {
                selectMode, FakeEntries(), 3, 0, @"D:\策划配置" })!;
            dlg.ShowDialog(owner);
            break;
        }

        case "linksetup":
        {
            var t = asm.GetType("wxd.files.winforms.ColumnLinkSetupDialog")!;
            var cfgType = asm.GetType("wxd.files.winforms.ColumnLinkConfig")!;
            var cfg = Activator.CreateInstance(cfgType)!;
            cfgType.GetProperty("ValueExtractionEnabled")!.SetValue(cfg, true);
            cfgType.GetProperty("DisplayColumnHeaders")!.SetValue(cfg, "道具#数量#权重");
            using var dlg = (Form)Activator.CreateInstance(t, new object?[] {
                FakeEntries(), "cfg_monster.csv", 6, @"D:\策划配置", 3, cfg, 0 })!;
            dlg.ShowDialog(owner);
            break;
        }

        case "dupdetail":
        {
            // 只截图，owner 传 null：双击跳转才会用到它
            var t = asm.GetType("wxd.files.winforms.DuplicateDetailForm")!;
            var grpType = asm.GetType("wxd.files.winforms.TableBrowserControl+DuplicateGroup")!;
            var data = asm.GetType("wxd.files.winforms.FileTableData")!;
            var rowType = asm.GetType("wxd.files.winforms.TableRow")!;

            var fileData = Activator.CreateInstance(data)!;
            var rows = (System.Collections.IList)data.GetProperty("Rows")!.GetValue(fileData)!;
            for (int i = 1; i <= 30; i++)
            {
                var row = Activator.CreateInstance(rowType)!;
                // DisplayLineNumber 是只读的派生值，行号口径来自物理起始行
                rowType.GetProperty("PhysicalStartLine")!.SetValue(row, i);
                rows.Add(row);
            }

            var groups = (System.Collections.IList)Activator.CreateInstance(
                typeof(List<>).MakeGenericType(grpType))!;
            string[] vals = { "4002#1", "3001#2", "100#5" };
            for (int g = 0; g < vals.Length; g++)
            {
                var grp = Activator.CreateInstance(grpType)!;
                grpType.GetProperty("Value")!.SetValue(grp, vals[g]);
                int n = 4 - g;
                grpType.GetProperty("Count")!.SetValue(grp, n);
                var idxs = (System.Collections.IList)grpType.GetProperty("RowIndices")!.GetValue(grp)!;
                for (int k = 0; k < n; k++) idxs.Add(g * 5 + k);
                groups.Add(grp);
            }

            using var dlg = (Form)Activator.CreateInstance(t, new object?[] { groups, fileData, null })!;
            dlg.ShowDialog(owner);
            break;
        }

        case "help":
        {
            var t = asm.GetType("wxd.files.winforms.HelpForm")!;
            string body = "◆ 工具栏操作\n\n" +
                "  · 【设置】—— 展开/收起设置面板\n" +
                "    ├ 窗格数量：1~10 个窗格自由设置\n" +
                "    ├ 布局方式：横向、纵向、纵横交错流式三种模式\n" +
                "    └ 主题切换：皮肤与深浅色\n\n" +
                "◆ 键盘快捷键\n\n" +
                "  · F2 —— 重命名当前选中的文件/文件夹\n" +
                "  · F5 —— 刷新当前窗格的文件列表\n" +
                "  · Delete —— 删除选中的文件/文件夹\n" +
                "  · Ctrl+A —— 全选当前窗格中的所有文件";
            using var dlg = (Form)Activator.CreateInstance(t, new object?[] { "使用说明 - 测试", body })!;
            dlg.ShowDialog(owner);
            break;
        }

        case "update":
        {
            var t = asm.GetType("wxd.files.winforms.UpdateForm")!;
            using var dlg = (Form)Activator.CreateInstance(t)!;
            dlg.ShowDialog(owner);
            break;
        }

        case "updn":
        {
            var t = asm.GetType("wxd.files.winforms.UpdateNotifyForm")!;
            var resType = asm.GetType("wxd.files.winforms.CheckResult")!;
            var res = Activator.CreateInstance(resType)!;
            resType.GetProperty("HasUpdate")!.SetValue(res, true);
            resType.GetProperty("CurrentVersion")!.SetValue(res, "1.4.0");
            resType.GetProperty("RemoteVersion")!.SetValue(res, "1.4.2");
            resType.GetProperty("RemoteMd5")!.SetValue(res, "0123456789ABCDEF");
            resType.GetProperty("ReleaseNotes")!.SetValue(res,
                "v1.4.2\n  · 修复冻结列在重新绑定后丢失\n  · 表格管理器支持整行选中\n\n" +
                "v1.4.1\n  · 新增两套皮肤\n  · 深色模式下补齐对话框配色");
            using var dlg = (Form)Activator.CreateInstance(t, new object?[] { res })!;
            dlg.ShowDialog(owner);
            break;
        }

        case "pbar":
        {
            // ThemeProgressBar 是 internal，只能反射造；四种状态摆一屏，静态图也能看出色块
            var t = asm.GetType("wxd.files.winforms.ThemeProgressBar")!;
            using var dlg = new Form { Text = "进度条自检", ClientSize = new Size(520, 200),
                StartPosition = FormStartPosition.CenterParent, ShowIcon = false };
            int y = 20;
            foreach (var v in new[] { 0, 45, 100 })
            {
                var pb = (Control)Activator.CreateInstance(t)!;
                pb.SetBounds(20, y, 460, 20);
                t.GetProperty("Value")!.SetValue(pb, v);
                dlg.Controls.Add(pb);
                y += 34;
            }
            var mq = (Control)Activator.CreateInstance(t)!;
            mq.SetBounds(20, y, 460, 20);
            t.GetProperty("Marquee")!.SetValue(mq, true);
            dlg.Controls.Add(mq);
            dlg.ShowDialog(owner);
            break;
        }

        case "note":
        {
            // 只读地打开笔记窗体：NoteStore.GetAll 不写盘，截图不会污染用户数据
            var t = asm.GetType("wxd.files.winforms.NoteForm")!;
            using var dlg = (Form)Activator.CreateInstance(t)!;
            // GRIDSHOT_FLIP=1：开窗后运行时切一次深浅色，验「已经打开的笔记」会不会把正文颜色留在旧主题上
            if (Environment.GetEnvironmentVariable("GRIDSHOT_FLIP") == "1")
            {
                var themeT = asm.GetType("wxd.files.winforms.Theme")!;
                var skinE = asm.GetType("wxd.files.winforms.UiSkin")!;
                dlg.Shown += (_, _) =>
                {
                    var tmr = new System.Windows.Forms.Timer { Interval = 2500 };
                    tmr.Tick += (_, _) =>
                    {
                        tmr.Stop(); tmr.Dispose();
                        bool curDark = (bool)themeT.GetProperty("Dark")!.GetValue(null)!;
                        int curSkin = (int)themeT.GetProperty("Skin")!.GetValue(null)!;
                        themeT.GetMethod("Set")!.Invoke(null,
                            new object[] { Enum.GetValues(skinE).GetValue(curSkin), !curDark });
                    };
                    tmr.Start();
                };
            }
            dlg.ShowDialog(owner);
            break;
        }

        case "noteitem":
        {
            // 列表项单独摆三种状态：普通 / 悬停 / 选中
            var itemType = asm.GetType("wxd.files.winforms.NoteItem")!;
            var ctrlType = asm.GetType("wxd.files.winforms.NoteListItemControl")!;
            using var dlg = new Form { Text = "笔记列表项", ClientSize = new Size(420, 170),
                StartPosition = FormStartPosition.CenterParent, ShowIcon = false };
            for (int i = 0; i < 3; i++)
            {
                var note = Activator.CreateInstance(itemType)!;
                itemType.GetProperty("Title")!.SetValue(note, i == 2 ? "(无标题)" : $"配置表冻结行的坑{i + 1}");
                itemType.GetProperty("UpdatedAt")!.SetValue(note, new DateTime(2026, 11, 8, 13, 5, 30 + i));
                var c = (Control)Activator.CreateInstance(ctrlType, new object?[] { note })!;
                dlg.Controls.Add(c);
                if (i == 1) c.GetType().GetMethod("SetSelected")!.Invoke(c, new object[] { true });
            }
            dlg.ShowDialog(owner);
            break;
        }

        case "conflict":
        {
            // 冲突解决是 UserControl，宿主进普通窗体；两份 FakeTable 各造一处新增/修改
            var t = asm.GetType("wxd.files.winforms.ConflictResolveControl")!;
            var vcsType = asm.GetType("wxd.files.winforms.VcsType")!;
            var remote = FakeTable();
            var local = FakeTable();
            remote.Rows.Add("1003", "岩蜥王", "25", "200", "150", "0.7", "只有远程有的新增行", "monster_1003");
            local.Rows.Add("1004", "霜狼", "9", "60", "30", "2.2", "只有本地有的新增行", "monster_1004");
            remote.Rows[3]["atk"] = "120";
            local.Rows[4]["def"] = "88";

            var ctor = t.GetConstructors().First(c => c.GetParameters().Length == 12);
            var ctl = (Control)ctor.Invoke(new object?[] {
                @"D:\策划配置\cfg_monster.csv", Enum.GetValues(vcsType).GetValue(2),
                remote, local, 3, 1, @"D:\策划配置", 0, 2, null, null, null })!;

            using var dlg = new Form { Text = "冲突解决自检", ClientSize = new Size(1300, 860),
                StartPosition = FormStartPosition.CenterParent, ShowIcon = false };
            ctl.Dock = DockStyle.Fill;
            dlg.Controls.Add(ctl);
            var load = (Task)t.GetMethod("LoadComparisonAsync")!.Invoke(ctl, Array.Empty<object>())!;
            int deadline = Environment.TickCount + 15000;
            while (!load.IsCompleted && Environment.TickCount < deadline)
                Application.DoEvents();
            dlg.ShowDialog(owner);
            break;
        }

        case "vcmp":
        {
            var t = asm.GetType("wxd.files.winforms.VersionCompareForm")!;
            var vcsType = asm.GetType("wxd.files.winforms.VcsType")!;
            var ctor = t.GetConstructors().First(c => c.GetParameters().Length == 9);
            using var dlg = (Form)ctor.Invoke(new object?[] {
                @"D:\策划配置\cfg_monster.csv", Enum.GetValues(vcsType).GetValue(2),
                3, 1, @"D:\策划配置", null, 0, -1, null })!;
            dlg.ShowDialog(owner);
            break;
        }

        default:
            throw new ArgumentException("unknown dialog: " + which);
    }
}

/// <summary>顶栏自检：导出每个子控件的尺寸/上边距/相对顶栏的纵坐标，
/// 用来判断动态命令按钮和固定按钮是否在同一基线（rebuilt=1 表示模拟"命令管理"关闭后的重建路径）
static void BarGeom(Type mainType, Form main, string log, bool hold)
{
    var flags = BindingFlags.Instance | BindingFlags.NonPublic;
    var report = new List<string>();

    void Dump(string stage)
    {
        var flow = (Control)mainType.GetField("_flowPanel", flags)!.GetValue(main)!;
        var bar = (Control)mainType.GetField("_pnlTopBar", flags)!.GetValue(main)!;
        report.Add($"[{stage}] barH={bar.Height} flow={flow.ClientSize.Width}x{flow.ClientSize.Height}");
        foreach (Control c in flow.Controls)
            report.Add($"  {c.GetType().Name,-14} {c.Text,-10} bounds=({c.Left},{c.Top},{c.Width}x{c.Height}) " +
                       $"margin=({c.Margin.Left},{c.Margin.Top},{c.Margin.Right},{c.Margin.Bottom}) " +
                       $"absTop={c.Top - flow.Top}");

        // 右侧区顺带导出屏幕坐标，方便用真实鼠标点「命令管理」复现关闭后的状态
        var right = (Control)mainType.GetField("_pnlTopRight", flags)!.GetValue(main)!;
        foreach (Control c in right.Controls)
        {
            var r = c.RectangleToScreen(c.Bounds);
            report.Add($"  R {c.Text,-10} screen=({r.Left + r.Width / 2},{r.Top + r.Height / 2}) size={r.Width}x{r.Height}");
        }
    }

    main.Shown += (s, e) =>
    {
        try
        {
            Dump("startup");
            mainType.GetMethod("RebuildCmdActionButtons", flags)!.Invoke(main, Array.Empty<object>());
            Dump("rebuild-only");
            // hold=true 停在"命令管理关闭后"的状态，让截图抓到真实像素
            if (!hold)
                mainType.GetMethod("LayoutTopBar", flags)!.Invoke(main, Array.Empty<object>());
            Dump("rebuild+layout");
        }
        catch (Exception ex)
        {
            report.Add("EX:" + (ex.InnerException ?? ex).GetType().Name + " " + (ex.InnerException ?? ex).Message);
        }
        File.WriteAllLines(log, report);
    };
}

/// <summary>顶栏分段菜单自检：开菜单、点项、看勾选是否跟着走</summary>
static void SmokeTest(Type mainType, Form main, string log)
{
    var flags = BindingFlags.Instance | BindingFlags.NonPublic;
    var report = new List<string>();

    void Fire(int seg) => mainType.GetMethod("OnSegmentClicked", flags)!
        .Invoke(main, new object[] { seg });
    ContextMenuStrip Menu(string field) =>
        (ContextMenuStrip)mainType.GetField(field, flags)!.GetValue(main)!;
    void ClickItem(ContextMenuStrip menu, string text) =>
        menu.Items.OfType<ToolStripMenuItem>().First(i => i.Text == text).PerformClick();
    int Panes() => ((Array)mainType.GetField("_panels", flags)!.GetValue(main)!).Length;
    string Checks(ContextMenuStrip m) => string.Join("|",
        m.Items.OfType<ToolStripMenuItem>().Where(i => i.Checked).Select(i => i.Text));
    string Tags(ContextMenuStrip m) => string.Join(";",
        m.Items.OfType<ToolStripMenuItem>().Select(i => $"{i.Text}={i.Tag}"));

    main.Shown += (s, e) =>
    {
        try
        {
            foreach (int seg in new[] { 0, 1, 2 })
            {
                Fire(seg);
                report.Add($"seg{seg} visible=" + (seg switch
                {
                    0 => Menu("_menuPane"),
                    1 => Menu("_menuLayout"),
                    _ => Menu("_menuTheme")
                }).Visible);
            }
            Menu("_menuPane").Close();
            Menu("_menuLayout").Close();
            Menu("_menuTheme").Close();

            report.Add("start panes=" + Panes() + " paneChecks=" + Checks(Menu("_menuPane"))
                + " layoutChecks=" + Checks(Menu("_menuLayout")) + " themeChecks=" + Checks(Menu("_menuTheme")));
            report.Add("paneTags=" + Tags(Menu("_menuPane")));

            ClickItem(Menu("_menuPane"), "3 个窗格");
            report.Add("afterPane panes=" + Panes() + " checks=" + Checks(Menu("_menuPane")));

            ClickItem(Menu("_menuLayout"), "横向布局");
            report.Add("afterLayout checks=" + Checks(Menu("_menuLayout")));

            ClickItem(Menu("_menuTheme"), "深色");
            report.Add("afterTheme checks=" + Checks(Menu("_menuTheme")));

            ClickItem(Menu("_menuPane"), "显示隐藏文件");
            report.Add("afterHidden checks=" + Checks(Menu("_menuPane")));
        }
        catch (Exception ex)
        {
            report.Add("FAIL " + ex);
        }
        finally
        {
            File.WriteAllLines(log, report);
            main.Close();
        }
    };
}

// 递归遍历控件树（含 DockContent 之类嵌套容器）
static IEnumerable<Control> Walk(Control root)
{
    foreach (Control c in root.Controls)
    {
        yield return c;
        foreach (var child in Walk(c))
            yield return child;
    }
}

static class Native
{
    public const int LVM_GETHEADER = 0x100D;

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool InvalidateRect(IntPtr h, IntPtr rect, bool erase);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool UpdateWindow(IntPtr h);

    [System.Runtime.InteropServices.DllImport("uxtheme.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    public static extern int SetWindowTheme(IntPtr h, string? subAppName, string? subColorSet);
}
