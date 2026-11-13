using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    public partial class Form1 : Form
    {
        // ==================== SetWindowPos P/Invoke（静默控制 Z-order，避免 TopMost 闪烁） ====================
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        private static readonly IntPtr HWND_TOP = IntPtr.Zero;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private ThemeSegmented? _segSettings;                 // 窗格 / 布局 / 主题 分段入口
        private ContextMenuStrip? _menuPane;                  // 窗格菜单
        private ContextMenuStrip? _menuLayout;                // 布局菜单
        private ContextMenuStrip? _menuTheme;                 // 主题与皮肤菜单
        private Button? _btnOpenCsvReader;
        private Button? _btnNotes;                          // 笔记按钮
        private NoteForm? _noteForm;                        // 笔记窗体（单例）
        private Panel _pnlContainer = new();
        private TabbedFileBrowserPanel[] _panels = Array.Empty<TabbedFileBrowserPanel>();
        private Panel? _pnlTopBar;
        private Panel? _pnlTopRight;                        // 顶栏右侧固定区
        private Button? _btnCmdManager;                      // 命令管理按钮
        private Button? _btnSaveState;                        // 保存状态按钮
        private FlowLayoutPanel? _flowPanel;                  // 顶部流式布局面板
        private readonly List<Button> _cmdActionButtons = new(); // 动态生成的命令按钮

        private int _paneCount = 7;                           // 当前窗格数量
        private int _layoutIndex = 2;                         // 当前布局方式序号

        // ==================== 启动Loading相关（独立STA线程方案） ====================
        private StartupLoadingForm? _startupLoadingForm;      // 独立线程运行的Loading窗体
        private Thread? _loadingThread;                       // Loading所在线程

        private enum LayoutMode { Horizontal, Vertical, Flow }
        private enum ThemeMode { Dark, Light }

        private ThemeMode _currentTheme = ThemeMode.Light;

        public Form1()
        {
            InitializeComponent();
            // 在标题栏追加程序集版本号
            var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            if (ver != null)
                Text = $"{Text} v{ver}";

            // 先按存档恢复窗格数量/布局/明暗，再激活令牌，后面控件创建期取色和菜单勾选才是最终值
            _paneCount = Math.Clamp(CsvSettings.PaneCount, 1, 10);
            _layoutIndex = Math.Clamp(CsvSettings.LayoutIndex, 0, 2);
            _currentTheme = (ThemeMode)CsvSettings.ThemeMode;
            Theme.Set((UiSkin)CsvSettings.ThemeSkin, _currentTheme == ThemeMode.Dark);

            CreateTopBar();
            InitControlPanel();

            // 应用主题（面板创建后再应用到面板）
            ApplyTheme();

            // 显示Loading遮罩（独立STA线程，确保动画流畅）
            ShowStartupLoading();

            // 首次可见时异步加载所有面板（此时 ClientRectangle 已是最终尺寸）
            VisibleChanged += OnFirstVisible;
        }

        /// <summary>仅在首次变为可见时触发一次异步加载</summary>
        private void OnFirstVisible(object? sender, EventArgs e)
        {
            if (Visible)
            {
                VisibleChanged -= OnFirstVisible;
                // 用 BeginInvoke 确保在当前消息处理完后执行，此时布局已完成
                BeginInvoke(AsyncInitializePanels);
            }
        }

        /// <summary>启动独立STA线程的Loading窗体（仿照VersionCompareForm方案）</summary>
        private void ShowStartupLoading()
        {
            if (_startupLoadingForm != null || _loadingThread != null) return;

            var ownerHandle = this.Handle;
            var ownerLocation = this.Location;
            var ownerSize = this.Size;

            // ★ 信号量：Loading 窗体 Shown 后通知主线程
            var readyEvent = new ManualResetEvent(false);

            _loadingThread = new Thread(() =>
            {
                _startupLoadingForm = new StartupLoadingForm();
                _startupLoadingForm.Load += (s, e) =>
                {
                    _startupLoadingForm.Location = new Point(
                        ownerLocation.X + (ownerSize.Width - _startupLoadingForm.Width) / 2,
                        ownerLocation.Y + (ownerSize.Height - _startupLoadingForm.Height) / 2);
                };
                _startupLoadingForm.Shown += (s, e) =>
                {
                    // ★ 窗体完全显示后通知主线程
                    readyEvent.Set();
                };
                _startupLoadingForm.FormClosed += (s, e) =>
                {
                    _startupLoadingForm = null;
                };
                Application.Run(_startupLoadingForm);
            });
            _loadingThread.SetApartmentState(ApartmentState.STA);
            _loadingThread.IsBackground = true;
            _loadingThread.Start();

            // ★ 等待 Loading 窗体完全显示（最多等 3 秒，防止卡死）
            readyEvent.WaitOne(3000);

            // 禁用主窗体，阻止用户交互
            Enabled = false;
        }

        /// <summary>隐藏Loading窗体（跨线程关闭，带淡出效果）</summary>
        private async Task HideStartupLoadingAsync()
        {
            // 恢复主窗体交互
            Enabled = true;

            // ★★★ 先显示面板内容，再关闭 Loading ★★★
            SuspendLayout();
            foreach (var panel in _panels)
            {
                if (!panel.IsDisposed)
                    panel.Visible = true;
            }
            ResumeLayout(false);
            Refresh();

            // ★ 强制处理 WM_PAINT，确保面板完全渲染
            Application.DoEvents();

            // ★★★ 关键：静默地将 Form1 置于 TopMost 层（不激活，不闪烁） ★★★
            // Loading 关闭后 Form1 就在最前面，不会被 VS 等其他窗口挡住
            SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

            // 关闭 LoadingForm（跨线程）
            _startupLoadingForm?.BeginInvoke(new Action(() =>
            {
                _startupLoadingForm?.Close();
            }));
            _loadingThread?.Join(1000);
            _startupLoadingForm?.Dispose();
            _startupLoadingForm = null;
            _loadingThread = null;

            // ★ Loading 关闭后，用 BeginInvoke 延迟取消 TopMost 并温和激活
            // BeginInvoke 确保在 Loading 线程完全退出后再执行
            await Task.Delay(50);

            BeginInvoke(new Action(() =>
            {
                // 1. 取消 TopMost，回到普通 Z-order（此时 Form1 在 Loading 之前的位置，即最前面）
                SetWindowPos(Handle, HWND_NOTOPMOST, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

                // 2. 温和激活（不带 SWP_NOACTIVATE，会激活窗口但不改变 Z-order）
                SetWindowPos(Handle, IntPtr.Zero, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE);
            }));
        }

        /// <summary>异步初始化所有面板：先创建空面板做布局，再逐个异步加载数据（后台线程执行）</summary>
        private async void AsyncInitializePanels()
        {
            // ★ 启动时加载 AppStateStore
            AppStateStore.Load();

            int panelCount = _paneCount;

            try
            {
                // ===== 第一步：创建轻量级面板（只建UI，不加载数据）=====
                CreatePanelsLightweight(panelCount);
                ApplyLayout(GetCurrentLayout());

                // ★ ShowStartupLoading 已通过 ManualResetEvent 确保 Loading 窗体完全显示，
                // 不再需要 Task.Delay 猜测等待时间

                // ===== 第二步：逐个异步初始化面板，等待每个面板的文件列表加载完成 =====
                // InitializeDeferredAsync 内部 NavigateToAsync → await LoadFileListAsync，
                // LoadFileListAsync 在后台线程读取文件，不会长时间阻塞 UI 线程。
                for (int i = 0; i < _panels.Length; i++)
                {
                    // ★ 跨线程更新进度（通过 BeginInvoke 发送到 LoadingForm 的独立消息泵）
                    int percent = (i * 100) / _panels.Length;
                    _startupLoadingForm?.BeginInvoke(new Action(() =>
                    {
                        _startupLoadingForm?.UpdateProgress(percent,
                            $"正在加载窗格 {i + 1}/{_panels.Length}...");
                    }));

                    // ★ 等待文件列表加载完成后再进入下一个面板
                    await _panels[i].InitializeDeferredAsync();
                }

                // 更新进度到 100%
                _startupLoadingForm?.BeginInvoke(new Action(() =>
                {
                    _startupLoadingForm?.UpdateProgress(100, "初始化完成！");
                }));

                await Task.Delay(300); // 让用户看到 100% 状态

                // 隐藏 Loading 遮罩
                await HideStartupLoadingAsync();

                // 界面完全就绪后，延迟 3 秒再触发更新检查
                var updateTimer = new System.Windows.Forms.Timer { Interval = 3000 };
                System.Windows.Forms.Timer timerRef = updateTimer;
                updateTimer.Tick += async (s, e) =>
                {
                    timerRef.Stop();
                    timerRef.Dispose();
                    try { await UpdateHelper.CheckSilentAsync(); }
                    catch { /* 静默 */ }
                };
                timerRef.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AsyncInitializePanels error: {ex}");
                // 即使出错也要隐藏loading，并显示面板
                foreach (var panel in _panels)
                    if (!panel.IsDisposed) panel.Visible = true;
                await HideStartupLoadingAsync();
            }
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            // 
            // Form1
            // 
            ClientSize = new Size(1300, 750);
            Font = Theme.Body;
            Icon = WindowIcons.App;
            MinimumSize = new Size(600, 450);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "多窗格文件夹管理器";
            WindowState = FormWindowState.Maximized;
            ResumeLayout(false);
        }

        private void CreateTopBar()
        {
            SuspendLayout();

            // 顶部工具栏面板：高度与内边距跟其它工具条同源，皮肤切换时由 LayoutTopBar 重算
            _pnlTopBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = ThemeMetrics.BarHeight,
                Padding = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Margin, 0),
                BackColor = Theme.P.Chrome
            };

            // 容器面板
            _pnlContainer.Dock = DockStyle.Fill;
            _pnlContainer.Padding = new Padding(ThemeMetrics.Gap);
            _pnlContainer.BackColor = BackColor;

            Controls.Add(_pnlContainer);
            Controls.Add(_pnlTopBar);

            ResumeLayout(false);
        }

        private void InitControlPanel()
        {
            if (_pnlTopBar == null) return;

            // 流式布局面板：左侧按「标题 → 设置入口 → 主操作」从左到右排开
            // 顶栏高度固定，控件不再靠容器内边距下移，而是各自的 Margin.Top 居中
            _flowPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = false,
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };

            var lblTitle = new Label
            {
                Text = "多窗格文件管理",
                AutoSize = true,
                Font = Theme.Title,
                ForeColor = Theme.P.Text,
                Tag = "title",
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
            };

            // 窗格 / 布局 / 主题：三个低频入口合成一颗分段按钮，各点各的菜单
            _segSettings = new ThemeSegmented
            {
                Items = new[] { "窗格", "布局", "主题" },
                ActiveIndex = 0,
                Margin = new Padding(0, 0, ThemeMetrics.Group, 0)
            };
            _segSettings.SegmentClicked += OnSegmentClicked;

            _menuPane = BuildPaneMenu();
            _menuLayout = BuildLayoutMenu();
            _menuTheme = BuildThemeMenu();

            // 打开CSV阅读器按钮
            _btnOpenCsvReader = new ThemeButton
            {
                Text = "阅读器",
                Font = Theme.Body,
                Primary = true,
                Margin = new Padding(0, 0, ThemeMetrics.Gap, 0)
            };
            _btnOpenCsvReader.Click += OnOpenCsvReaderClick;

            // 打开大文件阅读器按钮
            var btnLargeFileViewer = new ThemeButton
            {
                Text = "日志分析",
                Font = Theme.Body,
                Margin = new Padding(0, 0, ThemeMetrics.Gap, 0)
            };
            btnLargeFileViewer.Click += OnOpenLargeFileViewerClick;

            // 笔记按钮
            _btnNotes = new ThemeButton
            {
                Text = "笔记",
                Font = Theme.Body,
                Margin = new Padding(0, 0, ThemeMetrics.Gap, 0)
            };
            _btnNotes.Click += OnOpenNotesClick;

            _flowPanel.Controls.AddRange(new Control[]
            {
                lblTitle, _segSettings, _btnOpenCsvReader, btnLargeFileViewer, _btnNotes
            });

            // 右侧固定区：管理入口 + 关于，从右往左排，宽度随内容自适应
            _pnlTopRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };

            var btnHelp = new ThemeButton
            {
                Text = "关于",
                Font = Theme.Body,
                Ghost = true,
                Margin = new Padding(ThemeMetrics.Gap, 0, 0, 0)
            };
            btnHelp.Click += OnHelpClick;

            _btnCmdManager = new ThemeButton
            {
                Text = "命令管理",
                Font = Theme.Body,
                Ghost = true,
                Margin = new Padding(ThemeMetrics.Gap, 0, 0, 0)
            };
            _btnCmdManager.Click += OnCmdManagerClick;

            _btnSaveState = new ThemeButton
            {
                Text = "保存状态",
                Font = Theme.Body,
                Ghost = true,
                Margin = new Padding(ThemeMetrics.Gap, 0, 0, 0)
            };
            _btnSaveState.Click += (s, e) => SaveAllState();

            _pnlTopRight.Controls.AddRange(new Control[] { btnHelp, _btnCmdManager, _btnSaveState });

            // 停靠顺序：后加入的先停靠，所以底部分隔线要最后加才能横跨整条顶栏
            _pnlTopBar.Controls.Add(_flowPanel);
            _pnlTopBar.Controls.Add(_pnlTopRight);
            _pnlTopBar.Controls.Add(new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 1,
                BackColor = Theme.P.Border
            });

            // 重建动态命令按钮（内部会顺带做一次顶栏居中）
            RebuildCmdActionButtons();
        }

        /// <summary>
        /// 顶栏的尺寸全部从 ThemeMetrics 推导：皮肤切换会改行高，
        /// 所以按钮按文字量宽、控件按顶栏高度垂直居中，都要能建一次也能重算。
        /// </summary>
        private void LayoutTopBar()
        {
            if (_pnlTopBar == null) return;

            int barH = ThemeMetrics.BarHeight;
            _pnlTopBar.Height = barH;

            if (_flowPanel != null) LayoutTopBarChildren(_flowPanel, ThemeMetrics.RowH, barH);
            if (_pnlTopRight != null) LayoutTopBarChildren(_pnlTopRight, ThemeMetrics.RowH, barH);
        }

        private static void LayoutTopBarChildren(Control host, int rowH, int barH)
        {
            foreach (Control c in host.Controls)
            {
                int h = c switch
                {
                    ThemeSegmented => rowH + 4,
                    Label { AutoSize: true } => c.GetPreferredSize(Size.Empty).Height,
                    _ => rowH
                };

                if (c is Button btn)
                    btn.Size = new Size(ThemeButton.FitWidth(btn.Text), h);
                else if (c is ThemeSegmented)
                    c.Size = c.GetPreferredSize(Size.Empty);
                else if (c is not Label)
                    c.Height = h;

                // 流式面板只按上边距对齐，所以每个控件自己补居中偏移
                var m = c.Margin;
                c.Margin = new Padding(m.Left, Math.Max(0, (barH - c.Height) / 2), m.Right, 0);
            }
        }

        // ==================== 打开CSV阅读器（先选目录） ====================
        private void OnOpenCsvReaderClick(object? sender, EventArgs e)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "选择包含CSV/Excel文件的目录",
                ShowNewFolderButton = false
            };

            var lastDir = CsvSettings.CsvReaderLastDirectory;
            if (!string.IsNullOrEmpty(lastDir) && Directory.Exists(lastDir))
                dialog.SelectedPath = lastDir;

            if (dialog.ShowDialog() != DialogResult.OK) return;

            // 记忆本次选择的目录
            CsvSettings.CsvReaderLastDirectory = dialog.SelectedPath;

            var csvForm = new CsvReaderForm(dialog.SelectedPath);
            csvForm.Show();
        }

        // ==================== 打开大文件阅读器（先选目录） ====================
        private void OnOpenLargeFileViewerClick(object? sender, EventArgs e)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "选择包含文本文件的目录",
                ShowNewFolderButton = false
            };

            var lastDir = CsvSettings.LargeFileViewerLastDirectory;
            if (!string.IsNullOrEmpty(lastDir) && Directory.Exists(lastDir))
                dialog.SelectedPath = lastDir;

            if (dialog.ShowDialog() != DialogResult.OK) return;

            CsvSettings.LargeFileViewerLastDirectory = dialog.SelectedPath;

            var f = new LargeFileViewer(dialog.SelectedPath);
            f.Show();
        }

        // ==================== 打开笔记窗体（单例） ====================
        private void OnOpenNotesClick(object? sender, EventArgs e)
        {
            if (_noteForm != null && !_noteForm.IsDisposed)
            {
                if (_noteForm.WindowState == FormWindowState.Minimized)
                    _noteForm.WindowState = FormWindowState.Normal;
                _noteForm.Activate();
                return;
            }

            _noteForm = new NoteForm();
            _noteForm.Owner = null;
            _noteForm.FormClosed += (s, e2) => _noteForm = null;
            _noteForm.Show();
        }

        // ==================== 关于 / 更新 ====================
        private void OnHelpClick(object? sender, EventArgs e)
        {
            UpdateForm.ShowUpdateForm(this);
        }

        // ==================== 分段入口：弹出对应菜单 ====================
        private void OnSegmentClicked(int index)
        {
            if (_segSettings == null) return;
            _segSettings.ActiveIndex = index;

            var menu = index switch { 0 => _menuPane, 1 => _menuLayout, _ => _menuTheme };
            if (menu == null) return;

            var rect = _segSettings.SegmentRect(index);
            menu.Show(_segSettings, new Point(rect.Left, rect.Bottom + 2));
        }

        private static ContextMenuStrip NewMenu() => new()
        {
            Font = Theme.List,
            ShowImageMargin = false,
            ShowCheckMargin = true
        };

        /// <summary>往菜单里加一项，Tag 记下它属于哪一组、代表什么值，勾选时按组比对</summary>
        private static ToolStripMenuItem AddMenuEntry(ContextMenuStrip menu, string text,
            string group, object value, Action onPick)
        {
            var item = new ToolStripMenuItem(text) { Tag = ValueTuple.Create(group, value) };
            item.Click += (s, e) => onPick();
            menu.Items.Add(item);
            return item;
        }

        /// <summary>同一组里只勾当前值，其余取消；别的组不受影响</summary>
        private static void CheckGroup(ContextMenuStrip menu, string group, object value)
        {
            foreach (ToolStripItem it in menu.Items)
            {
                if (it is not ToolStripMenuItem mi) continue;
                if (mi.Tag is not ValueTuple<string, object> tag || tag.Item1 != group) continue;
                mi.Checked = Equals(tag.Item2, value);
            }
        }

        private ContextMenuStrip BuildPaneMenu()
        {
            var menu = NewMenu();
            for (int i = 1; i <= 10; i++)
            {
                int count = i;
                AddMenuEntry(menu, $"{count} 个窗格", "pane", count, () => SetPaneCount(count));
            }
            menu.Items.Add(new ToolStripSeparator());
            AddMenuEntry(menu, "显示隐藏文件", "hidden", true, () => SetShowHidden(!CsvSettings.ShowHiddenFiles));
            CheckGroup(menu, "pane", _paneCount);
            CheckGroup(menu, "hidden", CsvSettings.ShowHiddenFiles);
            return menu;
        }

        private ContextMenuStrip BuildLayoutMenu()
        {
            var menu = NewMenu();
            AddMenuEntry(menu, "横向布局", "layout", 0, () => SetLayoutIndex(0));
            AddMenuEntry(menu, "纵向布局", "layout", 1, () => SetLayoutIndex(1));
            AddMenuEntry(menu, "纵横交错流式布局", "layout", 2, () => SetLayoutIndex(2));
            CheckGroup(menu, "layout", _layoutIndex);
            return menu;
        }

        private ContextMenuStrip BuildThemeMenu()
        {
            var menu = NewMenu();
            AddMenuEntry(menu, "深色", "mode", 0, () => SetThemeMode(ThemeMode.Dark));
            AddMenuEntry(menu, "浅色", "mode", 1, () => SetThemeMode(ThemeMode.Light));
            menu.Items.Add(new ToolStripSeparator());
            AddMenuEntry(menu, "天蓝提质", "skin", 0, () => SetSkin(0));
            AddMenuEntry(menu, "密集工具", "skin", 1, () => SetSkin(1));
            AddMenuEntry(menu, "北欧蓝灰", "skin", 2, () => SetSkin(2));
            AddMenuEntry(menu, "暖沙护眼", "skin", 3, () => SetSkin(3));
            AddMenuEntry(menu, "复古暖棕", "skin", 4, () => SetSkin(4));
            AddMenuEntry(menu, "德古拉紫", "skin", 5, () => SetSkin(5));
            AddMenuEntry(menu, "薄荷绿", "skin", 6, () => SetSkin(6));
            AddMenuEntry(menu, "淑女粉", "skin", 7, () => SetSkin(7));
            CheckGroup(menu, "mode", (int)_currentTheme);
            CheckGroup(menu, "skin", CsvSettings.ThemeSkin);
            return menu;
        }

        // ==================== 窗格数量变更 ====================
        private void SetPaneCount(int count)
        {
            _paneCount = count;
            CheckGroup(_menuPane!, "pane", count);
            CreatePanels(count);
            ApplyLayout(GetCurrentLayout());
            SaveGlobalState();
        }

        // ==================== 布局方式变更 ====================
        private void SetLayoutIndex(int index)
        {
            _layoutIndex = index;
            CheckGroup(_menuLayout!, "layout", index);
            ApplyLayout(GetCurrentLayout());
            SaveGlobalState();
        }

        // ==================== 明暗切换 ====================
        private void SetThemeMode(ThemeMode mode)
        {
            _currentTheme = mode;
            CheckGroup(_menuTheme!, "mode", (int)mode);
            ApplyTheme();
            SaveGlobalState();
        }

        // ==================== 皮肤切换 ====================
        private void SetSkin(int skin)
        {
            CsvSettings.ThemeSkin = skin;
            CheckGroup(_menuTheme!, "skin", skin);
            ApplyTheme();
        }

        // ==================== 显示隐藏文件 ====================
        private void SetShowHidden(bool show)
        {
            CsvSettings.ShowHiddenFiles = show;
            CheckGroup(_menuPane!, "hidden", show);
            foreach (var panel in _panels)
            {
                if (!panel.IsDisposed)
                    panel.RefreshCurrentDirectory();
            }
        }

        /// <summary>实时保存全局状态（窗格数量+布局模式+主题）</summary>
        private void SaveGlobalState()
        {
            CsvSettings.PaneCount = _panels.Length;
            CsvSettings.LayoutIndex = _layoutIndex;
            CsvSettings.ThemeMode = (int)_currentTheme;
        }

        /// <summary>保存所有状态（包含每个面板的页签状态）</summary>
        private void SaveAllState()
        {
            // 强制保存每个面板的所有页签状态
            for (int i = 0; i < _panels.Length; i++)
            {
                _panels[i].ForceSaveState();
            }

            // 持久化到文件
            AppStateStore.Save();

            // 保存全局状态
            SaveGlobalState();
        }


        /// <summary>应用主题到整个窗体</summary>
        private void ApplyTheme()
        {
            var isDark = _currentTheme == ThemeMode.Dark;

            // 先刷新全局令牌，后面所有控件的取色都走 Theme.P
            Theme.Set((UiSkin)CsvSettings.ThemeSkin, isDark);

            // 窗体背景
            Font = Theme.Body;
            BackColor = Theme.P.Window;
            _pnlContainer.BackColor = BackColor;

            // 顶部工具栏
            if (_pnlTopBar != null)
            {
                _pnlTopBar.BackColor = Theme.P.Chrome;
                // 更新分隔线颜色
                foreach (Control c in _pnlTopBar.Controls)
                {
                    if (c is Panel p && p.Height == 1)
                        p.BackColor = Theme.P.Border;
                }
            }

            // 标签文字颜色 - 遍历_flowPanel中的控件
            var labelFg = Theme.P.Muted;
            var titleFg = Theme.P.Text;

            // 顶栏控件重新取色
            _segSettings?.RefreshTheme();
            foreach (Control c in _flowPanel.Controls)
            {
                if (c is Label lbl)
                {
                    // 标题靠 Tag 认出，别用字号猜：皮肤改了字号就会误判成普通标签
                    bool isTitle = Equals(lbl.Tag, "title");
                    lbl.Font = isTitle ? Theme.Title : Theme.Body;
                    lbl.ForeColor = isTitle ? titleFg : labelFg;
                }
                else if (c is ThemeButton tb)
                {
                    tb.Invalidate();
                }
            }

            if (_pnlTopRight != null)
            {
                foreach (Control c in _pnlTopRight.Controls)
                {
                    if (c is ThemeButton rb) rb.Invalidate();
                }
            }

            // 顶栏尺寸放到最后重算：标题字体刚换过，量出来的高度才是最终值
            LayoutTopBar();

            // 应用到所有子窗格
            foreach (var panel in _panels)
            {
                ApplyThemeToPanel(panel, isDark);
            }
        }

        /// <summary>应用主题到单个文件浏览面板</summary>
        private void ApplyThemeToPanel(TabbedFileBrowserPanel panel, bool isDark)
        {
            panel.ApplyTheme(isDark);
        }

        private LayoutMode GetCurrentLayout() => _layoutIndex switch
        {
            1 => LayoutMode.Vertical,
            2 => LayoutMode.Flow,
            _ => LayoutMode.Horizontal
        };

        // ==================== 创建/销毁窗格 ====================

        /// <summary>创建面板并立即导航到保存的目录（运行时切换数量时使用）</summary>
        private void CreatePanels(int count)
        {
            foreach (var p in _panels) p.Dispose();
            _pnlContainer.Controls.Clear();

            _panels = new TabbedFileBrowserPanel[count];
            for (int i = 0; i < count; i++)
            {
                _panels[i] = new TabbedFileBrowserPanel(i)
                {
                    Dock = DockStyle.None,
                    // 圆角卡片由面板自绘，方形边框会盖住圆弧
                    BorderStyle = BorderStyle.None
                };
                // 恢复页签配置（不跳过自动导航）
                _panels[i].RestoreTabs(skipAutoNavigate: false);
                // 应用当前主题到新面板
                _panels[i].ApplyTheme(_currentTheme == ThemeMode.Dark);
                _pnlContainer.Controls.Add(_panels[i]);
            }
        }

        /// <summary>创建轻量级面板（只建UI，不触发导航，用于启动时快速显示）</summary>
        private void CreatePanelsLightweight(int count)
        {
            foreach (var p in _panels) p.Dispose();
            _pnlContainer.Controls.Clear();

            _panels = new TabbedFileBrowserPanel[count];
            for (int i = 0; i < count; i++)
            {
                // 使用轻量构造模式：skipAutoNavigate=true 跳过构造函数中的自动导航
                _panels[i] = new TabbedFileBrowserPanel(i)
                {
                    Dock = DockStyle.None,
                    BorderStyle = BorderStyle.None,
                    Visible = false  // 初始隐藏，避免空白闪烁
                };
                // 恢复页签配置（跳过自动导航，后续由 InitializeDeferredAsync 加载）
                _panels[i].RestoreTabs(skipAutoNavigate: true);
                // 应用当前主题到新面板
                _panels[i].ApplyTheme(_currentTheme == ThemeMode.Dark);
                _pnlContainer.Controls.Add(_panels[i]);
            }
        }

        // ==================== 应用布局 ====================
        private void ApplyLayout(LayoutMode mode)
        {
            int count = _panels.Length;
            if (count == 0) return;

            // 容器可用区域（已减去 Padding）
            int availW = _pnlContainer.ClientRectangle.Width - _pnlContainer.Padding.Horizontal;
            int availH = _pnlContainer.ClientRectangle.Height - _pnlContainer.Padding.Top - _pnlContainer.Padding.Bottom;

            // 保护最小值，避免极端情况
            if (availW < 50) availW = 50;
            if (availH < 50) availH = 50;

            switch (mode)
            {
                case LayoutMode.Horizontal: DoHorizontal(count, availW, availH); break;
                case LayoutMode.Vertical: DoVertical(count, availW, availH); break;
                case LayoutMode.Flow: DoFlow(count, availW, availH); break;
            }
        }

        /// <summary>横向排列：所有窗格一行等分宽度</summary>
        private void DoHorizontal(int count, int w, int h)
        {
            int panelW = w / count;
            for (int i = 0; i < count; i++)
            {
                int x = _pnlContainer.Padding.Left + i * panelW;
                // 最后一个窗格撑满剩余空间（消除除法舍入误差）
                int pw = (i == count - 1) ? (w - i * panelW) : panelW;
                _panels[i].SetBounds(x, _pnlContainer.Padding.Top, pw, h);
            }
        }

        /// <summary>纵向排列：所有窗格一列等分高度</summary>
        private void DoVertical(int count, int w, int h)
        {
            int panelH = h / count;
            for (int i = 0; i < count; i++)
            {
                int y = _pnlContainer.Padding.Top + i * panelH;
                int ph = (i == count - 1) ? (h - i * panelH) : panelH;
                _panels[i].SetBounds(_pnlContainer.Padding.Left, y, w, ph);
            }
        }

        /// <summary>流式排列：3/5/7个时左边流式+右边一个占满高度，其他数量正常流式布局</summary>
        private void DoFlow(int count, int w, int h)
        {
            const int gap = 4; // 窗格之间的间隙

            // 3、5、7个窗格：特殊布局，左边流式，右边最后一个占满高度
            if (count == 3 || count == 5 || count == 7)
            {
                int leftCount = count - 1; // 左边窗格数量
                int leftCols = leftCount <= 2 ? 1 : 2; // 左边列数：2个以内1列，4个2列，6个2列
                int leftRows = (leftCount + leftCols - 1) / leftCols; // 左边行数（向上取整）

                int rightW = (w - gap) / 3; // 右边占1/3宽度
                int leftW = w - gap - rightW; // 左边占剩余宽度

                int cellW = (leftW - gap * (leftCols - 1)) / leftCols;
                int cellH = (h - gap * (leftRows - 1)) / leftRows;

                // 布局左边窗格
                for (int i = 0; i < leftCount; i++)
                {
                    int row = i / leftCols;
                    int col = i % leftCols;
                    bool isLastRow = (row == leftRows - 1);
                    bool isLastCol = (col == leftCols - 1);

                    int x = _pnlContainer.Padding.Left + col * (cellW + gap);
                    int y = _pnlContainer.Padding.Top + row * (cellH + gap);
                    int cw = isLastCol ? (leftW - col * (cellW + gap)) : cellW;
                    int ch = isLastRow ? (h - row * (cellH + gap)) : cellH;

                    _panels[i].SetBounds(x, y, cw, ch);
                }

                // 右边最后一个窗格占满高度
                _panels[leftCount].SetBounds(_pnlContainer.Padding.Left + leftW + gap, _pnlContainer.Padding.Top, rightW, h);
            }
            else
            {
                // 其他数量：正常流式网格布局
                int cols, rows;

                if (count <= 1)
                {
                    cols = 1;
                    rows = 1;
                }
                else if (count == 2)
                {
                    cols = 2;
                    rows = 1;
                }
                else if (count == 4)
                {
                    cols = 2;
                    rows = 2;
                }
                else if (count == 6)
                {
                    cols = 3;
                    rows = 2;
                }
                else if (count == 8)
                {
                    cols = 4;
                    rows = 2;
                }
                else if (count == 9)
                {
                    cols = 3;
                    rows = 3;
                }
                else // count == 10
                {
                    cols = 5;
                    rows = 2;
                }

                int cellW = (w - gap * (cols - 1)) / cols;
                int cellH = (h - gap * (rows - 1)) / rows;

                for (int i = 0; i < count; i++)
                {
                    int row = i / cols;
                    int col = i % cols;
                    bool isLastCol = (col == cols - 1);
                    bool isLastRow = (row == rows - 1);

                    int x = _pnlContainer.Padding.Left + col * (cellW + gap);
                    int y = _pnlContainer.Padding.Top + row * (cellH + gap);
                    int cw = isLastCol ? (w - col * (cellW + gap)) : cellW;
                    int ch = isLastRow ? (h - row * (cellH + gap)) : cellH;

                    _panels[i].SetBounds(x, y, cw, ch);
                }
            }
        }

        // ==================== 窗口大小改变时重新布局 ====================
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_panels.Length > 0 && _pnlContainer.ClientRectangle.Width > 10)
            {
                ApplyLayout(GetCurrentLayout());
            }
        }

        // ==================== 关闭时保存全局状态 ====================
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("OnFormClosing called");
            // 强制保存每个面板的全部状态（含目录、过滤条件、视图模式）
            for (int i = 0; i < _panels.Length; i++)
            {
                _panels[i].ForceSaveState();
            }

            // 持久化到文件
            AppStateStore.Save();

            CsvSettings.PaneCount = _panels.Length;
            CsvSettings.LayoutIndex = _layoutIndex;
            CsvSettings.ThemeMode = (int)_currentTheme;

            _menuPane?.Dispose();
            _menuLayout?.Dispose();
            _menuTheme?.Dispose();
            base.OnFormClosing(e);
        }

        // ==================== 打开命令管理面板 ====================
        private void OnCmdManagerClick(object? sender, EventArgs e)
        {
            using var form = new CustomCommandManagerForm();
            form.ShowDialog(this);
            // 关闭后刷新动态按钮
            RebuildCmdActionButtons();
        }

        // ==================== 重建动态命令按钮 ====================
        private void RebuildCmdActionButtons()
        {
            // 清除旧按钮
            foreach (var btn in _cmdActionButtons)
            {
                _flowPanel?.Controls.Remove(btn);
                btn.Dispose();
            }
            _cmdActionButtons.Clear();

            // 从配置读取命令列表
            var commands = CsvSettings.GetCustomCommands();
            if (commands.Count == 0 || _flowPanel == null)
            {
                LayoutTopBar();
                return;
            }

            // 按分组归类：独立按钮 + 分组按钮
            var standalone = commands.Where(c => string.IsNullOrWhiteSpace(c.GroupName)).ToList();
            var grouped = commands.Where(c => !string.IsNullOrWhiteSpace(c.GroupName))
                                  .GroupBy(c => c.GroupName)
                                  .ToList();

            // 先渲染独立按钮
            foreach (var cmd in standalone)
            {
                Color btnColor;
                try { btnColor = ColorTranslator.FromHtml(cmd.ButtonColor); }
                catch { btnColor = Theme.P.Accent; }

                var btn = new ThemeButton
                {
                    Text = cmd.Name,
                    Size = new Size(ThemeButton.FitWidth(cmd.Name), ThemeMetrics.RowH),
                    Font = Theme.Body,
                    FillOverride = btnColor,
                    Margin = new Padding(0, 0, ThemeMetrics.Gap, 0),
                    Tag = cmd
                };

                SetButtonTooltip(btn, cmd);
                btn.Click += OnCmdActionButtonClick;

                _cmdActionButtons.Add(btn);
                _flowPanel.Controls.Add(btn);
            }

            // 再渲染分组按钮（下拉菜单）
            foreach (var group in grouped)
            {
                var firstCmd = group.First();
                Color groupColor;
                try { groupColor = ColorTranslator.FromHtml(firstCmd.ButtonColor); }
                catch { groupColor = Theme.P.Accent; }

                var groupText = group.Key + " ▾";
                var groupBtn = new ThemeButton
                {
                    Text = groupText,
                    Size = new Size(ThemeButton.FitWidth(groupText), ThemeMetrics.RowH),
                    Font = Theme.Body,
                    FillOverride = groupColor,
                    Margin = new Padding(0, 0, ThemeMetrics.Gap, 0),
                    Tag = group.ToList()
                };

                var tipNames = group.Select(c => c.Name).ToList();
                var groupTip = new ToolTip { AutoPopDelay = 10000, InitialDelay = 300, ReshowDelay = 100 };
                groupTip.SetToolTip(groupBtn, string.Join("\n", tipNames));

                groupBtn.Click += OnGroupButtonClick;

                _cmdActionButtons.Add(groupBtn);
                _flowPanel.Controls.Add(groupBtn);
            }

            // 新建的按钮只带左右间距，垂直居中要靠顶栏统一重算，否则会比固定按钮高一截
            LayoutTopBar();
        }

        // ==================== 分组按钮点击：弹出下拉菜单 ====================
        private void OnGroupButtonClick(object? sender, EventArgs e)
        {
            if (sender is not Button btn || btn.Tag is not List<CustomCommandItem> groupCommands)
                return;

            var menu = new ContextMenuStrip
            {
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
                Font = Theme.Body,
                ShowImageMargin = false
            };

            foreach (var cmd in groupCommands)
            {
                var menuItem = new ToolStripMenuItem(cmd.Name)
                {
                    Tag = cmd,
                    ForeColor = Theme.P.Text,
                    BackColor = Color.Transparent,
                    Font = Theme.Body
                };
                var tipParts = new List<string>();
                if (!string.IsNullOrWhiteSpace(cmd.Command))
                    tipParts.Add($"命令: {cmd.Command}");
                if (!string.IsNullOrWhiteSpace(cmd.WorkDir))
                    tipParts.Add($"目录: {cmd.WorkDir}");
                if (tipParts.Count > 0)
                    menuItem.ToolTipText = string.Join("\n", tipParts);

                menuItem.Click += (s2, e2) =>
                {
                    if (s2 is ToolStripMenuItem mi && mi.Tag is CustomCommandItem item)
                        ExecuteCustomCommand(item);
                };
                menu.Items.Add(menuItem);
            }

            menu.Show(btn, new Point(0, btn.Height));
        }

        // ==================== 动态命令按钮点击 ====================
        private void OnCmdActionButtonClick(object? sender, EventArgs e)
        {
            if (sender is not Button btn || btn.Tag is not CustomCommandItem item)
                return;

            ExecuteCustomCommand(item);
        }

        /// <summary>执行自定义命令</summary>
        private void ExecuteCustomCommand(CustomCommandItem item)
        {
            var cmd = item.Command;
            if (string.IsNullOrWhiteSpace(cmd)) return;

            // 多行命令按顺序执行；空行不能变成 " & " 里的空语句
            var fullCmd = string.Join(" & ",
                cmd.Split(new[] { GlobalConst.CrLf, GlobalConst.Cr, GlobalConst.Lf }, StringSplitOptions.None)
                   .Select(line => line.Trim())
                   .Where(line => line.Length > 0));
            if (fullCmd.Length == 0) return;

            try
            {
                var showWindow = item.ShowWindow;
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    // /d 跳过注册表 AutoRun；/s 保证整条命令按原样交给 cmd 解析，
                    // 命令自身带引号时不会被外层引号吃掉
                    Arguments = "/d /s /c \"" + fullCmd + "\"",
                    UseShellExecute = false,
                    CreateNoWindow = !showWindow
                };

                var workDir = item.WorkDir;
                if (!string.IsNullOrWhiteSpace(workDir) && System.IO.Directory.Exists(workDir))
                {
                    psi.WorkingDirectory = workDir;
                }

                System.Diagnostics.Process.Start(psi);
            }
            catch (Exception ex)
            {
                AlertForm.Show(null, $"执行命令 [{item.Name}] 失败：{ex.Message}", "错误", AlertType.Error, exception: ex);
            }
        }

        /// <summary>设置按钮 Tooltip</summary>
        private static void SetButtonTooltip(Button btn, CustomCommandItem cmd)
        {
            var tipParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(cmd.Command))
                tipParts.Add($"命令: {cmd.Command}");
            if (!string.IsNullOrWhiteSpace(cmd.WorkDir))
                tipParts.Add($"目录: {cmd.WorkDir}");
            if (tipParts.Count > 0)
            {
                var tip = new ToolTip { AutoPopDelay = 10000, InitialDelay = 300, ReshowDelay = 100 };
                tip.SetToolTip(btn, string.Join("\n", tipParts));
            }
        }

        // ==================== 启动Loading窗体（独立STA线程运行） ====================

        /// <summary>启动时的 Loading 遮罩窗体（独立 STA 线程运行，确保动画流畅）</summary>
        private class StartupLoadingForm : Form
        {
            private readonly System.Windows.Forms.Timer _animTimer;
            private int _angle;
            private readonly Label _lblTitle;
            private readonly Label _lblDetail;
            private readonly ThemeProgressBar _progressBar;
            private readonly Panel _pnlDots;
            private readonly Label[] _dotLabels;

            public StartupLoadingForm()
            {
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;
                ShowInTaskbar = false;
                Size = new Size(420, 200);
                // 令牌在主窗构造时就 Set 过了，这里取到的是用户存档的皮肤色
                BackColor = Theme.P.Surface;
                TopMost = true;

                SetStyle(ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.UserPaint |
                         ControlStyles.DoubleBuffer |
                         ControlStyles.ResizeRedraw, true);
                UpdateStyles();

                // 标题
                _lblTitle = new Label
                {
                    Text = "正在初始化文件管理器...",
                    AutoSize = false,
                    Size = new Size(380, 28),
                    Location = new Point(20, 15),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = Theme.TitleBold,
                    ForeColor = Theme.P.Accent,
                    BackColor = Color.Transparent
                };

                // 进度条
                _progressBar = new ThemeProgressBar
                {
                    Size = new Size(380, 8),
                    Location = new Point(20, 55),
                    Value = 0
                };

                // 详细信息
                _lblDetail = new Label
                {
                    Text = "",
                    AutoSize = false,
                    Size = new Size(380, 20),
                    Location = new Point(20, 72),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = Theme.Body,
                    ForeColor = Theme.P.Muted,
                    BackColor = Color.Transparent
                };

                // 动画圆点
                _pnlDots = new Panel
                {
                    Size = new Size(380, 20),
                    Location = new Point(20, 100),
                    BackColor = Color.Transparent
                };
                _dotLabels = new Label[3];
                for (int d = 0; d < 3; d++)
                {
                    _dotLabels[d] = new Label
                    {
                        Text = "●",
                        AutoSize = true,
                        Font = new Font("Segoe UI", Theme.Body.Size + 1f),
                        ForeColor = Theme.P.Accent,
                        Location = new Point(150 + d * 28, 0),
                        BackColor = Color.Transparent
                    };
                    _pnlDots.Controls.Add(_dotLabels[d]);
                }

                Controls.Add(_lblTitle);
                Controls.Add(_progressBar);
                Controls.Add(_lblDetail);
                Controls.Add(_pnlDots);

                // 绘制圆角边框 + 旋转圆弧
                Paint += OnPaintBorder;

                // 旋转动画 ~60 FPS
                _animTimer = new System.Windows.Forms.Timer { Interval = 16 };
                _animTimer.Tick += (s, e) =>
                {
                    _angle = (_angle + 6) % 360;
                    // 跳动圆点动画
                    int step = (_angle / 10) % 4;
                    for (int d = 0; d < 3; d++)
                        _dotLabels[d].ForeColor = (d == step % 3)
                            ? Theme.P.Accent
                            : Theme.P.Border;
                    Invalidate(new Rectangle(20, 130, 72, 72)); // 只刷新圆弧区域
                };
            }

            /// <summary>更新进度（线程安全，通过 BeginInvoke 调用）</summary>
            public void UpdateProgress(int value, string detail)
            {
                if (InvokeRequired)
                {
                    BeginInvoke(new Action(() => UpdateProgress(value, detail)));
                    return;
                }
                _progressBar.Value = Math.Min(100, Math.Max(0, value));
                _lblDetail.Text = detail;
            }

            /// <summary>设置标题文字（线程安全）</summary>
            public void SetTitle(string title)
            {
                if (InvokeRequired)
                {
                    BeginInvoke(new Action(() => SetTitle(title)));
                    return;
                }
                _lblTitle.Text = title;
            }

            protected override void OnShown(EventArgs e)
            {
                base.OnShown(e);
                _animTimer.Start();
            }

            protected override void OnFormClosed(FormClosedEventArgs e)
            {
                _animTimer.Stop();
                _animTimer.Dispose();
                base.OnFormClosed(e);
            }

            private void OnPaintBorder(object? sender, PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                // 圆角矩形边框
                var rect = new Rectangle(0, 0, Width - 1, Height - 1);
                using var pen = new Pen(Theme.P.Border, 1f);
                using var path = ThemeButton.RoundRect(rect, 8);
                e.Graphics.DrawPath(pen, path);

                // 旋转圆弧动画
                DrawSpinner(e.Graphics);
            }

            private void DrawSpinner(Graphics g)
            {
                int cx = 56, cy = 155, radius = 18;
                float penWidth = 3f;

                // 背景圆环
                using var bgPen = new Pen(Theme.P.Divider, penWidth);
                g.DrawEllipse(bgPen, cx - radius, cy - radius, radius * 2, radius * 2);

                // 旋转强调色弧线
                using var arcPen = new Pen(Theme.P.Accent, penWidth)
                {
                    StartCap = System.Drawing.Drawing2D.LineCap.Round,
                    EndCap = System.Drawing.Drawing2D.LineCap.Round
                };
                g.DrawArc(arcPen,
                    cx - radius, cy - radius, radius * 2, radius * 2,
                    _angle, 270);
            }
        }
    }
}
