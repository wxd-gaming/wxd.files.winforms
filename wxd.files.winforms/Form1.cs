using System;
using System.Collections.Generic;
using System.Drawing;
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
        private ComboBox? _cboCount;
        private Label? _lblCount;
        private ComboBox? _cboLayout;
        private Label? _lblLayout;
        private ComboBox? _cboTheme;
        private Label? _lblTheme;
        private Button? _btnOpenCsvReader;
        private Panel _pnlContainer = new();
        private FileBrowserPanel[] _panels = Array.Empty<FileBrowserPanel>();
        private Panel? _pnlTopBar;
        private Button? _btnCmdManager;                      // 命令管理按钮
        private Button? _btnSettings;                         // ≡ 设置按钮
        private CheckBox? _chkShowHidden;                     // 显示隐藏文件夹复选框
        private Button? _btnSaveState;                        // 保存状态按钮
        private FlowLayoutPanel? _flowPanel;                  // 顶部流式布局面板
        private bool _showSettings = false;                   // 是否显示设置界面
        private readonly List<Button> _cmdActionButtons = new(); // 动态生成的命令按钮

        // ==================== 启动Loading相关（独立STA线程方案） ====================
        private StartupLoadingForm? _startupLoadingForm;      // 独立线程运行的Loading窗体
        private Thread? _loadingThread;                       // Loading所在线程

        private enum LayoutMode { Horizontal, Vertical, Flow }
        private enum ThemeMode { Dark, Light }

        // ==================== 浅色天蓝配色 ====================
        private static readonly Color _colorLightTopBg = Color.FromArgb(235, 245, 255);
        private static readonly Color _colorLightTopBorder = Color.FromArgb(173, 209, 237);
        private static readonly Color _colorLightTitleFg = Color.FromArgb(20, 60, 100);
        private static readonly Color _colorLightLabelFg = Color.FromArgb(90, 130, 170);
        private static readonly Color _colorLightFormBg = Color.FromArgb(235, 248, 255);

        // ==================== 深色配色 (VS风格) ====================
        private static readonly Color _colorDarkTopBg = Color.FromArgb(37, 37, 38);      // VS顶部栏
        private static readonly Color _colorDarkTopBorder = Color.FromArgb(62, 62, 66);      // VS边框
        private static readonly Color _colorDarkTitleFg = Color.FromArgb(241, 241, 241);   // VS白色文字
        private static readonly Color _colorDarkLabelFg = Color.FromArgb(200, 200, 200);   // VS灰色文字
        private static readonly Color _colorDarkFormBg = Color.FromArgb(30, 30, 30);      // VS深色背景

        private static readonly int GAP = 6; // 窗格间距
        private ThemeMode _currentTheme = ThemeMode.Light;

        public Form1()
        {
            InitializeComponent();
            // 在标题栏追加程序集版本号
            var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            if (ver != null)
                Text = $"{Text} v{ver}";
            CreateTopBar();
            InitControlPanel();

            // 恢复全局状态（窗格数量、布局方式、主题）
            var panelCount = CsvSettings.PaneCount;
            var layoutIndex = CsvSettings.LayoutIndex;
            var themeIndex = CsvSettings.ThemeMode;
            _currentTheme = (ThemeMode)themeIndex;
            _cboCount!.SelectedIndex = Math.Min(panelCount - 1, _cboCount.Items.Count - 1);
            _cboLayout!.SelectedIndex = Math.Min(layoutIndex, _cboLayout.Items.Count - 1);
            _cboTheme!.SelectedIndex = Math.Min(themeIndex, _cboTheme.Items.Count - 1);

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

            var isDark = _currentTheme == ThemeMode.Dark;
            var ownerHandle = this.Handle;
            var ownerLocation = this.Location;
            var ownerSize = this.Size;

            // ★ 信号量：Loading 窗体 Shown 后通知主线程
            var readyEvent = new ManualResetEvent(false);

            _loadingThread = new Thread(() =>
            {
                _startupLoadingForm = new StartupLoadingForm(isDark);
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
            int panelCount = _cboCount!.SelectedIndex + 1;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            SuspendLayout();
            // 
            // Form1
            // 
            ClientSize = new Size(1300, 750);
            Font = new Font("Microsoft YaHei UI", 9F);
            Icon = (Icon)resources.GetObject("$this.Icon");
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

            // 顶部工具栏面板
            _pnlTopBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 48,
                Padding = new Padding(12, 6, 12, 6),
                BackColor = _colorLightTopBg
            };

            var sep = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = _colorLightTopBorder };
            _pnlTopBar.Controls.Add(sep);

            // 容器面板
            _pnlContainer.Dock = DockStyle.Fill;
            _pnlContainer.Padding = new Padding(GAP);
            _pnlContainer.BackColor = BackColor;

            Controls.Add(_pnlContainer);
            Controls.Add(_pnlTopBar);

            ResumeLayout(false);
        }

        private void InitControlPanel()
        {
            if (_pnlTopBar == null) return;

            float fontSize = 9f;
            var labelFont = new Font("Microsoft YaHei UI", fontSize);
            var comboFont = new Font("Microsoft YaHei UI", fontSize);
            int ctrlHeight = 26;

            // 流式布局面板（自动排列控件，收起/展开时自动流动）
            _flowPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = false,
                Padding = new Padding(12, 4, 12, 0),
                BackColor = Color.Transparent
            };

            var lblTitle = new Label
            {
                Text = "多窗格文件管理",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 11f, FontStyle.Regular),
                ForeColor = _colorLightTitleFg,
                Margin = new Padding(0, 6, 16, 0)
            };

            // ≡ 设置按钮（点击显示/隐藏配置项，与标题保持足够间距）
            _btnSettings = new Button
            {
                Text = "≡",
                Size = new Size(32, ctrlHeight),
                Font = new Font("Microsoft YaHei UI", 12f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(144, 180, 210),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(120, 160, 200) },
                Margin = new Padding(0, 3, 4, 0)
            };
            _btnSettings.Click += OnSettingsClick;

            // 显示隐藏文件夹复选框（在设置按钮前面，始终可见）
            _chkShowHidden = new CheckBox
            {
                Text = "显示隐藏",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _colorLightLabelFg,
                Margin = new Padding(4, 6, 4, 0),
                Checked = CsvSettings.ShowHiddenFiles
            };
            _chkShowHidden.CheckedChanged += (s, e) =>
            {
                CsvSettings.ShowHiddenFiles = _chkShowHidden.Checked;
                // 通知所有面板刷新
                foreach (var panel in _panels)
                {
                    if (!panel.IsDisposed)
                        panel.RefreshCurrentDirectory();
                }
            };

            _lblCount = new Label
            {
                Text = "窗格数量",
                AutoSize = true,
                Font = labelFont,
                ForeColor = _colorLightLabelFg,
                Margin = new Padding(6, 6, 2, 0)
            };

            _cboCount = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Size = new Size(55, ctrlHeight),
                Font = comboFont,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 6, 0)
            };
            for (int i = 1; i <= 10; i++) _cboCount.Items.Add(i.ToString());
            _cboCount.SelectedIndex = 6;  // 默认7个窗格
            _cboCount.SelectedIndexChanged += OnCountChanged;

            _lblLayout = new Label
            {
                Text = "布局方式",
                AutoSize = true,
                Font = labelFont,
                ForeColor = _colorLightLabelFg,
                Margin = new Padding(6, 6, 2, 0)
            };

            _cboLayout = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Size = new Size(120, ctrlHeight),
                Font = comboFont,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 6, 0)
            };
            _cboLayout.Items.Add("横向布局");
            _cboLayout.Items.Add("纵向布局");
            _cboLayout.Items.Add("纵横交错流式布局");
            _cboLayout.SelectedIndex = 2;
            _cboLayout.SelectedIndexChanged += OnLayoutChanged;

            _lblTheme = new Label
            {
                Text = "主题",
                AutoSize = true,
                Font = labelFont,
                ForeColor = _colorLightLabelFg,
                Margin = new Padding(6, 6, 2, 0)
            };

            _cboTheme = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Size = new Size(80, ctrlHeight),
                Font = comboFont,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 6, 0)
            };
            _cboTheme.Items.Add("深色");
            _cboTheme.Items.Add("浅色");
            _cboTheme.SelectedIndex = 1; // 默认浅色
            _cboTheme.SelectedIndexChanged += OnThemeChanged;

            // 保存状态按钮
            _btnSaveState = new Button
            {
                Text = "保存状态",
                Size = new Size(70, ctrlHeight),
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(76, 175, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(56, 142, 60) },
                Margin = new Padding(4, 2, 2, 0)
            };
            _btnSaveState.Click += (s, e) => SaveAllState();

            // --- 命令管理按钮 ---
            _btnCmdManager = new Button
            {
                Text = "命令管理",
                Size = new Size(75, ctrlHeight),
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(38, 166, 154),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(26, 130, 120) },
                Margin = new Padding(2, 2, 4, 0)
            };
            _btnCmdManager.Click += OnCmdManagerClick;

            // 打开CSV阅读器按钮
            _btnOpenCsvReader = new Button
            {
                Text = "阅读器",
                Size = new Size(80, ctrlHeight),
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) },
                Margin = new Padding(4, 2, 2, 0)
            };
            _btnOpenCsvReader.Click += OnOpenCsvReaderClick;

            // 打开大文件阅读器按钮
            var btnLargeFileViewer = new Button
            {
                Text = "大文件阅读",
                Size = new Size(80, ctrlHeight),
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(171, 71, 188),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(123, 31, 162) },
                Margin = new Padding(2, 2, 4, 0)
            };
            btnLargeFileViewer.Click += OnOpenLargeFileViewerClick;

            // 关于 / 更新按钮
            var btnHelp = new Button
            {
                Text = "关于",
                Size = new Size(75, ctrlHeight),
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(120, 144, 156),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(96, 125, 139) },
                Margin = new Padding(2, 2, 4, 0)
            };
            btnHelp.Click += OnHelpClick;

            // 按顺序添加到流式布局面板：
            // 标题 → 显示隐藏复选框 → ≡按钮 → 配置项(默认隐藏) → 保存状态 → 命令管理(默认隐藏) → 阅读器 → 大文件阅读 → 使用说明
            _flowPanel.Controls.AddRange(new Control[] {
                lblTitle, _chkShowHidden, _btnSettings,
                _lblCount, _cboCount, _lblLayout, _cboLayout, _lblTheme, _cboTheme,
                _btnSaveState, _btnCmdManager,
                _btnOpenCsvReader, btnLargeFileViewer, btnHelp
            });

            _pnlTopBar.Controls.Add(_flowPanel);

            // 设置配置项控件默认隐藏
            UpdateSettingsVisibility();

            // 重建动态命令按钮
            RebuildCmdActionButtons();
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

        // ==================== 关于 / 更新 ====================
        private void OnHelpClick(object? sender, EventArgs e)
        {
            var isDark = _currentTheme == ThemeMode.Dark;
            UpdateForm.ShowUpdateForm(this, isDark);
        }

        // ==================== 窗格数量变更 ====================
        private void OnCountChanged(object? sender, EventArgs e)
        {
            if (_cboCount?.SelectedItem == null) return;
            CreatePanels(_cboCount.SelectedIndex + 1);
            ApplyLayout(GetCurrentLayout());
            SaveGlobalState();
        }

        // ==================== 布局方式变更 ====================
        private void OnLayoutChanged(object? sender, EventArgs e)
        {
            if (_cboLayout?.SelectedItem == null) return;
            ApplyLayout(GetCurrentLayout());
            SaveGlobalState();
        }

        /// <summary>实时保存全局状态（窗格数量+布局模式+主题）</summary>
        private void SaveGlobalState()
        {
            if (_cboCount != null && _cboLayout != null && _cboTheme != null)
            {
                CsvSettings.PaneCount = _panels.Length;
                CsvSettings.LayoutIndex = _cboLayout.SelectedIndex;
                CsvSettings.ThemeMode = _cboTheme.SelectedIndex;
            }
        }

        /// <summary>保存所有状态（包含每个面板的过滤文本）</summary>
        private void SaveAllState()
        {
            // 强制保存每个面板的过滤文本和目录
            for (int i = 0; i < _panels.Length; i++)
            {
                _panels[i].ForceSaveState();
            }

            // 保存全局状态
            SaveGlobalState();
        }

        // ==================== 主题切换 ====================
        private void OnThemeChanged(object? sender, EventArgs e)
        {
            if (_cboTheme?.SelectedItem == null) return;
            _currentTheme = (ThemeMode)_cboTheme.SelectedIndex;
            ApplyTheme();
            SaveGlobalState();
        }

        // ==================== ≡ 设置按钮点击：显示/隐藏配置项 ====================
        private void OnSettingsClick(object? sender, EventArgs e)
        {
            _showSettings = !_showSettings;
            UpdateSettingsVisibility();
        }

        /// <summary>根据 _showSettings 显示或隐藏配置项控件</summary>
        private void UpdateSettingsVisibility()
        {
            bool visible = _showSettings;

            if (_lblCount != null) _lblCount.Visible = visible;
            if (_cboCount != null) _cboCount.Visible = visible;
            if (_lblLayout != null) _lblLayout.Visible = visible;
            if (_cboLayout != null) _cboLayout.Visible = visible;
            if (_lblTheme != null) _lblTheme.Visible = visible;
            if (_cboTheme != null) _cboTheme.Visible = visible;
            if (_btnSaveState != null) _btnSaveState.Visible = visible;
            if (_btnCmdManager != null) _btnCmdManager.Visible = visible;
        }

        /// <summary>应用主题到整个窗体</summary>
        private void ApplyTheme()
        {
            var isDark = _currentTheme == ThemeMode.Dark;

            // 窗体背景
            BackColor = isDark ? _colorDarkFormBg : _colorLightFormBg;
            _pnlContainer.BackColor = BackColor;

            // 顶部工具栏
            if (_pnlTopBar != null)
            {
                _pnlTopBar.BackColor = isDark ? _colorDarkTopBg : _colorLightTopBg;
                // 更新分隔线颜色
                foreach (Control c in _pnlTopBar.Controls)
                {
                    if (c is Panel p && p.Height == 1)
                        p.BackColor = isDark ? _colorDarkTopBorder : _colorLightTopBorder;
                }
            }

            // 标签文字颜色 - 遍历_flowPanel中的控件
            var labelFg = isDark ? _colorDarkLabelFg : _colorLightLabelFg;
            var titleFg = isDark ? _colorDarkTitleFg : _colorLightTitleFg;

            if (_flowPanel != null)
            {
                foreach (Control c in _flowPanel.Controls)
                {
                    if (c is Label lbl)
                    {
                        if (lbl.Font.Size > 10) // 标题
                            lbl.ForeColor = titleFg;
                        else
                            lbl.ForeColor = labelFg;
                    }
                }
            }

            // 应用到所有子窗格
            foreach (var panel in _panels)
            {
                ApplyThemeToPanel(panel, isDark);
            }
        }

        /// <summary>应用主题到单个文件浏览面板</summary>
        private void ApplyThemeToPanel(FileBrowserPanel panel, bool isDark)
        {
            // 使用反射或公共方法设置面板主题
            panel.ApplyTheme(isDark);
        }

        private LayoutMode GetCurrentLayout() => _cboLayout?.SelectedItem switch
        {
            "纵向布局" => LayoutMode.Vertical,
            "纵横交错流式布局" => LayoutMode.Flow,
            _ => LayoutMode.Horizontal
        };

        // ==================== 创建/销毁窗格 ====================

        /// <summary>创建面板并立即导航到保存的目录（运行时切换数量时使用）</summary>
        private void CreatePanels(int count)
        {
            foreach (var p in _panels) p.Dispose();
            _pnlContainer.Controls.Clear();

            _panels = new FileBrowserPanel[count];
            for (int i = 0; i < count; i++)
            {
                _panels[i] = new FileBrowserPanel(i)
                {
                    Dock = DockStyle.None,
                    BorderStyle = BorderStyle.FixedSingle,
                    Padding = new Padding(0)
                };
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

            _panels = new FileBrowserPanel[count];
            for (int i = 0; i < count; i++)
            {
                // 使用轻量构造模式：skipAutoNavigate=true 跳过构造函数中的自动导航
                _panels[i] = new FileBrowserPanel(i, skipAutoNavigate: true)
                {
                    Dock = DockStyle.None,
                    BorderStyle = BorderStyle.FixedSingle,
                    Padding = new Padding(0),
                    Visible = false  // 初始隐藏，避免空白闪烁
                };
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

            // 容器可用区域（已减去 Padding/GAP）
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
            if (_panels.Length > 0 && _cboLayout != null && _pnlContainer.ClientRectangle.Width > 10)
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
                var filter = _panels[i].FilterText;
                System.Diagnostics.Debug.WriteLine($"OnFormClosing: Panel {i} Filter={filter}");
                CsvSettings.SetPaneFilter(i, filter);
                _panels[i].ForceSaveState();
            }

            CsvSettings.PaneCount = _panels.Length;
            CsvSettings.LayoutIndex = _cboLayout?.SelectedIndex ?? 2;
            CsvSettings.ThemeMode = _cboTheme?.SelectedIndex ?? 1;
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
            if (commands.Count == 0 || _flowPanel == null) return;

            for (int i = 0; i < commands.Count; i++)
            {
                var cmd = commands[i];
                Color btnColor;
                try { btnColor = ColorTranslator.FromHtml(cmd.ButtonColor); }
                catch { btnColor = Color.FromArgb(123, 31, 162); }

                var btn = new Button
                {
                    Text = cmd.Name,
                    Size = new Size(85, 26),
                    Font = new Font("Microsoft YaHei UI", 9f),
                    Cursor = Cursors.Hand,
                    BackColor = btnColor,
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(100, 30, 140) },
                    Margin = new Padding(2, 2, 2, 0),
                    Tag = cmd
                };

                // Tooltip：显示执行的命令和工作目录
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

                btn.Click += OnCmdActionButtonClick;

                _cmdActionButtons.Add(btn);
                // 在流式布局面板末尾追加
                _flowPanel.Controls.Add(btn);
            }
        }

        // ==================== 动态命令按钮点击 ====================
        private void OnCmdActionButtonClick(object? sender, EventArgs e)
        {
            if (sender is not Button btn || btn.Tag is not CustomCommandItem item)
                return;

            var cmd = item.Command;
            if (string.IsNullOrWhiteSpace(cmd)) return;

            // 将多行命令用 & 连接，实现逐条依次执行
            var lines = cmd.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var fullCmd = string.Join(" & ", lines);

            try
            {
                var showWindow = item.ShowWindow;
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c " + fullCmd,
                    UseShellExecute = false,
                    CreateNoWindow = !showWindow
                };

                // 如果配置了工作目录且存在，则设置工作目录
                var workDir = item.WorkDir;
                if (!string.IsNullOrWhiteSpace(workDir) && System.IO.Directory.Exists(workDir))
                {
                    psi.WorkingDirectory = workDir;
                }

                System.Diagnostics.Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行命令 [{item.Name}] 失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            private readonly ProgressBar _progressBar;
            private readonly Panel _pnlDots;
            private readonly Label[] _dotLabels;

            public StartupLoadingForm(bool isDark)
            {
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;
                ShowInTaskbar = false;
                Size = new Size(420, 200);
                BackColor = isDark ? Color.FromArgb(45, 45, 48) : Color.White;
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
                    Font = new Font("Microsoft YaHei UI", 12f, FontStyle.Bold),
                    ForeColor = isDark ? Color.FromArgb(241, 241, 241) : Color.FromArgb(30, 136, 229),
                    BackColor = Color.Transparent
                };

                // 进度条
                _progressBar = new ProgressBar
                {
                    Size = new Size(380, 8),
                    Location = new Point(20, 55),
                    Style = ProgressBarStyle.Continuous,
                    Minimum = 0,
                    Maximum = 100,
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
                    Font = new Font("Microsoft YaHei UI", 9f),
                    ForeColor = isDark ? Color.FromArgb(180, 180, 180) : Color.FromArgb(120, 160, 200),
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
                        Font = new Font("Segoe UI", 10f),
                        ForeColor = Color.FromArgb(66, 165, 245),
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
                            ? Color.FromArgb(100, 181, 246)
                            : Color.FromArgb(180, 200, 220);
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
                using var pen = new Pen(Color.FromArgb(200, 210, 225), 1f);
                using var path = GetRoundedRectPath(rect, 8);
                e.Graphics.DrawPath(pen, path);

                // 旋转圆弧动画
                DrawSpinner(e.Graphics);
            }

            private void DrawSpinner(Graphics g)
            {
                int cx = 56, cy = 155, radius = 18;
                float penWidth = 3f;

                // 背景圆环
                using var bgPen = new Pen(Color.FromArgb(230, 235, 240), penWidth);
                g.DrawEllipse(bgPen, cx - radius, cy - radius, radius * 2, radius * 2);

                // 旋转蓝色弧线
                using var arcPen = new Pen(Color.FromArgb(30, 136, 229), penWidth)
                {
                    StartCap = System.Drawing.Drawing2D.LineCap.Round,
                    EndCap = System.Drawing.Drawing2D.LineCap.Round
                };
                g.DrawArc(arcPen,
                    cx - radius, cy - radius, radius * 2, radius * 2,
                    _angle, 270);
            }

            private static System.Drawing.Drawing2D.GraphicsPath GetRoundedRectPath(
                Rectangle rect, int radius)
            {
                var path = new System.Drawing.Drawing2D.GraphicsPath();
                int d = radius * 2;
                path.AddArc(rect.X, rect.Y, d, d, 180, 90);
                path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
                path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
                path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }
        }
    }
}
