using System;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    public partial class Form1 : Form
    {
        private ComboBox? _cboCount;
        private Label?    _lblCount;
        private ComboBox? _cboLayout;
        private Label?    _lblLayout;
        private ComboBox? _cboTheme;
        private Label?    _lblTheme;
        private Button?   _btnOpenCsvReader;
        private Panel     _pnlContainer = new();
        private FileBrowserPanel[] _panels = Array.Empty<FileBrowserPanel>();
        private Panel?    _pnlTopBar;

        private enum LayoutMode { Horizontal, Vertical, Flow }
        private enum ThemeMode { Dark, Light }

        // ==================== 浅色配色 ====================
        private static readonly Color _colorLightTopBg      = Color.FromArgb(245, 246, 247);
        private static readonly Color _colorLightTopBorder   = Color.FromArgb(218, 220, 224);
        private static readonly Color _colorLightTitleFg     = Color.FromArgb(32, 33, 36);
        private static readonly Color _colorLightLabelFg     = Color.FromArgb(95, 99, 104);
        private static readonly Color _colorLightFormBg      = Color.FromArgb(250, 250, 250);

        // ==================== 深色配色 (VS风格) ====================
        private static readonly Color _colorDarkTopBg        = Color.FromArgb(37, 37, 38);      // VS顶部栏
        private static readonly Color _colorDarkTopBorder    = Color.FromArgb(62, 62, 66);      // VS边框
        private static readonly Color _colorDarkTitleFg      = Color.FromArgb(241, 241, 241);   // VS白色文字
        private static readonly Color _colorDarkLabelFg      = Color.FromArgb(200, 200, 200);   // VS灰色文字
        private static readonly Color _colorDarkFormBg       = Color.FromArgb(30, 30, 30);      // VS深色背景

        private static readonly int GAP = 6; // 窗格间距
        private ThemeMode _currentTheme = ThemeMode.Light;

        public Form1()
        {
            InitializeComponent();
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

            CreatePanels(_cboCount.SelectedIndex + 1);
            // 应用主题
            ApplyTheme();
            // 首次可见时执行布局（此时 ClientRectangle 已是最终尺寸）
            VisibleChanged += OnFirstVisible;
        }

        /// <summary>仅在首次变为可见时触发一次布局</summary>
        private void OnFirstVisible(object? sender, EventArgs e)
        {
            if (Visible)
            {
                VisibleChanged -= OnFirstVisible;
                // 用 BeginInvoke 确保在当前消息处理完后执行，此时布局已完成
                BeginInvoke(() => ApplyLayout(GetCurrentLayout()));
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
                Dock      = DockStyle.Top,
                Height    = 48,
                Padding   = new Padding(12, 6, 12, 6),
                BackColor = _colorLightTopBg
            };

            var sep = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = _colorLightTopBorder };
            _pnlTopBar.Controls.Add(sep);

            // 容器面板
            _pnlContainer.Dock     = DockStyle.Fill;
            _pnlContainer.Padding  = new Padding(GAP);
            _pnlContainer.BackColor = BackColor;

            Controls.Add(_pnlContainer);
            Controls.Add(_pnlTopBar);

            ResumeLayout(false);
        }

        private void InitControlPanel()
        {
            if (_pnlTopBar == null) return;

            int y = 12;
            float fontSize = 9f;
            var labelFont = new Font("Microsoft YaHei UI", fontSize);
            var comboFont = new Font("Microsoft YaHei UI", fontSize);

            var lblTitle = new Label
            {
                Text      = "多窗格文件管理",
                Location  = new Point(14, y),
                AutoSize  = true,
                Font      = new Font("Microsoft YaHei UI", 11f, FontStyle.Regular),
                ForeColor = _colorLightTitleFg
            };

            _lblCount = new Label
            {
                Text     = "窗格数量",
                Location = new Point(180, y + 2),
                AutoSize = true,
                Font     = labelFont,
                ForeColor = _colorLightLabelFg
            };

            _cboCount = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location      = new Point(245, y - 1),
                Size          = new Size(55, 26),
                Font          = comboFont,
                FlatStyle     = FlatStyle.System
            };
            for (int i = 1; i <= 10; i++) _cboCount.Items.Add(i.ToString());
            _cboCount.SelectedIndex = 3;
            _cboCount.SelectedIndexChanged += OnCountChanged;

            _lblLayout = new Label
            {
                Text     = "布局方式",
                Location = new Point(315, y + 2),
                AutoSize = true,
                Font     = labelFont,
                ForeColor = _colorLightLabelFg
            };

            _cboLayout = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location      = new Point(380, y - 1),
                Size          = new Size(120, 26),
                Font          = comboFont,
                FlatStyle     = FlatStyle.System
            };
            _cboLayout.Items.Add("横向布局");
            _cboLayout.Items.Add("纵向布局");
            _cboLayout.Items.Add("纵横交错流式布局");
            _cboLayout.SelectedIndex = 2;
            _cboLayout.SelectedIndexChanged += OnLayoutChanged;

            _lblTheme = new Label
            {
                Text     = "主题",
                Location = new Point(510, y + 2),
                AutoSize = true,
                Font     = labelFont,
                ForeColor = _colorLightLabelFg
            };

            _cboTheme = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location      = new Point(545, y - 1),
                Size          = new Size(80, 26),
                Font          = comboFont,
                FlatStyle     = FlatStyle.System
            };
            _cboTheme.Items.Add("深色");
            _cboTheme.Items.Add("浅色");
            _cboTheme.SelectedIndex = 1; // 默认浅色
            _cboTheme.SelectedIndexChanged += OnThemeChanged;

            // 打开CSV阅读器按钮
            _btnOpenCsvReader = new Button
            {
                Text       = "CSV阅读器",
                Location   = new Point(635, y - 1),
                Size       = new Size(80, 26),
                Font       = new Font("Microsoft YaHei UI", 9f),
                Cursor     = Cursors.Hand,
                BackColor  = Color.FromArgb(0, 122, 204),
                ForeColor  = Color.White,
                FlatStyle  = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 }
            };
            _btnOpenCsvReader.Click += OnOpenCsvReaderClick;

            _pnlTopBar.Controls.AddRange(new Control[] { lblTitle, _lblCount, _cboCount, _lblLayout, _cboLayout, _lblTheme, _cboTheme, _btnOpenCsvReader });
        }

        // ==================== 打开CSV阅读器 ====================
        private void OnOpenCsvReaderClick(object? sender, EventArgs e)
        {
            var csvForm = new CsvReaderForm("");  // 传入空文件夹
            csvForm.Show();
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

        // ==================== 主题切换 ====================
        private void OnThemeChanged(object? sender, EventArgs e)
        {
            if (_cboTheme?.SelectedItem == null) return;
            _currentTheme = (ThemeMode)_cboTheme.SelectedIndex;
            ApplyTheme();
            SaveGlobalState();
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

            // 标签文字颜色
            var labelFg = isDark ? _colorDarkLabelFg : _colorLightLabelFg;
            var titleFg = isDark ? _colorDarkTitleFg : _colorLightTitleFg;

            foreach (Control c in _pnlTopBar?.Controls!)
            {
                if (c is Label lbl)
                {
                    if (lbl.Font.Size > 10) // 标题
                        lbl.ForeColor = titleFg;
                    else
                        lbl.ForeColor = labelFg;
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
            "纵向布局"           => LayoutMode.Vertical,
            "纵横交错流式布局" => LayoutMode.Flow,
            _                    => LayoutMode.Horizontal
        };

        // ==================== 创建/销毁窗格 ====================
        private void CreatePanels(int count)
        {
            foreach (var p in _panels) p.Dispose();
            _pnlContainer.Controls.Clear();

            _panels = new FileBrowserPanel[count];
            for (int i = 0; i < count; i++)
            {
                _panels[i] = new FileBrowserPanel(i)
                {
                    Dock       = DockStyle.None,
                    BorderStyle = BorderStyle.FixedSingle,
                    Padding     = new Padding(0)
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
            int availW = _pnlContainer.ClientRectangle.Width  - _pnlContainer.Padding.Horizontal;
            int availH = _pnlContainer.ClientRectangle.Height - _pnlContainer.Padding.Top - _pnlContainer.Padding.Bottom;

            // 保护最小值，避免极端情况
            if (availW < 50) availW = 50;
            if (availH < 50) availH = 50;

            switch (mode)
            {
                case LayoutMode.Horizontal: DoHorizontal(count, availW, availH); break;
                case LayoutMode.Vertical:   DoVertical(count, availW, availH); break;
                case LayoutMode.Flow:       DoFlow(count, availW, availH); break;
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
            CsvSettings.PaneCount = _panels.Length;
            CsvSettings.LayoutIndex = _cboLayout?.SelectedIndex ?? 2;
            CsvSettings.ThemeMode = _cboTheme?.SelectedIndex ?? 1;
            base.OnFormClosing(e);
        }
    }
}
