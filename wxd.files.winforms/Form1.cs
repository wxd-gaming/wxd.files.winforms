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
        private Panel     _pnlContainer = new();
        private FileBrowserPanel[] _panels = Array.Empty<FileBrowserPanel>();
        private Panel?    _pnlTopBar;

        private enum LayoutMode { Horizontal, Vertical, Flow }

        // ==================== 配色 ====================
        private static readonly Color _colorTopBg      = Color.FromArgb(245, 246, 247);
        private static readonly Color _colorTopBorder   = Color.FromArgb(218, 220, 224);
        private static readonly Color _colorTitleFg     = Color.FromArgb(32, 33, 36);
        private static readonly Color _colorLabelFg     = Color.FromArgb(95, 99, 104);

        private static readonly int GAP = 6; // 窗格间距

        public Form1()
        {
            InitializeComponent();
            CreateTopBar();
            InitControlPanel();

            // 恢复全局状态（窗格数量、布局方式）
            var (panelCount, layoutIndex) = FileBrowserPanel.LoadGlobalState();
            _cboCount!.SelectedIndex = Math.Min(panelCount - 1, _cboCount.Items.Count - 1);
            _cboLayout!.SelectedIndex = Math.Min(layoutIndex, _cboLayout.Items.Count - 1);

            CreatePanels(_cboCount.SelectedIndex + 1);
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
            SuspendLayout();

            ClientSize = new Size(1300, 750);
            Text       = "多窗格文件夹管理器";
            MinimumSize = new Size(600, 450);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor   = Color.FromArgb(250, 250, 250);
            Font        = new Font("Microsoft YaHei UI", 9f);

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
                BackColor = _colorTopBg
            };

            var sep = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = _colorTopBorder };
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
                ForeColor = _colorTitleFg
            };

            _lblCount = new Label
            {
                Text     = "窗格数量",
                Location = new Point(180, y + 2),
                AutoSize = true,
                Font     = labelFont,
                ForeColor = _colorLabelFg
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
                ForeColor = _colorLabelFg
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

            _pnlTopBar.Controls.AddRange(new Control[] { lblTitle, _lblCount, _cboCount, _lblLayout, _cboLayout });
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

        /// <summary>实时保存全局状态（窗格数量+布局模式）</summary>
        private void SaveGlobalState()
        {
            if (_cboCount != null && _cboLayout != null)
                FileBrowserPanel.SaveGlobalState(_panels.Length, _cboLayout.SelectedIndex);
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

        /// <summary>流式排列：自动换行网格布局</summary>
        private void DoFlow(int count, int w, int h)
        {
            // 计算最佳列数：使单元格接近正方形
            int cols = Math.Max(1, (int)Math.Round(Math.Sqrt((double)count * w / Math.Max(h, 1))));
            cols = Math.Min(cols, count);
            int rows = (count + cols - 1) / cols;

            int cellW = w / cols;
            int cellH = h / rows;

            for (int i = 0; i < count; i++)
            {
                int row = i / cols;
                int col = i % cols;
                bool isLastCol = (col == cols - 1);
                bool isLastRow = (row == rows - 1);

                int x = _pnlContainer.Padding.Left + col * cellW;
                int y = _pnlContainer.Padding.Top  + row  * cellH;

                int cw = isLastCol ? (w - col * cellW) : cellW;
                int ch = isLastRow ? (h - row * cellH) : cellH;

                _panels[i].SetBounds(x, y, cw, ch);
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
            FileBrowserPanel.SaveGlobalState(_panels.Length, _cboLayout?.SelectedIndex ?? 2);
            base.OnFormClosing(e);
        }
    }
}
