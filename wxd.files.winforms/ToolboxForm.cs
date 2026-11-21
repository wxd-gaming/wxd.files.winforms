using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 工具集窗口：JSON、YAML、时间戳、编码、哈希、Cron 六个页签。
    /// JSON 页签沿用原来的文本 + 树布局，其余页签都是即用即走的转换工具。
    /// </summary>
    internal sealed class ToolboxForm : Form
    {
        private readonly ThemeTabControl _tabs = new();
        private readonly JsonToolPanel _json = new();
        private TabPage? _jsonPage;

        internal ToolboxForm()
        {
            InitializeComponent();
            BuildTabs();
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            FormClosing += OnFormClosing;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;

            _json.CaptionChanged += (_, _) => UpdateJsonCaption();
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            ClientSize = new Size(1320, 820);
            MinimumSize = new Size(1000, 620);
            Text = "工具箱";
            StartPosition = FormStartPosition.CenterScreen;
            Font = Theme.Body;
            Icon = WindowIcons.Reader;
            KeyPreview = true;
            ResumeLayout(false);
        }

        // ==================== 页签 ====================

        private void BuildTabs()
        {
            _tabs.Dock = DockStyle.Fill;
            _tabs.Font = Theme.Body;
            _tabs.ItemSize = new Size(170, Theme.P.ButtonHeight + 8);
            _tabs.SizeMode = TabSizeMode.Fixed;
            _tabs.Multiline = true;
            _tabs.BackColor = Theme.P.Window;
            _tabs.Padding = new Point(12, 4);
            _tabs.DrawMode = TabDrawMode.OwnerDrawFixed;
            _tabs.DrawItem += DrawTab;

            _jsonPage = AddTab("JSON", _json);
            AddTab("YAML", new YamlToolPanel());
            AddTab("时间戳", new TimeToolPanel());
            AddTab("编码", new EncodingToolPanel());
            AddTab("哈希", new HashToolPanel());
            AddTab("Cron", new CronToolPanel());

            Controls.Add(_tabs);
        }

        private TabPage AddTab(string text, Control content)
        {
            var page = new TabPage(text)
            {
                BackColor = Theme.P.Window,
                Padding = new Padding(0)
            };
            content.Dock = DockStyle.Fill;
            page.Controls.Add(content);
            _tabs.TabPages.Add(page);
            return page;
        }

        private void UpdateJsonCaption()
        {
            if (_jsonPage == null || _jsonPage.IsDisposed) return;
            _jsonPage.Text = _json.TabCaption;
            _tabs.Invalidate();
        }

        private void ApplyTheme()
        {
            if (IsDisposed) return;
            var p = Theme.P;

            Font = Theme.Body;
            BackColor = p.Window;
            _tabs.Font = Theme.Body;
            _tabs.BackColor = p.Window;
            _tabs.ItemSize = new Size(170, p.ButtonHeight + 8);

            foreach (TabPage page in _tabs.TabPages)
                page.BackColor = p.Window;

            _tabs.Invalidate();
            Invalidate(true);
        }

        private void DrawTab(object? sender, DrawItemEventArgs e)
        {
            var page = _tabs.TabPages[e.Index];
            Rectangle tabRect = _tabs.GetTabRect(e.Index);
            bool selected = _tabs.SelectedIndex == e.Index;
            var g = e.Graphics;

            // 页签本体四周留了 2px 缝，先按皮肤底色铺一层，否则缝里露的是 comctl32 的系统灰
            using (var band = new SolidBrush(Theme.P.Window))
                g.FillRectangle(band, tabRect);

            int bottomExtra = selected ? 2 : 0;
            var roundRect = new Rectangle(
                tabRect.X + 2, tabRect.Y + 2,
                tabRect.Width - 4, tabRect.Height - 2 + bottomExtra);
            int radius = Theme.P.Radius + 2;
            var backColor = selected ? Theme.P.Accent : Theme.P.Selection;
            var borderColor = selected ? Theme.P.Accent : Theme.P.Border;

            using (var path = new GraphicsPath())
            {
                path.AddArc(roundRect.X, roundRect.Y, radius * 2, radius * 2, 180, 90);
                path.AddArc(roundRect.Right - radius * 2, roundRect.Y, radius * 2, radius * 2, 270, 90);
                path.AddLine(roundRect.Right, roundRect.Y + radius, roundRect.Right, roundRect.Bottom);
                path.AddLine(roundRect.Right, roundRect.Bottom, roundRect.X, roundRect.Bottom);
                path.AddLine(roundRect.X, roundRect.Bottom, roundRect.X, roundRect.Y + radius);
                path.CloseFigure();

                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var brush = new SolidBrush(backColor))
                    g.FillPath(brush, path);
                using (var pen = new Pen(borderColor, 1f))
                    g.DrawPath(pen, path);
            }

            var textRect = new Rectangle(tabRect.X + 8, tabRect.Y + 2, tabRect.Width - 16, tabRect.Height - 4);
            var textColor = selected ? Theme.P.AccentText : Theme.P.Text;
            TextRenderer.DrawText(g, page.Text, selected ? Theme.BodyBold : Theme.Body, textRect, textColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            if (!_json.ConfirmClose()) e.Cancel = true;
        }
    }
}
