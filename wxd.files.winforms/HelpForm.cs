using System;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>通用帮助信息展示窗体，配色跟随当前皮肤</summary>
    public class HelpForm : Form
    {
        /// <summary>正文区：更新记录窗体继承后直接往这里填内容</summary>
        protected readonly RichTextBox _rtbContent = new();
        private readonly Label _lblTitle = new();
        private readonly Panel _topBar = new();
        private readonly Panel _topDivider = new();
        private readonly Panel _bar = new();
        private readonly ThemeButton _btnClose = new();

        /// <summary>创建帮助窗体</summary>
        /// <param name="title">标题栏文字</param>
        /// <param name="content">帮助内容（RTF格式，纯文本也可）</param>
        public HelpForm(string title, string content)
        {
            Text = title;
            ClientSize = new Size(700, 560);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;

            _topBar.Dock = DockStyle.Top;
            _topBar.Height = ThemeMetrics.BarHeight;
            _topDivider.Dock = DockStyle.Bottom;
            _topDivider.Height = 1;
            _topBar.Controls.Add(_topDivider);

            // 标题和窗口标题重复，但正文需要一个可见的抬头，保留并和其他对话框的页眉同规格
            _lblTitle.Text = title;
            _lblTitle.AutoSize = false;
            _lblTitle.TextAlign = ContentAlignment.MiddleLeft;
            _lblTitle.BackColor = Color.Transparent;
            _topBar.Controls.Add(_lblTitle);
            _topBar.Resize += (_, _) => LayoutTopBar();

            _rtbContent.BorderStyle = BorderStyle.None;
            _rtbContent.ReadOnly = true;
            _rtbContent.ScrollBars = RichTextBoxScrollBars.Vertical;
            _rtbContent.DetectUrls = false;
            _rtbContent.WordWrap = true;
            _rtbContent.Dock = DockStyle.Fill;
            _rtbContent.Text = content;

            // 正文要缩进，但页眉页脚的分隔线要铺满整宽，所以只给正文套一层留白容器
            var body = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(ThemeMetrics.Margin + 6, ThemeMetrics.Gap,
                                     ThemeMetrics.Margin + 6, ThemeMetrics.RowPad)
            };
            body.Controls.Add(_rtbContent);

            _bar = DialogButtonBar.Create();
            _btnClose.Text = "关闭";
            _btnClose.Primary = true;
            _btnClose.DialogResult = DialogResult.OK;
            _btnClose.Click += (s, e) => Close();
            _bar.Controls.Add(_btnClose);
            _bar.Resize += (_, _) => LayoutBar();

            // 停靠顺序：同一条边上后加入的靠外，所以正文先加、页眉页脚后加
            Controls.Add(body);
            Controls.Add(_topBar);
            Controls.Add(_bar);

            AcceptButton = _btnClose;
            CancelButton = _btnClose;

            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
        }

        private void LayoutTopBar()
        {
            _lblTitle.SetBounds(ThemeMetrics.Margin, 0,
                _topBar.Width - ThemeMetrics.Margin * 2, ThemeMetrics.BarHeight - 1);
        }

        private void LayoutBar() => DialogButtonBar.LayoutRight(_bar, _btnClose);

        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _topBar.BackColor = p.Chrome;
            _bar.BackColor = p.Chrome;
            _topDivider.BackColor = p.Divider;

            _lblTitle.ForeColor = p.Text;
            _lblTitle.Font = Theme.Head;
            _rtbContent.BackColor = p.Surface;
            _rtbContent.ForeColor = p.Text;
            _rtbContent.Font = Theme.Body;
            _btnClose.BackColor = p.Chrome;

            LayoutTopBar();
            LayoutBar();
            Invalidate(true);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // 滚动到顶部
            _rtbContent.SelectionStart = 0;
            _rtbContent.ScrollToCaret();
            _btnClose.Focus();
        }

        /// <summary>便捷方法：显示帮助窗体</summary>
        public static void ShowHelp(IWin32Window owner, string title, string content)
        {
            using var form = new HelpForm(title, content);
            form.ShowDialog(owner);
        }
    }
}
