using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>更新记录展示窗体，显示最近 N 次 Release 的更新日志</summary>
    public class ChangelogForm : Form
    {
        private readonly RichTextBox _rtbContent;
        private readonly Button _btnClose;
        private readonly bool _isDark;

        // 浅色配色
        private static readonly Color _lightBg = Color.FromArgb(235, 248, 255);
        private static readonly Color _lightTitleFg = Color.FromArgb(20, 60, 100);
        private static readonly Color _lightBtnBg = Color.FromArgb(66, 165, 245);
        private static readonly Color _lightBtnBorder = Color.FromArgb(30, 136, 229);

        // 深色配色
        private static readonly Color _darkBg = Color.FromArgb(30, 30, 30);
        private static readonly Color _darkTitleFg = Color.FromArgb(241, 241, 241);
        private static readonly Color _darkBtnBg = Color.FromArgb(0, 122, 204);
        private static readonly Color _darkBtnBorder = Color.FromArgb(0, 95, 160);

        public ChangelogForm(bool isDark = false)
        {
            _isDark = isDark;
            Text = "更新记录";
            ClientSize = new Size(700, 560);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = isDark ? _darkBg : _lightBg;
            Padding = new Padding(16, 16, 16, 12);

            // === 标题标签 ===
            var lblTitle = new Label
            {
                Text = "📋 更新记录",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 14f, FontStyle.Bold),
                ForeColor = isDark ? _darkTitleFg : _lightTitleFg,
                Location = new Point(16, 12),
                BackColor = Color.Transparent
            };

            // === 分隔线 ===
            var sep = new Panel
            {
                Location = new Point(16, 42),
                Size = new Size(668, 1),
                BackColor = isDark ? Color.FromArgb(62, 62, 66) : Color.FromArgb(173, 209, 237)
            };

            // === 富文本框（只读可滚动） ===
            _rtbContent = new RichTextBox
            {
                Location = new Point(16, 52),
                Size = new Size(668, 450),
                Font = new Font("Microsoft YaHei UI", 10f),
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                DetectUrls = false,
                WordWrap = true,
                BackColor = isDark ? Color.FromArgb(37, 37, 38) : Color.White,
                ForeColor = isDark ? Color.FromArgb(220, 220, 220) : Color.FromArgb(30, 60, 100)
            };
            _rtbContent.Text = "正在获取更新记录...";

            // === 关闭按钮 ===
            _btnClose = new Button
            {
                Text = "关闭",
                Size = new Size(100, 30),
                Location = new Point(584, 514),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = isDark ? _darkBtnBg : _lightBtnBg,
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.OK
            };
            _btnClose.FlatAppearance.BorderSize = 1;
            _btnClose.FlatAppearance.BorderColor = isDark ? _darkBtnBorder : _lightBtnBorder;
            _btnClose.Click += (s, e) => Close();

            Controls.AddRange(new Control[] { lblTitle, sep, _rtbContent, _btnClose });

            AcceptButton = _btnClose;
            CancelButton = _btnClose;
        }

        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            _btnClose.Focus();

            // 异步加载更新记录
            try
            {
                var notes = await UpdateHelper.GetRecentReleaseNotesAsync(5);
                if (string.IsNullOrWhiteSpace(notes))
                {
                    _rtbContent.Text = "暂无更新记录。";
                }
                else
                {
                    _rtbContent.Text = notes;
                }
            }
            catch
            {
                _rtbContent.Text = "获取更新记录失败，请检查网络连接后重试。";
            }

            _rtbContent.SelectionStart = 0;
            _rtbContent.ScrollToCaret();
        }

        /// <summary>便捷方法：显示更新记录窗体</summary>
        public static void ShowChangelog(IWin32Window owner, bool isDark = false)
        {
            using var form = new ChangelogForm(isDark);
            form.ShowDialog(owner);
        }
    }
}
