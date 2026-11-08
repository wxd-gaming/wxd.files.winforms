using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>更新提示窗体：发现更新 → 下载进度 → 完成退出</summary>
    public class UpdateNotifyForm : Form
    {
        private readonly CheckResult _result;

        /// <summary>用户选择：0=未选, 1=立即更新, 2=跳过版本, 3=取消</summary>
        public int UserChoice { get; private set; }

        private readonly Label _lblMsg = new();
        private readonly Label _lblTitle = new();
        private readonly Label _lblNotesTitle = new();
        private readonly Label _lblNoNotes = new();
        private readonly ThemeButton _btnUpdate = new();
        private readonly ThemeButton _btnSkip = new();
        private readonly ThemeButton _btnCancel = new();
        private readonly Label _lblProgressText = new();
        private readonly ThemeProgressBar _progressBar = new();
        private readonly Panel _topBar = new();
        private readonly Panel _topDivider = new();
        private readonly Panel _body = new();
        private readonly Panel _progressRow = new();
        private readonly Panel _bar = new();
        private readonly RichTextBox? _rtbReleaseNotes;

        public UpdateNotifyForm(CheckResult result)
        {
            _result = result;
            UserChoice = 0;

            bool hasNotes = !string.IsNullOrEmpty(result.ReleaseNotes);

            Text = "发现更新";
            ClientSize = new Size(600, hasNotes ? 520 : 200);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;
            TopMost = true;

            _topBar.Dock = DockStyle.Top;
            _topBar.Height = ThemeMetrics.BarHeight;
            _topDivider.Dock = DockStyle.Bottom;
            _topDivider.Height = 1;
            _topBar.Controls.Add(_topDivider);

            _lblTitle.Text = "发现新版本";
            _lblTitle.AutoSize = false;
            _lblTitle.TextAlign = ContentAlignment.MiddleLeft;
            _lblTitle.BackColor = Color.Transparent;
            _topBar.Controls.Add(_lblTitle);
            _topBar.Resize += (_, _) => LayoutTopBar();

            _body.Dock = DockStyle.Fill;

            _lblMsg.AutoSize = false;
            _lblMsg.TextAlign = ContentAlignment.MiddleLeft;
            _lblMsg.BackColor = Color.Transparent;
            _lblMsg.Text = $"当前版本：v{result.CurrentVersion}\n最新版本：v{result.RemoteVersion}";
            _body.Controls.Add(_lblMsg);

            if (hasNotes)
            {
                _lblNotesTitle.Text = "更新内容：";
                _lblNotesTitle.AutoSize = true;
                _lblNotesTitle.BackColor = Color.Transparent;
                _body.Controls.Add(_lblNotesTitle);

                _rtbReleaseNotes = new RichTextBox
                {
                    Text = result.ReleaseNotes,
                    ReadOnly = true,
                    BorderStyle = BorderStyle.None,
                    ScrollBars = RichTextBoxScrollBars.Vertical,
                    DetectUrls = false,
                    WordWrap = true,
                    TabStop = false
                };
                _body.Controls.Add(new ThemeBordered(_rtbReleaseNotes));
            }
            else
            {
                _lblNoNotes.Text = "暂无更新记录";
                _lblNoNotes.AutoSize = true;
                _lblNoNotes.BackColor = Color.Transparent;
                _body.Controls.Add(_lblNoNotes);
            }
            _body.Resize += (_, _) => LayoutBody();

            // 进度行贴在按钮条上方，未开始下载时整行隐藏
            _progressRow.Dock = DockStyle.Bottom;
            _progressRow.Height = ThemeMetrics.RowH + ThemeMetrics.Gap;
            _progressRow.Visible = false;
            _lblProgressText.AutoSize = false;
            _lblProgressText.TextAlign = ContentAlignment.MiddleRight;
            _lblProgressText.BackColor = Color.Transparent;
            _progressRow.Controls.Add(_progressBar);
            _progressRow.Controls.Add(_lblProgressText);
            _progressRow.Resize += (_, _) => LayoutProgressRow();

            _bar = DialogButtonBar.Create();
            _btnSkip.Text = "跳过版本";
            _btnSkip.Click += (_, _) => { UserChoice = 2; Close(); };
            _btnCancel.Text = "取消";
            _btnCancel.Click += (_, _) => { UserChoice = 3; Close(); };
            _btnUpdate.Text = "立即更新";
            _btnUpdate.Primary = true;
            _btnUpdate.Click += OnUpdateClick;
            _bar.Controls.AddRange(new Control[] { _btnSkip, _btnCancel, _btnUpdate });
            _bar.Resize += (_, _) => LayoutBar();

            // 停靠顺序：同一条边上后加入的靠外，所以正文先加、页眉页脚后加
            Controls.Add(_body);
            Controls.Add(_progressRow);
            Controls.Add(_topBar);
            Controls.Add(_bar);

            AcceptButton = _btnUpdate;

            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
        }

        private void LayoutTopBar()
        {
            _lblTitle.SetBounds(ThemeMetrics.Margin, 0,
                _topBar.Width - ThemeMetrics.Margin * 2, ThemeMetrics.BarHeight - 1);
        }

        private void LayoutBody()
        {
            // 正文区自己排，留白比页眉的边距多留一点，读起来不至于贴边
            int m = ThemeMetrics.Margin + 6;
            int x = m, y = ThemeMetrics.Group;
            int w = _body.ClientSize.Width - m * 2;

            _lblMsg.SetBounds(x, y, w, ThemeMetrics.RowH * 2);
            y += _lblMsg.Height + ThemeMetrics.Group;

            if (_rtbReleaseNotes != null)
            {
                _lblNotesTitle.Location = new Point(x, y);
                y = _lblNotesTitle.Bottom + ThemeMetrics.Gap;
                var host = (Control)_rtbReleaseNotes.Parent!;
                host.SetBounds(x, y, w,
                    Math.Max(ThemeMetrics.RowH, _body.ClientSize.Height - y - ThemeMetrics.RowPad));
            }
            else
            {
                _lblNoNotes.Location = new Point(x, y);
            }
        }

        private void LayoutProgressRow()
        {
            const int textW = 44;
            int barH = Math.Max(10, ThemeMetrics.RowH - 8);
            int y = (_progressRow.Height - barH) / 2;
            _lblProgressText.SetBounds(_progressRow.Width - textW, y, textW, barH);
            _progressBar.SetBounds(0, y,
                Math.Max(1, _progressRow.Width - textW - ThemeMetrics.Gap), barH);
        }

        private void LayoutBar()
            => DialogButtonBar.LayoutRight(_bar, _btnSkip, _btnCancel, _btnUpdate);

        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _topBar.BackColor = p.Chrome;
            _bar.BackColor = p.Chrome;
            _topDivider.BackColor = p.Divider;
            _progressRow.BackColor = p.Window;

            _lblTitle.ForeColor = p.Text;
            _lblTitle.Font = Theme.Head;
            _lblMsg.ForeColor = p.Text;
            _lblNotesTitle.ForeColor = p.Muted;
            _lblNotesTitle.Font = Theme.Head;
            _lblNoNotes.ForeColor = p.Muted;
            _lblProgressText.ForeColor = p.Muted;

            if (_rtbReleaseNotes != null)
            {
                _rtbReleaseNotes.BackColor = p.Surface;
                _rtbReleaseNotes.ForeColor = p.Text;
            }

            _btnSkip.BackColor = p.Chrome;
            _btnCancel.BackColor = p.Chrome;
            _btnUpdate.BackColor = p.Chrome;

            LayoutTopBar();
            LayoutBody();
            LayoutProgressRow();
            LayoutBar();
            Invalidate(true);
        }

        private async void OnUpdateClick(object? sender, EventArgs e)
        {
            UserChoice = 1;

            // 禁用按钮，显示进度条
            _btnUpdate.Enabled = false;
            _btnSkip.Enabled = false;
            _btnCancel.Enabled = false;
            _btnUpdate.Text = "下载中...";
            _progressRow.Visible = true;
            LayoutProgressRow();
            LayoutBar();

            try
            {
                // ★ 在后台线程执行下载，通过 BeginInvoke 安全更新 UI
                //    使用 Action<int> 而非 Progress<T>，避免 Progress 内部的回调合并机制
                //    导致进度条跳过中间值（只显示 0% → 100%）
                await Task.Run(async () =>
                {
                    await UpdateHelper.DoUpdateAsync(_result.RemoteVersion, _result.RemoteMd5, p =>
                    {
                        // 此回调在线程池线程上执行，通过 BeginInvoke 安全更新 UI
                        BeginInvoke(new Action(() =>
                        {
                            if (!IsDisposed)
                            {
                                _progressBar.Value = p;
                                _lblProgressText.Text = $"{p}%";
                            }
                        }));
                    });
                });

                // DoUpdateAsync 会执行 Environment.Exit(0)，不会执行到这里
            }
            catch (Exception ex)
            {
                _progressRow.Visible = false;
                _lblProgressText.Text = "";

                AlertForm.Show(null, $"更新失败：{ex.Message}", "更新错误", AlertType.Error, exception: ex);

                // 失败后这颗按钮的作用变成关掉窗口，实心留给它，不用再改层级
                _btnUpdate.Text = "关闭";
                _btnUpdate.Enabled = true;
                _btnUpdate.Click -= OnUpdateClick;
                _btnUpdate.Click += (_, _) => Close();
                AcceptButton = _btnUpdate;
                LayoutBar();
            }
        }
    }
}
