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
        private readonly bool _isDark;

        /// <summary>用户选择：0=未选, 1=立即更新, 2=跳过版本, 3=取消</summary>
        public int UserChoice { get; private set; }

        // 控件
        private readonly Label _lblMsg;
        private readonly Button _btnUpdate;
        private readonly Button _btnSkip;
        private readonly Button _btnCancel;
        private readonly Label _lblProgressText;
        private readonly ProgressBar _progressBar;

        public UpdateNotifyForm(CheckResult result, bool isDark = false)
        {
            _result = result;
            _isDark = isDark;
            UserChoice = 0;

            var lightBg = Color.FromArgb(235, 248, 255);
            var darkBg = Color.FromArgb(30, 30, 30);
            var lightBtnBg = Color.FromArgb(66, 165, 245);
            var darkBtnBg = Color.FromArgb(0, 122, 204);

            Text = "发现更新";
            ClientSize = new Size(420, 200);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;
            TopMost = true;
            Font = new Font("Microsoft YaHei UI", 10f);
            BackColor = _isDark ? darkBg : lightBg;

            var msg = $"发现新版本 v{result.RemoteVersion}！\n\n" +
                      $"当前版本: v{result.CurrentVersion}\n" +
                      $"最新版本: v{result.RemoteVersion}";

            _lblMsg = new Label
            {
                Text = msg,
                AutoSize = false,
                Size = new Size(380, 80),
                Location = new Point(20, 20),
                Font = new Font("Microsoft YaHei UI", 10f),
                ForeColor = _isDark ? Color.FromArgb(241, 241, 241) : Color.FromArgb(20, 60, 100),
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _btnUpdate = new Button
            {
                Text = "立即更新",
                Size = new Size(110, 34),
                Location = new Point(20, 120),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = _isDark ? darkBtnBg : Color.FromArgb(76, 175, 80),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            _btnUpdate.FlatAppearance.BorderSize = 0;

            _btnSkip = new Button
            {
                Text = "跳过版本",
                Size = new Size(110, 34),
                Location = new Point(150, 120),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = _isDark ? darkBtnBg : Color.FromArgb(255, 152, 0),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            _btnSkip.FlatAppearance.BorderSize = 0;

            _btnCancel = new Button
            {
                Text = "取消",
                Size = new Size(110, 34),
                Location = new Point(280, 120),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = _isDark ? darkBtnBg : Color.FromArgb(158, 158, 158),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            _btnCancel.FlatAppearance.BorderSize = 0;

            // 进度条（在立即更新按钮下方，默认隐藏）
            _progressBar = new ProgressBar
            {
                Location = new Point(20, 162),
                Size = new Size(300, 18),
                Visible = false,
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };

            // 进度百分比
            _lblProgressText = new Label
            {
                Text = "",
                AutoSize = false,
                Size = new Size(40, 18),
                Location = new Point(324, 162),
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = _isDark ? Color.FromArgb(200, 200, 200) : Color.FromArgb(90, 130, 170),
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft,
                Visible = false
            };

            Controls.Add(_lblMsg);
            Controls.Add(_btnUpdate);
            Controls.Add(_btnSkip);
            Controls.Add(_btnCancel);
            Controls.Add(_progressBar);
            Controls.Add(_lblProgressText);

            AcceptButton = _btnUpdate;

            _btnUpdate.Click += OnUpdateClick;
            _btnSkip.Click += (_, _) => { UserChoice = 2; Close(); };
            _btnCancel.Click += (_, _) => { UserChoice = 3; Close(); };
        }

        private async void OnUpdateClick(object? sender, EventArgs e)
        {
            UserChoice = 1;

            // 禁用按钮，显示进度条在立即更新按钮下方
            _btnUpdate.Enabled = false;
            _btnSkip.Enabled = false;
            _btnCancel.Enabled = false;
            _btnUpdate.Text = "下载中...";
            _progressBar.Visible = true;
            _lblProgressText.Visible = true;

            var progress = new Progress<int>(p =>
            {
                if (!IsDisposed)
                {
                    _progressBar.Value = p;
                    _lblProgressText.Text = $"{p}%";
                }
            });

            try
            {
                await UpdateHelper.DoUpdateAsync(_result.RemoteVersion, _result.RemoteMd5, progress);
                // DoUpdateAsync 会执行 Environment.Exit(0)，不会执行到这里
            }
            catch (Exception ex)
            {
                _progressBar.Visible = false;
                _lblProgressText.Text = "";

                MessageBox.Show($"更新失败：{ex.Message}", "更新错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                // 恢复关闭功能
                _btnUpdate.Text = "关闭";
                _btnUpdate.Enabled = true;
                _btnUpdate.Click -= OnUpdateClick;
                _btnUpdate.Click += (_, _) => Close();
                AcceptButton = _btnUpdate;
            }
        }
    }
}
