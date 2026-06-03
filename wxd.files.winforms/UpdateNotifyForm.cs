using System;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>更新提示窗体：立即更新 / 跳过版本 / 取消</summary>
    public class UpdateNotifyForm : Form
    {
        private readonly CheckResult _result;

        /// <summary>用户选择：0=未选, 1=立即更新, 2=跳过版本, 3=取消</summary>
        public int UserChoice { get; private set; }

        public UpdateNotifyForm(CheckResult result, bool isDark = false)
        {
            _result = result;
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
            BackColor = isDark ? darkBg : lightBg;

            var msg = $"发现新版本 v{result.RemoteVersion}！\n\n" +
                      $"当前版本: v{result.CurrentVersion}\n" +
                      $"最新版本: v{result.RemoteVersion}";

            var lbl = new Label
            {
                Text = msg,
                AutoSize = false,
                Size = new Size(380, 80),
                Location = new Point(20, 20),
                Font = new Font("Microsoft YaHei UI", 10f),
                ForeColor = isDark ? Color.FromArgb(241, 241, 241) : Color.FromArgb(20, 60, 100),
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var btnUpdate = new Button
            {
                Text = "立即更新",
                Size = new Size(110, 34),
                Location = new Point(20, 120),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = isDark ? darkBtnBg : Color.FromArgb(76, 175, 80),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Yes
            };
            btnUpdate.FlatAppearance.BorderSize = 0;

            var btnSkip = new Button
            {
                Text = "跳过版本",
                Size = new Size(110, 34),
                Location = new Point(150, 120),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = isDark ? darkBtnBg : Color.FromArgb(255, 152, 0),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Ignore
            };
            btnSkip.FlatAppearance.BorderSize = 0;

            var btnCancel = new Button
            {
                Text = "取消",
                Size = new Size(110, 34),
                Location = new Point(280, 120),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = isDark ? darkBtnBg : Color.FromArgb(158, 158, 158),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;

            Controls.Add(lbl);
            Controls.Add(btnUpdate);
            Controls.Add(btnSkip);
            Controls.Add(btnCancel);

            AcceptButton = btnUpdate;

            btnUpdate.Click += (_, _) => { UserChoice = 1; Close(); };
            btnSkip.Click += (_, _) => { UserChoice = 2; Close(); };
            btnCancel.Click += (_, _) => { UserChoice = 3; Close(); };
        }
    }
}
