using System;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>弹窗类型</summary>
    public enum AlertType
    {
        Information,
        Warning,
        Error
    }

    /// <summary>自定义弹窗，支持复制按钮（便于复制异常信息）</summary>
    public class AlertForm : Form
    {
        private readonly string _message;
        private readonly AlertType _type;
        private readonly bool _showCopy;
        private readonly Exception? _exception;
        private readonly RichTextBox _rtbMessage;
        private readonly Button _btnCopy;
        private readonly Button _btnOk;

        // 配色
        private static readonly Color _colorBg = Color.FromArgb(245, 250, 255);
        private static readonly Color _colorPrimary = Color.FromArgb(66, 165, 245);
        private static readonly Color _colorWarning = Color.FromArgb(255, 160, 0);
        private static readonly Color _colorError = Color.FromArgb(239, 68, 68);
        private static readonly Color _colorTitleFg = Color.FromArgb(20, 60, 100);
        private static readonly Color _colorBtnBorder = Color.FromArgb(30, 136, 229);

        private AlertForm(string title, string message, AlertType type, bool showCopy, Exception? exception)
        {
            _message = message;
            _type = type;
            _showCopy = showCopy;
            _exception = exception;

            // 如果 exception 不为 null，将堆栈信息追加到消息中显示在弹窗，同时输出到输出窗口
            if (exception != null)
            {
                var sb = new StringBuilder();
                sb.AppendLine("═══════════════════════════════════════");
                sb.AppendLine("[AlertForm] 异常标题: " + title);
                sb.AppendLine("[AlertForm] 异常消息: " + exception.Message);
                sb.AppendLine("[AlertForm] 异常类型: " + exception.GetType().FullName);
                sb.AppendLine("[AlertForm] 异常堆栈:");
                sb.AppendLine(exception.StackTrace ?? "(无堆栈信息)");
                if (exception.InnerException != null)
                {
                    sb.AppendLine("[AlertForm] 内部异常: " + exception.InnerException);
                    sb.AppendLine("[AlertForm] 内部异常堆栈:");
                    sb.AppendLine(exception.InnerException.StackTrace ?? "(无堆栈信息)");
                }
                sb.AppendLine("═══════════════════════════════════════");
                Debug.Write(sb.ToString());

                // 将堆栈信息追加到弹窗消息文本中，使用方传入了 exception 就自动显示堆栈
                _message = message + Environment.NewLine + Environment.NewLine + sb.ToString();
            }

            Text = title;
            ClientSize = new Size(500, 260);
            MinimumSize = new Size(400, 200);
            MaximumSize = new Size(800, 600);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = _colorBg;

            // 图标 + 标题
            var iconText = type switch
            {
                AlertType.Warning => "\u26A0",     // ⚠
                AlertType.Error => "\u2716",       // ✖
                _ => "\u2139"                      // ℹ
            };
            var iconColor = type switch
            {
                AlertType.Warning => _colorWarning,
                AlertType.Error => _colorError,
                _ => _colorPrimary
            };

            var lblIcon = new Label
            {
                Text = iconText,
                AutoSize = true,
                Font = new Font("Segoe UI", 20f, FontStyle.Regular),
                ForeColor = iconColor,
                Location = new Point(18, 16),
                BackColor = Color.Transparent
            };

            var lblTitle = new Label
            {
                Text = title,
                AutoSize = false,
                Height = 30,
                Location = new Point(58, 16),
                Width = ClientSize.Width - 80,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei UI", 12f, FontStyle.Bold),
                ForeColor = _colorTitleFg,
                BackColor = Color.Transparent
            };

            // 分隔线
            var sep = new Panel
            {
                Location = new Point(16, 50),
                Size = new Size(ClientSize.Width - 32, 1),
                BackColor = Color.FromArgb(173, 209, 237)
            };

            // 消息文本框（只读，可选择复制）
            _rtbMessage = new RichTextBox
            {
                Location = new Point(16, 58),
                Size = new Size(ClientSize.Width - 32, ClientSize.Height - 120),
                Font = new Font("Microsoft YaHei UI", 9.5f),
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(30, 60, 100),
                WordWrap = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                DetectUrls = false,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            _rtbMessage.Text = message;
            _rtbMessage.Select(0, 0); // 不选中任何文本

            // 底部按钮
            var bottomPanel = new Panel
            {
                Height = 44,
                Dock = DockStyle.Bottom,
                BackColor = Color.Transparent,
                Padding = new Padding(10, 6, 10, 6)
            };

            _btnOk = new Button
            {
                Text = "确定",
                Size = new Size(90, 30),
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = _colorPrimary,
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Yes
            };
            _btnOk.FlatAppearance.BorderSize = 1;
            _btnOk.FlatAppearance.BorderColor = _colorBtnBorder;
            _btnOk.FlatAppearance.MouseOverBackColor = Color.FromArgb(50, 145, 235);
            _btnOk.Location = new Point(bottomPanel.Width - _btnOk.Width - 10, 6);

            bottomPanel.Controls.Add(_btnOk);

            if (_showCopy)
            {
                _btnCopy = new Button
                {
                    Text = "复制",
                    Size = new Size(90, 30),
                    Anchor = AnchorStyles.Right,
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(120, 144, 156),
                    ForeColor = Color.White,
                    Font = new Font("Microsoft YaHei UI", 9f),
                    Cursor = Cursors.Hand
                };
                _btnCopy.FlatAppearance.BorderSize = 1;
                _btnCopy.FlatAppearance.BorderColor = Color.FromArgb(100, 124, 136);
                _btnCopy.FlatAppearance.MouseOverBackColor = Color.FromArgb(100, 124, 136);
                _btnCopy.Location = new Point(_btnOk.Left - _btnCopy.Width - 8, 6);
                _btnCopy.Click += BtnCopy_Click;

                bottomPanel.Controls.Add(_btnCopy);
            }

            Controls.Add(_rtbMessage);
            Controls.Add(sep);
            Controls.Add(lblTitle);
            Controls.Add(lblIcon);
            Controls.Add(bottomPanel);

            AcceptButton = _btnOk;
            CancelButton = _btnOk;

            // 自适应宽度
            lblTitle.Width = ClientSize.Width - 80;
            _rtbMessage.Width = ClientSize.Width - 32;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // 自动调整高度：根据内容高度
            var textHeight = _rtbMessage.GetPositionFromCharIndex(_rtbMessage.TextLength).Y + 30;
            var idealHeight = Math.Max(200, Math.Min(600, textHeight + 130));
            if (ClientSize.Height != idealHeight)
            {
                ClientSize = new Size(ClientSize.Width, idealHeight);
            }
            _btnOk.Focus();
        }

        private void BtnCopy_Click(object? sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(_message);
                _btnCopy.Text = "已复制";
                _btnCopy.Enabled = false;
                var timer = new System.Windows.Forms.Timer { Interval = 1500 };
                timer.Tick += (s, args) =>
                {
                    _btnCopy.Text = "复制";
                    _btnCopy.Enabled = true;
                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();
            }
            catch
            {
                // 剪贴板可能被占用
            }
        }

        /// <summary>显示自定义弹窗</summary>
        /// <param name="owner">父窗口</param>
        /// <param name="message">消息内容</param>
        /// <param name="title">标题</param>
        /// <param name="type">弹窗类型（决定图标颜色）</param>
        /// <param name="showCopy">是否显示"复制"按钮，方便复制异常信息</param>
        /// <param name="exception">可选异常对象，不为 null 时将输出堆栈到输出窗口</param>
        public static DialogResult Show(IWin32Window owner, string message, string title,
            AlertType type = AlertType.Information, bool showCopy = false, Exception? exception = null)
        {
            using var form = new AlertForm(title, message, type, showCopy, exception);
            return form.ShowDialog(owner);
        }

        /// <summary>显示带"确定"/"取消"按钮的确认弹窗，确定返回 <see cref="DialogResult.Yes"/>，取消返回 <see cref="DialogResult.No"/></summary>
        public static DialogResult ShowConfirm(IWin32Window owner, string message, string title,
            AlertType type = AlertType.Warning)
        {
            using var form = new AlertForm(title, message, type, showCopy: false, exception: null);
            // 在底部面板添加"取消"按钮
            var cancelBtn = new Button
            {
                Text = "取消",
                Size = new Size(90, 30),
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(180, 190, 200),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.No
            };
            cancelBtn.FlatAppearance.BorderSize = 1;
            cancelBtn.FlatAppearance.BorderColor = Color.FromArgb(160, 170, 180);
            cancelBtn.FlatAppearance.MouseOverBackColor = Color.FromArgb(160, 170, 180);

            // 获取底部面板（最后一个控件）
            var bottomPanel = form.Controls[^1] as Panel;
            if (bottomPanel != null)
            {
                // 调整确定按钮位置
                form._btnOk.Location = new Point(bottomPanel.Width - form._btnOk.Width - 10, 6);
                cancelBtn.Location = new Point(form._btnOk.Left - cancelBtn.Width - 8, 6);
                bottomPanel.Controls.Add(cancelBtn);
            }

            form.CancelButton = cancelBtn;
            return form.ShowDialog(owner);
        }

        /// <summary>显示带三个按钮的确认弹窗，三个按钮分别返回 <see cref="DialogResult.Yes"/>、<see cref="DialogResult.No"/>、<see cref="DialogResult.Cancel"/></summary>
        public static DialogResult ShowConfirmThreeButton(IWin32Window owner, string message, string title,
            string yesText = "确定", string noText = "取消", string cancelText = "继续编辑",
            AlertType type = AlertType.Warning)
        {
            using var form = new AlertForm(title, message, type, showCopy: false, exception: null);

            var btnYes = new Button
            {
                Text = yesText,
                Size = new Size(110, 30),
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Yes
            };
            btnYes.FlatAppearance.BorderSize = 1;
            btnYes.FlatAppearance.BorderColor = Color.FromArgb(30, 136, 229);

            var btnNo = new Button
            {
                Text = noText,
                Size = new Size(110, 30),
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(180, 190, 200),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.No
            };
            btnNo.FlatAppearance.BorderSize = 1;
            btnNo.FlatAppearance.BorderColor = Color.FromArgb(160, 170, 180);

            var btnCancel = new Button
            {
                Text = cancelText,
                Size = new Size(100, 30),
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(210, 215, 225),
                ForeColor = Color.FromArgb(60, 70, 90),
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 1;
            btnCancel.FlatAppearance.BorderColor = Color.FromArgb(190, 195, 205);

            var bottomPanel = form.Controls[^1] as Panel;
            if (bottomPanel != null)
            {
                form._btnOk.Visible = false;
                btnYes.Location = new Point(bottomPanel.Width - btnYes.Width - 10, 6);
                btnNo.Location = new Point(btnYes.Left - btnNo.Width - 8, 6);
                btnCancel.Location = new Point(btnNo.Left - btnCancel.Width - 8, 6);
                bottomPanel.Controls.Add(btnYes);
                bottomPanel.Controls.Add(btnNo);
                bottomPanel.Controls.Add(btnCancel);
            }

            form.AcceptButton = btnYes;
            form.CancelButton = btnCancel;
            return form.ShowDialog(owner);
        }

        /// <summary>显示错误弹窗（自动显示复制按钮）</summary>
        public static DialogResult ShowError(IWin32Window owner, string message, string title = "错误", Exception? exception = null)
        {
            return Show(owner, message, title, AlertType.Error, showCopy: true, exception: exception);
        }

        /// <summary>显示异常弹窗（自动包含异常信息 + 复制按钮 + 输出堆栈）</summary>
        public static DialogResult ShowException(IWin32Window owner, Exception ex, string title = "异常")
        {
            var sb = new StringBuilder();
            sb.AppendLine(ex.Message);
            sb.AppendLine();
            sb.AppendLine("--- 详细信息 ---");
            sb.AppendLine(ex.ToString());
            return Show(owner, sb.ToString(), title, AlertType.Error, showCopy: true, exception: ex);
        }
    }
}
