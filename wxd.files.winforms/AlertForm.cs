using System;
using System.Collections.Generic;
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
        /// <summary>一颗底部按钮的声明：文案、点完返回什么、是不是主操作。</summary>
        private readonly record struct ButtonSpec(string Text, DialogResult Result, bool Primary);

        private readonly string _message;
        private readonly AlertType _type;
        private readonly RichTextBox _rtbMessage;
        private readonly Panel _bar;
        private readonly List<ThemeButton> _barButtons = new();
        private readonly ThemeButton? _btnCopy;
        private readonly ThemeButton _defaultButton;
        private readonly Label _lblIcon;
        private readonly Label _lblTitle;
        private readonly Panel _sep;

        private AlertForm(string title, string message, AlertType type, bool showCopy,
            Exception? exception, ButtonSpec[]? buttons)
        {
            _message = message;
            _type = type;

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

            _lblIcon = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 20f, FontStyle.Regular),
                BackColor = Color.Transparent
            };
            _lblTitle = new Label
            {
                Text = title,
                AutoSize = false,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.Transparent
            };
            _sep = new Panel { Dock = DockStyle.Bottom, Height = 1 };

            var header = new Panel { Dock = DockStyle.Top, Height = 58 };
            header.Controls.Add(_sep);
            header.Controls.Add(_lblTitle);
            header.Controls.Add(_lblIcon);
            header.Resize += (_, _) => LayoutHeader();

            _rtbMessage = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                WordWrap = true,
                DetectUrls = false,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Text = _message
            };
            _rtbMessage.Select(0, 0); // 不选中任何文本

            // 正文要缩进，但页眉页脚的分隔线要铺满整宽，所以只给正文套一层留白容器
            var body = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(ThemeMetrics.Margin + 6, ThemeMetrics.Gap,
                                     ThemeMetrics.Margin + 6, ThemeMetrics.RowPad)
            };
            body.Controls.Add(_rtbMessage);

            // 底部按钮条：一条分隔线 + 右对齐的主题按钮，和其他对话框共用同一套排布
            _bar = DialogButtonBar.Create();
            var specs = buttons ?? new[] { new ButtonSpec("确定", DialogResult.Yes, true) };
            if (showCopy)
            {
                _btnCopy = DialogButtonBar.Make("复制", BtnCopy_Click);
                _barButtons.Add(_btnCopy);
            }
            ThemeButton? primary = null;
            foreach (var s in specs)
            {
                var b = DialogButtonBar.Make(s.Text, null, s.Primary);
                b.DialogResult = s.Result;
                if (s.Primary && primary == null) primary = b;
                _barButtons.Add(b);
            }
            _defaultButton = primary ?? _barButtons[^1];
            _bar.Controls.AddRange(_barButtons.ToArray());
            LayoutBar();
            _bar.Resize += (_, _) => LayoutBar();

            // 停靠顺序：后加入的靠外，所以填充区先加、页眉页脚后加
            Controls.Add(body);
            Controls.Add(header);
            Controls.Add(_bar);

            AcceptButton = _defaultButton;
            // Esc 必须落在"放弃"那颗。LayoutRight 是数组末尾靠右，主按钮（确定/Yes）在最右，
            // 原来取 _barButtons[^1] 等于把 Esc 变成确认，删除笔记、强制重装这类会被一键误触发。
            // 单按钮时没有放弃项，Esc 就还是确定本身。
            ThemeButton? esc = null;
            foreach (var s in specs)
            {
                if (s.Result != DialogResult.Cancel && s.Result != DialogResult.No) continue;
                esc = _barButtons.FirstOrDefault(b => b.DialogResult == s.Result);
                if (esc != null) break;
            }
            CancelButton = esc ?? _defaultButton;

            ApplyTheme();
            // 独立弹窗，主窗翻明暗时遍历不到，只能自己订阅
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
        }

        private void LayoutBar() => DialogButtonBar.LayoutRight(_bar, _barButtons.ToArray());

        /// <summary>标题从图标右侧开始，宽度跟着页眉走。</summary>
        private void LayoutHeader()
            => _lblTitle.SetBounds(58, 16, Math.Max(40, _sep.Width - 66), 30);

        /// <summary>取色全部走当前皮肤令牌；换皮肤或翻明暗时由 Theme.Changed 重新走一遍。</summary>
        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _bar.BackColor = p.Chrome;

            _lblIcon.Text = _type switch
            {
                AlertType.Warning => "\u26A0",     // ⚠
                AlertType.Error => "\u2716",       // ✖
                _ => "\u2139"                      // ℹ
            };
            _lblIcon.ForeColor = _type switch
            {
                AlertType.Warning => Theme.Signal(Color.FromArgb(255, 160, 0)),
                AlertType.Error => p.Danger,
                _ => p.Accent
            };
            _lblTitle.ForeColor = p.Text;
            // 用 Theme 持有的字体实例，别在这里 new：每次切皮肤 new 一个就漏一个 GDI 字体句柄
            _lblTitle.Font = Theme.TitleBold;
            _sep.BackColor = p.Divider;

            _rtbMessage.BackColor = p.Surface;
            _rtbMessage.ForeColor = p.Text;
            _rtbMessage.Font = Theme.List;

            foreach (var b in _barButtons)
            {
                b.BackColor = p.Chrome;
                b.Font = Theme.Body;
            }

            LayoutHeader();
            LayoutBar();
            Invalidate(true);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // 自适应高度：正文有多长，窗体就撑多高
            var textHeight = _rtbMessage.GetPositionFromCharIndex(_rtbMessage.TextLength).Y + 30;
            var idealHeight = Math.Max(200, Math.Min(600, textHeight + 130));
            if (ClientSize.Height != idealHeight)
                ClientSize = new Size(ClientSize.Width, idealHeight);
            _defaultButton.Focus();
        }

        private void BtnCopy_Click(object? sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(_message);
                _btnCopy!.Text = "已复制";
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
        /// <param name="exception">可选异常对象，不为 null 时将输出堆栈到弹窗和输出窗口</param>
        public static DialogResult Show(IWin32Window owner, string message, string title,
            AlertType type = AlertType.Information, bool showCopy = false, Exception? exception = null)
        {
            using var form = new AlertForm(title, message, type, showCopy, exception, null);
            return form.ShowDialog(owner);
        }

        /// <summary>显示带"确定"/"取消"按钮的确认弹窗，确定返回 <see cref="DialogResult.Yes"/>，取消返回 <see cref="DialogResult.No"/></summary>
        public static DialogResult ShowConfirm(IWin32Window owner, string message, string title,
            AlertType type = AlertType.Warning)
        {
            var specs = new[]
            {
                new ButtonSpec("取消", DialogResult.No, false),
                new ButtonSpec("确定", DialogResult.Yes, true)
            };
            using var form = new AlertForm(title, message, type, false, null, specs);
            return form.ShowDialog(owner);
        }

        /// <summary>显示带三个按钮的确认弹窗，三个按钮分别返回 <see cref="DialogResult.Yes"/>、<see cref="DialogResult.No"/>、<see cref="DialogResult.Cancel"/></summary>
        public static DialogResult ShowConfirmThreeButton(IWin32Window owner, string message, string title,
            string yesText = "确定", string noText = "取消", string cancelText = "继续编辑",
            AlertType type = AlertType.Warning)
        {
            // 从左到右：次要的排前面，主操作永远落在最右
            var specs = new[]
            {
                new ButtonSpec(cancelText, DialogResult.Cancel, false),
                new ButtonSpec(noText, DialogResult.No, false),
                new ButtonSpec(yesText, DialogResult.Yes, true)
            };
            using var form = new AlertForm(title, message, type, false, null, specs);
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
