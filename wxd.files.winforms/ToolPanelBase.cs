using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 工具集页签的公共骨架：顶部工具条 + 底部状态栏 + 主题刷新。
    /// 子类先放内容控件，再调用 <see cref="FinishChrome"/> 让顶部/底部贴边。
    /// </summary>
    internal abstract class ToolPanelBase : UserControl
    {
        private readonly List<ThemeButton> _buttons = new();

        protected readonly FlowLayoutPanel Top = new();
        protected readonly Label Status = new();

        protected ToolPanelBase()
        {
            Dock = DockStyle.Fill;
            BackColor = Theme.P.Window;
            Font = Theme.Body;

            Top.Dock = DockStyle.Top;
            Top.Height = 46;
            Top.FlowDirection = FlowDirection.LeftToRight;
            Top.WrapContents = false;
            Top.AutoScroll = true;
            Top.Padding = new Padding(ThemeMetrics.Margin, 9, 0, 0);
            Top.BackColor = Theme.P.Chrome;

            Status.Dock = DockStyle.Bottom;
            Status.Height = 26;
            Status.TextAlign = ContentAlignment.MiddleLeft;
            Status.Padding = new Padding(ThemeMetrics.Margin, 0, 0, 0);
            Status.Font = Theme.Caption;
            Status.BackColor = Theme.P.Window;
            Status.ForeColor = Theme.P.Muted;

            // 不能在这里调用 ApplyTheme：ApplyThemeCore 是虚方法，派生类字段此时还没初始化
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
        }

        /// <summary>按 Top → 内容 → Bottom 的贴边顺序挂载；必须在子类构造函数末尾调用一次。</summary>
        protected void FinishChrome(Control content)
        {
            Controls.Add(content);
            Controls.Add(Top);
            Controls.Add(Status);
        }

        protected ThemeButton AddButton(string text, EventHandler onClick, bool primary = false)
        {
            var button = new ThemeButton
            {
                Text = text,
                Font = Theme.Body,
                Primary = primary,
                Margin = new Padding(0, 0, ThemeMetrics.Gap, 0),
                Size = new Size(ThemeButton.FitWidth(text), ThemeMetrics.RowH),
                BackColor = Theme.P.Chrome
            };
            button.Click += onClick;
            _buttons.Add(button);
            Top.Controls.Add(button);
            return button;
        }

        protected ThemeField AddField(string placeholder, int width)
        {
            var field = new ThemeField
            {
                Size = new Size(width, ThemeMetrics.RowH),
                Margin = new Padding(0, 0, ThemeMetrics.Gap, 0),
                Placeholder = placeholder
            };
            Top.Controls.Add(field);
            return field;
        }

        protected ThemeCheckBox AddCheckBox(string text)
        {
            var box = new ThemeCheckBox
            {
                Text = text,
                Font = Theme.Body,
                AutoSize = true,
                Margin = new Padding(0, 5, ThemeMetrics.Gap, 0)
            };
            Top.Controls.Add(box);
            return box;
        }

        private void ApplyTheme()
        {
            if (IsDisposed) return;
            var p = Theme.P;

            Font = Theme.Body;
            BackColor = p.Window;
            Top.BackColor = p.Chrome;
            Status.BackColor = p.Window;
            Status.ForeColor = p.Muted;
            Status.Font = Theme.Caption;

            foreach (var button in _buttons)
            {
                button.Font = Theme.Body;
                button.BackColor = p.Chrome;
                button.Invalidate();
            }

            ApplyThemeCore();
            Invalidate(true);
        }

        protected virtual void ApplyThemeCore() { }

        protected void SetStatus(string text, bool error = false)
        {
            Status.Text = text;
            Status.ForeColor = error ? Theme.P.Danger : Theme.P.Muted;
        }

        protected static TextBox MakeEditor(bool readOnly = false)
        {
            return new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsTab = true,
                BorderStyle = BorderStyle.None,
                Font = Theme.Mono,
                ReadOnly = readOnly,
                HideSelection = false,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text
            };
        }

        protected static void ApplyEditorTheme(TextBox box)
        {
            box.Font = Theme.Mono;
            box.BackColor = Theme.P.Surface;
            box.ForeColor = Theme.P.Text;
        }

        protected void CopyToClipboard(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                SetStatus("没有可复制的内容", error: true);
                return;
            }

            try
            {
                Clipboard.SetText(text);
                SetStatus($"已复制到剪贴板（{text.Length} 字符）");
            }
            catch (Exception ex)
            {
                SetStatus("复制失败：" + ex.Message, error: true);
            }
        }
    }
}
