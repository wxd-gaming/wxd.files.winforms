using System;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 笔记列表项控件（替代 ListView 的 OwnerDraw 方案）。
    /// 由两个 Label 组成：标题（粗体） + 更新时间（小字灰色）。
    /// 自带悬停/选中背景切换，点击时触发 Selected 事件通知父窗体。
    /// </summary>
    public class NoteListItemControl : UserControl
    {
        // ==================== 字段 ====================
        private readonly Label _lblTitle;
        private readonly Label _lblTime;
        private bool _isSelected;
        private bool _isHover;

        /// <summary>当前绑定的笔记</summary>
        public NoteItem Note { get; }

        /// <summary>被点击/选中时触发，参数为自身</summary>
        public event Action<NoteListItemControl>? Selected;

        public NoteListItemControl(NoteItem note)
        {
            Note = note ?? throw new ArgumentNullException(nameof(note));

            // 控件本身：固定高度，Dock=Top 在可滚动 Panel 中自动从顶部堆叠并占满宽度
            Height = 44;
            Dock = DockStyle.Top;
            Margin = new Padding(0);
            Padding = new Padding(8, 5, 8, 4);
            Cursor = Cursors.Hand;
            DoubleBuffered = true;

            // 标题 Label
            _lblTitle = new Label
            {
                AutoSize = false,
                Location = new Point(8, 5),
                Size = new Size(Width - 16, 20),
                Font = Theme.Head,
                Text = string.IsNullOrEmpty(note.Title) ? "(无标题)" : note.Title,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 时间 Label
            _lblTime = new Label
            {
                AutoSize = false,
                Location = new Point(8, 27),
                Size = new Size(Width - 16, 14),
                Font = Theme.Caption,
                Text = note.UpdatedAt.ToString("yyyy-MM-dd HH:mm:ss"),
                TextAlign = ContentAlignment.MiddleLeft
            };

            Controls.Add(_lblTitle);
            Controls.Add(_lblTime);

            // 子控件鼠标事件冒泡到本控件
            foreach (Control c in Controls)
            {
                c.MouseEnter += (s, e) => OnMouseEnter(EventArgs.Empty);
                c.MouseLeave += (s, e) => OnMouseLeave(EventArgs.Empty);
                c.MouseDown += (s, e) => OnMouseDown(e);
                c.Click += (s, e) => OnClick(e);
            }

            ApplyColors();
            Theme.Changed += ApplyColors;
            Disposed += (_, _) => Theme.Changed -= ApplyColors;
        }

        /// <summary>更新标题显示</summary>
        public void UpdateTitle()
        {
            _lblTitle.Text = string.IsNullOrEmpty(Note.Title) ? "(无标题)" : Note.Title;
        }

        /// <summary>更新时间显示</summary>
        public void UpdateTime()
        {
            _lblTime.Text = Note.UpdatedAt.ToString("yyyy-MM-dd HH:mm:ss");
        }

        /// <summary>设置选中状态（由父窗体管理单选）</summary>
        public void SetSelected(bool selected)
        {
            if (_isSelected == selected) return;
            _isSelected = selected;
            ApplyColors();
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            _isHover = true;
            if (!_isSelected) ApplyColors();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _isHover = false;
            if (!_isSelected) ApplyColors();
        }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            SetSelected(true);
            Selected?.Invoke(this);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            // 构造函数早期 / 释放后可能为 null
            if (_lblTitle == null || _lblTime == null || IsDisposed) return;
            // 宽度变化时同步 Label 宽度
            var w = Width - 16;
            if (w > 0)
            {
                _lblTitle.Width = w;
                _lblTime.Width = w;
            }
        }

        private void ApplyColors()
        {
            var p = Theme.P;
            Color bg = _isSelected ? p.Selection : _isHover ? p.Hover : p.Surface;

            BackColor = bg;
            _lblTitle.BackColor = bg;
            _lblTime.BackColor = bg;
            _lblTitle.ForeColor = p.Text;
            _lblTime.ForeColor = p.Muted;
        }
    }
}
