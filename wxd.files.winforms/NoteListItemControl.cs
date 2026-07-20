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
        // ==================== 配色（与 NoteForm 保持一致） ====================
        private static readonly Color _lightBg = Color.White;
        private static readonly Color _lightHover = Color.FromArgb(235, 243, 252);
        private static readonly Color _lightSel = Color.FromArgb(210, 230, 250);
        private static readonly Color _lightTitleFg = Color.FromArgb(20, 60, 100);
        private static readonly Color _lightSubFg = Color.FromArgb(120, 140, 160);

        private static readonly Color _darkBg = Color.FromArgb(45, 45, 48);
        private static readonly Color _darkHover = Color.FromArgb(62, 62, 66);
        private static readonly Color _darkSel = Color.FromArgb(81, 113, 170);
        private static readonly Color _darkTitleFg = Color.FromArgb(241, 241, 241);
        private static readonly Color _darkSubFg = Color.FromArgb(160, 160, 160);

        // ==================== 字段 ====================
        private readonly Label _lblTitle;
        private readonly Label _lblTime;
        private readonly bool _isDark;
        private bool _isSelected;
        private bool _isHover;

        /// <summary>当前绑定的笔记</summary>
        public NoteItem Note { get; }

        /// <summary>被点击/选中时触发，参数为自身</summary>
        public event Action<NoteListItemControl>? Selected;

        public NoteListItemControl(NoteItem note, bool isDark)
        {
            Note = note ?? throw new ArgumentNullException(nameof(note));
            _isDark = isDark;

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
                Font = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold),
                Text = string.IsNullOrEmpty(note.Title) ? "(无标题)" : note.Title,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 时间 Label
            _lblTime = new Label
            {
                AutoSize = false,
                Location = new Point(8, 27),
                Size = new Size(Width - 16, 14),
                Font = new Font("Microsoft YaHei UI", 7.5f),
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

            ApplyBackColor();
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
            ApplyBackColor();
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            _isHover = true;
            if (!_isSelected) ApplyBackColor();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _isHover = false;
            if (!_isSelected) ApplyBackColor();
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

        private void ApplyBackColor()
        {
            Color bg;
            if (_isSelected)
                bg = _isDark ? _darkSel : _lightSel;
            else if (_isHover)
                bg = _isDark ? _darkHover : _lightHover;
            else
                bg = _isDark ? _darkBg : _lightBg;

            BackColor = bg;
            _lblTitle.BackColor = bg;
            _lblTime.BackColor = bg;
            _lblTitle.ForeColor = _isDark ? _darkTitleFg : _lightTitleFg;
            _lblTime.ForeColor = _isDark ? _darkSubFg : _lightSubFg;
        }
    }
}
