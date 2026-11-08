using System.Drawing.Drawing2D;

namespace wxd.files.winforms
{
    /// <summary>
    /// 圆角扁平按钮：替换默认方角按钮，配色实时取当前皮肤令牌。
    /// FillOverride 用于保留用户自定义命令按钮的颜色。
    /// </summary>
    internal class ThemeButton : Button
    {
        /// <summary>主操作按钮：实底强调色</summary>
        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public bool Primary { get; set; }

        /// <summary>次要按钮：只留文字，无底无框</summary>
        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public bool Ghost { get; set; }

        /// <summary>调用方指定的底色（自定义命令按钮的颜色）</summary>
        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Color? FillOverride { get; set; }

        private bool _hover;
        private bool _down;

        public ThemeButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint
                   | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat;
            Cursor = Cursors.Hand;
            MouseEnter += (_, _) => { _hover = true; Invalidate(); };
            MouseLeave += (_, _) => { _hover = false; _down = false; Invalidate(); };
            MouseDown += (_, _) => { _down = true; Invalidate(); };
            MouseUp += (_, _) => { _down = false; Invalidate(); };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var p = Theme.P;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var face = FaceColor();
            g.Clear(face);

            var flat = Primary || FillOverride.HasValue;
            var fill = FillOverride ?? (Primary ? p.Accent : face);
            var line = flat ? fill : p.Border;
            var fore = flat
                ? (Theme.IsDark(fill) ? Color.White : p.Text)
                : Ghost ? p.Muted : p.Text;

            if (!Enabled)
            {
                fill = Theme.Mix(fill, face, 0.55f);
                line = Theme.Mix(line, face, 0.55f);
                fore = Theme.Mix(fore, face, 0.45f);
            }
            else if (_down)
            {
                fill = Theme.Mix(fill, Theme.IsDark(fill) ? Color.White : p.Text, 0.16f);
                line = fill;
            }
            else if (_hover)
            {
                fill = Theme.Mix(fill, flat ? Color.White : p.Text, flat ? 0.14f : 0.06f);
                if (!Ghost) line = Theme.Mix(line, flat ? Color.White : p.Text, 0.14f);
            }

            var rect = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            var radius = (float)Math.Min(p.Radius, Height / 2f);

            using (var path = RoundRect(rect, radius))
            {
                if (!Ghost)
                {
                    using var br = new SolidBrush(fill);
                    g.FillPath(br, path);
                }
                // 实底按钮的边框和底色同色，画了只是噪声；空心/幽灵按钮只在悬停时给反馈
                if (!flat && (Ghost ? _hover : true))
                {
                    using var pen = new Pen(line, 1f);
                    g.DrawPath(pen, path);
                }
            }

            TextRenderer.DrawText(g, Text, Font, ClientRectangle, fore,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
              | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }

        /// <summary>
        /// 取真正垫在按钮下面的底色：容器可能设了 Transparent，
        /// 直接 Clear(Transparent) 会画成黑块，要往上找到第一个不透明的父级。
        /// </summary>
        private Color FaceColor() => FaceColorOf(this);

        /// <summary>
        /// 取真正垫在按钮下面的底色：容器可能设了 Transparent，
        /// 直接 Clear(Transparent) 会画成黑块，要往上找到第一个不透明的父级。
        /// </summary>
        internal static Color FaceColorOf(Control c)
        {
            for (var p = c.Parent; p != null; p = p.Parent)
            {
                var bc = p.BackColor;
                if (bc.A > 0) return bc;
            }
            return Theme.P.Window;
        }

        /// <summary>
        /// 工具条按钮统一按文字量宽：左右各留 12px，再兜一个最小宽度。
        /// 写死 84/92 这类数字会让两个字数的按钮和四个字的按钮看着不成比例。
        /// </summary>
        internal static int FitWidth(string text, int minWidth = 52)
        {
            int w = TextRenderer.MeasureText(text, Theme.Body).Width + 24;
            return Math.Max(w, minWidth);
        }

        internal static GraphicsPath RoundRect(RectangleF r, float radius)
        {
            var path = new GraphicsPath();
            // 半径超过宽高一半时四段弧会互相越过，路径画成扭曲的交叉形，这里统一钳住
            radius = Math.Min(radius, Math.Min(r.Width, r.Height) / 2f);
            if (radius <= 0f)
            {
                path.AddRectangle(r);
                path.CloseFigure();
                return path;
            }
            var d = radius * 2f;
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    /// <summary>
    /// 复选框：原生 CheckBox 的方框走系统视觉样式，在深色皮肤里会留一个白块，
    /// 所以勾框自己画，文字仍用 TextRenderer 保持和其它控件一致的排版。
    /// </summary>
    internal class ThemeCheckBox : CheckBox
    {
        /// <summary>
        /// 自绘把文字起点挪到了勾选框右侧 20px，比原生预留的 13px 宽，
        /// AutoSize 若只按基类算就会截掉尾巴，这里按实际排版重新量一次。
        /// </summary>
        public override Size GetPreferredSize(Size proposedSize)
        {
            var sz = base.GetPreferredSize(proposedSize);
            if (!AutoSize || string.IsNullOrEmpty(Text)) return sz;
            int need = TextRenderer.MeasureText(Text, Font, Size.Empty, TextFormatFlags.NoPrefix).Width + 24;
            sz.Width = Math.Max(sz.Width, need);
            return sz;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var p = Theme.P;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(ThemeButton.FaceColorOf(this));

            const int box = 14;
            int top = (Height - box) / 2;
            var rect = new RectangleF(0.5f, top + 0.5f, box - 1f, box - 1f);
            var fill = Checked ? p.Accent : p.Surface;
            var line = Checked ? p.Accent : p.Border;
            var fore = Enabled ? p.Text : p.Muted;
            if (!Enabled)
            {
                fill = Theme.Mix(fill, ThemeButton.FaceColorOf(this), 0.55f);
                line = Theme.Mix(line, ThemeButton.FaceColorOf(this), 0.55f);
            }

            using (var path = ThemeButton.RoundRect(rect, Math.Min(p.Radius, 3f)))
            {
                using var br = new SolidBrush(fill);
                g.FillPath(br, path);
                using var pen = new Pen(line, 1f);
                g.DrawPath(pen, path);
            }

            if (Checked)
            {
                using var mark = new Pen(Theme.IsDark(fill) ? Color.White : p.AccentText, 1.8f);
                g.DrawLines(mark, new[]
                {
                    new PointF(3.2f, top + box / 2f),
                    new PointF(5.8f, top + box - 4f),
                    new PointF(box - 3.2f, top + 3.6f)
                });
            }

            int textX = box + 6;
            TextRenderer.DrawText(g, Text, Font,
                new Rectangle(textX, 0, Math.Max(0, Width - textX), Height), fore,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left
              | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }
    }

    /// <summary>
    /// 分组框：原生 GroupBox 的边框和标题底色同样走系统色，深色皮肤里会发灰，改成描边 + 主题标题。
    /// 只换绘制，容器行为和子控件坐标都不动，原来的绝对定位照样能用。
    /// </summary>
    internal class ThemeGroupBox : GroupBox
    {
        protected override void OnPaint(PaintEventArgs e)
        {
            var p = Theme.P;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(BackColor);
            if (string.IsNullOrEmpty(Text)) return;

            int pad = 8;
            var ts = TextRenderer.MeasureText(g, Text, Font, new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.NoPrefix);
            var rect = new RectangleF(pad + 4.5f, ts.Height / 2f - 0.5f,
                Width - pad * 2 - 9f, Height - ts.Height / 2f - 1f);
            using (var path = ThemeButton.RoundRect(rect, p.Radius))
            using (var pen = new Pen(p.Border, 1f))
                g.DrawPath(pen, path);

            // 标题压在边框线上，所以先把标题底下那段线擦掉再写字
            using (var bg = new SolidBrush(BackColor))
                g.FillRectangle(bg, pad + 2, 0, ts.Width + 8, ts.Height);
            TextRenderer.DrawText(g, Text, Font, new Point(pad + 4, 0),
                Enabled ? p.Text : p.Muted, TextFormatFlags.NoPrefix);
        }
    }

    /// <summary>
    /// 下拉框：原生 ComboBox 的下拉箭头块固定用系统灰 ButtonFace 铺底，皮肤色盖不住它，
    /// 所以整个控件自绘（边框、底色、文字、箭头），下拉列表仍交给原生绘制。
    /// </summary>
    internal class ThemeComboBox : ComboBox
    {
        private const int ButtonWidth = 18;

        public ThemeComboBox()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            // ComboBox 默认就是可编辑的 DropDown，构造里无条件开自绘会让内嵌编辑框露黑底；
            // 绘制开关统一按当前样式决定（见下面 DropDownStyle setter）
            SetStyle(ControlStyles.UserPaint, DropDownStyle != ComboBoxStyle.DropDown);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            var body = BackColor == SystemColors.Control ? p.Surface : BackColor;
            var full = new Rectangle(0, 0, Width - 1, Height - 1);

            g.Clear(body);
            using (var border = new Pen(ContainsFocus ? p.Accent : p.Border))
                g.DrawRectangle(border, full);

            if (!string.IsNullOrWhiteSpace(Text))
            {
                var area = new Rectangle(3, 0, Width - ButtonWidth - 5, Height);
                TextRenderer.DrawText(g, Text, Font, area, ForeColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                    TextFormatFlags.EndEllipsis);
            }

            var btn = new Rectangle(Width - ButtonWidth, 1, ButtonWidth - 1, Height - 2);
            using (var split = new Pen(p.Border))
                g.DrawLine(split, btn.Left, 2, btn.Left, btn.Bottom - 1);

            int cx = btn.Left + btn.Width / 2;
            int cy = btn.Top + btn.Height / 2;
            var arrow = new[]
            {
                new Point(cx - 4, cy - 2),
                new Point(cx + 4, cy - 2),
                new Point(cx, cy + 3)
            };
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var head = new SolidBrush(p.Muted);
            g.FillPolygon(head, arrow);
        }

        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }

        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

        /// <summary>
        /// 可编辑样式（DropDown）的文字由内嵌编辑框子窗口绘制，自绘盖不住它、还会让它露出黑底，
        /// 所以只有非编辑样式才接管绘制。
        /// </summary>
        [System.ComponentModel.DefaultValue(ComboBoxStyle.DropDown)]
        public new ComboBoxStyle DropDownStyle
        {
            get => base.DropDownStyle;
            set
            {
                base.DropDownStyle = value;
                SetStyle(ControlStyles.UserPaint, value != ComboBoxStyle.DropDown);
            }
        }
    }

    /// <summary>
    /// 分段按钮：把几个同类入口（窗格 / 布局 / 主题）合成一颗带内分隔线的胶囊，
    /// 点哪段回调哪段，段矩形对外可见，方便把下拉菜单挂到对应段下面。
    /// </summary>
    internal class ThemeSegmented : Control
    {
        private const int PadX = 13;

        private string[] _items = Array.Empty<string>();
        private int _activeIndex = -1;
        private int _hoverIndex = -1;
        private int _downIndex = -1;

        /// <summary>点击某段时回调，参数为段序号</summary>
        public event Action<int>? SegmentClicked;

        /// <summary>当前选中的段，-1 表示没有</summary>
        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public int ActiveIndex
        {
            get => _activeIndex;
            set { _activeIndex = value; Invalidate(); }
        }

        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public string[] Items
        {
            get => _items;
            set
            {
                _items = value ?? Array.Empty<string>();
                Size = CalcSize();
                Invalidate();
            }
        }

        public ThemeSegmented()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint
                   | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Cursor = Cursors.Hand;
            Font = Theme.Body;
        }

        /// <summary>皮肤切换后字体和字号都变了，尺寸要重新量</summary>
        public void RefreshTheme()
        {
            Font = Theme.Body;
            Size = CalcSize();
            Invalidate();
        }

        /// <summary>某段在控件坐标系里的位置，用于把菜单贴到它下面</summary>
        public Rectangle SegmentRect(int index) => SegmentBounds(index, BodyRect());

        private Rectangle BodyRect() => new(1, 1, Math.Max(0, Width - 2), Math.Max(0, Height - 2));

        private Size CalcSize()
        {
            int w = 0;
            foreach (var segW in SegmentWidths())
                w += segW;
            return new Size(w + 2, Theme.P.ButtonHeight + 4);
        }

        private int[]? _segmentWidths;
        private Font? _measuredFont;
        private string[]? _measuredItems;

        /// <summary>段宽只在文字或字体变了才重量：绘制、命中测试、菜单定位都要用它，每帧逐个量纯属浪费。</summary>
        private int[] SegmentWidths()
        {
            if (_segmentWidths != null && _segmentWidths.Length == _items.Length
                && ReferenceEquals(_measuredItems, _items) && ReferenceEquals(_measuredFont, Font))
                return _segmentWidths;

            _measuredItems = _items;
            _measuredFont = Font;
            _segmentWidths = new int[_items.Length];
            for (int i = 0; i < _items.Length; i++)
                _segmentWidths[i] = TextRenderer.MeasureText(_items[i], Font).Width + PadX * 2;
            return _segmentWidths;
        }

        private Rectangle SegmentBounds(int index, Rectangle body)
        {
            var widths = SegmentWidths();
            int x = body.Left;
            for (int i = 0; i < index; i++)
                x += widths[i];
            return new Rectangle(x, body.Top, widths[index], body.Height);
        }

        private int HitIndex(Point pt)
        {
            var body = BodyRect();
            for (int i = 0; i < _items.Length; i++)
                if (SegmentBounds(i, body).Contains(pt)) return i;
            return -1;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(ThemeButton.FaceColorOf(this));
            if (_items.Length == 0) return;

            var rect = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            var radius = (float)Math.Min(p.Radius + 1, Height / 2f);
            var body = BodyRect();

            using (var path = ThemeButton.RoundRect(rect, radius))
            {
                using (var back = new SolidBrush(p.Surface))
                    g.FillPath(back, path);

                var oldClip = g.Clip;
                g.SetClip(path, CombineMode.Replace);

                for (int i = 0; i < _items.Length; i++)
                {
                    var seg = SegmentBounds(i, body);
                    if (i == _downIndex)
                        using (var br = new SolidBrush(Theme.Mix(p.Selection, p.Text, 0.10f)))
                            g.FillRectangle(br, seg);
                    else if (i == _activeIndex)
                        using (var br = new SolidBrush(p.Selection))
                            g.FillRectangle(br, seg);
                    else if (i == _hoverIndex)
                        using (var br = new SolidBrush(p.Hover))
                            g.FillRectangle(br, seg);

                    if (i > 0)
                        using (var div = new Pen(p.Divider))
                            g.DrawLine(div, seg.Left, 3, seg.Left, seg.Bottom - 3);
                }

                for (int i = 0; i < _items.Length; i++)
                {
                    var seg = SegmentBounds(i, body);
                    var fore = i == _activeIndex ? p.Accent : i == _hoverIndex ? p.Text : p.Muted;
                    TextRenderer.DrawText(g, _items[i], Font, seg, fore,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                      | TextFormatFlags.NoPrefix);
                }

                g.Clip = oldClip;
                using (var pen = new Pen(p.Border, 1f))
                    g.DrawPath(pen, path);
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int hit = HitIndex(e.Location);
            if (hit == _hoverIndex) return;
            _hoverIndex = hit;
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            _downIndex = HitIndex(e.Location);
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _downIndex = -1;
            Invalidate();
            int hit = HitIndex(e.Location);
            if (hit >= 0) SegmentClicked?.Invoke(hit);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _hoverIndex = -1;
            _downIndex = -1;
            Invalidate();
        }
    }

    /// <summary>
    /// 圆角输入框：原生 TextBox 的方角边框和系统底色盖不住皮肤色，
    /// 所以把它的边框去掉，由外层面板自绘圆角底与描边，聚焦时换成 2px 强调色。
    /// </summary>
    internal class ThemeField : Panel
    {
        /// <summary>真正接收输入的无边框文本框，调用方直接拿它挂事件、读写 Text。</summary>
        public TextBox Edit { get; }

        /// <summary>占位文案，转给内部的 TextBox。</summary>
        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public string Placeholder
        {
            get => Edit.PlaceholderText;
            set => Edit.PlaceholderText = value;
        }

        public ThemeField()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            Edit = new TextBox
            {
                BorderStyle = BorderStyle.None,
                Font = Theme.Body,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
            };
            Controls.Add(Edit);
            Edit.Enter += (_, _) => Invalidate();
            Edit.Leave += (_, _) => Invalidate();
            Height = Theme.P.ButtonHeight;
            LayoutEdit();
        }

        /// <summary>皮肤或深浅色变了以后重新取令牌色、字体和高度并重画。</summary>
        public void ApplyTheme()
        {
            Height = Theme.P.ButtonHeight;
            Edit.Font = Theme.Body;
            Edit.BackColor = Theme.P.Surface;
            Edit.ForeColor = Theme.P.Text;
            LayoutEdit();
            Invalidate();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            LayoutEdit();
        }

        private void LayoutEdit()
        {
            Edit.Left = 8;
            Edit.Width = Math.Max(20, Width - 16);
            Edit.Top = Math.Max(2, (Height - Edit.Height) / 2);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (Width <= Theme.P.Radius * 2 || Height <= Theme.P.Radius * 2) return;

            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            bool focused = Edit.Focused;
            float stroke = focused ? 2f : 1f;
            var rect = new RectangleF(stroke / 2f, stroke / 2f, Width - stroke, Height - stroke);
            using var path = ThemeButton.RoundRect(rect, Theme.P.Radius);
            using (var fill = new SolidBrush(Theme.P.Surface)) g.FillPath(fill, path);
            using var pen = new Pen(focused ? Theme.P.Accent : Theme.P.Border, stroke);
            g.DrawPath(pen, path);
        }
    }

    /// <summary>
    /// 给任意内容套一圈主题描边。
    /// 原生控件的 FixedSingle 边框走系统色，深色皮肤下会露出一块灰白，只能自己画：
    /// 面板本身涂成描边色，留 1px Padding，内容填满剩下的区域就是描边效果。
    /// </summary>
    internal class ThemeBordered : Panel
    {
        public ThemeBordered(Control content)
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            Padding = new Padding(1);
            content.Dock = DockStyle.Fill;
            Controls.Add(content);
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
        }

        public void ApplyTheme()
        {
            BackColor = Theme.P.Border;
            Invalidate();
        }
    }

    /// <summary>
    /// 主题进度条。原生 ProgressBar 走系统视觉样式，深色皮肤里会留一条亮灰底槽。
    /// Marquee 用内部计时器来回扫，不需要调用方自己驱动。
    /// </summary>
    internal class ThemeProgressBar : Control
    {
        private const int MarqueeInterval = 40;
        private const float MarqueeSweepSpeed = 0.012f;

        private readonly System.Windows.Forms.Timer _marqueeTimer = new();
        private int _value;
        private bool _marquee;
        private float _marqueePos;
        private float _marqueeDir = 1f;

        public ThemeProgressBar()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint
                   | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            _marqueeTimer.Interval = MarqueeInterval;
            _marqueeTimer.Tick += (_, _) =>
            {
                _marqueePos += MarqueeSweepSpeed * _marqueeDir;
                if (_marqueePos > 1f) { _marqueePos = 1f; _marqueeDir = -1f; }
                if (_marqueePos < 0f) { _marqueePos = 0f; _marqueeDir = 1f; }
                Invalidate();
            };
            ApplyTheme();
            Theme.Changed += OnThemeChanged;
            Disposed += (_, _) =>
            {
                Theme.Changed -= OnThemeChanged;
                _marqueeTimer.Dispose();
            };
        }

        /// <summary>
        /// 启动遮罩这类控件跑在独立的 STA 线程上，静态 Theme.Changed 却是主线程发的，
        /// 直接刷会让它跨线程碰自己的句柄，所以切回创建线程再重取令牌。
        /// </summary>
        private void OnThemeChanged()
        {
            if (!IsHandleCreated) { ApplyTheme(); return; }
            if (InvokeRequired)
            {
                try { BeginInvoke((Action)ApplyTheme); } catch { }
                return;
            }
            ApplyTheme();
        }

        /// <summary>百分比进度，固定 0~100。</summary>
        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public int Value
        {
            get => _value;
            set
            {
                _value = Math.Clamp(value, 0, 100);
                Invalidate();
            }
        }

        /// <summary>不确定进度：一段色块来回扫。</summary>
        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public bool Marquee
        {
            get => _marquee;
            set
            {
                if (_marquee == value) return;
                _marquee = value;
                if (_marquee) _marqueeTimer.Start();
                else _marqueeTimer.Stop();
                Invalidate();
            }
        }

        public void ApplyTheme()
        {
            BackColor = Theme.P.Window;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            g.Clear(BackColor);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);
            float radius = Math.Min(p.Radius, Height / 2f);
            using (var path = ThemeButton.RoundRect(rect, radius))
            {
                using var bg = new SolidBrush(p.Surface);
                using var pen = new Pen(p.Border, 1f);
                g.FillPath(bg, path);
                g.DrawPath(pen, path);
            }

            // 色块往里缩一点，免得顶到端头的弧上
            var inner = RectangleF.Inflate(rect, -2f, -2f);
            if (inner.Width <= 2f || inner.Height <= 2f) return;
            float fillRadius = Math.Max(1f, radius - 2f);

            if (_marquee)
            {
                const float span = 0.28f;
                float start = _marqueePos * (1f - span);
                using var seg = ThemeButton.RoundRect(
                    new RectangleF(inner.X + start * inner.Width, inner.Y, span * inner.Width, inner.Height),
                    fillRadius);
                using var fill = new SolidBrush(p.Accent);
                g.FillPath(fill, seg);
                return;
            }

            float ratio = _value / 100f;
            if (ratio <= 0f) return;

            using var bar = ThemeButton.RoundRect(
                new RectangleF(inner.X, inner.Y, inner.Width * ratio, inner.Height), fillRadius);
            using var accent = new SolidBrush(p.Accent);
            g.FillPath(accent, bar);
        }
    }

    /// <summary>
    /// 全项目共用的一套排版数字。之前每个窗口各写一份（8/8、10/6、12/10、Padding(10,6)…），
    /// 同一屏里混着两三种节奏就显乱；改这里等于一次改完所有界面的间距。
    /// </summary>
    internal static class ThemeMetrics
    {
        /// <summary>容器左右边距</summary>
        public const int Margin = 10;

        /// <summary>同一组控件之间的间距</summary>
        public const int Gap = 6;

        /// <summary>不同功能组之间的间距</summary>
        public const int Group = 14;

        /// <summary>首行距容器顶、末行距容器底的留白</summary>
        public const int RowPad = 8;

        /// <summary>按钮/输入框统一高度，跟着皮肤走</summary>
        public static int RowH => Theme.P.ButtonHeight;

        /// <summary>底部按钮条的高度：一行控件加上下留白。</summary>
        public static int BarHeight => RowH + RowPad * 2;
    }

    /// <summary>
    /// 对话框底部按钮条：一颗实心主按钮 + 若干描边按钮，右对齐。
    /// 各对话框原先自己写 100×30、110×34 和品牌蓝 FromArgb(66,165,245)，
    /// 统一到这里以后，改一次全部对话框跟着变。
    /// </summary>
    internal static class DialogButtonBar
    {
        /// <summary>造一颗按文字量宽的按钮。</summary>
        public static ThemeButton Make(string text, EventHandler? onClick = null, bool primary = false)
        {
            var b = new ThemeButton
            {
                Text = text,
                Primary = primary,
                Font = Theme.Body,
                BackColor = Theme.P.Chrome,
                Size = new Size(ThemeButton.FitWidth(text), ThemeMetrics.RowH)
            };
            if (onClick != null) b.Click += onClick;
            return b;
        }

        /// <summary>造一条贴底的按钮容器，顶上一圈分隔线。</summary>
        public static Panel Create(out Panel divider)
        {
            var bar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = ThemeMetrics.BarHeight,
                BackColor = Theme.P.Chrome,
                Padding = new Padding(0)
            };
            divider = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Theme.P.Divider };
            bar.Controls.Add(divider);
            return bar;
        }

        public static Panel Create() => Create(out _);

        /// <summary>
        /// 按传入顺序从左到右排进容器，整组靠右对齐。
        /// 容器变窄时按钮贴着右边走，所以顺手挂上 Right 锚定，不然拉伸窗口按钮会留在原地。
        /// </summary>
        public static void LayoutRight(Control host, params ThemeButton[] buttons)
        {
            int x = host.Width - ThemeMetrics.Margin;
            for (int i = buttons.Length - 1; i >= 0; i--)
            {
                var b = buttons[i];
                int w = ThemeButton.FitWidth(b.Text);
                b.Size = new Size(w, ThemeMetrics.RowH);
                b.Location = new Point(x - w, ThemeMetrics.RowPad);
                x -= w + ThemeMetrics.Gap;
                b.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            }
        }

        /// <summary>
        /// 按传入顺序从左到右贴左排列。放「清除设置」这类和窗口动作不同性质的按钮，
        /// 让它们待在左侧，不挤占右边 确定/取消 的位置。
        /// </summary>
        public static void LayoutLeft(Control host, params ThemeButton[] buttons)
        {
            int x = ThemeMetrics.Margin;
            foreach (var b in buttons)
            {
                int w = ThemeButton.FitWidth(b.Text);
                b.Size = new Size(w, ThemeMetrics.RowH);
                b.Location = new Point(x, ThemeMetrics.RowPad);
                x += w + ThemeMetrics.Gap;
                b.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            }
        }
    }

    /// <summary>
    /// 工具型 DataGridView 的统一取色。
    /// 主表格（TableBrowserControl）自己带冻结行、斑马纹开关，样式要按它的选项算，所以不走这里；
    /// 剩下的重复值详情、版本对比这类小表格只需要"跟当前皮肤一致"，用这一份就够。
    /// </summary>
    internal static class ThemeGrid
    {
        public static void Apply(DataGridView dgv)
        {
            var p = Theme.P;
            dgv.BackgroundColor = p.Surface;
            dgv.GridColor = p.GridLine;
            dgv.DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = p.Surface,
                ForeColor = p.Text,
                Font = Theme.Body,
                SelectionBackColor = p.Selection,
                SelectionForeColor = p.Text,
                Padding = new Padding(4, 0, 4, 0)
            };
            dgv.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = p.Zebra,
                SelectionBackColor = p.Selection,
                SelectionForeColor = p.Text
            };
            dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = p.Header,
                ForeColor = p.Text,
                Font = Theme.Head
            };
            dgv.ColumnHeadersHeight = p.RowHeight + 4;
            dgv.RowTemplate.Height = p.RowHeight;
            dgv.EnableHeadersVisualStyles = false;
        }
    }
}
