using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>DataGridView 关联字段的自定义 tooltip（使用 ToolStripDropDown + 自绘 Panel）</summary>
    public class DataGridLinkToolTip : IDisposable
    {
        private readonly List<DataGridView> _grids = new();
        private DataGridView? _activeGrid;
        private readonly ToolStripDropDown _dropDown;
        private readonly CustomToolTipPanel _panel;
        private readonly System.Windows.Forms.Timer _hoverTimer;
        private string? _currentToolTipText;
        private string? _currentTitle;
        private List<TipLine> _currentLines = new();
        private int _lastRow = -1, _lastCol = -1;
        private bool _disposed;
        private int _dropDownWidth;

        private static readonly Font _titleFont = new("Microsoft YaHei UI", 10f, FontStyle.Bold);
        private static readonly Font _textFont = new("Microsoft YaHei UI", 9f);
        private static readonly Font _hintFont = new("Microsoft YaHei UI", 8f, FontStyle.Italic);

        private static readonly Color _bgColor = Color.FromArgb(252, 252, 254);
        private static readonly Color _borderColor = Color.FromArgb(200, 210, 225);
        private static readonly Color _titleColor = Color.FromArgb(30, 100, 180);
        private static readonly Color _normalTextColor = Color.FromArgb(60, 60, 70);
        private static readonly Color _warningTextColor = Color.FromArgb(200, 40, 40);
        private static readonly Color _hintColor = Color.FromArgb(140, 150, 170);
        private static readonly Color _separatorColor = Color.FromArgb(220, 228, 240);

        /// <summary>用于外部提供 tooltip 内容的委托（grid: 当前鼠标所在的 DataGridView）</summary>
        public Func<DataGridView, int, int, (string? title, List<TipLine> lines, string? rawText)>? GetToolTipContent { get; set; }

        /// <summary>tooltip 中的一行信息</summary>
        public class TipLine
        {
            public string Text { get; set; } = "";
            public bool IsWarning { get; set; }
        }

        public DataGridLinkToolTip(DataGridView primaryGrid)
        {
            _activeGrid = primaryGrid;
            RegisterGrid(primaryGrid);

            _panel = new CustomToolTipPanel(this);

            _dropDown = new ToolStripDropDown
            {
                AutoClose = false,
                Padding = Padding.Empty,
                Margin = Padding.Empty,
                DropShadowEnabled = false,
                BackColor = _bgColor,
                Renderer = new NoBorderRenderer(),
            };

            var host = new ToolStripControlHost(_panel)
            {
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                AutoSize = false
            };
            _dropDown.Items.Add(host);

            _dropDown.MouseLeave += (s, e) => Hide();

            _hoverTimer = new System.Windows.Forms.Timer { Interval = 400 };
            _hoverTimer.Tick += OnHoverTimerTick;
        }

        /// <summary>注册一个额外的 DataGridView（如固定行表头）</summary>
        public void RegisterGrid(DataGridView grid)
        {
            if (_grids.Contains(grid)) return;
            _grids.Add(grid);
            grid.ShowCellToolTips = false;
            grid.MouseLeave += (s, e) => Hide();
            grid.CellMouseMove += OnGridCellMouseMove;
        }

        private void OnGridCellMouseMove(object? sender, DataGridViewCellMouseEventArgs e)
        {
            // 允许列标题（RowIndex=-1）也触发自定义 ToolTip
            if (e.ColumnIndex < 0)
            {
                Hide();
                return;
            }

            if (sender is DataGridView grid)
                _activeGrid = grid;

            if (e.RowIndex != _lastRow || e.ColumnIndex != _lastCol)
            {
                // 单元格切换：立即隐藏旧 tooltip，然后延迟显示新 tooltip
                _dropDown.Close();
                _lastRow = e.RowIndex;
                _lastCol = e.ColumnIndex;
                _hoverTimer?.Stop();
                _hoverTimer?.Start();
            }
        }

        private void OnHoverTimerTick(object? sender, EventArgs e)
        {
            _hoverTimer?.Stop();
            Show();
        }

        private void Show()
        {
            if (_lastRow < 0 || _lastCol < 0 || GetToolTipContent == null || _activeGrid == null) return;

            var content = GetToolTipContent(_activeGrid, _lastRow, _lastCol);
            if (content.title == null && content.lines.Count == 0 && string.IsNullOrEmpty(content.rawText))
            {
                Hide();
                return;
            }

            _currentTitle = content.title;
            _currentLines = content.lines;
            _currentToolTipText = content.rawText;

            // 测量并计算布局
            _panel.MeasureLayout();

            // 计算显示位置（使用鼠标位置，紧跟鼠标，不远离）
            var mouseScreenPos = Cursor.Position;
            // 默认显示在鼠标右下方
            var screenPos = new Point(mouseScreenPos.X + 14, mouseScreenPos.Y + 8);

            // 屏幕边界调整
            var screen = Screen.FromPoint(screenPos).WorkingArea;

            // 限制 tooltip 最大高度（不超过屏幕可用高度的 80%，避免内容过长时无法显示）
            int maxTipHeight = (int)(screen.Height * 0.8);
            if (_panel.Height > maxTipHeight)
                _panel.Height = maxTipHeight;

            // X 坐标修正：优先显示在鼠标右侧，空间不够则翻转到鼠标左侧
            if (screenPos.X + _dropDownWidth > screen.Right)
            {
                int leftPos = mouseScreenPos.X - _dropDownWidth - 14;
                if (leftPos >= screen.Left)
                    screenPos.X = leftPos;
                else
                    screenPos.X = screen.Right - _dropDownWidth - 5;
            }
            if (screenPos.X < screen.Left)
                screenPos.X = screen.Left + 5;

            // Y 坐标修正：优先显示在鼠标下方，空间不够则翻转到鼠标上方
            if (screenPos.Y + _panel.Height > screen.Bottom)
            {
                int topPos = mouseScreenPos.Y - _panel.Height - 8;
                if (topPos >= screen.Top)
                    screenPos.Y = topPos;
                else
                    screenPos.Y = Math.Max(screen.Top, screen.Bottom - _panel.Height - 5);
            }
            if (screenPos.Y < screen.Top)
                screenPos.Y = screen.Top + 5;

            if (!_disposed)
            {
                _panel.Size = new Size(_dropDownWidth, _panel.Height);
                var host = (ToolStripControlHost)_dropDown.Items[0];
                host.Size = _panel.Size;
                _dropDown.Size = _panel.Size;
                _dropDown.Show(screenPos);
            }
        }

        public void Hide()
        {
            _currentToolTipText = null;
            _currentTitle = null;
            _currentLines.Clear();
            _hoverTimer?.Stop();
            _dropDown.Close();
        }

        /// <summary>创建一个固定的 tooltip 窗口（独立于悬浮 tooltip）</summary>
        public PinnedTipWindow? CreatePinnedWindow(string? title, List<TipLine> lines, string? rawText, Point screenPos)
        {
            if (string.IsNullOrEmpty(title) && lines.Count == 0 && string.IsNullOrEmpty(rawText))
                return null;
            return new PinnedTipWindow(title, lines, rawText, screenPos);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _hoverTimer?.Dispose();
            _dropDown.Dispose();
            _panel.Dispose();
            // 注意：_titleFont/_textFont/_hintFont 是 static 字段，生命周期应持续到 AppDomain 卸载，
            // 不能在这里 Dispose，否则后续新创建的实例会使用已释放的字体导致 tooltip 文字无法绘制。
        }

        /// <summary>无边框渲染器，禁止 ToolStripDropDown 绘制系统边框</summary>
        private class NoBorderRenderer : ToolStripProfessionalRenderer
        {
            public NoBorderRenderer() : base(new NoBorderColorTable()) { }
            protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { /* 不绘制边框 */ }
        }

        private class NoBorderColorTable : ProfessionalColorTable
        {
            public override Color ToolStripDropDownBackground => Color.Transparent;
            public override Color ImageMarginGradientBegin => Color.Transparent;
            public override Color ImageMarginGradientMiddle => Color.Transparent;
            public override Color ImageMarginGradientEnd => Color.Transparent;
        }

        /// <summary>自绘 Panel，负责渲染 tooltip 内容</summary>
        private class CustomToolTipPanel : Panel
        {
            private readonly DataGridLinkToolTip _owner;
            private int _titleHeight;
            private int _lineHeight;
            private int _hintHeight;
            private List<(string text, bool isWarning)> _layoutLines = new();
            private string _hintText = "";
            private bool _hasLinkMapping; // 是否有关联映射行（控制底部提示显示）

            // 布局常量
            private const int PaddingH = 14;
            private const int PaddingV = 10;
            private const int LineSpacing = 3;
            private const int SeparatorHeight = 1;
            private const int SeparatorMargin = 6;
            private const int CornerRadius = 6;
            private const int MinWidth = 160;
            private const int MaxWidth = 520;

            public CustomToolTipPanel(DataGridLinkToolTip owner)
            {
                _owner = owner;
                DoubleBuffered = true;
                BackColor = Color.Transparent;
            }

            /// <summary>安全绘制字符串，防止特殊字符导致 GDI+ 崩溃</summary>
            internal static void SafeDrawString(Graphics g, string text, Font font, Brush brush, RectangleF rect, StringFormat? format = null)
            {
                if (string.IsNullOrEmpty(text)) return;
                try
                {
                    if (format != null)
                        g.DrawString(text, font, brush, rect, format);
                    else
                        g.DrawString(text, font, brush, rect);
                }
                catch
                {
                    // 过滤不可打印字符后重试
                    var safe = text.Length > 2000 ? text.Substring(0, 2000) : text;
                    safe = new string(safe.Where(c => c >= 0x20 || c == '\n' || c == '\r' || c == '\t').ToArray());
                    try
                    {
                        if (format != null)
                            g.DrawString(safe, font, brush, rect, format);
                        else
                            g.DrawString(safe, font, brush, rect);
                    }
                    catch
                    {
                        // 放弃绘制
                    }
                }
            }

            /// <summary>安全测量字符串，防止特殊字符导致 GDI+ 崩溃</summary>
            private static SizeF SafeMeasureString(Graphics g, string text, Font font)
            {
                if (string.IsNullOrEmpty(text))
                    return SizeF.Empty;
                try
                {
                    return g.MeasureString(text, font);
                }
                catch
                {
                    // GDI+ MeasureString 可能因特殊字符而崩溃
                    // 回退：使用 TextRenderer 测量（更稳定）
                    try
                    {
                        var size = TextRenderer.MeasureText(g, text, font);
                        return new SizeF(size.Width, size.Height);
                    }
                    catch
                    {
                        // 最终回退：限制长度并清除控制字符后再次尝试
                        var safe = text.Length > 2000 ? text.Substring(0, 2000) : text;
                        safe = new string(safe.Where(c => c >= 0x20 || c == '\n' || c == '\r' || c == '\t').ToArray());
                        try
                        {
                            var size = TextRenderer.MeasureText(g, safe, font);
                            return new SizeF(size.Width, size.Height);
                        }
                        catch
                        {
                            return new SizeF(50, 16); // 兜底值
                        }
                    }
                }
            }

            /// <summary>安全测量字符串（指定最大宽度，用于计算自动换行后的高度）</summary>
            private static SizeF SafeMeasureString(Graphics g, string text, Font font, int maxWidth)
            {
                if (string.IsNullOrEmpty(text))
                    return SizeF.Empty;
                try
                {
                    return g.MeasureString(text, font, maxWidth);
                }
                catch
                {
                    try
                    {
                        var size = TextRenderer.MeasureText(g, text, font, new Size(maxWidth, int.MaxValue),
                            TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
                        return new SizeF(size.Width, size.Height);
                    }
                    catch
                    {
                        var safe = text.Length > 2000 ? text.Substring(0, 2000) : text;
                        safe = new string(safe.Where(c => c >= 0x20 || c == '\n' || c == '\r' || c == '\t').ToArray());
                        try
                        {
                            var size = TextRenderer.MeasureText(g, safe, font, new Size(maxWidth, int.MaxValue),
                                TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
                            return new SizeF(size.Width, size.Height);
                        }
                        catch
                        {
                            return new SizeF(50, 16);
                        }
                    }
                }
            }

            public void MeasureLayout()
            {
                using var g = CreateGraphics();

                // 测量标题
                _titleHeight = 0;
                int maxWidth = MinWidth;
                if (!string.IsNullOrEmpty(_owner._currentTitle))
                {
                    var titleSize = SafeMeasureString(g, _owner._currentTitle, _titleFont);
                    _titleHeight = (int)Math.Ceiling(titleSize.Height);
                    // ★ 标题行宽度也参与 maxWidth 计算（拆分换行符逐行测量）
                    var titleLines = _owner._currentTitle.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                    foreach (var tl in titleLines)
                    {
                        int tw = (int)Math.Ceiling(SafeMeasureString(g, tl, _titleFont).Width);
                        if (tw > maxWidth) maxWidth = tw;
                    }
                }

                // 测量各行（支持换行符拆分 + 自动换行）
                _layoutLines.Clear();
                _hasLinkMapping = false;
                _lineHeight = (int)Math.Ceiling(SafeMeasureString(g, "Ag", _textFont).Height);

                if (_owner._currentLines.Count > 0)
                {
                    foreach (var line in _owner._currentLines)
                    {
                        // 拆分换行符
                        var subLines = line.Text.Split(new[] { '\n' }, StringSplitOptions.None);
                        foreach (var sub in subLines)
                        {
                            _layoutLines.Add((sub, line.IsWarning));
                            var w = (int)Math.Ceiling(SafeMeasureString(g, sub, _textFont).Width);
                            if (w > maxWidth) maxWidth = w;
                        }
                        if (line.Text.Contains("→")) _hasLinkMapping = true;
                    }
                }
                else if (!string.IsNullOrEmpty(_owner._currentToolTipText))
                {
                    var subLines = _owner._currentToolTipText.Split(new[] { '\n' }, StringSplitOptions.None);
                    foreach (var sub in subLines)
                    {
                        _layoutLines.Add((sub, false));
                        var w = (int)Math.Ceiling(SafeMeasureString(g, sub, _textFont).Width);
                        if (w > maxWidth) maxWidth = w;
                    }
                }

                // 底部提示（仅有关联映射行时显示）
                _hintText = _hasLinkMapping ? "右键 → 打开关联表" : "";
                _hintHeight = _hasLinkMapping ? (int)Math.Ceiling(SafeMeasureString(g, _hintText, _hintFont).Height) : 0;

                _owner._dropDownWidth = Math.Min(maxWidth + PaddingH * 2, MaxWidth);
                if (_owner._dropDownWidth < MinWidth) _owner._dropDownWidth = MinWidth;

                // 计算高度时，对每行文本使用最大宽度约束测量，支持自动换行
                int contentMaxWidth = _owner._dropDownWidth - PaddingH * 2;
                int totalHeight = PaddingV;
                if (_titleHeight > 0)
                {
                    // 标题也需要考虑换行
                    var titleSizeConstrained = SafeMeasureString(g, _owner._currentTitle!, _titleFont, contentMaxWidth);
                    _titleHeight = (int)Math.Ceiling(titleSizeConstrained.Height);
                    totalHeight += _titleHeight + SeparatorMargin + SeparatorHeight + SeparatorMargin;
                }
                if (_layoutLines.Count > 0)
                {
                    foreach (var line in _layoutLines)
                    {
                        // 在限定宽度下测量实际高度（支持自动换行）
                        var constrainedSize = SafeMeasureString(g, line.text, _textFont, contentMaxWidth);
                        totalHeight += (int)Math.Ceiling(constrainedSize.Height) + LineSpacing;
                    }
                    totalHeight -= LineSpacing; // 最后一行不需要额外的 LineSpacing
                }
                if (_hasLinkMapping)
                {
                    totalHeight += SeparatorMargin + SeparatorHeight + SeparatorMargin;
                    totalHeight += _hintHeight;
                }
                totalHeight += PaddingV;

                Height = totalHeight;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                var rect = new Rectangle(0, 0, Width - 1, Height - 1);

                // 绘制圆角背景
                using (var path = GetRoundedRectPath(rect, CornerRadius))
                using (var bgBrush = new SolidBrush(_bgColor))
                {
                    g.FillPath(bgBrush, path);
                }

                // 绘制边框
                using (var path = GetRoundedRectPath(rect, CornerRadius))
                using (var borderPen = new Pen(_borderColor, 1f))
                {
                    g.DrawPath(borderPen, path);
                }

                int y = PaddingV;
                var contentWidth = Width - PaddingH * 2;

                // 绘制标题
                if (!string.IsNullOrEmpty(_owner._currentTitle))
                {
                    var titleRect = new Rectangle(PaddingH, y, contentWidth, _titleHeight);
                    using var titleBrush = new SolidBrush(_titleColor);
                    SafeDrawString(g, _owner._currentTitle, _titleFont, titleBrush, titleRect);
                    y += _titleHeight + SeparatorMargin;

                    // 分隔线
                    using var sepPen = new Pen(_separatorColor, 1f);
                    g.DrawLine(sepPen, PaddingH, y, Width - PaddingH, y);
                    y += SeparatorHeight + SeparatorMargin;
                }

                // 绘制映射行（使用约束宽度下的实际高度，支持自动换行）
                if (_layoutLines.Count > 0)
                {
                    foreach (var line in _layoutLines)
                    {
                        var color = line.isWarning ? _warningTextColor : _normalTextColor;
                        using var textBrush = new SolidBrush(color);
                        var fmt = new StringFormat
                        {
                            Trimming = StringTrimming.None,
                            FormatFlags = 0
                        };
                        var constrainedSize = SafeMeasureString(g, line.text, _textFont, contentWidth);
                        var textRect = new Rectangle(PaddingH, y, contentWidth, (int)Math.Ceiling(constrainedSize.Height) + LineSpacing);
                        SafeDrawString(g, line.text, _textFont, textBrush, textRect, fmt);
                        y += textRect.Height + LineSpacing;
                    }
                }

                // 底部提示（仅有关联映射行时显示）
                if (_hasLinkMapping)
                {
                    y += SeparatorMargin;
                    using (var sepPen = new Pen(_separatorColor, 1f))
                        g.DrawLine(sepPen, PaddingH, y, Width - PaddingH, y);
                    y += SeparatorHeight + SeparatorMargin;

                    var hintRect = new Rectangle(PaddingH, y, contentWidth, _hintHeight);
                    using var hintBrush = new SolidBrush(_hintColor);
                    SafeDrawString(g, _hintText, _hintFont, hintBrush, hintRect);
                }
            }

            private static GraphicsPath GetRoundedRectPath(Rectangle rect, int radius)
            {
                var path = new GraphicsPath();
                int d = radius * 2;
                path.AddArc(rect.X, rect.Y, d, d, 180, 90);
                path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
                path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
                path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }
        }

        /// <summary>固定的 tooltip 窗口（独立于悬浮 tooltip，支持多个并存）</summary>
        public class PinnedTipWindow : IDisposable
        {
            private readonly ToolStripDropDown _dropDown;
            private readonly PinnedTipPanel _panel;
            private bool _disposed;

            // 关闭按钮
            private const int CloseBtnSize = 20;
            private const int CloseBtnMargin = 6;

            private static readonly Font _titleFont2 = new("Microsoft YaHei UI", 10f, FontStyle.Bold);
            private static readonly Font _textFont2 = new("Microsoft YaHei UI", 9f);

            private static readonly Color _bgColor2 = Color.FromArgb(255, 255, 248);
            private static readonly Color _borderColor2 = Color.FromArgb(180, 190, 140);
            private static readonly Color _titleColor2 = Color.FromArgb(120, 100, 30);
            private static readonly Color _normalTextColor2 = Color.FromArgb(60, 60, 70);
            private static readonly Color _warningTextColor2 = Color.FromArgb(200, 40, 40);
            private static readonly Color _separatorColor2 = Color.FromArgb(220, 215, 195);

            private const int PaddingH = 14;
            private const int PaddingV = 10;
            private const int LineSpacing = 3;
            private const int SeparatorHeight = 1;
            private const int SeparatorMargin = 6;
            private const int CornerRadius = 6;
            private const int MinWidth = 160;
            private const int MaxWidth = 520;

            public PinnedTipWindow(string? title, List<TipLine> lines, string? rawText, Point screenPos)
            {
                _panel = new PinnedTipPanel(this, title, lines, rawText);

                _dropDown = new ToolStripDropDown
                {
                    AutoClose = false,
                    Padding = Padding.Empty,
                    Margin = Padding.Empty,
                    DropShadowEnabled = false,
                    BackColor = _bgColor2,
                    Renderer = new NoBorderRenderer(),
                };

                var host = new ToolStripControlHost(_panel)
                {
                    Margin = Padding.Empty,
                    Padding = Padding.Empty,
                    AutoSize = false
                };
                _dropDown.Items.Add(host);

                // 测量并设置大小
                _panel.MeasureSize();
                _panel.Size = new Size(_panel._measuredWidth, _panel._measuredHeight);
                host.Size = _panel.Size;
                _dropDown.Size = _panel.Size;

                // 屏幕边界调整
                var screen = Screen.FromPoint(screenPos).WorkingArea;

                // 限制最大高度
                int maxTipHeight = (int)(screen.Height * 0.8);
                if (_panel._measuredHeight > maxTipHeight)
                    _panel._measuredHeight = maxTipHeight;

                // X 修正：优先右侧，不够则翻转到左侧
                if (screenPos.X + _panel._measuredWidth > screen.Right)
                {
                    int leftPos = screenPos.X - _panel._measuredWidth - 6;
                    screenPos.X = leftPos >= screen.Left ? leftPos : screen.Right - _panel._measuredWidth - 5;
                }
                if (screenPos.X < screen.Left) screenPos.X = screen.Left + 5;

                // Y 修正：优先下方，不够则翻转到上方
                if (screenPos.Y + _panel._measuredHeight > screen.Bottom)
                    screenPos.Y = Math.Max(screen.Top, screen.Bottom - _panel._measuredHeight - 5);

                _dropDown.Show(screenPos);
            }

            public void Close()
            {
                if (_disposed) return;
                _dropDown.Close();
                Dispose();
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                _dropDown.Dispose();
                _panel.Dispose();
            }

            /// <summary>自绘 Panel —— 固定窗口的内容面板（带关闭按钮和拖动支持）</summary>
            private class PinnedTipPanel : Panel
            {
                private readonly PinnedTipWindow _owner;
                private readonly string? _title;
                private readonly List<TipLine> _lines;
                private readonly string? _rawText;

                private List<(string text, bool isWarning)> _layoutLines = new();
                private int _titleHeight;
                private int _lineHeight;
                private Rectangle _closeBtnRect;
                private bool _closeHovered;
                private Point _dragStart;
                private bool _dragging;

                public int _measuredWidth;
                public int _measuredHeight;

                public PinnedTipPanel(PinnedTipWindow owner, string? title, List<TipLine> lines, string? rawText)
                {
                    _owner = owner;
                    _title = title;
                    _lines = lines ?? new List<TipLine>();
                    _rawText = rawText;
                    DoubleBuffered = true;
                    BackColor = Color.Transparent;
                    Cursor = Cursors.SizeAll;
                }

                /// <summary>安全测量字符串（PinnedTipPanel 版本）</summary>
                private static SizeF SafeMeasureStringPinned(Graphics g, string text, Font font)
                {
                    if (string.IsNullOrEmpty(text))
                        return SizeF.Empty;
                    try
                    {
                        return g.MeasureString(text, font);
                    }
                    catch
                    {
                        try
                        {
                            var size = TextRenderer.MeasureText(g, text, font);
                            return new SizeF(size.Width, size.Height);
                        }
                        catch
                        {
                            var safe = text.Length > 2000 ? text.Substring(0, 2000) : text;
                            safe = new string(safe.Where(c => c >= 0x20 || c == '\n' || c == '\r' || c == '\t').ToArray());
                            try
                            {
                                var size = TextRenderer.MeasureText(g, safe, font);
                                return new SizeF(size.Width, size.Height);
                            }
                            catch
                            {
                                return new SizeF(50, 16);
                            }
                        }
                    }
                }

                public void MeasureSize()
                {
                    using var g = CreateGraphics();

                    _titleHeight = 0;
                    if (!string.IsNullOrEmpty(_title))
                    {
                        var titleSize = SafeMeasureStringPinned(g, _title, _titleFont2);
                        _titleHeight = (int)Math.Ceiling(titleSize.Height);
                    }

                    _layoutLines.Clear();
                    _lineHeight = (int)Math.Ceiling(SafeMeasureStringPinned(g, "Ag", _textFont2).Height);
                    int maxWidth = MinWidth;

                    // 为关闭按钮留出右侧空间
                    int titleRightMargin = CloseBtnMargin + CloseBtnSize + CloseBtnMargin;

                    if (_lines.Count > 0)
                    {
                        foreach (var line in _lines)
                        {
                            var subLines = line.Text.Split(new[] { '\n' }, StringSplitOptions.None);
                            foreach (var sub in subLines)
                            {
                                _layoutLines.Add((sub, line.IsWarning));
                                var w = (int)Math.Ceiling(SafeMeasureStringPinned(g, sub, _textFont2).Width);
                                if (w > maxWidth) maxWidth = w;
                            }
                        }
                    }
                    else if (!string.IsNullOrEmpty(_rawText))
                    {
                        var subLines = _rawText.Split(new[] { '\n' }, StringSplitOptions.None);
                        foreach (var sub in subLines)
                        {
                            _layoutLines.Add((sub, false));
                            var w = (int)Math.Ceiling(SafeMeasureStringPinned(g, sub, _textFont2).Width);
                            if (w > maxWidth) maxWidth = w;
                        }
                    }

                    // 确保宽度至少能容纳标题+关闭按钮
                    if (!string.IsNullOrEmpty(_title))
                    {
                        int titleW = (int)Math.Ceiling(SafeMeasureStringPinned(g, _title, _titleFont2).Width) + PaddingH + titleRightMargin;
                        if (titleW > maxWidth) maxWidth = titleW;
                    }
                    else if (maxWidth < MinWidth)
                    {
                        maxWidth = MinWidth;
                    }

                    _measuredWidth = Math.Min(maxWidth + PaddingH * 2, MaxWidth);
                    if (_measuredWidth < MinWidth) _measuredWidth = MinWidth;

                    int totalHeight = PaddingV;
                    if (_titleHeight > 0) totalHeight += _titleHeight + SeparatorMargin + SeparatorHeight + SeparatorMargin;
                    if (_layoutLines.Count > 0)
                        totalHeight += _layoutLines.Count * (_lineHeight + LineSpacing) - LineSpacing;
                    totalHeight += PaddingV;

                    _measuredHeight = totalHeight;
                }

                protected override void OnPaint(PaintEventArgs e)
                {
                    base.OnPaint(e);
                    var g = e.Graphics;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                    var rect = new Rectangle(0, 0, Width - 1, Height - 1);

                    // 圆角背景（淡黄色，区别于悬浮 tooltip）
                    using (var path = GetRoundedRectPath(rect, CornerRadius))
                    using (var bgBrush = new SolidBrush(_bgColor2))
                    {
                        g.FillPath(bgBrush, path);
                    }

                    // 边框
                    using (var path = GetRoundedRectPath(rect, CornerRadius))
                    using (var borderPen = new Pen(_borderColor2, 1f))
                    {
                        g.DrawPath(borderPen, path);
                    }

                    // 关闭按钮（右上角，自绘 X）
                    _closeBtnRect = new Rectangle(Width - CloseBtnMargin - CloseBtnSize, CloseBtnMargin, CloseBtnSize, CloseBtnSize);
                    using (var cbBgBrush = new SolidBrush(_closeHovered ? Color.FromArgb(240, 220, 220) : Color.FromArgb(250, 248, 240)))
                    {
                        g.FillRectangle(cbBgBrush, _closeBtnRect);
                    }
                    using (var cbBorderPen = new Pen(_closeHovered ? Color.FromArgb(220, 80, 80) : Color.FromArgb(200, 200, 180), 1f))
                    {
                        g.DrawRectangle(cbBorderPen, _closeBtnRect);
                    }
                    // 绘制 X
                    int pad = 5;
                    using var xPen = new Pen(_closeHovered ? Color.FromArgb(200, 50, 50) : Color.FromArgb(150, 140, 120), 1.8f);
                    g.DrawLine(xPen, _closeBtnRect.Left + pad, _closeBtnRect.Top + pad,
                               _closeBtnRect.Right - pad, _closeBtnRect.Bottom - pad);
                    g.DrawLine(xPen, _closeBtnRect.Right - pad, _closeBtnRect.Top + pad,
                               _closeBtnRect.Left + pad, _closeBtnRect.Bottom - pad);

                    int y = PaddingV;
                    // 标题区域要给关闭按钮让路
                    int titleMaxWidth = Width - PaddingH - CloseBtnMargin - CloseBtnSize - CloseBtnMargin - PaddingH;
                    if (titleMaxWidth < 40) titleMaxWidth = 40;
                    var contentWidth = Width - PaddingH * 2;

                    // 绘制标题
                    if (!string.IsNullOrEmpty(_title))
                    {
                        var titleRect = new Rectangle(PaddingH, y, titleMaxWidth, _titleHeight);
                        using var titleBrush = new SolidBrush(_titleColor2);
                        var titleFmt = new StringFormat { Trimming = StringTrimming.EllipsisCharacter };
                        CustomToolTipPanel.SafeDrawString(g, _title, _titleFont2, titleBrush, titleRect, titleFmt);
                        y += _titleHeight + SeparatorMargin;

                        using var sepPen = new Pen(_separatorColor2, 1f);
                        g.DrawLine(sepPen, PaddingH, y, Width - PaddingH, y);
                        y += SeparatorHeight + SeparatorMargin;
                    }

                    // 绘制内容行
                    if (_layoutLines.Count > 0)
                    {
                        foreach (var line in _layoutLines)
                        {
                            var color = line.isWarning ? _warningTextColor2 : _normalTextColor2;
                            using var textBrush = new SolidBrush(color);
                            var fmt = new StringFormat
                            {
                                Trimming = StringTrimming.None,
                                FormatFlags = 0
                            };
                            var textSize = SafeMeasureStringPinned(g, line.text, _textFont2);
                            var textRect = new Rectangle(PaddingH, y, contentWidth, (int)Math.Ceiling(textSize.Height) + LineSpacing);
                            CustomToolTipPanel.SafeDrawString(g, line.text, _textFont2, textBrush, textRect, fmt);
                            y += textRect.Height + LineSpacing;
                        }
                    }
                }

                protected override void OnMouseClick(MouseEventArgs e)
                {
                    if (_closeBtnRect.Contains(e.Location))
                        _owner.Close();
                    base.OnMouseClick(e);
                }

                protected override void OnMouseMove(MouseEventArgs e)
                {
                    base.OnMouseMove(e);
                    bool newHovered = _closeBtnRect.Contains(e.Location);
                    if (newHovered != _closeHovered)
                    {
                        _closeHovered = newHovered;
                        Cursor = _closeHovered ? Cursors.Hand : Cursors.SizeAll;
                        Invalidate();
                    }

                    if (_dragging)
                    {
                        int dx = e.X - _dragStart.X;
                        int dy = e.Y - _dragStart.Y;
                        if (Math.Abs(dx) > 2 || Math.Abs(dy) > 2)
                        {
                            var newPos = new Point(
                                _owner._dropDown.Left + dx,
                                _owner._dropDown.Top + dy);
                            _owner._dropDown.Location = newPos;
                        }
                    }
                }

                protected override void OnMouseDown(MouseEventArgs e)
                {
                    base.OnMouseDown(e);
                    if (!_closeBtnRect.Contains(e.Location))
                    {
                        _dragging = true;
                        _dragStart = e.Location;
                    }
                }

                protected override void OnMouseUp(MouseEventArgs e)
                {
                    base.OnMouseUp(e);
                    _dragging = false;
                }

                private static GraphicsPath GetRoundedRectPath(Rectangle rect, int radius)
                {
                    var path = new GraphicsPath();
                    int d = radius * 2;
                    path.AddArc(rect.X, rect.Y, d, d, 180, 90);
                    path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
                    path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
                    path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
                    path.CloseFigure();
                    return path;
                }
            }
        }
    }
}
