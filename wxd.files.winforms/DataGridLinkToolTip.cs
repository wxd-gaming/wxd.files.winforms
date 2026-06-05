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
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
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

            // 计算显示位置（在单元格右下侧）
            var cellRect = _activeGrid.GetCellDisplayRectangle(_lastCol, _lastRow, false);
            var screenPos = _activeGrid.PointToScreen(new Point(cellRect.Right + 6, cellRect.Top));

            // 屏幕边界调整
            var screen = Screen.FromPoint(screenPos).WorkingArea;
            if (screenPos.X + _dropDownWidth > screen.Right)
                screenPos.X = screen.Right - _dropDownWidth - 5;
            if (screenPos.X < screen.Left)
                screenPos.X = screen.Left + 5;
            if (screenPos.Y + _panel.Height > screen.Bottom)
                screenPos.Y = screen.Bottom - _panel.Height - 5;

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

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _hoverTimer?.Dispose();
            _dropDown.Dispose();
            _panel.Dispose();
            _titleFont.Dispose();
            _textFont.Dispose();
            _hintFont.Dispose();
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

            public void MeasureLayout()
            {
                using var g = CreateGraphics();

                // 测量标题
                _titleHeight = 0;
                if (!string.IsNullOrEmpty(_owner._currentTitle))
                {
                    var titleSize = g.MeasureString(_owner._currentTitle, _titleFont);
                    _titleHeight = (int)Math.Ceiling(titleSize.Height);
                }

                // 测量各行（支持换行符拆分）
                _layoutLines.Clear();
                _hasLinkMapping = false;
                _lineHeight = (int)Math.Ceiling(g.MeasureString("Ag", _textFont).Height);
                int maxWidth = MinWidth;

                if (_owner._currentLines.Count > 0)
                {
                    foreach (var line in _owner._currentLines)
                    {
                        // 拆分换行符
                        var subLines = line.Text.Split(new[] { '\n' }, StringSplitOptions.None);
                        foreach (var sub in subLines)
                        {
                            _layoutLines.Add((sub, line.IsWarning));
                            var w = (int)Math.Ceiling(g.MeasureString(sub, _textFont).Width);
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
                        var w = (int)Math.Ceiling(g.MeasureString(sub, _textFont).Width);
                        if (w > maxWidth) maxWidth = w;
                    }
                }

                // 底部提示（仅有关联映射行时显示）
                _hintText = _hasLinkMapping ? "右键 → 打开关联表" : "";
                _hintHeight = _hasLinkMapping ? (int)Math.Ceiling(g.MeasureString(_hintText, _hintFont).Height) : 0;

                _owner._dropDownWidth = Math.Min(maxWidth + PaddingH * 2, MaxWidth);
                if (_owner._dropDownWidth < MinWidth) _owner._dropDownWidth = MinWidth;

                int totalHeight = PaddingV;
                if (_titleHeight > 0) totalHeight += _titleHeight + SeparatorMargin + SeparatorHeight + SeparatorMargin;
                if (_layoutLines.Count > 0)
                    totalHeight += _layoutLines.Count * (_lineHeight + LineSpacing) - (_layoutLines.Count > 0 ? LineSpacing : 0);
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
                    g.DrawString(_owner._currentTitle, _titleFont, titleBrush, titleRect);
                    y += _titleHeight + SeparatorMargin;

                    // 分隔线
                    using var sepPen = new Pen(_separatorColor, 1f);
                    g.DrawLine(sepPen, PaddingH, y, Width - PaddingH, y);
                    y += SeparatorHeight + SeparatorMargin;
                }

                // 绘制映射行
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
                        var textSize = g.MeasureString(line.text, _textFont);
                        var textRect = new Rectangle(PaddingH, y, contentWidth, (int)Math.Ceiling(textSize.Height) + LineSpacing);
                        g.DrawString(line.text, _textFont, textBrush, textRect, fmt);
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
                    g.DrawString(_hintText, _hintFont, hintBrush, hintRect);
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
    }
}
