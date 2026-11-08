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
        private List<string>? _currentTableHeaders; // 表模式列头
        private int _lastRow = -1, _lastCol = -1;
        private bool _disposed;
        private int _dropDownWidth;
        /// <summary>由本实例创建、尚未关闭的固定 Tips 窗口（随 Dispose 一起关闭，避免残留孤儿窗口）</summary>
        private readonly List<PinnedTipWindow> _pinnedWindows = new();

        private static readonly Font _titleFont = new("Microsoft YaHei UI", 10f, FontStyle.Bold);
        private static readonly Font _textFont = new("Microsoft YaHei UI", 9f);
        private static readonly Font _hintFont = new("Microsoft YaHei UI", 8f, FontStyle.Italic);

        private static Color _bgColor => Theme.P.Surface;
        private static Color _borderColor => Theme.P.Border;
        private static Color _titleColor => Theme.P.Link;
        private static Color _normalTextColor => Theme.P.Text;
        private static Color _warningTextColor => Theme.Signal(Color.FromArgb(200, 40, 40));
        private static Color _hintColor => Theme.P.Muted;
        private static Color _separatorColor => Theme.P.Divider;

        /// <summary>用于外部提供 tooltip 内容的委托（grid: 当前鼠标所在的 DataGridView）
        /// 返回值新增 tableHeaders: 表模式的列头列表（非表模式返回 null）</summary>
        public Func<DataGridView, int, int, (string? title, List<TipLine> lines, string? rawText, List<string>? tableHeaders)>? GetToolTipContent { get; set; }

        /// <summary>tooltip 中的一行信息</summary>
        public class TipLine
        {
            public string Text { get; set; } = "";
            public bool IsWarning { get; set; }

            /// <summary>是否为表头行（表模式第一行，加粗蓝色文字+浅蓝背景）</summary>
            public bool IsHeader { get; set; }

            /// <summary>表模式：单元格列表（用 \t 分隔的多列数据），非表模式时为 null</summary>
            public List<string>? Cells { get; set; }

            /// <summary>是否为表模式行（Cells != null 且有内容）</summary>
            public bool IsTableRow => Cells != null && Cells.Count > 0;
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

            // Tips 里的行是每次弹出按当前令牌重画的，只有下拉框底色是构造时钉死的，切皮肤要补这一次
            Theme.Changed += OnThemeChanged;
        }

        private void OnThemeChanged()
        {
            if (_disposed) return;
            _dropDown.BackColor = Theme.P.Surface;
            _panel.Invalidate();
        }

        /// <summary>注册一个额外的 DataGridView（如固定行表头）</summary>
        public void RegisterGrid(DataGridView grid)
        {
            if (_grids.Contains(grid)) return;
            _grids.Add(grid);
            grid.ShowCellToolTips = false;
            grid.MouseLeave += (s, e) => Hide();
            grid.CellMouseMove += OnGridCellMouseMove;
            grid.MouseMove += OnGridMouseMove;
            // ★ 鼠标按下时立即隐藏 ToolTip，防止 ToolTip 遮挡右键菜单
            grid.CellMouseDown += (s, e) => Hide();
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

        private void OnGridMouseMove(object? sender, MouseEventArgs e)
        {
            if (sender is DataGridView grid && grid == _activeGrid)
            {
                var hit = grid.HitTest(e.X, e.Y);
                if (hit.Type != DataGridViewHitTestType.Cell)
                {
                    // 鼠标在空白区域，停止延迟定时器，避免定时器到期后显示残留 tooltip
                    _hoverTimer?.Stop();
                    _lastRow = -1;
                    _lastCol = -1;
                }
            }
        }

        private void OnHoverTimerTick(object? sender, EventArgs e)
        {
            _hoverTimer?.Stop();

            // 验证鼠标是否仍在有效单元格上（防止快速划过表格后停在空白处仍显示 tooltip）
            if (_activeGrid != null)
            {
                var pos = _activeGrid.PointToClient(Cursor.Position);
                var hit = _activeGrid.HitTest(pos.X, pos.Y);
                if (hit.Type != DataGridViewHitTestType.Cell)
                {
                    _lastRow = -1;
                    _lastCol = -1;
                    return;
                }
            }

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
            _currentTableHeaders = content.tableHeaders;

            // 测量并计算布局（传入最大高度限制，确保提示语始终可见）
            var mouseScreenPos = Cursor.Position;
            var screen = Screen.FromPoint(new Point(mouseScreenPos.X + 14, mouseScreenPos.Y + 8)).WorkingArea;
            int maxTipHeight = (int)(screen.Height * 0.8);

            _panel.MeasureLayout(maxTipHeight);

            // 计算显示位置（使用鼠标位置，紧跟鼠标，不远离）
            var screenPos = new Point(mouseScreenPos.X + 14, mouseScreenPos.Y + 8);

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
            _currentTableHeaders = null;
            _hoverTimer?.Stop();
            _dropDown.Close();
        }

        /// <summary>创建一个固定的 tooltip 窗口（独立于悬浮 tooltip）</summary>
        public PinnedTipWindow? CreatePinnedWindow(string? title, List<TipLine> lines, string? rawText, Point screenPos, List<string>? tableHeaders = null)
        {
            if (string.IsNullOrEmpty(title) && lines.Count == 0 && string.IsNullOrEmpty(rawText))
                return null;
            // 顺手清理用户已关闭的窗口引用
            _pinnedWindows.RemoveAll(p => p.IsDisposed);
            var pinned = new PinnedTipWindow(title, lines, rawText, screenPos, tableHeaders);
            _pinnedWindows.Add(pinned);
            return pinned;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Theme.Changed -= OnThemeChanged;
            _hoverTimer?.Dispose();
            _dropDown.Dispose();
            _panel.Dispose();
            // 宿主控件关闭时一并关闭固定 Tips，否则 AutoClose=false 的窗口会残留
            foreach (var pinned in _pinnedWindows.ToList())
            {
                try { pinned.Close(); } catch { }
            }
            _pinnedWindows.Clear();
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

        /// <summary>自绘 Panel，负责渲染 tooltip 内容（支持传统行模式和表模式 DataGridView）</summary>
        private class CustomToolTipPanel : Panel
        {
            private readonly DataGridLinkToolTip _owner;
            private int _titleHeight;
            private int _lineHeight;
            private int _hintHeight;
            private List<(string text, bool isWarning)> _layoutLines = new();
            private string _hintText = "";
            private bool _hasLinkMapping; // 是否有关联映射行（控制底部提示显示）
            private List<int> _lineMeasuredHeights = new(); // 每行预计算的绘制高度
            private const int MaxHoverLines = 10; // 悬浮模式最大显示行数

            // ★ 表模式：嵌入 DataGridView
            private bool _isTableMode;             // 是否为表模式
            private DataGridView? _tableGrid;      // 表模式 DataGridView
            private int _tableGridHeight;          // DataGridView 高度
            private List<TipLine> _tableModeLines = new(); // 表模式原始行数据（含 Cells）
            private List<string>? _tableHeaders;   // 表模式列头

            // 布局常量
            private const int PaddingH = 14;
            private const int PaddingV = 10;
            private const int LineSpacing = 3;
            private const int SeparatorHeight = 1;
            private const int SeparatorMargin = 6;
            private const int CornerRadius = 6;
            private const int MinWidth = 160;
            private const int MaxWidth = 520;
            private const int DgvRowHeight = 24;  // DataGridView 行高
            private const int DgvHeaderHeight = 28; // DataGridView 表头高度

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

            /// <summary>测量布局，maxTipHeight 限制最大高度（确保提示语始终可见）</summary>
            public void MeasureLayout(int maxTipHeight = 0)
            {
                using var g = CreateGraphics();

                // ★ 共享变量（避免同一方法不同分支的局部变量名冲突）
                int titleTotal = 0;
                int hintTotal = 0;
                int totalHeight = 0;
                bool truncated = false;

                // ★ 判断是否为表模式
                bool wasTableMode = _isTableMode;
                _isTableMode = _owner._currentLines.Any(l => l.IsTableRow);
                _tableModeLines = _isTableMode ? _owner._currentLines : new List<TipLine>();
                _tableHeaders = _owner._currentTableHeaders;

                // ★ 从表模式切换到非表模式时，清理旧的 DataGridView
                if (wasTableMode && !_isTableMode && _tableGrid != null)
                {
                    Controls.Remove(_tableGrid);
                    _tableGrid.Dispose();
                    _tableGrid = null;
                }

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

                // ★ 表模式：使用 DataGridView 嵌入
                if (_isTableMode)
                {
                    _layoutLines.Clear();
                    _hasLinkMapping = true; // 表模式始终显示底部提示
                    _lineHeight = (int)Math.Ceiling(SafeMeasureString(g, "Ag", _textFont).Height);

                    // 创建/重建 DataGridView
                    BuildTableGrid(g);

                    // 计算宽度：基于 DataGridView 的列宽总和 + padding
                    int dgvTotalWidth = 0;
                    if (_tableGrid != null)
                    {
                        foreach (DataGridViewColumn col in _tableGrid.Columns)
                            dgvTotalWidth += col.Width;
                        dgvTotalWidth += 2; // 边距
                    }
                    maxWidth = Math.Max(maxWidth, dgvTotalWidth);

                    // 计算悬浮模式下 DataGridView 最大行数
                    int maxDataRows = MaxHoverLines;
                    int dataRowsCount = _tableGrid?.Rows.Count ?? 0;
                    truncated = dataRowsCount > maxDataRows;

                    // DataGridView 高度（悬浮模式限制行数，无表头）
                    int visibleRows = Math.Min(dataRowsCount, maxDataRows);
                    _tableGridHeight = visibleRows * DgvRowHeight + 2;

                    // 底部提示
                    _hintText = truncated ? "更多内容 → 右键固定查看" : "右键 → 查看关联表";
                    _hintHeight = (int)Math.Ceiling(SafeMeasureString(g, _hintText, _hintFont).Height);

                    _owner._dropDownWidth = Math.Min(maxWidth + PaddingH * 2, MaxWidth);
                    if (_owner._dropDownWidth < MinWidth) _owner._dropDownWidth = MinWidth;

                    // 计算标题高度
                    titleTotal = 0;
                    if (_titleHeight > 0)
                    {
                        var titleSizeConstrained = SafeMeasureString(g, _owner._currentTitle!, _titleFont, _owner._dropDownWidth - PaddingH * 2);
                        _titleHeight = (int)Math.Ceiling(titleSizeConstrained.Height);
                        titleTotal = _titleHeight + SeparatorMargin + SeparatorHeight + SeparatorMargin;
                    }

                    hintTotal = SeparatorMargin + SeparatorHeight + SeparatorMargin + _hintHeight;

                    // 总高度（不含自绘表头行，表头已作为 DataGridView 第一行）
                    totalHeight = PaddingV + titleTotal + _tableGridHeight + hintTotal + PaddingV;

                    // ★ 如果有高度限制，调整 DataGridView 显示行数
                    if (maxTipHeight > 0)
                    {
                        int fixedH = PaddingV + titleTotal + hintTotal + PaddingV;
                        int availableForGrid = maxTipHeight - fixedH;
                        if (_tableGridHeight > availableForGrid)
                        {
                            // 减少可见行数（无表头）
                            int gridArea = availableForGrid - 2;
                            if (gridArea < DgvRowHeight)
                                gridArea = DgvRowHeight; // 至少显示1行
                            visibleRows = gridArea / DgvRowHeight;
                            if (visibleRows < 1) visibleRows = 1;
                            _tableGridHeight = visibleRows * DgvRowHeight + 2;
                            if (visibleRows < dataRowsCount && string.IsNullOrEmpty(_hintText))
                            {
                                _hintText = "更多内容 → 右键固定查看";
                                _hintHeight = (int)Math.Ceiling(SafeMeasureString(g, _hintText, _hintFont).Height);
                            }
                            totalHeight = PaddingV + titleTotal + _tableGridHeight + hintTotal + PaddingV;
                        }
                    }

                    // 设置 DataGridView 尺寸
                    if (_tableGrid != null)
                    {
                        _tableGrid.Height = _tableGridHeight;
                        _tableGrid.Width = _owner._dropDownWidth - PaddingH * 2;
                    }

                    Height = totalHeight;
                    return;
                }

                // ========== 传统行模式（以下为原有逻辑） ==========

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

                // 悬浮模式：截取前 MaxHoverLines 行，超出部分不显示（右键固定窗口可查看全部）
                truncated = _layoutLines.Count > MaxHoverLines;
                if (truncated)
                    _layoutLines = _layoutLines.Take(MaxHoverLines).ToList();

                // 底部提示（有关联映射行或截断时显示）
                _hintText = _hasLinkMapping ? "右键 → 查看关联表" : (truncated ? "更多内容 → 右键固定查看" : "");
                _hintHeight = (!string.IsNullOrEmpty(_hintText)) ? (int)Math.Ceiling(SafeMeasureString(g, _hintText, _hintFont).Height) : 0;

                _owner._dropDownWidth = Math.Min(maxWidth + PaddingH * 2, MaxWidth);
                if (_owner._dropDownWidth < MinWidth) _owner._dropDownWidth = MinWidth;

                // 计算高度时，对每行文本使用最大宽度约束测量，支持自动换行
                int contentMaxWidth = _owner._dropDownWidth - PaddingH * 2;

                // 计算每行的实际高度（约束宽度测量）
                _lineMeasuredHeights.Clear();
                foreach (var line in _layoutLines)
                {
                    var constrainedSize = SafeMeasureString(g, line.text, _textFont, contentMaxWidth);
                    _lineMeasuredHeights.Add((int)Math.Ceiling(constrainedSize.Height));
                }

                // 计算标题高度
                titleTotal = 0;
                if (_titleHeight > 0)
                {
                    var titleSizeConstrained = SafeMeasureString(g, _owner._currentTitle!, _titleFont, contentMaxWidth);
                    _titleHeight = (int)Math.Ceiling(titleSizeConstrained.Height);
                    titleTotal = _titleHeight + SeparatorMargin + SeparatorHeight + SeparatorMargin;
                }

                // 计算提示语区域高度
                hintTotal = 0;
                if (!string.IsNullOrEmpty(_hintText))
                    hintTotal = SeparatorMargin + SeparatorHeight + SeparatorMargin + _hintHeight;

                // ★ 如果有高度限制，从后往前减少内容行数，确保提示语始终可见
                if (maxTipHeight > 0)
                {
                    // 固定部分：上边距 + 标题 + 下边距 + 提示语区域
                    int fixedHeight = PaddingV + titleTotal + PaddingV + hintTotal;
                    int availableForLines = maxTipHeight - fixedHeight;

                    // 从后往前裁剪行，直到内容行总高度 <= availableForLines
                    while (_layoutLines.Count > 1 && _lineMeasuredHeights.Count > 1)
                    {
                        int linesTotalHeight = _lineMeasuredHeights.Sum() + (_layoutLines.Count - 1) * LineSpacing;
                        if (linesTotalHeight <= availableForLines) break;

                        // 移除最后一行
                        _layoutLines.RemoveAt(_layoutLines.Count - 1);
                        _lineMeasuredHeights.RemoveAt(_lineMeasuredHeights.Count - 1);

                        // 被裁剪了，需要显示提示语
                        if (string.IsNullOrEmpty(_hintText))
                        {
                            _hintText = "更多内容 → 右键固定查看";
                            _hintHeight = (int)Math.Ceiling(SafeMeasureString(g, _hintText, _hintFont).Height);
                            hintTotal = SeparatorMargin + SeparatorHeight + SeparatorMargin + _hintHeight;
                            // 重新计算可用空间（提示语区域高度可能变化）
                            availableForLines = maxTipHeight - PaddingV - titleTotal - PaddingV - hintTotal;
                        }
                    }
                }

                // 计算最终总高度
                totalHeight = PaddingV;
                totalHeight += titleTotal;
                if (_layoutLines.Count > 0)
                {
                    for (int i = 0; i < _lineMeasuredHeights.Count; i++)
                        totalHeight += _lineMeasuredHeights[i] + LineSpacing;
                    totalHeight -= LineSpacing; // 最后一行不需要额外的 LineSpacing
                }
                totalHeight += hintTotal;
                totalHeight += PaddingV;

                Height = totalHeight;
            }

            /// <summary>创建/重建表模式的 DataGridView（透明风格，无边框）</summary>
            private void BuildTableGrid(Graphics g)
            {
                // 清理旧的 DataGridView
                if (_tableGrid != null)
                {
                    Controls.Remove(_tableGrid);
                    _tableGrid.Dispose();
                    _tableGrid = null;
                }

                if (_tableModeLines.Count == 0) return;

                // 确定列数：从 Cells 列表获取
                int colCount = _tableModeLines.Max(l => l.Cells?.Count ?? 0);
                if (colCount == 0) return;

                // 确定列头
                var headers = new List<string>();
                if (_tableHeaders != null && _tableHeaders.Count >= colCount)
                    headers = _tableHeaders;
                else
                {
                    for (int i = 0; i < colCount; i++)
                        headers.Add($"列{i}");
                }

                // 排序行：警告行（IsWarning）在最前面
                var orderedLines = _tableModeLines
                    .Where(l => l.IsWarning)
                    .Concat(_tableModeLines.Where(l => !l.IsWarning))
                    .ToList();

                // ★ 如果有自定义表头，作为第一行插入（这样列宽自然与内容对齐）
                if (_tableHeaders != null && _tableHeaders.Count > 0)
                {
                    orderedLines.Insert(0, new TipLine
                    {
                        Cells = new List<string>(_tableHeaders),
                        IsWarning = false,
                        Text = "",
                        IsHeader = true // 标记为表头行
                    });
                }

                // 创建 DataGridView（完全去表格化：无行号、无表头、无网格线、纯色背景）
                _tableGrid = new DataGridView
                {
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AllowUserToResizeRows = false,
                    AllowUserToResizeColumns = false,
                    ReadOnly = true,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = false,
                    RowHeadersVisible = false,
                    ColumnHeadersVisible = false,
                    RowTemplate = new DataGridViewRow { Height = DgvRowHeight },
                    BorderStyle = BorderStyle.None,
                    BackgroundColor = _bgColor,
                    CellBorderStyle = DataGridViewCellBorderStyle.None,
                    EnableHeadersVisualStyles = false,
                    ScrollBars = ScrollBars.None,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                    Font = _textFont,
                };

                // 统一纯色背景单元格样式（无网格线、无交替色）
                var cellStyle = new DataGridViewCellStyle
                {
                    BackColor = _bgColor,
                    ForeColor = _normalTextColor,
                    Font = _textFont,
                    Alignment = DataGridViewContentAlignment.MiddleLeft,
                    Padding = new Padding(2, 0, 2, 0),
                    SelectionBackColor = _bgColor,
                    SelectionForeColor = _normalTextColor,
                };
                _tableGrid.DefaultCellStyle = cellStyle;

                // 创建列
                for (int i = 0; i < colCount; i++)
                {
                    var col = new DataGridViewTextBoxColumn
                    {
                        HeaderText = headers[i],
                        Name = $"Col{i}",
                        Resizable = DataGridViewTriState.False,
                        SortMode = DataGridViewColumnSortMode.NotSortable,
                    };
                    _tableGrid.Columns.Add(col);
                }

                // 填充数据行
                foreach (var line in orderedLines)
                {
                    var row = new DataGridViewRow();
                    row.Height = DgvRowHeight;

                    for (int i = 0; i < colCount; i++)
                    {
                        var cellVal = (line.Cells != null && i < line.Cells.Count) ? line.Cells[i] : "-";
                        var cell = new DataGridViewTextBoxCell
                        {
                            Value = cellVal,
                        };

                        // ★ 表头行：加粗蓝色文字+浅蓝背景
                        if (line.IsHeader)
                        {
                            var headerStyle = new DataGridViewCellStyle(cellStyle)
                            {
                                Font = new Font(_textFont, FontStyle.Bold),
                                ForeColor = _titleColor,
                                BackColor = Theme.CellTint(Color.FromArgb(235, 245, 255)),
                                SelectionForeColor = _titleColor,
                                SelectionBackColor = Theme.CellTint(Color.FromArgb(235, 245, 255)),
                            };
                            cell.Style = headerStyle;
                        }
                        // ★ 警告行：红色文字+浅红背景
                        else if (line.IsWarning)
                        {
                            var warningStyle = new DataGridViewCellStyle(cellStyle)
                            {
                                ForeColor = _warningTextColor,
                                BackColor = Theme.CellTint(Color.FromArgb(255, 240, 240)),
                                SelectionForeColor = _warningTextColor,
                                SelectionBackColor = Theme.CellTint(Color.FromArgb(255, 240, 240)),
                            };
                            cell.Style = warningStyle;
                        }

                        row.Cells.Add(cell);
                    }

                    _tableGrid.Rows.Add(row);
                }

                // ★ 动态计算列宽（基于最大内容宽度）
                for (int i = 0; i < colCount; i++)
                {
                    int maxColWidth = (int)Math.Ceiling(SafeMeasureString(g, headers[i], _textFont).Width) + 8;
                    foreach (DataGridViewRow row in _tableGrid.Rows)
                    {
                        var cellVal = row.Cells[i].Value?.ToString() ?? "";
                        int w = (int)Math.Ceiling(SafeMeasureString(g, cellVal, _textFont).Width) + 8;
                        if (w > maxColWidth) maxColWidth = w;
                    }
                    _tableGrid.Columns[i].Width = Math.Max(40, Math.Min(maxColWidth, 200));
                }

                Controls.Add(_tableGrid);
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

                // ★ 表模式：定位 DataGridView（表头已作为第一行数据嵌入，无需自绘）
                if (_isTableMode && _tableGrid != null)
                {
                    _tableGrid.Location = new Point(PaddingH, y);
                    _tableGrid.Width = contentWidth;
                    _tableGrid.Height = _tableGridHeight;
                    y += _tableGridHeight;
                }

                // ★ 传统行模式：绘制映射行（使用预计算高度，确保与 MeasureLayout 一致）
                if (!_isTableMode && _layoutLines.Count > 0)
                {
                    for (int i = 0; i < _layoutLines.Count; i++)
                    {
                        var line = _layoutLines[i];
                        var color = line.isWarning ? _warningTextColor : _normalTextColor;
                        using var textBrush = new SolidBrush(color);
                        var fmt = new StringFormat
                        {
                            Trimming = StringTrimming.None,
                            FormatFlags = 0
                        };
                        int lineH = _lineMeasuredHeights.Count > i ? _lineMeasuredHeights[i] : _lineHeight;
                        var textRect = new Rectangle(PaddingH, y, contentWidth, lineH);
                        SafeDrawString(g, line.text, _textFont, textBrush, textRect, fmt);
                        y += lineH + LineSpacing;
                    }
                }

                // 底部提示（有关联映射行或截断时显示）
                if (!string.IsNullOrEmpty(_hintText))
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

        /// <summary>固定的 tooltip 窗口（独立于悬浮 tooltip，支持多个并存，带滚动条）</summary>
        public class PinnedTipWindow : IDisposable
        {
            private readonly ToolStripDropDown _dropDown;
            private readonly PinnedTipPanel _panel;
            private bool _disposed;

            /// <summary>是否已关闭（供创建方清理引用）</summary>
            public bool IsDisposed => _disposed;

            // 关闭按钮
            private const int CloseBtnSize = 20;
            private const int CloseBtnMargin = 6;

            private static readonly Font _titleFont2 = new("Microsoft YaHei UI", 10f, FontStyle.Bold);
            private static readonly Font _textFont2 = new("Microsoft YaHei UI", 9f);

            private static Color _bgColor2 => Theme.P.Surface;
            private static Color _borderColor2 => Theme.P.Border;
            private static Color _titleColor2 => Theme.P.Link;
            private static Color _normalTextColor2 => Theme.P.Text;
            private static Color _warningTextColor2 => Theme.Signal(Color.FromArgb(200, 40, 40));
            private static Color _separatorColor2 => Theme.P.Divider;

            private const int PaddingH = 14;
            private const int PaddingV = 10;
            private const int LineSpacing = 3;
            private const int SeparatorHeight = 1;
            private const int SeparatorMargin = 6;
            private const int CornerRadius = 6;
            private const int MinWidth = 160;
            private const int MaxWidth = 520;

            public PinnedTipWindow(string? title, List<TipLine> lines, string? rawText, Point screenPos, List<string>? tableHeaders = null)
            {
                _panel = new PinnedTipPanel(this, title, lines, rawText, tableHeaders);

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

                // 屏幕边界调整
                var screen = Screen.FromPoint(screenPos).WorkingArea;

                // 限制最大可见高度（不超过 450px）
                int maxVisibleHeight = Math.Min(450, (int)(screen.Height * 0.7));
                _panel.SetupScroll(maxVisibleHeight);

                _panel.Size = new Size(_panel._measuredWidth, _panel._visibleHeight);
                host.Size = _panel.Size;
                _dropDown.Size = _panel.Size;

                // X 修正：优先右侧，不够则翻转到左侧
                if (screenPos.X + _panel._measuredWidth > screen.Right)
                {
                    int leftPos = screenPos.X - _panel._measuredWidth - 6;
                    screenPos.X = leftPos >= screen.Left ? leftPos : screen.Right - _panel._measuredWidth - 5;
                }
                if (screenPos.X < screen.Left) screenPos.X = screen.Left + 5;

                // Y 修正：优先下方，不够则翻转到上方
                if (screenPos.Y + _panel._visibleHeight > screen.Bottom)
                    screenPos.Y = Math.Max(screen.Top, screen.Bottom - _panel._visibleHeight - 5);

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

            /// <summary>自绘 Panel —— 固定窗口的内容面板（带关闭按钮、拖动支持和滚动条）
            /// 支持传统行模式和表模式 DataGridView</summary>
            private class PinnedTipPanel : Panel
            {
                private readonly PinnedTipWindow _owner;
                private readonly string? _title;
                private readonly List<TipLine> _lines;
                private readonly string? _rawText;
                private readonly List<string>? _tableHeaders;

                private List<(string text, bool isWarning)> _layoutLines = new();
                private List<int> _lineYPositions = new(); // 每行的 Y 轴起始位置（像素）
                private List<int> _lineHeights = new();    // 每行的实际绘制高度（像素）
                private int _titleHeight;
                private int _lineHeight;
                private Rectangle _closeBtnRect;
                private bool _closeHovered;
                private Point _dragStart;
                private bool _dragging;

                // ★ 表模式
                private bool _isTableMode;
                private DataGridView? _tableGrid;      // 表模式 DataGridView
                private int _tableGridHeight;          // DataGridView 高度
                private List<TipLine> _tableModeLines = new();

                // 滚动相关
                private VScrollBar? _vScrollBar;
                private int _scrollOffset = 0;        // 当前滚动偏移量（像素）
                private int _contentHeight = 0;       // 内容总高度（像素）
                private bool _needsScroll = false;    // 是否需要滚动条
                private const int ScrollBarWidth = 18;
                private const int MaxVisibleLinesBeforeScroll = 15; // 超过此行数后启用滚动

                public int _measuredWidth;
                public int _measuredHeight;
                public int _visibleHeight;  // 可见区域高度（窗口实际显示的高度）
                private const int DgvRowHeight = 24;
                private const int DgvHeaderHeight = 28;

                public PinnedTipPanel(PinnedTipWindow owner, string? title, List<TipLine> lines, string? rawText, List<string>? tableHeaders)
                {
                    _owner = owner;
                    _title = title;
                    _lines = lines ?? new List<TipLine>();
                    _rawText = rawText;
                    _tableHeaders = tableHeaders;
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

                /// <summary>安全测量字符串（指定最大宽度，用于计算自动换行后的高度）</summary>
                private static SizeF SafeMeasureStringPinned(Graphics g, string text, Font font, int maxWidth)
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

                public void MeasureSize()
                {
                    using var g = CreateGraphics();

                    // ★ 共享变量
                    int totalWidth;
                    int rightMargin;
                    int y;

                    // ★ 判断是否为表模式
                    bool wasTableMode = _isTableMode;
                    _isTableMode = _lines.Any(l => l.IsTableRow);
                    _tableModeLines = _isTableMode ? _lines : new List<TipLine>();

                    // ★ 从表模式切换到非表模式时清理旧 DataGridView
                    if (wasTableMode && !_isTableMode && _tableGrid != null)
                    {
                        Controls.Remove(_tableGrid);
                        _tableGrid.Dispose();
                        _tableGrid = null;
                    }

                    _titleHeight = 0;
                    if (!string.IsNullOrEmpty(_title))
                    {
                        var titleSize = SafeMeasureStringPinned(g, _title, _titleFont2);
                        _titleHeight = (int)Math.Ceiling(titleSize.Height);
                    }

                    _layoutLines.Clear();
                    _lineYPositions.Clear();
                    _lineHeights.Clear();
                    _lineHeight = (int)Math.Ceiling(SafeMeasureStringPinned(g, "Ag", _textFont2).Height);
                    int maxWidth = MinWidth;

                    // 为关闭按钮留出右侧空间
                    int titleRightMargin = CloseBtnMargin + CloseBtnSize + CloseBtnMargin;
                    // 滚动条占用宽度（可能需要）
                    int scrollBarRightMargin = ScrollBarWidth + 2;

                    // ★ 表模式：使用 DataGridView
                    if (_isTableMode)
                    {
                        // 创建 DataGridView
                        BuildPinnedTableGrid(g);

                        // 计算宽度
                        if (_tableGrid != null)
                        {
                            int dgvTotalWidth = 0;
                            foreach (DataGridViewColumn col in _tableGrid.Columns)
                                dgvTotalWidth += col.Width;
                            dgvTotalWidth += 2;
                            maxWidth = Math.Max(maxWidth, dgvTotalWidth);
                        }

                        // 固定窗口：DataGridView 自身处理滚动，高度不超过 450px
                        int dataRowsCount = _tableGrid?.Rows.Count ?? 0;
                        int fullTableHeight = dataRowsCount * DgvRowHeight + 2;
                        _tableGridHeight = Math.Min(fullTableHeight, 450);

                        _needsScroll = false; // 表模式用 DataGridView 自身滚动

                        rightMargin = titleRightMargin;

                        if (!string.IsNullOrEmpty(_title))
                        {
                            int titleW = (int)Math.Ceiling(SafeMeasureStringPinned(g, _title, _titleFont2).Width) + PaddingH + rightMargin;
                            if (titleW > maxWidth) maxWidth = titleW;
                        }

                        totalWidth = Math.Min(maxWidth + PaddingH * 2, MaxWidth);
                        if (totalWidth < MinWidth) totalWidth = MinWidth;
                        _measuredWidth = totalWidth;

                        // 设置 DataGridView 尺寸
                        if (_tableGrid != null)
                        {
                            _tableGrid.Width = _measuredWidth - PaddingH * 2;
                            _tableGrid.Height = _tableGridHeight;
                        }

                        // 计算总高度（表头已作为 DataGridView 第一行，无需额外空间）
                        y = PaddingV;
                        if (_titleHeight > 0)
                        {
                            y += _titleHeight + SeparatorMargin + SeparatorHeight + SeparatorMargin;
                        }
                        y += _tableGridHeight;
                        y += PaddingV;

                        _contentHeight = y;
                        _measuredHeight = y;
                        return;
                    }

                    // ========== 传统行模式（原有逻辑） ==========

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

                    // 判断是否需要滚动条
                    _needsScroll = _layoutLines.Count > MaxVisibleLinesBeforeScroll;

                    // 确保宽度至少能容纳标题+关闭按钮（如果有滚动条也需考虑）
                    rightMargin = titleRightMargin;
                    if (_needsScroll) rightMargin = Math.Max(titleRightMargin, scrollBarRightMargin);

                    if (!string.IsNullOrEmpty(_title))
                    {
                        int titleW = (int)Math.Ceiling(SafeMeasureStringPinned(g, _title, _titleFont2).Width) + PaddingH + rightMargin;
                        if (titleW > maxWidth) maxWidth = titleW;
                    }
                    else if (maxWidth < MinWidth)
                    {
                        maxWidth = MinWidth;
                    }

                    // 有滚动条时宽度增加 ScrollBarWidth
                    totalWidth = Math.Min(maxWidth + PaddingH * 2, MaxWidth);
                    if (_needsScroll) totalWidth += ScrollBarWidth;
                    if (totalWidth < MinWidth) totalWidth = MinWidth;
                    _measuredWidth = totalWidth;

                    // 计算内容区域宽度（用于文本换行测量）
                    int contentMaxWidth = _measuredWidth - PaddingH * 2;
                    if (_needsScroll) contentMaxWidth -= ScrollBarWidth;

                    // 计算每行的 Y 位置和高度，以及总内容高度
                    y = PaddingV;
                    if (_titleHeight > 0)
                    {
                        // 标题也考虑换行
                        var titleSizeConstrained = SafeMeasureStringPinned(g, _title!, _titleFont2, contentMaxWidth);
                        _titleHeight = (int)Math.Ceiling(titleSizeConstrained.Height);
                        y += _titleHeight + SeparatorMargin + SeparatorHeight + SeparatorMargin;
                    }

                    foreach (var line in _layoutLines)
                    {
                        var constrainedSize = SafeMeasureStringPinned(g, line.text, _textFont2, contentMaxWidth);
                        int h = (int)Math.Ceiling(constrainedSize.Height);
                        _lineYPositions.Add(y);
                        _lineHeights.Add(h);
                        y += h + LineSpacing;
                    }
                    if (_layoutLines.Count > 0)
                        y -= LineSpacing; // 去掉最后一行多余的 LineSpacing

                    y += PaddingV;

                    _contentHeight = y;
                    _measuredHeight = y;
                }

                /// <summary>创建固定窗口的表模式 DataGridView（支持自身滚动）</summary>
                private void BuildPinnedTableGrid(Graphics g)
                {
                    // 清理旧 DataGridView
                    if (_tableGrid != null)
                    {
                        Controls.Remove(_tableGrid);
                        _tableGrid.Dispose();
                        _tableGrid = null;
                    }

                    if (_tableModeLines.Count == 0) return;

                    int colCount = _tableModeLines.Max(l => l.Cells?.Count ?? 0);
                    if (colCount == 0) return;

                    // 确定列头
                    var headers = new List<string>();
                    if (_tableHeaders != null && _tableHeaders.Count >= colCount)
                        headers = _tableHeaders;
                    else
                    {
                        for (int i = 0; i < colCount; i++)
                            headers.Add($"列{i}");
                    }

                    // 排序行：警告行在最前面
                    var orderedLines = _tableModeLines
                        .Where(l => l.IsWarning)
                        .Concat(_tableModeLines.Where(l => !l.IsWarning))
                        .ToList();

                    // ★ 如果有自定义表头，作为第一行插入
                    if (_tableHeaders != null && _tableHeaders.Count > 0)
                    {
                        orderedLines.Insert(0, new TipLine
                        {
                            Cells = new List<string>(_tableHeaders),
                            IsWarning = false,
                            Text = "",
                            IsHeader = true
                        });
                    }

                    // 固定模式 DataGridView 最大高度 450px（纯内容区域，不含标题）
                    int maxDataGridViewHeight = 450;

                    // 创建 DataGridView（支持垂直滚动，完全去表格化：无行号、无表头、无网格线、纯色背景）
                    _tableGrid = new DataGridView
                    {
                        AllowUserToAddRows = false,
                        AllowUserToDeleteRows = false,
                        AllowUserToResizeRows = false,
                        AllowUserToResizeColumns = false,
                        ReadOnly = true,
                        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                        MultiSelect = false,
                        RowHeadersVisible = false,
                        ColumnHeadersVisible = false,
                        RowTemplate = new DataGridViewRow { Height = DgvRowHeight },
                        BorderStyle = BorderStyle.None,
                        BackgroundColor = _bgColor2,
                        CellBorderStyle = DataGridViewCellBorderStyle.None,
                        EnableHeadersVisualStyles = false,
                        ScrollBars = ScrollBars.Vertical,
                        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                        Font = _textFont2,
                    };

                    // 统一纯色背景单元格样式（无网格线、无交替色）
                    var cellStyle = new DataGridViewCellStyle
                    {
                        BackColor = _bgColor2,
                        ForeColor = _normalTextColor2,
                        Font = _textFont2,
                        Alignment = DataGridViewContentAlignment.MiddleLeft,
                        Padding = new Padding(2, 0, 2, 0),
                        SelectionBackColor = _bgColor2,
                        SelectionForeColor = _normalTextColor2,
                    };
                    _tableGrid.DefaultCellStyle = cellStyle;

                    // 创建列
                    for (int i = 0; i < colCount; i++)
                    {
                        var col = new DataGridViewTextBoxColumn
                        {
                            HeaderText = headers[i],
                            Name = $"Col{i}",
                            Resizable = DataGridViewTriState.False,
                            SortMode = DataGridViewColumnSortMode.NotSortable,
                        };
                        _tableGrid.Columns.Add(col);
                    }

                    // 填充数据行
                    foreach (var line in orderedLines)
                    {
                        var row = new DataGridViewRow();
                        row.Height = DgvRowHeight;

                        for (int i = 0; i < colCount; i++)
                        {
                            var cellVal = (line.Cells != null && i < line.Cells.Count) ? line.Cells[i] : "";
                            var cell = new DataGridViewTextBoxCell { Value = cellVal };

                            // ★ 表头行：加粗蓝色文字+浅蓝背景
                            if (line.IsHeader)
                            {
                                var headerStyle = new DataGridViewCellStyle(cellStyle)
                                {
                                    Font = new Font(_textFont2, FontStyle.Bold),
                                    ForeColor = _titleColor2,
                                    BackColor = Theme.CellTint(Color.FromArgb(235, 245, 255)),
                                    SelectionForeColor = _titleColor2,
                                    SelectionBackColor = Theme.CellTint(Color.FromArgb(235, 245, 255)),
                                };
                                cell.Style = headerStyle;
                            }
                            else if (line.IsWarning)
                            {
                                var warningStyle = new DataGridViewCellStyle(cellStyle)
                                {
                                    ForeColor = _warningTextColor2,
                                    BackColor = Theme.CellTint(Color.FromArgb(255, 240, 230)),
                                };
                                cell.Style = warningStyle;
                            }

                            row.Cells.Add(cell);
                        }

                        _tableGrid.Rows.Add(row);
                    }

                    // 动态计算列宽
                    for (int i = 0; i < colCount; i++)
                    {
                        int maxColWidth = (int)Math.Ceiling(SafeMeasureStringPinned(g, headers[i], _textFont2).Width) + 8;
                        foreach (DataGridViewRow row in _tableGrid.Rows)
                        {
                            var cellVal = row.Cells[i].Value?.ToString() ?? "";
                            int w = (int)Math.Ceiling(SafeMeasureStringPinned(g, cellVal, _textFont2).Width) + 8;
                            if (w > maxColWidth) maxColWidth = w;
                        }
                        _tableGrid.Columns[i].Width = Math.Max(40, Math.Min(maxColWidth, 200));
                    }

                    Controls.Add(_tableGrid);
                }

                /// <summary>设置滚动：根据最大可见高度决定是否启用滚动条</summary>
                public void SetupScroll(int maxVisibleHeight)
                {
                    if (!_needsScroll)
                    {
                        // 不需要滚动，直接显示全部内容
                        _visibleHeight = _measuredHeight;
                        return;
                    }

                    // 限制可见高度不超过 maxVisibleHeight
                    _visibleHeight = Math.Min(_measuredHeight, maxVisibleHeight);

                    // 创建滚动条
                    if (_vScrollBar == null)
                    {
                        _vScrollBar = new VScrollBar
                        {
                            Width = ScrollBarWidth,
                            SmallChange = _lineHeight + LineSpacing, // 单行滚动
                            LargeChange = (_lineHeight + LineSpacing) * 5, // 5行滚动
                        };
                        _vScrollBar.Scroll += OnScrollBarScroll;
                        Controls.Add(_vScrollBar);
                    }

                    // 设置滚动条范围和位置
                    int scrollRange = _contentHeight - _visibleHeight;
                    if (scrollRange < 0) scrollRange = 0;
                    _vScrollBar.Minimum = 0;
                    _vScrollBar.Maximum = scrollRange + _vScrollBar.LargeChange - 1;
                    _vScrollBar.Value = 0;
                    _vScrollBar.Height = _visibleHeight - PaddingV * 2;
                    _vScrollBar.Location = new Point(_measuredWidth - ScrollBarWidth - 2, PaddingV);
                    _vScrollBar.Visible = true;

                    _scrollOffset = 0;
                }

                private void OnScrollBarScroll(object? sender, ScrollEventArgs e)
                {
                    if (_vScrollBar == null) return;
                    int newOffset = _vScrollBar.Value;
                    if (newOffset != _scrollOffset)
                    {
                        _scrollOffset = newOffset;
                        Invalidate();
                    }
                }

                protected override void OnMouseWheel(MouseEventArgs e)
                {
                    if (!_needsScroll || _vScrollBar == null) return;

                    int delta = e.Delta > 0 ? -(_lineHeight + LineSpacing) * 3 : (_lineHeight + LineSpacing) * 3;
                    int newOffset = _scrollOffset + delta;
                    int maxOffset = _contentHeight - _visibleHeight;
                    newOffset = Math.Max(0, Math.Min(newOffset, maxOffset));

                    if (newOffset != _scrollOffset)
                    {
                        _scrollOffset = newOffset;
                        _vScrollBar.Value = newOffset;
                        Invalidate();
                    }

                    base.OnMouseWheel(e);
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
                    int closeBtnX = Width - CloseBtnMargin - CloseBtnSize;
                    _closeBtnRect = new Rectangle(closeBtnX, CloseBtnMargin, CloseBtnSize, CloseBtnSize);
                    using (var cbBgBrush = new SolidBrush(_closeHovered ? Theme.CellTint(Color.FromArgb(240, 220, 220)) : Theme.P.Surface))
                    {
                        g.FillRectangle(cbBgBrush, _closeBtnRect);
                    }
                    using (var cbBorderPen = new Pen(_closeHovered ? Theme.Signal(Color.FromArgb(220, 80, 80)) : Theme.P.Border, 1f))
                    {
                        g.DrawRectangle(cbBorderPen, _closeBtnRect);
                    }
                    // 绘制 X
                    int pad = 5;
                    using var xPen = new Pen(_closeHovered ? Theme.Signal(Color.FromArgb(200, 50, 50)) : Theme.P.Muted, 1.8f);
                    g.DrawLine(xPen, _closeBtnRect.Left + pad, _closeBtnRect.Top + pad,
                               _closeBtnRect.Right - pad, _closeBtnRect.Bottom - pad);
                    g.DrawLine(xPen, _closeBtnRect.Right - pad, _closeBtnRect.Top + pad,
                               _closeBtnRect.Left + pad, _closeBtnRect.Bottom - pad);

                    // 内容区域宽度
                    int contentWidth = Width - PaddingH * 2;
                    int titleMaxWidth = contentWidth - CloseBtnMargin - CloseBtnSize - CloseBtnMargin;
                    if (titleMaxWidth < 40) titleMaxWidth = 40;

                    g.SetClip(new Rectangle(0, 0, Width, Height));

                    int y = PaddingV - _scrollOffset;

                    // 绘制标题
                    if (!string.IsNullOrEmpty(_title))
                    {
                        var titleRect = new Rectangle(PaddingH, y, titleMaxWidth, _titleHeight);
                        using var titleBrush = new SolidBrush(_titleColor2);
                        var titleFmt = new StringFormat { Trimming = StringTrimming.EllipsisCharacter };
                        CustomToolTipPanel.SafeDrawString(g, _title, _titleFont2, titleBrush, titleRect, titleFmt);
                        y += _titleHeight + SeparatorMargin;

                        using var sepPen = new Pen(_separatorColor2, 1f);
                        g.DrawLine(sepPen, PaddingH, y, PaddingH + contentWidth, y);
                        y += SeparatorHeight + SeparatorMargin;
                    }

                    // ★ 表模式：定位 DataGridView（表头已作为第一行嵌入，无需自绘）
                    if (_isTableMode && _tableGrid != null)
                    {
                        _tableGrid.Location = new Point(PaddingH, y);
                        // DataGridView 自身处理滚动，不需要 OnPaint 滚动偏移
                    }
                    else if (_layoutLines.Count > 0)
                    {
                        // 传统行模式：绘制内容行（使用预计算的 Y 位置和高度，通过 _scrollOffset 偏移）
                        for (int i = 0; i < _layoutLines.Count; i++)
                        {
                            int lineY = _lineYPositions[i] - _scrollOffset;
                            int lineH = _lineHeights[i];

                            // 跳过完全在可见区域之外的行（性能优化）
                            if (lineY + lineH + LineSpacing < PaddingV) continue;
                            if (lineY > Height - PaddingV) break;

                            var line = _layoutLines[i];
                            var color = line.isWarning ? _warningTextColor2 : _normalTextColor2;
                            using var textBrush = new SolidBrush(color);
                            var fmt = new StringFormat
                            {
                                Trimming = StringTrimming.None,
                                FormatFlags = 0
                            };
                            var textRect = new Rectangle(PaddingH, lineY, contentWidth, lineH + LineSpacing);
                            CustomToolTipPanel.SafeDrawString(g, line.text, _textFont2, textBrush, textRect, fmt);
                        }
                    }

                    g.ResetClip();
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

                protected override void Dispose(bool disposing)
                {
                    if (disposing && _vScrollBar != null)
                    {
                        _vScrollBar.Dispose();
                        _vScrollBar = null;
                    }
                    base.Dispose(disposing);
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
