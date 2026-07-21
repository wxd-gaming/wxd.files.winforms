using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 自定义虚拟化行列表控件，完全自绘，无 Win32 ListView HRGN 裁剪限制。
    /// 支持百万/亿级行（仅绘制可见行）、选中高亮、双击、Ctrl+C 复制、水平/垂直滚动。
    /// </summary>
    public class LineListView : Control
    {
        // ==================== 状态 ====================
        private long _totalLines;
        private int _maxLineWidth;          // 内容总宽（像素），由外层测量最长行后设置
        private int _vPos;                   // 垂直滚动位置（首行索引）
        private int _hPos;                   // 水平滚动位置（像素）
        private readonly HashSet<long> _selected = new();
        private long _anchorIndex = -1;      // Shift 多选锚点

        // ==================== 外观 ====================
        private const int RowHeightPadding = 6;
        private const int LeftMargin = 4;
        private static readonly Color SelectionBackColor = Color.FromArgb(66, 165, 245);
        private static readonly Color SelectionTextColor = Color.White;
        private static readonly Color BackColorDefault = Color.White;
        private static readonly Color ForeColorDefault = Color.FromArgb(30, 30, 30);

        // ==================== 滚动条 ====================
        private readonly VScrollBar _vScroll;
        private readonly HScrollBar _hScroll;
        private const int MAX_SCROLL_RANGE = 10 * 1024 * 1024;

        // ==================== 回调 ====================
        /// <summary>按需获取行文本的回调（外层提供缓存/读取逻辑）</summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Func<long, string>? GetLineText { get; set; }

        /// <summary>选中行变化时触发</summary>
        public event EventHandler? SelectedIndexChanged;

        /// <summary>双击行时触发，参数为行索引</summary>
        public event Action<long>? ItemDoubleClick;

        /// <summary>垂直滚动位置变化时触发（外层用于预读缓存）</summary>
        public event EventHandler? VScrollChanged;

        public LineListView()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer
                   | ControlStyles.UserPaint
                   | ControlStyles.AllPaintingInWmPaint
                   | ControlStyles.ResizeRedraw, true);
            BackColor = BackColorDefault;
            ForeColor = ForeColorDefault;
            Font = new Font("Consolas", 10F);
            DoubleBuffered = true;

            _vScroll = new VScrollBar { Dock = DockStyle.Right, Visible = false };
            _hScroll = new HScrollBar { Dock = DockStyle.Bottom, Visible = false };
            _vScroll.Scroll += OnVScroll;
            _hScroll.Scroll += OnHScroll;
            Controls.Add(_vScroll);
            Controls.Add(_hScroll);
        }

        // ==================== 公开 API ====================
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public long TotalLines => _totalLines;
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int MaxLineWidth => _maxLineWidth;
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int VPos => _vPos;
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int HPos => _hPos;
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int RowHeight => Font.Height + RowHeightPadding;

        /// <summary>当前选中行集合（只读视图）</summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public IReadOnlyCollection<long> SelectedIndices => _selected;

        /// <summary>设置总行数，更新垂直滚动条</summary>
        public void SetTotalLines(long count)
        {
            _totalLines = Math.Max(0, count);
            ClampVPos();
            UpdateVScroll();
            Invalidate();
        }

        /// <summary>设置最大行宽（像素），更新水平滚动条</summary>
        public void SetMaxLineWidth(int width)
        {
            _maxLineWidth = Math.Max(0, Math.Min(width, MAX_SCROLL_RANGE));
            ClampHPos();
            UpdateHScroll();
            Invalidate();
        }

        /// <summary>选中单个行并滚动到可见</summary>
        public void SelectRow(long index)
        {
            if (index < 0 || index >= _totalLines) return;
            _selected.Clear();
            _selected.Add(index);
            _anchorIndex = index;
            EnsureVisible(index);
            SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
            Invalidate();
        }

        public void ClearSelection()
        {
            if (_selected.Count == 0) return;
            _selected.Clear();
            _anchorIndex = -1;
            SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
            Invalidate();
        }

        /// <summary>获取首个选中行文本（供 Ctrl+C 使用）</summary>
        public string? GetSelectedLineText()
        {
            foreach (var idx in _selected)
                return GetLineText?.Invoke(idx) ?? "";
            return null;
        }

        /// <summary>获取所有选中行文本（换行连接）</summary>
        public string GetSelectedText()
        {
            if (_selected.Count == 0) return "";
            var sb = new StringBuilder();
            // 排序输出
            var list = new List<long>(_selected);
            list.Sort();
            foreach (var idx in list)
            {
                string text = GetLineText?.Invoke(idx) ?? "";
                sb.AppendLine(text);
            }
            return sb.ToString().TrimEnd('\r', '\n');
        }

        /// <summary>滚动使指定行可见</summary>
        public void EnsureVisible(long index)
        {
            if (index < 0 || index >= _totalLines) return;
            int visible = FullyVisibleRowCount;
            int idx = (int)Math.Min(index, int.MaxValue);
            if (idx < _vPos)
                _vPos = idx;
            else if (idx >= _vPos + visible)
                _vPos = Math.Max(0, idx - visible + 1);
            ClampVPos();
            UpdateVScroll();
            Invalidate();
        }

        /// <summary>获取当前顶部行索引（兼容旧代码 TopItem.Index）</summary>
        public int TopIndex => _vPos;

        // ==================== 绘制 ====================
        // 完全可见的行数（向下取整，仅用于滚动判定/翻页/滚动条配置）。
        // 注意：OnPaint 的绘制范围用此值即可，无需 +1；底部多余像素由客户区自然裁剪。
        private int FullyVisibleRowCount => Math.Max(1, ClientRectangle.Height / RowHeight);

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(BackColor);

            if (_totalLines <= 0) return;

            int rowH = RowHeight;
            int clientW = ClientRectangle.Width;
            int clientH = ClientRectangle.Height;
            // 背景宽度：取内容宽与客户区宽的较大者（选中行蓝底铺满）
            int bgWidth = Math.Max(_maxLineWidth, clientW);

            int visible = FullyVisibleRowCount;
            long start = _vPos;
            long end = Math.Min(_totalLines, start + visible);

            // 绘制可见行
            for (long i = start; i < end; i++)
            {
                int y = (int)(i - start) * rowH;
                if (y >= clientH) break;
                bool selected = _selected.Contains(i);

                // 背景
                using (var backBrush = new SolidBrush(selected ? SelectionBackColor : BackColor))
                {
                    g.FillRectangle(backBrush, 0, y, bgWidth, rowH);
                }

                // 文本
                string text = GetLineText?.Invoke(i) ?? "";
                if (!string.IsNullOrEmpty(text))
                {
                    // 截断保护：GDI DrawText 无法渲染超长字符串
                    if (text.Length > 8 * 1024)
                        text = text.Substring(0, 8 * 1024) + "...更多";

                    int drawX = LeftMargin - _hPos;
                    int drawY = y + (rowH - Font.Height) / 2;
                    TextRenderer.DrawText(g, text, Font,
                        new Point(drawX, drawY),
                        selected ? SelectionTextColor : ForeColor,
                        TextFormatFlags.Left | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);
                }
            }

            // 绘制选中行聚焦边框（可选，简单实现）
            // 省略：FocusRectangle
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            ClampVPos();
            UpdateVScroll();
            UpdateHScroll();
            Invalidate();
        }

        // ==================== 滚动条 ====================
        private void ClampVPos()
        {
            int max = Math.Max(0, (int)Math.Min(_totalLines - FullyVisibleRowCount, int.MaxValue));
            if (_vPos > max) _vPos = max;
            if (_vPos < 0) _vPos = 0;
        }

        private void ClampHPos()
        {
            int max = Math.Max(0, _maxLineWidth - ClientRectangle.Width);
            if (_hPos > max) _hPos = max;
            if (_hPos < 0) _hPos = 0;
        }

        private void UpdateVScroll()
        {
            int visible = FullyVisibleRowCount;
            bool need = _totalLines > visible;
            _vScroll.Visible = need;
            if (!need) return;
            int max = (int)Math.Min(_totalLines, int.MaxValue);
            _vScroll.Minimum = 0;
            _vScroll.Maximum = Math.Max(0, max - 1); // Maximum 是上界，LargeChange 会减去
            _vScroll.LargeChange = Math.Max(1, visible);
            _vScroll.SmallChange = 1;
            if (_vPos > _vScroll.Maximum - _vScroll.LargeChange + 1)
                _vPos = Math.Max(0, _vScroll.Maximum - _vScroll.LargeChange + 1);
            _vScroll.Value = _vPos;
        }

        private void UpdateHScroll()
        {
            int clientW = ClientRectangle.Width;
            bool need = _maxLineWidth > clientW;
            _hScroll.Visible = need;
            if (!need) return;
            _hScroll.Minimum = 0;
            _hScroll.Maximum = Math.Max(0, _maxLineWidth - 1);
            _hScroll.LargeChange = Math.Max(1, clientW);
            _hScroll.SmallChange = 40;
            if (_hPos > _hScroll.Maximum - _hScroll.LargeChange + 1)
                _hPos = Math.Max(0, _hScroll.Maximum - _hScroll.LargeChange + 1);
            _hScroll.Value = _hPos;
        }

        private void OnVScroll(object? s, ScrollEventArgs e)
        {
            int newVal = e.NewValue;
            int max = Math.Max(0, _vScroll.Maximum - _vScroll.LargeChange + 1);
            if (newVal > max) newVal = max;
            if (newVal < 0) newVal = 0;
            if (newVal == _vPos) return;
            _vPos = newVal;
            Invalidate();
            VScrollChanged?.Invoke(this, EventArgs.Empty);
        }

        private void OnHScroll(object? s, ScrollEventArgs e)
        {
            int newVal = e.NewValue;
            int max = Math.Max(0, _hScroll.Maximum - _hScroll.LargeChange + 1);
            if (newVal > max) newVal = max;
            if (newVal < 0) newVal = 0;
            if (newVal == _hPos) return;
            _hPos = newVal;
            Invalidate();
        }

        // ==================== 鼠标 ====================
        protected override void OnMouseDown(MouseEventArgs e)
        {
            Focus();
            long idx = HitTest(e.Y);
            if (idx < 0 || idx >= _totalLines) return;

            if ((Control.ModifierKeys & Keys.Shift) != 0 && _anchorIndex >= 0)
            {
                // Shift 多选区间
                _selected.Clear();
                long lo = Math.Min(_anchorIndex, idx);
                long hi = Math.Max(_anchorIndex, idx);
                for (long i = lo; i <= hi; i++) _selected.Add(i);
            }
            else if ((Control.ModifierKeys & Keys.Control) != 0)
            {
                if (_selected.Contains(idx)) _selected.Remove(idx);
                else _selected.Add(idx);
                _anchorIndex = idx;
            }
            else
            {
                _selected.Clear();
                _selected.Add(idx);
                _anchorIndex = idx;
            }
            SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
            Invalidate();
            base.OnMouseDown(e);
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            long idx = HitTest(e.Y);
            if (idx >= 0 && idx < _totalLines)
                ItemDoubleClick?.Invoke(idx);
            base.OnMouseDoubleClick(e);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            int delta = e.Delta / 120;
            int step = (Control.ModifierKeys & Keys.Shift) != 0 ? FullyVisibleRowCount : 3;
            int newPos = _vPos - delta * step;
            if (newPos < 0) newPos = 0;
            int max = Math.Max(0, (int)Math.Min(_totalLines - FullyVisibleRowCount, int.MaxValue));
            if (newPos > max) newPos = max;
            if (newPos != _vPos)
            {
                _vPos = newPos;
                UpdateVScroll();
                Invalidate();
                VScrollChanged?.Invoke(this, EventArgs.Empty);
            }
            base.OnMouseWheel(e);
        }

        private long HitTest(int y)
        {
            int rowH = RowHeight;
            if (rowH <= 0) return -1;
            return _vPos + (y / rowH);
        }

        // ==================== 键盘 ====================
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                string text = GetSelectedText();
                if (!string.IsNullOrEmpty(text))
                {
                    try { Clipboard.SetText(text); }
                    catch { try { Clipboard.SetDataObject(text, true); } catch { } }
                }
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (_totalLines <= 0) { base.OnKeyDown(e); return; }

            long cur = _anchorIndex >= 0 && _anchorIndex < _totalLines ? _anchorIndex : 0;
            long newIdx = cur;
            switch (e.KeyCode)
            {
                case Keys.Up: newIdx = cur - 1; break;
                case Keys.Down: newIdx = cur + 1; break;
                case Keys.PageUp: newIdx = cur - FullyVisibleRowCount; break;
                case Keys.PageDown: newIdx = cur + FullyVisibleRowCount; break;
                case Keys.Home: newIdx = 0; break;
                case Keys.End: newIdx = _totalLines - 1; break;
                case Keys.Left:
                    {
                        int step = (e.Modifiers & Keys.Shift) != 0 ? ClientRectangle.Width : 40;
                        int newPos = _hPos - step;
                        if (newPos < 0) newPos = 0;
                        if (newPos != _hPos) { _hPos = newPos; UpdateHScroll(); Invalidate(); }
                        e.Handled = true;
                        return;
                    }
                case Keys.Right:
                    {
                        int step = (e.Modifiers & Keys.Shift) != 0 ? ClientRectangle.Width : 40;
                        int max = Math.Max(0, _maxLineWidth - ClientRectangle.Width);
                        int newPos = _hPos + step;
                        if (newPos > max) newPos = max;
                        if (newPos != _hPos) { _hPos = newPos; UpdateHScroll(); Invalidate(); }
                        e.Handled = true;
                        return;
                    }
                default:
                    base.OnKeyDown(e);
                    return;
            }

            if (newIdx < 0) newIdx = 0;
            if (newIdx >= _totalLines) newIdx = _totalLines - 1;
            if (newIdx == cur) { e.Handled = true; return; }

            if ((e.Modifiers & Keys.Shift) != 0 && _anchorIndex >= 0)
            {
                // Shift+方向键扩展选择
                _selected.Clear();
                long lo = Math.Min(_anchorIndex, newIdx);
                long hi = Math.Max(_anchorIndex, newIdx);
                for (long i = lo; i <= hi; i++) _selected.Add(i);
            }
            else
            {
                _selected.Clear();
                _selected.Add(newIdx);
                _anchorIndex = newIdx;
            }
            EnsureVisible(newIdx);
            SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
        }

        protected override bool IsInputKey(Keys keyData)
        {
            switch (keyData & Keys.KeyCode)
            {
                case Keys.Up:
                case Keys.Down:
                case Keys.Left:
                case Keys.Right:
                case Keys.PageUp:
                case Keys.PageDown:
                case Keys.Home:
                case Keys.End:
                    return true;
            }
            return base.IsInputKey(keyData);
        }
    }
}
