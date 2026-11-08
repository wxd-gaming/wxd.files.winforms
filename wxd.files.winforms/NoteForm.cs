using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 笔记主窗体（非模态、单例）
    /// 左侧：笔记列表（标题 + 更新时间）+ 新建/删除按钮 + 搜索框
    /// 右侧：标题输入 + RichTextBox 富文本（支持 Ctrl+V 粘贴QQ截图）
    /// </summary>
    public class NoteForm : Form
    {
        // ==================== 控件 ====================
        private SplitContainer _splitter = null!;
        private Panel _leftPanel = null!;
        private Panel _toolBarPanel = null!;
        private Panel _searchRow = null!;
        private ThemeButton _btnNew = null!;
        private ThemeButton _btnDelete = null!;
        private ThemeField _fldSearch = null!;
        private TextBox _txtSearch = null!;
        private ThemeButton _btnSearch = null!;
        private Panel _listPanel = null!;
        private Panel _rightPanel = null!;
        private Panel _cardPanel = null!;
        private Panel _titlePanel = null!;
        private Label _lblTitleHint = null!;
        private TextBox _txtTitle = null!;
        private RichTextBox _rtbContent = null!;
        private Panel _statusBar = null!;
        private Label _lblStatus = null!;
        private Label _lblCount = null!;

        // ==================== 状态 ====================
        private List<NoteItem> _allNotes = new();           // 全部笔记
        private List<NoteItem> _filtered = new();           // 过滤后显示
        private string? _currentId;                          // 当前选中的笔记 Id
        private readonly System.Windows.Forms.Timer _debounceTimer;
        private bool _dirty;                                 // 是否有待保存的修改
        // 编辑内容缓存：在 TextChanged 事件中写入，FlushSave 只读取缓存，
        // 避免防抖保存时访问 RichTextBox 的 .Rtf/.Text 打断 IME 组合输入
        private string _pendingTitle = "";
        private string _pendingRtf = "";
        private string _pendingPlain = "";

        public NoteForm()
        {
            _debounceTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            _debounceTimer.Tick += OnDebounceTick;

            BuildUi();
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
            LoadNotes();
        }

        // ==================== UI 构建 ====================
        private void BuildUi()
        {
            Text = "📝 笔记";
            ClientSize = new Size(900, 580);
            MinimumSize = new Size(640, 420);
            StartPosition = FormStartPosition.CenterScreen;
            Font = Theme.Body;
            ShowIcon = false;
            KeyPreview = true;

            _splitter = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 4,
                FixedPanel = FixedPanel.Panel1,
                // 初始值,OnLoad 中根据 ClientSize 重新设置以确保生效
                SplitterDistance = 300
            };
            Controls.Add(_splitter);

            BuildLeftPanel(_splitter.Panel1);
            BuildRightPanel(_splitter.Panel2);

            // 快捷键：Ctrl+S 立即保存
            KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.S)
                {
                    FlushSave();
                    _lblStatus.Text = "已保存 ✓";
                    e.SuppressKeyPress = true;
                }
            };
        }

        private void BuildLeftPanel(Panel parent)
        {
            _leftPanel = new Panel { Dock = DockStyle.Fill };
            parent.Controls.Add(_leftPanel);

            // 工具栏：新建 + 删除
            _toolBarPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = ThemeMetrics.BarHeight
            };

            _btnNew = DialogButtonBar.Make("新建");
            _btnNew.Click += OnNewClick;
            _btnDelete = DialogButtonBar.Make("删除");
            _btnDelete.Click += OnDeleteClick;
            _toolBarPanel.Controls.AddRange(new Control[] { _btnNew, _btnDelete });
            _toolBarPanel.Resize += (_, _) => LayoutToolBar();
            LayoutToolBar();

            // 搜索行：输入框占满剩余宽度，搜索按钮贴右
            _searchRow = new Panel
            {
                Dock = DockStyle.Top,
                Height = ThemeMetrics.BarHeight
            };
            _fldSearch = new ThemeField { Placeholder = "搜索标题或内容..." };
            _txtSearch = _fldSearch.Edit;
            _txtSearch.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { ApplyFilter(); e.SuppressKeyPress = true; }
            };
            _btnSearch = DialogButtonBar.Make("搜索");
            _btnSearch.Click += (s, e) => ApplyFilter();
            _searchRow.Controls.AddRange(new Control[] { _fldSearch, _btnSearch });
            _searchRow.Resize += (_, _) => LayoutSearchRow(_searchRow);
            LayoutSearchRow(_searchRow);

            // 列表（用可滚动 Panel 承载 NoteListItemControl，子控件 Dock=Top 自动堆叠并占满宽度）
            _listPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BorderStyle = BorderStyle.None,
                Font = Theme.Body
            };

            // Dock 布局顺序：WinForms 按 z-order 逆序 Dock（索引大的先布局,占外侧位置）。
            // 要让 Top 控件出现在顶部且不被 Fill 覆盖,Fill 必须最先 Add(z 最低),
            // Top 控件按"越靠外的越后 Add"的顺序加入。
            // 工具栏(最外顶) → 最后 Add, 搜索框(次外顶) → 第二 Add, 列表(Fill) → 第一 Add
            _leftPanel.Controls.Add(_listPanel);
            _leftPanel.Controls.Add(_searchRow);
            _leftPanel.Controls.Add(_toolBarPanel);
        }

        private void LayoutToolBar()
        {
            int x = ThemeMetrics.Margin, y = ThemeMetrics.RowPad;
            int rowH = ThemeMetrics.RowH;
            _btnNew.SetBounds(x, y, ThemeButton.FitWidth(_btnNew.Text), rowH);
            x += _btnNew.Width + ThemeMetrics.Gap;
            _btnDelete.SetBounds(x, y, ThemeButton.FitWidth(_btnDelete.Text), rowH);
        }

        private void LayoutSearchRow(Panel row)
        {
            int m = ThemeMetrics.Margin, y = ThemeMetrics.RowPad;
            int rowH = ThemeMetrics.RowH;
            int btnW = ThemeButton.FitWidth(_btnSearch.Text);
            _btnSearch.SetBounds(row.Width - m - btnW, y, btnW, rowH);
            _fldSearch.SetBounds(m, y, Math.Max(60, _btnSearch.Left - m - ThemeMetrics.Gap), rowH);
        }

        private void BuildRightPanel(Panel parent)
        {
            _rightPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(ThemeMetrics.Group, ThemeMetrics.Group,
                                     ThemeMetrics.Group, ThemeMetrics.RowPad)   // 窗口底色留白，衬托卡片
            };
            parent.Controls.Add(_rightPanel);

            // 笔记正文画成圆角卡片，和文件管理器那边的窗格卡片同一套画法：
            // Surface 填充 + 1px 描边，聚焦时描边转 Accent
            _cardPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 6, 10, 6)
            };
            _cardPanel.Paint += (s, e) =>
            {
                var p = Theme.P;
                float stroke = _rtbContent.Focused || _txtTitle.Focused ? 2f : 1f;
                var rect = new RectangleF(stroke / 2f, stroke / 2f,
                    _cardPanel.Width - stroke, _cardPanel.Height - stroke);
                using var path = ThemeButton.RoundRect(rect, p.Radius);
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var fill = new SolidBrush(p.Surface)) e.Graphics.FillPath(fill, path);
                using var pen = new Pen(_rtbContent.Focused || _txtTitle.Focused ? p.Accent : p.Border, stroke);
                e.Graphics.DrawPath(pen, path);
            };
            _rightPanel.Controls.Add(_cardPanel);

            // 标题区：Label + 无边框输入框 + 底部 1px 细分隔线，边界由卡片提供
            _titlePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                Padding = new Padding(0, 4, 0, 6)
            };
            _titlePanel.Paint += (s, e) =>
            {
                var r = _titlePanel.ClientRectangle;
                using var pen = new Pen(Theme.P.Divider, 1);
                e.Graphics.DrawLine(pen, 0, r.Bottom - 1, r.Width, r.Bottom - 1);
            };

            _lblTitleHint = new Label
            {
                Text = "标题：",
                AutoSize = true,
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _txtTitle = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = Theme.Body,
                PlaceholderText = "输入标题..."
            };
            _txtTitle.TextChanged += OnTitleChanged;
            // 按 Enter 跳转到内容编辑区
            _txtTitle.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { _rtbContent.Focus(); e.SuppressKeyPress = true; }
            };

            // Dock 顺序: Fill 先 Add(z 低,最后布局占剩余), Left 后 Add(z 高,先布局占左侧)
            _titlePanel.Controls.Add(_txtTitle);
            _titlePanel.Controls.Add(_lblTitleHint);

            // 状态栏（Bottom）：Panel + Label，跟其他窗口一致。
            // 原来用 StatusStrip，但 ToolStripStatusLabel 不认 ForeColor（系统渲染器走自己的取色），
            // 深色下会画成黑底黑字，只剩 ClearType 边缘那点蓝色像素。
            _statusBar = new Panel { Dock = DockStyle.Bottom, Height = Theme.P.RowHeight };
            _statusBar.Paint += (s, e) =>
            {
                var r = _statusBar.ClientRectangle;
                using var pen = new Pen(Theme.P.Divider, 1);
                e.Graphics.DrawLine(pen, 0, 0, r.Width, 0);
            };
            _lblCount = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(2, 0, 0, 0)
            };
            _lblStatus = new Label
            {
                Dock = DockStyle.Right,
                AutoSize = false,
                Width = 130,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 6, 0)
            };
            // Dock 顺序：Fill 先 Add 占剩余，Right 后 Add 先挤右侧
            _statusBar.Controls.Add(_lblCount);
            _statusBar.Controls.Add(_lblStatus);

            // 内容编辑区（Fill）：无边框，边界由卡片提供，消除"框中框"
            _rtbContent = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = Theme.Body,
                WordWrap = true,
                DetectUrls = true,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            _rtbContent.TextChanged += OnContentChanged;
            // 聚焦态画在卡片描边上，进出焦点要重画一次卡片
            foreach (Control c in new Control[] { _rtbContent, _txtTitle })
            {
                c.Enter += (_, _) => _cardPanel.Invalidate();
                c.Leave += (_, _) => _cardPanel.Invalidate();
            }

            // Dock 布局顺序：Fill 先 Add(z 低)，Top/Bottom 后 Add(z 高,先从边缘挤占)
            _cardPanel.Controls.Add(_rtbContent);
            _cardPanel.Controls.Add(_titlePanel);
            _rightPanel.Controls.Add(_statusBar);
        }

        // ==================== 主题 ====================
        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;

            BackColor = p.Window;
            _leftPanel.BackColor = p.Window;
            _rightPanel.BackColor = p.Window;
            _toolBarPanel.BackColor = p.Chrome;
            _searchRow.BackColor = p.Chrome;
            _splitter.BackColor = p.Window;   // 分隔条与底色一致，消除突兀灰条

            foreach (var b in new[] { _btnNew, _btnDelete, _btnSearch })
                b.BackColor = p.Chrome;
            _fldSearch.ApplyTheme();

            // 列表项自己会按 Surface/Hover/Selection 上色，这里只兜住空隙
            _listPanel.BackColor = p.Surface;
            _listPanel.ForeColor = p.Text;

            // 卡片：本体 Surface，圆角外圈露出的部分跟窗口底色一致
            _cardPanel.BackColor = p.Window;
            _titlePanel.BackColor = p.Surface;
            _lblTitleHint.ForeColor = p.Muted;
            _txtTitle.BackColor = p.Surface;
            _txtTitle.ForeColor = p.Text;
            _rtbContent.BackColor = p.Surface;
            _rtbContent.ForeColor = p.Text;
            if (_rtbContent.IsHandleCreated)
            {
                // 换肤刷字色不是用户编辑：不压住事件就会把笔记标脏，防抖一过自动存盘并挪动 UpdatedAt
                _suppressContentEvents = true;
                try { RecolorContentText(); }
                finally { _suppressContentEvents = false; }
            }

            _statusBar.BackColor = p.Chrome;
            _lblCount.ForeColor = p.Muted;
            _lblStatus.ForeColor = p.Muted;

            LayoutToolBar();
            LayoutSearchRow(_searchRow);
            _cardPanel.Invalidate();
            Invalidate(true);
        }

        // ==================== 数据加载 ====================
        private void LoadNotes()
        {
            _allNotes = NoteStore.GetAll();
            ApplyFilter();
            UpdateCount();

            // 列表非空则默认选中并显示第一条
            if (_filtered.Count > 0)
            {
                _currentId = _filtered[0].Id;
                LoadNoteToEditor(_filtered[0]);
                FindControlByNoteId(_currentId)?.SetSelected(true);
            }
        }

        private void ApplyFilter()
        {
            // 切换前保存当前
            FlushSave();

            var kw = (_txtSearch.Text ?? "").Trim();
            if (string.IsNullOrEmpty(kw))
            {
                _filtered = new List<NoteItem>(_allNotes);
            }
            else
            {
                _filtered = _allNotes.Where(n =>
                    (n.Title ?? "").IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (n.PlainText ?? "").IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0
                ).ToList();
            }

            // 重建列表控件（Panel + Dock=Top: 先 Add 的在底部，需逆序 Add 让第一条在顶部）
            _listPanel.SuspendLayout();
            _listPanel.Controls.Clear();
            NoteListItemControl? selectCtrl = null;
            for (var i = _filtered.Count - 1; i >= 0; i--)
            {
                var n = _filtered[i];
                var ctrl = new NoteListItemControl(n);
                ctrl.Selected += OnItemSelected;
                _listPanel.Controls.Add(ctrl);
                if (n.Id == _currentId) selectCtrl = ctrl;
            }
            _listPanel.ResumeLayout(true);

            // 保持当前选中
            if (selectCtrl != null)
            {
                selectCtrl.SetSelected(true);
                selectCtrl.ScrollIntoViewSafe();
            }
            UpdateCount();
        }

        /// <summary>列表项被点击选中时</summary>
        private void OnItemSelected(NoteListItemControl ctrl)
        {
            // 取消其他项的选中
            foreach (Control c in _listPanel.Controls)
            {
                if (c is NoteListItemControl ic && ic != ctrl) ic.SetSelected(false);
            }

            // 切换前保存上一个笔记
            FlushSave();

            _currentId = ctrl.Note.Id;
            LoadNoteToEditor(ctrl.Note);
        }

        private void UpdateCount()
        {
            _lblCount.Text = $"共 {_filtered.Count} / {_allNotes.Count} 条笔记";
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // 在窗体已布局、ClientSize 已确定后设置 SplitterDistance,确保左侧宽度生效
            // 左侧需要容纳工具栏(两个按钮 ~170) + 搜索行(搜索框 + 按钮 ~230),取 280 保证不挤
            var desired = 280;
            var max = _splitter.Width - _splitter.SplitterWidth - _splitter.Panel2MinSize;
            if (max < desired) desired = Math.Max(_splitter.Panel1MinSize, max);
            _splitter.SplitterDistance = desired;
        }

        /// <summary>面板尺寸变化时不需要额外处理，Dock=Top 的子控件会自动占满宽度</summary>
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
        }

        /// <summary>
        /// 正文按当前主题重刷取色。
        /// 笔记存的是 RTF，里面钉着写入当时主题的正文色（浅色下是 30,50,80 这种近黑蓝），
        /// 直接 Rtf= 赋值会把控件 ForeColor 顶掉，深色底上就变成看不见。
        /// 编辑器本身没有改字色的功能，所以读回来一律刷成 Theme.P.Text，光标位置原样还回去。
        /// </summary>
        private void RecolorContentText()
        {
            if (_rtbContent.TextLength == 0) return;
            int caret = _rtbContent.SelectionStart;
            int len = _rtbContent.SelectionLength;
            _rtbContent.SelectAll();
            _rtbContent.SelectionColor = Theme.P.Text;
            _rtbContent.Select(caret, len);
        }

        private void LoadNoteToEditor(NoteItem note)
        {
            _suppressContentEvents = true;
            try
            {
                _txtTitle.Text = note.Title;
                if (!string.IsNullOrEmpty(note.RtfContent))
                {
                    try { _rtbContent.Rtf = note.RtfContent; }
                    catch { _rtbContent.Text = note.PlainText; }
                    RecolorContentText();
                }
                else
                {
                    _rtbContent.Clear();
                }
                // 同步缓存，使 FlushSave 在未发生新的编辑时也有正确的值
                _pendingTitle = note.Title;
                _pendingRtf = _rtbContent.Rtf;
                _pendingPlain = _rtbContent.Text;
                _dirty = false;
                _debounceTimer.Stop();
                _lblStatus.Text = "";
            }
            finally
            {
                _suppressContentEvents = false;
            }
        }

        // ==================== 事件：新建 / 删除 ====================
        private void OnNewClick(object? sender, EventArgs e)
        {
            FlushSave();
            var note = NoteStore.Add();
            _allNotes.Insert(0, note);
            // 清空搜索以便显示新笔记
            _txtSearch.Text = "";
            ApplyFilter();
            // 直接选中新建的笔记（ApplyFilter 后它位于列表顶部）
            _currentId = note.Id;
            LoadNoteToEditor(note);
            var ctrl = FindControlByNoteId(_currentId);
            if (ctrl != null)
            {
                ctrl.SetSelected(true);
                ctrl.ScrollIntoViewSafe();
            }
            _txtTitle.Focus();
            _txtTitle.SelectAll();
        }

        private void OnDeleteClick(object? sender, EventArgs e)
        {
            if (_currentId == null) return;
            var note = _allNotes.FirstOrDefault(n => n.Id == _currentId);
            if (note == null) return;

            var title = string.IsNullOrEmpty(note.Title) ? "(无标题)" : note.Title;
            if (AlertForm.ShowConfirm(this, $"确定删除笔记「{title}」吗？此操作不可恢复。", "删除确认") != DialogResult.Yes)
                return;

            NoteStore.Delete(_currentId);
            _allNotes.RemoveAll(n => n.Id == _currentId);
            _currentId = null;
            _txtTitle.Clear();
            _rtbContent.Clear();
            ApplyFilter();
            // 选中第一条（_filtered[0] 是最新的第一条）
            if (_filtered.Count > 0)
            {
                var firstNote = _filtered[0];
                _currentId = firstNote.Id;
                LoadNoteToEditor(firstNote);
                var ctrl = FindControlByNoteId(_currentId);
                if (ctrl != null)
                {
                    ctrl.SetSelected(true);
                    ctrl.ScrollIntoViewSafe();
                }
            }
        }

        // ==================== 事件：内容变更（防抖自动保存） ====================
        private bool _suppressContentEvents;

        private void OnTitleChanged(object? sender, EventArgs e)
        {
            if (_suppressContentEvents || _currentId == null) return;
            _pendingTitle = _txtTitle.Text;   // 缓存标题，FlushSave 不再读取 TextBox
            _dirty = true;
            _debounceTimer.Stop();
            _debounceTimer.Start();
            _lblStatus.Text = "编辑中...";
            // 同步列表项控件显示的标题（实时，不依赖保存）
            var ctrl = FindControlByNoteId(_currentId);
            if (ctrl != null)
            {
                ctrl.Note.Title = _pendingTitle;
                ctrl.UpdateTitle();
            }
        }

        private void OnContentChanged(object? sender, EventArgs e)
        {
            if (_suppressContentEvents || _currentId == null) return;
            // 缓存 RTF 和纯文本，FlushSave 不再访问 RichTextBox（避免打断 IME 组合输入）
            _pendingRtf = _rtbContent.Rtf;
            _pendingPlain = _rtbContent.Text;
            _dirty = true;
            _debounceTimer.Stop();
            _debounceTimer.Start();
            _lblStatus.Text = "编辑中...";
        }

        private void OnDebounceTick(object? sender, EventArgs e)
        {
            _debounceTimer.Stop();
            FlushSave();
            _lblStatus.Text = "已自动保存 ✓";
            // 3 秒后清空提示
            var t = new System.Windows.Forms.Timer { Interval = 3000 };
            t.Tick += (s, _) => { _lblStatus.Text = ""; t.Stop(); t.Dispose(); };
            t.Start();
        }

        /// <summary>立即保存当前编辑内容到 Store</summary>
        private void FlushSave()
        {
            if (!_dirty || _currentId == null) return;
            // 只使用缓存值，不访问 RichTextBox/TextBox，避免打断 IME 组合输入
            var title = _pendingTitle;
            var rtf = _pendingRtf;
            var plain = _pendingPlain;
            NoteStore.Update(_currentId, title, rtf, plain);

            // 同步内存中的笔记（标题已在 OnTitleChanged 实时更新到列表控件，
            // 此处只更新数据层；时间/排序在切换笔记、新建、删除等非输入时机由 ApplyFilter 刷新）
            var n = _allNotes.FirstOrDefault(x => x.Id == _currentId);
            if (n != null)
            {
                n.Title = title;
                n.RtfContent = rtf;
                n.PlainText = plain;
                n.UpdatedAt = DateTime.Now;
                // 移到最前
                _allNotes.Remove(n);
                _allNotes.Insert(0, n);
            }
            _dirty = false;
            // 注意：不调用 ctrl.UpdateTitle()/UpdateTime()，避免防抖保存触发 UI 重绘打断输入
        }

        /// <summary>按笔记 Id 查找对应的列表项控件</summary>
        private NoteListItemControl? FindControlByNoteId(string? id)
        {
            if (id == null) return null;
            foreach (Control c in _listPanel.Controls)
            {
                if (c is NoteListItemControl ic && ic.Note.Id == id) return ic;
            }
            return null;
        }

        // ==================== 关闭时保存 ====================
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            _debounceTimer.Stop();
            FlushSave();
        }
    }

    internal static class StringExtensions
    {
        public static bool IsEmpty(this string? s) => string.IsNullOrEmpty(s);
    }

    /// <summary>列表项控件的扩展方法</summary>
    internal static class NoteListItemControlExtensions
    {
        /// <summary>让控件滚动到 FlowLayoutPanel 的可视区域（防止未创建句柄时报错）</summary>
        public static void ScrollIntoViewSafe(this NoteListItemControl ctrl)
        {
            try
            {
                if (!ctrl.IsHandleCreated || ctrl.Parent == null) return;
                var panel = ctrl.Parent as ScrollableControl;
                if (panel == null) return;
                panel.ScrollControlIntoView(ctrl);
            }
            catch { /* 忽略滚动异常 */ }
        }
    }
}
