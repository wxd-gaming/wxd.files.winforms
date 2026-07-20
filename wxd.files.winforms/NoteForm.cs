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
        // ==================== 配色（浅色/深色） ====================
        private readonly bool _isDark;
        private static readonly Color _lightBg = Color.FromArgb(245, 250, 255);
        private static readonly Color _lightPanel = Color.White;
        private static readonly Color _lightBorder = Color.FromArgb(200, 215, 230);
        private static readonly Color _lightListItemBg = Color.White;
        private static readonly Color _lightListItemHover = Color.FromArgb(235, 243, 252);
        private static readonly Color _lightListItemSel = Color.FromArgb(210, 230, 250);
        private static readonly Color _lightTitleFg = Color.FromArgb(20, 60, 100);
        private static readonly Color _lightSubFg = Color.FromArgb(120, 140, 160);
        private static readonly Color _lightTextFg = Color.FromArgb(30, 50, 80);

        private static readonly Color _darkBg = Color.FromArgb(30, 30, 30);
        private static readonly Color _darkPanel = Color.FromArgb(45, 45, 48);
        private static readonly Color _darkBorder = Color.FromArgb(62, 62, 66);
        private static readonly Color _darkListItemBg = Color.FromArgb(45, 45, 48);
        private static readonly Color _darkListItemHover = Color.FromArgb(62, 62, 66);
        private static readonly Color _darkListItemSel = Color.FromArgb(81, 113, 170);
        private static readonly Color _darkTitleFg = Color.FromArgb(241, 241, 241);
        private static readonly Color _darkSubFg = Color.FromArgb(160, 160, 160);
        private static readonly Color _darkTextFg = Color.FromArgb(220, 220, 220);

        // ==================== 控件 ====================
        private SplitContainer _splitter = null!;
        private Panel _leftPanel = null!;
        private Panel _toolBarPanel = null!;
        private Button _btnNew = null!;
        private Button _btnDelete = null!;
        private TextBox _txtSearch = null!;
        private Button _btnSearch = null!;
        private Panel _listPanel = null!;
        private Panel _rightPanel = null!;
        private TextBox _txtTitle = null!;
        private RichTextBox _rtbContent = null!;
        private StatusStrip _statusBar = null!;
        private ToolStripStatusLabel _lblStatus = null!;
        private ToolStripStatusLabel _lblCount = null!;

        // ==================== 状态 ====================
        private List<NoteItem> _allNotes = new();           // 全部笔记
        private List<NoteItem> _filtered = new();           // 过滤后显示
        private string? _currentId;                          // 当前选中的笔记 Id
        private readonly System.Windows.Forms.Timer _debounceTimer;
        private bool _dirty;                                 // 是否有待保存的修改

        public NoteForm(bool isDark)
        {
            _isDark = isDark;
            _debounceTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            _debounceTimer.Tick += OnDebounceTick;

            BuildUi();
            ApplyTheme();
            LoadNotes();
        }

        // ==================== UI 构建 ====================
        private void BuildUi()
        {
            Text = "📝 笔记";
            ClientSize = new Size(900, 580);
            MinimumSize = new Size(640, 420);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9f);
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
                Height = 36,
                Padding = new Padding(6, 4, 6, 4)
            };

            _btnNew = new Button
            {
                Text = "+ 新建",
                Size = new Size(80, 26),
                Location = new Point(6, 5),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand
            };
            _btnNew.FlatAppearance.BorderSize = 1;
            _btnNew.Click += OnNewClick;
            _toolBarPanel.Controls.Add(_btnNew);

            _btnDelete = new Button
            {
                Text = "🗑 删除",
                Size = new Size(80, 26),
                Location = new Point(6 + 80 + 6, 5),   // 紧贴新建按钮右侧
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 9f),
                Cursor = Cursors.Hand
            };
            _btnDelete.FlatAppearance.BorderSize = 1;
            _btnDelete.Click += OnDeleteClick;
            _toolBarPanel.Controls.Add(_btnDelete);

            // 搜索行
            var searchPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 34,
                Padding = new Padding(6, 2, 6, 4)
            };

            _txtSearch = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f),
                PlaceholderText = "搜索标题或内容..."
            };
            _txtSearch.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { ApplyFilter(); e.SuppressKeyPress = true; }
            };

            _btnSearch = new Button
            {
                Text = "🔍",
                Size = new Size(40, 24),
                Dock = DockStyle.Right,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f),
                Cursor = Cursors.Hand
            };
            _btnSearch.FlatAppearance.BorderSize = 1;
            _btnSearch.Click += (s, e) => ApplyFilter();

            // Dock 顺序: Fill 先 Add(z 低,最后布局占剩余), Right 后 Add(z 高,先占右侧)
            searchPanel.Controls.Add(_txtSearch);
            searchPanel.Controls.Add(_btnSearch);

            // 列表（用可滚动 Panel 承载 NoteListItemControl，子控件 Dock=Top 自动堆叠并占满宽度）
            _listPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BorderStyle = BorderStyle.None,
                Font = new Font("Microsoft YaHei UI", 9f)
            };

            // Dock 布局顺序：WinForms 按 z-order 逆序 Dock（索引大的先布局,占外侧位置）。
            // 要让 Top 控件出现在顶部且不被 Fill 覆盖,Fill 必须最先 Add(z 最低),
            // Top 控件按"越靠外的越后 Add"的顺序加入。
            // 工具栏(最外顶) → 最后 Add, 搜索框(次外顶) → 第二 Add, 列表(Fill) → 第一 Add
            _leftPanel.Controls.Add(_listPanel);
            _leftPanel.Controls.Add(searchPanel);
            _leftPanel.Controls.Add(_toolBarPanel);
        }

        private void BuildRightPanel(Panel parent)
        {
            _rightPanel = new Panel { Dock = DockStyle.Fill };
            parent.Controls.Add(_rightPanel);

            // 标题
            var titlePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 34,
                Padding = new Padding(8, 4, 8, 4)
            };

            var lblTitleHint = new Label
            {
                Text = "标题：",
                AutoSize = true,
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei UI", 9f)
            };

            _txtTitle = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 10f)
            };
            _txtTitle.TextChanged += OnTitleChanged;
            // 按 Enter 跳转到内容编辑区
            _txtTitle.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { _rtbContent.Focus(); e.SuppressKeyPress = true; }
            };

            // Dock 顺序: Fill 先 Add(z 低,最后布局占剩余), Left 后 Add(z 高,先布局占左侧)
            // 这样 Label 在视觉上覆盖 TextBox 左侧, TextBox Fill 剩余空间
            titlePanel.Controls.Add(_txtTitle);
            titlePanel.Controls.Add(lblTitleHint);

            // 状态栏（Bottom）
            _statusBar = new StatusStrip { BackColor = Color.Transparent };
            _lblCount = new ToolStripStatusLabel { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
            _lblStatus = new ToolStripStatusLabel { TextAlign = ContentAlignment.MiddleRight };
            _statusBar.Items.Add(_lblCount);
            _statusBar.Items.Add(_lblStatus);

            // 内容编辑区（Fill）
            _rtbContent = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new Font("Microsoft YaHei UI", 10.5f),
                WordWrap = true,
                DetectUrls = true,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            _rtbContent.TextChanged += OnContentChanged;

            // Dock 布局顺序：先添加 Fill，再添加 Top/Bottom
            // （WinForms 按 z-order 逆序分配空间，先 Add 的在顶层，
            //  Fill 必须最先添加，使其在最底层，Top/Bottom 再从边缘挤占）
            _rightPanel.Controls.Add(_rtbContent);
            _rightPanel.Controls.Add(_statusBar);
            _rightPanel.Controls.Add(titlePanel);
        }

        // ==================== 主题 ====================
        private void ApplyTheme()
        {
            var bg = _isDark ? _darkBg : _lightBg;
            var panel = _isDark ? _darkPanel : _lightPanel;
            var border = _isDark ? _darkBorder : _lightBorder;
            var titleFg = _isDark ? _darkTitleFg : _lightTitleFg;
            var subFg = _isDark ? _darkSubFg : _lightSubFg;
            var textFg = _isDark ? _darkTextFg : _lightTextFg;

            BackColor = bg;
            _leftPanel.BackColor = panel;
            _rightPanel.BackColor = bg;
            _toolBarPanel.BackColor = panel;
            _splitter.BackColor = border;

            // 按钮（新建=琥珀色，删除=灰红）
            StyleButton(_btnNew, Color.FromArgb(255, 152, 0), Color.FromArgb(245, 124, 0));
            StyleButton(_btnDelete, Color.FromArgb(229, 115, 115), Color.FromArgb(198, 78, 78));
            StyleButton(_btnSearch, _isDark ? Color.FromArgb(70, 70, 74) : Color.FromArgb(227, 235, 245),
                         _isDark ? Color.FromArgb(90, 90, 94) : Color.FromArgb(180, 200, 220));
            _btnSearch.ForeColor = _isDark ? Color.White : Color.FromArgb(40, 60, 90);

            _txtSearch.BackColor = _isDark ? Color.FromArgb(37, 37, 38) : Color.White;
            _txtSearch.ForeColor = textFg;
            _txtSearch.BorderStyle = BorderStyle.FixedSingle;

            _listPanel.BackColor = panel;
            _listPanel.ForeColor = titleFg;

            _txtTitle.BackColor = _isDark ? Color.FromArgb(37, 37, 38) : Color.White;
            _txtTitle.ForeColor = titleFg;

            _rtbContent.BackColor = _isDark ? Color.FromArgb(30, 30, 30) : Color.White;
            _rtbContent.ForeColor = textFg;

            _statusBar.BackColor = panel;
            _lblCount.ForeColor = subFg;
            _lblStatus.ForeColor = subFg;
        }

        private static void StyleButton(Button btn, Color bg, Color border)
        {
            btn.BackColor = bg;
            btn.ForeColor = Color.White;
            btn.FlatAppearance.BorderColor = border;
            btn.FlatAppearance.MouseOverBackColor = ControlPaint.Light(bg, 0.1f);
        }

        // ==================== 数据加载 ====================
        private void LoadNotes()
        {
            _allNotes = NoteStore.GetAll();
            ApplyFilter();
            UpdateCount();
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
                var ctrl = new NoteListItemControl(n, _isDark);
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

            // 如无选中，自动选中第一条
            if (_currentId == null && _filtered.Count > 0)
            {
                _currentId = _filtered[0].Id;
                LoadNoteToEditor(_filtered[0]);
                // 标记第一个控件为选中
                if (_listPanel.Controls.Count > 0 && _listPanel.Controls[0] is NoteListItemControl firstCtrl)
                    firstCtrl.SetSelected(true);
            }
        }

        /// <summary>面板尺寸变化时不需要额外处理，Dock=Top 的子控件会自动占满宽度</summary>
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
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
                }
                else
                {
                    _rtbContent.Clear();
                }
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
            // 选中新建的笔记（位于列表第一项）
            if (_listPanel.Controls.Count > 0 && _listPanel.Controls[0] is NoteListItemControl newCtrl)
            {
                newCtrl.SetSelected(true);
                OnItemSelected(newCtrl);
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
            // 选中第一条
            if (_listPanel.Controls.Count > 0 && _listPanel.Controls[0] is NoteListItemControl firstCtrl)
            {
                firstCtrl.SetSelected(true);
                OnItemSelected(firstCtrl);
            }
        }

        // ==================== 事件：内容变更（防抖自动保存） ====================
        private bool _suppressContentEvents;

        private void OnTitleChanged(object? sender, EventArgs e)
        {
            if (_suppressContentEvents || _currentId == null) return;
            _dirty = true;
            _debounceTimer.Stop();
            _debounceTimer.Start();
            _lblStatus.Text = "编辑中...";
            // 同步列表项控件显示的标题
            var ctrl = FindControlByNoteId(_currentId);
            if (ctrl != null) ctrl.UpdateTitle();
        }

        private void OnContentChanged(object? sender, EventArgs e)
        {
            if (_suppressContentEvents || _currentId == null) return;
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
            var title = _txtTitle.Text;
            var rtf = _rtbContent.Rtf;
            var plain = _rtbContent.Text;
            NoteStore.Update(_currentId, title, rtf, plain);

            // 同步内存中的笔记
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

            // 刷新列表显示（更新时间与排序），并恢复当前选中
            ApplyFilter();
            var ctrl = FindControlByNoteId(_currentId);
            if (ctrl != null)
            {
                ctrl.SetSelected(true);
                ctrl.ScrollIntoViewSafe();
            }
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
