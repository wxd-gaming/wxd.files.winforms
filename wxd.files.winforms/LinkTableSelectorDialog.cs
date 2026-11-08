using System.Data;

namespace wxd.files.winforms
{
    /// <summary>关联表选择对话框 — 展示左侧已加载的文件列表，带搜索筛选</summary>
    public class LinkTableSelectorDialog : Form
    {
        private ListView _lstFiles = new();
        private ThemeField _fldFilter = new();
        private ThemeButton _btnOK = new();
        private ThemeButton _btnClear = new();
        private Panel _topBar = new();
        private Panel _topDivider = new();
        private Panel _bar = new();
        private List<(string display, string fullPath)> _fileEntries = new();
        private List<(string display, string fullPath)> _filteredEntries = new();

        public string SelectedFilePath { get; private set; } = "";

        public LinkTableSelectorDialog(List<(string display, string fullPath)> fileEntries)
        {
            _fileEntries = fileEntries;
            InitializeComponent();
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
            PopulateList("");
        }

        private void PopulateList(string filter)
        {
            _lstFiles.BeginUpdate();
            _lstFiles.Items.Clear();
            _filteredEntries.Clear();

            var query = string.IsNullOrEmpty(filter)
                ? _fileEntries
                : _fileEntries.Where(f => f.display.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var entry in query.OrderBy(f => f.display))
            {
                _filteredEntries.Add(entry);
                var item = new ListViewItem(entry.display)
                {
                    Tag = entry.fullPath,
                    ToolTipText = entry.fullPath
                };
                _lstFiles.Items.Add(item);
            }

            _lstFiles.EndUpdate();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "选择关联表";
            ClientSize = new Size(500, 400);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;

            _topBar = new Panel { Dock = DockStyle.Top, Height = ThemeMetrics.BarHeight };
            _topDivider = new Panel { Dock = DockStyle.Bottom, Height = 1 };
            _topBar.Controls.Add(_topDivider);

            _fldFilter = new ThemeField { Placeholder = "输入关键字筛选文件..." };
            _fldFilter.Edit.TextChanged += (s, e) => PopulateList(_fldFilter.Edit.Text.Trim());
            _topBar.Controls.Add(_fldFilter);
            LayoutTopBar();
            _topBar.Resize += (_, _) => LayoutTopBar();

            _lstFiles = new ListView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                GridLines = false,
                HeaderStyle = ColumnHeaderStyle.None,
                MultiSelect = false
            };
            _lstFiles.Columns.Add("文件", 460);
            _lstFiles.DoubleClick += (s, e) => AcceptSelection();

            _bar = DialogButtonBar.Create();
            _btnClear = DialogButtonBar.Make("清除", (s, e) =>
            {
                SelectedFilePath = "";
                DialogResult = DialogResult.Cancel;
            });
            _btnOK = DialogButtonBar.Make("确定", (s, e) => AcceptSelection(), primary: true);
            _bar.Controls.AddRange(new Control[] { _btnClear, _btnOK });
            LayoutBar();
            _bar.Resize += (_, _) => LayoutBar();

            // 停靠顺序：后加入的靠外，所以列表先加、页眉页脚后加
            Controls.Add(_lstFiles);
            Controls.Add(_topBar);
            Controls.Add(_bar);
            AcceptButton = _btnOK;
            CancelButton = _btnClear;

            ResumeLayout(false);
        }

        private void AcceptSelection()
        {
            if (_lstFiles.SelectedItems.Count == 0) return;
            SelectedFilePath = _lstFiles.SelectedItems[0].Tag as string ?? "";
            DialogResult = DialogResult.OK;
        }

        private void LayoutTopBar()
            => _fldFilter.SetBounds(ThemeMetrics.Margin, ThemeMetrics.RowPad,
                _topBar.Width - ThemeMetrics.Margin * 2, ThemeMetrics.RowH);

        private void LayoutBar()
            => DialogButtonBar.LayoutRight(_bar, _btnClear, _btnOK);

        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _topBar.BackColor = p.Chrome;
            _bar.BackColor = p.Chrome;
            _topDivider.BackColor = p.Divider;

            _lstFiles.BackColor = p.Surface;
            _lstFiles.ForeColor = p.Text;
            _lstFiles.Font = Theme.List;

            _fldFilter.ApplyTheme();
            _btnClear.BackColor = p.Chrome;
            _btnOK.BackColor = p.Chrome;

            LayoutTopBar();
            LayoutBar();
            Invalidate(true);
        }
    }
}
