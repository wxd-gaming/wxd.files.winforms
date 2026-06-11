using System.Data;

namespace wxd.files.winforms
{
    /// <summary>关联表选择对话框 — 展示左侧已加载的文件列表，带搜索筛选</summary>
    public class LinkTableSelectorDialog : Form
    {
        private ListView _lstFiles = new();
        private TextBox _txtFilter = new();
        private Button _btnOK = new();
        private Button _btnCancel = new();
        private List<(string display, string fullPath)> _fileEntries = new();
        private List<(string display, string fullPath)> _filteredEntries = new();

        public string SelectedFilePath { get; private set; } = "";

        public LinkTableSelectorDialog(List<(string display, string fullPath)> fileEntries)
        {
            _fileEntries = fileEntries;
            InitializeComponent();
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
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);
            MinimizeBox = false;
            MaximizeBox = false;

            // 筛选框
            _txtFilter = new TextBox
            {
                Location = new Point(10, 10),
                Size = new Size(480, 23),
                Font = new Font("Microsoft YaHei UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                PlaceholderText = "输入关键字筛选文件..."
            };
            _txtFilter.TextChanged += (s, e) => PopulateList(_txtFilter.Text.Trim());

            // 文件列表
            _lstFiles = new ListView
            {
                Location = new Point(10, 40),
                Size = new Size(480, 310),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 9f),
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                GridLines = false,
                HeaderStyle = ColumnHeaderStyle.None,
                MultiSelect = false
            };
            _lstFiles.Columns.Add("文件", 460);
            _lstFiles.DoubleClick += (s, e) =>
            {
                if (_lstFiles.SelectedItems.Count > 0)
                {
                    SelectedFilePath = _lstFiles.SelectedItems[0].Tag as string ?? "";
                    DialogResult = DialogResult.OK;
                    Close();
                }
            };

            // OK 按钮
            _btnOK = new Button
            {
                Text = "确定",
                Location = new Point(320, 360),
                Size = new Size(80, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
            };
            _btnOK.Click += (s, e) =>
            {
                if (_lstFiles.SelectedItems.Count > 0)
                {
                    SelectedFilePath = _lstFiles.SelectedItems[0].Tag as string ?? "";
                    DialogResult = DialogResult.OK;
                    Close();
                }
            };

            // 取消按钮
            _btnCancel = new Button
            {
                Text = "清除",
                Location = new Point(408, 360),
                Size = new Size(70, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel,
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237) }
            };

            Controls.AddRange(new Control[] { _txtFilter, _lstFiles, _btnOK, _btnCancel });
            AcceptButton = _btnOK;
            CancelButton = _btnCancel;

            ResumeLayout(false);
        }
    }
}
