using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>重复值详情窗体（可搜索、可复制、可跳转到行）</summary>
    public class DuplicateDetailForm : Form
    {
        private readonly List<TableBrowserControl.DuplicateGroup> _duplicates;
        private readonly FileTableData _fileData;
        private readonly TableBrowserControl _owner;
        private DataGridView _grid = new();
        private TextBox _txtSearch = new();
        private Label _lblSummary = new();
        private Button _btnCopy = new();
        private Button _btnClose = new();
        private DataTable _displayTable = new();

        private static readonly Color _colorBg = Color.FromArgb(245, 250, 255);
        private static readonly Color _colorPrimary = Color.FromArgb(66, 165, 245);

        public DuplicateDetailForm(
            List<TableBrowserControl.DuplicateGroup> duplicates,
            FileTableData fileData,
            TableBrowserControl owner)
        {
            _duplicates = duplicates;
            _fileData = fileData;
            _owner = owner;

            InitializeComponent();
            SetupGrid();
            PopulateGrid();
        }

        private void InitializeComponent()
        {
            Text = "重复值详情";
            ClientSize = new Size(750, 550);
            MinimumSize = new Size(500, 350);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = _colorBg;
            FormBorderStyle = FormBorderStyle.Sizable;
            ShowInTaskbar = false;

            var totalRedundant = _duplicates.Sum(d => d.Count - 1);

            // 摘要标签
            _lblSummary = new Label
            {
                Text = $"共 {_duplicates.Count} 个重复值，{totalRedundant} 处冗余  |  双击某行可跳转到对应位置",
                AutoSize = false,
                Height = 30,
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 10, 0),
                ForeColor = Color.FromArgb(180, 50, 30),
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                BackColor = Color.FromArgb(255, 240, 240),
            };

            // 搜索框
            var searchPanel = new Panel { Dock = DockStyle.Top, Height = 38, Padding = new Padding(10, 4, 10, 4) };
            var lblSearch = new Label
            {
                Text = "搜索:",
                AutoSize = true,
                Location = new Point(0, 8),
                ForeColor = Color.FromArgb(90, 130, 170),
            };
            _txtSearch = new TextBox
            {
                Width = 260,
                Anchor = AnchorStyles.Left,
                Font = new Font("Microsoft YaHei UI", 9f),
                Location = new Point(45, 6),
            };
            _txtSearch.TextChanged += (s, e) => ApplyFilter();
            searchPanel.Controls.Add(lblSearch);
            searchPanel.Controls.Add(_txtSearch);

            // DataGridView
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor = Color.FromArgb(200, 225, 245),
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(240, 246, 252),
                    ForeColor = Color.FromArgb(20, 60, 100),
                    Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    ForeColor = Color.FromArgb(20, 60, 100),
                    SelectionBackColor = Color.FromArgb(227, 242, 253),
                    SelectionForeColor = Color.FromArgb(20, 60, 100),
                },
            };
            _grid.CellDoubleClick += Grid_CellDoubleClick;

            // 底部按钮
            var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 46, Padding = new Padding(10, 6, 10, 6) };
            _btnCopy = new Button
            {
                Text = "复制到剪贴板",
                Size = new Size(120, 30),
                Anchor = AnchorStyles.Left,
                FlatStyle = FlatStyle.Flat,
                BackColor = _colorPrimary,
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9f),
            };
            _btnCopy.Click += BtnCopy_Click;
            _btnClose = new Button
            {
                Text = "关闭",
                Size = new Size(80, 30),
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(200, 210, 220),
                ForeColor = Color.FromArgb(60, 70, 80),
                Font = new Font("Microsoft YaHei UI", 9f),
            };
            _btnClose.Click += (s, e) => Close();
            bottomPanel.Controls.Add(_btnCopy);
            bottomPanel.Controls.Add(_btnClose);
            _btnClose.Location = new Point(bottomPanel.Width - 90, 6);

            Controls.Add(_grid);
            Controls.Add(bottomPanel);
            Controls.Add(searchPanel);
            Controls.Add(_lblSummary);
        }

        private void SetupGrid()
        {
            _displayTable = new DataTable();
            _displayTable.Columns.Add("Value", typeof(string));
            _displayTable.Columns.Add("Count", typeof(int));
            _displayTable.Columns.Add("Rows", typeof(string));

            _grid.DataSource = _displayTable;

            if (_grid.Columns.Count >= 3)
            {
                _grid.Columns[0].HeaderText = "重复值";
                _grid.Columns[0].FillWeight = 40;
                _grid.Columns[1].HeaderText = "次数";
                _grid.Columns[1].FillWeight = 10;
                _grid.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                _grid.Columns[2].HeaderText = "所在行号";
                _grid.Columns[2].FillWeight = 50;
            }
        }

        private void PopulateGrid()
        {
            _displayTable.Rows.Clear();
            foreach (var d in _duplicates.OrderByDescending(d => d.Count))
            {
                var rowNums = d.RowIndices
                    .Take(20)
                    .Select(idx => idx < _fileData.Rows.Count
                        ? (_fileData.Rows[idx].OriginalRowIndex + 1).ToString()
                        : (idx + 1).ToString());
                var rowsStr = string.Join(", ", rowNums);
                if (d.RowIndices.Count > 20)
                    rowsStr += $" ... 还有 {d.RowIndices.Count - 20} 行";

                _displayTable.Rows.Add(d.Value, d.Count, rowsStr);
            }
        }

        private void ApplyFilter()
        {
            var filter = _txtSearch.Text.Trim();
            if (string.IsNullOrEmpty(filter))
            {
                _grid.DataSource = _displayTable;
                return;
            }

            var filtered = _displayTable.AsEnumerable()
                .Where(r => (r.Field<string>("Value") ?? "").IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                         || (r.Field<string>("Rows") ?? "").IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            if (filtered.Count > 0)
                _grid.DataSource = filtered.CopyToDataTable();
            else
            {
                var emptyTable = _displayTable.Clone();
                _grid.DataSource = emptyTable;
            }
        }

        private void Grid_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var sourceTable = _grid.DataSource as DataTable;
            if (sourceTable == null) return;

            var row = sourceTable.Rows[e.RowIndex];
            var value = row["Value"]?.ToString() ?? "";
            var dup = _duplicates.FirstOrDefault(d => d.Value == value);
            if (dup != null && dup.RowIndices.Count > 0)
            {
                var targetRow = dup.RowIndices[0];
                _owner.NavigateToRow(targetRow);
                Close();
            }
        }

        private void BtnCopy_Click(object? sender, EventArgs e)
        {
            var sb = new StringBuilder();
            sb.AppendLine("重复值\t出现次数\t所在行号");
            foreach (var d in _duplicates.OrderByDescending(d => d.Count))
            {
                var rowNums = string.Join(", ", d.RowIndices.Take(50).Select(idx =>
                    idx < _fileData.Rows.Count
                        ? (_fileData.Rows[idx].OriginalRowIndex + 1).ToString()
                        : (idx + 1).ToString()));
                if (d.RowIndices.Count > 50)
                    rowNums += $" ... 还有 {d.RowIndices.Count - 50} 行";
                sb.AppendLine($"{d.Value}\t{d.Count}\t{rowNums}");
            }
            Clipboard.SetText(sb.ToString());
            AlertForm.Show(this, "已复制到剪贴板", "提示", AlertType.Information);
        }
    }
}
