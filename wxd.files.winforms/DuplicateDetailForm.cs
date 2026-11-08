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
        private ThemeField _fldSearch = new();
        private Label _lblSummary = new();
        private ThemeButton _btnCopy = new();
        private ThemeButton _btnClose = new();
        private Panel _topBar = new();
        private Panel _topDivider = new();
        private Panel _bar = new();
        private DataTable _displayTable = new();

        public DuplicateDetailForm(
            List<TableBrowserControl.DuplicateGroup> duplicates,
            FileTableData fileData,
            TableBrowserControl owner)
        {
            _duplicates = duplicates;
            _fileData = fileData;
            _owner = owner;

            InitializeComponent();
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
            SetupGrid();
            PopulateGrid();
        }

        private void InitializeComponent()
        {
            Text = "重复值详情";
            ClientSize = new Size(750, 550);
            MinimumSize = new Size(500, 350);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            ShowInTaskbar = false;

            var totalRedundant = _duplicates.Sum(d => d.Count - 1);
            _lblSummary = new Label
            {
                Text = $"共 {_duplicates.Count} 个重复值，{totalRedundant} 处冗余  |  双击某行可跳转到对应位置",
                AutoSize = false,
                Height = 30,
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Margin, 0),
                Font = Theme.Head
            };

            _topBar = new Panel { Dock = DockStyle.Top, Height = ThemeMetrics.BarHeight };
            _topDivider = new Panel { Dock = DockStyle.Bottom, Height = 1 };
            _topBar.Controls.Add(_topDivider);

            var lblSearch = new Label { Text = "搜索:", AutoSize = true };
            _fldSearch = new ThemeField { Placeholder = "按重复值或行号筛选..." };
            _fldSearch.Edit.TextChanged += (s, e) => ApplyFilter();
            _topBar.Controls.AddRange(new Control[] { lblSearch, _fldSearch });
            LayoutTopBar(lblSearch);
            _topBar.Resize += (_, _) => LayoutTopBar(lblSearch);

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
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            };
            _grid.CellDoubleClick += Grid_CellDoubleClick;

            _bar = DialogButtonBar.Create();
            _btnCopy = DialogButtonBar.Make("复制到剪贴板", BtnCopy_Click, primary: true);
            _btnClose = DialogButtonBar.Make("关闭", (s, e) => Close());
            _bar.Controls.AddRange(new Control[] { _btnCopy, _btnClose });
            LayoutBar();
            _bar.Resize += (_, _) => LayoutBar();

            // 停靠顺序：同一条边上后加入的靠外，所以摘要条要最后加
            Controls.Add(_grid);
            Controls.Add(_bar);
            Controls.Add(_topBar);
            Controls.Add(_lblSummary);
        }

        private void LayoutTopBar(Label lblSearch)
        {
            int rowH = ThemeMetrics.RowH;
            lblSearch.Location = new Point(ThemeMetrics.Margin, (ThemeMetrics.BarHeight - lblSearch.Height) / 2);
            _fldSearch.SetBounds(lblSearch.Right + ThemeMetrics.Gap, ThemeMetrics.RowPad,
                _topBar.Width - lblSearch.Right - ThemeMetrics.Margin * 2 - ThemeMetrics.Gap, rowH);
        }

        private void LayoutBar()
            => DialogButtonBar.LayoutRight(_bar, _btnClose, _btnCopy);

        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _topBar.BackColor = p.Chrome;
            _bar.BackColor = p.Chrome;
            _topDivider.BackColor = p.Divider;

            // 摘要条是提示性的问题清单，底色从危险色淡出来，文字保住对比度
            var signal = Theme.Signal(Color.FromArgb(180, 50, 30));
            _lblSummary.ForeColor = signal;
            _lblSummary.BackColor = Theme.Mix(signal, p.Surface, 0.88f);

            ThemeGrid.Apply(_grid);
            _fldSearch.ApplyTheme();
            _btnCopy.BackColor = p.Chrome;
            _btnClose.BackColor = p.Chrome;

            Invalidate(true);
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
                        ? (_fileData.Rows[idx].DisplayLineNumber ?? idx + 1).ToString()
                        : (idx + 1).ToString());
                var rowsStr = string.Join(", ", rowNums);
                if (d.RowIndices.Count > 20)
                    rowsStr += $" ... 还有 {d.RowIndices.Count - 20} 行";

                _displayTable.Rows.Add(d.Value, d.Count, rowsStr);
            }
        }

        private void ApplyFilter()
        {
            var filter = _fldSearch.Edit.Text.Trim();
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
                        ? (_fileData.Rows[idx].DisplayLineNumber ?? idx + 1).ToString()
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
