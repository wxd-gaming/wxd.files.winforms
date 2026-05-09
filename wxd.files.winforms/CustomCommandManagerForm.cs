using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms;

/// <summary>自定义命令管理面板</summary>
public class CustomCommandManagerForm : Form
{
    private DataGridView _dgv;
    private Button _btnAdd;
    private Button _btnDelete;
    private Button _btnCopy;
    private Button _btnMoveUp;
    private Button _btnMoveDown;
    private Button _btnOk;
    private Button _btnCancel;

    private BindingList<CustomCommandItem> _binding;

    public CustomCommandManagerForm()
    {
        Text = "自定义命令管理";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(1250, 650);
        MinimizeBox = false;
        MaximizeBox = false;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        ShowInTaskbar = false;

        // --- DataGridView ---
        _dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = true,
            RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            AllowUserToOrderColumns = false
        };

        var colName = new DataGridViewTextBoxColumn
        {
            HeaderText = "命令名",
            DataPropertyName = "Name",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 25
        };
        var colCmd = new DataGridViewTextBoxColumn
        {
            HeaderText = "执行的命令",
            Name = "colCmd",
            DataPropertyName = "Command",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 40,
            DefaultCellStyle = { WrapMode = DataGridViewTriState.True }
        };
        var colWorkDir = new DataGridViewTextBoxColumn
        {
            HeaderText = "工作目录",
            DataPropertyName = "WorkDir",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 30
        };
        var colShowWindow = new DataGridViewCheckBoxColumn
        {
            HeaderText = "显示窗口",
            DataPropertyName = "ShowWindow",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 10
        };
        _dgv.Columns.Add(colName);
        _dgv.Columns.Add(colCmd);
        _dgv.Columns.Add(colWorkDir);
        _dgv.Columns.Add(colShowWindow);

        // 按钮颜色列（只读，点击弹出取色器）
        var colColor = new DataGridViewTextBoxColumn
        {
            HeaderText = "按钮颜色",
            Name = "colColor",
            DataPropertyName = "ButtonColor",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 12,
            ReadOnly = true,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        _dgv.Columns.Add(colColor);

        // 绘制颜色单元格
        _dgv.CellFormatting += (s, e) =>
        {
            if (_dgv.Columns[e.ColumnIndex].Name == "colColor" && e.Value is string colorStr)
            {
                try
                {
                    var color = ColorTranslator.FromHtml(colorStr);
                    e.CellStyle.BackColor = color;
                    e.CellStyle.ForeColor = color; // 隐藏文字，只显示纯色块
                    e.CellStyle.SelectionBackColor = color;
                    e.CellStyle.SelectionForeColor = color;
                }
                catch { }
            }
        };

        // 点击颜色列弹出取色器
        _dgv.CellClick += (s, e) =>
        {
            if (_dgv.Columns[e.ColumnIndex].Name == "colColor" && e.RowIndex >= 0)
            {
                var row = _dgv.Rows[e.RowIndex];
                if (row.DataBoundItem is CustomCommandItem item)
                {
                    using var cd = new ColorDialog();
                    try { cd.Color = ColorTranslator.FromHtml(item.ButtonColor); }
                    catch { cd.Color = Color.FromArgb(123, 31, 162); }
                    if (cd.ShowDialog() == DialogResult.OK)
                    {
                        item.ButtonColor = ColorTranslator.ToHtml(cd.Color);
                        _dgv.InvalidateCell(e.ColumnIndex, e.RowIndex);
                    }
                }
            }
        };

        // 取消颜色列的选择焦点（避免全选文字变白不可见）
        _dgv.SelectionChanged += (s, e) =>
        {
            if (_dgv.CurrentCell?.OwningColumn?.Name == "colColor")
            {
                // 跳转到名称列
                var row = _dgv.CurrentCell.RowIndex;
                if (_dgv.Rows[row].Cells[0].Visible)
                    _dgv.CurrentCell = _dgv.Rows[row].Cells[0];
            }
        };

        // 命令列支持多行编辑
        _dgv.EditingControlShowing += OnEditingControlShowing;

        // 行高自适应：根据命令列内容自动调整行高
        _dgv.RowsAdded += (s, e) => AutoSizeAllRows();
        _dgv.CellValueChanged += (s, e) =>
        {
            if (e.ColumnIndex == colCmd.Index)
                AutoSizeRow(e.RowIndex);
        };

        // --- 底部按钮 ---
        _btnAdd = new Button { Text = "新增", Width = 75, Height = 28 };
        _btnAdd.Click += (s, e) => AddRow();

        _btnCopy = new Button { Text = "复制", Width = 75, Height = 28 };
        _btnCopy.Click += (s, e) => CopySelectedRow();

        _btnDelete = new Button { Text = "删除", Width = 75, Height = 28 };
        _btnDelete.Click += (s, e) => DeleteSelectedRow();

        _btnMoveUp = new Button { Text = "上移", Width = 65, Height = 28 };
        _btnMoveUp.Click += (s, e) => MoveRow(-1);

        _btnMoveDown = new Button { Text = "下移", Width = 65, Height = 28 };
        _btnMoveDown.Click += (s, e) => MoveRow(1);

        _btnOk = new Button
        {
            Text = "确定",
            Width = 90,
            Height = 28,
            DialogResult = DialogResult.OK,
            BackColor = Color.FromArgb(0, 122, 204),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            FlatAppearance = { BorderSize = 0 }
        };

        _btnCancel = new Button
        {
            Text = "取消",
            Width = 90,
            Height = 28,
            DialogResult = DialogResult.Cancel
        };

        var bottomPanel = new Panel
        {
            Height = 44,
            Dock = DockStyle.Bottom,
            Padding = new Padding(8)
        };
        _btnAdd.Location = new Point(8, 8);
        _btnCopy.Location = new Point(_btnAdd.Right + 6, 8);
        _btnDelete.Location = new Point(_btnCopy.Right + 6, 8);
        _btnMoveUp.Location = new Point(_btnDelete.Right + 12, 8);
        _btnMoveDown.Location = new Point(_btnMoveUp.Right + 6, 8);
        _btnCancel.Location = new Point(bottomPanel.Width - 90 - 8, 8);
        _btnOk.Location = new Point(_btnCancel.Left - 90 - 6, 8);

        bottomPanel.Controls.AddRange(new Control[] { _btnAdd, _btnCopy, _btnDelete, _btnMoveUp, _btnMoveDown, _btnOk, _btnCancel });
        bottomPanel.Resize += (s, e) =>
        {
            _btnCancel.Left = bottomPanel.Width - 90 - 8;
            _btnOk.Left = _btnCancel.Left - 90 - 6;
        };

        // 主体区域：DataGridView
        var mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8, 8, 8, 0) };
        mainPanel.Controls.Add(_dgv);

        Controls.Add(mainPanel);
        Controls.Add(bottomPanel);

        // 从配置加载
        var list = CsvSettings.GetCustomCommands();
        _binding = new BindingList<CustomCommandItem>(list);
        _dgv.DataSource = _binding;

        AcceptButton = _btnOk;
        CancelButton = _btnCancel;
    }

    private void AddRow()
    {
        var item = new CustomCommandItem { Name = "新命令", Command = "", WorkDir = "" };
        _binding.Add(item);
        // 定位到最后一行并进入编辑
        _dgv.CurrentCell = _dgv.Rows[_binding.Count - 1].Cells[0];
        _dgv.BeginEdit(true);
    }

    private void DeleteSelectedRow()
    {
        if (_dgv.CurrentRow == null || _dgv.CurrentRow.IsNewRow) return;
        var item = (CustomCommandItem)_dgv.CurrentRow.DataBoundItem;
        _binding.Remove(item);
    }

    /// <summary>复制选中的命令行为新行</summary>
    private void CopySelectedRow()
    {
        if (_dgv.CurrentRow == null || _dgv.CurrentRow.IsNewRow) return;
        var source = (CustomCommandItem)_dgv.CurrentRow.DataBoundItem;
        var copy = new CustomCommandItem
        {
            Name = source.Name + " (副本)",
            Command = source.Command,
            WorkDir = source.WorkDir,
            ShowWindow = source.ShowWindow,
            ButtonColor = source.ButtonColor
        };
        var idx = _binding.IndexOf(source);
        _binding.Insert(idx + 1, copy);
        // 定位到复制行并进入编辑
        _dgv.CurrentCell = _dgv.Rows[idx + 1].Cells[0];
        _dgv.BeginEdit(true);
    }

    /// <summary>上移/下移选中的行</summary>
    private void MoveRow(int direction)
    {
        if (_dgv.CurrentRow == null || _dgv.CurrentRow.IsNewRow) return;
        var item = (CustomCommandItem)_dgv.CurrentRow.DataBoundItem;
        var idx = _binding.IndexOf(item);
        var newIdx = idx + direction;
        if (newIdx < 0 || newIdx >= _binding.Count) return;

        _binding.RemoveAt(idx);
        _binding.Insert(newIdx, item);
        _dgv.CurrentCell = _dgv.Rows[newIdx].Cells[0];
    }

    /// <summary>自动调整所有行的命令列高度</summary>
    private void AutoSizeAllRows()
    {
        foreach (DataGridViewRow row in _dgv.Rows)
            AutoSizeRow(row.Index);
    }

    /// <summary>根据命令列内容自动调整指定行高，容纳多行文字</summary>
    private void AutoSizeRow(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= _dgv.Rows.Count) return;
        var row = _dgv.Rows[rowIndex];
        var cell = row.Cells["colCmd"] ?? _dgv.Rows[rowIndex].Cells[1];
        if (cell.Value == null) return;

        var text = cell.Value.ToString() ?? "";
        if (string.IsNullOrEmpty(text))
        {
            row.Height = 52;
            return;
        }
        // 根据换行符估算行数，每行约16像素，最小基础高度60，再加30px间隔更清晰
        var lines = text.Split('\n').Length;
        var estimatedHeight = Math.Max(lines * 16, 60) + 30;
        row.Height = Math.Min(estimatedHeight, 230);
    }

    /// <summary>进入编辑时，将命令列设为多行模式</summary>
    private void OnEditingControlShowing(object? sender, DataGridViewEditingControlShowingEventArgs e)
    {
        if (_dgv.CurrentCell?.OwningColumn?.DataPropertyName == "Command"
            && e.Control is TextBox tb)
        {
            tb.Multiline = true;
            // 设置合适高度以容纳多行输入
            _dgv.Rows[_dgv.CurrentCell.RowIndex].Height = 90;
        }
    }

    /// <summary>确定时保存到配置</summary>
    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (DialogResult == DialogResult.OK)
        {
            // 结束可能的编辑状态
            _dgv.EndEdit();
            var list = new List<CustomCommandItem>(_binding);
            CsvSettings.SetCustomCommands(list);
        }
        base.OnFormClosing(e);
    }
}
