using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
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
            ReadOnly = true, // 禁用单元格内编辑，改为双击/右键弹窗编辑
            SortMode = DataGridViewColumnSortMode.NotSortable,
            DefaultCellStyle = { WrapMode = DataGridViewTriState.True }
        };
        var colWorkDir = new DataGridViewTextBoxColumn
        {
            HeaderText = "工作目录",
            DataPropertyName = "WorkDir",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 30,
            ReadOnly = true,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        var colShowWindow = new DataGridViewCheckBoxColumn
        {
            HeaderText = "显示窗口",
            DataPropertyName = "ShowWindow",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 10
        };
        var colGroupName = new DataGridViewComboBoxColumn
        {
            HeaderText = "分组",
            Name = "colGroupName",
            DataPropertyName = "GroupName",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 15,
            SortMode = DataGridViewColumnSortMode.NotSortable,
            FlatStyle = FlatStyle.Flat
        };
        // 初始填充已有分组名（后续动态更新）
        colGroupName.Items.Add(""); // 空值 = 独立按钮（保证数据验证通过）
        _dgv.Columns.Add(colName);
        _dgv.Columns.Add(colCmd);
        _dgv.Columns.Add(colWorkDir);
        _dgv.Columns.Add(colShowWindow);
        _dgv.Columns.Add(colGroupName);

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
            // 分组列：空值显示"(独立按钮)"
            if (_dgv.Columns[e.ColumnIndex].Name == "colGroupName" && e.RowIndex >= 0)
            {
                // 直接从数据源获取真实 GroupName，避免 DataError 回滚导致 e.Value 为空
                var realGroupName = "";
                if (_dgv.Rows[e.RowIndex].DataBoundItem is CustomCommandItem item)
                    realGroupName = item.GroupName ?? "";
                if (string.IsNullOrEmpty(realGroupName))
                {
                    e.Value = "(独立按钮)";
                }
                else
                {
                    e.Value = realGroupName;
                }
                e.FormattingApplied = true;
            }
        };

        // 双击工作目录列弹出目录选择对话框；双击命令列弹出文本编辑对话框
        _dgv.CellDoubleClick += (s, e) =>
        {
            if (e.RowIndex < 0) return;
            var dataProp = _dgv.Columns[e.ColumnIndex].DataPropertyName;
            if (dataProp == "WorkDir")
            {
                var row = _dgv.Rows[e.RowIndex];
                if (row.DataBoundItem is CustomCommandItem item)
                {
                    using var fbd = new FolderBrowserDialog();
                    if (!string.IsNullOrEmpty(item.WorkDir))
                        fbd.SelectedPath = item.WorkDir;
                    if (fbd.ShowDialog() == DialogResult.OK)
                    {
                        item.WorkDir = fbd.SelectedPath;
                    }
                }
            }
            else if (dataProp == "Command")
            {
                EditCommandViaDialog(e.RowIndex);
            }
        };

        // 右键命令列弹出文本编辑对话框
        _dgv.CellMouseDown += (s, e) =>
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0
                && _dgv.Columns[e.ColumnIndex].DataPropertyName == "Command")
            {
                EditCommandViaDialog(e.RowIndex);
            }
        };

        // 分组列编辑控件：动态刷新分组下拉选项
        ComboBox? _currentCombo = null; // 追踪当前编辑的 ComboBox，避免重复注册 Leave
        _dgv.EditingControlShowing += (s, e) =>
        {
            if (_dgv.CurrentCell?.OwningColumn?.Name == "colGroupName" && e.Control is ComboBox combo)
            {
                // 先移除旧的 Leave 处理
                if (_currentCombo != null)
                    _currentCombo.Leave -= OnComboLeave;
                _currentCombo = combo;

                // 收集已有分组名
                var existingGroups = new HashSet<string>();
                foreach (var item in _binding)
                {
                    if (!string.IsNullOrWhiteSpace(item.GroupName))
                        existingGroups.Add(item.GroupName);
                }
                combo.Items.Clear();
                combo.Items.Add(""); // 空值 = 独立按钮
                foreach (var g in existingGroups.OrderBy(x => x))
                    combo.Items.Add(g);
                // 允许用户输入新的分组名
                combo.DropDownStyle = ComboBoxStyle.DropDown;
                combo.Leave += OnComboLeave;
            }
        };

        // 处理 DataError，阻止无限弹窗，并接受用户输入的新分组名
        _dgv.DataError += (s, e) =>
        {
            // 阻止默认错误对话框弹出
            e.ThrowException = false;

            // 如果是分组名列，接受用户输入的新值，并同步到 Items 列表
            if (_dgv.Columns[e.ColumnIndex].Name == "colGroupName"
                && (e.Context & DataGridViewDataErrorContexts.Commit) != 0
                && e.RowIndex >= 0)
            {
                // 获取编辑控件中用户实际输入的内容
                if (_dgv.EditingControl is ComboBox combo)
                {
                    var inputText = combo.Text?.Trim();
                    if (!string.IsNullOrEmpty(inputText))
                    {
                        // 1. 将新值加入列级别的 Items，下次验证就能通过
                        if (!colGroupName.Items.Contains(inputText))
                            colGroupName.Items.Add(inputText);

                        // 2. 强制设置单元格值为用户输入
                        _dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = inputText;

                        // 3. 结束编辑让 DataGridView 重新验证
                        _dgv.EndEdit();
                        return;
                    }
                }
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
            BackColor = Color.FromArgb(66, 165, 245),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229) }
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

        // 同步列级 Items：将所有已有的分组名加入 ComboBox 列，避免显示空白
        foreach (var item in list)
        {
            var g = item.GroupName ?? "";
            if (!string.IsNullOrEmpty(g) && !colGroupName.Items.Contains(g))
                colGroupName.Items.Add(g);
        }

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
            ButtonColor = source.ButtonColor,
            GroupName = source.GroupName
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

    /// <summary>通过弹窗编辑指定行的命令（双击/右键触发）</summary>
    private void EditCommandViaDialog(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= _dgv.Rows.Count) return;
        if (_dgv.Rows[rowIndex].DataBoundItem is not CustomCommandItem item) return;

        var cmdColIndex = _dgv.Columns["colCmd"].Index;
        using var dlg = new CellEditDialog(item.Command, rowIndex, cmdColIndex, "", "执行的命令",
            customTitle: $"编辑命令 — {item.Name}", closeOnSave: true);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            item.Command = dlg.CellValue;          // CellValue 已将换行归一化为 Lf
            _dgv.InvalidateCell(cmdColIndex, rowIndex); // 刷新单元格显示
            AutoSizeRow(rowIndex);                 // 按新内容重算行高
        }
    }

    /// <summary>分组列 ComboBox 失去焦点时，处理用户输入的新分组名</summary>
    private void OnComboLeave(object? sender, EventArgs e)
    {
        if (sender is ComboBox combo)
        {
            // 用户输入新分组名后，确保数据绑定更新
            if (combo.SelectedItem == null && !string.IsNullOrWhiteSpace(combo.Text))
            {
                _dgv.CurrentCell.Value = combo.Text.Trim();
            }
        }
    }

    /// <summary>拦截 Ctrl+C：编辑状态下只复制单元格内选中的文本，不复制整行</summary>
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == (Keys.Control | Keys.C))
        {
            // 检测是否处于编辑状态，且编辑控件中有选中文本
            if (_dgv.IsCurrentCellInEditMode && _dgv.EditingControl is TextBox tb && tb.SelectionLength > 0)
            {
                Clipboard.SetText(tb.SelectedText);
                return true; // 已处理，不再交给 DGV
            }
        }
        return base.ProcessCmdKey(ref msg, keyData);
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
