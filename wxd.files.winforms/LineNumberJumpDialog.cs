namespace wxd.files.winforms
{
    /// <summary>行号跳转输入框：输入行号区显示的行号（源文件物理行号），回车定位到对应行</summary>
    public class LineNumberJumpDialog : Form
    {
        private readonly ThemeField _fldLine = new();
        private readonly Label _lblRange = new();
        private readonly ThemeButton _btnOk;
        private readonly ThemeButton _btnCancel;
        private readonly Panel _bar;
        private readonly int _maxLine;

        /// <summary>确认后的行号（取消时为 -1）</summary>
        public long LineNumber { get; private set; } = -1;

        public LineNumberJumpDialog(int maxLine, long initialLine = 1)
        {
            _maxLine = maxLine;
            Text = "跳转到行号";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;

            int m = ThemeMetrics.Margin;
            int gap = ThemeMetrics.Gap;
            int bodyW = 280 - m * 2;

            _lblRange.AutoSize = true;
            _lblRange.Text = $"有效行号：1–{_maxLine}";
            _lblRange.Location = new Point(m, m);

            _fldLine.SetBounds(m, _lblRange.Bottom + gap, bodyW, ThemeMetrics.RowH);
            _fldLine.Edit.Text = initialLine.ToString();

            _bar = DialogButtonBar.Create();
            _btnCancel = DialogButtonBar.Make("取消");
            _btnCancel.DialogResult = DialogResult.Cancel;
            _btnOk = DialogButtonBar.Make("跳转", OnOkClick, primary: true);
            _bar.Controls.AddRange(new Control[] { _btnCancel, _btnOk });

            // 停靠顺序：同一条边上后加入的靠外，正文先加、按钮条后加
            Controls.Add(_fldLine);
            Controls.Add(_lblRange);
            Controls.Add(_bar);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;

            ClientSize = new Size(280, _bar.Top + ThemeMetrics.BarHeight);
            DialogButtonBar.LayoutRight(_bar, _btnCancel, _btnOk);

            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;

            _fldLine.Edit.Select();
            _fldLine.Edit.SelectAll();
        }

        private void ApplyTheme()
        {
            Font = Theme.Body;
            BackColor = Theme.P.Window;
            _bar.BackColor = Theme.P.Chrome;
            _lblRange.Font = Theme.Caption;
            _fldLine.ApplyTheme();
            // 报错时前景被换成 Danger，切皮肤后按旧令牌→新令牌平移回来
            ThemeRecolor.Apply(_lblRange);
            _lblRange.Invalidate();
        }

        private void OnOkClick(object? sender, EventArgs e)
        {
            if (long.TryParse(_fldLine.Edit.Text.Trim(), out var line) && line >= 1 && line <= _maxLine)
            {
                LineNumber = line;
                DialogResult = DialogResult.OK;
                Close();
                return;
            }
            // 出错提示沿用同一个位置，避免对话框在报错时改变高度
            _lblRange.ForeColor = Theme.P.Danger;
            _lblRange.Text = $"请输入 1–{_maxLine} 之间的行号";
        }
    }
}
