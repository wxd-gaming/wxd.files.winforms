namespace wxd.files.winforms
{
    /// <summary>行号跳转输入框：输入行号区显示的行号（源文件物理行号），回车定位到对应行</summary>
    public class LineNumberJumpDialog : Form
    {
        private readonly TextBox _txtLine = new();
        private readonly Label _lblRange = new();
        private readonly Button _btnOk = new();
        private readonly Button _btnCancel = new();

        /// <summary>确认后的行号（取消时为 -1）</summary>
        public long LineNumber { get; private set; } = -1;

        public LineNumberJumpDialog(int maxLine, long initialLine = 1)
        {
            Text = "跳转到行号";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9f);
            ClientSize = new Size(280, 118);

            _txtLine.Location = new Point(12, 14);
            _txtLine.Size = new Size(256, 23);
            _txtLine.Text = initialLine.ToString();

            _lblRange.Text = $"有效行号：1–{maxLine}";
            _lblRange.AutoSize = true;
            _lblRange.ForeColor = Color.FromArgb(110, 110, 110);
            _lblRange.Location = new Point(12, 42);

            _btnOk.Text = "跳转";
            _btnOk.Size = new Size(80, 28);
            _btnOk.Location = new Point(100, 72);
            _btnOk.BackColor = Color.FromArgb(144, 180, 210);
            _btnOk.ForeColor = Color.White;
            _btnOk.FlatStyle = FlatStyle.Flat;

            _btnCancel.Text = "取消";
            _btnCancel.Size = new Size(80, 28);
            _btnCancel.Location = new Point(188, 72);
            _btnCancel.DialogResult = DialogResult.Cancel;

            _btnOk.Click += (s, e) =>
            {
                if (long.TryParse(_txtLine.Text.Trim(), out var line) && line >= 1 && line <= maxLine)
                {
                    LineNumber = line;
                    DialogResult = DialogResult.OK;
                    Close();
                    return;
                }
                _lblRange.ForeColor = Color.FromArgb(200, 30, 30);
                _lblRange.Text = $"请输入 1–{maxLine} 之间的行号";
            };
            // 回车直接跳转，Esc 取消
            AcceptButton = _btnOk;
            CancelButton = _btnCancel;

            Controls.AddRange(new Control[] { _txtLine, _lblRange, _btnOk, _btnCancel });
            _txtLine.Select();
            _txtLine.SelectAll();
        }
    }
}
