namespace wxd.files.winforms
{
    /// <summary>单元格编辑对话框</summary>
    public class CellEditDialog : Form
    {
        private RichTextBox _txtValue = new();
        private Button _btnSave = new();
        private Button _btnCancel = new();
        private Button _btnDefault = new();
        private Button _btnRaw = new();
        private Panel _pnlBottom = new();
        private Panel _pnlTop = new();
        private Label _lblSaveTime = new();  // 保存时间标签
        private Label _lblReadOnlyHint = new(); // 只读提示标签
        private readonly Action<string>? _onSave;

        public string CellValue => _txtValue.Text.Replace(GlobalConst.CrLf, GlobalConst.Lf);

        /// <summary>
        /// 将各种换行符转换为 Windows 格式 (\r\n)，供 RichTextBox 多行模式正确显示
        /// </summary>
        private static string NormalizeLineBreaksForDisplay(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            // 先将已有的 \r\n 保护，再处理 \r 和 \n，最后统一输出 \r\n
            return text.Replace(GlobalConst.CrLf, GlobalConst.Lf)
                       .Replace(GlobalConst.Cr, GlobalConst.Lf)
                       .Replace(GlobalConst.Lf, GlobalConst.CrLf);
        }

        public CellEditDialog(string initialValue, int row, int col, string fileName = "", string colName = "",
            Action<string>? onSave = null, bool readOnly = false, string readOnlyMessage = "")
        {
            _onSave = onSave;
            InitializeComponent(readOnly, readOnlyMessage);
            _txtValue.Text = NormalizeLineBreaksForDisplay(initialValue);

            string colDisplay = string.IsNullOrEmpty(colName)
                ? (col >= 0 && col < 26 ? ((char)('A' + col)).ToString() : $"Col{col}")
                : colName;
            if (!string.IsNullOrEmpty(fileName))
                Text = $"单元格编辑 — {Path.GetFileName(fileName)} — 列：{colDisplay}，第 {row + 1} 行";
            else
                Text = $"编辑单元格 — 行 {row + 1}  列 {colDisplay}({col})";
        }

        private static void SetActiveButton(Button active, Button inactive)
        {
            active.BackColor = Color.FromArgb(66, 165, 245);
            active.ForeColor = Color.White;
            active.FlatAppearance.BorderSize = 1;
            active.FlatAppearance.BorderColor = Color.FromArgb(30, 136, 229);
            inactive.BackColor = Color.FromArgb(235, 248, 255);
            inactive.ForeColor = Color.FromArgb(20, 60, 100);
            inactive.FlatAppearance.BorderSize = 1;
            inactive.FlatAppearance.BorderColor = Color.FromArgb(173, 209, 237);
        }

        private void InitializeComponent(bool readOnly = false, string readOnlyMessage = "")
        {
            SuspendLayout();

            Text = "编辑单元格";
            ClientSize = new Size(800, 600);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = true;
            MaximizeBox = true;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = Color.FromArgb(242, 249, 255);

            // ==================== 顶部格式切换面板 ====================
            _pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = Color.FromArgb(235, 245, 255),
                Padding = new Padding(8, 7, 8, 7)
            };

            _btnDefault = new Button
            {
                Text = "默认格式",
                Location = new Point(8, 7),
                Size = new Size(90, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 0, MouseOverBackColor = Color.FromArgb(227, 242, 253) }
            };

            _btnRaw = new Button
            {
                Text = "自动换行",
                Location = new Point(106, 7),
                Size = new Size(90, 28),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 0, MouseOverBackColor = Color.FromArgb(227, 242, 253) }
            };

            _btnDefault.Click += (s, e) =>
            {
                SetActiveButton(_btnDefault, _btnRaw);
                _txtValue.WordWrap = false;
                _txtValue.ScrollBars = RichTextBoxScrollBars.Both;
            };
            _btnRaw.Click += (s, e) =>
            {
                SetActiveButton(_btnRaw, _btnDefault);
                _txtValue.WordWrap = true;
                _txtValue.ScrollBars = RichTextBoxScrollBars.Both;
            };

            // 默认选中"默认格式"
            SetActiveButton(_btnDefault, _btnRaw);

            _pnlTop.Controls.Add(_btnDefault);
            _pnlTop.Controls.Add(_btnRaw);

            // ==================== 底部按钮面板 ====================
            _pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 52,
                BackColor = Color.FromArgb(235, 245, 255)
            };

            // 只读提示标签（左侧）
            _lblReadOnlyHint = new Label
            {
                Text = readOnly ? readOnlyMessage : "",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                ForeColor = Color.Red,
                Location = new Point(10, 14)
            };

            // 保存按钮
            _btnSave = new Button
            {
                Text = "保存",
                Size = new Size(80, 30),
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                BackColor = Color.FromArgb(66, 165, 245),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(30, 136, 229), MouseOverBackColor = Color.FromArgb(30, 136, 229) },
                Enabled = !readOnly
            };
            _btnSave.Click += (s, e) =>
            {
                _lblSaveTime.Text = $"已保存 {DateTime.Now:HH:mm:ss}";
                _onSave?.Invoke(CellValue);
            };

            // 保存时间标签
            _lblSaveTime = new Label
            {
                Text = "",
                Size = new Size(140, 28),
                Font = new Font("Microsoft YaHei UI", 8.5f),
                ForeColor = Color.FromArgb(76, 175, 80),
                TextAlign = ContentAlignment.MiddleRight
            };

            // 取消（关闭）按钮
            _btnCancel = new Button
            {
                Text = "关闭",
                Size = new Size(80, 30),
                Font = new Font("Microsoft YaHei UI", 9f),
                BackColor = Color.FromArgb(235, 248, 255),
                ForeColor = Color.FromArgb(20, 60, 100),
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(173, 209, 237), MouseOverBackColor = Color.FromArgb(200, 225, 245) }
            };
            _btnCancel.Click += (s, e) => { Close(); };

            _pnlBottom.Controls.Add(_lblReadOnlyHint);
            _pnlBottom.Controls.Add(_lblSaveTime);
            _pnlBottom.Controls.Add(_btnSave);
            _pnlBottom.Controls.Add(_btnCancel);

            // 窗口调整大小时重新定位按钮（保持在右下角）
            void LayoutButtons()
            {
                _btnCancel.Left = _pnlBottom.Width - _btnCancel.Width - 10;
                _btnSave.Left = _btnCancel.Left - _btnSave.Width - 8;
                _lblSaveTime.Left = _btnSave.Left - _lblSaveTime.Width - 8;
                _lblSaveTime.Top = 12;
                _btnSave.Top = 10;
                _btnCancel.Top = 10;
                _lblReadOnlyHint.Top = 14;
            }
            LayoutButtons();
            _pnlBottom.Resize += (s, e) => LayoutButtons();

            // ==================== 编辑区（RichTextBox）====================
            _txtValue = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10f),
                BorderStyle = BorderStyle.None,
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both,
                AcceptsTab = true,
                BackColor = Color.White
            };

            Controls.Add(_txtValue);
            Controls.Add(_pnlBottom);
            Controls.Add(_pnlTop);

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;

            ResumeLayout(false);
        }
    }
}
