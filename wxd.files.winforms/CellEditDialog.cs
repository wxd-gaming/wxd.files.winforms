namespace wxd.files.winforms
{
    /// <summary>单元格编辑对话框</summary>
    public class CellEditDialog : Form
    {
        private RichTextBox _txtValue = new();
        private ThemeButton _btnSave = new();
        private ThemeButton _btnCancel = new();
        // 两种互斥的显示方式，用分段控件代替两颗会互相抢底色的切换按钮
        private ThemeSegmented _segView = new();
        private Panel _pnlBottom = new();
        private Panel _pnlTop = new();
        private Panel _topDivider = new();
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
            Action<string>? onSave = null, bool readOnly = false, string readOnlyMessage = "",
            bool closeOnSave = false, string? customTitle = null)
        {
            _onSave = onSave;
            InitializeComponent(readOnly, readOnlyMessage, closeOnSave);
            _txtValue.Text = NormalizeLineBreaksForDisplay(initialValue);

            ApplyTheme();
            // 独立对话框，主窗翻明暗时遍历不到，只能自己订阅
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;

            // 自定义标题：直接用作窗口标题（模态弹窗场景）
            if (!string.IsNullOrEmpty(customTitle))
            {
                Text = customTitle;
                if (closeOnSave)
                    _btnCancel.Text = "取消";
                return;
            }

            string colDisplay = string.IsNullOrEmpty(colName)
                ? (col >= 0 && col < 26 ? ((char)('A' + col)).ToString() : $"Col{col}")
                : colName;
            if (!string.IsNullOrEmpty(fileName))
                Text = $"单元格编辑 — {Path.GetFileName(fileName)} — 列：{colDisplay}，第 {row + 1} 行";
            else
                Text = $"编辑单元格 — 行 {row + 1}  列 {colDisplay}({col})";
        }

        private void InitializeComponent(bool readOnly = false, string readOnlyMessage = "", bool closeOnSave = false)
        {
            SuspendLayout();

            Text = "编辑单元格";
            ClientSize = new Size(800, 600);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = true;
            MaximizeBox = true;

            // ==================== 顶部格式切换面板 ====================
            _pnlTop = new Panel { Dock = DockStyle.Top, Height = ThemeMetrics.BarHeight };
            _topDivider = new Panel { Dock = DockStyle.Bottom, Height = 1 };
            _pnlTop.Controls.Add(_topDivider);

            _segView = new ThemeSegmented { Items = new[] { "默认格式", "自动换行" }, ActiveIndex = 0 };
            _segView.SegmentClicked += i =>
            {
                _txtValue.WordWrap = i == 1;
                _txtValue.ScrollBars = RichTextBoxScrollBars.Both;
            };
            _pnlTop.Controls.Add(_segView);
            _pnlTop.Resize += (_, _) => LayoutTop();

            // ==================== 底部按钮面板 ====================
            _pnlBottom = DialogButtonBar.Create();

            _lblReadOnlyHint = new Label
            {
                Text = readOnly ? readOnlyMessage : "",
                AutoSize = true,
                Font = Theme.Body
            };

            _btnSave = DialogButtonBar.Make("保存", null, primary: true);
            _btnSave.Enabled = !readOnly;
            _btnSave.Click += (s, e) =>
            {
                if (closeOnSave)
                {
                    // 模态模式：保存后立即关闭，由调用方读取 CellValue
                    _onSave?.Invoke(CellValue);
                    DialogResult = DialogResult.OK;
                    Close();
                }
                else
                {
                    _lblSaveTime.Text = $"已保存 {DateTime.Now:HH:mm:ss}";
                    _onSave?.Invoke(CellValue);
                    LayoutBottom();
                }
            };

            _lblSaveTime = new Label
            {
                Text = "",
                Size = new Size(140, 28),
                TextAlign = ContentAlignment.MiddleRight
            };

            _btnCancel = DialogButtonBar.Make("关闭");
            _btnCancel.Click += (s, e) => { Close(); };

            _pnlBottom.Controls.Add(_lblReadOnlyHint);
            _pnlBottom.Controls.Add(_lblSaveTime);
            _pnlBottom.Controls.AddRange(new Control[] { _btnSave, _btnCancel });
            LayoutBottom();
            _pnlBottom.Resize += (_, _) => LayoutBottom();

            // ==================== 编辑区（RichTextBox）====================
            _txtValue = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both,
                AcceptsTab = true
            };

            // 停靠顺序：后加入的靠外，所以编辑区先加、页眉页脚后加
            Controls.Add(_txtValue);
            Controls.Add(_pnlBottom);
            Controls.Add(_pnlTop);

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;

            ResumeLayout(false);
        }

        private void LayoutTop()
            => _segView.Location = new Point(ThemeMetrics.Margin,
                (_pnlTop.Height - _topDivider.Height - _segView.Height) / 2);

        /// <summary>按钮整组靠右，保存时间和只读提示分别贴在按钮左边和容器左边。</summary>
        private void LayoutBottom()
        {
            DialogButtonBar.LayoutRight(_pnlBottom, _btnSave, _btnCancel);
            int y = (_pnlBottom.Height - _lblSaveTime.Height) / 2;
            _lblSaveTime.Location = new Point(_btnSave.Left - _lblSaveTime.Width - ThemeMetrics.Gap, y);
            _lblReadOnlyHint.Location = new Point(ThemeMetrics.Margin,
                (_pnlBottom.Height - _lblReadOnlyHint.Height) / 2);
        }

        /// <summary>取色全部走当前皮肤令牌；换皮肤或翻明暗时由 Theme.Changed 重新走一遍。</summary>
        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _pnlTop.BackColor = p.Chrome;
            _pnlBottom.BackColor = p.Chrome;
            _topDivider.BackColor = p.Divider;

            _txtValue.BackColor = p.Surface;
            _txtValue.ForeColor = p.Text;
            _txtValue.Font = Theme.Code;

            _lblSaveTime.ForeColor = Theme.Signal(Color.FromArgb(22, 163, 74));
            _lblReadOnlyHint.ForeColor = p.Danger;

            _segView.RefreshTheme();
            LayoutTop();
            LayoutBottom();
            Invalidate(true);
        }
    }
}
