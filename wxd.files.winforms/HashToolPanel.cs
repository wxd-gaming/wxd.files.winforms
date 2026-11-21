using System;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>哈希页签：文本按 UTF-8 计算 MD5 / SHA256，输入时实时刷新。</summary>
    internal sealed class HashToolPanel : ToolPanelBase
    {
        private readonly TextBox _input = MakeEditor();
        private readonly TextBox _output = MakeEditor(readOnly: true);
        private readonly ThemeBordered _inputBorder;
        private readonly ThemeBordered _outputBorder;
        private readonly ThemeCheckBox _upper;
        private readonly SplitContainer _split = new SafeSplitContainer();

        internal HashToolPanel()
        {
            AddButton("复制结果", (_, _) => CopyToClipboard(_output.Text), primary: true);
            AddButton("清空", (_, _) =>
            {
                _input.Clear();
                _input.Focus();
            });
            _upper = AddCheckBox("大写输出");
            _upper.CheckedChanged += (_, _) => RefreshHash();

            _input.TextChanged += (_, _) => RefreshHash();

            _inputBorder = new ThemeBordered(_input) { Dock = DockStyle.Fill };
            _outputBorder = new ThemeBordered(_output) { Dock = DockStyle.Fill };

            _split.Dock = DockStyle.Fill;
            _split.Orientation = Orientation.Horizontal;
            _split.SplitterWidth = 6;
            _split.SizeChanged += (_, _) => ApplySplitterLayout();
            _split.Panel1.Controls.Add(_inputBorder);
            _split.Panel2.Controls.Add(_outputBorder);

            FinishChrome(_split);

            Load += (_, _) => ApplySplitterLayout();

            RefreshHash();
        }

        private bool _applyingSplitter;

        private void ApplySplitterLayout()
        {
            if (_applyingSplitter || _split.IsDisposed) return;
            int height = _split.Height;
            if (height < 100) return;
            int min = Math.Min(100, Math.Max(25, (height - _split.SplitterWidth) / 3));
            int max = Math.Max(min, height - _split.SplitterWidth - min);
            int distance = Math.Clamp(height / 2, min, max);
            _applyingSplitter = true;
            try
            {
                if (_split.Panel1MinSize != min) _split.Panel1MinSize = min;
                if (_split.Panel2MinSize != min) _split.Panel2MinSize = min;
                if (_split.SplitterDistance != distance) _split.SplitterDistance = distance;
            }
            catch (InvalidOperationException) { }
            finally { _applyingSplitter = false; }
        }

        protected override void ApplyThemeCore()
        {
            _inputBorder.ApplyTheme();
            _outputBorder.ApplyTheme();
            ApplyEditorTheme(_input);
            ApplyEditorTheme(_output);
            _split.BackColor = Theme.P.Divider;
            _split.Panel1.BackColor = Theme.P.Surface;
            _split.Panel2.BackColor = Theme.P.Surface;
        }

        private void RefreshHash()
        {
            string text = _input.Text;
            bool upper = _upper.Checked;
            int bytes = System.Text.Encoding.UTF8.GetByteCount(text);

            _output.Text = string.Join(Environment.NewLine,
                "MD5    : " + HashToolEngine.Md5Hex(text, upper),
                "SHA256 : " + HashToolEngine.Sha256Hex(text, upper),
                "",
                $"字符数 : {text.Length}",
                $"UTF-8  : {bytes} 字节");

            SetStatus(text.Length == 0
                ? "空文本也有哈希值（UTF-8 0 字节）"
                : "已计算 · 文本按 UTF-8 取字节");
        }
    }
}
