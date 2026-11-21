using System;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 编码页签：左边输入、右边输出，支持 Unicode 转义、UTF-8 Hex、
    /// Base64、URL 和 HTML 的编码/解码，文本统一按 UTF-8 处理。
    /// </summary>
    internal sealed class EncodingToolPanel : ToolPanelBase
    {
        private readonly TextBox _input = MakeEditor();
        private readonly TextBox _output = MakeEditor(readOnly: true);
        private readonly ThemeBordered _inputBorder;
        private readonly ThemeBordered _outputBorder;
        private readonly ThemeCheckBox _asciiOnly;
        private readonly SplitContainer _split = new SafeSplitContainer();

        internal EncodingToolPanel()
        {
            Top.Height = 84;
            Top.WrapContents = true;

            AddButton("Unicode 转义", (_, _) => Run(text => TextToolResult.Ok(
                EncodingToolEngine.EscapeUnicode(text, _asciiOnly.Checked)), "Unicode 转义"), primary: true);
            AddButton("Unicode 还原", (_, _) =>
                Run(EncodingToolEngine.UnescapeUnicode, "Unicode 还原"));
            _asciiOnly = AddCheckBox("仅转义非 ASCII");

            AddButton("UTF-8 Hex", (_, _) => Run(text => TextToolResult.Ok(
                EncodingToolEngine.ToUtf8Hex(text)), "UTF-8 Hex"));
            AddButton("Hex → 文本", (_, _) =>
                Run(EncodingToolEngine.FromUtf8Hex, "Hex 转文本"));
            AddButton("Base64 编码", (_, _) => Run(text => TextToolResult.Ok(
                EncodingToolEngine.ToBase64(text)), "Base64 编码"));
            AddButton("Base64 解码", (_, _) =>
                Run(EncodingToolEngine.FromBase64, "Base64 解码"));
            AddButton("URL 编码", (_, _) => Run(text => TextToolResult.Ok(
                EncodingToolEngine.UrlEncode(text)), "URL 编码"));
            AddButton("URL 解码", (_, _) => Run(text => TextToolResult.Ok(
                EncodingToolEngine.UrlDecode(text)), "URL 解码"));
            AddButton("HTML 编码", (_, _) => Run(text => TextToolResult.Ok(
                EncodingToolEngine.HtmlEncode(text)), "HTML 编码"));
            AddButton("HTML 解码", (_, _) => Run(text => TextToolResult.Ok(
                EncodingToolEngine.HtmlDecode(text)), "HTML 解码"));
            AddButton("复制结果", (_, _) => CopyToClipboard(_output.Text));
            AddButton("结果→输入", (_, _) =>
            {
                _input.Text = _output.Text;
                SetStatus("已把结果移到输入");
            });

            _inputBorder = new ThemeBordered(_input) { Dock = DockStyle.Fill };
            _outputBorder = new ThemeBordered(_output) { Dock = DockStyle.Fill };

            _split.Dock = DockStyle.Fill;
            _split.Orientation = Orientation.Vertical;
            _split.SplitterWidth = 6;
            _split.SizeChanged += (_, _) => ApplySplitterLayout();
            _split.Panel1.Controls.Add(_inputBorder);
            _split.Panel2.Controls.Add(_outputBorder);

            FinishChrome(_split);

            Load += (_, _) => ApplySplitterLayout();

            SetStatus("左边输入、右边输出；文本按 UTF-8 取字节，解码失败会明确报错");
        }

        private bool _applyingSplitter;

        private void ApplySplitterLayout()
        {
            if (_applyingSplitter || _split.IsDisposed) return;
            int width = _split.Width;
            if (width < 120) return;
            int min = Math.Min(180, Math.Max(25, (width - _split.SplitterWidth) / 3));
            int max = Math.Max(min, width - _split.SplitterWidth - min);
            int distance = Math.Clamp(width / 2, min, max);
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

        private void Run(Func<string, TextToolResult> transform, string action)
        {
            if (_input.TextLength == 0)
            {
                SetStatus("请先在左侧输入内容", error: true);
                return;
            }

            var result = transform(_input.Text);
            if (!result.Success)
            {
                SetStatus(result.Error ?? "转换失败", error: true);
                return;
            }

            _output.Text = result.Text;
            _output.SelectionStart = 0;
            _output.SelectionLength = 0;
            SetStatus($"{action} · 输出 {result.Text.Length} 字符");
        }
    }
}
