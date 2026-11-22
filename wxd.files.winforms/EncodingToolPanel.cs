using System;
using System.Drawing;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 编码页签：左边输入、右边输出，支持 Unicode 转义、UTF-8 Hex、
    /// Base64、URL 和 HTML 的编码/解码，以及 MD5 / SHA256 哈希，
    /// 文本统一按 UTF-8 处理。
    /// </summary>
    internal sealed class EncodingToolPanel : ToolPanelBase
    {
        private readonly TextBox _input = MakeEditor();
        private readonly TextBox _output = MakeEditor(readOnly: true);
        private readonly ThemeBordered _inputBorder;
        private readonly ThemeBordered _outputBorder;
        private readonly ThemeCheckBox _asciiOnly;
        private readonly ThemeCheckBox _upper;
        private readonly ThemeButton _swap = new()
        {
            Text = "⇄",
            Size = new Size(30, 30),
            // Segoe UI Symbol 保证双向箭头字形在所有皮肤字体下都不缺字
            Font = new Font("Segoe UI Symbol", 12f)
        };
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
            AddButton("MD5", (_, _) =>
                RunHash(text => HashToolEngine.Md5Hex(text, _upper.Checked), "MD5"));
            AddButton("SHA256", (_, _) =>
                RunHash(text => HashToolEngine.Sha256Hex(text, _upper.Checked), "SHA256"));
            _upper = AddCheckBox("大写输出");
            AddButton("文件 MD5", (_, _) => PickFileMd5());
            AddButton("文件夹 MD5", (_, _) => PickFolderMd5());
            AddButton("复制结果", (_, _) => CopyToClipboard(_output.Text));

            _swap.Click += (_, _) =>
            {
                string input = _input.Text;
                _input.Text = _output.Text;
                _output.Text = input;
                SetStatus("已交换左右两侧内容");
            };

            _inputBorder = new ThemeBordered(_input) { Dock = DockStyle.Fill };
            _outputBorder = new ThemeBordered(_output) { Dock = DockStyle.Fill };

            _split.Dock = DockStyle.Fill;
            _split.Orientation = Orientation.Vertical;
            _split.SplitterWidth = 6;
            _split.SizeChanged += (_, _) => ApplySplitterLayout();
            _split.SplitterMoved += (_, _) => PlaceSwapButton();
            _split.Panel1.Controls.Add(_inputBorder);
            _split.Panel2.Controls.Add(_outputBorder);
            // SplitContainer 自身的 Controls 是只读集合，按钮只能挂在面板上盖在两栏之上；
            // 先于 FinishChrome 添加，让它处于 z 序最前
            Controls.Add(_swap);
            var swapTip = new ToolTip();
            swapTip.SetToolTip(_swap, "左右交换");

            FinishChrome(_split);

            Load += (_, _) => ApplySplitterLayout();

            SetStatus("左边输入、右边输出；中间按钮可左右交换，MD5 / SHA256 按 UTF-8 取字节");
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
            finally
            {
                _applyingSplitter = false;
                PlaceSwapButton();
            }
        }

        private void PlaceSwapButton()
        {
            if (_split.IsDisposed) return;
            _swap.Location = new Point(
                _split.Left + _split.SplitterDistance + (_split.SplitterWidth - _swap.Width) / 2,
                _split.Top + (_split.Height - _swap.Height) / 2);
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
            _swap.FillOverride = Theme.P.Surface;
            _swap.Invalidate();
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

        // 哈希和编码不同：空文本也有摘要（UTF-8 0 字节），不做非空校验
        private void RunHash(Func<string, string> hash, string label)
        {
            string text = _input.Text;
            string result = hash(text);
            _output.Text = result;
            _output.SelectionStart = 0;
            _output.SelectionLength = 0;
            SetStatus(text.Length == 0
                ? $"空文本也有{label}（UTF-8 0 字节）"
                : $"{label} · 输出 {result.Length} 字符");
        }

        private void PickFileMd5()
        {
            using var dialog = new OpenFileDialog { Title = "选择文件计算 MD5" };
            if (dialog.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                string hex = HashToolEngine.FileMd5Hex(dialog.FileName, _upper.Checked);
                var info = new FileInfo(dialog.FileName);
                _output.Text = string.Join(Environment.NewLine,
                    "文件   : " + dialog.FileName,
                    "MD5    : " + hex,
                    $"大小   : {info.Length:N0} 字节");
                _output.SelectionStart = 0;
                _output.SelectionLength = 0;
                SetStatus("已计算文件 MD5");
            }
            catch (Exception ex)
            {
                SetStatus("读取失败：" + ex.Message, error: true);
            }
        }

        private bool _folderHashing;

        private async void PickFolderMd5()
        {
            using var dialog = new FolderBrowserDialog { Description = "选择文件夹计算 MD5" };
            if (dialog.ShowDialog(this) != DialogResult.OK) return;

            if (_folderHashing)
            {
                SetStatus("上一次文件夹计算还在进行，请稍等", error: true);
                return;
            }
            _folderHashing = true;
            string root = dialog.SelectedPath;
            bool upper = _upper.Checked;
            SetStatus("正在逐个文件计算 MD5…");
            try
            {
                // 后台线程逐文件流式计算，避免大文件夹卡住界面；无权限的文件跳过并计数
                var (lines, files, totalBytes, failed) = await Task.Run(() =>
                {
                    var list = new List<string>();
                    int count = 0;
                    long bytes = 0;
                    int errors = 0;
                    var options = new EnumerationOptions
                    {
                        RecurseSubdirectories = true,
                        IgnoreInaccessible = true
                    };
                    foreach (string path in Directory.EnumerateFiles(root, "*", options))
                    {
                        string rel = Path.GetRelativePath(root, path);
                        string hex;
                        try
                        {
                            hex = HashToolEngine.FileMd5Hex(path, upper);
                            bytes += new FileInfo(path).Length;
                        }
                        catch (Exception)
                        {
                            list.Add(rel + " : 读取失败");
                            errors++;
                            continue;
                        }
                        list.Add(rel + " : " + hex);
                        count++;
                    }
                    list.Sort(StringComparer.OrdinalIgnoreCase);
                    return (list, count, bytes, errors);
                });

                var all = new List<string>
                {
                    "文件夹 : " + root,
                    $"共 {files} 个文件，{totalBytes:N0} 字节"
                        + (failed > 0 ? $"，{failed} 个读取失败" : ""),
                    ""
                };
                all.AddRange(lines);
                _output.Text = string.Join(Environment.NewLine, all);
                _output.SelectionStart = 0;
                _output.SelectionLength = 0;
                SetStatus(files == 0
                    ? "文件夹里没有文件"
                    : $"文件夹 MD5 完成 · {files} 个文件" + (failed > 0 ? $"，{failed} 个失败" : ""));
            }
            catch (Exception ex)
            {
                SetStatus("文件夹读取失败：" + ex.Message, error: true);
            }
            finally
            {
                _folderHashing = false;
            }
        }
    }
}
