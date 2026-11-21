using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// JSON 页签：左边原始文本、右边可折叠的 JSON 树。
    /// 支持格式化、压缩成一行、递归 key 排序、全部展开/折叠、按 key/值搜索。
    /// 解析用 System.Text.Json，保留数字原始写法、中文不转义。
    /// </summary>
    internal sealed class JsonToolPanel : UserControl
    {
        private static readonly JsonSerializerOptions PreviewOptions = new()
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        private readonly FlowLayoutPanel _top = new();
        private readonly ThemeButton _btnOpen = new();
        private readonly ThemeButton _btnSave = new();
        private readonly ThemeButton _btnSaveAs = new();
        private readonly ThemeButton _btnFormat = new();
        private readonly ThemeButton _btnMinify = new();
        private readonly ThemeButton _btnSort = new();
        private readonly ThemeButton _btnExpand = new();
        private readonly ThemeButton _btnCollapse = new();
        private readonly ThemeButton _btnCopy = new();
        private readonly ThemeButton _btnPrev = new();
        private readonly ThemeButton _btnNext = new();
        private readonly ThemeField _fldSearch = new();
        private readonly TextBox _txtSource = new();
        private readonly ThemeBordered _sourceBorder;
        private readonly TreeView _tree = new();
        private readonly SplitContainer _split = new SafeSplitContainer();
        private readonly Label _lblStatus = new();
        private readonly System.Windows.Forms.Timer _parseTimer = new();
        private readonly List<TreeNode> _matches = new();

        private string _searchTerm = "";
        private int _matchIndex = -1;
        private string? _currentPath;
        private System.Text.Encoding? _currentEncoding;
        private bool _dirty;
        private bool _suppressTextChanged;
        private bool _hasComments;
        private JsonNode? _root;
        private string _tabCaption = "JSON";

        /// <summary>页签标题（文件名 / 未保存 *）变化，由工具集外壳刷新页签文字。</summary>
        internal event EventHandler? CaptionChanged;

        internal string TabCaption => _tabCaption;

        internal JsonToolPanel()
        {
            _sourceBorder = new ThemeBordered(_txtSource);
            BuildLayout();
            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) =>
            {
                Theme.Changed -= ApplyTheme;
                _parseTimer.Stop();
                _parseTimer.Dispose();
            };
        }

        // ==================== 布局 ====================

        private void BuildLayout()
        {
            SuspendLayout();

            Dock = DockStyle.Fill;
            BackColor = Theme.P.Window;
            Font = Theme.Body;
            AllowDrop = true;
            DragEnter += OnDragEnter;
            DragDrop += OnFileDrop;

            _top.Dock = DockStyle.Top;
            _top.Height = 46;
            _top.FlowDirection = FlowDirection.LeftToRight;
            _top.WrapContents = false;
            _top.AutoScroll = true;   // 窗口窄时按钮可横向滚动，避免被裁掉
            _top.Padding = new Padding(ThemeMetrics.Margin, 9, 0, 0);
            _top.BackColor = Color.Transparent;

            ConfigureButton(_btnOpen, "打开文件…", (_, _) => OpenFile());
            ConfigureButton(_btnSave, "保存", (_, _) => Save());
            ConfigureButton(_btnSaveAs, "另存为…", (_, _) => SaveAs());
            ConfigureButton(_btnFormat, "格式化", (_, _) => FormatJson());
            ConfigureButton(_btnMinify, "压缩", (_, _) => MinifyJson());
            ConfigureButton(_btnSort, "key 排序", (_, _) => SortKeys());
            ConfigureButton(_btnExpand, "全部展开", (_, _) => SetAllExpanded(true));
            ConfigureButton(_btnCollapse, "全部折叠", (_, _) => SetAllExpanded(false));
            ConfigureButton(_btnCopy, "复制结果", (_, _) => CopyResult());

            _btnPrev.Text = "上一个";
            _btnPrev.Font = Theme.Body;
            _btnPrev.Size = new Size(ThemeButton.FitWidth("上一个"), ThemeMetrics.RowH);
            _btnPrev.Margin = new Padding(ThemeMetrics.Gap, 0, 0, 0);
            _btnPrev.Click += (_, _) => RunSearch(forward: false);

            _btnNext.Text = "下一个";
            _btnNext.Font = Theme.Body;
            _btnNext.Size = new Size(ThemeButton.FitWidth("下一个"), ThemeMetrics.RowH);
            _btnNext.Margin = new Padding(ThemeMetrics.Gap, 0, 0, 0);
            _btnNext.Click += (_, _) => RunSearch(forward: true);

            _fldSearch.Size = new Size(210, ThemeMetrics.RowH);
            _fldSearch.Margin = new Padding(ThemeMetrics.Group, 0, ThemeMetrics.Gap, 0);
            _fldSearch.Placeholder = "搜索 key / 值";
            _fldSearch.Edit.KeyDown += OnSearchKeyDown;

            _top.Controls.AddRange(new Control[]
            {
                _btnOpen, _btnSave, _btnSaveAs, _btnFormat, _btnMinify, _btnSort,
                _btnExpand, _btnCollapse, _btnCopy, _fldSearch, _btnPrev, _btnNext
            });

            _txtSource.Multiline = true;
            _txtSource.ScrollBars = ScrollBars.Both;
            _txtSource.WordWrap = false;
            _txtSource.AcceptsTab = true;
            _txtSource.BorderStyle = BorderStyle.None;
            _txtSource.Font = Theme.Mono;
            _txtSource.TextChanged += OnSourceTextChanged;
            _txtSource.AllowDrop = true;
            _txtSource.DragEnter += OnDragEnter;
            _txtSource.DragDrop += OnFileDrop;

            _sourceBorder.Dock = DockStyle.Fill;

            _tree.Dock = DockStyle.Fill;
            _tree.BorderStyle = BorderStyle.None;
            _tree.HideSelection = false;
            _tree.ShowLines = true;
            _tree.ShowPlusMinus = true;
            _tree.ShowRootLines = true;
            _tree.Font = Theme.Mono;
            _tree.ItemHeight = 22;
            _tree.FullRowSelect = false;
            _tree.AllowDrop = true;
            _tree.DragEnter += OnDragEnter;
            _tree.DragDrop += OnFileDrop;

            _split.Dock = DockStyle.Fill;
            _split.Orientation = Orientation.Vertical;
            _split.SplitterWidth = 6;
            _split.SizeChanged += (_, _) => ApplySplitterLayout();

            _split.Panel1.Controls.Add(_sourceBorder);
            _split.Panel2.Controls.Add(_tree);

            _lblStatus.Dock = DockStyle.Bottom;
            _lblStatus.Height = 26;
            _lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            _lblStatus.Padding = new Padding(ThemeMetrics.Margin, 0, 0, 0);
            _lblStatus.Font = Theme.Caption;
            _lblStatus.Text = "打开 JSON 文件或直接粘贴内容";

            Controls.Add(_split);
            Controls.Add(_top);
            Controls.Add(_lblStatus);

            _parseTimer.Interval = 400;
            _parseTimer.Tick += (_, _) =>
            {
                _parseTimer.Stop();
                ParseAndRender();
            };

            Load += (_, _) =>
            {
                ApplySplitterLayout();
                _txtSource.Focus();
            };

            ResumeLayout(false);
            PerformLayout();
        }

        private void ConfigureButton(ThemeButton button, string text, EventHandler onClick)
        {
            button.Text = text;
            button.Font = Theme.Body;
            button.Margin = new Padding(0, 0, ThemeMetrics.Gap, 0);
            button.Size = new Size(ThemeButton.FitWidth(text), ThemeMetrics.RowH);
            button.Click += onClick;
        }

        private bool _applyingSplitter;

        private void ApplySplitterLayout()
        {
            // 面板创建初期宽度很小，固定 240 的 PanelNMinSize 会让 SplitContainer
            // 内部布局反复抛 InvalidOperationException，按当前宽度自适应并夹紧。
            if (_applyingSplitter || _split.IsDisposed) return;
            int width = _split.Width;
            if (width < 120) return;
            int min = Math.Min(240, Math.Max(25, (width - _split.SplitterWidth) / 3));
            int max = Math.Max(min, width - _split.SplitterWidth - min);
            int distance = Math.Clamp(width / 2, min, max);
            _applyingSplitter = true;
            try
            {
                if (_split.Panel1MinSize != min) _split.Panel1MinSize = min;
                if (_split.Panel2MinSize != min) _split.Panel2MinSize = min;
                if (_split.SplitterDistance != distance) _split.SplitterDistance = distance;
                NormalizePaneLayout();
            }
            catch (InvalidOperationException) { }
            finally { _applyingSplitter = false; }
        }

        private void NormalizePaneLayout()
        {
            _sourceBorder.Width = _split.Panel1.ClientSize.Width - 1;
            _tree.Width = _split.Panel2.ClientSize.Width;
        }

        // ==================== 主题 ====================

        private void ApplyTheme()
        {
            if (IsDisposed) return;
            var p = Theme.P;

            Font = Theme.Body;
            BackColor = p.Window;
            _top.BackColor = p.Chrome;

            foreach (var button in new[]
                     {
                         _btnOpen, _btnSave, _btnSaveAs, _btnFormat, _btnMinify, _btnSort,
                         _btnExpand, _btnCollapse, _btnCopy, _btnPrev, _btnNext
                     })
            {
                button.Font = Theme.Body;
                button.BackColor = p.Chrome;
                button.Invalidate();
            }

            _fldSearch.ApplyTheme();
            _sourceBorder.ApplyTheme();

            _txtSource.BackColor = p.Surface;
            _txtSource.ForeColor = p.Text;
            _txtSource.Font = Theme.Mono;

            _tree.BackColor = p.Surface;
            _tree.ForeColor = p.Text;
            _tree.Font = Theme.Mono;
            _tree.LineColor = p.Border;

            _split.BackColor = p.Divider;
            _split.Panel1.BackColor = p.Surface;
            _split.Panel2.BackColor = p.Surface;

            _lblStatus.BackColor = p.Window;
            _lblStatus.ForeColor = p.Muted;
            _lblStatus.Font = Theme.Caption;

            Invalidate(true);
        }

        // ==================== 打开 / 解析 ====================

        private void OpenFile()
        {
            using var dialog = new OpenFileDialog
            {
                Title = "打开 JSON 文件",
                CheckFileExists = true,
                Filter = "JSON 文件 (*.json;*.jsonc)|*.json;*.jsonc|所有文件 (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            LoadFile(dialog.FileName);
        }

        private bool LoadFile(string path)
        {
            if (_dirty
                && !string.Equals(path, _currentPath, StringComparison.OrdinalIgnoreCase)
                && !ConfirmDiscardChanges())
                return false;

            try
            {
                string text = TextFileReader.ReadAllText(path);
                var result = JsonToolEngine.Parse(text);

                _currentPath = path;
                _currentEncoding = TextFileReader.DetectEncoding(path);
                _hasComments = JsonToolEngine.ContainsComments(text);

                if (result.Success)
                {
                    // 含注释的 JSONC 不能自动格式化：JsonNode 序列化会丢掉注释。
                    // 其余文件先格式化方便看结构；打开不算"未保存"（磁盘文件没动）。
                    string display = _hasComments ? text : JsonToolEngine.Format(result.Root);
                    SetSourceText(display, suppressChanged: true);
                    _dirty = false;
                    UpdateCaption();
                    ParseAndRender(_hasComments
                        ? "已打开 " + Path.GetFileName(path) + " · 含注释，未自动格式化"
                        : "已打开 " + Path.GetFileName(path) + " · 已格式化");
                }
                else
                {
                    SetSourceText(text, suppressChanged: true);
                    _dirty = false;
                    UpdateCaption();
                    ParseAndRender();
                }

                UpdateCaption();
                return true;
            }
            catch (Exception ex)
            {
                SetStatus("读取失败：" + ex.Message, error: true);
                return false;
            }
        }

        private bool ConfirmDiscardChanges()
            => AlertForm.ShowConfirm(this,
                "当前文件有未保存的修改，打开新文件会丢失，是否继续？",
                "JSON 工具", AlertType.Warning) == DialogResult.Yes;

        /// <summary>保存到当前文件；没打开过文件时退化为"另存为"。</summary>
        private void Save()
        {
            if (string.IsNullOrEmpty(_currentPath))
            {
                SaveAs();
                return;
            }
            WriteToFile(_currentPath, _currentEncoding ?? new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        private void SaveAs()
        {
            using var dialog = new SaveFileDialog
            {
                Title = "保存 JSON",
                Filter = "JSON 文件 (*.json)|*.json|所有文件 (*.*)|*.*",
                FileName = _currentPath == null ? "untitled.json" : Path.GetFileName(_currentPath),
                InitialDirectory = _currentPath == null ? "" : Path.GetDirectoryName(_currentPath) ?? ""
            };

            if (dialog.ShowDialog(this) != DialogResult.OK) return;

            // 新文件统一 UTF-8 无 BOM
            WriteToFile(dialog.FileName, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        private void WriteToFile(string path, System.Text.Encoding encoding)
        {
            // 先写同目录临时文件再替换：写到一半崩溃不会把原文件截断
            if (!AtomicFileWriter.TryWriteAllText(path, _txtSource.Text, encoding, out string error))
            {
                SetStatus("保存失败（没有写权限或文件被占用）：" + error, error: true);
                return;
            }

            _currentPath = path;
            _currentEncoding = encoding;
            _dirty = false;
            UpdateCaption();
            SetStatus("已保存：" + path);
        }

        /// <summary>工具集外壳关闭前调用：有未保存修改时询问是否丢弃。</summary>
        internal bool ConfirmClose()
        {
            if (!_dirty) return true;

            return AlertForm.ShowConfirm(this,
                "JSON 页签有未保存的修改，关闭后丢失。确定关闭吗？", "JSON 工具", AlertType.Warning)
                == DialogResult.Yes;
        }

        private void OnSourceTextChanged(object? sender, EventArgs e)
        {
            if (_suppressTextChanged) return;
            MarkDirty();
            _parseTimer.Stop();
            _parseTimer.Start();
        }

        private void MarkDirty()
        {
            if (_dirty) return;
            _dirty = true;
            UpdateCaption();
        }

        private void UpdateCaption()
        {
            string suffix = _currentPath == null ? "" : " · " + Path.GetFileName(_currentPath);
            string caption = "JSON" + suffix + (_dirty ? " *" : "");
            if (caption == _tabCaption) return;

            _tabCaption = caption;
            CaptionChanged?.Invoke(this, EventArgs.Empty);
        }

        private void SetSourceText(string text, bool suppressChanged)
        {
            _suppressTextChanged = suppressChanged;
            try
            {
                _txtSource.Text = text;
                _txtSource.SelectionStart = 0;
                _txtSource.SelectionLength = 0;
            }
            finally
            {
                _suppressTextChanged = false;
            }
        }

        private void ParseAndRender(string? successMessage = null)
        {
            var result = JsonToolEngine.Parse(_txtSource.Text);
            _hasComments = JsonToolEngine.ContainsComments(_txtSource.Text);

            _matches.Clear();
            _searchTerm = "";
            _matchIndex = -1;

            if (!result.Success)
            {
                _root = null;
                _tree.Nodes.Clear();
                SetStatus(result.Error ?? "JSON 解析失败", error: true);
                return;
            }

            _root = result.Root;
            RebuildTree();

            int count = JsonToolEngine.CountNodes(result.Root);
            string message = successMessage == null
                ? $"解析成功 · {count} 个节点"
                : $"{successMessage} · {count} 个节点";
            if (result.NormalizedKeys)
                message += " · 已为数字 key 补引号";
            SetStatus(message);
        }

        private void RebuildTree()
        {
            _tree.BeginUpdate();
            try
            {
                _tree.Nodes.Clear();
                var rootNode = BuildTreeNode("$ " + Preview(_root), _root);
                _tree.Nodes.Add(rootNode);
                rootNode.Expand();
            }
            finally
            {
                _tree.EndUpdate();
            }
        }

        private TreeNode BuildTreeNode(string label, JsonNode? node)
        {
            var treeNode = new TreeNode(label);

            if (node is JsonObject obj)
            {
                foreach (var property in obj)
                {
                    string childLabel = QuoteKey(property.Key) + ": " + Preview(property.Value);
                    treeNode.Nodes.Add(BuildTreeNode(childLabel, property.Value));
                }
            }
            else if (node is JsonArray array)
            {
                for (int i = 0; i < array.Count; i++)
                {
                    string childLabel = "[" + i + "]: " + Preview(array[i]);
                    treeNode.Nodes.Add(BuildTreeNode(childLabel, array[i]));
                }
            }

            return treeNode;
        }

        private static string QuoteKey(string key)
            => JsonValue.Create(key)?.ToJsonString(PreviewOptions) ?? key;

        private static string Preview(JsonNode? node)
        {
            switch (node)
            {
                case null:
                    return "null";
                case JsonObject obj:
                    return "{ }" + (obj.Count > 0 ? $" ({obj.Count})" : "");
                case JsonArray array:
                    return "[ ]" + (array.Count > 0 ? $" ({array.Count})" : "");
                case JsonValue value:
                {
                    string text = value.ToJsonString(PreviewOptions);
                    return text.Length > 80 ? text[..77] + "…" : text;
                }
                default:
                    return node.ToJsonString();
            }
        }

        // ==================== 操作 ====================

        private void FormatJson() => ApplyTransform(root => JsonToolEngine.Format(root), "已格式化");

        private void MinifyJson() => ApplyTransform(root => JsonToolEngine.Minify(root), "已压缩为一行");

        private void SortKeys()
        {
            ApplyTransform(root =>
            {
                var sorted = JsonToolEngine.SortKeys(root);
                return JsonToolEngine.Format(sorted);
            }, "已按 key 排序（递归）");
        }

        private void ApplyTransform(Func<JsonNode?, string> transform, string successMessage)
        {
            var result = JsonToolEngine.Parse(_txtSource.Text);
            if (!result.Success)
            {
                SetStatus(result.Error ?? "JSON 解析失败", error: true);
                return;
            }

            if (_hasComments && AlertForm.ShowConfirm(this,
                    "当前 JSON 包含注释，此操作会移除注释，是否继续？",
                    "JSON 工具", AlertType.Warning) != DialogResult.Yes)
                return;

            string transformed = transform(result.Root);
            if (string.Equals(transformed, _txtSource.Text, StringComparison.Ordinal))
            {
                // 已经是目标格式：不标脏，避免"点一下格式化就被问保存"
                ParseAndRender(successMessage);
                return;
            }

            SetSourceText(transformed, suppressChanged: true);
            MarkDirty();
            ParseAndRender(successMessage);
        }

        private void CopyResult()
        {
            if (string.IsNullOrEmpty(_txtSource.Text))
            {
                SetStatus("没有可复制的内容", error: true);
                return;
            }

            try
            {
                Clipboard.SetText(_txtSource.Text);
                SetStatus($"已复制到剪贴板（{_txtSource.Text.Length} 字符）");
            }
            catch (Exception ex)
            {
                SetStatus("复制失败：" + ex.Message, error: true);
            }
        }

        private void SetAllExpanded(bool expanded)
        {
            _tree.BeginUpdate();
            try
            {
                SetExpanded(_tree.Nodes, expanded);
            }
            finally
            {
                _tree.EndUpdate();
            }
        }

        private static void SetExpanded(TreeNodeCollection nodes, bool expanded)
        {
            foreach (TreeNode node in nodes)
            {
                if (expanded) node.Expand();
                else node.Collapse();
                SetExpanded(node.Nodes, expanded);
            }
        }

        // ==================== 搜索 ====================

        private void OnSearchKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;
            e.Handled = true;
            e.SuppressKeyPress = true;
            RunSearch(forward: !e.Shift);
        }

        private void RunSearch(bool forward)
        {
            string term = _fldSearch.Edit.Text.Trim();
            if (term.Length == 0)
            {
                ClearSearch();
                return;
            }

            if (!string.Equals(term, _searchTerm, StringComparison.Ordinal))
            {
                _searchTerm = term;
                CollectMatches(term);
                _matchIndex = -1;
            }

            if (_matches.Count == 0)
            {
                SetStatus($"未找到“{term}”", error: true);
                return;
            }

            _matchIndex = forward
                ? (_matchIndex + 1) % _matches.Count
                : (_matchIndex - 1 + _matches.Count) % _matches.Count;

            var node = _matches[_matchIndex];
            ExpandAncestors(node);
            _tree.SelectedNode = node;
            node.EnsureVisible();

            SetStatus($"搜索“{term}”：第 {_matchIndex + 1}/{_matches.Count} 处");
        }

        private void CollectMatches(string term)
        {
            _matches.Clear();
            ResetNodeColors(_tree.Nodes);

            void Walk(TreeNodeCollection nodes)
            {
                foreach (TreeNode node in nodes)
                {
                    if (node.Text.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        _matches.Add(node);
                        node.ForeColor = Theme.P.Danger;
                    }
                    Walk(node.Nodes);
                }
            }

            Walk(_tree.Nodes);
        }

        private void ResetNodeColors(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                node.ForeColor = Theme.P.Text;
                ResetNodeColors(node.Nodes);
            }
        }

        private static void ExpandAncestors(TreeNode node)
        {
            for (var parent = node.Parent; parent != null; parent = parent.Parent)
                parent.Expand();
        }

        private void ClearSearch()
        {
            _searchTerm = "";
            _matchIndex = -1;
            _matches.Clear();
            ResetNodeColors(_tree.Nodes);
            SetStatus("搜索已清除");
        }

        // ==================== 杂项 ====================

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.O))
            {
                OpenFile();
                return true;
            }
            if (keyData == (Keys.Control | Keys.S))
            {
                Save();
                return true;
            }
            if (keyData == (Keys.Control | Keys.F))
            {
                _fldSearch.Edit.Focus();
                _fldSearch.Edit.SelectAll();
                return true;
            }
            if (keyData == Keys.F3 || keyData == (Keys.Shift | Keys.F3))
            {
                RunSearch(forward: keyData == Keys.F3);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private static void OnDragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effect = DragDropEffects.Copy;
        }

        private void OnFileDrop(object? sender, DragEventArgs e)
        {
            var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;
            LoadFile(files[0]);
        }

        private void SetStatus(string text, bool error = false)
        {
            _lblStatus.Text = text;
            _lblStatus.ForeColor = error ? Theme.P.Danger : Theme.P.Muted;
        }
    }
}
