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
    /// JSON 面板：左边原始文本、右边可折叠的 JSON 树。
    /// 支持格式化、压缩成一行、递归 key 排序、全部展开/折叠、按 key/值搜索。
    /// 解析用 System.Text.Json，保留数字原始写法、中文不转义。
    /// </summary>
    internal sealed class JsonToolForm : Form
    {
        private static readonly JsonSerializerOptions PreviewOptions = new()
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        private readonly FlowLayoutPanel _top = new();
        private readonly ThemeButton _btnOpen = new();
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
        private readonly SplitContainer _split = new();
        private readonly Label _lblStatus = new();
        private readonly System.Windows.Forms.Timer _parseTimer = new();
        private readonly List<TreeNode> _matches = new();

        private string _searchTerm = "";
        private int _matchIndex = -1;
        private bool _suppressTextChanged;
        private JsonNode? _root;

        internal JsonToolForm()
        {
            _sourceBorder = new ThemeBordered(_txtSource);
            InitializeComponent();
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

        private void InitializeComponent()
        {
            SuspendLayout();
            ClientSize = new Size(1280, 780);
            MinimumSize = new Size(900, 560);
            Text = "JSON 工具";
            StartPosition = FormStartPosition.CenterScreen;
            Font = Theme.Body;
            Icon = WindowIcons.Reader;
            KeyPreview = true;
            ResumeLayout(false);
        }

        // ==================== 布局 ====================

        private void BuildLayout()
        {
            SuspendLayout();

            _top.Dock = DockStyle.Top;
            _top.Height = 46;
            _top.FlowDirection = FlowDirection.LeftToRight;
            _top.WrapContents = false;
            _top.Padding = new Padding(ThemeMetrics.Margin, 9, 0, 0);
            _top.BackColor = Color.Transparent;

            ConfigureButton(_btnOpen, "打开文件…", (_, _) => OpenFile());
            ConfigureButton(_btnFormat, "格式化", (_, _) => FormatJson());
            ConfigureButton(_btnMinify, "压缩", (_, _) => MinifyJson());
            ConfigureButton(_btnSort, "key 排序", (_, _) => SortKeys());
            ConfigureButton(_btnExpand, "全部展开", (_, _) => SetAllExpanded(true));
            ConfigureButton(_btnCollapse, "全部折叠", (_, _) => SetAllExpanded(false));
            ConfigureButton(_btnCopy, "复制结果", (_, _) => CopyResult());

            _btnPrev.Text = "上一个";
            _btnPrev.Font = Theme.Body;
            _btnPrev.Margin = new Padding(ThemeMetrics.Gap, 0, 0, 0);
            _btnPrev.Click += (_, _) => RunSearch(forward: false);

            _btnNext.Text = "下一个";
            _btnNext.Font = Theme.Body;
            _btnNext.Margin = new Padding(ThemeMetrics.Gap, 0, 0, 0);
            _btnNext.Click += (_, _) => RunSearch(forward: true);

            _fldSearch.Size = new Size(210, ThemeMetrics.RowH);
            _fldSearch.Margin = new Padding(ThemeMetrics.Group, 0, ThemeMetrics.Gap, 0);
            _fldSearch.Placeholder = "搜索 key / 值";
            _fldSearch.Edit.KeyDown += OnSearchKeyDown;

            _top.Controls.AddRange(new Control[]
            {
                _btnOpen, _btnFormat, _btnMinify, _btnSort, _btnExpand, _btnCollapse, _btnCopy,
                _fldSearch, _btnPrev, _btnNext
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

            AllowDrop = true;
            DragEnter += OnDragEnter;
            DragDrop += OnFileDrop;

            Shown += (_, _) =>
            {
                try
                {
                    _split.SplitterDistance = Math.Max(320, _split.Width / 2);
                    _split.Panel1MinSize = 240;
                    _split.Panel2MinSize = 240;
                    NormalizePaneLayout();
                }
                catch (InvalidOperationException) { }
                _txtSource.Focus();
            };

            KeyDown += OnFormKeyDown;

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
                     { _btnOpen, _btnFormat, _btnMinify, _btnSort, _btnExpand, _btnCollapse, _btnCopy, _btnPrev, _btnNext })
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

        private void LoadFile(string path)
        {
            try
            {
                string text = TextFileReader.ReadAllText(path);
                var result = JsonToolEngine.Parse(text);

                if (result.Success)
                {
                    // 打开时先格式化一次，方便直接看结构
                    SetSourceText(JsonToolEngine.Format(result.Root), suppressChanged: true);
                    ParseAndRender("已打开 " + Path.GetFileName(path) + " · 已格式化");
                }
                else
                {
                    SetSourceText(text, suppressChanged: true);
                    ParseAndRender();
                }

                Text = "JSON 工具 - " + Path.GetFileName(path);
            }
            catch (Exception ex)
            {
                SetStatus("读取失败：" + ex.Message, error: true);
            }
        }

        private void OnSourceTextChanged(object? sender, EventArgs e)
        {
            if (_suppressTextChanged) return;
            _parseTimer.Stop();
            _parseTimer.Start();
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
            SetStatus(successMessage == null
                ? $"解析成功 · {count} 个节点"
                : $"{successMessage} · {count} 个节点");
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

            string transformed = transform(result.Root);
            SetSourceText(transformed, suppressChanged: true);
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

        private void OnFormKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.O)
            {
                OpenFile();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.F)
            {
                _fldSearch.Edit.Focus();
                _fldSearch.Edit.SelectAll();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F3)
            {
                RunSearch(forward: !e.Shift);
                e.Handled = true;
            }
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
