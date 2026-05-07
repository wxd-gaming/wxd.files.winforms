using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>大文件阅读器控件，可嵌入 TabPage</summary>
    public class LargeFileReaderControl : UserControl
    {
        private Button _btnOpenFile = null!, _btnGotoLine = null!;
        private TextBox _txtGotoLine = null!, _txtSearch = null!;
        private CheckBox _chkCaseSensitive = null!, _chkRegex = null!;
        private Button _btnSearchFirst = null!, _btnSearchPrev = null!, _btnSearchNext = null!, _btnSearchLast = null!, _btnExport = null!;
        private Label _lblStatus = null!, _lblIndexProgress = null!, _lblSearchCount = null!;
        private HScrollListView _lstContent = null!;
        private System.Windows.Forms.Timer _timer = null!, _timerScroll = null!;
        private TableLayoutPanel _topPanel = null!;

        // 搜索状态
        private List<int> _searchResults = new();
        private int _searchCurrentIndex = -1;
        private int _lastNavigatedMatchLine = -1;
        private CancellationTokenSource? _searchCts;
        private string _lastSearchKeyword = "";
        private bool _lastSearchCaseSensitive;
        private bool _lastSearchRegex;

        private string? _filePath;
        private long _fileLength;
        private volatile long[]? _lineOffsets;
        private volatile bool _indexReady;
        private volatile bool _cancel;
        private const int CHUNK = 4 * 1024 * 1024;
        private const int MAX_LINE = 64 * 1024;
        private volatile int _idxProg, _idxCount;

        private FileStream? _readStream;
        private readonly object _readLock = new();
        private readonly ConcurrentDictionary<int, string> _cache = new();
        private volatile int _cS = -1, _cE = -1;

        /// <summary>当前打开的文件的完整路径</summary>
        public string? FilePath => _filePath;

        public EventHandler? FileClosed;
        /// <summary>导出成功后触发</summary>
        public EventHandler? FileExported;

        public LargeFileReaderControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            _btnOpenFile = new Button { Text = "打开文件", Size = new Size(90, 28), Location = new Point(10, 6), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White };
            _btnOpenFile.Click += OnOpen;
            _btnGotoLine = new Button { Text = "跳转", Size = new Size(55, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(80, 80, 80), ForeColor = Color.White };
            _btnGotoLine.Click += OnGoto;
            _txtGotoLine = new TextBox { Size = new Size(120, 28), PlaceholderText = "行号" };
            _txtGotoLine.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OnGoto(s, e); };
            _lblIndexProgress = new Label { Text = "就绪", AutoSize = true, ForeColor = Color.Gray, Font = new Font("Microsoft YaHei UI", 8F) };
            _lblIndexProgress.Margin = new Padding(3, 7, 0, 0);
            _lblStatus = new Label { Text = "请打开一个文本文件", AutoSize = true, ForeColor = Color.FromArgb(95, 99, 104), Font = new Font("Microsoft YaHei UI", 8.5F) };
            _lblStatus.Margin = new Padding(3, 7, 0, 0);

            _topPanel = new TableLayoutPanel { Dock = DockStyle.Top, Height = 88, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(245, 246, 247) };
            _topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            _topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 1));
            _topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 43));

            var row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10, 8, 0, 0) };
            row1.Controls.AddRange(new Control[] { _btnOpenFile, _txtGotoLine, _btnGotoLine, _lblIndexProgress, _lblStatus });

            var separator = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(218, 220, 224) };

            _txtSearch = new TextBox { Size = new Size(200, 28), PlaceholderText = "搜索..." };
            _txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OnSearch(); };
            _txtSearch.TextChanged += (s, e) => { _lastSearchKeyword = ""; };
            _chkCaseSensitive = new CheckBox { Text = "Aa", AutoSize = true, Checked = false, TextAlign = ContentAlignment.MiddleRight };
            _chkCaseSensitive.Margin = new Padding(3, 6, 0, 0);
            _chkCaseSensitive.CheckedChanged += (s, e) => { _searchResults.Clear(); _searchCurrentIndex = -1; _lblSearchCount.Text = ""; _lastSearchKeyword = ""; };
            _chkRegex = new CheckBox { Text = ".*", AutoSize = true, Checked = false, TextAlign = ContentAlignment.MiddleRight };
            _chkRegex.Margin = new Padding(3, 6, 0, 0);
            _chkRegex.CheckedChanged += (s, e) => { _searchResults.Clear(); _searchCurrentIndex = -1; _lblSearchCount.Text = ""; _lastSearchKeyword = ""; };
            _btnSearchFirst = new Button { Text = "|<", Size = new Size(35, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(80, 80, 80), ForeColor = Color.White };
            _btnSearchFirst.Click += OnSearchFirst;
            _btnSearchPrev = new Button { Text = "上一个", Size = new Size(65, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(80, 80, 80), ForeColor = Color.White };
            _btnSearchPrev.Click += OnSearchPrev;
            _btnSearchNext = new Button { Text = "下一个", Size = new Size(65, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(80, 80, 80), ForeColor = Color.White };
            _btnSearchNext.Click += OnSearchNext;
            _btnSearchLast = new Button { Text = ">|", Size = new Size(35, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(80, 80, 80), ForeColor = Color.White };
            _btnSearchLast.Click += OnSearchLast;
            _lblSearchCount = new Label { Text = "", AutoSize = true, ForeColor = Color.Gray, Font = new Font("Microsoft YaHei UI", 8F), MinimumSize = new Size(180, 0) };
            _lblSearchCount.Margin = new Padding(3, 7, 0, 0);

            var row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10, 8, 0, 0) };
            _btnExport = new Button { Text = "导出结果", Size = new Size(75, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White };
            _btnExport.Click += OnExport;
            row2.Controls.AddRange(new Control[] { _txtSearch, _chkCaseSensitive, _chkRegex, _btnSearchFirst, _btnSearchPrev, _btnSearchNext, _btnSearchLast, _btnExport, _lblSearchCount });

            _topPanel.Controls.Add(row1, 0, 0);
            _topPanel.Controls.Add(separator, 0, 1);
            _topPanel.Controls.Add(row2, 0, 2);

            _lstContent = new HScrollListView();
            _lstContent.Dock = DockStyle.Fill;
            _lstContent.View = View.Details;
            _lstContent.FullRowSelect = true;
            _lstContent.HeaderStyle = ColumnHeaderStyle.None;
            _lstContent.VirtualMode = true;
            _lstContent.HideSelection = false;
            _lstContent.Scrollable = true;
            _lstContent.OwnerDraw = true;
            _lstContent.Font = new Font("Consolas", 10F);
            _lstContent.BackColor = Color.White;
            _lstContent.ForeColor = Color.FromArgb(30, 30, 30);
            _lstContent.Columns.Add("");
            _lstContent.RetrieveVirtualItem += OnRetrieve;
            _lstContent.CacheVirtualItems += OnCache;
            _lstContent.DrawItem += OnDrawItem;
            _lstContent.DrawSubItem += (s, e) => { };
            _lstContent.DrawColumnHeader += (s, e) => { };
            _lstContent.MouseDoubleClick += OnDoubleClick;
            _lstContent.DrawSubItem += (s, e) => { };
            _lstContent.DrawColumnHeader += (s, e) => { };
            Controls.Add(_lstContent);
            Controls.Add(_topPanel);
            ResumeLayout(false);

            _timer = new System.Windows.Forms.Timer { Interval = 200 };
            _timer.Tick += (_, _) => Tick();
            _timerScroll = new System.Windows.Forms.Timer { Interval = 300 };
            _timerScroll.Tick += (_, _) => UpdateScrollPos();
        }

        private void OnOpen(object? s, EventArgs e)
        {
            using var d = new OpenFileDialog { Filter = "文本文件|*.txt;*.log;*.csv;*.json;*.xml;*.cs;*.cpp;*.h;*.py;*.js|所有文件|*.*" };
            if (d.ShowDialog() != DialogResult.OK) return;
            Open(d.FileName);
        }

        public void Open(string path)
        {
            Cleanup();
            if (!File.Exists(path)) { MessageBox.Show(ParentForm, $"文件不存在: {path}", "错误"); return; }

            _filePath = path;
            _fileLength = new FileInfo(path).Length;
            _cancel = false; _indexReady = false; _idxProg = 0; _idxCount = 0;
            _lstContent.HScrollPos = 0;
            _lstContent.MaxContentWidth = _lstContent.ClientRectangle.Width;

            int loaded = 0;
            try
            {
                using var sr = new StreamReader(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.SequentialScan), Encoding.UTF8, true, 65536);
                for (int i = 0; i < 500; i++)
                {
                    string? line = sr.ReadLine();
                    if (line == null) break;
                    _cache[i] = line; loaded++;
                }
            }
            catch { }

            _lstContent.VirtualListSize = Math.Max(loaded, 1);
            _cS = 0; _cE = loaded - 1;

            _readStream?.Dispose();
            _readStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.RandomAccess);

            using (var g = _lstContent.CreateGraphics())
            {
                int maxW = _lstContent.ClientRectangle.Width;
                foreach (var kv in _cache)
                    if (!string.IsNullOrEmpty(kv.Value))
                    {
                        var sz = g.MeasureString(kv.Value, _lstContent.Font);
                        if (sz.Width + 10 > maxW) maxW = (int)sz.Width + 10;
                    }
                _lstContent.MaxContentWidth = maxW;
            }
            _lstContent.InitHScroll();

            Status($"前 {loaded} 行已加载，正在构建索引...");
            _lblIndexProgress.Text = "索引中...";
            try { _lstContent.RedrawItems(0, Math.Max(loaded - 1, 0), true); } catch { }

            _timer.Start();
            _ = Task.Run(() => BuildIndex());
        }

        // ==================== OwnerDraw 绘制 ====================
        private void OnDrawItem(object? s, DrawListViewItemEventArgs e)
        {
            e.DrawDefault = false;
            int hPos = _lstContent.HScrollPos;
            string? text = null;
            _cache.TryGetValue(e.ItemIndex, out text);

            if (string.IsNullOrEmpty(text))
            {
                e.Graphics.FillRectangle(new SolidBrush(_lstContent.BackColor), e.Bounds);
                return;
            }

            bool selected = _lstContent.SelectedIndices.Contains(e.ItemIndex);
            using var backBrush = new SolidBrush(selected ? Color.FromArgb(0x33, 0x77, 0xBB) : _lstContent.BackColor);
            using var textBrush = new SolidBrush(selected ? Color.White : _lstContent.ForeColor);

            e.Graphics.FillRectangle(backBrush, e.Bounds);

            int drawX = e.Bounds.X - hPos + 4;
            int drawY = e.Bounds.Y + (e.Bounds.Height - _lstContent.Font.Height) / 2;
            TextRenderer.DrawText(e.Graphics, text, _lstContent.Font,
                new Rectangle(drawX, drawY, int.MaxValue, e.Bounds.Height),
                selected ? Color.White : _lstContent.ForeColor,
                TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);

            if ((e.State & ListViewItemStates.Focused) != 0)
                e.DrawFocusRectangle();
        }

        // ==================== 索引构建 ====================
        private void BuildIndex()
        {
            try
            {
                var est = Math.Max(_fileLength / 40L, 1024L);
                if (est > int.MaxValue) est = int.MaxValue;
                var off = new List<long>((int)est) { 0 };

                using var fs = new FileStream(_filePath!, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, CHUNK, FileOptions.SequentialScan);
                var buf = new byte[CHUNK];
                long pos = 0; int n;
                while ((n = fs.Read(buf, 0, buf.Length)) > 0)
                {
                    if (_cancel) return;
                    for (int i = 0; i < n; i++, pos++)
                        if (buf[i] == '\n') off.Add(pos + 1);
                    _idxCount = off.Count - 1;
                    _idxProg = (int)(pos * 100 / _fileLength);
                }
                if (off.Count > 1 && off[off.Count - 1] >= _fileLength) off.RemoveAt(off.Count - 1);
                _lineOffsets = off.ToArray();
                _idxCount = _lineOffsets.Length;
                _idxProg = 100; _indexReady = true;
            }
            catch (Exception ex) { Debug.WriteLine($"Index: {ex}"); _indexReady = false; }
        }

        private void Tick()
        {
            if (_indexReady)
            {
                _timer.Stop();
                _lblIndexProgress.Text = $"完成 {_idxCount:N0} 行";
                Status($"{FormatSize(_fileLength)} | {_idxCount:N0} 行");
                _lstContent.VirtualListSize = _idxCount;
                _timerScroll.Start();
                try { _lstContent.RedrawItems(0, Math.Min(_lstContent.VirtualListSize - 1, _lstContent.ClientRectangle.Height / Math.Max(_lstContent.Font.Height + 6, 1)), true); } catch { }
                return;
            }
            _lblIndexProgress.Text = $"索引中... {_idxCount:N0} 行 ({_idxProg}%)";
        }

        private void UpdateScrollPos()
        {
            try
            {
                int line = (_lstContent.TopItem?.Index ?? 0) + 1;
                _lblIndexProgress.Text = $"行 {line:N0}/{_idxCount:N0}";
            }
            catch { }
        }

        private string ReadLine(int idx)
        {
            var off = _lineOffsets;
            if (off == null || _readStream == null || idx >= off.Length) return "";
            long s = off[idx], e = (idx + 1 < off.Length) ? off[idx + 1] : _fileLength;
            int len = Math.Min(Math.Max((int)(e - s), 0), MAX_LINE);
            if (len <= 0) return "";
            lock (_readLock)
            {
                try
                {
                    _readStream.Seek(s, SeekOrigin.Begin);
                    var buf = new byte[len]; int got = 0;
                    while (got < len) { int r = _readStream.Read(buf, got, len - got); if (r <= 0) break; got += r; }
                    return got > 0 ? Encoding.UTF8.GetString(buf, 0, got).TrimEnd('\r', '\n') : "";
                }
                catch { return "[err]"; }
            }
        }

        private void OnRetrieve(object? s, RetrieveVirtualItemEventArgs e)
        {
            e.Item = new ListViewItem(_filePath != null && _cache.TryGetValue(e.ItemIndex, out var v) ? v : "");
        }

        private void OnCache(object? s, CacheVirtualItemsEventArgs e)
        {
            if (_filePath == null || !_indexReady) return;
            int start = Math.Max(0, e.StartIndex), end = Math.Min(_lstContent.VirtualListSize - 1, e.EndIndex);
            if (start > end) return;
            if (_cS >= 0 && start >= _cS && end <= _cE) return;
            for (int i = _cS; i <= _cE; i++) if (i < start || i > end) _cache.TryRemove(i, out _);
            for (int i = start; i <= end; i++) if (!_cache.ContainsKey(i)) _cache[i] = ReadLine(i);
            _cS = start; _cE = end;
        }

        private void Status(string s) { if (_lblStatus != null) try { BeginInvoke(() => _lblStatus.Text = s); } catch { } }
        static string FormatSize(long b) { string[] u = ["B", "KB", "MB", "GB", "TB"]; double v = b; int o = 0; while (v >= 1024 && o < 4) { o++; v /= 1024; } return $"{v:0.##}{u[o]}"; }

        private void OnDoubleClick(object? s, MouseEventArgs e)
        {
            int idx = _lstContent.SelectedIndices.Count > 0
                ? _lstContent.SelectedIndices[0]
                : (e.Y / (_lstContent.Font.Height + 6) + (_lstContent.TopItem?.Index ?? 0));
            if (idx < 0 || idx >= _lstContent.VirtualListSize) return;
            _lstContent.SelectedIndices.Clear();
            _lstContent.SelectedIndices.Add(idx);

            string originalText = _cache.TryGetValue(idx, out var v) ? v : "";
            if (string.IsNullOrEmpty(originalText)) return;

            string formattedText = BuildFormattedText(originalText);

            var f = new Form
            {
                Text = $"行 {idx + 1}", Size = new Size(950, 550), StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false, MaximizeBox = false, ShowIcon = false, ShowInTaskbar = false
            };

            // 顶部工具栏
            var topPanel = new Panel { Height = 44, Dock = DockStyle.Top, BackColor = Color.FromArgb(245, 246, 247) };
            var btnFormat = new Button { Text = "格式化", Size = new Size(90, 30), Location = new Point(10, 7), FlatStyle = FlatStyle.Flat };
            var btnRaw = new Button { Text = "原始", Size = new Size(90, 30), Location = new Point(110, 7), FlatStyle = FlatStyle.Flat };

            void SetActive(Button active, Button inactive)
            {
                active.BackColor = Color.FromArgb(0, 122, 204);
                active.ForeColor = Color.White;
                inactive.BackColor = Color.FromArgb(210, 210, 210);
                inactive.ForeColor = Color.FromArgb(80, 80, 80);
            }
            SetActive(btnFormat, btnRaw);

            var rtb = new RichTextBox
            {
                Multiline = true, ReadOnly = true, Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F), ScrollBars = RichTextBoxScrollBars.Both,
                WordWrap = false, BackColor = Color.White
            };

            // 默认显示格式化
            ShowFormattedRtf(rtb, formattedText);

            btnFormat.Click += (_, _) => { SetActive(btnFormat, btnRaw); ShowFormattedRtf(rtb, formattedText); };
            btnRaw.Click += (_, _) => { SetActive(btnRaw, btnFormat); rtb.Text = originalText; };

            topPanel.Controls.Add(btnFormat);
            topPanel.Controls.Add(btnRaw);
            f.Controls.Add(rtb);
            f.Controls.Add(topPanel);
            f.ShowDialog(ParentForm);
        }

        private static void ShowFormattedRtf(RichTextBox rtb, string formattedText)
        {
            try { rtb.Rtf = BuildRtfForLine(formattedText); }
            catch { rtb.Text = formattedText; }
        }

        /// <summary>构建格式化文本：JSON 缩进换行，其余内容按逗号换行</summary>
        private static string BuildFormattedText(string line)
        {
            var sb = new StringBuilder();
            int i = 0;
            while (i < line.Length)
            {
                // JSON 对象 { ... }
                if (line[i] == '{')
                {
                    int end = FindJsonEnd(line, i, '{', '}');
                    if (end > 0)
                    {
                        string raw = line.Substring(i, end - i + 1);
                        string? beautified = TryBeautifyJson(raw);
                        if (beautified != null)
                        {
                            if (sb.Length > 0 && sb[^1] != '\n') sb.AppendLine();
                            sb.Append(beautified);
                            i = end + 1;
                            continue;
                        }
                        // 尝试将 key=value 格式转为 JSON
                        string? kvJson = TryConvertEqualsToJson(raw);
                        if (kvJson != null)
                        {
                            if (sb.Length > 0 && sb[^1] != '\n') sb.AppendLine();
                            sb.Append(kvJson);
                            i = end + 1;
                            continue;
                        }
                        // 外层 JSON 解析失败 → 剥除外层，逐条目拆分
                        // 例如 {1:{...}, 2:{...}} → 1:{...} 每条单独处理
                        string inner = raw.Substring(1, raw.Length - 2).Trim();
                        var entries = SplitTopLevelEntries(inner);
                        foreach (var entry in entries)
                        {
                            if (string.IsNullOrEmpty(entry)) continue;
                            string trimmedEntry = entry.Trim();
                            if (sb.Length > 0) sb.AppendLine();
                            // 尝试提取内部 JSON 块
                            int jsonStart = trimmedEntry.IndexOf('{');
                            int jsonStart2 = trimmedEntry.IndexOf('[');
                            int jsonStartPos = -1;
                            if (jsonStart >= 0 && jsonStart2 >= 0) jsonStartPos = Math.Min(jsonStart, jsonStart2);
                            else if (jsonStart >= 0) jsonStartPos = jsonStart;
                            else if (jsonStart2 >= 0) jsonStartPos = jsonStart2;

                            if (jsonStartPos > 0)
                            {
                                // key 部分
                                string keyPart = trimmedEntry[..jsonStartPos].Trim().TrimEnd(':');
                                if (!string.IsNullOrEmpty(keyPart))
                                {
                                    sb.Append(keyPart);
                                }
                                // 子 JSON 块
                                char brace = trimmedEntry[jsonStartPos];
                                int subEnd = FindJsonEnd(trimmedEntry, jsonStartPos, brace, brace == '{' ? '}' : ']');
                                if (subEnd > jsonStartPos)
                                {
                                    string subRaw = trimmedEntry.Substring(jsonStartPos, subEnd - jsonStartPos + 1);
                                    string? subBeautified = TryBeautifyJson(subRaw);
                                    if (subBeautified != null)
                                    {
                                        sb.AppendLine();
                                        sb.Append(IndentBlock(subBeautified, "  "));
                                    }
                                    else
                                    {
                                        sb.AppendLine();
                                        sb.Append(subRaw);
                                    }
                                }
                            }
                            else
                            {
                                // 整个条目没有 JSON → 按逗号再拆分
                                sb.Append(trimmedEntry);
                            }
                        }
                        i = end + 1;
                        continue;
                    }
                }
                // JSON 数组 [ ... ]
                if (line[i] == '[')
                {
                    int end = FindJsonEnd(line, i, '[', ']');
                    if (end > 0)
                    {
                        string raw = line.Substring(i, end - i + 1);
                        string? beautified = TryBeautifyJson(raw);
                        if (beautified != null)
                        {
                            if (sb.Length > 0 && sb[^1] != '\n') sb.AppendLine();
                            sb.Append(beautified);
                            i = end + 1;
                            continue;
                        }
                        // 尝试将数组元素从 key=value 格式转为 JSON
                        string? kvArrayJson = TryConvertArrayToJson(raw);
                        if (kvArrayJson != null)
                        {
                            if (sb.Length > 0 && sb[^1] != '\n') sb.AppendLine();
                            sb.Append(kvArrayJson);
                            i = end + 1;
                            continue;
                        }
                        if (sb.Length > 0 && sb[^1] != '\n') sb.AppendLine();
                        sb.Append(raw);
                        i = end + 1;
                        continue;
                    }
                }

                // 非 JSON 内容，累积到下一个 JSON 起始或结尾
                int start = i;
                while (i < line.Length && line[i] is not '{' and not '[') i++;
                if (i > start)
                {
                    string chunk = line.Substring(start, i - start).Trim();
                    if (!string.IsNullOrEmpty(chunk))
                    {
                        chunk = chunk.Trim().TrimStart(',').TrimEnd(',');
                        var parts = chunk.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                        foreach (var part in parts)
                        {
                            if (!string.IsNullOrEmpty(part))
                            {
                                if (sb.Length > 0) sb.AppendLine();
                                sb.Append(part.Trim());
                            }
                        }
                    }
                }
            }
            return sb.ToString();
        }

        /// <summary>在逗号分隔的字符串中按顶层（忽略嵌套 {}[]）拆分成条目</summary>
        private static List<string> SplitTopLevelEntries(string s)
        {
            var result = new List<string>();
            int depth = 0;
            int start = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c is '{' or '[') depth++;
                else if (c is '}' or ']') depth--;
                else if (c == ',' && depth == 0)
                {
                    result.Add(s.Substring(start, i - start));
                    start = i + 1;
                }
            }
            if (start < s.Length)
                result.Add(s.Substring(start));
            return result;
        }

        /// <summary>给多行文本每行添加缩进</summary>
        private static string IndentBlock(string text, string indent)
        {
            var lines = text.Split('\n');
            return string.Join("\n", lines.Select(l => indent + l));
        }

        /// <summary>尝试美化 JSON，失败先修复未引号 key 再试，还失败返回 null</summary>
        private static string? TryBeautifyJson(string json)
        {
            // 第一次尝试原样解析
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                return System.Text.Json.JsonSerializer.Serialize(doc.RootElement,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true, IndentSize = 2 });
            }
            catch { }

            // 修复未引号的 key 后再试
            string fixedJson = FixUnquotedKeys(json);
            if (fixedJson != json)
            {
                try
                {
                    using var doc = System.Text.Json.JsonDocument.Parse(fixedJson);
                    return System.Text.Json.JsonSerializer.Serialize(doc.RootElement,
                        new System.Text.Json.JsonSerializerOptions { WriteIndented = true, IndentSize = 2 });
                }
                catch { }
            }
            return null;
        }

        /// <summary>将 JSON 中未加引号的 key（如 {1:..., key:...}）加上引号</summary>
        private static string FixUnquotedKeys(string json)
        {
            var sb = new StringBuilder(json.Length + 32);
            bool inStr = false;
            char strChar = '\0';

            for (int i = 0; i < json.Length; i++)
            {
                char c = json[i];

                // 字符串内：原样复制
                if (inStr)
                {
                    if (c == '\\') { sb.Append(c); i++; if (i < json.Length) sb.Append(json[i]); continue; }
                    sb.Append(c);
                    if (c == strChar) inStr = false;
                    continue;
                }

                if (c is '"' or '\'') { inStr = true; strChar = c; sb.Append(c); continue; }

                // 遇到 { 或 , 时，检查后面是否跟着一个未引号的 key (word/digits : )
                if (c is '{' or ',')
                {
                    sb.Append(c);
                    int j = i + 1;
                    while (j < json.Length && char.IsWhiteSpace(json[j])) j++;
                    if (j < json.Length && (char.IsLetterOrDigit(json[j]) || json[j] == '_'))
                    {
                        int k = j;
                        while (k < json.Length && (char.IsLetterOrDigit(json[k]) || json[k] is '_' or '.' or '#' or '@' or '-')) k++;
                        int l = k;
                        while (l < json.Length && char.IsWhiteSpace(json[l])) l++;
                        if (l < json.Length && json[l] == ':')
                        {
                            // 这是未引号的 key → 加双引号
                            sb.Append('"');
                            sb.Append(json, j, k - j);
                            sb.Append('"');
                            i = k - 1; // 循环会自增 i
                            continue;
                        }
                    }
                    continue;
                }

                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>将 key=value 格式的 { id=123, name=abc } 转换为标准 JSON</summary>
        private static string? TryConvertEqualsToJson(string text)
        {
            if (text.Length < 2 || text[0] != '{' || text[^1] != '}') return null;
            string inner = text.Substring(1, text.Length - 2).Trim();
            if (string.IsNullOrEmpty(inner)) return null;
            // 如果包含标准 JSON 的引号格式（"key": val 或 'key': val），交由 FixUnquotedKeys 处理
            if (inner.Contains("\":")) return null;
            // 检查是否大部分条目是 key=value 格式
            int eqCount = 0, colonCount = 0, totalEntries = 0;
            var entries = SplitTopLevelEntries(inner);
            foreach (var entry in entries)
            {
                var t = entry.Trim();
                if (string.IsNullOrEmpty(t)) continue;
                totalEntries++;
                if (t.Contains('=')) eqCount++;
                if (t.Contains(':')) colonCount++;
            }
            if (totalEntries == 0) return null;
            // 超过一半是 = 格式才转换
            if (eqCount < totalEntries / 2.0) return null;

            var jsonSb = new StringBuilder();
            jsonSb.AppendLine("{");
            for (int idx = 0; idx < entries.Count; idx++)
            {
                var entry = entries[idx].Trim();
                if (string.IsNullOrEmpty(entry)) continue;
                int eqPos = entry.IndexOf('=');
                if (eqPos > 0)
                {
                    string key = entry[..eqPos].Trim();
                    string val = entry[(eqPos + 1)..].Trim();
                    jsonSb.Append($"  \"{EscapeJsonString(key)}\": ");
                    AppendJsonValue(jsonSb, val);
                }
                else
                {
                    jsonSb.Append("  ");
                    AppendJsonValue(jsonSb, entry);
                }
                jsonSb.AppendLine(idx < entries.Count - 1 ? "," : "");
            }
            jsonSb.Append("}");
            return jsonSb.ToString();
        }

        /// <summary>尝试将数组内容从 key=value 对象格式转为 JSON 数组</summary>
        private static string? TryConvertArrayToJson(string text)
        {
            if (text.Length < 2 || text[0] != '[' || text[^1] != ']') return null;
            string inner = text.Substring(1, text.Length - 2).Trim();
            if (string.IsNullOrEmpty(inner)) return null;
            // 已包含标准 JSON 结构则跳过
            if (inner.TrimStart().StartsWith("{")) return null;
            // 按顶层逗号拆分条目
            var items = SplitTopLevelEntries(inner);
            if (items.Count == 0) return null;
            // 检查是否每个条目都包含 {...}
            bool allHaveBrace = items.All(i => i.Contains('{'));
            if (!allHaveBrace) return null;

            var jsonSb = new StringBuilder();
            var convertedItems = new List<string>();
            foreach (string item in items)
            {
                string trimmed = item.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;
                int braceStart = trimmed.IndexOf('{');
                if (braceStart < 0) { convertedItems.Add($"  {trimmed}"); continue; }
                string prefix = braceStart > 0 ? trimmed[..braceStart].Trim() : "";
                string braceContent = trimmed[braceStart..].Trim();
                string? objJson = TryBeautifyJson(braceContent);
                if (objJson == null)
                    objJson = TryConvertEqualsToJson(braceContent);
                if (objJson != null)
                {
                    // 合并 prefix 和 JSON 对象
                    if (!string.IsNullOrEmpty(prefix))
                    {
                        var lines = objJson.Split('\n');
                        for (int li = 0; li < lines.Length; li++)
                        {
                            if (li == 0)
                                objJson = $"  {prefix} {lines[li]}";
                            else
                                objJson += $"\n  {lines[li]}";
                        }
                    }
                    else
                    {
                        objJson = "  " + objJson.Replace("\n", "\n  ");
                    }
                    convertedItems.Add(objJson);
                }
                else
                {
                    convertedItems.Add($"  {trimmed}");
                }
            }
            jsonSb.AppendLine("[");
            jsonSb.Append(string.Join(",\n", convertedItems));
            jsonSb.Append("\n]");
            return jsonSb.ToString();
        }

        /// <summary>转义 JSON 字符串中的特殊字符</summary>
        private static string EscapeJsonString(string s)
        {
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '\"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default: sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        /// <summary>将值智能处理：数字不加引号，其他加引号</summary>
        private static void AppendJsonValue(StringBuilder sb, string val)
        {
            if (string.IsNullOrEmpty(val)) { sb.Append("\"\""); return; }
            // true/false/null
            if (val == "true" || val == "false" || val == "null") { sb.Append(val); return; }
            // 纯数字
            if (long.TryParse(val, out _) || double.TryParse(val, out _)) { sb.Append(val); return; }
            // 有 { 或 [ 的内容原样输出
            if (val.Contains('{') || val.Contains('[')) { sb.Append(val); return; }
            sb.Append('\"');
            sb.Append(EscapeJsonString(val));
            sb.Append('\"');
        }

        // ==================== 混合内容格式化引擎 ====================

        /// <summary>片段类型</summary>
        private enum SegType { JsonBlock, KeyValuePair, QuotedString, Number, BooleanNull, PlainText }

        /// <summary>将已格式化文本（多行）渲染为带语法高亮的 RTF</summary>
        private static string BuildRtfForLine(string formattedText)
        {
            var lines = formattedText.Split('\n');
            var sb = new StringBuilder();
            sb.AppendLine(@"{\rtf1\ansi\deff0");
            sb.AppendLine(@"{\fonttbl{\f0\fnil Consolas;}}");
            sb.AppendLine(@"{\colortbl ;\red0\green122\blue204;\red22\green138\blue8;\red156\green39\blue176;\red117\green117\blue117;\red216\green67\blue21;\red0\green121\blue203;\red30\green30\blue30;}");
            sb.Append(@"\f0\fs20 ");

            bool inJson = false;
            foreach (var rawLine in lines)
            {
                sb.Append(@"\par ");
                string rl = rawLine.TrimEnd('\r');
                if (rl.Length == 0) { sb.Append(' '); continue; }

                string trimmed = rl.TrimStart();
                // 判断 JSON 块的开始/结束
                if (!inJson && trimmed.Length > 0 && (trimmed[0] is '{' or '['))
                    inJson = true;

                if (inJson)
                {
                    // JSON 块整体蓝色
                    sb.Append(@"\cf1 ");
                    sb.Append(EscapeRtf(rl));
                    if (trimmed.Length > 0 && (trimmed[^1] is '}' or ']'))
                        inJson = false;
                }
                else
                {
                    // 非 JSON 行 → 逐 token 语法着色
                    var tokens = TokenizeLine(rl);
                    foreach (var (type, text) in tokens)
                    {
                        string escaped = EscapeRtf(text);
                        switch (type)
                        {
                            case SegType.JsonBlock:
                                sb.Append(@"\cf1 ").Append(escaped.Replace("\n", @"\par "));
                                break;
                            case SegType.KeyValuePair:
                                int eqIdx = text.IndexOf('=');
                                if (eqIdx > 0)
                                {
                                    sb.Append(@"\cf5 ").Append(EscapeRtf(text[..eqIdx].TrimEnd()));
                                    sb.Append(@"\cf7 = ");
                                    sb.Append(@"\cf6 ").Append(EscapeRtf(text[(eqIdx + 1)..].TrimStart()));
                                }
                                else { sb.Append(@"\cf5 ").Append(escaped); }
                                break;
                            case SegType.QuotedString:
                                sb.Append(@"\cf2 ").Append(escaped);
                                break;
                            case SegType.Number:
                                sb.Append(@"\cf3 ").Append(escaped);
                                break;
                            case SegType.BooleanNull:
                                sb.Append(@"\b\cf4 ").Append(escaped).Append(@"\b0");
                                break;
                            default:
                                sb.Append(@"\cf7 ").Append(escaped);
                                break;
                        }
                    }
                }
            }
            sb.AppendLine("}");
            return sb.ToString();
        }

        /// <summary>RTF 转义</summary>
        private static string EscapeRtf(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            // 处理特殊 RTF 字符和 Unicode
            var sb = new StringBuilder(text.Length);
            foreach (char c in text)
            {
                switch (c)
                {
                    case '\\': sb.Append(@"\\"); break;
                    case '{':  sb.Append(@"\{ "); break;
                    case '}':  sb.Append(@"\} "); break;
                    case '\n': sb.Append(@"\par "); break;
                    case '\r': break; // skip
                    default:
                        if (c > 127)
                        {
                            // Unicode char: \uNNNN?
                            sb.Append($@"\u{(int)c}?");
                        }
                        else if (c < 32)
                        {
                            // Control char
                            sb.Append($@"\'{((int)c):X2}");
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            return sb.ToString();
        }

        /// <summary>将一行混合文本分词为带类型的片段列表</summary>
        private static List<(SegType Type, string Text)> TokenizeLine(string line)
        {
            var result = new List<(SegType, string)>();
            int i = 0;
            while (i < line.Length)
            {
                // 1) JSON 对象 { ... }
                if (line[i] == '{')
                {
                    int end = FindJsonEnd(line, i, '{', '}');
                    if (end > 0)
                    {
                        string raw = line.Substring(i, end - i + 1);
                        string beautified = BeautifyJson(raw);
                        result.Add((SegType.JsonBlock, beautified));
                        i = end + 1;
                        continue;
                    }
                }

                // 2) JSON 数组 [ ... ]
                if (line[i] == '[')
                {
                    int end = FindJsonEnd(line, i, '[', ']');
                    if (end > 0)
                    {
                        string raw = line.Substring(i, end - i + 1);
                        string beautified = BeautifyJson(raw);
                        result.Add((SegType.JsonBlock, beautified));
                        i = end + 1;
                        continue;
                    }
                }

                // 3) 键值对 key=value
                var rest = line.AsSpan(i);
                var kvM = System.Text.RegularExpressions.Regex.Match(rest.ToString(), @"^([\w.@\-]+)\s*=\s*([^,\s;\]\}\)]+)");
                if (kvM.Success && kvM.Groups[1].Value.Length > 0)
                {
                    result.Add((SegType.KeyValuePair, kvM.Value));
                    i += kvM.Length;
                    continue;
                }

                // 4) 引号字符串
                if (line[i] is '"' or '\'')
                {
                    char quote = line[i];
                    int end = i + 1;
                    while (end < line.Length && line[end] != quote)
                    {
                        if (line[end] == '\\') end++; // 跳过转义
                        end++;
                    }
                    if (end < line.Length) end++; // 闭合引号
                    result.Add((SegType.QuotedString, line.Substring(i, end - i)));
                    i = end;
                    continue;
                }

                // 5) 数字
                int numStart = i;
                if (line[i] == '-' && i + 1 < line.Length && char.IsDigit(line[i + 1])) i++;
                if (i < line.Length && char.IsDigit(line[i]))
                {
                    while (i < line.Length && (char.IsDigit(line[i]) || line[i] is '.' or 'e' or 'E' or '+' or '-' || line[i] == 'f' || line[i] == 'd' || line[i] == 'm'))
                    {
                        if ((line[i] is '+' or '-') && i > numStart && line[i - 1] is not 'e' and not 'E')
                        {
                            // +/- not preceded by scientific notation → end of number
                            break;
                        }
                        i++;
                    }
                    // 回退: 纯数字不应该包含字母后缀（除了 e/E/f/d/m）
                    while (i > numStart + 1 && !char.IsDigit(line[i - 1]) && line[i - 1] is not '.' and not 'f' and not 'd' and not 'm')
                        i--;
                    result.Add((SegType.Number, line.Substring(numStart, i - numStart)));
                    continue;
                }
                i = numStart; // 不是数字，回退

                // 6) true / false / null
                if (i + 3 <= line.Length)
                {
                    var sub = line.AsSpan(i);
                    if (sub.StartsWith("true") || sub.StartsWith("null") ||
                        (sub.Length >= 5 && sub.StartsWith("false")))
                    {
                        int kwLen = sub.StartsWith("false") ? 5 : 4;
                        result.Add((SegType.BooleanNull, line.Substring(i, kwLen)));
                        i += kwLen;
                        continue;
                    }
                }

                // 7) 普通文本 — 累积到遇到特殊字符为止
                int ps = i;
                while (i < line.Length)
                {
                    char c = line[i];
                    if (c is '{' or '[' or '"' or '\'') break;
                    // 检查是否是 KV 起始: word 后跟 =
                    if (char.IsLetterOrDigit(c) || c is '-' or '.' or '@' or '_')
                    {
                        int j = i;
                        while (j < line.Length && (char.IsLetterOrDigit(line[j]) || line[j] is '.' or '-' or '@' or '_')) j++;
                        if (j < line.Length && line[j] == '=')
                        {
                            // 后面有 =，让 KV 匹配器处理
                            break;
                        }
                    }
                    // 检查是否是数字/关键词起始
                    if (char.IsDigit(c)) break;
                    if (c == '-' && i + 1 < line.Length && char.IsDigit(line[i + 1])) break;
                    if ((c == 't' && i + 3 < line.Length && line.AsSpan(i, 4).SequenceEqual("true")) ||
                        (c == 'f' && i + 4 < line.Length && line.AsSpan(i, 5).SequenceEqual("false")) ||
                        (c == 'n' && i + 3 < line.Length && line.AsSpan(i, 4).SequenceEqual("null")))
                        break;
                    i++;
                }
                if (i > ps)
                    result.Add((SegType.PlainText, line.Substring(ps, i - ps)));
            }
            return result;
        }

        /// <summary>找到匹配的括号结束位置，支持嵌套</summary>
        private static int FindJsonEnd(string s, int start, char open, char close)
        {
            if (start >= s.Length || s[start] != open) return -1;
            int depth = 1;
            bool inString = false;
            char stringChar = '\0';
            for (int i = start + 1; i < s.Length; i++)
            {
                if (inString)
                {
                    if (s[i] == '\\') { i++; continue; }
                    if (s[i] == stringChar) inString = false;
                    continue;
                }
                if (s[i] is '"' or '\'') { inString = true; stringChar = s[i]; continue; }
                if (s[i] == open) { depth++; continue; }
                if (s[i] == close) { depth--; if (depth == 0) return i; }
            }
            return -1;
        }

        /// <summary>美化 JSON 片段（缩进展开）</summary>
        private static string BeautifyJson(string json)
        {
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                return System.Text.Json.JsonSerializer.Serialize(doc.RootElement,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true, IndentSize = 2 });
            }
            catch
            {
                return json; // 解析失败则原样保留
            }
        }

        private void OnGoto(object? s, EventArgs e)
        {
            if (!int.TryParse(_txtGotoLine.Text, out int ln) || ln < 1) { MessageBox.Show(ParentForm, "请输入有效行号", "提示"); return; }
            if (ln > _lstContent.VirtualListSize) { MessageBox.Show(ParentForm, $"超出范围，共 {_lstContent.VirtualListSize:N0} 行", "提示"); return; }
            int idx = ln - 1; _lstContent.EnsureVisible(idx); _lstContent.SelectedIndices.Clear(); _lstContent.SelectedIndices.Add(idx); _lstContent.Focus();
        }

        // ==================== 搜索功能 ====================

        private int GetCurrentLine()
        {
            try
            {
                if (_lstContent.SelectedIndices.Count > 0)
                    return _lstContent.SelectedIndices[0];
                return _lstContent.TopItem?.Index ?? 0;
            }
            catch { return 0; }
        }

        private static int FindNearestIndex(List<int> results, int currentLine)
        {
            if (results.Count == 0) return -1;
            int lo = 0, hi = results.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                if (results[mid] == currentLine) return mid;
                if (results[mid] < currentLine) lo = mid + 1;
                else hi = mid - 1;
            }
            if (hi >= 0) return hi;
            return 0;
        }

        private void OnSearch()
        {
            string keyword = _txtSearch.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                _lblSearchCount.Text = "";
                return;
            }
            if (_lineOffsets == null)
            {
                _lblSearchCount.Text = "请等待索引完成";
                return;
            }
            _searchCts?.Cancel();
            _searchCts = new CancellationTokenSource();
            var token = _searchCts.Token;
            bool caseSensitive = _chkCaseSensitive.Checked;
            bool useRegex = _chkRegex.Checked;
            _lastSearchKeyword = keyword;
            _lastSearchCaseSensitive = caseSensitive;
            _lastSearchRegex = useRegex;
            _lblSearchCount.Text = "搜索中...";
            _searchResults.Clear();
            _searchCurrentIndex = -1;
            _lastNavigatedMatchLine = -1;
            int curLine = GetCurrentLine();
            _ = Task.Run(() => DoSearch(keyword, caseSensitive, useRegex, token, curLine));
        }

        private bool SearchParamsMatch()
        {
            return _txtSearch.Text.Trim() == _lastSearchKeyword
                && _chkCaseSensitive.Checked == _lastSearchCaseSensitive
                && _chkRegex.Checked == _lastSearchRegex;
        }

        private void DoSearch(string keyword, bool caseSensitive, bool useRegex, CancellationToken token, int curLine)
        {
            try
            {
                var results = new List<int>();
                StringComparison cmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

                Func<string, bool> matcher;
                if (useRegex)
                {
                    var opts = caseSensitive ? System.Text.RegularExpressions.RegexOptions.None : System.Text.RegularExpressions.RegexOptions.IgnoreCase;
                    var regex = new System.Text.RegularExpressions.Regex(keyword, opts);
                    matcher = (line) => regex.IsMatch(line);
                }
                else
                {
                    matcher = (line) => line.Contains(keyword, cmp);
                }

                int total = _lineOffsets!.Length;
                for (int i = 0; i < total; i++)
                {
                    if (token.IsCancellationRequested) return;
                    string line = ReadLine(i);
                    if (!string.IsNullOrEmpty(line) && matcher(line))
                        results.Add(i);
                    if (i % 500 == 0)
                    {
                        int cnt = results.Count;
                        int pct = (i + 1) * 100 / total;
                        BeginInvoke(() => _lblSearchCount.Text = $"搜索中... {cnt} 个 ({pct}%)");
                    }
                }

                int nearestLine = curLine;
                BeginInvoke(() =>
                {
                    _searchResults = results;
                    if (results.Count > 0)
                    {
                        int idx2 = FindNearestIndex(results, nearestLine);
                        NavigateTo(idx2);
                    }
                    else
                    {
                        _searchCurrentIndex = -1;
                        _lastNavigatedMatchLine = -1;
                        _lblSearchCount.Text = "未找到匹配";
                    }
                });
            }
            catch { }
        }

        private void NavigateTo(int index)
        {
            _searchCurrentIndex = index;
            int line = _searchResults[index];
            _lstContent.EnsureVisible(line);
            _lstContent.SelectedIndices.Clear();
            _lstContent.SelectedIndices.Add(line);
            _lstContent.Focus();
            _lastNavigatedMatchLine = line;
            _lblSearchCount.Text = $"第 {index + 1}/{_searchResults.Count} 个 | 行 {line + 1:N0}";
        }

        private bool UserHasMovedAway()
        {
            if (_lastNavigatedMatchLine < 0) return true;
            if (_lstContent.SelectedIndices.Count > 0 && _lstContent.SelectedIndices[0] != _lastNavigatedMatchLine)
                return true;
            int top = _lstContent.TopItem?.Index ?? 0;
            if (Math.Abs(top - _lastNavigatedMatchLine) > (_lstContent.ClientRectangle.Height / (_lstContent.Font.Height + 6)))
                return true;
            return false;
        }

        private void OnSearchPrev(object? s, EventArgs e)
        {
            if (!SearchParamsMatch())
                _searchResults.Clear();

            if (_searchResults.Count > 0)
            {
                if (UserHasMovedAway())
                {
                    int curLine = GetCurrentLine();
                    int idx2 = FindNearestIndex(_searchResults, curLine);
                    if (_searchResults[idx2] >= curLine)
                        idx2--;
                    if (idx2 < 0)
                    {
                        MessageBox.Show(ParentForm, "已到顶部，没有更多匹配项", "提示");
                        return;
                    }
                    NavigateTo(idx2);
                }
                else
                {
                    if (_searchCurrentIndex <= 0)
                    {
                        MessageBox.Show(ParentForm, "已到顶部，没有更多匹配项", "提示");
                        return;
                    }
                    NavigateTo(_searchCurrentIndex - 1);
                }
                return;
            }

            OnSearch();
        }

        private void OnSearchNext(object? s, EventArgs e)
        {
            if (!SearchParamsMatch())
                _searchResults.Clear();

            if (_searchResults.Count > 0)
            {
                if (UserHasMovedAway())
                {
                    int curLine = GetCurrentLine();
                    int idx2 = FindNearestIndex(_searchResults, curLine);
                    if (_searchResults[idx2] <= curLine)
                        idx2++;
                    if (idx2 >= _searchResults.Count)
                    {
                        MessageBox.Show(ParentForm, "已到底部，没有更多匹配项", "提示");
                        return;
                    }
                    NavigateTo(idx2);
                }
                else
                {
                    if (_searchCurrentIndex >= _searchResults.Count - 1)
                    {
                        MessageBox.Show(ParentForm, "已到底部，没有更多匹配项", "提示");
                        return;
                    }
                    NavigateTo(_searchCurrentIndex + 1);
                }
                return;
            }

            OnSearch();
        }

        private void OnSearchFirst(object? s, EventArgs e)
        {
            if (!SearchParamsMatch())
                _searchResults.Clear();
            if (_searchResults.Count > 0)
            {
                NavigateTo(0);
                return;
            }
            OnSearch();
        }

        private void OnSearchLast(object? s, EventArgs e)
        {
            if (!SearchParamsMatch())
                _searchResults.Clear();
            if (_searchResults.Count > 0)
            {
                NavigateTo(_searchResults.Count - 1);
                return;
            }
            OnSearch();
        }

        private void OnExport(object? s, EventArgs e)
        {
            string? sourcePath = _filePath;
            if (string.IsNullOrEmpty(sourcePath))
            {
                MessageBox.Show(ParentForm, "请先打开文件", "提示");
                return;
            }

            string keyword = _txtSearch.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                MessageBox.Show(ParentForm, "请先输入搜索关键词", "提示");
                return;
            }
            if (_lineOffsets == null)
            {
                MessageBox.Show(ParentForm, "请等待索引完成", "提示");
                return;
            }

            // 搜索条件已变更 → 重新搜索并导出
            if (!SearchParamsMatch())
                _searchResults.Clear();

            // 没有搜索结果 → 先搜索再导出
            if (_searchResults.Count == 0)
            {
                _searchCts?.Cancel();
                _searchCts = new CancellationTokenSource();
                var token = _searchCts.Token;
                bool caseSensitive = _chkCaseSensitive.Checked;
                bool useRegex = _chkRegex.Checked;
                _lastSearchKeyword = keyword;
                _lastSearchCaseSensitive = caseSensitive;
                _lastSearchRegex = useRegex;
                _lblSearchCount.Text = "搜索中...";
                _searchCurrentIndex = -1;
                _lastNavigatedMatchLine = -1;

                Status($"正在搜索 \"{keyword}\"...");
                _ = Task.Run(() => DoSearchAndExport(keyword, caseSensitive, useRegex, token));
                return;
            }

            // 已有搜索结果且条件匹配 → 直接导出
            DoExport();
        }

        /// <summary>搜索完成后自动导出</summary>
        private void DoSearchAndExport(string keyword, bool caseSensitive, bool useRegex, CancellationToken token)
        {
            try
            {
                var results = new List<int>();
                StringComparison cmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

                Func<string, bool> matcher;
                if (useRegex)
                {
                    var opts = caseSensitive ? System.Text.RegularExpressions.RegexOptions.None : System.Text.RegularExpressions.RegexOptions.IgnoreCase;
                    var regex = new System.Text.RegularExpressions.Regex(keyword, opts);
                    matcher = (line) => regex.IsMatch(line);
                }
                else
                {
                    matcher = (line) => line.Contains(keyword, cmp);
                }

                int total = _lineOffsets!.Length;
                for (int i = 0; i < total; i++)
                {
                    if (token.IsCancellationRequested) return;
                    string line = ReadLine(i);
                    if (!string.IsNullOrEmpty(line) && matcher(line))
                        results.Add(i);
                    if (i % 500 == 0)
                    {
                        int cnt = results.Count;
                        int pct = (i + 1) * 100 / total;
                        BeginInvoke(() => _lblSearchCount.Text = $"搜索中... {cnt} 个 ({pct}%)");
                    }
                }

                BeginInvoke(() =>
                {
                    _searchResults = results;
                    if (results.Count > 0)
                    {
                        int idx2 = FindNearestIndex(results, GetCurrentLine());
                        NavigateTo(idx2);
                        DoExport();
                    }
                    else
                    {
                        _searchCurrentIndex = -1;
                        _lastNavigatedMatchLine = -1;
                        _lblSearchCount.Text = "未找到匹配，无法导出";
                        Status("未找到匹配行");
                    }
                });
            }
            catch { }
        }

        private void DoExport()
        {
            string? sourcePath = _filePath;
            if (_searchResults.Count == 0) return;

            string dir = Path.GetDirectoryName(sourcePath)!;
            string fileName = Path.GetFileName(sourcePath);     // 如 xx.log
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string outName = $"{fileName}-{timestamp}{Path.GetExtension(sourcePath)}";
            string outPath = Path.Combine(dir, outName);

            try
            {
                Cursor = Cursors.WaitCursor;
                Status($"正在导出 {_searchResults.Count} 行到 {outName}...");
                using var sw = new StreamWriter(outPath, false, Encoding.UTF8);
                for (int i = 0; i < _searchResults.Count; i++)
                {
                    int lineIdx = _searchResults[i];
                    string lineText = _cache.TryGetValue(lineIdx, out var v) ? v : ReadLine(lineIdx);
                    sw.WriteLine(lineText);
                }
                sw.Flush();
                Status($"已导出 {_searchResults.Count} 行 -> {outName}");
                MessageBox.Show(ParentForm, $"导出完成！\n\n文件: {outName}\n路径: {dir}\n行数: {_searchResults.Count:N0}", "导出成功");
                try { BeginInvoke(() => FileExported?.Invoke(this, EventArgs.Empty)); } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ParentForm, $"导出失败: {ex.Message}", "错误");
                Status("导出失败");
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void Cleanup()
        {
            _cancel = true;
            try { _timer.Stop(); } catch { }
            try { _timerScroll?.Stop(); } catch { }
            _indexReady = false;
            _searchCts?.Cancel();
            _searchResults.Clear();
            _searchCurrentIndex = -1;
            _lastNavigatedMatchLine = -1;
            _lastSearchKeyword = "";
            if (IsHandleCreated)
                try { BeginInvoke(() => _lblSearchCount.Text = ""); } catch { }
            lock (_readLock) { _readStream?.Dispose(); _readStream = null; }
            _filePath = null; _lineOffsets = null; _cache.Clear(); _cS = _cE = -1;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Cleanup();
                _timer?.Dispose();
                _timerScroll?.Dispose();
                _searchCts?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 自定义 ListView：重写 WndProc 处理水平滚动，用 Win32 API 控制真正的水平滚动
    /// </summary>
    public class HScrollListView : ListView
    {
        private const int WM_HSCROLL = 0x114;
        private const int WM_SIZE = 0x0005;
        private const int WM_LBUTTONDBLCLK = 0x0203;
        private const int SB_HORZ = 0;

        [StructLayout(LayoutKind.Sequential)]
        private struct SCROLLINFO
        {
            public int cbSize;
            public int fMask;
            public int nMin;
            public int nMax;
            public int nPage;
            public int nPos;
            public int nTrackPos;
        }

        [DllImport("user32.dll")]
        private static extern int SetScrollInfo(IntPtr hWnd, int fnBar, ref SCROLLINFO lpsi, bool fRedraw);
        [DllImport("user32.dll")]
        private static extern int GetScrollInfo(IntPtr hWnd, int fnBar, ref SCROLLINFO lpsi);

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int HScrollPos { get; set; }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int MaxContentWidth { get; set; } = 1000;

        public event EventHandler<MouseEventArgs>? ItemDoubleClick;

        public void InitHScroll()
        {
            if (!IsHandleCreated) return;
            if (Columns.Count > 0)
                Columns[0].Width = Math.Max(MaxContentWidth, ClientRectangle.Width) + 50;
            var si = new SCROLLINFO();
            si.cbSize = Marshal.SizeOf(si);
            si.fMask = 0x0003;
            si.nMin = 0;
            si.nMax = Math.Max(MaxContentWidth, ClientRectangle.Width);
            si.nPage = ClientRectangle.Width;
            si.nPos = 0;
            SetScrollInfo(Handle, SB_HORZ, ref si, true);
            HScrollPos = 0;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HSCROLL)
            {
                var si = new SCROLLINFO();
                si.cbSize = Marshal.SizeOf(si);
                si.fMask = 0x0001 | 0x0010 | 0x0004;
                GetScrollInfo(Handle, SB_HORZ, ref si);

                int scrollCode = (int)m.WParam.ToInt64() & 0xFFFF;
                int newPos = HScrollPos;
                int pageSize = ClientRectangle.Width;
                int maxScroll = Math.Max(MaxContentWidth - pageSize, 0);

                switch (scrollCode)
                {
                    case 0: newPos -= 40; break;
                    case 1: newPos += 40; break;
                    case 2: newPos -= pageSize; break;
                    case 3: newPos += pageSize; break;
                    case 4:
                    case 5:
                        newPos = si.nTrackPos; break;
                    case 6: newPos = 0; break;
                    case 7: newPos = maxScroll; break;
                }

                newPos = Math.Max(0, Math.Min(newPos, maxScroll));
                if (newPos != HScrollPos)
                {
                    HScrollPos = newPos;
                    si.nPos = newPos;
                    si.fMask = 0x0002;
                    SetScrollInfo(Handle, SB_HORZ, ref si, true);
                    Invalidate();
                }
                return;
            }
            if (m.Msg == WM_SIZE)
            {
                base.WndProc(ref m);
                BeginInvoke(new Action(InitHScroll));
                return;
            }
            base.WndProc(ref m);
        }
    }
}
