using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Encodings.Web;
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
        private Button _btnSearchFirst = null!, _btnSearchPrev = null!, _btnSearchNext = null!, _btnSearchLast = null!, _btnExport = null!, _btnReverseExport = null!;
        private Label _lblStatus = null!, _lblIndexProgress = null!, _lblSearchCount = null!;
        private NumericUpDown _numContextLines = null!;
        private Label _lblContextLines = null!;
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
        private bool _hasuserNavigated; // 用户在搜索过程中是否已手动跳转过

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

            _btnOpenFile = new Button { Text = "打开文件", Size = new Size(90, 28), Location = new Point(10, 6), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(66, 165, 245), ForeColor = Color.White };
            _btnOpenFile.FlatAppearance.BorderSize = 1;
            _btnOpenFile.FlatAppearance.BorderColor = Color.FromArgb(30, 136, 229);
            _btnOpenFile.Click += OnOpen;
            _btnGotoLine = new Button { Text = "跳转", Size = new Size(55, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(144, 180, 210), ForeColor = Color.White };
            _btnGotoLine.FlatAppearance.BorderSize = 1;
            _btnGotoLine.FlatAppearance.BorderColor = Color.FromArgb(120, 160, 200);
            _btnGotoLine.Click += OnGoto;
            _txtGotoLine = new TextBox { Size = new Size(120, 28), PlaceholderText = "行号" };
            _txtGotoLine.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OnGoto(s, e); };
            _lblIndexProgress = new Label { Text = "就绪", AutoSize = true, ForeColor = Color.FromArgb(90, 130, 170), Font = new Font("Microsoft YaHei UI", 8F) };
            _lblIndexProgress.Margin = new Padding(3, 7, 0, 0);
            _lblStatus = new Label { Text = "请打开一个文本文件", AutoSize = true, ForeColor = Color.FromArgb(90, 130, 170), Font = new Font("Microsoft YaHei UI", 8.5F) };
            _lblStatus.Margin = new Padding(3, 7, 0, 0);

            _topPanel = new TableLayoutPanel { Dock = DockStyle.Top, Height = 88, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(235, 245, 255) };
            _topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            _topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 1));
            _topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 43));

            var row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10, 8, 0, 0) };
            row1.Controls.AddRange(new Control[] { _btnOpenFile, _txtGotoLine, _btnGotoLine, _lblIndexProgress, _lblStatus });

            var separator = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(173, 209, 237) };

            _txtSearch = new TextBox { Size = new Size(320, 28), PlaceholderText = "搜索..." };
            _txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OnSearch(); };
            _txtSearch.TextChanged += (s, e) => { _lastSearchKeyword = ""; };
            _chkCaseSensitive = new CheckBox { Text = "Aa", AutoSize = true, Checked = false, TextAlign = ContentAlignment.MiddleRight };
            _chkCaseSensitive.Margin = new Padding(3, 6, 0, 0);
            _chkCaseSensitive.CheckedChanged += (s, e) => { _searchResults.Clear(); _searchCurrentIndex = -1; _lblSearchCount.Text = ""; _lastSearchKeyword = ""; };
            _chkRegex = new CheckBox { Text = ".*", AutoSize = true, Checked = false, TextAlign = ContentAlignment.MiddleRight };
            _chkRegex.Margin = new Padding(3, 6, 0, 0);
            _chkRegex.CheckedChanged += (s, e) => { _searchResults.Clear(); _searchCurrentIndex = -1; _lblSearchCount.Text = ""; _lastSearchKeyword = ""; };
            _btnSearchFirst = new Button { Text = "|<", Size = new Size(35, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(144, 180, 210), ForeColor = Color.White };
            _btnSearchFirst.FlatAppearance.BorderSize = 1;
            _btnSearchFirst.FlatAppearance.BorderColor = Color.FromArgb(120, 160, 200);
            _btnSearchFirst.Click += OnSearchFirst;
            _btnSearchPrev = new Button { Text = "上一个", Size = new Size(65, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(144, 180, 210), ForeColor = Color.White };
            _btnSearchPrev.FlatAppearance.BorderSize = 1;
            _btnSearchPrev.FlatAppearance.BorderColor = Color.FromArgb(120, 160, 200);
            _btnSearchPrev.Click += OnSearchPrev;
            _btnSearchNext = new Button { Text = "下一个", Size = new Size(65, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(144, 180, 210), ForeColor = Color.White };
            _btnSearchNext.FlatAppearance.BorderSize = 1;
            _btnSearchNext.FlatAppearance.BorderColor = Color.FromArgb(120, 160, 200);
            _btnSearchNext.Click += OnSearchNext;
            _btnSearchLast = new Button { Text = ">|", Size = new Size(35, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(144, 180, 210), ForeColor = Color.White };
            _btnSearchLast.FlatAppearance.BorderSize = 1;
            _btnSearchLast.FlatAppearance.BorderColor = Color.FromArgb(120, 160, 200);
            _btnSearchLast.Click += OnSearchLast;
            _lblSearchCount = new Label { Text = "", AutoSize = true, ForeColor = Color.FromArgb(90, 130, 170), Font = new Font("Microsoft YaHei UI", 8F) };
            _lblSearchCount.Margin = new Padding(3, 7, 0, 0);

            var row2 = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10, 8, 0, 0) };
            _btnExport = new Button { Text = "导出结果", Size = new Size(75, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(66, 165, 245), ForeColor = Color.White };
            _btnExport.FlatAppearance.BorderSize = 1;
            _btnExport.FlatAppearance.BorderColor = Color.FromArgb(30, 136, 229);
            _btnExport.Click += OnExport;
            _btnReverseExport = new Button { Text = "反向导出", Size = new Size(75, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(255, 152, 0), ForeColor = Color.White };
            _btnReverseExport.FlatAppearance.BorderSize = 1;
            _btnReverseExport.FlatAppearance.BorderColor = Color.FromArgb(245, 124, 0);
            _btnReverseExport.Click += OnReverseExport;
            _lblContextLines = new Label { Text = "上下行:", AutoSize = true, ForeColor = Color.FromArgb(90, 130, 170), Font = new Font("Microsoft YaHei UI", 8F), TextAlign = ContentAlignment.MiddleRight };
            _lblContextLines.Margin = new Padding(8, 7, 0, 0);
            _numContextLines = new NumericUpDown { Minimum = 0, Maximum = 999, Value = 0, Width = 55 };
            row2.Controls.AddRange(new Control[] { _txtSearch, _chkCaseSensitive, _chkRegex, _btnSearchFirst, _btnSearchPrev, _btnSearchNext, _btnSearchLast, _lblContextLines, _numContextLines, _btnExport, _btnReverseExport, _lblSearchCount });

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
            using var backBrush = new SolidBrush(selected ? Color.FromArgb(66, 165, 245) : _lstContent.BackColor);
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

        private async void OnDoubleClick(object? s, MouseEventArgs e)
        {
            int idx = _lstContent.SelectedIndices.Count > 0
                ? _lstContent.SelectedIndices[0]
                : (e.Y / (_lstContent.Font.Height + 6) + (_lstContent.TopItem?.Index ?? 0));
            if (idx < 0 || idx >= _lstContent.VirtualListSize) return;
            _lstContent.SelectedIndices.Clear();
            _lstContent.SelectedIndices.Add(idx);

            string originalText = _cache.TryGetValue(idx, out var v) ? v : "";
            if (string.IsNullOrEmpty(originalText)) return;

            var f = new Form
            {
                Text = $"行 {idx + 1}", Size = new Size(950, 550), StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = true, MaximizeBox = true, ShowIcon = false, ShowInTaskbar = false
            };

            // 顶部工具栏
            var topPanel = new Panel { Height = 44, Dock = DockStyle.Top, BackColor = Color.FromArgb(235, 245, 255) };
            var btnFormat = new Button { Text = "格式化", Size = new Size(90, 30), Location = new Point(10, 7), FlatStyle = FlatStyle.Flat };
            var btnRaw = new Button { Text = "原始", Size = new Size(90, 30), Location = new Point(110, 7), FlatStyle = FlatStyle.Flat };

            void SetActive(Button active, Button inactive)
            {
                active.BackColor = Color.FromArgb(66, 165, 245);
                active.ForeColor = Color.White;
                active.FlatAppearance.BorderSize = 1;
                active.FlatAppearance.BorderColor = Color.FromArgb(30, 136, 229);
                inactive.BackColor = Color.FromArgb(200, 220, 240);
                inactive.ForeColor = Color.FromArgb(50, 80, 120);
                inactive.FlatAppearance.BorderSize = 1;
                inactive.FlatAppearance.BorderColor = Color.FromArgb(160, 190, 220);
            }
            SetActive(btnFormat, btnRaw);

            var rtb = new RichTextBox
            {
                Multiline = true, ReadOnly = true, Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F), ScrollBars = RichTextBoxScrollBars.Both,
                WordWrap = false, BackColor = Color.White,
                Text = "正在格式化..."
            };

            string formattedText = "";
            btnFormat.Click += async (_, _) =>
            {
                SetActive(btnFormat, btnRaw);
                rtb.Text = "正在构建 RTF...";
                rtb.WordWrap = false;
                rtb.ScrollBars = RichTextBoxScrollBars.Both;
                var clickDeadline = Deadline.FromSeconds(5.0);
                string? rtf = null;
                try { rtf = await Task.Run(() => BuildRtfForLine(formattedText, clickDeadline)).ConfigureAwait(true); }
                catch { /* 解析失败，显示原始文本 */ }
                if (f.IsDisposed) return;
                try { rtb.Rtf = rtf ?? originalText; }
                catch { try { rtb.Text = originalText; } catch { } }
            };
            btnRaw.Click += (_, _) => {
                SetActive(btnRaw, btnFormat);
                rtb.WordWrap = true;
                rtb.ScrollBars = RichTextBoxScrollBars.Vertical;
                rtb.Text = originalText;
            };

            topPanel.Controls.Add(btnFormat);
            topPanel.Controls.Add(btnRaw);
            f.Controls.Add(rtb);
            f.Controls.Add(topPanel);

            f.Show();

            // 格式化在后台线程完成，完成后自动显示 RTF
            try { formattedText = await Task.Run(() => BuildFormattedText(originalText)).ConfigureAwait(true); }
            catch { formattedText = originalText; }
            if (f.IsDisposed) return;

            var autoDeadline = Deadline.FromSeconds(5.0);
            string? rtf = null;
            try { rtf = await Task.Run(() => BuildRtfForLine(formattedText, autoDeadline)).ConfigureAwait(true); }
            catch { /* 解析失败，显示原始文本 */ }
            if (f.IsDisposed) return;
            try { rtb.Rtf = rtf ?? originalText; }
            catch { try { rtb.Text = originalText; } catch { } }
        }

        private static void ShowFormattedRtf(RichTextBox rtb, string formattedText) => ShowFormattedRtf(rtb, formattedText, Deadline.Default);

        private static void ShowFormattedRtf(RichTextBox rtb, string formattedText, Deadline deadline)
        {
            try
            {
                string? rtf = BuildRtfForLine(formattedText, deadline);
                rtb.Text = rtf ?? formattedText; // null 表示超时或过大，直接显示纯文本
            }
            catch { rtb.Text = formattedText; }
        }

        /// <summary>超时截止检查器，防止复杂内容格式化导致卡死</summary>
        private readonly struct Deadline
        {
            private readonly long _limitTicks;
            private Deadline(long limitTicks) { _limitTicks = limitTicks; }
            public static Deadline FromSeconds(double seconds) => new(DateTime.UtcNow.Ticks + (long)(seconds * TimeSpan.TicksPerSecond));
            public static Deadline Default => FromSeconds(2.0);
            public bool IsExpired => DateTime.UtcNow.Ticks >= _limitTicks;
            public void ThrowIfExpired()
            {
                if (IsExpired) throw new InvalidOperationException("__DEADLINE_EXPIRED__");
            }
        }

        /// <summary>构建格式化文本：JSON 缩进换行，其余内容按逗号换行</summary>
        private static string BuildFormattedText(string line)
        {
            Deadline deadline = Deadline.Default;
            try
            {
                return BuildFormattedTextImpl(line, ref deadline);
            }
            catch (InvalidOperationException ex) when (ex.Message == "__DEADLINE_EXPIRED__")
            {
                return line; // 超时保底：返回原始文本
            }
            catch
            {
                return line; // 解析异常保底：返回原始文本
            }
        }

        /// <summary>内容片段类型</summary>
        private enum ContentSegType { JsonObject, JsonArray, BracketBlock, PlainText }

        /// <summary>内容片段</summary>
        private readonly struct ContentSegment
        {
            public readonly ContentSegType Type;
            public readonly int Start;
            public readonly int End; // inclusive

            public ContentSegment(ContentSegType type, int start, int end)
            {
                Type = type; Start = start; End = end;
            }
        }

        private static string BuildFormattedTextImpl(string line, ref Deadline deadline)
        {
            deadline.ThrowIfExpired();

            // Phase 1: 提取所有顶层片段：JSON对象、JSON数组、【】块、纯文本
            var segments = ExtractTopLevelSegments(line, deadline);
            if (segments.Count == 0) return line;

            // Phase 2: 格式化每个片段
            var sb = new StringBuilder();
            foreach (var seg in segments)
            {
                deadline.ThrowIfExpired();
                string raw = line.Substring(seg.Start, seg.End - seg.Start + 1);

                switch (seg.Type)
                {
                    case ContentSegType.JsonObject:
                        AppendFormattedJson(sb, raw, deadline);
                        break;
                    case ContentSegType.JsonArray:
                        AppendFormattedJson(sb, raw, deadline);
                        break;
                    case ContentSegType.BracketBlock:
                        AppendFormattedBracketBlock(sb, raw, deadline);
                        break;
                    case ContentSegType.PlainText:
                        AppendFormattedPlainText(sb, raw, deadline);
                        break;
                }
            }
            return sb.ToString();
        }

        /// <summary>将一行文本拆分成顶层片段序列</summary>
        private static List<ContentSegment> ExtractTopLevelSegments(string line, Deadline deadline)
        {
            var result = new List<ContentSegment>();
            int i = 0;
            // 记录括号深度，用于在【】内跳过嵌套的{}/[]
            int bracketDepth = -1; // -1 表示不在【】内
            int bracketStart = -1;

            while (i < line.Length)
            {
                deadline.ThrowIfExpired();

                if (bracketDepth >= 0)
                {
                    // 在【】内部，需要找到闭合的】
                    if (line[i] == '】')
                    {
                        bracketDepth--;
                        if (bracketDepth < 0)
                        {
                            result.Add(new ContentSegment(ContentSegType.BracketBlock, bracketStart, i));
                            bracketDepth = -1;
                            bracketStart = -1;
                            i++;
                            continue;
                        }
                    }
                    // 嵌套的【只计数不分割
                    else if (line[i] == '【')
                    {
                        bracketDepth++;
                    }
                    i++;
                    continue;
                }

                char c = line[i];

                // JSON 对象 { ... }
                if (c == '{')
                {
                    int end = FindJsonEnd(line, i, '{', '}', deadline);
                    if (end > 0)
                    {
                        // 先把之前的纯文本收尾
                        FlushPlainSegments(result, line, i, deadline);
                        result.Add(new ContentSegment(ContentSegType.JsonObject, i, end));
                        i = end + 1;
                        continue;
                    }
                }

                // JSON 数组 [ ... ] — 仅当内容可解析为有效 JSON 时才识别
                if (c == '[')
                {
                    int end = FindJsonEnd(line, i, '[', ']', deadline);
                    if (end > 0)
                    {
                        string candidate = line.Substring(i, end - i + 1);
                        if (IsValidJsonArray(candidate, deadline))
                        {
                            FlushPlainSegments(result, line, i, deadline);
                            result.Add(new ContentSegment(ContentSegType.JsonArray, i, end));
                            i = end + 1;
                            continue;
                        }
                        // 不是有效 JSON，当作普通文本继续
                    }
                }

                // 【 块开始
                if (c == '【')
                {
                    bracketDepth = 0;
                    bracketStart = i;
                    i++;
                    continue;
                }

                // 普通字符 → 累积到纯文本段
                // 但纯文本段由 FlushPlainSegments 在处理完特殊段落时创建
                // 所以我们先跳过特殊字符，在 FlushPlainSegments 中取前面的纯文本
                i++;
            }

            // 最后的纯文本
            FlushPlainSegments(result, line, line.Length, deadline);
            return result;
        }

        /// <summary>将 [prevStart, currStart) 范围的纯文本加入片段列表</summary>
        private static void FlushPlainSegments(List<ContentSegment> segments, string line, int currStart, Deadline deadline)
        {
            if (segments.Count == 0 && currStart <= 0) return;
            int prevEnd = -1;
            if (segments.Count > 0)
                prevEnd = segments[^1].End;

            int pStart = prevEnd + 1;
            if (pStart >= currStart) return;

            // 跳过纯空白
            string chunk = line.Substring(pStart, currStart - pStart);
            if (string.IsNullOrWhiteSpace(chunk)) return;

            segments.Add(new ContentSegment(ContentSegType.PlainText, pStart, currStart - 1));
        }

        /// <summary>检查字符串是否可解析为有效的 JSON 数组</summary>
        private static bool IsValidJsonArray(string text, Deadline deadline)
        {
            if (text.Length < 2 || text[0] != '[' || text[^1] != ']') return false;
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(text);
                return doc.RootElement.ValueKind == System.Text.Json.JsonValueKind.Array;
            }
            catch { return false; }
        }

        /// <summary>格式化 JSON 对象/数组（美化缩进）</summary>
        private static void AppendFormattedJson(StringBuilder sb, string raw, Deadline deadline)
        {
            // 尝试标准 JSON 美化
            string? beautified = TryBeautifyJson(raw, deadline);
            if (beautified != null)
            {
                if (sb.Length > 0) sb.AppendLine();
                sb.Append(beautified);
                return;
            }

            // 尝试 key=value → JSON 转换
            if (raw[0] == '{')
            {
                string? kvJson = TryConvertEqualsToJson(raw, deadline);
                if (kvJson != null)
                {
                    if (sb.Length > 0) sb.AppendLine();
                    sb.Append(kvJson);
                    return;
                }
            }

            // 数组：尝试转换 key=value 对象数组
            if (raw[0] == '[')
            {
                string? kvArrayJson = TryConvertArrayToJson(raw, deadline);
                if (kvArrayJson != null)
                {
                    if (sb.Length > 0) sb.AppendLine();
                    sb.Append(kvArrayJson);
                    return;
                }
            }

            // 都失败 → 原样保留
            if (sb.Length > 0) sb.AppendLine();
            sb.Append(raw);
        }

        /// <summary>格式化【】块：提取内部 JSON 美化，其余按逗号分行</summary>
        private static void AppendFormattedBracketBlock(StringBuilder sb, string raw, Deadline deadline)
        {
            // 去掉最外层【】
            string inner = raw;
            if (inner.Length >= 2 && inner[0] == '【' && inner[^1] == '】')
                inner = inner.Substring(1, inner.Length - 2);

            if (string.IsNullOrWhiteSpace(inner))
            {
                if (sb.Length > 0) sb.AppendLine();
                sb.Append(raw);
                return;
            }

            // 查找内部是否有 JSON { } 或 [ ] 块
            int jsonStart = -1;
            char jsonOpen = '\0';
            char jsonClose = '\0';
            for (int i = 0; i < inner.Length; i++)
            {
                if (i % 64 == 0) deadline.ThrowIfExpired();
                if (inner[i] == '{') { jsonStart = i; jsonOpen = '{'; jsonClose = '}'; break; }
                if (inner[i] == '[') { jsonStart = i; jsonOpen = '['; jsonClose = ']'; break; }
            }

            if (jsonStart >= 0)
            {
                int jsonEnd = FindJsonEnd(inner, jsonStart, jsonOpen, jsonClose, deadline);
                if (jsonEnd > jsonStart)
                {
                    // 提取 JSON 前的文本（可能是 ClassName）
                    string prefix = jsonStart > 0 ? inner[..jsonStart].Trim() : "";
                    string jsonRaw = inner.Substring(jsonStart, jsonEnd - jsonStart + 1);
                    string? beautified = TryBeautifyJson(jsonRaw, deadline);

                    // 提取 JSON 后的文本
                    string suffix = jsonEnd + 1 < inner.Length ? inner[(jsonEnd + 1)..].Trim() : "";

                    if (sb.Length > 0) sb.AppendLine();

                    // 输出前缀（过滤出标识符字符）
                    if (!string.IsNullOrEmpty(prefix))
                    {
                        int nameEnd = 0;
                        while (nameEnd < prefix.Length && (char.IsLetterOrDigit(prefix[nameEnd]) || prefix[nameEnd] == '_')) nameEnd++;
                        if (nameEnd > 0)
                        {
                            sb.Append(prefix[..nameEnd]);
                            sb.AppendLine();
                        }
                    }

                    if (beautified != null)
                    {
                        sb.Append(beautified);
                    }
                    else
                    {
                        sb.Append(jsonRaw);
                    }

                    if (!string.IsNullOrEmpty(suffix))
                    {
                        sb.AppendLine();
                        sb.Append("剩余: " + suffix);
                    }
                    return;
                }
            }

            // 内部没有 JSON → 按逗号分离
            if (sb.Length > 0) sb.AppendLine();
            var parts = inner.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0) sb.AppendLine();
                sb.Append(parts[i].Trim());
            }
        }

        /// <summary>格式化纯文本：不盲目拆分</summary>
        private static void AppendFormattedPlainText(StringBuilder sb, string raw, Deadline deadline)
        {
            string text = raw.Trim().TrimStart(',').TrimEnd(',');
            if (string.IsNullOrEmpty(text)) return;

            // 判断是否为 Java 栈跟踪（包含 <= 符号）
            bool isStackTrace = text.Contains("<=") || text.Contains("->");
            // 判断是否包含中文
            bool hasChinese = text.Any(c => c >= 0x4E00 && c <= 0x9FFF);

            if (isStackTrace)
            {
                // 堆栈跟踪：按 <= 分割换行，保持可读性
                var stackParts = text.Split(new[] { "<=" }, StringSplitOptions.None);
                for (int i = 0; i < stackParts.Length; i++)
                {
                    var sp = stackParts[i].Trim().TrimEnd("<= ".ToCharArray()).Trim();
                    if (string.IsNullOrEmpty(sp)) continue;
                    if (sb.Length > 0) sb.AppendLine();
                    if (i > 0) sb.Append("  ");
                    sb.Append(sp);
                }
            }
            else if (hasChinese)
            {
                // 含中文的日志行：保留整行，不做逗号拆分（防止拆分包含 = 的混合内容）
                if (sb.Length > 0) sb.AppendLine();
                sb.Append(text);
            }
            else
            {
                // 通用文本（无中文）：包含 key=value 的对按逗号拆分，否则整段保留
                bool hasKeyValue = text.Contains('=') || text.Contains("：");
                if (hasKeyValue)
                {
                    var kvParts = text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                    foreach (var part in kvParts)
                    {
                        var tp = part.Trim();
                        if (string.IsNullOrEmpty(tp)) continue;
                        if (sb.Length > 0) sb.AppendLine();
                        sb.Append(tp);
                    }
                }
                else
                {
                    if (sb.Length > 0) sb.AppendLine();
                    sb.Append(text);
                }
            }
        }

        /// <summary>按中文标点拆分段</summary>
        private static List<string> SplitChineseText(string text)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            foreach (char c in text)
            {
                if (c is '，' or '；' or '。' or ',' or ';')
                {
                    if (sb.Length > 0)
                    {
                        result.Add(sb.ToString());
                        sb.Clear();
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }
            if (sb.Length > 0) result.Add(sb.ToString());
            return result;
        }

        /// <summary>在逗号分隔的字符串中按顶层（忽略嵌套 {}[]）拆分成条目</summary>
        private static List<string> SplitTopLevelEntries(string s, Deadline deadline)
        {
            var result = new List<string>();
            int depth = 0;
            int start = 0;
            for (int i = 0; i < s.Length; i++)
            {
                if (i % 256 == 0) deadline.ThrowIfExpired();
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
        private static string? TryBeautifyJson(string json, Deadline deadline)
        {
            // 第一次尝试原样解析
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                return System.Text.Json.JsonSerializer.Serialize(doc.RootElement,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
            }
            catch { }

            // 修复未引号的 key 后再试
            string fixedJson = FixUnquotedKeys(json, deadline);
            if (fixedJson != json)
            {
                try
                {
                    using var doc = System.Text.Json.JsonDocument.Parse(fixedJson);
                    return System.Text.Json.JsonSerializer.Serialize(doc.RootElement,
                        new System.Text.Json.JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
                }
                catch { }
            }
            return null;
        }

        /// <summary>将 JSON 中未加引号的 key（如 {1:..., key:...}）加上引号，带超时检查</summary>
        private static string FixUnquotedKeys(string json, Deadline deadline)
        {
            var sb = new StringBuilder(json.Length + 32);
            bool inStr = false;
            char strChar = '\0';

            for (int i = 0; i < json.Length; i++)
            {
                if (i % 256 == 0) deadline.ThrowIfExpired();
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
                if (c is '{' or '[' or ',')
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
                        if (l < json.Length && json[l] is ':' or '=')
                        {
                            // 这是未引号的 key → 加双引号
                            sb.Append('"');
                            sb.Append(json, j, k - j);
                            sb.Append("\":");
                            i = k; // 跳到 : 或 =，循环 i++ 会再跳一位
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
        private static string? TryConvertEqualsToJson(string text, Deadline deadline)
        {
            if (text.Length < 2 || text[0] != '{' || text[^1] != '}') return null;
            string inner = text.Substring(1, text.Length - 2).Trim();
            if (string.IsNullOrEmpty(inner)) return null;
            // 如果包含标准 JSON 的引号格式（"key": val 或 'key': val），交由 FixUnquotedKeys 处理
            if (inner.Contains("\":")) return null;
            // 检查是否大部分条目是 key=value 格式
            int eqCount = 0, colonCount = 0, totalEntries = 0;
            var entries = SplitTopLevelEntries(inner, deadline);
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
                    jsonSb.Append($"  \"{key}\": ");
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
        private static string? TryConvertArrayToJson(string text, Deadline deadline)
        {
            if (text.Length < 2 || text[0] != '[' || text[^1] != ']') return null;
            string inner = text.Substring(1, text.Length - 2).Trim();
            if (string.IsNullOrEmpty(inner)) return null;
            // 已包含标准 JSON 结构则跳过
            if (inner.TrimStart().StartsWith("{")) return null;
            // 按顶层逗号拆分条目
            var items = SplitTopLevelEntries(inner, deadline);
            if (items.Count == 0) return null;
            // 检查是否每个条目都包含 {...}
            bool allHaveBrace = items.All(i => i.Contains('{'));
            if (!allHaveBrace) return null;

            var jsonSb = new StringBuilder();
            var convertedItems = new List<string>();
            foreach (string item in items)
            {
                deadline.ThrowIfExpired();
                string trimmed = item.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;
                int braceStart = trimmed.IndexOf('{');
                if (braceStart < 0) { convertedItems.Add($"  {trimmed}"); continue; }
                string prefix = braceStart > 0 ? trimmed[..braceStart].Trim() : "";
                string braceContent = trimmed[braceStart..].Trim();
                string? objJson = TryBeautifyJson(braceContent, deadline);
                if (objJson == null)
                    objJson = TryConvertEqualsToJson(braceContent, deadline);
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

        /// <summary>将值智能处理：数字不加引号，其他加引号（显示用，不转义引号）</summary>
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
            sb.Append(val);
            sb.Append('\"');
        }

        // ==================== 混合内容格式化引擎 ====================

        /// <summary>片段类型</summary>
        private enum SegType { JsonBlock, KeyValuePair, QuotedString, Number, BooleanNull, PlainText }

        /// <summary>将已格式化文本（多行）渲染为带语法高亮的 RTF</summary>
        private static string BuildRtfForLine(string formattedText) => BuildRtfForLine(formattedText, Deadline.Default);

        private static string? BuildRtfForLine(string formattedText, Deadline deadline)
        {
            // 超大文本跳过 RTF 着色，返回 null 表示用纯文本
            if (formattedText.Length > 512 * 1024) return null;

            var lines = formattedText.Split('\n');
            var sb = new StringBuilder();
            sb.AppendLine($@"{{\rtf1\ansi\ansicpg{Encoding.Default.CodePage}\deff0");
            sb.AppendLine($@"{{\fonttbl{{\f0\fnil\fcharset{Encoding.Default.CodePage} Consolas;}}}}");
            sb.AppendLine(@"{\colortbl ;\red0\green122\blue204;\red22\green138\blue8;\red156\green39\blue176;\red117\green117\blue117;\red216\green67\blue21;\red0\green121\blue203;\red30\green30\blue30;}");
            sb.Append(@"\f0\fs20 ");

            int lineIdx = 0;
            bool inJson = false;
            foreach (var rawLine in lines)
            {
                if (lineIdx++ % 256 == 0) deadline.ThrowIfExpired();

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
                    var tokens = TokenizeLine(rl, deadline);
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

                /// <summary>RTF 转义（ANSI 模式，所有字节用 \'HH hex 编码）</summary>
        private static string EscapeRtf(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var sb = new StringBuilder(text.Length * 2);
            foreach (byte b in Encoding.Default.GetBytes(text))
            {
                switch (b)
                {
                    case 92:  sb.Append(@"\\"); break;  // \
                    case 123: sb.Append(@"\{ "); break; // {
                    case 125: sb.Append(@"\} "); break; // }
                    case 10:  sb.Append(@"\par "); break; // \n
                    case 13:  break;                  // \r
                    default:  sb.Append($"\\'{(int)b:X2}"); break;
                }
            }
            return sb.ToString();
        }

private static List<(SegType Type, string Text)> TokenizeLine(string line) => TokenizeLine(line, Deadline.Default);

        private static List<(SegType Type, string Text)> TokenizeLine(string line, Deadline deadline)
        {
            var result = new List<(SegType, string)>();
            int i = 0;
            while (i < line.Length)
            {
                if (i % 256 == 0) deadline.ThrowIfExpired();

                // 1) JSON 对象 { ... }
                if (line[i] == '{')
                {
                    int end = FindJsonEnd(line, i, '{', '}', deadline);
                    if (end > 0)
                    {
                        string raw = line.Substring(i, end - i + 1);
                        string beautified;
                        try { beautified = BeautifyJson(raw, deadline); }
                        catch { beautified = raw; }
                        result.Add((SegType.JsonBlock, beautified));
                        i = end + 1;
                        continue;
                    }
                }

                // 2) JSON 数组 [ ... ]
                if (line[i] == '[')
                {
                    int end = FindJsonEnd(line, i, '[', ']', deadline);
                    if (end > 0)
                    {
                        string raw = line.Substring(i, end - i + 1);
                        string beautified;
                        try { beautified = BeautifyJson(raw, deadline); }
                        catch { beautified = raw; }
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

        /// <summary>找到匹配的括号结束位置，支持嵌套（无超时版本，供 TokenizeLine 等调用）</summary>
        private static int FindJsonEnd(string s, int start, char open, char close) => FindJsonEnd(s, start, open, close, Deadline.Default);

        /// <summary>找到匹配的括号结束位置，支持嵌套，带超时检查</summary>
        private static int FindJsonEnd(string s, int start, char open, char close, Deadline deadline)
        {
            if (start >= s.Length || s[start] != open) return -1;
            int depth = 1;
            bool inString = false;
            char stringChar = '\0';
            for (int i = start + 1; i < s.Length; i++)
            {
                if (i % 512 == 0) deadline.ThrowIfExpired();
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
        private static string BeautifyJson(string json) => BeautifyJson(json, Deadline.Default);

        private static string BeautifyJson(string json, Deadline deadline)
        {
            deadline.ThrowIfExpired();
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(json);
                return System.Text.Json.JsonSerializer.Serialize(doc.RootElement,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
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
            _hasuserNavigated = false; // ★ 新搜索开始，重置跳转标记
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
                        int pct = (int)((i + 1L) * 100 / total);
                        // ★ 增量更新 _searchResults，让导航按钮在搜索过程中就能使用已有结果
                        if (cnt > 0 && cnt % 500 == 0)
                        {
                            var snapshot = new List<int>(results);
                            BeginInvoke(() => _searchResults = snapshot);
                        }
                        BeginInvoke(() => _lblSearchCount.Text = $"搜索中... {cnt} 个 ({pct}%)");
                    }
                }

                int nearestLine = curLine;
                BeginInvoke(() =>
                {
                    _searchResults = results;
                    if (results.Count > 0)
                    {
                        // ★ 如果用户已经手动跳转过，不再强制跳转到第一个匹配项
                        if (!_hasuserNavigated)
                        {
                            int idx2 = FindNearestIndex(results, nearestLine);
                            NavigateTo(idx2);
                        }
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

        private bool userHasMovedAway()
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
            {
                _searchResults.Clear();
                _searchCts?.Cancel();
                OnSearch();
                return;
            }

            if (_searchResults.Count > 0)
            {
                _hasuserNavigated = true; // ★ 标记用户已手动跳转
                if (userHasMovedAway())
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

            // ★ 关键字没变，但结果为 0 且搜索正在进行中 → 不重复触发
            if (_searchCts != null && !_searchCts.IsCancellationRequested)
            {
                _lblSearchCount.Text = "搜索中，请稍候...";
                return;
            }

            OnSearch();
        }

        private void OnSearchNext(object? s, EventArgs e)
        {
            if (!SearchParamsMatch())
            {
                _searchResults.Clear();
                _searchCts?.Cancel();
                OnSearch();
                return;
            }

            if (_searchResults.Count > 0)
            {
                _hasuserNavigated = true; // ★ 标记用户已手动跳转
                if (userHasMovedAway())
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

            // ★ 关键字没变，但结果为 0 且搜索正在进行中 → 不重复触发
            if (_searchCts != null && !_searchCts.IsCancellationRequested)
            {
                _lblSearchCount.Text = "搜索中，请稍候...";
                return;
            }

            OnSearch();
        }

        private void OnSearchFirst(object? s, EventArgs e)
        {
            if (!SearchParamsMatch())
            {
                _searchResults.Clear();
                _searchCts?.Cancel();
                OnSearch();
                return;
            }
            if (_searchResults.Count > 0)
            {
                _hasuserNavigated = true; // ★ 标记用户已手动跳转
                NavigateTo(0);
                return;
            }

            // ★ 关键字没变，但结果为 0 且搜索正在进行中 → 不重复触发
            if (_searchCts != null && !_searchCts.IsCancellationRequested)
            {
                _lblSearchCount.Text = "搜索中，请稍候...";
                return;
            }

            OnSearch();
        }

        private void OnSearchLast(object? s, EventArgs e)
        {
            if (!SearchParamsMatch())
            {
                _searchResults.Clear();
                _searchCts?.Cancel();
                OnSearch();
                return;
            }
            if (_searchResults.Count > 0)
            {
                _hasuserNavigated = true; // ★ 标记用户已手动跳转
                NavigateTo(_searchResults.Count - 1);
                return;
            }

            // ★ 关键字没变，但结果为 0 且搜索正在进行中 → 不重复触发
            if (_searchCts != null && !_searchCts.IsCancellationRequested)
            {
                _lblSearchCount.Text = "搜索中，请稍候...";
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

        /// <summary>反向导出：排除匹配行，保留不匹配的所有行</summary>
        private void OnReverseExport(object? s, EventArgs e)
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

            // 搜索条件已变更 → 清空旧结果
            if (!SearchParamsMatch())
                _searchResults.Clear();

            // 没有搜索结果 → 先搜索再反向导出
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
                _ = Task.Run(() => DoSearchAndReverseExport(keyword, caseSensitive, useRegex, token));
                return;
            }

            // 已有搜索结果 → 直接反向导出
            DoReverseExport();
        }

        /// <summary>搜索完成后自动反向导出</summary>
        private void DoSearchAndReverseExport(string keyword, bool caseSensitive, bool useRegex, CancellationToken token)
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
                        int pct = (int)((i + 1L) * 100 / total);
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
                        DoReverseExport();
                    }
                    else
                    {
                        _searchCurrentIndex = -1;
                        _lastNavigatedMatchLine = -1;
                        _lblSearchCount.Text = "未找到匹配，无法反向导出";
                        Status("未找到匹配行");
                    }
                });
            }
            catch { }
        }

        /// <summary>反向导出：利用已有的 _searchResults 索引，导出所有不匹配的行</summary>
        private void DoReverseExport()
        {
            string? sourcePath = _filePath;
            if (_searchResults.Count == 0) return;

            // 快照：避免后台线程读取时被并发修改
            var searchResultsSnapshot = _searchResults.ToList();
            var lineOffsets = _lineOffsets!;
            string dir = Path.GetDirectoryName(sourcePath)!;
            string fileName = Path.GetFileName(sourcePath);
            string sourceExt = Path.GetExtension(sourcePath);
            string filterPart = SanitizeFileName(_lastSearchKeyword, 50);
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            Status("正在准备反向导出...");

            _ = Task.Run(() =>
            {
                // 计算反向导出的总行数 = 总行数 - 匹配行数
                int totalLines = lineOffsets.Length;
                int excludedCount = searchResultsSnapshot.Count;
                int exportCount = totalLines - excludedCount;

                // 命名规则：原始文件名.过滤字符_排除.时间戳.源文件后缀
                string safeFileName;
                if (!string.IsNullOrEmpty(filterPart))
                    safeFileName = $"{SanitizeFileName(fileName, 200)}.{filterPart}_排除.{timestamp}{sourceExt}";
                else
                    safeFileName = $"{SanitizeFileName(fileName, 200)}_排除.{timestamp}{sourceExt}";

                string outName = safeFileName;
                string outPath = Path.Combine(dir, outName);

                BeginInvoke(() => Status($"正在反向导出 {exportCount} 行到 {outName}..."));

                try
                {
                    using var sw = new StreamWriter(outPath, false, Encoding.UTF8);
                    int matchIdx = 0;
                    int exported = 0;

                    for (int lineIdx = 0; lineIdx < totalLines; lineIdx++)
                    {
                        // 如果当前行在匹配列表中 → 跳过
                        if (matchIdx < excludedCount && searchResultsSnapshot[matchIdx] == lineIdx)
                        {
                            matchIdx++;
                            continue;
                        }
                        // 不在匹配列表中 → 输出
                        string lineText = ReadLine(lineIdx);
                        sw.WriteLine(lineText);
                        exported++;
                    }
                    sw.Flush();

                    BeginInvoke(() =>
                    {
                        Status($"已反向导出 {exported} 行 -> {outName}");
                        string msg = "反向导出完成！" + Environment.NewLine +
                                     Environment.NewLine +
                                     "文件: " + outName + Environment.NewLine +
                                     "路径: " + dir + Environment.NewLine +
                                     "总行数: " + totalLines.ToString("N0") + Environment.NewLine +
                                     "排除行: " + excludedCount.ToString("N0") + Environment.NewLine +
                                     "导出行: " + exported.ToString("N0");
                        MessageBox.Show(ParentForm, msg, "导出成功");
                        try { FileExported?.Invoke(this, EventArgs.Empty); } catch { }
                    });
                }
                catch (Exception ex)
                {
                    BeginInvoke(() =>
                    {
                        MessageBox.Show(ParentForm, $"反向导出失败: {ex.Message}", "错误");
                        Status("反向导出失败");
                    });
                }
            });
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
                        int pct = (int)((i + 1L) * 100 / total);
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

            int contextLines = (int)_numContextLines.Value;

            // 快照：避免后台线程读取时被并发修改
            var searchResultsSnapshot = _searchResults.ToList();
            var lineOffsets = _lineOffsets!;
            string dir = Path.GetDirectoryName(sourcePath)!;
            string fileName = Path.GetFileName(sourcePath);
            string sourceExt = Path.GetExtension(sourcePath);
            string filterPart = SanitizeFileName(_lastSearchKeyword, 50);
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            Status("正在准备导出...");

            // 所有耗时操作都在后台线程执行，避免 UI 卡死
            _ = Task.Run(() =>
            {
                // ===== 以下全部在后台线程执行 =====

                // 合并重叠的上下文区间，避免重复输出
                var ranges = new List<(int Start, int End)>(searchResultsSnapshot.Count);
                foreach (int lineIdx in searchResultsSnapshot)
                {
                    int start = Math.Max(0, lineIdx - contextLines);
                    int end = Math.Min(lineOffsets.Length - 1, lineIdx + contextLines);
                    ranges.Add((start, end));
                }
                ranges.Sort((a, b) => a.Start.CompareTo(b.Start));

                var mergedRanges = new List<(int Start, int End)>();
                foreach (var range in ranges)
                {
                    if (mergedRanges.Count == 0)
                    {
                        mergedRanges.Add(range);
                    }
                    else
                    {
                        var last = mergedRanges[^1];
                        if (range.Start <= last.End + 1) // 相邻或重叠 → 合并
                        {
                            mergedRanges[^1] = (last.Start, Math.Max(last.End, range.End));
                        }
                        else
                        {
                            mergedRanges.Add(range);
                        }
                    }
                }

                // 计算总行数
                int totalLines = 0;
                foreach (var (s, e) in mergedRanges)
                    totalLines += (e - s + 1);

                // 命名规则：原始文件名.过滤字符.时间戳.源文件后缀
                string safeFileName;
                if (!string.IsNullOrEmpty(filterPart))
                    safeFileName = $"{SanitizeFileName(fileName, 200)}.{filterPart}.{timestamp}{sourceExt}";
                else
                    safeFileName = $"{SanitizeFileName(fileName, 200)}.{timestamp}{sourceExt}";

                string outName = safeFileName;
                string outPath = Path.Combine(dir, outName);

                // 更新状态（回 UI 线程）
                BeginInvoke(() => Status($"正在导出 {totalLines} 行到 {outName}..."));

                try
                {
                    using var sw = new StreamWriter(outPath, false, Encoding.UTF8);
                    foreach (var (start, end) in mergedRanges)
                    {
                        for (int lineIdx = start; lineIdx <= end; lineIdx++)
                        {
                            string lineText = ReadLine(lineIdx);
                            sw.WriteLine(lineText);
                        }
                    }
                    sw.Flush();

                    BeginInvoke(() =>
                    {
                        Status($"已导出 {totalLines} 行 -> {outName}");
                        MessageBox.Show(ParentForm,
                            $"导出完成！\n\n文件: {outName}\n路径: {dir}\n行数: {totalLines:N0}\n匹配行: {searchResultsSnapshot.Count:N0}\n上下行数: {contextLines}",
                            "导出成功");
                        try { FileExported?.Invoke(this, EventArgs.Empty); } catch { }
                    });
                }
                catch (Exception ex)
                {
                    BeginInvoke(() =>
                    {
                        MessageBox.Show(ParentForm, $"导出失败: {ex.Message}", "错误");
                        Status("导出失败");
                    });
                }
            });
        }

        /// <summary>过滤文件名中的非法字符，避免文件创建失败</summary>
        private static string SanitizeFileName(string name, int maxLength = 255)
        {
            if (string.IsNullOrEmpty(name))
                return "";
            
            // Windows 文件名禁止的字符：\/:*?"<>|
            var invalidChars = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
            {
                if (!invalidChars.Contains(c) && !char.IsControl(c))
                    sb.Append(c);
                else
                    sb.Append('_'); // 非法字符替换为下划线
            }
            
            string result = sb.ToString();
            
            // 去除开头和结尾的空格或点号（Windows 限制）
            result = result.TrimStart(' ', '.').TrimEnd(' ', '.');
            
            // 空或全是空格/点的情况
            if (string.IsNullOrEmpty(result))
                return "_";
            
            // 截断到最大长度，但保留结尾的结构
            if (result.Length > maxLength)
                result = result.Substring(0, maxLength).TrimEnd(' ', '.');
            
            return result;
        }

        public void Cleanup()
        {
            _cancel = true;
            try { _timer.Stop(); } catch { }
            try { _timerScroll?.Stop(); } catch { }
            // ★ 提前置 null，防止 OnCache/OnRetrieve 在清理过程中访问已释放的资源
            _indexReady = false;
            _filePath = null;
            _lineOffsets = null;
            _searchCts?.Cancel();
            _searchResults.Clear();
            _searchCurrentIndex = -1;
            _lastNavigatedMatchLine = -1;
            _lastSearchKeyword = "";
            // ★ 先清空 VirtualListSize，大幅减少 ListView 在 Dispose 时的清理开销
            try { _lstContent.VirtualListSize = 0; } catch { }
            // ★ 清空缓存
            _cache.Clear();
            _cS = _cE = -1;
            // ★ 最后关闭文件流
            lock (_readLock) { _readStream?.Dispose(); _readStream = null; }
            if (IsHandleCreated)
                try { BeginInvoke(() => _lblSearchCount.Text = ""); } catch { }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // ★ 不再重复调用 Cleanup()，CloseTab 已经调用过了
                // 只释放 WinForms 控件本身
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