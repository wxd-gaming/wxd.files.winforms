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
        private Button _btnOpenFile = null!, _btnGotoLine = null!, _btnClearCache = null!;
        private TextBox _txtGotoLine = null!;
        private ComboBox _cboSearch = null!;
        private CheckBox _chkCaseSensitive = null!, _chkRegex = null!;
        private Button _btnSearchFirst = null!, _btnSearchPrev = null!, _btnSearchNext = null!, _btnSearchLast = null!, _btnExport = null!, _btnReverseExport = null!;
        private Label _lblStatus = null!;
        private NumericUpDown _numContextLines = null!;
        private Label _lblContextLines = null!;
        // ★ 底部状态栏
        private FlowLayoutPanel _statusBar = null!;
        private Label _lblFileModify = null!, _lblFileSize = null!, _lblFileLines = null!;
        private Label _lblCacheSize = null!, _lblSearchPos = null!, _lblCurLine = null!;
        private HScrollListView _lstContent = null!;
        private System.Windows.Forms.Timer _timer = null!, _timerScroll = null!;
        private TableLayoutPanel _topPanel = null!;

        // 搜索状态
        // ★ 行号统一改为 long，支持超过 21 亿行的超大文件（极端超短行场景）
        // ★ 不再有 _searchResults 字段，导航时直接从 _recentSearches 缓存获取结果快照
        private long _searchCurrentIndex = -1;
        private long _lastNavigatedMatchLine = -1;
        private CancellationTokenSource? _searchCts;
        // ★ 当前生效的搜索 key，格式 "keyword|caseSensitive|useRegex"
        private string _lastSearch = "";
        private bool _hasuserNavigated; // 用户在搜索过程中是否已手动跳转过

        // ★ 控件内最近搜索结果缓存（与 FileCacheStore 独立，不受 5000 万行限制）
        // 解决：大结果不缓存到 FileCacheStore 时，切回来仍能立即跳转
        // 使用 LinkedList 维护插入顺序，便于 LRU 淘汰
        private readonly LinkedList<KeyValuePair<string, List<long>>> _recentSearches = new();
        private long _recentSearchesBytes;
        private const long MaxRecentSearchesBytes = 800L * 1024 * 1024; // 800MB
        private const int MaxRecentSearchesCount = 5;

        private string? _filePath;
        private long _fileLength;
        private volatile long[]? _lineOffsets;
        private volatile bool _indexReady;
        private volatile bool _cancel;
        private const int CHUNK = 4 * 1024 * 1024;
        private const int DISPLAY_LINE_LIMIT = 8 * 1024;          // 列表面板显示截断：8KB
        private const int FULL_LINE_LIMIT = 64 * 1024 * 1024;     // 弹窗原始内容上限：64MB
        // ★ 修复D：Win32 SetScrollInfo 的 nMax 为 int，负值会静默失败。
        // 同时水平滚动范围过大也会导致滚动条精度问题。上限夹紧到 10MB 像素。
        private const int MAX_SCROLL_RANGE = 10 * 1024 * 1024;
        // ★ _idxCount 改为 long，行数也可能超 int.MaxValue（极端超短行场景）
        private long _idxCount;
        private volatile int _idxProg;

        // ★ 并行索引构建相关常量与字段
        // 文件按字节大小动态切块，每块约 100MB，最大 4 线程并行构建行索引
        private const long BLOCK_SIZE = 100L * 1024 * 1024;  // 每块目标大小：100MB
        private const int MAX_BLOCK_COUNT = 128;             // 块数上限
        private const int MAX_PARALLEL = 4;
        private int _blockCount;  // 动态块数，根据文件大小在 BuildIndex 中计算
        // 单块构建结果：块号 → BlockIndexResult
        private readonly ConcurrentDictionary<int, BlockIndexResult> _blockResults = new();
        private CancellationTokenSource? _indexCts;
        private volatile int _completedBlocks;        // 已完成块数（原子计数）
        private long _partialMaxLine;                 // 已连续构建的最大行号（long）
        // ★ 横向滚动条：索引构建期间按字节长度跟踪的最长行（索引完成后一次性测量并设置滚动条，之后不再重算）
        private long _longestLineOffset = -1;         // 最长行起始字节偏移（-1 表示尚未发现）
        private long _longestLineBytes = 0;           // 最长行字节长度
        // 并行搜索：每块的局部结果
        private readonly ConcurrentDictionary<int, List<long>> _blockSearchResults = new();

        /// <summary>单块索引构建结果</summary>
        private sealed class BlockIndexResult
        {
            public int BlockId;
            public long BlockStart;         // 块起始字节（已修正跨块边界）
            public long BlockEnd;           // 块结束字节（不含）
            public List<long> LineStarts;   // 块内每行起始偏移（绝对偏移，long）
            public bool EndsWithNewline;    // 块尾是否以LF换行符结尾
            public long FirstLineStart;     // 块第一行真正起始偏移（已修正跨块）
            public long MaxLineLen;         // 块内最长行字节长度
            public long MaxLineOffset;      // 块内最长行起始偏移（-1 表示无）
        }

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
            // ★ 加载全局搜索历史到下拉框
            try { RefreshSearchHistory(); } catch { }
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            _btnOpenFile = new Button { Text = "打开文件", Size = new Size(90, 28), Location = new Point(10, 6), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(66, 165, 245), ForeColor = Color.White };
            _btnOpenFile.FlatAppearance.BorderSize = 1;
            _btnOpenFile.FlatAppearance.BorderColor = Color.FromArgb(30, 136, 229);
            _btnOpenFile.Click += OnOpen;
            _btnClearCache = new Button { Text = "清理缓存", Size = new Size(80, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(200, 210, 220), ForeColor = Color.FromArgb(50, 70, 90) };
            _btnClearCache.FlatAppearance.BorderSize = 1;
            _btnClearCache.FlatAppearance.BorderColor = Color.FromArgb(160, 180, 200);
            _btnClearCache.Click += OnClearCache;
            _btnGotoLine = new Button { Text = "跳转", Size = new Size(55, 28), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(144, 180, 210), ForeColor = Color.White };
            _btnGotoLine.FlatAppearance.BorderSize = 1;
            _btnGotoLine.FlatAppearance.BorderColor = Color.FromArgb(120, 160, 200);
            _btnGotoLine.Click += OnGoto;
            _txtGotoLine = new TextBox { Size = new Size(120, 28), PlaceholderText = "行号" };
            _txtGotoLine.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OnGoto(s, e); };
            _lblStatus = new Label { Text = "请打开一个文本文件", AutoSize = true, ForeColor = Color.FromArgb(90, 130, 170), Font = new Font("Microsoft YaHei UI", 8.5F) };
            _lblStatus.Margin = new Padding(3, 7, 0, 0);

            _topPanel = new TableLayoutPanel { Dock = DockStyle.Top, Height = 88, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(235, 245, 255) };
            _topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            _topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 1));
            _topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 43));

            var row1 = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10, 8, 0, 0) };
            row1.Controls.AddRange(new Control[] { _btnOpenFile, _btnClearCache, _txtGotoLine, _btnGotoLine });

            var separator = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(173, 209, 237) };

            _cboSearch = new ComboBox { Size = new Size(400, 28), DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems, FlatStyle = FlatStyle.Flat };
            _cboSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OnSearch(); };
            _cboSearch.SelectedIndexChanged += OnSearchHistorySelected;
            // ★ 不监听 TextChanged 清空 _lastSearch：导航按钮会自行比较搜索框 key 与 _lastSearch
            _chkCaseSensitive = new CheckBox { Text = "Aa", AutoSize = true, Checked = false, TextAlign = ContentAlignment.MiddleRight };
            _chkCaseSensitive.Margin = new Padding(3, 6, 0, 0);
            _chkRegex = new CheckBox { Text = ".*", AutoSize = true, Checked = false, TextAlign = ContentAlignment.MiddleRight };
            _chkRegex.Margin = new Padding(3, 6, 0, 0);
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
            row2.Controls.AddRange(new Control[] { _cboSearch, _chkCaseSensitive, _chkRegex, _btnSearchFirst, _btnSearchPrev, _btnSearchNext, _btnSearchLast, _lblContextLines, _numContextLines, _btnExport, _btnReverseExport });

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
            // ★ 修复：OwnerDraw + FullRowSelect 下，空 DrawSubItem 处理器会让系统默认选中色（蓝色）
            // 覆盖 OnDrawItem 已绘制的自定义色。必须显式设置 e.DrawDefault = false，让 OnDrawItem
            // 完全接管 SubItem 绘制，避免系统补绘选中背景。
            _lstContent.DrawSubItem += (s, e) => { e.DrawDefault = false; };
            _lstContent.DrawColumnHeader += (s, e) => { e.DrawDefault = true; };
            _lstContent.MouseDoubleClick += OnDoubleClick;
            _lstContent.KeyDown += OnListViewKeyDown;

            // ★ 底部状态栏：6 个 Label，流式自动排列
            _lblFileModify = MakeStatusLabel("最后修改：-");
            _lblFileSize = MakeStatusLabel("大小：-");
            _lblFileLines = MakeStatusLabel("行数：-");
            _lblCacheSize = MakeStatusLabel("缓存：0B");
            _lblSearchPos = MakeStatusLabel("");
            _lblCurLine = MakeStatusLabel("行：1");
            _statusBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 24,
                Padding = new Padding(8, 3, 8, 2),
                BackColor = Color.FromArgb(235, 245, 255),
                AutoScroll = false,
                WrapContents = false
            };
            _statusBar.Controls.AddRange(new Control[] { _lblFileModify, _lblFileSize, _lblCacheSize, _lblCurLine, _lblFileLines, _lblSearchPos });

            Controls.Add(_lstContent);
            Controls.Add(_statusBar);
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
            if (!File.Exists(path)) { AlertForm.Show(ParentForm, $"文件不存在: {path}", "错误", AlertType.Error); return; }

            _filePath = path;
            _fileLength = new FileInfo(path).Length;
            _cancel = false; _indexReady = false; _idxProg = 0; _idxCount = 0;
            _longestLineOffset = -1; _longestLineBytes = 0;
            _lstContent.HScrollPos = 0;
            _lstContent.MaxContentWidth = _lstContent.ClientRectangle.Width;

            // ★ 检测文件是否变化，变化则清理磁盘缓存（后台执行，不阻塞）
            try { FileCacheStore.CheckFileValid(path, _fileLength, File.GetLastWriteTime(path)); } catch { }

            // ★ 底部状态栏：文件信息（打开时设置一次）
            try
            {
                var fi = new FileInfo(path);
                _lblFileModify.Text = $"最后修改：{fi.LastWriteTime:yyyy-MM-dd HH:mm:ss}";
                _lblFileSize.Text = $"大小：{FormatSize(_fileLength)}";
                _lblFileLines.Text = "行数：索引中...";
                _lblCacheSize.Text = "缓存：0B";
                _lblSearchPos.Text = "";
                _lblCurLine.Text = "行：1";
            }
            catch { }

            // ★ 首屏快速加载：用 _readStream 手动扫描前若干行偏移，通过截断逻辑统一读取
            // 避免 StreamReader.ReadLine() 无截断导致超长行在 TextRenderer.DrawText 中静默失败
            int loaded = 0;
            _readStream?.Dispose();
            _readStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.RandomAccess);

            try
            {
                // ★ 循环扫描文件直到找到足够多的换行符，避免只读前64KB导致大行/多行无法扫描
                var firstOffsets = new List<long>() { 0 };
                var scanBuf = new byte[65536];
                long scanPos = 0;
                while (firstOffsets.Count <= 50 && scanPos < _fileLength)
                {
                    _readStream.Seek(scanPos, SeekOrigin.Begin);
                    int n = _readStream.Read(scanBuf, 0, scanBuf.Length);
                    if (n <= 0) break;
                    for (int i = 0; i < n; i++)
                    {
                        if (scanBuf[i] == GlobalConst.LfChar)
                        {
                            // ★ scanPos + i + 1：scanPos是全局偏移，i是buf内偏移，+1跳过LF本身
                            firstOffsets.Add(scanPos + i + 1);
                            if (firstOffsets.Count > 50) break;
                        }
                    }
                    scanPos += n;
                }
                if (firstOffsets.Count == 1 && _fileLength > 0)
                    firstOffsets.Add(_fileLength);

                for (int idx = 0; idx < firstOffsets.Count - 1 && idx < 50; idx++)
                {
                    long s = firstOffsets[idx], e = firstOffsets[idx + 1];
                    int lineLen = (int)Math.Min(e - s, FULL_LINE_LIMIT);
                    if (lineLen <= 0) continue;
                    _readStream.Seek(s, SeekOrigin.Begin);
                    var lineBuf = new byte[lineLen];
                    int got = 0;
                    while (got < lineLen) { int r = _readStream.Read(lineBuf, got, lineLen - got); if (r <= 0) break; got += r; }
                    string fullText = got > 0
                        ? Encoding.UTF8.GetString(lineBuf, 0, got).TrimEnd(GlobalConst.CrChar, GlobalConst.LfChar)
                        : "";
                    _cache[idx] = TruncateLineText(fullText);
                    loaded++;
                }
            }
            catch { }

            _lstContent.VirtualListSize = Math.Max(loaded, 1);
            _cS = 0; _cE = loaded - 1;

            Status($"前 {loaded} 行已加载，正在构建索引...");
            _lblFileLines.Text = "行数：索引中...";
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

            // ★ 绘制时截断保护：GDI DrawText 无法渲染超过约 32767 字符的字符串，
            // 超长文本会导致静默不绘制（界面空白）。截断为 DISPLAY_LINE_LIMIT 字符。
            if (text.Length > DISPLAY_LINE_LIMIT)
                text = text.Substring(0, DISPLAY_LINE_LIMIT) + "...更多";

            bool selected = _lstContent.SelectedIndices.Contains(e.ItemIndex);
            using var backBrush = new SolidBrush(selected ? Color.FromArgb(66, 165, 245) : _lstContent.BackColor);
            using var textBrush = new SolidBrush(selected ? Color.White : _lstContent.ForeColor);

            e.Graphics.FillRectangle(backBrush, e.Bounds);

            int drawX = e.Bounds.X - hPos + 4;
            int drawY = e.Bounds.Y + (e.Bounds.Height - _lstContent.Font.Height) / 2;
            // ★ 修复D：超长行（>1024字符）输出绘制诊断日志，便于排查水平滚动条问题
            if (text.Length > 1024)
            {
                var msz = TextRenderer.MeasureText(e.Graphics, text, _lstContent.Font,
                    new Size(int.MaxValue, int.MaxValue),
                    TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
                Debug.WriteLine($"[Draw] 行{e.ItemIndex} text.Length={text.Length} 测量宽={msz.Width} drawX={drawX} hPos={hPos} MaxContentWidth={_lstContent.MaxContentWidth}");
            }
            TextRenderer.DrawText(e.Graphics, text, _lstContent.Font,
                new Rectangle(drawX, drawY, int.MaxValue, e.Bounds.Height),
                selected ? Color.White : _lstContent.ForeColor,
                TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);

            if ((e.State & ListViewItemStates.Focused) != 0)
                e.DrawFocusRectangle();
        }

        // ==================== 索引构建 ====================
        // ★ 并行索引构建：文件按字节大小动态切块（每块约 100MB），最大 4 线程并行构建行索引
        // 每块独立 FileStream 扫描 ，完成后原子计数 + 增量 UI 反馈
        // 全部完成后按 0→N-1 顺序合并，去重跨块边界行
        private void BuildIndex()
        {
            try
            {
                _indexCts?.Cancel();
                _indexCts = new CancellationTokenSource();
                var token = _indexCts.Token;

                // ★ 动态计算块数：每块约 100MB，小文件单块，大文件最多 MAX_BLOCK_COUNT 块
                _blockCount = _fileLength <= BLOCK_SIZE ? 1 : (int)Math.Min((_fileLength + BLOCK_SIZE - 1) / BLOCK_SIZE, MAX_BLOCK_COUNT);

                // ★ 小文件快速路径：退化为单块单线程，避免并行调度开销
                if (_blockCount <= 1)
                {
                    BuildIndexSingleThread(token);
                    return;
                }

                var bounds = ComputeBlockBoundaries();
                _blockResults.Clear();
                _completedBlocks = 0;
                _partialMaxLine = 0;

                var options = new ParallelOptions
                {
                    MaxDegreeOfParallelism = MAX_PARALLEL,
                    CancellationToken = token
                };

                try
                {
                    Parallel.For(0, _blockCount, options, blockId =>
                    {
                        if (_cancel) return;
                        var result = BuildBlockIndex(blockId, bounds[blockId].start, bounds[blockId].end, token);
                        if (result == null) return;
                        _blockResults[blockId] = result;

                        // ★ 增量反馈：原子计数 + 从块0开始连续已完成时扩大 VirtualListSize
                        int done = Interlocked.Increment(ref _completedBlocks);
                        UpdatePartialIndex(done);
                    });
                }
                catch (OperationCanceledException) { return; }

                if (_cancel) return;

                // 全部完成 → 顺序合并 64 块
                _lineOffsets = MergeBlockResults();
                _idxCount = _lineOffsets.Length;
                _idxProg = 100;
                _indexReady = true;
            }
            catch (Exception ex) { Debug.WriteLine($"Index: {ex}"); _indexReady = false; }
        }

        /// <summary>小文件快速路径：保留原单线程顺序扫描逻辑</summary>
        private void BuildIndexSingleThread(CancellationToken token)
        {
            try
            {
                var est = Math.Max(_fileLength / 40L, 1024L);
                if (est > int.MaxValue) est = int.MaxValue;
                var off = new List<long>((int)est) { 0 };

                using var fs = new FileStream(_filePath!, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, CHUNK, FileOptions.SequentialScan);
                var buf = new byte[CHUNK];
                long pos = 0; int n;
                long lineStart = 0;        // 当前行起始偏移
                long maxLen = 0, maxOff = -1;  // 本文件最长行（按字节长度）
                while ((n = fs.Read(buf, 0, buf.Length)) > 0)
                {
                    if (_cancel || token.IsCancellationRequested) return;
                    for (int i = 0; i < n; i++, pos++)
                    {
                        if (buf[i] == GlobalConst.LfChar)
                        {
                            off.Add(pos + 1);
                            long lineLen = pos - lineStart;  // 不含 LF 的行字节长度
                            if (lineLen > maxLen) { maxLen = lineLen; maxOff = lineStart; }
                            lineStart = pos + 1;
                        }
                    }
                    _idxCount = off.Count - 1;
                    _idxProg = (int)(pos * 100 / _fileLength);
                }
                // 文件末尾无换行符的最后一行也要参与最长行统计
                if (_fileLength > lineStart)
                {
                    long lineLen = _fileLength - lineStart;
                    if (lineLen > maxLen) { maxLen = lineLen; maxOff = lineStart; }
                }
                _longestLineOffset = maxOff;
                _longestLineBytes = maxLen;
                if (off.Count > 1 && off[off.Count - 1] >= _fileLength) off.RemoveAt(off.Count - 1);
                _lineOffsets = off.ToArray();
                _idxCount = _lineOffsets.Length;
                _idxProg = 100; _indexReady = true;
            }
            catch (Exception ex) { Debug.WriteLine($"Index: {ex}"); _indexReady = false; }
        }

        /// <summary>计算 _blockCount 块的字节边界，按字节大小均匀分配（余数均匀分到前 N 块）</summary>
        private (long start, long end)[] ComputeBlockBoundaries()
        {
            var bounds = new (long, long)[_blockCount];
            long baseSize = _fileLength / _blockCount;
            long remainder = _fileLength % _blockCount;
            long acc = 0;
            for (int i = 0; i < _blockCount; i++)
            {
                long size = baseSize + (i < remainder ? 1 : 0);
                bounds[i] = (acc, acc + size);
                acc += size;
            }
            return bounds;
        }

        /// <summary>从 position 往前回退查找上一个  的位置，返回该  之后的位置（即本块第一行真正起始）</summary>
        private long FindPrevNewlineEnd(long position, FileStream fs)
        {
            // 回退最多 MAX_LINE 字节查找上一个换行符
            long back = Math.Min(position, FULL_LINE_LIMIT);
            long searchStart = position - back;
            fs.Position = searchStart;
            var buf = new byte[back];
            int got = 0;
            while (got < buf.Length)
            {
                int r = fs.Read(buf, got, buf.Length - got);
                if (r <= 0) break;
                got += r;
            }
            // 从后往前找最后一个 

            for (int i = got - 1; i >= 0; i--)
            {
                if (buf[i] == GlobalConst.LfChar)
                    return searchStart + i + 1;
            }
            // 没找到换行符，使用 searchStart（块边界）
            return searchStart;
        }

        /// <summary>单块索引构建（独立 FileStream，处理跨块边界）</summary>
        private BlockIndexResult? BuildBlockIndex(int blockId, long blockStart, long blockEnd, CancellationToken token)
        {
            var lineStarts = new List<long>();
            long firstLineStart;

            using var fs = new FileStream(_filePath!, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite, 65536, FileOptions.RandomAccess);

            // 块0：第一行起始为 0
            // 块N>0：回退找上一个 ，其 +1 为本块第一行真正起始（跨块边界行修正）
            if (blockId > 0)
                firstLineStart = FindPrevNewlineEnd(blockStart, fs);
            else
                firstLineStart = 0;

            fs.Position = firstLineStart;
            var buf = new byte[CHUNK];
            long pos = firstLineStart;
            bool endsWithNewline = false;
            long lineStart = firstLineStart;         // 当前行起始偏移
            long maxLen = 0, maxOff = -1;            // 块内最长行（按字节长度）

            while (pos < blockEnd)
            {
                if (token.IsCancellationRequested || _cancel) return null;
                int toRead = (int)Math.Min(buf.Length, blockEnd - pos);
                int n = fs.Read(buf, 0, toRead);
                if (n <= 0) break;

                for (int i = 0; i < n; i++)
                {
                    if (buf[i] == GlobalConst.LfChar)
                    {
                        long nextLineStart = pos + i + 1;
                        // 统计本行长度（不含 LF）；跨块首行完整落入本块统计，块尾部分行仅部分长度，取全局 max 仍正确
                        long lineLen = pos + i - lineStart;
                        if (lineLen > maxLen) { maxLen = lineLen; maxOff = lineStart; }
                        if (nextLineStart < blockEnd)
                            lineStarts.Add(nextLineStart);
                        else
                            endsWithNewline = true;
                        lineStart = nextLineStart;
                    }
                }
                pos += n;
            }
            // 块尾无换行符的部分行也参与统计（其完整长度会在下一块被完整统计，此处仅为兜底）
            if (pos > lineStart)
            {
                long lineLen = pos - lineStart;
                if (lineLen > maxLen) { maxLen = lineLen; maxOff = lineStart; }
            }

            return new BlockIndexResult
            {
                BlockId = blockId,
                BlockStart = firstLineStart,
                BlockEnd = blockEnd,
                LineStarts = lineStarts,
                EndsWithNewline = endsWithNewline,
                FirstLineStart = firstLineStart,
                MaxLineLen = maxLen,
                MaxLineOffset = maxOff
            };
        }

        /// <summary>按 0→63 顺序合并块结果，去重跨块边界行</summary>
        private long[] MergeBlockResults()
        {
            var merged = new List<long>(_lineOffsets?.Length ?? 65536) { 0 }; // 第一行起始为 0

            for (int i = 0; i < _blockCount; i++)
            {
                if (!_blockResults.TryGetValue(i, out var br)) continue;
                foreach (var off in br.LineStarts)
                {
                    // 去重：跨块边界行偏移可能与 merged 最后一项相同
                    if (merged.Count == 0 || merged[merged.Count - 1] != off)
                        merged.Add(off);
                }
            }

            // 修正：如果最后一块尾部不是 LF，移除最后一个虚拟行偏移
            if (merged.Count > 1 && merged[merged.Count - 1] >= _fileLength)
                merged.RemoveAt(merged.Count - 1);

            // ★ 汇总各块最长行，取全局最大（跨块首行完整落入其起始块统计，取 max 正确）
            long maxLen = 0, maxOff = -1;
            for (int i = 0; i < _blockCount; i++)
            {
                if (!_blockResults.TryGetValue(i, out var br)) continue;
                if (br.MaxLineLen > maxLen) { maxLen = br.MaxLineLen; maxOff = br.MaxLineOffset; }
            }
            _longestLineOffset = maxOff;
            _longestLineBytes = maxLen;

            return merged.ToArray();
        }

        /// <summary>计算从块0开始连续已完成的行数，用于增量 UI 反馈</summary>
        private long CountContinuousBuiltLines()
        {
            long total = 1; // 第 0 行起始为 0
            long lastOffset = 0;
            for (int i = 0; i < _blockCount; i++)
            {
                if (!_blockResults.TryGetValue(i, out var br)) break;
                foreach (var off in br.LineStarts)
                {
                    if (off > lastOffset)
                    {
                        total++;
                        lastOffset = off;
                    }
                }
            }
            return total;
        }

        /// <summary>增量更新部分索引：从块0开始连续完成时扩大 VirtualListSize，让用户可提前查看</summary>
        private void UpdatePartialIndex(int doneBlocks)
        {
            long continuous = CountContinuousBuiltLines();
            if (continuous <= _partialMaxLine) return;
            _partialMaxLine = continuous;
            _idxCount = continuous;
            _idxProg = doneBlocks * 100 / _blockCount;

            // 回 UI 线程扩大 VirtualListSize
            try
            {
                BeginInvoke(() =>
                {
                    if (_indexReady) return;
                    int newSize = (int)Math.Min(_partialMaxLine, int.MaxValue);
                    if (newSize > _lstContent.VirtualListSize)
                    {
                        _lstContent.VirtualListSize = newSize;
                        _lblFileLines.Text = $"行数：已构建 {continuous:N0} 行 ({doneBlocks}/{_blockCount}块 {_idxProg}%)";
                    }
                });
            }
            catch { }
        }

        private void Tick()
        {
            if (_indexReady)
            {
                _timer.Stop();
                Status($"{FormatSize(_fileLength)} | {_idxCount:N0} 行");
                // ★ 底部状态栏：行数 + 当前缓存大小
                _lblFileLines.Text = $"行数：{_idxCount:N0}";
                RefreshCacheSizeStatus();
                // ★ long → int 安全转换（WinForms ListView 限制为 int）
                _lstContent.VirtualListSize = (int)Math.Min(_idxCount, int.MaxValue);
                // ★ 索引完成后一次性测量最长行并设置横向滚动条，之后纵向滚动不再重算（与 Sublime 一致）
                SetHorizontalScrollOnce();
                _timerScroll.Start();
                try { _lstContent.RedrawItems(0, Math.Min(_lstContent.VirtualListSize - 1, _lstContent.ClientRectangle.Height / Math.Max(_lstContent.Font.Height + 6, 1)), true); } catch { }
                return;
            }
            // ★ 索引构建过程中实时刷新底部状态栏的行数和缓存大小
            // 部分索引可用时显示已构建行数 + 块进度（UpdatePartialIndex 会单独更新文案）
            if (_partialMaxLine == 0)
                _lblFileLines.Text = $"行数：{_idxCount:N0} ({_idxProg}%)";
            RefreshCacheSizeStatus();
        }

        private void UpdateScrollPos()
        {
            try
            {
                int line = (_lstContent.TopItem?.Index ?? 0) + 1;
                // ★ 底部状态栏：当前行（滚动时更新）
                _lblCurLine.Text = $"行：第{line:N0}行";
                // ★ 同时定时刷新缓存大小（缓存随时可能因搜索而增长）
                RefreshCacheSizeStatus();
            }
            catch { }
        }

        /// <summary>对行文本进行截断处理（与 ReadLine/预加载 共用），超长行截断为 DISPLAY_LINE_LIMIT 字符</summary>
        private static string TruncateLineText(string fullText)
        {
            if (fullText.Length > DISPLAY_LINE_LIMIT)
                return fullText.Substring(0, DISPLAY_LINE_LIMIT) + "...更多";
            return fullText;
        }

        /// <summary>索引完成后一次性设置横向滚动条：测量索引期发现的最长行（已截断 8K），只设置一次，之后纵向滚动不再重算</summary>
        private void SetHorizontalScrollOnce()
        {
            try
            {
                int clientW = _lstContent.ClientRectangle.Width;
                string? text = null;
                int lineIdx = -1;

                // 定位最长行：通过行偏移数组二分查找行号，再 ReadLine 读取（内部已截断 DISPLAY_LINE_LIMIT）
                var off = _lineOffsets;
                if (off != null && _longestLineOffset >= 0)
                {
                    int idx = Array.BinarySearch(off, _longestLineOffset);
                    if (idx >= 0)
                    {
                        lineIdx = idx;
                        text = ReadLine(idx);
                    }
                }
                // 兜底：最长行定位失败时，退化为缓存中最长文本
                if (text == null && !_cache.IsEmpty)
                {
                    var kv = _cache.OrderByDescending(p => p.Value?.Length ?? 0).FirstOrDefault();
                    text = kv.Value;
                    lineIdx = kv.Key;
                }

                int maxW = clientW;
                if (!string.IsNullOrEmpty(text))
                {
                    using var g = _lstContent.CreateGraphics();
                    // TextRenderer.MeasureText 与绘制路径 TextRenderer.DrawText 同栈，结果一致（避免 GDI+ 超长返回 Infinity）
                    var sz = TextRenderer.MeasureText(g, text, _lstContent.Font,
                        new Size(int.MaxValue, int.MaxValue),
                        TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
                    int candidate = sz.Width + 10;
                    if (candidate > 0)
                        maxW = Math.Max(clientW, candidate);
                    else
                        Debug.WriteLine($"[Measure] 最长行{lineIdx} 异常测量 candidate={candidate} text.Length={text.Length} sz.Width={sz.Width}");
                }
                maxW = Math.Min(maxW, MAX_SCROLL_RANGE);

                Debug.WriteLine($"[Measure] SetHorizontalScrollOnce MaxContentWidth {_lstContent.MaxContentWidth} -> {maxW} (最长行={lineIdx}, 字节数={_longestLineBytes})");
                _lstContent.MaxContentWidth = maxW;
                _lstContent.InitHScroll();
            }
            catch (Exception ex) { Debug.WriteLine($"[Measure] SetHorizontalScrollOnce 异常: {ex.Message}"); }
        }

        private string ReadLine(long idx)
        {
            var off = _lineOffsets;
            if (off == null || _readStream == null || idx >= off.Length) return "";
            long s = off[idx], e = (idx + 1 < off.Length) ? off[idx + 1] : _fileLength;
            int lineLen = (int)Math.Min(e - s, FULL_LINE_LIMIT);
            if (lineLen <= 0) return "";
            lock (_readLock)
            {
                try
                {
                    _readStream.Seek(s, SeekOrigin.Begin);
                    var buf = new byte[lineLen]; int got = 0;
                    while (got < lineLen) { int r = _readStream.Read(buf, got, lineLen - got); if (r <= 0) break; got += r; }
                    string fullText = got > 0
                        ? Encoding.UTF8.GetString(buf, 0, got).TrimEnd(GlobalConst.CrChar, GlobalConst.LfChar)
                        : "";
                    return TruncateLineText(fullText);
                }
                catch { return "[err]"; }
            }
        }

        /// <summary>使用指定 FileStream 读取行（无锁，供并行搜索/导出使用，避免 _readLock 争抢）</summary>
        private string ReadLineWithStream(FileStream fs, long idx)
        {
            var off = _lineOffsets;
            if (off == null || idx >= off.Length) return "";
            long s = off[idx], e = (idx + 1 < off.Length) ? off[idx + 1] : _fileLength;
            int lineLen = (int)Math.Min(e - s, FULL_LINE_LIMIT);
            if (lineLen <= 0) return "";
            try
            {
                fs.Seek(s, SeekOrigin.Begin);
                var buf = new byte[lineLen]; int got = 0;
                while (got < lineLen) { int r = fs.Read(buf, got, lineLen - got); if (r <= 0) break; got += r; }
                string fullText = got > 0
                    ? Encoding.UTF8.GetString(buf, 0, got).TrimEnd(GlobalConst.CrChar, GlobalConst.LfChar)
                    : "";
                if (fullText.Length > DISPLAY_LINE_LIMIT)
                    return fullText.Substring(0, DISPLAY_LINE_LIMIT) + "...更多";
                return fullText;
            }
            catch { return ""; }
        }

        /// <summary>读取行的完整内容（无截断），上限 64MB</summary>
        /// <param name="idx">行号</param>
        /// <param name="truncated">是否因超过 64MB 被截断</param>
        /// <returns>完整行文本，超过 64MB 返回截断版本；读取失败返回 null</returns>
        private string? ReadLineFull(long idx, out bool truncated)
        {
            truncated = false;
            var off = _lineOffsets;
            if (off == null || _readStream == null || idx >= off.Length) { truncated = false; return null; }
            long s = off[idx], e = (idx + 1 < off.Length) ? off[idx + 1] : _fileLength;
            long lineLen = e - s;
            if (lineLen <= 0) { truncated = false; return null; }
            if (lineLen > FULL_LINE_LIMIT)
            {
                truncated = true;
                lineLen = FULL_LINE_LIMIT;
            }
            int len = (int)lineLen;
            lock (_readLock)
            {
                try
                {
                    _readStream.Seek(s, SeekOrigin.Begin);
                    var buf = new byte[len]; int got = 0;
                    while (got < len) { int r = _readStream.Read(buf, got, len - got); if (r <= 0) break; got += r; }
                    return got > 0
                        ? Encoding.UTF8.GetString(buf, 0, got).TrimEnd(GlobalConst.CrChar, GlobalConst.LfChar)
                        : null;
                }
                catch { return null; }
            }
        }

        private void OnRetrieve(object? s, RetrieveVirtualItemEventArgs e)
        {
            e.Item = new ListViewItem(_filePath != null && _cache.TryGetValue(e.ItemIndex, out var v) ? v : "");
        }

        private void OnCache(object? s, CacheVirtualItemsEventArgs e)
        {
            if (_filePath == null) return;
            // ★ 部分索引可用时也允许构建缓存，但需检查是否在已构建范围
            // _indexReady 为 true 时 _lineOffsets 完整可用
            // _indexReady 为 false 但 _partialMaxLine > 0 时，只允许访问 [0, _partialMaxLine)
            if (!_indexReady)
            {
                if (_partialMaxLine <= 0 || _lineOffsets == null) return;
            }
            else if (_lineOffsets == null) return;
            int start = Math.Max(0, e.StartIndex), end = Math.Min(_lstContent.VirtualListSize - 1, e.EndIndex);
            if (start > end) return;
            if (_cS >= 0 && start >= _cS && end <= _cE) return;
            for (int i = _cS; i <= _cE; i++) if (i < start || i > end) _cache.TryRemove(i, out _);
            for (int i = start; i <= end; i++) if (!_cache.ContainsKey(i)) _cache[i] = ReadLine(i);
            _cS = start; _cE = end;
        }

        private void Status(string s) { if (_lblStatus != null) try { BeginInvoke(() => _lblStatus.Text = s); } catch { } }
        static string FormatSize(long b) { string[] u = ["B", "KB", "MB", "GB", "TB"]; double v = b; int o = 0; while (v >= 1024 && o < 4) { o++; v /= 1024; } return $"{v:0.##}{u[o]}"; }

        /// <summary>创建底部状态栏的 Label</summary>
        private static Label MakeStatusLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Margin = new Padding(0, 2, 16, 0),
                ForeColor = Color.FromArgb(90, 130, 170),
                Font = new Font("Microsoft YaHei UI", 8F),
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        /// <summary>刷新底部状态栏的缓存大小</summary>
        private void RefreshCacheSizeStatus()
        {
            try
            {
                long bytes = string.IsNullOrEmpty(_filePath) ? 0 : FileCacheStore.GetMemorySize(_filePath);
                _lblCacheSize.Text = $"缓存：{FormatSize(bytes)}";
            }
            catch { }
        }

        private void OnListViewKeyDown(object? s, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectedLines();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void CopySelectedLines()
        {
            var indices = _lstContent.SelectedIndices;
            if (indices.Count == 0) return;

            var sb = new StringBuilder();
            foreach (int idx in indices)
            {
                string text = ReadLine(idx);
                sb.AppendLine(text);
            }

            if (sb.Length > 0)
            {
                string text = sb.ToString().TrimEnd('\r', '\n');
                try { Clipboard.SetText(text); }
                catch
                {
                    try { Thread.Sleep(100); Clipboard.SetDataObject(text, true); }
                    catch { }
                }
            }
        }

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

            // 读取完整行内容用于弹窗（不截断），并计算内容大小
            bool isFullTruncated = false;
            string? fullLine = ReadLineFull(idx, out isFullTruncated);
            if (string.IsNullOrEmpty(fullLine)) fullLine = originalText;
            string sizeLabel = FormatSize(Encoding.UTF8.GetByteCount(fullLine));

            var f = new Form
            {
                Text = $"行 {idx + 1} — {sizeLabel}", Size = new Size(950, 550), StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = true, MaximizeBox = true, ShowIcon = false, ShowInTaskbar = false
            };

            // 如果超过 64MB，弹出警告
            if (isFullTruncated)
            {
                MessageBox.Show("该行内容超过 64MB，已截断显示前 64MB。" + GlobalConst.NewLine + $"行号：{idx + 1}", "内容过大", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

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
                try { rtb.Rtf = rtf ?? fullLine; }
                catch { try { rtb.Text = fullLine; } catch { } }
            };
            btnRaw.Click += (_, _) => {
                SetActive(btnRaw, btnFormat);
                rtb.WordWrap = true;
                rtb.ScrollBars = RichTextBoxScrollBars.Vertical;
                rtb.Text = fullLine;
            };

            topPanel.Controls.Add(btnFormat);
            topPanel.Controls.Add(btnRaw);
            f.Controls.Add(rtb);
            f.Controls.Add(topPanel);

            f.Show();

            // 格式化在后台线程完成，完成后自动显示 RTF
            try { formattedText = await Task.Run(() => BuildFormattedText(fullLine)).ConfigureAwait(true); }
            catch { formattedText = fullLine; }
            if (f.IsDisposed) return;

            var autoDeadline = Deadline.FromSeconds(5.0);
            string? rtf = null;
            try { rtf = await Task.Run(() => BuildRtfForLine(formattedText, autoDeadline)).ConfigureAwait(true); }
            catch { /* 解析失败，显示原始文本 */ }
            if (f.IsDisposed) return;
            try { rtb.Rtf = rtf ?? fullLine; }
            catch { try { rtb.Text = fullLine; } catch { } }
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
            // ★ 行号支持 long（超过 int 的超大文件）
            if (!long.TryParse(_txtGotoLine.Text, out long ln) || ln < 1) { AlertForm.Show(ParentForm, "请输入有效行号", "提示", AlertType.Information); return; }
            // ★ long → int 安全转换（ListView 限制为 int）
            int virtualSize = _lstContent.VirtualListSize;
            if (ln > virtualSize) { AlertForm.Show(ParentForm, $"超出范围，共 {virtualSize:N0} 行", "提示", AlertType.Information); return; }
            int idx = (int)(ln - 1); _lstContent.EnsureVisible(idx); _lstContent.SelectedIndices.Clear(); _lstContent.SelectedIndices.Add(idx); _lstContent.Focus();
        }

        // ==================== 搜索功能 ====================

        private long GetCurrentLine()
        {
            try
            {
                if (_lstContent.SelectedIndices.Count > 0)
                    return _lstContent.SelectedIndices[0];
                return _lstContent.TopItem?.Index ?? 0;
            }
            catch { return 0; }
        }

        private static long FindNearestIndex(List<long> results, long currentLine)
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

        /// <summary>根据搜索框和选项计算当前搜索 key</summary>
        private string CurrentSearchKey()
        {
            return $"{_cboSearch.Text.Trim()}|{_chkCaseSensitive.Checked}|{_chkRegex.Checked}";
        }

        /// <summary>
        /// 写入控件内最近搜索缓存（不受 FileCacheStore 的 5000 万行限制）。
        /// 采用 LRU：按总字节数 800MB 和条数 5 的上限淘汰最旧条目。
        /// </summary>
        private void PutRecentSearch(string key, List<long> results)
        {
            // 先移除同 key 旧条目（避免重复）
            var node = _recentSearches.First;
            while (node != null)
            {
                if (node.Value.Key == key)
                {
                    _recentSearchesBytes -= EstimateSearchResultBytes(node.Value.Value);
                    _recentSearches.Remove(node);
                    break;
                }
                node = node.Next;
            }

            // 加入链表尾部（最新）
            _recentSearches.AddLast(new KeyValuePair<string, List<long>>(key, results));
            _recentSearchesBytes += EstimateSearchResultBytes(results);

            // LRU 淘汰：超过条数或字节数时从头部移除最旧
            while (_recentSearches.Count > MaxRecentSearchesCount || _recentSearchesBytes > MaxRecentSearchesBytes)
            {
                if (_recentSearches.Count == 0) break;
                var oldest = _recentSearches.First!;
                _recentSearchesBytes -= EstimateSearchResultBytes(oldest.Value.Value);
                _recentSearches.RemoveFirst();
            }
        }

        /// <summary>查询控件内最近搜索缓存，命中则返回结果（不移除，LRU 顺序不变）</summary>
        private List<long>? TryGetRecentSearch(string key)
        {
            foreach (var kv in _recentSearches)
            {
                if (kv.Key == key)
                    return kv.Value;
            }
            return null;
        }

        /// <summary>估算 List&lt;long&gt; 占用字节数（每 long 8 字节 + 列表开销）</summary>
        private static long EstimateSearchResultBytes(List<long> list)
        {
            // List<long> 内部 long 数组：count * 8，加上 List 对象本身约 32 字节
            return (long)list.Count * 8 + 64;
        }

        private async void OnSearch()
        {
            await EnsureSearchReadyAsync();
        }

        /// <summary>
        /// 确保当前搜索条件已就绪：
        /// 1. 若搜索框 key == _lastSearch → 直接使用（命中或正在搜索中）
        /// 2. 否则查控件内最近搜索缓存 _recentSearches（不受大小限制）
        /// 3. 再查 FileCacheStore 缓存（受 5000 万行限制）
        /// 4. 都未命中 → 启动后台搜索
        /// 返回 true 表示结果已就绪可导航；false 表示正在后台搜索中。
        /// </summary>
        private async Task<bool> EnsureSearchReadyAsync()
        {
            string keyword = _cboSearch.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                _lblSearchPos.Text = "";
                return false;
            }
            if (_lineOffsets == null)
            {
                _lblSearchPos.Text = "请等待索引完成";
                return false;
            }

            bool caseSensitive = _chkCaseSensitive.Checked;
            bool useRegex = _chkRegex.Checked;
            string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";

            // 1. 当前搜索结果仍然有效
            if (currentKey == _lastSearch)
            {
                // ★ 正在后台搜索中 → 返回 true（即使搜索未完成，也能使用已有部分结果导航）
                if (_searchCts != null && !_searchCts.IsCancellationRequested)
                {
                    var results = TryGetRecentSearch(currentKey);
                    return results != null && results.Count > 0;
                }
                return true; // 0 匹配也算就绪
            }

            // 2a. 先查控件内最近搜索缓存（同步，无 IO，不受大小限制）
            var recent = TryGetRecentSearch(currentKey);
            if (recent != null)
            {
                _searchCts = null;
                _lastSearch = currentKey;
                _searchCurrentIndex = -1;
                _lastNavigatedMatchLine = -1;
                _hasuserNavigated = false;

                if (recent.Count > 0)
                {
                    long idx2 = FindNearestIndex(recent, GetCurrentLine());
                    NavigateTo(recent, idx2);
                }
                else
                {
                    _lblSearchPos.Text = "搜索完成，0 个匹配（缓存）";
                }
                return true;
            }

            // 2b. 再查 FileCacheStore 缓存（后台读盘，不阻塞 UI）
            var cached = await FileCacheStore.TryGetSearchResultAsync(_filePath!, currentKey);

            // ★ await 回来后，用户可能已切换关键词；若已切换，丢弃本次结果
            if (CurrentSearchKey() != currentKey)
                return false;

            if (cached != null)
            {
                _searchCts = null; // 缓存命中，无后台任务
                _lastSearch = currentKey;
                _searchCurrentIndex = -1;
                _lastNavigatedMatchLine = -1;
                _hasuserNavigated = false;
                // ★ FileCacheStore 命中也写入控件内最近缓存，方便下次切换
                PutRecentSearch(currentKey, cached);

                if (cached.Count > 0)
                {
                    long idx2 = FindNearestIndex(cached, GetCurrentLine());
                    NavigateTo(cached, idx2);
                }
                else
                {
                    _lblSearchPos.Text = "搜索完成，0 个匹配（缓存）";
                }
                return true;
            }

            // 3. 缓存未命中 → 后台搜索
            _searchCts?.Cancel();
            _searchCts = new CancellationTokenSource();
            var token = _searchCts.Token;
            _lastSearch = currentKey;
            _lblSearchPos.Text = "搜索中...";
            _searchCurrentIndex = -1;
            _lastNavigatedMatchLine = -1;
            _hasuserNavigated = false;
            // ★ 向缓存中插入空列表作为搜索容器（DoSearch 中会增量追加）
            PutRecentSearch(currentKey, new List<long>());
            // ★ 实际发起搜索时才写入历史记录、刷新下拉框
            SearchHistoryStore.Add(keyword, caseSensitive, useRegex);
            RefreshSearchHistory();
            long curLine = GetCurrentLine();
            _ = Task.Run(() => DoSearch(keyword, caseSensitive, useRegex, token, curLine));
            return false;
        }

        /// <summary>刷新搜索下拉框的历史记录（不影响 _lastSearch，不影响搜索框输入）</summary>
        private void RefreshSearchHistory()
        {
            // ★ 临时解除 SelectedIndexChanged 事件，避免 Items 操作触发递归
            _cboSearch.SelectedIndexChanged -= OnSearchHistorySelected;
            try
            {
                var history = SearchHistoryStore.GetHistory();
                string current = _cboSearch.Text;
                _cboSearch.Items.Clear();
                foreach (var h in history) _cboSearch.Items.Add(h);
                _cboSearch.Text = current;
                _cboSearch.SelectionStart = current.Length;
                _cboSearch.SelectionLength = 0;
            }
            finally
            {
                _cboSearch.SelectedIndexChanged += OnSearchHistorySelected;
            }
        }

        /// <summary>从历史记录中选择搜索项时，恢复 checkbox 状态并触发搜索</summary>
        private void OnSearchHistorySelected(object? s, EventArgs e)
        {
            if (_cboSearch.SelectedItem is SearchHistoryItem item)
            {
                // ★ 恢复 checkbox 状态
                _chkCaseSensitive.Checked = item.CaseSensitive;
                _chkRegex.Checked = item.UseRegex;
                // ★ 触发搜索
                OnSearch();
            }
        }

        /// <summary>清理当前文件的搜索缓存</summary>
        private void OnClearCache(object? s, EventArgs e)
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                AlertForm.Show(ParentForm, "当前未打开文件", "提示", AlertType.Information);
                return;
            }
            FileCacheStore.ClearFile(_filePath);
            _searchCurrentIndex = -1;
            _lastNavigatedMatchLine = -1;
            _lastSearch = "";
            // ★ 清空控件内最近搜索缓存
            _recentSearches.Clear();
            _recentSearchesBytes = 0;
            _lblSearchPos.Text = "";
            // ★ 清理缓存后刷新底部状态栏的缓存大小
            RefreshCacheSizeStatus();
            AlertForm.Show(ParentForm, "已清理当前文件的搜索缓存", "提示", AlertType.Information);
        }

        private void DoSearch(string keyword, bool caseSensitive, bool useRegex, CancellationToken token, long curLine)
        {
            bool useParallel = false;
            try
            {
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

                long total = _lineOffsets!.Length;
                useParallel = total > 50000;
                string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";

                if (useParallel)
                {
                    // ★ 增量并行搜索：Task.Run + SemaphoreSlim 控制并发
                    // 每块完成后立即增量合并到缓存的列表，第一块完成即自动跳转
                    var lineBounds = ComputeLineBoundsForBlocks();
                    var semaphore = new SemaphoreSlim(MAX_PARALLEL);
                    var cts = CancellationTokenSource.CreateLinkedTokenSource(token);
                    var cancelToken = cts.Token;
                    int totalBlocks = _blockCount;
                    int completedCount = 0;
                    bool firstBlockHandled = false;

                    for (int blockId = 0; blockId < totalBlocks; blockId++)
                    {
                        int block = blockId; // 捕获闭包
                        semaphore.Wait(cancelToken);

                        Task.Run(() =>
                        {
                            try
                            {
                                if (cancelToken.IsCancellationRequested) return;

                                var (startLine, endLine) = lineBounds[block];
                                var local = new List<long>();
                                using var fs = new FileStream(_filePath!, FileMode.Open, FileAccess.Read,
                                    FileShare.ReadWrite, 65536, FileOptions.RandomAccess);

                                for (long i = startLine; i <= endLine; i++)
                                {
                                    if (cancelToken.IsCancellationRequested) return;
                                    string line = ReadLineWithStream(fs, i);
                                    if (!string.IsNullOrEmpty(line) && matcher(line))
                                        local.Add(i);
                                }

                                if (cancelToken.IsCancellationRequested) return;

                                // ★ 每块完成后，增量合并到缓存的列表（在 UI 线程上）
                                BeginInvoke(() =>
                                {
                                    if (_lastSearch != currentKey) return;

                                    // 从缓存获取列表引用，追加结果
                                    var cacheList = TryGetRecentSearch(currentKey);
                                    if (cacheList == null) return;

                                    if (local.Count > 0)
                                    {
                                        foreach (var lineNum in local)
                                        {
                                            int pos = cacheList.BinarySearch(lineNum);
                                            if (pos < 0) pos = ~pos;
                                            cacheList.Insert(pos, lineNum);
                                        }
                                    }

                                    int done = Interlocked.Increment(ref completedCount);

                                    // ★ 第一块完成且有匹配 → 自动跳转到最近匹配项
                                    if (!firstBlockHandled && done == 1)
                                    {
                                        firstBlockHandled = true;
                                        if (cacheList.Count > 0 && !_hasuserNavigated)
                                        {
                                            long idx2 = FindNearestIndex(cacheList, curLine);
                                            NavigateTo(cacheList, idx2);
                                        }
                                    }

                                    // 更新进度
                                    _lblSearchPos.Text = $"搜索中... {cacheList.Count} 个匹配 ({done}/{totalBlocks}块)";

                                    // ★ 全部完成 → 收尾并释放 semaphore
                                    if (done >= totalBlocks)
                                    {
                                        OnSearchComplete(keyword, caseSensitive, useRegex);
                                        try { semaphore.Dispose(); } catch { }
                                    }
                                });
                            }
                            finally
                            {
                                semaphore.Release();
                            }
                        }, cancelToken);
                    }
                }
                else
                {
                    // 单线程快速路径（增量模式）
                    for (long i = 0; i < total; i++)
                    {
                        if (token.IsCancellationRequested) return;
                        string line = ReadLine(i);
                        if (!string.IsNullOrEmpty(line) && matcher(line))
                        {
                            var cacheList = TryGetRecentSearch(currentKey);
                            if (cacheList != null)
                            {
                                cacheList.Add(i);
                            }
                        }
                        if (i % 500 == 0)
                        {
                            int cnt = TryGetRecentSearch(currentKey)?.Count ?? 0;
                            int pct = (int)((i + 1L) * 100 / total);
                            // ★ 有匹配就跳转（首次）
                            if (cnt > 0 && !_hasuserNavigated)
                            {
                                BeginInvoke(() =>
                                {
                                    if (_lastSearch != currentKey) return;
                                    if (!_hasuserNavigated)
                                    {
                                        var list = TryGetRecentSearch(currentKey);
                                        if (list != null && list.Count > 0)
                                        {
                                            long idx2 = FindNearestIndex(list, curLine);
                                            NavigateTo(list, idx2);
                                        }
                                    }
                                });
                            }
                            BeginInvoke(() => _lblSearchPos.Text = $"搜索中... {cnt} 个 ({pct}%)");
                        }
                    }

                    // 单线程完成
                    BeginInvoke(() =>
                    {
                        string currentKey2 = $"{keyword}|{caseSensitive}|{useRegex}";
                        if (_lastSearch != currentKey2) return;
                        OnSearchComplete(keyword, caseSensitive, useRegex);
                    });
                }
            }
            catch { }
            finally
            {
                // ★★★ Bug 修复：搜索结束后清理 _searchCts，避免导航按钮误判"搜索中" ★★★
                // ★ 注意：并行分支中所有 Task.Run 已通过 fire-and-forget 启动，
                // DoSearch 在启动所有任务后立即返回，所以 finally 在这里执行。
                // 对于并行分支，_searchCts 的清理由 OnSearchComplete 负责（在所有块完成后）。
                // 对于单线程分支，这里清理 _searchCts 是正确的。
                if (!useParallel && _searchCts != null && _searchCts.Token == token)
                {
                    var cts = _searchCts;
                    _searchCts = null;
                    try { cts.Dispose(); } catch { }
                }
            }
        }

        /// <summary>搜索全部完成后的收尾逻辑（写入缓存、更新状态）</summary>
        private void OnSearchComplete(string keyword, bool caseSensitive, bool useRegex)
        {
            string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";
            if (_lastSearch != currentKey) return;

            // ★ 清理 _searchCts（并行分支在 OnSearchComplete 中清理）
            if (_searchCts != null)
            {
                var cts = _searchCts;
                _searchCts = null;
                try { cts.Dispose(); } catch { }
            }

            // ★ 从缓存获取结果列表（搜索过程中已增量写入）
            var results = TryGetRecentSearch(currentKey);
            if (results == null) return;

            // ★ 写入 FileCacheStore 持久化缓存（后台写盘）
            _ = FileCacheStore.PutSearchResultAsync(_filePath!, currentKey, results);

            if (results.Count > 0)
            {
                // ★ 如果用户已经手动跳转过，不再强制跳转
                if (!_hasuserNavigated)
                {
                    long idx2 = FindNearestIndex(results, GetCurrentLine());
                    NavigateTo(results, idx2);
                }
                else
                {
                    _lblSearchPos.Text = $"搜索完成，共 {results.Count} 个匹配";
                }
            }
            else
            {
                _searchCurrentIndex = -1;
                _lastNavigatedMatchLine = -1;
                _lblSearchPos.Text = "搜索完成，0 个匹配";
            }
        }

        /// <summary>二分 _lineOffsets 计算每块对应的行范围 [startLine, endLine]</summary>
        private (long startLine, long endLine)[] ComputeLineBoundsForBlocks()
        {
            var bounds = ComputeBlockBoundaries();
            var lineBounds = new (long startLine, long endLine)[_blockCount];
            long[] off = _lineOffsets!;
            long totalLines = off.Length;
            for (int i = 0; i < _blockCount; i++)
            {
                long blockStart = bounds[i].start;
                long blockEnd = bounds[i].end;
                // 二分查找：第一个 >= blockStart 的行
                long startLine = LowerBound(off, blockStart);
                // 最后一个 < blockEnd 的行
                long endLine = LowerBound(off, blockEnd) - 1;
                if (endLine >= totalLines) endLine = totalLines - 1;
                if (endLine < startLine) endLine = startLine - 1; // 空块
                lineBounds[i] = (startLine, endLine);
            }
            // 第一块从第 0 行开始；最后一块到最后一行结束
            if (_blockCount > 0)
            {
                lineBounds[0] = (0, lineBounds[0].endLine);
                lineBounds[_blockCount - 1] = (lineBounds[_blockCount - 1].startLine, totalLines - 1);
            }
            return lineBounds;
        }

        /// <summary>二分查找：返回第一个 >= target 的索引，找不到返回 arr.Length</summary>
        private static long LowerBound(long[] arr, long target)
        {
            long lo = 0, hi = arr.Length;
            while (lo < hi)
            {
                long mid = (lo + hi) / 2;
                if (arr[mid] < target) lo = mid + 1;
                else hi = mid;
            }
            return lo;
        }

        private void NavigateTo(List<long> results, long index)
        {
            _searchCurrentIndex = index;
            long line = results[(int)index];
            // ★ long → int 安全转换（ListView 限制为 int）
            int listViewIdx = (int)Math.Min(line, int.MaxValue);
            _lstContent.EnsureVisible(listViewIdx);
            _lstContent.SelectedIndices.Clear();
            _lstContent.SelectedIndices.Add(listViewIdx);
            _lstContent.Focus();
            _lastNavigatedMatchLine = line;
            // ★ 底部状态栏：搜索位置 + 当前行
            _lblSearchPos.Text = $"搜索：第{index + 1}/{results.Count}个";
            _lblCurLine.Text = $"行：第{line + 1:N0}行";
        }

        private bool userHasMovedAway()
        {
            if (_lastNavigatedMatchLine < 0) return true;
            if (_lstContent.SelectedIndices.Count > 0 && _lstContent.SelectedIndices[0] != _lastNavigatedMatchLine)
                return true;
            int top = _lstContent.TopItem?.Index ?? 0;
            if (Math.Abs((long)top - _lastNavigatedMatchLine) > (_lstContent.ClientRectangle.Height / (_lstContent.Font.Height + 6)))
                return true;
            return false;
        }

        /// <summary>从缓存获取当前关键字的搜索结果快照（拷贝），供导航使用</summary>
        private List<long>? GetSearchResultsSnapshot()
        {
            string currentKey = CurrentSearchKey();
            if (string.IsNullOrEmpty(_cboSearch.Text.Trim()) || currentKey != _lastSearch)
                return null;
            var list = TryGetRecentSearch(currentKey);
            if (list == null || list.Count == 0)
                return null;
            return list.ToList(); // ★ 快照：防止导航过程中列表被后台搜索修改
        }

        private async void OnSearchPrev(object? s, EventArgs e)
        {
            bool ready = await EnsureSearchReadyAsync();
            var results = GetSearchResultsSnapshot();
            // ★ 搜索中但已有部分结果时，仍然可以导航
            if (!ready && results == null)
            {
                if (_searchCts != null && !_searchCts.IsCancellationRequested)
                    _lblSearchPos.Text = "搜索中，请稍候...";
                return;
            }
            if (results == null || results.Count == 0)
            {
                AlertForm.Show(ParentForm, "无匹配项", "提示", AlertType.Information);
                return;
            }

            _hasuserNavigated = true; // ★ 标记用户已手动跳转
            if (userHasMovedAway())
            {
                long curLine = GetCurrentLine();
                long idx2 = FindNearestIndex(results, curLine);
                if (results[(int)idx2] >= curLine)
                    idx2--;
                if (idx2 < 0)
                {
                    AlertForm.Show(ParentForm, "已到顶部，没有更多匹配项", "提示", AlertType.Information);
                    return;
                }
                NavigateTo(results, idx2);
            }
            else
            {
                if (_searchCurrentIndex <= 0)
                {
                    AlertForm.Show(ParentForm, "已到顶部，没有更多匹配项", "提示", AlertType.Information);
                    return;
                }
                NavigateTo(results, _searchCurrentIndex - 1);
            }
        }

        private async void OnSearchNext(object? s, EventArgs e)
        {
            bool ready = await EnsureSearchReadyAsync();
            var results = GetSearchResultsSnapshot();
            // ★ 搜索中但已有部分结果时，仍然可以导航
            if (!ready && results == null)
            {
                if (_searchCts != null && !_searchCts.IsCancellationRequested)
                    _lblSearchPos.Text = "搜索中，请稍候...";
                return;
            }
            if (results == null || results.Count == 0)
            {
                AlertForm.Show(ParentForm, "无匹配项", "提示", AlertType.Information);
                return;
            }

            _hasuserNavigated = true; // ★ 标记用户已手动跳转
            if (userHasMovedAway())
            {
                long curLine = GetCurrentLine();
                long idx2 = FindNearestIndex(results, curLine);
                if (results[(int)idx2] <= curLine)
                    idx2++;
                if (idx2 >= results.Count)
                {
                    AlertForm.Show(ParentForm, "已到底部，没有更多匹配项", "提示", AlertType.Information);
                    return;
                }
                NavigateTo(results, idx2);
            }
            else
            {
                if (_searchCurrentIndex >= results.Count - 1)
                {
                    AlertForm.Show(ParentForm, "已到底部，没有更多匹配项", "提示", AlertType.Information);
                    return;
                }
                NavigateTo(results, _searchCurrentIndex + 1);
            }
        }

        private async void OnSearchFirst(object? s, EventArgs e)
        {
            bool ready = await EnsureSearchReadyAsync();
            var results = GetSearchResultsSnapshot();
            // ★ 搜索中但已有部分结果时，仍然可以导航
            if (!ready && results == null)
            {
                if (_searchCts != null && !_searchCts.IsCancellationRequested)
                    _lblSearchPos.Text = "搜索中，请稍候...";
                return;
            }
            if (results == null || results.Count == 0)
            {
                AlertForm.Show(ParentForm, "无匹配项", "提示", AlertType.Information);
                return;
            }

            _hasuserNavigated = true; // ★ 标记用户已手动跳转
            NavigateTo(results, 0);
        }

        private async void OnSearchLast(object? s, EventArgs e)
        {
            bool ready = await EnsureSearchReadyAsync();
            var results = GetSearchResultsSnapshot();
            // ★ 搜索中但已有部分结果时，仍然可以导航
            if (!ready && results == null)
            {
                if (_searchCts != null && !_searchCts.IsCancellationRequested)
                    _lblSearchPos.Text = "搜索中，请稍候...";
                return;
            }
            if (results == null || results.Count == 0)
            {
                AlertForm.Show(ParentForm, "无匹配项", "提示", AlertType.Information);
                return;
            }

            _hasuserNavigated = true; // ★ 标记用户已手动跳转
            NavigateTo(results, results.Count - 1);
        }

        private void OnExport(object? s, EventArgs e)
        {
            string? sourcePath = _filePath;
            if (string.IsNullOrEmpty(sourcePath))
            {
                AlertForm.Show(ParentForm, "请先打开文件", "提示", AlertType.Information);
                return;
            }

            string keyword = _cboSearch.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                AlertForm.Show(ParentForm, "请先输入搜索关键词", "提示", AlertType.Information);
                return;
            }
            if (_lineOffsets == null)
            {
                AlertForm.Show(ParentForm, "请等待索引完成", "提示", AlertType.Information);
                return;
            }

            string currentKey = CurrentSearchKey();

            // 搜索条件已变更 → 重新搜索并导出
            if (currentKey != _lastSearch)
            {
                StartSearchForExport(keyword, false);
                return;
            }

            // 检查缓存中是否有结果
            var results = TryGetRecentSearch(currentKey);
            if (results == null || results.Count == 0)
            {
                StartSearchForExport(keyword, false);
                return;
            }

            // 已有搜索结果且条件匹配 → 直接导出
            DoExport(results);
        }

        /// <summary>启动搜索，完成后自动导出/反向导出</summary>
        private void StartSearchForExport(string keyword, bool reverse)
        {
            _searchCts?.Cancel();
            _searchCts = new CancellationTokenSource();
            var token = _searchCts.Token;
            bool caseSensitive = _chkCaseSensitive.Checked;
            bool useRegex = _chkRegex.Checked;
            string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";
            _lastSearch = currentKey;
            _lblSearchPos.Text = "搜索中...";
            _searchCurrentIndex = -1;
            _lastNavigatedMatchLine = -1;
            // ★ 向缓存中插入空列表作为搜索容器
            PutRecentSearch(currentKey, new List<long>());

            Status($"正在搜索 \"{keyword}\"...");
            if (reverse)
                _ = Task.Run(() => DoSearchAndReverseExport(keyword, caseSensitive, useRegex, token));
            else
                _ = Task.Run(() => DoSearchAndExport(keyword, caseSensitive, useRegex, token));
        }

        /// <summary>反向导出：排除匹配行，保留不匹配的所有行</summary>
        private void OnReverseExport(object? s, EventArgs e)
        {
            string? sourcePath = _filePath;
            if (string.IsNullOrEmpty(sourcePath))
            {
                AlertForm.Show(ParentForm, "请先打开文件", "提示", AlertType.Information);
                return;
            }

            string keyword = _cboSearch.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                AlertForm.Show(ParentForm, "请先输入搜索关键词", "提示", AlertType.Information);
                return;
            }
            if (_lineOffsets == null)
            {
                AlertForm.Show(ParentForm, "请等待索引完成", "提示", AlertType.Information);
                return;
            }

            string currentKey = CurrentSearchKey();

            // 搜索条件已变更 → 重新搜索
            if (currentKey != _lastSearch)
            {
                StartSearchForExport(keyword, true);
                return;
            }

            // 检查缓存中是否有结果
            var results = TryGetRecentSearch(currentKey);
            if (results == null || results.Count == 0)
            {
                StartSearchForExport(keyword, true);
                return;
            }

            // 已有搜索结果 → 直接反向导出
            DoReverseExport(results);
        }

        /// <summary>搜索完成后自动反向导出</summary>
        private void DoSearchAndReverseExport(string keyword, bool caseSensitive, bool useRegex, CancellationToken token)
        {
            try
            {
                string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";
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

                long total = _lineOffsets!.Length;
                for (long i = 0; i < total; i++)
                {
                    if (token.IsCancellationRequested) return;
                    string line = ReadLine(i);
                    if (!string.IsNullOrEmpty(line) && matcher(line))
                    {
                        var cacheList = TryGetRecentSearch(currentKey);
                        if (cacheList != null) cacheList.Add(i);
                    }
                    if (i % 500 == 0)
                    {
                        int cnt = TryGetRecentSearch(currentKey)?.Count ?? 0;
                        int pct = (int)((i + 1L) * 100 / total);
                        BeginInvoke(() => _lblSearchPos.Text = $"搜索中... {cnt} 个 ({pct}%)");
                    }
                }

                BeginInvoke(() =>
                {
                    // ★ 如果用户在搜索期间切换了关键词，丢弃本次结果
                    if (_lastSearch != currentKey) return;

                    var results = TryGetRecentSearch(currentKey);
                    if (results == null) return;

                    // ★ 写入 FileCacheStore 持久化缓存（后台写盘）
                    _ = FileCacheStore.PutSearchResultAsync(_filePath!, currentKey, results);
                    if (results.Count > 0)
                    {
                        long idx2 = FindNearestIndex(results, GetCurrentLine());
                        NavigateTo(results, idx2);
                        DoReverseExport(results);
                    }
                    else
                    {
                        _searchCurrentIndex = -1;
                        _lastNavigatedMatchLine = -1;
                        _lblSearchPos.Text = "未找到匹配，无法反向导出";
                        Status("未找到匹配行");
                    }

                    // ★ 清理 _searchCts
                    if (_searchCts != null)
                    {
                        var cts = _searchCts;
                        _searchCts = null;
                        try { cts.Dispose(); } catch { }
                    }
                });
            }
            catch { }
        }

        /// <summary>反向导出：利用已有的搜索结果索引，导出所有不匹配的行</summary>
        private void DoReverseExport(List<long> results)
        {
            string? sourcePath = _filePath;
            if (results.Count == 0) return;

            // 快照：避免后台线程读取时被并发修改
            var searchResultsSnapshot = results.ToList();
            var lineOffsets = _lineOffsets!;
            string dir = Path.GetDirectoryName(sourcePath)!;
            string fileName = Path.GetFileName(sourcePath);
            string sourceExt = Path.GetExtension(sourcePath);
            // ★ 从 _lastSearch 中取关键词部分（格式 "keyword|caseSensitive|useRegex"）
            string filterPart = SanitizeFileName(_lastSearch.Split('|')[0], 50);
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            Status("正在准备反向导出...");

            _ = Task.Run(() =>
            {
                // 计算反向导出的总行数 = 总行数 - 匹配行数
                long totalLines = lineOffsets.Length;
                int excludedCount = searchResultsSnapshot.Count;
                long exportCount = totalLines - excludedCount;

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
                    long exported = 0;

                    for (long lineIdx = 0; lineIdx < totalLines; lineIdx++)
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
                        AlertForm.Show(ParentForm, msg, "导出成功", AlertType.Information);
                        try { FileExported?.Invoke(this, EventArgs.Empty); } catch { }
                    });
                }
                catch (Exception ex)
                {
                    BeginInvoke(() =>
                    {
                        AlertForm.Show(ParentForm, $"反向导出失败: {ex.Message}", "错误", AlertType.Error, exception: ex);
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
                string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";
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

                long total = _lineOffsets!.Length;
                for (long i = 0; i < total; i++)
                {
                    if (token.IsCancellationRequested) return;
                    string line = ReadLine(i);
                    if (!string.IsNullOrEmpty(line) && matcher(line))
                    {
                        var cacheList = TryGetRecentSearch(currentKey);
                        if (cacheList != null) cacheList.Add(i);
                    }
                    if (i % 500 == 0)
                    {
                        int cnt = TryGetRecentSearch(currentKey)?.Count ?? 0;
                        int pct = (int)((i + 1L) * 100 / total);
                        BeginInvoke(() => _lblSearchPos.Text = $"搜索中... {cnt} 个 ({pct}%)");
                    }
                }

                BeginInvoke(() =>
                {
                    // ★ 如果用户在搜索期间切换了关键词，丢弃本次结果
                    if (_lastSearch != currentKey) return;

                    var results = TryGetRecentSearch(currentKey);
                    if (results == null) return;

                    // ★ 写入 FileCacheStore 持久化缓存（后台写盘）
                    _ = FileCacheStore.PutSearchResultAsync(_filePath!, currentKey, results);
                    if (results.Count > 0)
                    {
                        long idx2 = FindNearestIndex(results, GetCurrentLine());
                        NavigateTo(results, idx2);
                        DoExport(results);
                    }
                    else
                    {
                        _searchCurrentIndex = -1;
                        _lastNavigatedMatchLine = -1;
                        _lblSearchPos.Text = "未找到匹配，无法导出";
                        Status("未找到匹配行");
                    }

                    // ★ 清理 _searchCts
                    if (_searchCts != null)
                    {
                        var cts = _searchCts;
                        _searchCts = null;
                        try { cts.Dispose(); } catch { }
                    }
                });
            }
            catch { }
        }

        private void DoExport(List<long> results)
        {
            string? sourcePath = _filePath;
            if (results.Count == 0) return;

            int contextLines = (int)_numContextLines.Value;

            // 快照：避免后台线程读取时被并发修改
            var searchResultsSnapshot = results.ToList();
            var lineOffsets = _lineOffsets!;
            string dir = Path.GetDirectoryName(sourcePath)!;
            string fileName = Path.GetFileName(sourcePath);
            string sourceExt = Path.GetExtension(sourcePath);
            // ★ 从 _lastSearch 中取关键词部分（格式 "keyword|caseSensitive|useRegex"）
            string filterPart = SanitizeFileName(_lastSearch.Split('|')[0], 50);
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            Status("正在准备导出...");

            // 所有耗时操作都在后台线程执行，避免 UI 卡死
            _ = Task.Run(() =>
            {
                // ===== 以下全部在后台线程执行 =====

                // 合并重叠的上下文区间，避免重复输出
                var ranges = new List<(long Start, long End)>(searchResultsSnapshot.Count);
                foreach (long lineIdx in searchResultsSnapshot)
                {
                    long start = Math.Max(0, lineIdx - contextLines);
                    long end = Math.Min(lineOffsets.Length - 1, lineIdx + contextLines);
                    ranges.Add((start, end));
                }
                ranges.Sort((a, b) => a.Start.CompareTo(b.Start));

                var mergedRanges = new List<(long Start, long End)>();
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
                long totalLines = 0;
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
                        for (long lineIdx = start; lineIdx <= end; lineIdx++)
                        {
                            string lineText = ReadLine(lineIdx);
                            sw.WriteLine(lineText);
                        }
                    }
                    sw.Flush();

                    BeginInvoke(() =>
                    {
                        Status($"已导出 {totalLines} 行 -> {outName}");
                        AlertForm.Show(ParentForm,
                            $"导出完成！" + GlobalConst.NewLine + GlobalConst.NewLine + $"文件: {outName}" + GlobalConst.NewLine + $"路径: {dir}" + GlobalConst.NewLine + $"行数: {totalLines:N0}" + GlobalConst.NewLine + $"匹配行: {searchResultsSnapshot.Count:N0}" + GlobalConst.NewLine + $"上下行数: {contextLines}",
                            "导出成功", AlertType.Information);
                        try { FileExported?.Invoke(this, EventArgs.Empty); } catch { }
                    });
                }
                catch (Exception ex)
                {
                    BeginInvoke(() =>
                    {
                        AlertForm.Show(ParentForm, $"导出失败: {ex.Message}", "错误", AlertType.Error, exception: ex);
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
            // ★ 取消并行索引构建任务
            try { _indexCts?.Cancel(); } catch { }
            // ★ 提前保存文件路径，用于卸载内存缓存（_filePath 接下来会被置 null）
            string? oldPath = _filePath;
            // ★ 提前置 null，防止 OnCache/OnRetrieve 在清理过程中访问已释放的资源
            _indexReady = false;
            _partialMaxLine = 0;
            _completedBlocks = 0;
            _filePath = null;
            _lineOffsets = null;
            _blockResults.Clear();
            _blockSearchResults.Clear();
            _searchCts?.Cancel();
            _searchCurrentIndex = -1;
            _lastNavigatedMatchLine = -1;
            _lastSearch = "";
            // ★ 清空控件内最近搜索缓存
            _recentSearches.Clear();
            _recentSearchesBytes = 0;
            // ★ 关闭页签时卸载内存缓存，磁盘保留供下次打开
            if (!string.IsNullOrEmpty(oldPath))
            {
                try { FileCacheStore.UnloadFile(oldPath); } catch { }
            }
            // ★ 先清空 VirtualListSize，大幅减少 ListView 在 Dispose 时的清理开销
            try { _lstContent.VirtualListSize = 0; } catch { }
            // ★ 清空缓存
            _cache.Clear();
            _cS = _cE = -1;
            // ★ 最后关闭文件流
            lock (_readLock) { _readStream?.Dispose(); _readStream = null; }
            if (IsHandleCreated)
                try
                {
                    // ★ 只清空动态状态 Label，不动文件元信息（修改日期/大小/行数）
                    // 因为 Open 方法会调用 Cleanup，若用 BeginInvoke 异步清空会覆盖 Open 刚设置的文件信息
                    // 同步设置即可，Label.Text 无需消息循环
                    _lblSearchPos.Text = "";
                    _lblCacheSize.Text = "缓存：0B";
                    _lblCurLine.Text = "行：1";
                }
                catch { }
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
                _indexCts?.Dispose();
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
        private const int WM_PAINT = 0x000F;
        private const int WM_LBUTTONDBLCLK = 0x0203;
        private const int SB_HORZ = 0;
        // ★ 修复D：Win32 SetScrollInfo 的 nMax 为 int，负值会静默失败。
        // 与外层 LargeFileReaderControl.MAX_SCROLL_RANGE 保持一致。
        private const int MAX_SCROLL_RANGE = 10 * 1024 * 1024;

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
        [DllImport("user32.dll")]
        private static extern bool ShowScrollBar(IntPtr hWnd, int wBar, bool bShow);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private static readonly IntPtr HWND_TOP = IntPtr.Zero;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_FRAMECHANGED = 0x0020;
        private const uint SWP_NOACTIVATE = 0x0010;

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int HScrollPos { get; set; }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int MaxContentWidth { get; set; } = 1000;

        // ★ 修复B：WM_SIZE 去抖字段，避免重复异步触发 InitHScroll
        private int _lastClientWidthForSize = -1;
        private int _lastMaxContentWidthForSize = -1;

        public event EventHandler<MouseEventArgs>? ItemDoubleClick;

        public void InitHScroll()
        {
            if (!IsHandleCreated) return;
            // ★ 修复D：夹紧到 MAX_SCROLL_RANGE，防止 nMax 负值或超大值导致 SetScrollInfo 失败
            int clampedMax = Math.Min(Math.Max(MaxContentWidth, 1), MAX_SCROLL_RANGE);
            if (clampedMax != MaxContentWidth)
            {
                Debug.WriteLine($"[HScroll] MaxContentWidth 夹紧 {MaxContentWidth} -> {clampedMax}");
                MaxContentWidth = clampedMax;
            }
            // ★ 修复：ListView 在 Details 模式下水平滚动条由**列宽**驱动，而非 SCROLLINFO 的 nMax。
            // 若列宽 != nMax，ListView 在后续 WM_PAINT/WM_NCCALCSIZE 中会按列宽重新计算并隐藏滚动条。
            // 必须让列宽 = 内容宽度（clampedMax），与 nMax 严格对齐。
            // 仅当内容宽度 < ClientWidth 时（无需滚动），列宽兜底为 ClientWidth 以填满可视区。
            int colWidth = clampedMax < ClientRectangle.Width ? ClientRectangle.Width : clampedMax;
            if (Columns.Count > 0)
                Columns[0].Width = colWidth;
            // nMax 与列宽对齐
            int maxRange = colWidth;
            int page = ClientRectangle.Width;
            int maxScroll = Math.Max(0, maxRange - page + 1);
            // ★ 修复A：保留当前滚动位置，仅夹紧到合法范围，不再强制清零
            int newPos = HScrollPos;
            if (newPos < 0) newPos = 0;
            if (newPos > maxScroll) newPos = maxScroll;
            var si = new SCROLLINFO();
            si.cbSize = Marshal.SizeOf(si);
            // SIF_RANGE(0x0001) | SIF_PAGE(0x0002) | SIF_POS(0x0004) = 0x0007
            si.fMask = 0x0007;
            si.nMin = 0;
            si.nMax = maxRange;
            si.nPage = page;
            si.nPos = newPos;
            SetScrollInfo(Handle, SB_HORZ, ref si, true);
            // ★ v3 修复：ListView 在 OwnerDraw+VirtualMode+HeaderStyle.None 下会在 WM_NCCALCSIZE/WM_PAINT
            // 内部用列宽覆盖 SetScrollInfo 的 nMax。此处立即用 GetScrollInfo 验证是否生效，
            // 若 nMax 被改回（通常是改回 ClientWidth），则双重 SetScrollInfo 并触发 SWP_FRAMECHANGED
            // 强制 ListView 重算非客户区（含滚动条）。
            var siCheck = new SCROLLINFO();
            siCheck.cbSize = Marshal.SizeOf(siCheck);
            siCheck.fMask = 0x0001 | 0x0002 | 0x0004; // SIF_RANGE | SIF_PAGE | SIF_POS
            GetScrollInfo(Handle, SB_HORZ, ref siCheck);
            if (siCheck.nMax != maxRange || siCheck.nPage != page)
            {
                Debug.WriteLine($"[HScroll] 检测到 nMax 被覆盖: 期望 nMax={maxRange} nPage={page}, 实际 nMax={siCheck.nMax} nPage={siCheck.nPage} -> 双重重设");
                // 再次设置，并触发非客户区重算
                SetScrollInfo(Handle, SB_HORZ, ref si, true);
                SetWindowPos(Handle, HWND_TOP, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
                // 重算后再校验一次
                GetScrollInfo(Handle, SB_HORZ, ref siCheck);
                Debug.WriteLine($"[HScroll] 双重重设后: nMax={siCheck.nMax} nPage={siCheck.nPage} nPos={siCheck.nPos}");
            }
            // ★ 修复E：ListView 在 OwnerDraw+VirtualMode+HeaderStyle.None 下会自动隐藏水平滚动条，
            // 此处强制重新显示，确保用户可见可用
            if (maxRange > page)
                ShowScrollBar(Handle, SB_HORZ, true);
            HScrollPos = newPos;
            Debug.WriteLine($"[HScroll] nMax={maxRange} nPage={page} nPos={newPos} maxScroll={maxScroll} shown={maxRange > page}");
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
                // ★ 修复B：去抖判断，仅在 ClientWidth/MaxContentWidth 真正变化时才刷新滚动条
                int curW = ClientRectangle.Width;
                if (curW != _lastClientWidthForSize || MaxContentWidth != _lastMaxContentWidthForSize)
                {
                    _lastClientWidthForSize = curW;
                    _lastMaxContentWidthForSize = MaxContentWidth;
                    BeginInvoke(new Action(InitHScroll));
                }
                return;
            }
            if (m.Msg == WM_PAINT)
            {
                base.WndProc(ref m);
                // ★ v3 修复：兜底重设完整 SCROLLINFO（不仅是 ShowScrollBar）。
                // ListView 在 OwnerDraw+VirtualMode+HeaderStyle.None 下，即使 SetScrollInfo
                // 设置了正确的范围，WM_PAINT 内部仍可能用列宽覆盖 nMax 并隐藏滚动条。
                // 此处在每次 WM_PAINT 后检查：若内容需要滚动但 nMax 不正确，则重设并强制显示。
                int contentW = MaxContentWidth;
                int clientW = ClientRectangle.Width;
                if (contentW > clientW)
                {
                    var siPaint = new SCROLLINFO();
                    siPaint.cbSize = Marshal.SizeOf(siPaint);
                    siPaint.fMask = 0x0001 | 0x0002 | 0x0004; // SIF_RANGE | SIF_PAGE | SIF_POS
                    GetScrollInfo(Handle, SB_HORZ, ref siPaint);
                    int maxScroll = Math.Max(0, contentW - clientW + 1);
                    int pos = HScrollPos;
                    if (pos < 0) pos = 0;
                    if (pos > maxScroll) pos = maxScroll;
                    if (siPaint.nMax != contentW || siPaint.nPage != clientW || siPaint.nPos != pos)
                    {
                        siPaint.nMin = 0;
                        siPaint.nMax = contentW;
                        siPaint.nPage = clientW;
                        siPaint.nPos = pos;
                        siPaint.fMask = 0x0007;
                        SetScrollInfo(Handle, SB_HORZ, ref siPaint, true);
                        HScrollPos = pos;
                        Debug.WriteLine($"[HScroll] WM_PAINT 兜底重设: nMax={contentW} nPage={clientW} nPos={pos}");
                    }
                    ShowScrollBar(Handle, SB_HORZ, true);
                }
                return;
            }
            base.WndProc(ref m);
        }
    }
}