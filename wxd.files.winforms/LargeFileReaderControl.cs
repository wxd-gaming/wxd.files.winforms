using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
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
        private ThemeButton _btnOpenFile = null!, _btnGotoLine = null!, _btnClearCache = null!;
        private ThemeField _fldGoto = null!;
        private TextBox _txtGotoLine = null!;
        private ThemeComboBox _cboSearch = null!;
        // 可编辑下拉框在 Flat 下浅色没有边框、深色只有系统灰边，外面套一圈主题描边统一观感
        private SearchBoxHost _searchHost = null!;
        private ThemeCheckBox _chkCaseSensitive = null!, _chkRegex = null!;
        private ThemeButton _btnExport = null!, _btnReverseExport = null!, _btnCancelExport = null!;
        // 首条/上一个/下一个/末条合成一颗分段按钮，比四颗并排方按钮更像一个整体
        private ThemeSegmented _segNav = null!;
        private Label _lblStatus = null!;
        // 原生 NumericUpDown 的灰色微调块不跟随皮肤，改成和其他输入框同款的数字主题框
        private ThemeField _fldContext = null!;
        private TextBox _txtContext = null!;
        private Label _lblContextLines = null!;
        // ★ 底部状态栏
        private Panel _statusBar = null!;
        private Panel _statusDivider = null!;
        private FlowLayoutPanel _statusFlow = null!;
        private Label _lblFileModify = null!, _lblFileSize = null!, _lblFileLines = null!;
        private Label _lblCacheSize = null!, _lblSearchPos = null!, _lblCurLine = null!;
        private LineListView _lstContent = null!;
        private System.Windows.Forms.Timer _timer = null!, _timerScroll = null!;
        private Panel _toolPanel = null!;
        private Panel _toolDivider = null!;
        // 选中行详情框：列表只画截断的前 8KB，这里用 ReadLineFull 展示完整行，只读但可选中复制，
        // 省去「想看全/想复制就得双击弹窗」这一步。只留文本框本身，不另占表头行；
        // 行号/大小/截断提示放悬停 tooltip。贴底 docking 从列表区借高度，外框和状态栏都不长高。
        private Panel _detailPanel = null!;
        private Panel _detailDivider = null!;
        private TextBox _txtDetail = null!;
        private ToolTip _detailTip = null!;

        // 搜索状态
        // ★ 行号统一改为 long，支持超过 21 亿行的超大文件（极端超短行场景）
        // ★ 不再有 _searchResults 字段，导航时直接从 _recentSearches 缓存获取结果快照
        private long _searchCurrentIndex = -1;
        private long _lastNavigatedMatchLine = -1;
        private CancellationTokenSource? _searchCts;
        // 导出/反向导出共用：关掉页签或换文件时取消，半截输出会被删掉而不是报「导出成功」
        private CancellationTokenSource? _exportCts;
        // 导出并发闸：正向/反向同一时刻只跑一个，后来的提示等待而不是取消覆盖上一个
        private int _exportRunning;
        // ★ 当前生效的搜索 key，格式 "keyword|caseSensitive|useRegex"
        private string _lastSearch = "";
        private bool _hasuserNavigated; // 用户在搜索过程中是否已手动跳转过

        // ★ 控件内最近搜索结果缓存（与 FileCacheStore 独立，不受 5000 万行限制）
        // 解决：大结果不缓存到 FileCacheStore 时，切回来仍能立即跳转
        // 使用 LinkedList 维护插入顺序，便于 LRU 淘汰
        private readonly LinkedList<KeyValuePair<string, SearchResultSet>> _recentSearches = new();
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
        private const int DETAIL_LINE_LIMIT = 256 * 1024;         // 详情框显示上限：256KB，再大的行交给双击弹窗，免得 EDIT 控件卡死
        // ★ 修复D：Win32 SetScrollInfo 的 nMax 为 int，负值会静默失败。
        // 同时水平滚动范围过大也会导致滚动条精度问题。上限夹紧到 10MB 像素。
        private const int MAX_SCROLL_RANGE = 10 * 1024 * 1024;
        // ★ 文件切割相关常量
        private const int TAIL_EMPTY_LINES = 2;
        private const long SPLIT_LINE_THRESHOLD = 1_000_000_000L; // 10亿行
        private const long SPLIT_LINES_PER_FILE = 500_000_000L;   // 每5亿行
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
        // 并行搜索：行数超过这个值才值得按块铺开扫描
        private const long ParallelSearchLineThreshold = 50_000;
        // 单线程扫描时攒够这么多匹配就发布一段，避免段数过多拖慢视图重建
        private const int SearchSegmentLines = 65_536;
        // 进度回投单飞闸（1=已有一个委托在 UI 队列里排队）与首次自动跳转标记，每次新搜索重置
        private int _searchProgressQueued;
        private int _autoJumpTried;

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
            ApplyTheme();
            // 这个控件是独立窗口里的，主窗翻明暗时不会遍历到它，只能自己订阅
            Theme.Changed += OnThemeChanged;
            Disposed += (_, _) => Theme.Changed -= OnThemeChanged;
            // ★ 加载全局搜索历史到下拉框
            try { RefreshSearchHistory(); } catch { }
        }

        private void OnThemeChanged() => ApplyTheme();

        private void InitializeComponent()
        {
            SuspendLayout();

            // 一个区域只留一颗实心主按钮，其余动作统一走描边次要样式，别再混幽灵态
            _btnOpenFile = new ThemeButton { Text = "打开文件", Primary = true };
            _btnOpenFile.Click += OnOpen;
            _btnClearCache = new ThemeButton { Text = "清理缓存" };
            _btnClearCache.Click += OnClearCache;
            _fldGoto = new ThemeField { Placeholder = "跳转到行号" };
            _txtGotoLine = _fldGoto.Edit;
            _txtGotoLine.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OnGoto(s, e); };
            _btnGotoLine = new ThemeButton { Text = "跳转" };
            _btnGotoLine.Click += OnGoto;
            _lblStatus = new Label { Text = "请打开一个文本文件", AutoSize = true };

            // 工具区贴在列表下方、状态栏上方，顶部只留给内容
            _toolPanel = new Panel { Dock = DockStyle.Bottom, Height = 88 };
            _toolDivider = new Panel { Dock = DockStyle.Top, Height = 1 };
            _toolPanel.Controls.Add(_toolDivider);

            _cboSearch = new ThemeComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDown,
                // 原生 Standard 边框自带 2px 白色凹凸线，深色下压不住，换 Flat 让边框走绘制路径
                FlatStyle = FlatStyle.Flat,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            _searchHost = new SearchBoxHost();
            _searchHost.Controls.Add(_cboSearch);
            _searchHost.GotFocus += (s, e) => _cboSearch.Focus();
            _cboSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OnSearch(); };
            _cboSearch.SelectedIndexChanged += OnSearchHistorySelected;
            // ★ 不监听 TextChanged 清空 _lastSearch：导航按钮会自行比较搜索框 key 与 _lastSearch
            _chkCaseSensitive = new ThemeCheckBox { Text = "Aa", AutoSize = true };
            _chkRegex = new ThemeCheckBox { Text = ".*", AutoSize = true };
            _segNav = new ThemeSegmented { Items = new[] { "|<", "上一个", "下一个", ">|" } };
            _segNav.SegmentClicked += i =>
            {
                switch (i)
                {
                    case 0: OnSearchFirst(this, EventArgs.Empty); break;
                    case 1: OnSearchPrev(this, EventArgs.Empty); break;
                    case 2: OnSearchNext(this, EventArgs.Empty); break;
                    case 3: OnSearchLast(this, EventArgs.Empty); break;
                }
            };
            _btnExport = new ThemeButton { Text = "导出结果" };
            _btnExport.Click += OnExport;
            _btnReverseExport = new ThemeButton { Text = "反向导出" };
            _btnReverseExport.Click += OnReverseExport;
            // 取消按钮平时隐藏，导出跑起来才摆到进度文字前面（见 LayoutToolbar / ShowCancelExport）
            _btnCancelExport = new ThemeButton { Text = "取消", Visible = false };
            _btnCancelExport.Click += OnCancelExport;
            _lblContextLines = new Label { Text = "上下行", AutoSize = true };
            _fldContext = new ThemeField();
            _txtContext = _fldContext.Edit;
            _txtContext.Text = "0";
            // 原来的微调按钮没了，用上下键和滚轮补上调值手感
            _txtContext.KeyPress += OnContextKeyPress;
            _txtContext.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Up) { StepContextLines(1); e.Handled = true; }
                else if (e.KeyCode == Keys.Down) { StepContextLines(-1); e.Handled = true; }
            };
            _txtContext.MouseWheel += (s, e) => StepContextLines(e.Delta > 0 ? 1 : -1);

            _toolPanel.Controls.AddRange(new Control[]
            {
                _btnOpenFile, _btnClearCache, _fldGoto, _btnGotoLine, _lblStatus,
                _searchHost, _chkCaseSensitive, _chkRegex, _segNav,
                _lblContextLines, _fldContext, _btnExport, _btnReverseExport, _btnCancelExport
            });
            LayoutToolbar();
            _toolPanel.Resize += (s, e) => LayoutToolbar();

            _lstContent = new LineListView();
            _lstContent.Dock = DockStyle.Fill;
            _lstContent.GetLineText = idx => _cache.TryGetValue((int)idx, out var v) ? v : ReadLine(idx);
            _lstContent.SelectedIndexChanged += (s, e) =>
            {
                // 同步刷新选中行详情框（完整内容，方便阅读和复制）
                try { UpdateDetail(); } catch { }
                // 选中变化时刷新底部状态栏当前行（与滚动定时器同一口径）
                try { RefreshCurLineLabel(); } catch { }
            };
            _lstContent.ItemDoubleClick += OnDoubleClick;
            _lstContent.VScrollChanged += (s, e) => UpdateScrollPos();
            _lstContent.KeyDown += OnListViewKeyDown;

            // ★ 选中行详情框：列表只画截断的前 8KB，这里用 ReadLineFull 展示完整行，只读但可选中复制。
            // 不另占表头行：行号/大小/截断提示放悬停 tooltip，复制直接在框内选中后 Ctrl+C（或 Ctrl+A 全选）。
            _txtDetail = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                WordWrap = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                MaxLength = 0,
                Font = Theme.Code,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text
            };
            _detailTip = new ToolTip();
            _detailDivider = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Theme.P.Divider };
            _detailPanel = new Panel { Dock = DockStyle.Bottom, Height = 72, BackColor = Theme.P.Surface };
            // Fill 先加、Top 后加：分隔线在顶，文本框填满剩余
            _detailPanel.Controls.Add(_txtDetail);
            _detailPanel.Controls.Add(_detailDivider);

            // ★ 底部状态栏：一条分隔线 + 6 个流式排列的信息
            _lblFileModify = MakeStatusLabel("最后修改：-");
            _lblFileSize = MakeStatusLabel("大小：-");
            _lblFileLines = MakeStatusLabel("行数：-");
            _lblCacheSize = MakeStatusLabel("缓存：0B");
            _lblSearchPos = MakeStatusLabel("");
            _lblCurLine = MakeStatusLabel("行：1");
            _statusFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 2, 8, 0),
                AutoScroll = false,
                WrapContents = false
            };
            _statusBar = new Panel { Dock = DockStyle.Bottom, Height = Theme.P.RowHeight };
            _statusDivider = new Panel { Dock = DockStyle.Top, Height = 1 };
            _statusBar.Controls.Add(_statusFlow);
            _statusBar.Controls.Add(_statusDivider);
            _statusFlow.Controls.AddRange(new Control[] { _lblFileModify, _lblFileSize, _lblCacheSize, _lblCurLine, _lblFileLines, _lblSearchPos });

            // 同一条边上后加入的靠外：状态栏最底，其次工具区，再是详情框，列表填满剩余
            Controls.Add(_lstContent);
            Controls.Add(_detailPanel);
            Controls.Add(_toolPanel);
            Controls.Add(_statusBar);
            ResumeLayout(false);

            _timer = new System.Windows.Forms.Timer { Interval = 200 };
            _timer.Tick += (_, _) => Tick();
            _timerScroll = new System.Windows.Forms.Timer { Interval = 300 };
            _timerScroll.Tick += (_, _) => UpdateScrollPos();
        }

        /// <summary>
        /// 顶部工具区逐行算坐标：第一行左组管文件和跳转、右组管上下文和导出，中间的空档给进度文字；
        /// 第二行整条给搜索与导航。窗口窄到两组放不下时，导出组退回第二行尾部，再挤就折第三行，
        /// 而不是像流式布局那样把搜索框挤没。
        /// </summary>
        private void LayoutToolbar()
        {
            // 边距/间距与左侧文件区工具条同一套常量，按钮宽度按文字量，不再写死
            int margin = ThemeMetrics.Margin, gap = ThemeMetrics.Gap, group = ThemeMetrics.Group;
            const int searchMin = 220;
            // 进度文字至少留这么多才值得挤进第一行，否则宁可交给底部状态栏
            const int statusMin = 140;
            // 两行工具区是叠起来的，按单行的 RowPad=8 走会翻成两倍高；折到 4px，和表格阅读器过滤条一致
            const int rowPad = 4;
            int rowH = ThemeMetrics.RowH;
            int w = Math.Max(360, _toolPanel.Width - margin * 2);
            int y1 = rowPad;

            int wOpen = ThemeButton.FitWidth(_btnOpenFile.Text);
            int wClear = ThemeButton.FitWidth(_btnClearCache.Text);
            int wGoto = ThemeButton.FitWidth(_btnGotoLine.Text);
            int wCtx = 64;
            int wExport = ThemeButton.FitWidth(_btnExport.Text);
            int wReverse = ThemeButton.FitWidth(_btnReverseExport.Text);

            int flagsW = _chkCaseSensitive.Width + 2 + _chkRegex.Width;
            int ctxW = _lblContextLines.Width + 4 + wCtx;
            int expW = wExport + gap + wReverse;
            int leftRow1 = wOpen + gap + wClear + group + 130 + gap + wGoto;
            int rightRow1 = ctxW + group + expW;

            int x = margin;
            _btnOpenFile.SetBounds(x, y1, wOpen, rowH); x += wOpen + gap;
            _btnClearCache.SetBounds(x, y1, wClear, rowH); x += wClear + group;
            _fldGoto.SetBounds(x, y1, 130, rowH); x += 130 + gap;
            _btnGotoLine.SetBounds(x, y1, wGoto, rowH);

            // 上下文/导出这一组靠右对齐；两组之间留不出放进度文字的位子时，把这组退回第二行尾部
            bool splitRight = w - leftRow1 - rightRow1 >= group * 2 + statusMin;
            // 进度/结果文字跟在文件组后面（说的就是这个文件），空档留在两组之间；
            // 右组退回第二行时才贴到行尾，免得在空档里飘着。
            // 导出期间取消按钮常驻进度文字前面，它占掉的宽度要从文字可用空间里先扣掉
            int statusW = _lblStatus.Width;
            bool exporting = _btnCancelExport.Visible;
            int cancelW = exporting ? ThemeButton.FitWidth(_btnCancelExport.Text) : 0;
            int leadW = exporting ? cancelW + gap : 0;
            int statusX = splitRight ? _btnGotoLine.Right + group : margin + w - statusW - leadW;
            int statusLimit = splitRight ? margin + w - rightRow1 - group : margin + w;
            if (exporting)
                _btnCancelExport.SetBounds(statusX, y1, cancelW, rowH);
            bool showStatus = statusX + leadW + statusW <= statusLimit;
            _lblStatus.Visible = showStatus;
            if (showStatus)
                _lblStatus.Location = new Point(statusX + leadW, y1 + (rowH - _lblStatus.Height) / 2);

            // 第二行搜索框右侧的固定占位：搜索框吃掉剩下的全部宽度，整行刚好收在右边界
            int navTailW = gap + flagsW + group + _segNav.Width;
            int fullTailW = navTailW + group + ctxW + group + expW;
            int y2 = y1 + rowH + gap;
            int y = y2;

            x = margin;
            bool fold = !splitRight && w < searchMin + fullTailW;
            // 窄到连「搜索框 + Aa/.* + 导航」都不同行时，让搜索框独占一行、导航组另起一行。
            // 不这样做的话下面 Math.Max 的下限会把导航组顶到面板右边界之外，四个按钮一个都点不到。
            bool navFold = w < searchMin + navTailW;
            int searchW = navFold ? w : Math.Max(searchMin, w - (fold || splitRight ? navTailW : fullTailW));
            int navRowY = navFold ? y2 + rowH + gap : y2;
            // 独立成行时右对齐收边，和搜索框那一行的右端保持一条线
            int navX = navFold ? margin + w - (flagsW + group + _segNav.Width) : margin + searchW + gap;
            _searchHost.SetBounds(x, y2, searchW, rowH);
            // 内缩 2px：让下拉框自带的 1px 边和容器描边并成一条，别看成框中框
            _cboSearch.SetBounds(2, (rowH - _cboSearch.Height) / 2, searchW - 4, _cboSearch.Height);
            x = navX;
            _chkCaseSensitive.Location = new Point(x, navRowY + (rowH - _chkCaseSensitive.Height) / 2);
            x += _chkCaseSensitive.Width + 2;
            _chkRegex.Location = new Point(x, navRowY + (rowH - _chkRegex.Height) / 2);
            x = _chkRegex.Right + group;
            _segNav.SetBounds(x, navRowY + (rowH - _segNav.Height) / 2, _segNav.Width, _segNav.Height);

            if (splitRight)
            {
                // 上下文和导出挂在第一行右端，导出按钮的右边缘就是内容区右边界
                int ex = margin + w - expW;
                _btnExport.SetBounds(ex, y1, wExport, rowH);
                _btnReverseExport.SetBounds(ex + wExport + gap, y1, wReverse, rowH);
                int cx = ex - group - ctxW;
                _lblContextLines.Location = new Point(cx, y1 + (rowH - _lblContextLines.Height) / 2);
                _fldContext.SetBounds(cx + _lblContextLines.Width + 4, y1, wCtx, rowH);
            }
            else
            {
                if (fold || navFold)
                {
                    y = _segNav.Bottom + gap;
                    x = margin;
                }
                else x = _segNav.Right + group;

                _lblContextLines.Location = new Point(x, y + (rowH - _lblContextLines.Height) / 2);
                x += _lblContextLines.Width + 4;
                _fldContext.SetBounds(x, y, wCtx, rowH);
                int ex = margin + w - expW;
                _btnExport.SetBounds(ex, y, wExport, rowH);
                _btnReverseExport.SetBounds(ex + wExport + gap, y, wReverse, rowH);
            }

            int h = Math.Max(Math.Max(_searchHost.Bottom, _segNav.Bottom), _btnReverseExport.Bottom) + rowPad;
            if (_toolPanel.Height == h) return;
            // Bottom 停靠改 Height 时不会把 Top 一起抬上去：折成三行会盖住状态栏，收回两行又在中间留一条缝。
            // 高度落定后直接按状态栏顶边把面板摆回去，停靠宽度和列表占位都由停靠自己维护。
            _toolPanel.Height = h;
            // 构造期第一次排版时状态栏还没建，跳过摆位（那时面板尚未停靠，位置由停靠接管）
            if (_statusBar != null)
                _toolPanel.Location = new Point(0, _statusBar.Top - h);
        }

        /// <summary>数字框只收数字，粘贴和退路键照常放行。</summary>
        private void OnContextKeyPress(object? s, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar) || char.IsDigit(e.KeyChar)) return;
            e.Handled = true;
        }

        private void StepContextLines(int delta)
        {
            int v = ContextLines + delta;
            _txtContext.Text = Math.Clamp(v, 0, 999).ToString();
        }

        /// <summary>上下文行数，非法输入按 0 处理。</summary>
        private int ContextLines => int.TryParse(_txtContext.Text, out var v) ? Math.Clamp(v, 0, 999) : 0;

        /// <summary>只画一圈主题描边的容器，给去掉原生边框的可编辑下拉框补外形和聚焦环。</summary>
        private sealed class SearchBoxHost : Panel
        {
            public SearchBoxHost()
            {
                SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                var p = Theme.P;
                if (Width <= p.Radius * 2 || Height <= p.Radius * 2) return;

                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                bool focused = ContainsFocus;
                float stroke = focused ? 2f : 1f;
                var rect = new RectangleF(stroke / 2f, stroke / 2f, Width - stroke, Height - stroke);
                using var path = ThemeButton.RoundRect(rect, p.Radius);
                using (var fill = new SolidBrush(p.Surface)) g.FillPath(fill, path);
                using var pen = new Pen(focused ? p.Accent : p.Border, stroke);
                g.DrawPath(pen, path);
            }
        }

        /// <summary>取色全部走当前皮肤令牌；主窗翻明暗或换皮肤后也会重新走一遍。</summary>
        public void ApplyTheme()
        {
            var p = Theme.P;
            BackColor = p.Surface;
            _toolPanel.BackColor = p.Chrome;
            _statusBar.BackColor = p.Chrome;
            _statusBar.Height = p.RowHeight;
            _statusFlow.BackColor = p.Chrome;
            _statusDivider.BackColor = p.Divider;
            _toolDivider.BackColor = p.Divider;

            foreach (var b in new[] { _btnOpenFile, _btnClearCache, _btnGotoLine, _btnExport, _btnReverseExport, _btnCancelExport })
            {
                b.Font = Theme.Body;
                b.BackColor = p.Chrome;
                b.Invalidate();
            }

            _fldGoto.ApplyTheme();
            _fldContext.ApplyTheme();
            _searchHost.BackColor = p.Chrome;
            _searchHost.Invalidate();
            _cboSearch.Font = Theme.Body;
            _cboSearch.BackColor = p.Surface;
            _cboSearch.ForeColor = p.Text;
            _cboSearch.Invalidate();

            foreach (var c in new[] { _chkCaseSensitive, _chkRegex })
            {
                c.Font = Theme.Body;
                c.BackColor = p.Chrome;
                c.ForeColor = p.Text;
            }

            _segNav.RefreshTheme();

            _lblStatus.Font = Theme.Caption;
            _lblStatus.ForeColor = p.Muted;
            _lblStatus.BackColor = p.Chrome;
            _lblContextLines.Font = Theme.Caption;
            _lblContextLines.ForeColor = p.Muted;
            _lblContextLines.BackColor = p.Chrome;
            foreach (var lbl in new[] { _lblFileModify, _lblFileSize, _lblFileLines, _lblCacheSize, _lblSearchPos, _lblCurLine })
            {
                lbl.Font = Theme.Caption;
                lbl.ForeColor = p.Muted;
                lbl.BackColor = p.Chrome;
            }

            // 选中行详情框跟随皮肤
            _detailPanel.BackColor = p.Surface;
            _detailDivider.BackColor = p.Divider;
            _txtDetail.Font = Theme.Code;
            _txtDetail.BackColor = p.Surface;
            _txtDetail.ForeColor = p.Text;

            _lstContent.ApplyTheme();
            LayoutToolbar();
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
            _lstContent.SetMaxLineWidth(0);
            _lstContent.SetTotalLines(0);
            _lstContent.ClearSelection();

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

            _lstContent.SetTotalLines(Math.Max(loaded, 1));
            _cS = 0; _cE = loaded - 1;

            Status($"前 {loaded} 行已加载，正在构建索引...");
            _lblFileLines.Text = "行数：索引中...";
            _lstContent.Invalidate();

            _timer.Start();
            _ = Task.Run(() => BuildIndex());
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
                    if (newSize > _lstContent.TotalLines)
                    {
                        _lstContent.SetTotalLines(newSize);
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

                // ★ 检测行数是否超过切割阈值
                if (_lineOffsets != null && _lineOffsets.Length > SPLIT_LINE_THRESHOLD)
                {
                    CheckAndPromptSplit(_lineOffsets.Length);
                    return;
                }

                // 索引建完就把"正在构建索引"这条过程消息收掉：大小和行数底部状态栏已经有，
                // 在这里再写一遍会让工具条中间出现一段和按钮对不上号的重复信息
                Status("");
                // ★ 底部状态栏：行数 + 当前缓存大小
                _lblFileLines.Text = $"行数：{_idxCount:N0}";
                RefreshCacheSizeStatus();
                // ★ LineListView 支持 long 行数，末尾追加空白行确保最后一行完整显示
                _lstContent.SetTotalLines(_idxCount + TAIL_EMPTY_LINES);
                // ★ 索引完成后一次性测量最长行并设置横向滚动条，之后纵向滚动不再重算（与 Sublime 一致）
                SetHorizontalScrollOnce();
                _timerScroll.Start();
                _lstContent.Invalidate();
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
                // ★ 底部状态栏：当前行（滚动时更新）。有选中行时保持选中行号，
                // 不能拿视口顶行直接覆盖，否则 300ms 定时器会把选中行号重置回顶行。
                RefreshCurLineLabel();
                // ★ 同时定时刷新缓存大小（缓存随时可能因搜索而增长）
                RefreshCacheSizeStatus();
                // 预读可见行缓存
                EnsureVisibleCache();
            }
            catch { }
        }

        /// <summary>预读可见行到缓存，模拟原 ListView 的 CacheVirtualItems 行为</summary>
        private void EnsureVisibleCache()
        {
            if (_filePath == null || _lineOffsets == null) return;
            int top = _lstContent.TopIndex;
            int visible = Math.Max(1, _lstContent.ClientRectangle.Height / Math.Max(_lstContent.Font.Height + 6, 1));
            int start = Math.Max(0, top - 5);
            int end = (int)Math.Min(_lstContent.TotalLines - 1, top + visible + 5);
            if (start > end) return;
            // 若已在缓存范围内则跳过
            if (_cS >= 0 && start >= _cS && end <= _cE) return;
            for (int i = _cS; i <= _cE; i++) if (i < start || i > end) _cache.TryRemove(i, out _);
            for (int i = start; i <= end; i++) if (!_cache.ContainsKey(i)) _cache[i] = ReadLine(i);
            _cS = start; _cE = end;
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
                    int candidate = sz.Width + 60;  // 额外 60px 充裕空间，避免最长行贴边
                    if (candidate > 0)
                        maxW = Math.Max(clientW, candidate);
                    else
                        Debug.WriteLine($"[Measure] 最长行{lineIdx} 异常测量 candidate={candidate} text.Length={text.Length} sz.Width={sz.Width}");
                }
                maxW = Math.Min(maxW, MAX_SCROLL_RANGE);

                Debug.WriteLine($"[Measure] SetHorizontalScrollOnce MaxLineWidth {_lstContent.MaxLineWidth} -> {maxW} (最长行={lineIdx}, 字节数={_longestLineBytes})");
                _lstContent.SetMaxLineWidth(maxW);
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

        /// <summary>
        /// 导出用的行读取：只认调用方在任务开始时快照的偏移表和自己打开的流，不碰 <see cref="_lineOffsets"/> / <see cref="_readStream"/>。
        /// 关掉页签会走 <see cref="Cleanup"/>，那两个字段被置空后 <see cref="ReadLine"/> 静默返回空串，
        /// 剩下的行会以空行写满整个文件还报成功；这里读不到就返回 null，由调用方中止并删掉残缺输出。
        /// </summary>
        private static string? ReadLineForExport(FileStream fs, long[] offsets, long fileLength, long idx)
        {
            if ((ulong)idx >= (ulong)offsets.Length) return null;
            long s = offsets[idx];
            long e = idx + 1 < offsets.Length ? offsets[idx + 1] : fileLength;
            int lineLen = (int)Math.Min(e - s, FULL_LINE_LIMIT);
            if (lineLen <= 0) return "";
            try
            {
                fs.Seek(s, SeekOrigin.Begin);
                var buf = new byte[lineLen];
                int got = 0;
                while (got < lineLen)
                {
                    int r = fs.Read(buf, got, lineLen - got);
                    if (r <= 0) break;
                    got += r;
                }
                return got > 0
                    ? Encoding.UTF8.GetString(buf, 0, got).TrimEnd(GlobalConst.CrChar, GlobalConst.LfChar)
                    : "";
            }
            catch { return null; }
        }

        /// <summary>后台导出任务回 UI 线程：页签可能已经关掉，控件销毁时 BeginInvoke 会抛，这里统一吞掉。</summary>
        private void ExportUi(Action action)
        {
            try
            {
                if (IsDisposed || !IsHandleCreated) return;
                BeginInvoke(action);
            }
            catch { }
        }

        /// <summary>
        /// 导出进度节流上报：写盘循环每 65536 行才看一次表，距上次上报不足 250ms 就跳过。
        /// 千万行导出若每行都 BeginInvoke，UI 消息队列会被进度刷爆，反而卡界面。
        /// </summary>
        private void ReportExportProgress(string verb, long done, long total, Stopwatch sw, ref long lastMs)
        {
            long now = sw.ElapsedMilliseconds;
            if (now - lastMs < 250) return;
            lastMs = now;
            long pct = total > 0 ? done * 100 / total : 0;
            ExportUi(() => Status($"{verb} {done:N0}/{total:N0} 行 ({pct}%)"));
        }

        /// <summary>导出收尾：放开并发闸并清掉自己的 CTS。凭 token 比对，不替 Cleanup 或下一次导出清。</summary>
        private void ReleaseExport(CancellationToken token)
        {
            Interlocked.Exchange(ref _exportRunning, 0);
            if (_exportCts != null && _exportCts.Token == token)
            {
                var c = _exportCts;
                _exportCts = null;
                try { c.Dispose(); } catch { }
            }
            // 收尾在后台线程，隐藏取消按钮要回 UI；若此刻已有新导出抢到闸，就别误隐藏它的按钮
            ExportUi(() =>
            {
                if (Volatile.Read(ref _exportRunning) == 0)
                {
                    _btnCancelExport.Visible = false;
                    LayoutToolbar();
                }
            });
        }

        /// <summary>导出开跑：把取消按钮摆到进度文字前面（在 UI 线程、抢到并发闸之后调用）。</summary>
        private void ShowCancelExport()
        {
            _btnCancelExport.Enabled = true;
            _btnCancelExport.Visible = true;
            LayoutToolbar();
        }

        /// <summary>取消正在跑的导出：只发信号，半截产物由导出线程的 finally 自己删。灰掉按钮防连点。</summary>
        private void OnCancelExport(object? s, EventArgs e)
        {
            _btnCancelExport.Enabled = false;
            // 导出可能正好在这一刻收尾并 Dispose 掉 CTS，Cancel 会抛 ObjectDisposedException
            try { _exportCts?.Cancel(); } catch { }
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





        private void Status(string s)
        {
            if (_lblStatus == null) return;
            try
            {
                BeginInvoke(() =>
                {
                    _lblStatus.Text = s;
                    // 第一行放不放得下这句进度是按宽度算的，文字换了不重排就会照上一句的宽度摆，长消息压到按钮上
                    LayoutToolbar();
                });
            }
            catch { }
        }
        static string FormatSize(long b) { string[] u = ["B", "KB", "MB", "GB", "TB"]; double v = b; int o = 0; while (v >= 1024 && o < 4) { o++; v /= 1024; } return $"{v:0.##}{u[o]}"; }

        /// <summary>创建底部状态栏的 Label</summary>
        private static Label MakeStatusLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Margin = new Padding(0, 2, 16, 0),
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
            // Ctrl+C 已由 LineListView 内部处理，此处保留兼容
            if (e.Control && e.KeyCode == Keys.C)
            {
                string text = _lstContent.GetSelectedText();
                if (!string.IsNullOrEmpty(text))
                {
                    try { Clipboard.SetText(text); }
                    catch
                    {
                        try { Thread.Sleep(100); Clipboard.SetDataObject(text, true); }
                        catch { }
                    }
                }
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void CopySelectedLines()
        {
            string text = _lstContent.GetSelectedText();
            if (string.IsNullOrEmpty(text)) return;
            try { Clipboard.SetText(text); }
            catch
            {
                try { Thread.Sleep(100); Clipboard.SetDataObject(text, true); }
                catch { }
            }
        }

        /// <summary>
        /// 刷新选中行详情框：取首个选中行的完整内容（ReadLineFull，不是列表那份 8KB 截断版），
        /// 超过 256KB 只显示前段并提示双击看全，免得超大行把 EDIT 控件卡死。
        /// 行号/大小/截断提示放悬停 tooltip，不另占表头行。
        /// </summary>
        private void UpdateDetail()
        {
            if (_txtDetail == null) return;

            // 取最小选中行号（最靠上那行）；多选时只显示首行
            long idx = -1; int selCount = 0;
            foreach (var s in _lstContent.SelectedIndices)
            {
                selCount++;
                if (idx < 0 || s < idx) idx = s;
            }

            if (idx < 0 || idx >= _lstContent.TotalLines)
            {
                _txtDetail.Text = "";
                _detailTip.SetToolTip(_txtDetail, "");
                return;
            }

            bool over64 = false;
            string full = ReadLineFull(idx, out over64) ?? "";
            bool capped = full.Length > DETAIL_LINE_LIMIT;
            _txtDetail.Text = capped ? full.Substring(0, DETAIL_LINE_LIMIT) : full;

            // 行字节数直接查偏移表，O(1)，不必为超大行把整行再编码一遍
            long lineBytes = 0;
            var off = _lineOffsets;
            if (off != null && idx < off.Length)
                lineBytes = (idx + 1 < off.Length ? off[idx + 1] : _fileLength) - off[idx];

            string note = over64 ? "（超过 64MB，已截断）"
                        : capped ? "（超大行，仅显示前 256KB，双击查看完整）"
                        : selCount > 1 ? $"（已选 {selCount} 行，显示首行）"
                        : "";
            _detailTip.SetToolTip(_txtDetail, $"行 {idx + 1:N0} — {FormatSize(lineBytes)}{note}");
        }

        /// <summary>
        /// 底部状态栏「当前行」唯一口径：有选中行显示选中行，无选中才显示视口顶行。
        /// 滚动定时器和选中变化都走这里，否则 300ms 定时器会把选中行号覆盖回顶行。
        /// </summary>
        private void RefreshCurLineLabel()
        {
            long cur = -1;
            foreach (var s in _lstContent.SelectedIndices)
                if (cur < 0 || s < cur) cur = s;
            if (cur < 0) cur = _lstContent.TopIndex;
            _lblCurLine.Text = $"行：第{cur + 1:N0}行";
        }

        private async void OnDoubleClick(long idx)
        {
            if (idx < 0 || idx >= _lstContent.TotalLines) return;
            _lstContent.SelectRow(idx);

            string originalText = _cache.TryGetValue((int)idx, out var v) ? v : "";
            if (string.IsNullOrEmpty(originalText)) return;

            // 读取完整行内容用于弹窗（不截断），并计算内容大小
            bool isFullTruncated = false;
            string? fullLine = ReadLineFull(idx, out isFullTruncated);
            if (string.IsNullOrEmpty(fullLine)) fullLine = originalText;
            string sizeLabel = FormatSize(Encoding.UTF8.GetByteCount(fullLine));

            var f = new Form
            {
                Text = $"行 {idx + 1} — {sizeLabel}",
                Size = new Size(950, 550),
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = true,
                MaximizeBox = true,
                Icon = WindowIcons.Logs,
                ShowInTaskbar = false,
                BackColor = Theme.P.Window
            };

            // 如果超过 64MB，弹出警告
            if (isFullTruncated)
            {
                MessageBox.Show("该行内容超过 64MB，已截断显示前 64MB。" + GlobalConst.NewLine + $"行号：{idx + 1}", "内容过大", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // 顶部工具栏：格式化 / 原始 做成一颗分段切换，省得两个按钮互相抢底色
            var topPanel = new Panel { Height = Theme.P.ButtonHeight + 12, Dock = DockStyle.Top, BackColor = Theme.P.Chrome };
            var segView = new ThemeSegmented { Items = new[] { "格式化", "原始" }, ActiveIndex = 0 };
            segView.Location = new Point(10, (topPanel.Height - segView.Height) / 2);

            var rtb = new RichTextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = Theme.Code,
                ScrollBars = RichTextBoxScrollBars.Both,
                WordWrap = false,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
                Text = "正在格式化..."
            };

            string formattedText = "";
            segView.SegmentClicked += async i =>
            {
                segView.ActiveIndex = i;
                if (i == 1)
                {
                    rtb.WordWrap = true;
                    rtb.ScrollBars = RichTextBoxScrollBars.Vertical;
                    rtb.Text = fullLine;
                    return;
                }
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

            topPanel.Controls.Add(segView);
            // 原生边框是系统灰，深色下刺眼；换成令牌色 1px 分隔条，Chrome 和内容区才分得开
            var rtDivider = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Theme.P.Border };
            f.Controls.Add(rtb);
            f.Controls.Add(rtDivider);
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
                    case 92: sb.Append(@"\\"); break;  // \
                    case 123: sb.Append(@"\{ "); break; // {
                    case 125: sb.Append(@"\} "); break; // }
                    case 10: sb.Append(@"\par "); break; // \n
                    case 13: break;                  // \r
                    default: sb.Append($"\\'{(int)b:X2}"); break;
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
            long virtualSize = _lstContent.TotalLines;
            if (ln > virtualSize) { AlertForm.Show(ParentForm, $"超出范围，共 {virtualSize:N0} 行", "提示", AlertType.Information); return; }
            _lstContent.SelectRow(ln - 1);
            _lstContent.Focus();
        }

        // ==================== 搜索功能 ====================

        private long GetCurrentLine()
        {
            try
            {
                if (_lstContent.SelectedIndices.Count > 0)
                    return (new List<long>(_lstContent.SelectedIndices))[0];
                return _lstContent.TopIndex;
            }
            catch { return 0; }
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
        private void PutRecentSearch(string key, SearchResultSet results)
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
            _recentSearches.AddLast(new KeyValuePair<string, SearchResultSet>(key, results));
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
        private SearchResultSet? TryGetRecentSearch(string key)
        {
            foreach (var kv in _recentSearches)
            {
                if (kv.Key == key)
                    return kv.Value;
            }
            return null;
        }

        /// <summary>估算搜索结果集占用字节数</summary>
        private static long EstimateSearchResultBytes(SearchResultSet results)
        {
            return results.EstimatedBytes;
        }

        /// <summary>按关键词/大小写/正则构造逐行匹配器</summary>
        private static Func<string, bool> BuildLineMatcher(string keyword, bool caseSensitive, bool useRegex)
        {
            if (useRegex)
            {
                var opts = caseSensitive ? System.Text.RegularExpressions.RegexOptions.None : System.Text.RegularExpressions.RegexOptions.IgnoreCase;
                var regex = new System.Text.RegularExpressions.Regex(keyword, opts);
                return (line) => regex.IsMatch(line);
            }
            StringComparison cmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            return (line) => line.Contains(keyword, cmp);
        }

        /// <summary>搜索进行中的状态栏文案：分块扫描报块进度，单线程扫描报百分比</summary>
        private static string BuildSearchingText(long count, long doneUnits, long totalUnits, long scannedLines, long totalLines)
        {
            if (totalUnits > 1) return $"搜索中... {count:N0} 个匹配 ({doneUnits}/{totalUnits}块)";
            int pct = totalLines > 0 ? (int)((scannedLines + 1L) * 100 / totalLines) : 0;
            return $"搜索中... {count:N0} 个 ({pct}%)";
        }

        /// <summary>
        /// 请刷进度。后台每千行/每块都会调，这里用单飞闸把它压成「最多一个委托在排队」，
        /// 否则千万行扫描会往 UI 消息队列里塞几万个委托，界面被自己的进度条刷死。
        /// </summary>
        private void QueueSearchProgress(SearchResultSet results, string currentKey,
            long doneUnits, long totalUnits, long scannedLines, long totalLines)
        {
            if (Interlocked.CompareExchange(ref _searchProgressQueued, 1, 0) != 0) return;
            if (!IsHandleCreated) { _searchProgressQueued = 0; return; }
            long snapshotDone = doneUnits, snapshotTotal = totalUnits;
            long snapshotScanned = scannedLines, snapshotLines = totalLines;
            SafeBeginInvoke(() =>
            {
                _searchProgressQueued = 0;
                if (_lastSearch != currentKey) return;
                _lblSearchPos.Text = BuildSearchingText(results.Count, snapshotDone, snapshotTotal, snapshotScanned, snapshotLines);
            });
        }

        /// <summary>扫出第一个匹配时自动跳过去；用户已手动跳转或已跳过则不再打扰</summary>
        private void TryAutoJump(SearchResultSet results, string currentKey, long curLine)
        {
            if (_hasuserNavigated || results.Count == 0) return;
            if (Interlocked.CompareExchange(ref _autoJumpTried, 1, 0) != 0) return;
            SafeBeginInvoke(() =>
            {
                if (_lastSearch != currentKey || _hasuserNavigated) return;
                long idx = results.IndexOfNearest(curLine);
                if (idx >= 0) NavigateTo(results, idx);
            });
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
            // ★ 只认封了口的：切走关键词时上一次搜索可能被中途放弃，缓存里留的是半截结果，
            //   拿它当"搜索完成"会让导航和导出都少一截
            var recent = TryGetRecentSearch(currentKey);
            if (recent != null && recent.IsSealed)
            {
                _searchCts = null;
                _lastSearch = currentKey;
                _searchCurrentIndex = -1;
                _lastNavigatedMatchLine = -1;
                _hasuserNavigated = false;

                if (recent.Count > 0)
                {
                    NavigateTo(recent, recent.IndexOfNearest(GetCurrentLine()));
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
                    NavigateTo(cached, cached.IndexOfNearest(GetCurrentLine()));
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
            _autoJumpTried = 0;
            Volatile.Write(ref _searchProgressQueued, 0);
            // ★ 先建空容器进缓存，再把引用直接交给后台任务：
            // 后台若回头去查 _recentSearches，会和 UI 线程增删链表打架（链表非线程安全）
            var container = new SearchResultSet();
            PutRecentSearch(currentKey, container);
            // ★ 实际发起搜索时才写入历史记录、刷新下拉框
            SearchHistoryStore.Add(keyword, caseSensitive, useRegex);
            RefreshSearchHistory();
            long curLine = GetCurrentLine();
            _ = Task.Run(() => DoSearch(keyword, caseSensitive, useRegex, token, curLine, container));
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

        private void DoSearch(string keyword, bool caseSensitive, bool useRegex, CancellationToken token, long curLine, SearchResultSet results)
        {
            try
            {
                Func<string, bool> matcher = BuildLineMatcher(keyword, caseSensitive, useRegex);
                long total = _lineOffsets!.Length;
                string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";

                if (total > ParallelSearchLineThreshold)
                {
                    // ★ 并行分块搜索：每块把匹配行号按块号发布成独立一段。
                    // 块号即字节位置即行号顺序，所以段与段天然全局有序，不需要合并；
                    // 旧实现把各块结果逐个 Insert 进同一个有序 List（O(n) 搬移），
                    // 千万级匹配时光合并就要在 UI 线程跑几十分钟，界面直接假死。
                    var lineBounds = ComputeLineBoundsForBlocks();
                    int totalBlocks = lineBounds.Length;
                    int completedCount = 0;
                    var options = new ParallelOptions
                    {
                        MaxDegreeOfParallelism = MAX_PARALLEL,
                        CancellationToken = token
                    };

                    try
                    {
                        Parallel.For(0, totalBlocks, options, block =>
                        {
                            var (startLine, endLine) = lineBounds[block];
                            var local = new List<long>();
                            using var fs = new FileStream(_filePath!, FileMode.Open, FileAccess.Read,
                                FileShare.ReadWrite, 65536, FileOptions.RandomAccess);

                            for (long i = startLine; i <= endLine; i++)
                            {
                                if (token.IsCancellationRequested) return;
                                string line = ReadLineWithStream(fs, i);
                                if (!string.IsNullOrEmpty(line) && matcher(line))
                                    local.Add(i);
                            }

                            // 取消掉的块不落位：结果集只描述一次完整扫描
                            if (token.IsCancellationRequested) return;

                            results.Publish(block, local.ToArray());
                            int done = Interlocked.Increment(ref completedCount);
                            QueueSearchProgress(results, currentKey, done, totalBlocks, 0, 0);
                            TryAutoJump(results, currentKey, curLine);
                        });
                    }
                    catch (OperationCanceledException) { return; }

                    if (token.IsCancellationRequested) return;

                    results.Seal();
                    SafeBeginInvoke(() =>
                    {
                        if (_lastSearch != currentKey) return;
                        OnSearchComplete(keyword, caseSensitive, useRegex, results);
                    });
                }
                else
                {
                    // 小文件：单线程顺序扫，同样按段发布（≤5万行，一段就够）
                    bool cancelled = ScanLinesSequential(results, currentKey, matcher, token, curLine, autoJump: true);
                    if (cancelled) return;

                    SafeBeginInvoke(() =>
                    {
                        if (_lastSearch != currentKey) return;
                        OnSearchComplete(keyword, caseSensitive, useRegex, results);
                    });
                }
            }
            catch { }
            finally
            {
                // 两条分支都同步等到扫描彻底结束才返回，所以收尾统一放这里
                ReleaseSearchCts(token);
            }
        }

        /// <summary>
        /// 单线程顺序扫描全文件，匹配行号攒够一段就发布出去。
        /// 返回 true 表示中途被取消。发布出去的段不再改动，所以 UI 线程可以同时读。
        /// </summary>
        private bool ScanLinesSequential(SearchResultSet results, string currentKey,
            Func<string, bool> matcher, CancellationToken token, long curLine, bool autoJump)
        {
            long total = _lineOffsets!.Length;
            var local = new List<long>();
            for (long i = 0; i < total; i++)
            {
                if (token.IsCancellationRequested) break;
                string line = ReadLine(i);
                if (!string.IsNullOrEmpty(line) && matcher(line))
                {
                    local.Add(i);
                    if (local.Count >= SearchSegmentLines)
                    {
                        results.AppendSegment(local.ToArray());
                        local.Clear();
                    }
                }
                if ((i & 1023) == 1023)
                {
                    QueueSearchProgress(results, currentKey, 1, 1, i, total);
                    if (autoJump) TryAutoJump(results, currentKey, curLine);
                }
            }
            if (local.Count > 0) results.AppendSegment(local.ToArray());
            if (token.IsCancellationRequested) return true;
            results.Seal();
            return false;
        }

        /// <summary>后台线程往 UI 投委托时窗口可能已被关掉（句柄未建/已毁），这里吞掉那一下异常</summary>
        private void SafeBeginInvoke(Action action)
        {
            try
            {
                if (IsHandleCreated) BeginInvoke(action);
            }
            catch (InvalidOperationException) { } // 含 ObjectDisposedException：句柄已毁
        }

        /// <summary>
        /// 搜索任务结束时清掉自己的 CTS，避免导航按钮误判"搜索中"。
        /// 凭 token 比对：用户已发起新搜索时 _searchCts 是新的那个，不能替它清。
        /// </summary>
        private void ReleaseSearchCts(CancellationToken token)
        {
            if (_searchCts != null && _searchCts.Token == token)
            {
                var cts = _searchCts;
                _searchCts = null;
                try { cts.Dispose(); } catch { }
            }
        }

        /// <summary>搜索全部完成后的收尾逻辑（写入缓存、更新状态）</summary>
        private void OnSearchComplete(string keyword, bool caseSensitive, bool useRegex, SearchResultSet results)
        {
            string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";
            if (_lastSearch != currentKey) return;

            // ★ 写入 FileCacheStore 持久化缓存（内存存引用，磁盘后台分段写）
            _ = FileCacheStore.PutSearchResultAsync(_filePath!, currentKey, results);

            if (results.Count > 0)
            {
                // ★ 如果用户已经手动跳转过，不再强制跳转
                if (!_hasuserNavigated)
                {
                    NavigateTo(results, results.IndexOfNearest(GetCurrentLine()));
                }
                else
                {
                    _lblSearchPos.Text = $"搜索完成，共 {results.Count:N0} 个匹配";
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

        private void NavigateTo(SearchResultSet results, long index)
        {
            _searchCurrentIndex = index;
            long line = results[index];
            _lstContent.SelectRow(line);
            _lstContent.Focus();
            _lastNavigatedMatchLine = line;
            // ★ 底部状态栏：搜索位置 + 当前行
            _lblSearchPos.Text = $"搜索：第{index + 1:N0}/{results.Count:N0}个";
            RefreshCurLineLabel();
        }

        private bool userHasMovedAway()
        {
            if (_lastNavigatedMatchLine < 0) return true;
            long sel = -1;
            if (_lstContent.SelectedIndices.Count > 0)
                sel = (new List<long>(_lstContent.SelectedIndices))[0];
            if (sel != _lastNavigatedMatchLine)
                return true;
            int top = _lstContent.TopIndex;
            if (Math.Abs((long)top - _lastNavigatedMatchLine) > (_lstContent.ClientRectangle.Height / (_lstContent.Font.Height + 6)))
                return true;
            return false;
        }

        /// <summary>
        /// 取当前关键字的搜索结果集，供导航使用。
        /// 段发布后不再改动，所以直接给引用即可；旧实现每次导航 ToList() 拷一份，
        /// 千万级匹配就是每次几十 MB 的大对象堆分配，连按几下 F3 就把 GC 压垮了。
        /// </summary>
        private SearchResultSet? CurrentSearchResults()
        {
            string currentKey = CurrentSearchKey();
            if (string.IsNullOrEmpty(_cboSearch.Text.Trim()) || currentKey != _lastSearch)
                return null;
            var set = TryGetRecentSearch(currentKey);
            if (set == null || set.Count == 0)
                return null;
            return set;
        }

        private async void OnSearchPrev(object? s, EventArgs e)
        {
            bool ready = await EnsureSearchReadyAsync();
            var results = CurrentSearchResults();
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
                long idx2 = results.IndexOfNearest(curLine);
                if (results[idx2] >= curLine)
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
            var results = CurrentSearchResults();
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
                long idx2 = results.IndexOfNearest(curLine);
                if (results[idx2] <= curLine)
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
            var results = CurrentSearchResults();
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
            var results = CurrentSearchResults();
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

            // ★ 未封口就不能导：要么同一条件的搜索还在后台跑（容器里只有已落位的几段），
            //   要么切走关键词时留下了半截结果，两种情况导出去的都是残缺文件
            if (results != null && !results.IsSealed)
            {
                AlertForm.Show(ParentForm, "搜索尚未完成，请等待搜索完成后再导出", "提示", AlertType.Information);
                return;
            }

            if (results == null || results.Count == 0)
            {
                StartSearchForExport(keyword, false);
                return;
            }

            // 已有完整搜索结果且条件匹配 → 直接导出
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
            _autoJumpTried = 0;
            Volatile.Write(ref _searchProgressQueued, 0);
            // ★ 建空容器进缓存，引用直接交给后台任务，后台不再回查缓存链表
            var container = new SearchResultSet();
            PutRecentSearch(currentKey, container);

            Status($"正在搜索 \"{keyword}\"...");
            if (reverse)
                _ = Task.Run(() => DoSearchAndReverseExport(keyword, caseSensitive, useRegex, token, container));
            else
                _ = Task.Run(() => DoSearchAndExport(keyword, caseSensitive, useRegex, token, container));
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

            // ★ 未封口就不能反向导出：搜索还在跑时，尚未扫到的行会被当成"不匹配"写进产物
            if (results != null && !results.IsSealed)
            {
                AlertForm.Show(ParentForm, "搜索尚未完成，请等待搜索完成后再反向导出", "提示", AlertType.Information);
                return;
            }

            if (results == null || results.Count == 0)
            {
                StartSearchForExport(keyword, true);
                return;
            }

            // 已有完整搜索结果 → 直接反向导出
            DoReverseExport(results);
        }

        /// <summary>搜索完成后自动反向导出</summary>
        private void DoSearchAndReverseExport(string keyword, bool caseSensitive, bool useRegex, CancellationToken token, SearchResultSet results)
        {
            try
            {
                string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";
                Func<string, bool> matcher = BuildLineMatcher(keyword, caseSensitive, useRegex);

                // 导出要的是完整匹配集，扫描期间不做自动跳转，免得界面被带着跑
                if (ScanLinesSequential(results, currentKey, matcher, token, curLine: -1, autoJump: false)) return;

                SafeBeginInvoke(() =>
                {
                    // ★ 如果用户在搜索期间切换了关键词，丢弃本次结果
                    if (_lastSearch != currentKey) return;

                    // ★ 写入 FileCacheStore 持久化缓存（内存存引用，磁盘后台分段写）
                    _ = FileCacheStore.PutSearchResultAsync(_filePath!, currentKey, results);
                    if (results.Count > 0)
                    {
                        NavigateTo(results, results.IndexOfNearest(GetCurrentLine()));
                        // "正在搜索"是过程消息，命中后底部状态栏已经报了几处/第几处，这里收掉别让它常驻
                        Status("");
                        DoReverseExport(results);
                    }
                    else
                    {
                        _searchCurrentIndex = -1;
                        _lastNavigatedMatchLine = -1;
                        _lblSearchPos.Text = "未找到匹配，无法反向导出";
                        Status("未找到匹配行");
                    }
                });
            }
            catch { }
            finally
            {
                ReleaseSearchCts(token);
            }
        }

        /// <summary>反向导出：利用已有的搜索结果索引，导出所有不匹配的行</summary>
        private void DoReverseExport(SearchResultSet results)
        {
            string? sourcePath = _filePath;
            if (results.Count == 0 || string.IsNullOrEmpty(sourcePath)) return;

            // 段快照：发布出去的段不会再改，直接共享引用，不必再拷一份大列表
            var segments = results.SnapshotSegments();
            var lineOffsets = _lineOffsets!;
            long fileLength = _fileLength;
            string exportSource = sourcePath;
            // ★ 与正向导出共用并发闸：已有导出在跑就提示等待，不取消覆盖
            if (Interlocked.CompareExchange(ref _exportRunning, 1, 0) != 0)
            {
                AlertForm.Show(ParentForm, "当前有导出正在运行，请稍后重试", "提示", AlertType.Information);
                return;
            }
            _exportCts = new CancellationTokenSource();
            var exportToken = _exportCts.Token;
            string dir = Path.GetDirectoryName(sourcePath)!;
            string fileName = Path.GetFileName(sourcePath);
            string sourceExt = Path.GetExtension(sourcePath);
            // ★ 从 _lastSearch 中取关键词部分（格式 "keyword|caseSensitive|useRegex"）
            string filterPart = SanitizeFileName(_lastSearch.Split('|')[0], 50);
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            ShowCancelExport();
            Status("正在准备反向导出...");

            _ = Task.Run(() =>
            {
                // 计算反向导出的总行数 = 总行数 - 匹配行数
                long totalLines = lineOffsets.Length;
                long excludedCount = results.Count;
                long exportCount = totalLines - excludedCount;

                // 命名规则：原始文件名.过滤字符_排除.时间戳.源文件后缀
                string safeFileName;
                if (!string.IsNullOrEmpty(filterPart))
                    safeFileName = $"{SanitizeFileName(fileName, 200)}.{filterPart}_排除.{timestamp}{sourceExt}";
                else
                    safeFileName = $"{SanitizeFileName(fileName, 200)}_排除.{timestamp}{sourceExt}";

                string outName = safeFileName;
                string outPath = Path.Combine(dir, outName);

                ExportUi(() => Status($"正在反向导出 {exportCount} 行到 {outName}..."));

                StreamWriter? sw = null;
                FileStream? fs = null;
                long exported = 0;
                bool cancelled = false;
                bool done = false;
                var progSw = Stopwatch.StartNew();
                long lastProgMs = 0;

                // 匹配行号游标：段内、段间都是升序，顺着推即可，
                // 旧实现先 ToList() 拷一份再用下标走，千万级匹配等于白拷几十 MB
                int segIdx = 0, segOff = -1;
                long NextMatch()
                {
                    segOff++;
                    while (segIdx < segments.Length)
                    {
                        if (segOff < segments[segIdx].Length) return segments[segIdx][segOff];
                        segIdx++;
                        segOff = 0;
                    }
                    return long.MaxValue;
                }
                long matchLine = NextMatch();
                try
                {
                    fs = new FileStream(exportSource, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.RandomAccess);
                    sw = new StreamWriter(outPath, false, Encoding.UTF8);

                    for (long lineIdx = 0; lineIdx < totalLines; lineIdx++)
                    {
                        if (exportToken.IsCancellationRequested) { cancelled = true; break; }
                        // 如果当前行在匹配列表中 → 跳过
                        if (matchLine == lineIdx)
                        {
                            do { matchLine = NextMatch(); } while (matchLine == lineIdx);
                            continue;
                        }
                        // 不在匹配列表中 → 输出
                        string? lineText = ReadLineForExport(fs, lineOffsets, fileLength, lineIdx);
                        if (lineText == null) throw new EndOfStreamException($"源文件第 {lineIdx + 1} 行读取失败");
                        sw.WriteLine(lineText);
                        exported++;
                        if ((exported & 0xFFFF) == 0)
                            ReportExportProgress("反向导出中", exported, exportCount, progSw, ref lastProgMs);
                    }
                    sw.Flush();
                    if (cancelled)
                    {
                        ExportUi(() => Status("反向导出已取消"));
                        return;
                    }
                    done = true;

                    ExportUi(() =>
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
                    ExportUi(() =>
                    {
                        AlertForm.Show(ParentForm, $"反向导出失败: {ex.Message}", "错误", AlertType.Error, exception: ex);
                        Status("反向导出失败");
                    });
                }
                finally
                {
                    try { sw?.Dispose(); } catch { }
                    try { fs?.Dispose(); } catch { }
                    if (!done) { try { File.Delete(outPath); } catch { } }
                    ReleaseExport(exportToken);
                }
            });
        }

        /// <summary>搜索完成后自动导出</summary>
        private void DoSearchAndExport(string keyword, bool caseSensitive, bool useRegex, CancellationToken token, SearchResultSet results)
        {
            try
            {
                string currentKey = $"{keyword}|{caseSensitive}|{useRegex}";
                Func<string, bool> matcher = BuildLineMatcher(keyword, caseSensitive, useRegex);

                if (ScanLinesSequential(results, currentKey, matcher, token, curLine: -1, autoJump: false)) return;

                SafeBeginInvoke(() =>
                {
                    // ★ 如果用户在搜索期间切换了关键词，丢弃本次结果
                    if (_lastSearch != currentKey) return;

                    // ★ 写入 FileCacheStore 持久化缓存（内存存引用，磁盘后台分段写）
                    _ = FileCacheStore.PutSearchResultAsync(_filePath!, currentKey, results);
                    if (results.Count > 0)
                    {
                        NavigateTo(results, results.IndexOfNearest(GetCurrentLine()));
                        DoExport(results);
                    }
                    else
                    {
                        _searchCurrentIndex = -1;
                        _lastNavigatedMatchLine = -1;
                        _lblSearchPos.Text = "未找到匹配，无法导出";
                        Status("未找到匹配行");
                    }
                });
            }
            catch { }
            finally
            {
                ReleaseSearchCts(token);
            }
        }

        private void DoExport(SearchResultSet results)
        {
            string? sourcePath = _filePath;
            if (results.Count == 0 || string.IsNullOrEmpty(sourcePath)) return;

            int contextLines = ContextLines;

            // 段快照：发布出去的段不会再改，直接共享引用，无需拷成一份大列表
            var segments = results.SnapshotSegments();
            var lineOffsets = _lineOffsets!;
            long fileLength = _fileLength;
            string exportSource = sourcePath;
            // ★ 同一时刻只跑一个导出：正向/反向共用一个闸。旧实现这里 Cancel 上一个再开新的，
            //   用户连点两个按钮会把先发起的那次悄悄取消、产物被删，看起来像"被覆盖"
            if (Interlocked.CompareExchange(ref _exportRunning, 1, 0) != 0)
            {
                AlertForm.Show(ParentForm, "当前有导出正在运行，请稍后重试", "提示", AlertType.Information);
                return;
            }
            _exportCts = new CancellationTokenSource();
            var exportToken = _exportCts.Token;
            string dir = Path.GetDirectoryName(sourcePath)!;
            string fileName = Path.GetFileName(sourcePath);
            string sourceExt = Path.GetExtension(sourcePath);
            // ★ 从 _lastSearch 中取关键词部分（格式 "keyword|caseSensitive|useRegex"）
            string filterPart = SanitizeFileName(_lastSearch.Split('|')[0], 50);
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");

            ShowCancelExport();
            Status("正在准备导出...");

            // 所有耗时操作都在后台线程执行，避免 UI 卡死
            _ = Task.Run(() =>
            {
                // ===== 以下全部在后台线程执行 =====

                // 合并重叠的上下文区间，避免重复输出。
                // 匹配行号本来就全局升序，所以 [line-ctx, line+ctx] 也是升序，顺路合并即可；
                // 旧实现先给每个匹配建一条区间再整体 Sort，800w 匹配就是 128MB 加一次
                // 千万级排序，导出没卡死过才怪。
                var mergedRanges = new List<(long Start, long End)>();
                long lastStart = 0, lastEnd = -1;
                long lastLine = lineOffsets.Length - 1;
                foreach (var seg in segments)
                {
                    foreach (long lineIdx in seg)
                    {
                        long start = lineIdx > contextLines ? lineIdx - contextLines : 0;
                        long end = Math.Min(lastLine, lineIdx + contextLines);
                        if (lastEnd >= 0 && start <= lastEnd + 1)
                        {
                            if (end > lastEnd) lastEnd = end; // 与当前区间相邻/重叠 → 就地延长
                        }
                        else
                        {
                            if (lastEnd >= 0) mergedRanges.Add((lastStart, lastEnd));
                            lastStart = start;
                            lastEnd = end;
                        }
                    }
                }
                if (lastEnd >= 0) mergedRanges.Add((lastStart, lastEnd));

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
                ExportUi(() => Status($"正在导出 {totalLines} 行到 {outName}..."));

                StreamWriter? sw = null;
                FileStream? fs = null;
                bool cancelled = false;
                bool done = false;
                long exported = 0;
                var progSw = Stopwatch.StartNew();
                long lastProgMs = 0;
                try
                {
                    // 自带流 + 快照偏移：页签中途关掉也不影响这次导出的正确性
                    fs = new FileStream(exportSource, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.RandomAccess);
                    sw = new StreamWriter(outPath, false, Encoding.UTF8);
                    foreach (var (start, end) in mergedRanges)
                    {
                        for (long lineIdx = start; lineIdx <= end; lineIdx++)
                        {
                            if (exportToken.IsCancellationRequested) { cancelled = true; break; }
                            string? lineText = ReadLineForExport(fs, lineOffsets, fileLength, lineIdx);
                            if (lineText == null) throw new EndOfStreamException($"源文件第 {lineIdx + 1} 行读取失败");
                            sw.WriteLine(lineText);
                            exported++;
                            if ((exported & 0xFFFF) == 0)
                                ReportExportProgress("导出中", exported, totalLines, progSw, ref lastProgMs);
                        }
                        if (cancelled) break;
                    }
                    sw.Flush();
                    if (cancelled)
                    {
                        ExportUi(() => Status("导出已取消"));
                        return;
                    }
                    done = true;

                    ExportUi(() =>
                    {
                        Status($"已导出 {totalLines} 行 -> {outName}");
                        AlertForm.Show(ParentForm,
                            $"导出完成！" + GlobalConst.NewLine + GlobalConst.NewLine + $"文件: {outName}" + GlobalConst.NewLine + $"路径: {dir}" + GlobalConst.NewLine + $"行数: {totalLines:N0}" + GlobalConst.NewLine + $"匹配行: {results.Count:N0}" + GlobalConst.NewLine + $"上下行数: {contextLines}",
                            "导出成功", AlertType.Information);
                        try { FileExported?.Invoke(this, EventArgs.Empty); } catch { }
                    });
                }
                catch (Exception ex)
                {
                    ExportUi(() =>
                    {
                        AlertForm.Show(ParentForm, $"导出失败: {ex.Message}", "错误", AlertType.Error, exception: ex);
                        Status("导出失败");
                    });
                }
                finally
                {
                    try { sw?.Dispose(); } catch { }
                    try { fs?.Dispose(); } catch { }
                    // 半截文件（被取消或读失败）一律删掉：留着会被当成一次成功的导出
                    if (!done) { try { File.Delete(outPath); } catch { } }
                    ReleaseExport(exportToken);
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

        /// <summary>检测行数是否超过切割阈值，弹窗提示用户进行文件切割</summary>
        private void CheckAndPromptSplit(long totalLines)
        {
            if (_lineOffsets == null || string.IsNullOrEmpty(_filePath)) return;

            string baseName = Path.GetFileNameWithoutExtension(_filePath);
            string ext = Path.GetExtension(_filePath);
            int fileCount = FileSplitter.GetFileCount(totalLines, SPLIT_LINES_PER_FILE);

            var result = MessageBox.Show(
                ParentForm,
                $"文件 {Path.GetFileName(_filePath)} 总行数 {totalLines:N0}，\n" +
                $"超过显示上限，将按照每 {SPLIT_LINES_PER_FILE:N0} 行一个文件进行切割。\n\n" +
                $"切割后文件示例：\n" +
                $"  {baseName}-0{ext}\n" +
                $"  {baseName}-1{ext}\n" +
                $"  ...\n\n" +
                $"共切割为 {fileCount} 个文件，是否继续？",
                "文件切割",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning);

            if (result != DialogResult.OK) return;

            // CSV/TSV 切割后在第 2 个及以后的分片顶部补一份表头，避免单独打开时丢失列名
            int headerLineCount = 0;
            string extLower = Path.GetExtension(_filePath).ToLowerInvariant();
            if (extLower is ".csv" or ".tsv")
            {
                var keepHeader = AlertForm.ShowConfirm(ParentForm!,
                    "是否在每个分片文件顶部保留首行作为表头？\n\n" +
                    "选择“确定”保留（推荐），选择“取消”不保留。",
                    "文件切割", AlertType.Information);
                if (keepHeader == DialogResult.Yes) headerLineCount = 1;
            }

            Status("正在切割文件...");
            _lblFileLines.Text = "行数：切割中...";

            // 后台执行切割
            Task.Run(() =>
            {
                try
                {
                    var splitFiles = FileSplitter.Split(
                        _filePath, _lineOffsets, _fileLength, SPLIT_LINES_PER_FILE,
                        (current, total, percent) =>
                        {
                            BeginInvoke(() =>
                            {
                                Status($"正在切割文件... ({current}/{total}, {percent}%)");
                                _lblFileLines.Text = $"行数：切割中 {percent}%";
                            });
                        },
                        headerLineCount);

                    // 切割完成后在 UI 线程提示
                    BeginInvoke(() =>
                    {
                        Status($"切割完成，共生成 {splitFiles.Count} 个文件");
                        _lblFileLines.Text = "行数：已切割";
                        AlertForm.Show(ParentForm,
                            $"文件切割完成！\n\n" +
                            $"原始文件：{Path.GetFileName(_filePath)}\n" +
                            $"总行数：{totalLines:N0}\n" +
                            $"切割文件数：{splitFiles.Count}\n\n" +
                            $"请分别打开切割后的文件进行查看。",
                            "切割完成", AlertType.Information);
                    });
                }
                catch (Exception ex)
                {
                    BeginInvoke(() =>
                    {
                        Status("文件切割失败");
                        _lblFileLines.Text = "行数：切割失败";
                        AlertForm.Show(ParentForm, $"文件切割失败：{ex.Message}", "错误", AlertType.Error);
                    });
                }
            });
        }

        public void Cleanup()
        {
            _cancel = true;
            try { _timer.Stop(); } catch { }
            try { _timerScroll?.Stop(); } catch { }
            // ★ 取消并行索引构建任务
            try { _indexCts?.Cancel(); } catch { }
            // ★ 取消进行中的导出：源信息马上要失效，让它跑完只会写出半截文件
            try { _exportCts?.Cancel(); } catch { }
            // ★ 提前保存文件路径，用于卸载内存缓存（_filePath 接下来会被置 null）
            string? oldPath = _filePath;
            // ★ 提前置 null，防止 OnCache/OnRetrieve 在清理过程中访问已释放的资源
            _indexReady = false;
            _partialMaxLine = 0;
            _completedBlocks = 0;
            _filePath = null;
            _lineOffsets = null;
            _blockResults.Clear();
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
            // ★ 先清空总行数，减少 Dispose 时的清理开销
            try { _lstContent.SetTotalLines(0); } catch { }
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
                    // 详情框一并清空，换文件/关页签后不留上一行的内容
                    _txtDetail.Text = "";
                    _detailTip.SetToolTip(_txtDetail, "");
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
                _exportCts?.Dispose();
                // ToolTip 是组件、不在 Controls 集合里，控件销毁不会连带释放，需手动 Dispose
                _detailTip?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}