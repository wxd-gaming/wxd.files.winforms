using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// HTTP 页签：轻量请求工具（不追求完整 Postman）。
    /// 顶部工具条选方法 + 填 URL + 请求次数 / 线程数；
    /// 参数填写区左右分栏：右侧一条页签栏共五页——参数（单一键值网格，GET 拼 URL，
    /// POST/PUT 按 Body 单选组请求体）/ 请求头（键值可下拉 + 编码）/ 响应 / 请求 / 响应头；
    /// gzip / deflate 自动解压并自带 Accept-Encoding；次数 &gt; 1 按线程数并发发同一请求，响应区统计 QPS 与延迟。
    /// 左右分栏：左侧「已保存」列表（选中即加载，输入名称后保存覆盖），右侧请求详情；
    /// 整套输入存进 .wxd-work/http-requests.json。
    /// </summary>
    internal sealed class HttpToolPanel : ToolPanelBase
    {
        private static readonly HttpClient Client = new(new HttpClientHandler
        {
            // 自动带 Accept-Encoding: gzip, deflate，并透明解压响应体
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
        })
        { Timeout = TimeSpan.FromSeconds(30) };

        private static readonly string[] CommonHeaderKeys =
        {
            "Accept", "Accept-Encoding", "Accept-Language", "Authorization", "Cache-Control",
            "Connection", "Content-Encoding", "Content-Length", "Content-Type", "Cookie",
            "Host", "If-None-Match", "Origin", "Referer", "User-Agent",
            "X-Forwarded-For", "X-Requested-With"
        };

        // 值列的候选按当前行的键给；没列出的键自由输入，不限制
        private static readonly Dictionary<string, string[]> ValueSuggestions = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Content-Type"] = new[] { "multipart/form-data", "application/x-www-form-urlencoded", "application/json", "text/plain", "application/xml" },
            ["User-Agent"] = new[] { "Mozilla/5.0 (Windows NT 10.0; Win64; x64)", "PostmanRuntime/7.42.0" },
            ["Accept"] = new[] { "*/*", "application/json", "text/html" },
            ["Accept-Language"] = new[] { "zh-CN", "en-US" },
            ["Accept-Encoding"] = new[] { "gzip, deflate", "identity" },
        };

        private readonly ThemeComboBox _method;
        private readonly ThemeComboBox _encoding;
        private readonly ThemeField _url;
        private readonly ThemeField _count;
        private readonly ThemeField _threads;
        private readonly SplitContainer _split = new();
        private readonly TableLayoutPanel _savedPane = new();
        private readonly ListBox _savedList = new();
        private readonly TextBox _saveName = new();
        private bool _splitInitDone;
        private readonly List<SavedHttpRequest> _saved = new();
        // Body 单选的取值，下标对应 _bodyType；"" 表示 none（无请求体）
        private static readonly string[] BodyTypeValues =
            { "", "application/json", "application/x-www-form-urlencoded", "multipart/form-data" };
        private readonly ThemeTabControl _inputTabs = new();
        private readonly TableLayoutPanel _paramsPane = new();
        private readonly TableLayoutPanel _headersPane = new();
        private FlowLayoutPanel _bodyTypeRow = new();
        private readonly ThemeRadioButton[] _bodyType = new ThemeRadioButton[4];
        private readonly Label _noneHint;
        // 单一键值网格：GET / DELETE 自动拼到 URL，POST / PUT 按 Body 单选组请求体
        private readonly DataGridView _params = NewPairsGrid();
        private readonly DataGridView _headers = NewPairsGrid(suggestHeaderKeys: true);
        private readonly TableLayoutPanel _reqLayout = new();
        private readonly TableLayoutPanel _respHeadLayout = new();
        private readonly TextBox _reqLine = MakeEditor(readOnly: true);
        private readonly TextBox _reqBody = MakeEditor(readOnly: true);
        private readonly DataGridView _reqHeaders = NewReadOnlyPairsGrid();
        private readonly DataGridView _respHeaders = NewReadOnlyPairsGrid();
        private readonly Label _respMeta = MakeHint("");
        private readonly ThemeBordered _respBorder;
        private readonly TextBox _respBody = MakeEditor(readOnly: true);
        private readonly TableLayoutPanel _content = new();
        private readonly ThemeButton _sendButton;
        private readonly ThemeButton _saveButton;
        private readonly ThemeButton _deleteButton;
        private bool _sending;
        private bool _savedUpdating;

        internal HttpToolPanel()
        {
            _savedList.BorderStyle = BorderStyle.None;
            _savedList.DrawMode = DrawMode.OwnerDrawFixed;
            _savedList.ItemHeight = 24;
            _savedList.IntegralHeight = false;
            _savedList.Font = Theme.Body;
            _savedList.DrawItem += DrawSavedItem;
            _savedList.SelectedIndexChanged += (_, _) => LoadSavedSelection();

            _saveName.BorderStyle = BorderStyle.None;
            _saveName.Font = Theme.Body;

            _method = MakeTopCombo(new object[] { "GET", "POST", "PUT", "DELETE" }, 86);
            _method.SelectedIndex = 0;
            Top.Controls.Add(_method);

            _url = AddField("https://example.com/api", 400);

            Top.Controls.Add(MakeTopLabel("次数"));
            _count = AddField("1", 44);
            _count.Edit.Text = "1";
            Top.Controls.Add(MakeTopLabel("线程"));
            _threads = AddField("1", 44);
            _threads.Edit.Text = "1";

            // Body 类型单选组：none / json / urlencoded / multipart（互斥靠同一个父容器）
            // 只对 POST / PUT 有意义，GET / DELETE 时整行随方法切换隐藏
            string[] bodyTypeLabels = { "none", "application/json", "x-www-form-urlencoded", "multipart/form-data" };
            _bodyTypeRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Theme.P.Window,
                Margin = new Padding(0),
                WrapContents = false,
                Padding = new Padding(ThemeMetrics.Margin, 6, 0, 0)
            };
            for (int i = 0; i < bodyTypeLabels.Length; i++)
            {
                var radio = new ThemeRadioButton
                {
                    Text = bodyTypeLabels[i],
                    AutoSize = true,
                    Font = Theme.Body,
                    ForeColor = Theme.P.Text,
                    BackColor = Theme.P.Window,
                    Margin = new Padding(0, 4, 14, 0)
                };
                _bodyType[i] = radio;
                _bodyTypeRow.Controls.Add(radio);
            }
            _bodyType[3].Checked = true; // 默认 multipart/form-data
            // 初值设完再绑事件：CheckedChanged 里会切网格可见性，那时页面控件还没建好
            foreach (var radio in _bodyType) radio.CheckedChanged += (_, _) => UpdateBodyEditor();

            _encoding = MakeTopCombo(new object[] { "UTF-8", "GBK", "GB18030" }, 76);
            _encoding.SelectedIndex = 0;

            _sendButton = AddButton("发送", (_, _) => SendRequest(), primary: true);
            AddButton("格式化响应", (_, _) => PrettyPrintResponse());
            AddButton("复制结果", (_, _) => CopyToClipboard(_respBody.Text));

            _saveButton = MakeContentButton("保存", (_, _) => SaveCurrent());
            _deleteButton = MakeContentButton("删除", (_, _) => DeleteSaved());

            _content.Dock = DockStyle.Fill;
            _content.BackColor = Theme.P.Window;
            _content.ColumnCount = 1;
            _content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // 参数填写区 + 响应详情共用一条页签栏：参数 / 请求头 / 响应 / 请求 / 响应头
            _inputTabs.Dock = DockStyle.Fill;
            _inputTabs.Margin = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Margin, ThemeMetrics.Gap);
            _inputTabs.DrawMode = TabDrawMode.OwnerDrawFixed;
            _inputTabs.SizeMode = TabSizeMode.Fixed;
            _inputTabs.ItemSize = new Size(132, 26);
            _inputTabs.Padding = new Point(6, 2);
            _inputTabs.DrawItem += DrawToolTab;

            _paramsPane.Dock = DockStyle.Fill;
            _paramsPane.BackColor = Theme.P.Window;
            _paramsPane.ColumnCount = 1;
            _paramsPane.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            _paramsPane.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            _paramsPane.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _paramsPane.Controls.Add(MakeHint("参数（GET / DELETE 拼到 URL；POST / PUT 按下方类型组请求体）"), 0, 0);
            _bodyTypeRow.Dock = DockStyle.Fill;
            _bodyTypeRow.BackColor = Theme.P.Window;
            _bodyTypeRow.WrapContents = false;
            _bodyTypeRow.Padding = new Padding(ThemeMetrics.Margin, 6, 0, 0);
            _paramsPane.Controls.Add(_bodyTypeRow, 0, 1);
            var bodyEditorHost = new Panel { Dock = DockStyle.Fill, BackColor = Theme.P.Window, Margin = new Padding(0) };
            _noneHint = MakeHint("none：无请求体。请求头里手写 Content-Type 时，参数会按原样 k=v&k=v 组体（游戏服手写解析用）");
            _noneHint.Dock = DockStyle.Fill;
            _noneHint.TextAlign = ContentAlignment.MiddleCenter;
            _noneHint.Visible = false;
            bodyEditorHost.Controls.Add(new ThemeBordered(_params)
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Margin, ThemeMetrics.Gap)
            });
            bodyEditorHost.Controls.Add(_noneHint);
            _paramsPane.Controls.Add(bodyEditorHost, 0, 2);

            _headersPane.Dock = DockStyle.Fill;
            _headersPane.BackColor = Theme.P.Window;
            _headersPane.ColumnCount = 1;
            _headersPane.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            _headersPane.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var headersHintRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Theme.P.Window,
                Margin = new Padding(0),
                WrapContents = false
            };
            headersHintRow.Controls.Add(MakeHint("请求头（键、值可下拉选择，也可自定义）"));
            headersHintRow.Controls.Add(MakeHint("编码"));
            headersHintRow.Controls.Add(_encoding);
            _headersPane.Controls.Add(headersHintRow, 0, 0);
            _headersPane.Controls.Add(new ThemeBordered(_headers)
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Margin, ThemeMetrics.Gap)
            }, 0, 1);

            var paramsPage = new TabPage("参数 (Params)") { BackColor = Theme.P.Window };
            paramsPage.Controls.Add(_paramsPane);
            var headersPage = new TabPage("请求头 (Headers)") { BackColor = Theme.P.Window };
            headersPage.Controls.Add(_headersPane);
            _inputTabs.TabPages.Add(paramsPage);
            _inputTabs.TabPages.Add(headersPage);
            // 初值设完再绑事件：UpdateBodyEditor 会按方法切单选行可见性，这时页面控件必须已建好
            _method.SelectedIndexChanged += (_, _) => UpdateBodyEditor();
            UpdateBodyEditor();

            _respBorder = new ThemeBordered(_respBody)
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Margin, ThemeMetrics.Margin)
            };

            _reqLayout.Dock = DockStyle.Fill;
            _reqLayout.BackColor = Theme.P.Window;
            _reqLayout.ColumnCount = 1;
            _reqLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            _reqLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            _reqLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 110));
            _reqLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            _reqLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _reqLayout.Controls.Add(new ThemeBordered(_reqLine)
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Margin, ThemeMetrics.Gap)
            }, 0, 0);
            _reqLayout.Controls.Add(MakeHint("发送的请求头（Content-Type 为实际发送值，multipart 含 boundary）"), 0, 1);
            _reqLayout.Controls.Add(_reqHeaders, 0, 2);
            _reqLayout.Controls.Add(MakeHint("请求体"), 0, 3);
            _reqLayout.Controls.Add(new ThemeBordered(_reqBody)
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Margin, ThemeMetrics.Margin)
            }, 0, 4);

            _respHeadLayout.Dock = DockStyle.Fill;
            _respHeadLayout.BackColor = Theme.P.Window;
            _respHeadLayout.ColumnCount = 1;
            _respHeadLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            _respHeadLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _respMeta.Margin = new Padding(ThemeMetrics.Margin, 4, ThemeMetrics.Margin, 0);
            _respHeadLayout.Controls.Add(_respMeta, 0, 0);
            _respHeadLayout.Controls.Add(_respHeaders, 0, 1);

            var responsePage = new TabPage("响应") { BackColor = Theme.P.Window };
            responsePage.Controls.Add(_respBorder);
            var requestPage = new TabPage("请求") { BackColor = Theme.P.Window };
            requestPage.Controls.Add(_reqLayout);
            var respHeadPage = new TabPage("响应头") { BackColor = Theme.P.Window };
            respHeadPage.Controls.Add(_respHeadLayout);
            _inputTabs.TabPages.Add(responsePage);
            _inputTabs.TabPages.Add(requestPage);
            _inputTabs.TabPages.Add(respHeadPage);
            _content.Controls.Add(_inputTabs, 0, 0);

            // 左右分栏：左侧已保存请求列表（可拖分隔条），右侧请求详情
            _split.Dock = DockStyle.Fill;
            _split.Orientation = Orientation.Vertical;
            _split.FixedPanel = FixedPanel.Panel1;
            _split.Panel1MinSize = 170;
            _split.SplitterWidth = 4;
            // 构造期控件还没拿到真实宽度，直接设 SplitterDistance 会越界；
            // 等首次布局把宽度撑起来后设一次初始值，之后交给用户拖动
            _split.Resize += (_, _) =>
            {
                if (!_splitInitDone && _split.Width > 560)
                {
                    _splitInitDone = true;
                    _split.SplitterDistance = 218;
                }
            };

            _savedPane.Dock = DockStyle.Fill;
            _savedPane.BackColor = Theme.P.Window;
            _savedPane.ColumnCount = 1;
            _savedPane.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            _savedPane.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _savedPane.RowStyles.Add(new RowStyle(SizeType.Absolute, ThemeMetrics.RowH + ThemeMetrics.Gap));
            _savedPane.Controls.Add(MakeHint("已保存（点击加载；输入名称后保存）"), 0, 0);
            _savedPane.Controls.Add(new ThemeBordered(_savedList)
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(ThemeMetrics.Margin, 0, ThemeMetrics.Gap, ThemeMetrics.Gap)
            }, 0, 1);
            var saveRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(ThemeMetrics.Margin, 0, 0, 0),
                WrapContents = false
            };
            saveRow.Controls.Add(new ThemeBordered(_saveName)
            {
                Width = 80,
                Height = ThemeMetrics.RowH,
                Margin = new Padding(0, 0, ThemeMetrics.Gap, 0)
            });
            saveRow.Controls.Add(_saveButton);
            saveRow.Controls.Add(_deleteButton);
            _savedPane.Controls.Add(saveRow, 0, 2);

            _split.Panel1.Controls.Add(_savedPane);
            _split.Panel2.Controls.Add(_content);

            FinishChrome(_split);

            LoadSavedFile();

            SetStatus("GET 参数拼 URL，POST / PUT 按 Content-Type 组请求体；次数 > 1 按线程数并发并统计 QPS，30 秒超时");
        }

        /// <summary>参数填写区与详情区页签共用的自绘：圆角选中块 + 皮肤配色。</summary>
        private void DrawToolTab(object? sender, DrawItemEventArgs e)
        {
            var tabs = (TabControl)sender!;
            var tabPage = tabs.TabPages[e.Index];
            var tabRect = tabs.GetTabRect(e.Index);
            bool isSelected = tabs.SelectedIndex == e.Index;

            // 页签本体四周留了 2px 缝，先按皮肤底色铺一层，否则缝里露的是 comctl32 的系统灰。
            // 只能铺自己这一格：DrawItem 的 Graphics 没裁到本页签矩形，铺整条会把别的页签刷掉。
            using (var bandBrush = new SolidBrush(Theme.P.Window))
                e.Graphics.FillRectangle(bandBrush, tabRect);

            int bottomExtra = isSelected ? 2 : 0;
            var roundRect = new Rectangle(
                tabRect.X + 2, tabRect.Y + 2, tabRect.Width - 4, tabRect.Height - 2 + bottomExtra);
            int radius = Theme.P.Radius + 2;

            using (var path = new GraphicsPath())
            {
                path.AddArc(roundRect.X, roundRect.Y, radius * 2, radius * 2, 180, 90);
                path.AddArc(roundRect.Right - radius * 2, roundRect.Y, radius * 2, radius * 2, 270, 90);
                path.AddLine(roundRect.Right, roundRect.Y + radius, roundRect.Right, roundRect.Bottom);
                path.AddLine(roundRect.Right, roundRect.Bottom, roundRect.X, roundRect.Bottom);
                path.AddLine(roundRect.X, roundRect.Bottom, roundRect.X, roundRect.Y + radius);
                path.CloseFigure();

                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                var bgColor = isSelected ? Theme.P.Accent : Theme.P.Selection;
                var borderColor = isSelected ? Theme.P.Accent : Theme.P.Border;
                using (var brush = new SolidBrush(bgColor)) e.Graphics.FillPath(brush, path);
                using (var pen = new Pen(borderColor, 1f)) e.Graphics.DrawPath(pen, path);
            }

            var textRect = new Rectangle(tabRect.X + 10, tabRect.Y + 4, tabRect.Width - 20, tabRect.Height - 8);
            Color textColor = isSelected ? Theme.P.AccentText : Theme.P.Text;
            TextRenderer.DrawText(e.Graphics, tabPage.Text, tabs.Font, textRect, textColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
        }

        protected override void ApplyThemeCore()
        {
            _content.BackColor = Theme.P.Window;
            _respBorder.ApplyTheme();
            ApplyComboTheme(_method);
            ApplyComboTheme(_encoding);
            _split.BackColor = Theme.P.Border;
            _split.Panel1.BackColor = Theme.P.Window;
            _split.Panel2.BackColor = Theme.P.Window;
            _savedPane.BackColor = Theme.P.Window;
            _savedList.Font = Theme.Body;
            _savedList.BackColor = Theme.P.Surface;
            _savedList.ForeColor = Theme.P.Text;
            _savedList.Invalidate();
            _saveName.Font = Theme.Body;
            _saveName.BackColor = Theme.P.Surface;
            _saveName.ForeColor = Theme.P.Text;
            _inputTabs.Font = Theme.Body;
            _inputTabs.Invalidate();
            foreach (TabPage page in _inputTabs.TabPages) page.BackColor = Theme.P.Window;
            _paramsPane.BackColor = Theme.P.Window;
            _headersPane.BackColor = Theme.P.Window;
            _bodyTypeRow.BackColor = Theme.P.Window;
            _noneHint.Parent!.BackColor = Theme.P.Window; // 参数网格宿主面板（none 提示的父容器）
            foreach (var radio in _bodyType)
            {
                radio.Font = Theme.Body;
                radio.ForeColor = Theme.P.Text;
                radio.BackColor = Theme.P.Window;
            }
            _noneHint.Font = Theme.Caption;
            _noneHint.ForeColor = Theme.P.Muted;
            _noneHint.BackColor = Theme.P.Window;
            _reqLayout.BackColor = Theme.P.Window;
            _respHeadLayout.BackColor = Theme.P.Window;
            ApplyEditorTheme(_reqLine);
            ApplyEditorTheme(_reqBody);
            ApplyEditorTheme(_respBody);
            ApplyGridTheme(_params);
            ApplyGridTheme(_headers);
            ApplyGridTheme(_reqHeaders);
            ApplyGridTheme(_respHeaders);
            _respMeta.Font = Theme.Caption;
            _respMeta.ForeColor = Theme.P.Muted;
            _respMeta.BackColor = Theme.P.Window;
            _saveButton.Font = Theme.Body;
            _saveButton.BackColor = Theme.P.Window;
            _saveButton.Invalidate();
            _deleteButton.Font = Theme.Body;
            _deleteButton.BackColor = Theme.P.Window;
            _deleteButton.Invalidate();
            foreach (Control control in Top.Controls)
                if (control is Label topLabel)
                {
                    topLabel.Font = Theme.Caption;
                    topLabel.ForeColor = Theme.P.Muted;
                }
            foreach (Control control in _content.Controls)
            {
                if (control is Label label)
                {
                    label.Font = Theme.Caption;
                    label.ForeColor = Theme.P.Muted;
                    label.BackColor = Theme.P.Window;
                }
            }
            // 两个页签页内部的提示行（参数 / 请求头页；含 Body 单选行与 none 提示所在容器）
            foreach (var pane in new[] { _paramsPane, _headersPane })
                foreach (Control control in pane.Controls)
                    WalkHintLabels(control);
            foreach (Control control in _savedPane.Controls)
                if (control is Label savedLabel)
                {
                    savedLabel.Font = Theme.Caption;
                    savedLabel.ForeColor = Theme.P.Muted;
                    savedLabel.BackColor = Theme.P.Window;
                }
        }

        private static void WalkHintLabels(Control control)
        {
            if (control is Label label)
            {
                label.Font = Theme.Caption;
                label.ForeColor = Theme.P.Muted;
                label.BackColor = Theme.P.Window;
                return;
            }
            foreach (Control child in control.Controls) WalkHintLabels(child);
        }

        /// <summary>已保存列表自绘：选中行用皮肤选中色，超宽名称省略号截断。</summary>
        private void DrawSavedItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _savedList.Items.Count) return;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            using (var brush = new SolidBrush(selected ? Theme.P.Selection : _savedList.BackColor))
                e.Graphics.FillRectangle(brush, e.Bounds);
            TextRenderer.DrawText(e.Graphics, Convert.ToString(_savedList.Items[e.Index]), e.Font, e.Bounds,
                Theme.P.Text, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        private int BodyTypeIndex()
        {
            for (int i = 0; i < _bodyType.Length; i++)
                if (_bodyType[i].Checked) return i;
            return 0;
        }

        /// <summary>单选 none 时隐藏键值网格，露出说明文字。</summary>
        private void UpdateBodyEditor()
        {
            // 单一参数网格按方法分流：GET / DELETE 没有请求体，单选行整行收起，网格内容拼到 URL
            bool isBodyMethod = _method.SelectedItem is "POST" or "PUT";
            _bodyTypeRow.Visible = isBodyMethod;
            _paramsPane.RowStyles[1].Height = isBodyMethod ? 30 : 0;
            bool none = isBodyMethod && BodyTypeIndex() == 0;
            _noneHint.Visible = none;
            foreach (Control control in _noneHint.Parent!.Controls)
                if (control is ThemeBordered) control.Visible = !none;
        }

        private static ThemeButton MakeContentButton(string text, EventHandler onClick)
        {
            var button = new ThemeButton
            {
                Text = text,
                Font = Theme.Body,
                Size = new Size(ThemeButton.FitWidth(text), ThemeMetrics.RowH),
                Margin = new Padding(0, 0, ThemeMetrics.Gap, ThemeMetrics.Gap)
            };
            button.Click += onClick;
            return button;
        }

        private static Label MakeTopLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = Theme.Caption,
                ForeColor = Theme.P.Muted,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                MinimumSize = new Size(0, ThemeMetrics.RowH),
                Margin = new Padding(ThemeMetrics.Gap, 0, 2, ThemeMetrics.Gap)
            };
        }

        private static ThemeComboBox MakeTopCombo(object[] items, int width)
        {
            var combo = new ThemeComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = Theme.Body,
                BackColor = Theme.P.Surface,
                ForeColor = Theme.P.Text,
                Size = new Size(width, ThemeMetrics.RowH),
                Margin = new Padding(0, 0, ThemeMetrics.Gap, ThemeMetrics.Gap)
            };
            combo.Items.AddRange(items);
            return combo;
        }

        private static void ApplyComboTheme(ThemeComboBox combo)
        {
            combo.Font = Theme.Body;
            combo.BackColor = Theme.P.Surface;
            combo.ForeColor = Theme.P.Text;
        }

        private static DataGridView NewPairsGrid(bool suggestHeaderKeys = false)
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                RowHeadersVisible = true,
                RowHeadersWidth = 24,
                BorderStyle = BorderStyle.None,
                AllowUserToDeleteRows = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Margin = new Padding(ThemeMetrics.Margin, 2, ThemeMetrics.Margin, ThemeMetrics.Gap)
            };
            // AllowUserToAddRows 默认 true：ColumnCount 一执行新行就按当时的模板建好了，
            // 模板必须趁新行还不存在时换，否则空新行永远是 TextBox 没有下拉候选
            grid.AllowUserToAddRows = false;
            grid.ColumnCount = 2;
            grid.Columns[0].HeaderText = "键";
            grid.Columns[0].FillWeight = 35;
            grid.Columns[1].HeaderText = "值";
            grid.Columns[1].FillWeight = 65;
            if (suggestHeaderKeys)
            {
                // 键 / 值列换成可下拉的单元格模板；必须在建行之前换，新行单元格是按模板克隆的
                grid.Columns[0].CellTemplate = new HeaderComboCell();
                grid.Columns[1].CellTemplate = new HeaderComboCell();
            }
            grid.AllowUserToAddRows = true;

            // 点行头选中整行后按 Del 删整行；只选单元格时 Del 还是清空值
            grid.KeyDown += (_, e) =>
            {
                if (grid.ReadOnly) return;
                if (e.KeyCode != Keys.Delete || grid.SelectedRows.Count == 0) return;
                if (grid.SelectedRows[0].IsNewRow) return;
                grid.Rows.RemoveAt(grid.SelectedRows[0].Index);
                e.Handled = true;
            };

            ApplyGridTheme(grid);
            return grid;
        }

        /// <summary>详情页签里的展示网格：只读，不接受输入、不响应 Del。</summary>
        private static DataGridView NewReadOnlyPairsGrid()
        {
            var grid = NewPairsGrid();
            grid.AllowUserToAddRows = false;
            grid.ReadOnly = true;
            return grid;
        }

        /// <summary>
        /// 键值网格的可下拉编辑控件：DropDown 样式既能点开选常用值，也能自由输入，
        /// 输入时按前缀筛列表（Suggest）。挂在 DataGridView 里要自己实现编辑控件接口。
        /// </summary>
        private sealed class GridComboEditingControl : ComboBox, IDataGridViewEditingControl
        {
            private DataGridView? _grid;
            private bool _valueChanged;
            private bool _loading;

            public GridComboEditingControl()
            {
                DropDownStyle = ComboBoxStyle.DropDown;
                AutoCompleteMode = AutoCompleteMode.Suggest;
                AutoCompleteSource = AutoCompleteSource.ListItems;
                TabStop = false;
            }

            public void LoadValue(string? value)
            {
                _loading = true;
                try { Text = value ?? ""; }
                finally { _loading = false; }
                _autoDropped = false;
            }

            private bool _autoDropped;

            protected override void OnGotFocus(EventArgs e)
            {
                base.OnGotFocus(e);
                if (_autoDropped) return;
                _autoDropped = true;
                // 新行起编辑时网格会先插行再定位编辑控件，BeginInvoke 一次会被后续定位顶掉；
                // 用一次性定时器等消息队列排空后再展开，Focused 守卫避免用户已经离开时还弹
                var timer = new System.Windows.Forms.Timer { Interval = 200 };
                timer.Tick += (_, _) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    if (Focused && !DroppedDown && Items.Count > 0) DroppedDown = true;
                };
                timer.Start();
            }

            private void MarkDirty()
            {
                if (_loading) return;
                _valueChanged = true;
                _grid?.NotifyCurrentCellDirty(true);
            }

            protected override void OnTextChanged(EventArgs e) { base.OnTextChanged(e); MarkDirty(); }
            protected override void OnSelectedIndexChanged(EventArgs e) { base.OnSelectedIndexChanged(e); MarkDirty(); }

            protected override void OnValidating(CancelEventArgs e)
            {
                // .NET 10 的 ComboBox 在下拉/联想还开着时关窗，OnValidating 里的 NotifyAutoComplete
                // 会拿负长度抛 ArgumentOutOfRange（WinForms 内部 bug）；吞掉这一个异常，不影响取值。
                try { base.OnValidating(e); }
                catch (ArgumentOutOfRangeException) { }
            }

            // 接口属性走显式实现，避免 WinForms 分析器要求给控件属性配序列化
            DataGridView? IDataGridViewEditingControl.EditingControlDataGridView { get => _grid; set => _grid = value; }
            object? IDataGridViewEditingControl.EditingControlFormattedValue { get => Text; set => Text = value as string ?? ""; }
            int IDataGridViewEditingControl.EditingControlRowIndex { get; set; }
            bool IDataGridViewEditingControl.EditingControlValueChanged { get => _valueChanged; set => _valueChanged = value; }
            public Cursor EditingPanelCursor => Cursors.Default;
            public bool RepositionEditingControlOnValueChange => false;

            public void ApplyCellStyleToEditingControl(DataGridViewCellStyle style)
            {
                Font = style.Font;
                BackColor = style.BackColor;
                ForeColor = style.ForeColor;
            }

            public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context) => Text;

            public bool EditingControlWantsInputKey(Keys keyData, bool dataGridViewWantsInputKey)
            {
                // 下拉展开时方向键 / 回车 / Esc 归编辑控件用，其余听网格的
                switch (keyData & Keys.KeyCode)
                {
                    case Keys.Down:
                    case Keys.Up:
                    case Keys.Enter:
                    case Keys.Escape:
                        if (DroppedDown) return true;
                        break;
                }
                return !dataGridViewWantsInputKey;
            }

            public void PrepareEditingControlForEdit(bool selectAll)
            {
                if (selectAll) SelectAll();
            }

            public bool RepaintEditingControl() => false;
        }

        /// <summary>键值网格的单元格：编辑时换成可下拉的 ComboBox；键列候选是常用请求头，值列候选看当前行的键。</summary>
        private sealed class HeaderComboCell : DataGridViewTextBoxCell
        {
            public override Type EditType => typeof(GridComboEditingControl);

            public override void InitializeEditingControl(int rowIndex, object? initialFormattedValue, DataGridViewCellStyle style)
            {
                base.InitializeEditingControl(rowIndex, initialFormattedValue, style);
                if (DataGridView?.EditingControl is not GridComboEditingControl combo) return;
                combo.Items.Clear();
                bool isKeyColumn = OwningColumn != null && DataGridView.Columns.IndexOf(OwningColumn) == 0;
                if (isKeyColumn)
                    combo.Items.AddRange(CommonHeaderKeys);
                else if (ValueSuggestions.TryGetValue(
                        Convert.ToString(DataGridView.Rows[rowIndex].Cells[0].Value)?.Trim() ?? "",
                        out string[]? suggestions))
                    combo.Items.AddRange(suggestions);
                combo.LoadValue(Convert.ToString(initialFormattedValue));
            }
        }

        private static void ApplyGridTheme(DataGridView grid)
        {
            ThemeGrid.Apply(grid);
            grid.RowHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Theme.P.Header,
                SelectionBackColor = Theme.P.Selection,
                Font = Theme.Caption,
                Alignment = DataGridViewContentAlignment.MiddleCenter
            };
        }

        private static Label MakeHint(string text)
        {
            return new Label
            {
                Text = text,
                Font = Theme.Caption,
                ForeColor = Theme.P.Muted,
                BackColor = Theme.P.Window,
                AutoSize = true,
                Margin = new Padding(ThemeMetrics.Margin, 4, ThemeMetrics.Margin, 0)
            };
        }

        private static List<KeyValuePair<string, string>> ReadPairs(DataGridView grid)
        {
            var pairs = new List<KeyValuePair<string, string>>();
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow) continue;
                string key = Convert.ToString(row.Cells[0].Value)?.Trim() ?? "";
                string value = Convert.ToString(row.Cells[1].Value) ?? "";
                if (key.Length > 0) pairs.Add(new KeyValuePair<string, string>(key, value));
            }
            return pairs;
        }

        private async void SendRequest()
        {
            string url = _url.Edit.Text.Trim();
            if (url.Length == 0)
            {
                SetStatus("请先填 URL", error: true);
                return;
            }
            if (_sending)
            {
                SetStatus("上一个请求还在进行…", error: true);
                return;
            }

            string method = (string)(_method.SelectedItem ?? "GET");
            var headerPairs = ReadPairs(_headers);

            bool isBodyMethod = method is "POST" or "PUT";
            var pairs = ReadPairs(_params);
            // Body 类型由单选决定（none = 无请求体）；请求头里手写的 Content-Type 行优先作为实际发送值，
            // 单选 none 但手写了 Content-Type 时按原样 k=v&k=v 组体（游戏服手写解析的兜底通道）
            string radioCt = BodyTypeValues[BodyTypeIndex()];
            string manualCt = headerPairs.FirstOrDefault(kv =>
                kv.Key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase)).Value?.Trim() ?? "";
            bool wantBody = isBodyMethod && pairs.Count > 0 && (radioCt.Length > 0 || manualCt.Length > 0);
            string assembleCt = radioCt.Length > 0 ? radioCt : manualCt;

            Encoding encoding = HttpToolEngine.GetEncoding(_encoding.Text);

            string? bodyText = null;
            bool multipart = false;
            string boundary = "";
            string? bodyPreview = null;
            string? sentContentType = null;
            if (wantBody)
            {
                if (assembleCt.Contains("multipart", StringComparison.OrdinalIgnoreCase))
                {
                    // multipart 用真实边界组装；预览和实际发送共用同一边界，请求页签展示的才是真发出去的
                    multipart = true;
                    boundary = "----wxd" + Guid.NewGuid().ToString("N");
                    using var sample = new MultipartFormDataContent(boundary);
                    foreach (var kv in pairs) sample.Add(new StringContent(kv.Value, encoding), kv.Key);
                    bodyPreview = await sample.ReadAsStringAsync();
                    sentContentType = sample.Headers.ContentType!.ToString();
                }
                else
                {
                    bodyText = assembleCt.Contains("urlencoded", StringComparison.OrdinalIgnoreCase)
                        ? HttpToolEngine.BuildFormBody(pairs, encoding)
                        : assembleCt.Contains("json", StringComparison.OrdinalIgnoreCase)
                            ? HttpToolEngine.BuildJsonBody(pairs)
                            : HttpToolEngine.BuildRawBody(pairs);
                    bodyPreview = bodyText;
                    sentContentType = manualCt.Length > 0 ? manualCt
                        : radioCt.Length > 0 ? radioCt
                        : null;
                }
            }
            else if (!isBodyMethod && pairs.Count > 0)
            {
                url = HttpToolEngine.BuildUrl(url, pairs);
            }

            // 每个请求都要独立的 HttpRequestMessage，压测时按次数重建
            Func<HttpRequestMessage> makeRequest = () =>
            {
                var request = new HttpRequestMessage(new HttpMethod(method), url);
                if (multipart)
                {
                    var content = new MultipartFormDataContent(boundary);
                    foreach (var kv in pairs) content.Add(new StringContent(kv.Value, encoding), kv.Key);
                    request.Content = content;
                }
                else if (bodyText != null)
                {
                    var content = new ByteArrayContent(encoding.GetBytes(bodyText));
                    if (sentContentType != null)
                        content.Headers.TryAddWithoutValidation("Content-Type", sentContentType);
                    request.Content = content;
                }
                foreach (var kv in headerPairs)
                {
                    // Content-Type 是内容头，已经作用在请求体上；无请求体时它也不是合法的请求头
                    if (kv.Key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase)) continue;
                    request.Headers.TryAddWithoutValidation(kv.Key, kv.Value);
                }
                return request;
            };

            int count = ParseBounded(_count.Edit.Text, fallback: 1, min: 1, max: 500);
            int threads = Math.Clamp(ParseBounded(_threads.Edit.Text, fallback: 1, min: 1, max: 64), 1, count);

            _sending = true;
            _sendButton.Enabled = false;
            SetStatus(count > 1 ? $"压测中… {count} 次 / {threads} 线程" : "请求中…");
            try
            {
                if (count == 1)
                {
                    var watch = Stopwatch.StartNew();
                    using var response = await Client.SendAsync(makeRequest(), HttpCompletionOption.ResponseHeadersRead);
                    byte[] data = await response.Content.ReadAsByteArrayAsync();
                    watch.Stop();

                    string responseCharset = HttpToolEngine.ResolveCharset(response.Content.Headers.ContentType?.ToString(), encoding);
                    FillRequestView(method, url, headerPairs, sentContentType, bodyPreview);
                    FillResponseView(response, watch.ElapsedMilliseconds, data.Length, responseCharset);

                    string text = HttpToolEngine.DecodeBody(data, response.Content.Headers.ContentType?.ToString(), encoding)
                        ?? "(二进制响应，未显示)";
                    _respBody.Text = text;
                    _respBody.SelectionStart = 0;
                    _respBody.SelectionLength = 0;
                    SetStatus($"HTTP {(int)response.StatusCode} {response.ReasonPhrase} · {watch.ElapsedMilliseconds} ms · {data.Length:N0} 字节",
                        error: (int)response.StatusCode >= 400);
                }
                else
                {
                    var watch = Stopwatch.StartNew();
                    var latencies = new long[count];
                    int ok = 0, fail = 0;
                    string? firstError = null;
                    string[][]? lastRespHeads = null;
                    int lastStatus = 0, lastBytes = 0;
                    long lastMs = 0;
                    string lastCharset = "";

                    // 线程 t 认领 i = t, t+threads, t+2*threads…；latencies 按各自下标写，无竞争
                    var workers = Enumerable.Range(0, threads).Select(t => Task.Run(async () =>
                    {
                        for (int i = t; i < count; i += threads)
                        {
                            var sw = Stopwatch.StartNew();
                            try
                            {
                                using var request = makeRequest();
                                using var response = await Client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
                                byte[] data = await response.Content.ReadAsByteArrayAsync();
                                sw.Stop();
                                Interlocked.Increment(ref ok);
                                lastRespHeads = response.Headers.Select(h => new[] { h.Key, string.Join("; ", h.Value) })
                                    .Concat(response.Content.Headers.Select(h => new[] { h.Key, string.Join("; ", h.Value) }))
                                    .ToArray();
                                lastStatus = (int)response.StatusCode;
                                lastBytes = data.Length;
                                lastMs = sw.ElapsedMilliseconds;
                                lastCharset = HttpToolEngine.ResolveCharset(response.Content.Headers.ContentType?.ToString(), encoding);
                            }
                            catch (Exception ex)
                            {
                                Interlocked.Increment(ref fail);
                                firstError ??= ex.Message;
                            }
                            sw.Stop();
                            latencies[i] = sw.ElapsedMilliseconds;
                        }
                    })).ToArray();
                    await Task.WhenAll(workers);
                    watch.Stop();

                    FillRequestView(method, url, headerPairs, sentContentType, bodyPreview);
                    _respHeaders.Rows.Clear();
                    if (lastRespHeads != null)
                    {
                        _respMeta.Text = $"最后一次响应：HTTP {lastStatus} · {lastMs} ms · {lastBytes:N0} 字节 · {lastCharset}";
                        foreach (var pair in lastRespHeads) _respHeaders.Rows.Add(pair[0], pair[1]);
                    }
                    else
                    {
                        _respMeta.Text = "全部请求失败，没有响应头可显示";
                    }

                    double seconds = watch.Elapsed.TotalSeconds;
                    double qps = seconds > 0 ? count / seconds : 0;
                    string summary =
                        $"总请求 {count} · 线程 {threads} · 成功 {ok} / 失败 {fail}\r\n" +
                        $"总耗时 {seconds:F2} 秒 · QPS ≈ {qps:F1}\r\n" +
                        $"平均 {latencies.Average():N0} ms · 最快 {latencies.Min():N0} ms · 最慢 {latencies.Max():N0} ms";
                    if (firstError != null) summary += "\r\n首个错误：" + firstError;
                    _respBody.Text = summary;
                    _respBody.SelectionStart = 0;
                    _respBody.SelectionLength = 0;
                    SetStatus($"{count} 次 / {threads} 线程 · 成功 {ok} 失败 {fail} · QPS ≈ {qps:F1}",
                        error: ok == 0);
                }
            }
            catch (TaskCanceledException)
            {
                SetStatus("请求超时（30 秒）", error: true);
            }
            catch (Exception ex)
            {
                SetStatus("请求失败：" + ex.Message, error: true);
            }
            finally
            {
                _sending = false;
                _sendButton.Enabled = true;
            }
        }

        private static int ParseBounded(string text, int fallback, int min, int max)
        {
            if (!int.TryParse(text.Trim(), out int value)) value = fallback;
            return Math.Clamp(value, min, max);
        }

        private static string SavedFilePath
        {
            get
            {
                try { return Path.Combine(AppContext.BaseDirectory, ".wxd-work", "http-requests.json"); }
                catch { return Path.Combine(Path.GetTempPath(), "wxd-http-requests.json"); }
            }
        }

        private void LoadSavedFile()
        {
            try
            {
                string file = SavedFilePath;
                if (File.Exists(file))
                    _saved.AddRange(HttpToolEngine.DeserializeRequests(File.ReadAllText(file)));
            }
            catch
            {
                // 文件存在但读不出来（损坏/占用）：先留一份备份，避免后续写入把它彻底覆盖
                try
                {
                    string file = SavedFilePath;
                    if (File.Exists(file)) File.Copy(file, file + ".bad", overwrite: true);
                }
                catch { }
            }
            RefreshSavedList();
        }

        private void RefreshSavedList()
        {
            _savedUpdating = true;
            try
            {
                _savedList.Items.Clear();
                foreach (var request in _saved) _savedList.Items.Add(request.Name);
                _savedList.SelectedIndex = -1;
                _saveName.Text = "";
            }
            finally { _savedUpdating = false; }
        }

        private void LoadSavedSelection()
        {
            if (_savedUpdating) return;
            int index = _savedList.SelectedIndex;
            if (index < 0 || index >= _saved.Count) return;
            ApplySaved(_saved[index]);
            SetStatus($"已加载 {_saved[index].Name}");
        }

        private void ApplySaved(SavedHttpRequest saved)
        {
            string[] methods = { "GET", "POST", "PUT", "DELETE" };
            int methodIndex = Array.IndexOf(methods, saved.Method);
            _method.SelectedIndex = methodIndex >= 0 ? methodIndex : 0;
            bool bodyMethod = saved.Method is "POST" or "PUT";
            _url.Edit.Text = saved.Url;
            // 单一网格按方法取行：POST / PUT 的键值在 BodyParams（旧档没有此字段时 Params 就是请求体），其余在 Params
            var rows = bodyMethod ? (saved.BodyParams ?? saved.Params) : saved.Params;
            _params.Rows.Clear();
            foreach (var pair in rows) _params.Rows.Add(pair);
            _headers.Rows.Clear();
            foreach (var pair in saved.Headers) _headers.Rows.Add(pair);
            // Content-Type 映射到 Body 单选；认不出的自定义值放进请求头网格（none + 手写 CT 组 raw 体）。
            // null 才是旧档（没有该字段，按 BodyFormat 兜底）；新档 none 存的是空串，直接当 none
            string contentType;
            if (saved.ContentType == null)
                contentType = bodyMethod
                    ? (saved.BodyFormat == "Form" ? "application/x-www-form-urlencoded" : "application/json")
                    : "";
            else
                contentType = saved.ContentType;
            int radioIndex =
                contentType.Contains("multipart", StringComparison.OrdinalIgnoreCase) ? 3 :
                contentType.Contains("urlencoded", StringComparison.OrdinalIgnoreCase) ? 2 :
                contentType.Contains("json", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
            if (radioIndex == 0 && contentType.Length > 0)
                _headers.Rows.Add("Content-Type", contentType);
            _bodyType[radioIndex].Checked = true;
            _count.Edit.Text = saved.Count.ToString();
            _threads.Edit.Text = saved.Threads.ToString();
        }

        private void SaveCurrent()
        {
            string name = _saveName.Text.Trim();
            if (name.Length == 0)
            {
                SetStatus("先在左侧名称框里输入一个名称再保存", error: true);
                return;
            }

            // 实际发送的 Content-Type：请求头手写行优先，否则取 Body 单选的值（none 存 null）
            string manualCt = ReadPairs(_headers).FirstOrDefault(kv =>
                kv.Key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase)).Value?.Trim() ?? "";
            string contentType = manualCt.Length > 0 ? manualCt : BodyTypeValues[BodyTypeIndex()];
            bool bodyMethod = (string)(_method.SelectedItem ?? "GET") is "POST" or "PUT";
            var rows = ReadPairs(_params).Select(kv => new[] { kv.Key, kv.Value }).ToArray();
            var entry = new SavedHttpRequest
            {
                Name = name,
                Method = (string)(_method.SelectedItem ?? "GET"),
                Url = _url.Edit.Text.Trim(),
                ContentType = contentType.Length > 0 ? contentType : "",
                // BodyFormat 只是给旧版 exe 回滚时兜底的粗略提示
                BodyFormat = contentType.Contains("urlencoded", StringComparison.OrdinalIgnoreCase) ? "Form" : "JSON",
                Params = bodyMethod ? Array.Empty<string[]>() : rows,
                BodyParams = bodyMethod ? rows : Array.Empty<string[]>(),
                Headers = ReadPairs(_headers).Select(kv => new[] { kv.Key, kv.Value }).ToArray(),
                Count = ParseBounded(_count.Edit.Text, fallback: 1, min: 1, max: 500),
                Threads = ParseBounded(_threads.Edit.Text, fallback: 1, min: 1, max: 64),
            };

            int index = _saved.FindIndex(r => r.Name == name);
            bool overwritten = index >= 0;
            if (overwritten) _saved[index] = entry;
            else _saved.Add(entry);

            if (PersistSaved())
            {
                _savedUpdating = true;
                try
                {
                    if (!overwritten) _savedList.Items.Add(name);
                    _savedList.SelectedIndex = overwritten ? index : _savedList.Items.Count - 1;
                }
                finally { _savedUpdating = false; }
                SetStatus(overwritten ? $"已覆盖 {name}" : $"已保存 {name}");
            }
        }

        private void DeleteSaved()
        {
            int index = _savedList.SelectedIndex;
            if (index < 0 || index >= _saved.Count)
            {
                SetStatus("先在左侧列表里选中要删除的请求", error: true);
                return;
            }
            string name = _saved[index].Name;
            _saved.RemoveAt(index);
            if (PersistSaved())
            {
                RefreshSavedList();
                SetStatus($"已删除 {name}");
            }
        }

        private bool PersistSaved()
        {
            try
            {
                string file = SavedFilePath;
                Directory.CreateDirectory(Path.GetDirectoryName(file)!);
                string temp = file + ".tmp";
                File.WriteAllText(temp, HttpToolEngine.SerializeRequests(_saved));
                File.Move(temp, file, overwrite: true);
                return true;
            }
            catch (Exception ex)
            {
                SetStatus("保存失败：" + ex.Message, error: true);
                return false;
            }
        }

        private void FillRequestView(string method, string url,
            List<KeyValuePair<string, string>> headers, string? contentType, string? body)
        {
            _reqLine.Text = method + " " + url;
            _reqHeaders.Rows.Clear();
            foreach (var kv in headers)
            {
                if (kv.Key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase)) continue;
                _reqHeaders.Rows.Add(kv.Key, kv.Value);
            }
            if (body != null && !string.IsNullOrEmpty(contentType))
                _reqHeaders.Rows.Add("Content-Type", contentType);
            _reqBody.Text = body ?? "";
        }

        private void FillResponseView(HttpResponseMessage response, long ms, int bytes, string charset)
        {
            _respMeta.Text = $"HTTP {(int)response.StatusCode} {response.ReasonPhrase} · {ms} ms · {bytes:N0} 字节 · {charset}";
            _respHeaders.Rows.Clear();
            foreach (var header in response.Headers)
                _respHeaders.Rows.Add(header.Key, string.Join("; ", header.Value));
            foreach (var header in response.Content.Headers)
                _respHeaders.Rows.Add(header.Key, string.Join("; ", header.Value));
        }

        private void PrettyPrintResponse()
        {
            string text = _respBody.Text;
            if (text.Length == 0)
            {
                SetStatus("还没有响应内容", error: true);
                return;
            }
            try
            {
                var node = JsonNode.Parse(text);
                _respBody.Text = node!.ToJsonString(new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });
                _respBody.SelectionStart = 0;
                _respBody.SelectionLength = 0;
                SetStatus("响应 JSON 已格式化");
            }
            catch
            {
                SetStatus("响应不是 JSON，保持原样", error: true);
            }
        }
    }
}
