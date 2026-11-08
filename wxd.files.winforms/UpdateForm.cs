using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>更新管理窗体：使用说明 + 检查更新 + 进度条</summary>
    public class UpdateForm : Form
    {
        private readonly RichTextBox _rtbContent = new();
        private readonly ThemeButton _btnCheckUpdate = new();
        private readonly ThemeButton _btnChangelog = new();
        private readonly ThemeButton _btnForceReinstall = new();
        private readonly ThemeButton _btnClose = new();
        private readonly Label _lblStatus = new();
        private readonly ThemeProgressBar _progressBar = new();
        private readonly Label _lblProgressText = new();
        private readonly Label _lblLastCheck = new();
        private readonly Label _lblTitle = new();
        private readonly Panel _topBar = new();
        private readonly Panel _topDivider = new();
        private readonly Panel _statusRow = new();
        private readonly Panel _bar = new();

        private CheckResult? _lastCheckResult;

        private const string HelpContent = @"📂 多窗格文件管理器 - 使用说明
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

◆ 工具栏操作

  · 【≡ 设置按钮】—— 点击展开/收起设置面板
    ├ 窗格数量：1~10 个窗格自由设置
    ├ 布局方式：横向、纵向、纵横交错流式三种模式
    ├ 主题切换：深色模式 / 浅色模式
    ├ 保存状态：手动保存当前所有窗格的目录和过滤条件
    └ 命令管理：管理自定义命令按钮

  · 【显示隐藏】—— 切换显示/隐藏系统隐藏文件和文件夹
    （勾选后立即生效，所有面板同步刷新）

  · 【阅读器】—— 打开 CSV/Excel 阅读器独立窗口
    （需先选择包含 CSV 或 Excel 文件的目录）

  · 【日志分析】—— 打开大文件阅读器独立窗口
    （支持 GB 级超大文本文件，虚拟滚动高性能浏览）

  · 【自定义命令】—— 通过「命令管理」添加的自定义按钮
    （点击后执行配置的 cmd 命令，支持设置工作目录和显示控制台窗口）


◆ 更新检查

  · 程序启动时自动检查是否有新版本
  · 点击「关于」窗口中的「检查更新」手动检测
  · 发现新版本时显示更新日志，可选择下载或跳过
  · 支持跳过指定版本，下次不重复提示
  · 可通过「强制重装」重新下载当前版本


◆ 窗格文件管理

  · 路径框
    ├ 下拉选择磁盘盘符或最近访问的路径
    ├ 直接输入路径后按回车跳转
    └ 支持 UNC 网络路径

  · 上级按钮 —— 返回当前目录的上级目录

  · 文件列表
    ├ 双击文件夹进入该目录
    ├ 双击文件用系统默认程序打开
    └ 右键弹出 Windows 原生右键菜单（支持多选文件）

  · 视图切换 —— 点击面板顶部的视图按钮切换五种视图模式
    ├ 详细信息（默认）
    ├ 列表
    ├ 大图标
    ├ 小图标
    └ 平铺

  · 过滤框
    ├ 输入关键词实时过滤文件列表（300ms 防抖）
    ├ 支持 || 运算符：匹配任意关键词（或）
    ├ 支持 && 运算符：同时匹配所有关键词（与）
    └ 按 Esc 键快速清空过滤


◆ 键盘快捷键

  · F2      —— 重命名当前选中的文件/文件夹
  · F5      —— 刷新当前窗格的文件列表
  · Delete  —— 删除选中的文件/文件夹
  · Ctrl+A  —— 全选当前窗格中的所有文件


◆ 窗格数量与布局

  · 数量范围：1~10 个窗格
  · 横向布局：所有窗格一行等分宽度
  · 纵向布局：所有窗格一列等分高度
  · 纵横交错流式布局：
    ├ 3/5/7 个窗格：左侧流式排列 + 右侧一个占满高度
    └ 其他数量：网格流式排列


◆ 主题切换

  · 深色模式：Visual Studio 风格深色主题
  · 浅色模式：清爽天蓝配色


◆ 目录级配置

  · CSV/Excel 阅读器中每个目录的配置独立存储
  · 配置文件位于目录下的 .wxd-work/config.json
  · 包含：冻结行、表头行、列关联、列显隐、文件备注等
  · 配置随文件夹一起移植，无需重复设置


◆ 状态记忆（自动保存 & 自动恢复）

  · 每个窗格的当前目录、过滤条件、视图模式
  · 窗格数量、布局方式、主题
  · 显示隐藏文件开关状态
  · 自定义命令按钮配置
  · 关闭窗口时自动保存，重启程序自动恢复所有状态";
        public UpdateForm()
        {
            // 当前版本 - 优先从 config 读取（与更新检测对比用同一个版本号）
            var configVersion = UpdateHelper.GetCurrentVersion();
            string verStr;
            if (!string.IsNullOrEmpty(configVersion))
            {
                verStr = $"v{configVersion}";
            }
            else
            {
                var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                verStr = ver != null ? $"v{ver}" : "未知";
            }

            Text = "关于 - wxd.files.winforms";
            ClientSize = new Size(700, 720);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;

            // 页眉两行：标题 + 上次检测时间，比单行页眉高一档
            _topBar.Dock = DockStyle.Top;
            _topBar.Height = ThemeMetrics.BarHeight + ThemeMetrics.RowH;
            _topDivider.Dock = DockStyle.Bottom;
            _topDivider.Height = 1;
            _topBar.Controls.Add(_topDivider);

            _lblTitle.Text = $"wxd.files.winforms  {verStr}";
            _lblTitle.AutoSize = true;
            _lblTitle.BackColor = Color.Transparent;
            _lblLastCheck.AutoSize = true;
            _lblLastCheck.BackColor = Color.Transparent;
            _lblLastCheck.Text = LastCheckText();
            _topBar.Controls.AddRange(new Control[] { _lblTitle, _lblLastCheck });
            _topBar.Resize += (_, _) => LayoutTopBar();

            _rtbContent.ReadOnly = true;
            _rtbContent.BorderStyle = BorderStyle.None;
            _rtbContent.ScrollBars = RichTextBoxScrollBars.Vertical;
            _rtbContent.DetectUrls = false;
            _rtbContent.WordWrap = true;
            _rtbContent.Text = HelpContent;

            var body = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(ThemeMetrics.Margin + 6, ThemeMetrics.Group,
                                     ThemeMetrics.Margin + 6, ThemeMetrics.Gap)
            };
            body.Controls.Add(new ThemeBordered(_rtbContent) { Dock = DockStyle.Fill });

            // 状态行：左边状态文字，右边进度条 + 百分比，进度条平时留着不显示数值
            _statusRow.Dock = DockStyle.Bottom;
            _statusRow.Height = ThemeMetrics.BarHeight;
            _lblStatus.AutoSize = false;
            _lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            _lblStatus.BackColor = Color.Transparent;
            _lblStatus.Text = "就绪";
            _lblProgressText.AutoSize = false;
            _lblProgressText.TextAlign = ContentAlignment.MiddleRight;
            _lblProgressText.BackColor = Color.Transparent;
            _statusRow.Controls.AddRange(new Control[] { _lblStatus, _progressBar, _lblProgressText });
            _statusRow.Resize += (_, _) => LayoutStatusRow();

            _bar = DialogButtonBar.Create();
            _btnForceReinstall.Text = "强制重装";
            _btnForceReinstall.Click += OnForceReinstallClick;
            _btnChangelog.Text = "更新记录";
            _btnChangelog.Click += (_, _) => ChangelogForm.ShowChangelog(this);
            _btnCheckUpdate.Text = "检查更新";
            _btnCheckUpdate.Click += OnCheckUpdateClick;
            _btnClose.Text = "关闭";
            _btnClose.Primary = true;
            _btnClose.DialogResult = DialogResult.OK;
            _btnClose.Click += (s, e) => Close();
            _bar.Controls.AddRange(new Control[]
                { _btnForceReinstall, _btnChangelog, _btnCheckUpdate, _btnClose });
            _bar.Resize += (_, _) => LayoutBar();

            // 停靠顺序：同一条边上后加入的靠外，所以正文先加、页眉页脚后加
            Controls.Add(body);
            Controls.Add(_statusRow);
            Controls.Add(_topBar);
            Controls.Add(_bar);

            AcceptButton = _btnClose;
            CancelButton = _btnClose;

            ApplyTheme();
            Theme.Changed += ApplyTheme;
            Disposed += (_, _) => Theme.Changed -= ApplyTheme;
        }

        private static string LastCheckText()
        {
            var lastCheckDt = UpdateHelper.GetLastCheckTimePublic();
            if (lastCheckDt == DateTime.MinValue)
                return "尚未检测过更新（每天自动检测一次）";

            var checkedToday = lastCheckDt.Date == DateTime.Now.Date;
            string remainText = checkedToday ? "（今日已检测）" : "（可手动检测更新）";
            return $"上次检测: {lastCheckDt:yyyy-MM-dd HH:mm}  {remainText}";
        }

        private void LayoutTopBar()
        {
            int m = ThemeMetrics.Margin + 6;
            _lblTitle.Location = new Point(m, ThemeMetrics.RowPad);
            _lblLastCheck.Location = new Point(m + 1, _lblTitle.Bottom + 2);
        }

        private void LayoutStatusRow()
        {
            int m = ThemeMetrics.Margin + 6;
            int rowH = ThemeMetrics.RowH;
            int y = (_statusRow.Height - rowH) / 2;
            int barW = 200;
            const int textW = 44;

            _progressBar.SetBounds(_statusRow.Width - m - textW - ThemeMetrics.Gap - barW,
                (_statusRow.Height - (rowH - 8)) / 2, barW, rowH - 8);
            _lblProgressText.SetBounds(_progressBar.Right + ThemeMetrics.Gap, y, textW, rowH);
            _lblStatus.SetBounds(m, y, Math.Max(60, _progressBar.Left - m - ThemeMetrics.Group), rowH);
        }

        private void LayoutBar()
            => DialogButtonBar.LayoutRight(_bar, _btnForceReinstall, _btnChangelog,
                                           _btnCheckUpdate, _btnClose);

        private void ApplyTheme()
        {
            var p = Theme.P;
            Font = Theme.Body;
            BackColor = p.Window;
            _topBar.BackColor = p.Chrome;
            _statusRow.BackColor = p.Chrome;
            _bar.BackColor = p.Chrome;
            _topDivider.BackColor = p.Divider;

            _lblTitle.ForeColor = p.Text;
            _lblTitle.Font = Theme.Head;
            _lblLastCheck.ForeColor = p.Muted;
            _lblLastCheck.Font = Theme.Caption;
            _lblStatus.ForeColor = p.Muted;
            _lblProgressText.ForeColor = p.Muted;

            _rtbContent.BackColor = p.Surface;
            _rtbContent.ForeColor = p.Text;
            _rtbContent.Font = Theme.Body;

            foreach (var b in new[] { _btnForceReinstall, _btnChangelog, _btnCheckUpdate, _btnClose })
                b.BackColor = p.Chrome;

            LayoutTopBar();
            LayoutStatusRow();
            LayoutBar();
            Invalidate(true);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            _rtbContent.SelectionStart = 0;
            _rtbContent.ScrollToCaret();
            _btnClose.Focus();
        }

        private async void OnCheckUpdateClick(object? sender, EventArgs e)
        {
            _btnCheckUpdate.Enabled = false;
            _lblStatus.Text = "正在检查更新...";
            _progressBar.Marquee = true;

            try
            {
                _lastCheckResult = await UpdateHelper.CheckManualAsync();
                _lblLastCheck.Text = LastCheckText();

                if (_lastCheckResult.HasUpdate)
                {
                    _lblStatus.Text = $"发现新版本 v{_lastCheckResult.RemoteVersion}！";

                    // 统一使用 UpdateNotifyForm 弹窗（点击"立即更新"后内部切换为下载进度界面）
                    var notifyForm = new UpdateNotifyForm(_lastCheckResult);
                    notifyForm.ShowDialog(this);

                    if (notifyForm.UserChoice == 2)
                    {
                        UpdateHelper.AddSkippedVersion(_lastCheckResult.RemoteVersion);
                        _lblStatus.Text = $"已跳过版本 v{_lastCheckResult.RemoteVersion}";
                    }
                    else if (notifyForm.UserChoice == 3)
                    {
                        _lblStatus.Text = $"已忽略版本 v{_lastCheckResult.RemoteVersion}，下次启动时再次提示";
                    }
                    // UserChoice == 1 → UpdateNotifyForm 内部已处理下载和退出
                }
                else
                {
                    _lblStatus.Text = "已是最新版本！";
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "检查更新失败";
                AlertForm.Show(null, $"检查更新失败：{ex.Message}", "网络错误", AlertType.Warning, exception: ex);
            }
            finally
            {
                _progressBar.Marquee = false;
                _progressBar.Value = 0;
                _btnCheckUpdate.Enabled = true;
            }
        }

        private void OnForceReinstallClick(object? sender, EventArgs e)
        {
            if (_lastCheckResult == null)
            {
                OnCheckUpdateClick(sender, e);
                return;
            }

            var result = AlertForm.ShowConfirm(null,
                $"当前版本 v{_lastCheckResult.CurrentVersion}，确定要强制重装吗？" + GlobalConst.NewLine +
                "将重新下载并覆盖当前程序文件。",
                "强制重装", AlertType.Information);

            if (result != DialogResult.Yes) return;

            // 用 UpdateNotifyForm 展示强制重装的下载进度
            var notifyForm = new UpdateNotifyForm(_lastCheckResult);
            notifyForm.ShowDialog(this);
            // UserChoice == 1 → UpdateNotifyForm 内部已处理下载和退出
        }

        /// <summary>便捷方法：显示更新窗体</summary>
        public static void ShowUpdateForm(IWin32Window owner)
        {
            using var form = new UpdateForm();
            form.ShowDialog(owner);
        }
    }
}
