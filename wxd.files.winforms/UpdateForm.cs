using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>更新管理窗体：使用说明 + 检查更新 + 进度条</summary>
    public class UpdateForm : Form
    {
        private readonly RichTextBox _rtbContent;
        private readonly Button _btnCheckUpdate;
        private readonly Button _btnForceReinstall;
        private readonly Button _btnClose;
        private readonly Label _lblStatus;
        private readonly ProgressBar _progressBar;
        private readonly Label _lblProgressText;
        private readonly Label _lblLastCheck;
        private readonly bool _isDark;

        private CheckResult? _lastCheckResult;

        // 浅色配色
        private static readonly Color _lightBg = Color.FromArgb(235, 248, 255);
        private static readonly Color _lightTitleFg = Color.FromArgb(20, 60, 100);
        private static readonly Color _lightBtnBg = Color.FromArgb(66, 165, 245);
        private static readonly Color _lightBtnBorder = Color.FromArgb(30, 136, 229);
        private static readonly Color _lightStatusFg = Color.FromArgb(90, 130, 170);
        private static readonly Color _lightSubFg = Color.FromArgb(140, 160, 180);

        // 深色配色
        private static readonly Color _darkBg = Color.FromArgb(30, 30, 30);
        private static readonly Color _darkTitleFg = Color.FromArgb(241, 241, 241);
        private static readonly Color _darkBtnBg = Color.FromArgb(0, 122, 204);
        private static readonly Color _darkBtnBorder = Color.FromArgb(0, 95, 160);
        private static readonly Color _darkStatusFg = Color.FromArgb(200, 200, 200);
        private static readonly Color _darkSubFg = Color.FromArgb(130, 130, 130);

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

        public UpdateForm(bool isDark = false)
        {
            _isDark = isDark;
            Text = "关于 - wxd.files.winforms";
            ClientSize = new Size(700, 720);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            Font = new Font("Microsoft YaHei UI", 9f);
            BackColor = isDark ? _darkBg : _lightBg;

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

            // 顶栏背景
            var topPanel = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(700, 52),
                BackColor = isDark ? Color.FromArgb(40, 40, 43) : Color.FromArgb(200, 228, 245)
            };

            // 标题（版本号）
            var lblTitle = new Label
            {
                Text = $"wxd.files.winforms  {verStr}",
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 13f, FontStyle.Bold),
                ForeColor = isDark ? _darkTitleFg : _lightTitleFg,
                Location = new Point(16, 8),
                BackColor = Color.Transparent
            };

            // 上次检测时间 + 检测频率提示
            var lastCheckDt = UpdateHelper.GetLastCheckTimePublic();
            string lastCheckText;
            if (lastCheckDt == DateTime.MinValue)
            {
                lastCheckText = "尚未检测过更新（每天自动检测一次）";
            }
            else
            {
                var checkedToday = lastCheckDt.Date == DateTime.Now.Date;
                string remainText = checkedToday ? "（今日已检测）" : "（可手动检测更新）";
                lastCheckText = $"上次检测: {lastCheckDt:yyyy-MM-dd HH:mm}  {remainText}";
            }
            _lblLastCheck = new Label
            {
                Text = lastCheckText,
                AutoSize = true,
                Font = new Font("Microsoft YaHei UI", 8.5f),
                ForeColor = isDark ? _darkSubFg : _lightSubFg,
                Location = new Point(20, 32),
                BackColor = Color.Transparent
            };

            topPanel.Controls.AddRange(new Control[] { lblTitle, _lblLastCheck });

            // 分隔线
            var sep = new Panel
            {
                Location = new Point(0, 52),
                Size = new Size(700, 1),
                BackColor = isDark ? Color.FromArgb(62, 62, 66) : Color.FromArgb(173, 209, 237)
            };

            // 富文本框（使用说明）
            _rtbContent = new RichTextBox
            {
                Location = new Point(16, 62),
                Size = new Size(668, 570),
                Font = new Font("Microsoft YaHei UI", 10f),
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                DetectUrls = false,
                WordWrap = true,
                BackColor = isDark ? Color.FromArgb(37, 37, 38) : Color.White,
                ForeColor = isDark ? Color.FromArgb(220, 220, 220) : Color.FromArgb(30, 60, 100)
            };
            _rtbContent.Text = HelpContent;

            // 底栏背景
            var bottomPanel = new Panel
            {
                Location = new Point(0, 640),
                Size = new Size(700, 80),
                BackColor = isDark ? Color.FromArgb(37, 37, 38) : Color.FromArgb(220, 238, 250)
            };

            // 状态文字
            _lblStatus = new Label
            {
                Text = "就绪",
                AutoSize = false,
                Size = new Size(300, 22),
                Location = new Point(16, 8),
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = isDark ? _darkStatusFg : _lightStatusFg,
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 进度条
            _progressBar = new ProgressBar
            {
                Location = new Point(16, 34),
                Size = new Size(300, 18),
                Visible = false,
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };

            // 进度百分比文字
            _lblProgressText = new Label
            {
                Text = "",
                AutoSize = false,
                Size = new Size(45, 18),
                Location = new Point(320, 34),
                Font = new Font("Microsoft YaHei UI", 9f),
                ForeColor = isDark ? _darkStatusFg : _lightStatusFg,
                BackColor = Color.Transparent,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 检查更新按钮（右侧区域，不和左侧状态文字重叠）
            _btnCheckUpdate = new Button
            {
                Text = "检查更新",
                Size = new Size(72, 30),
                Location = new Point(396, 8),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = isDark ? _darkBtnBg : _lightBtnBg,
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            _btnCheckUpdate.FlatAppearance.BorderSize = 1;
            _btnCheckUpdate.FlatAppearance.BorderColor = isDark ? _darkBtnBorder : _lightBtnBorder;
            _btnCheckUpdate.Click += OnCheckUpdateClick;

            // 更新记录按钮
            var btnChangelog = new Button
            {
                Text = "更新记录",
                Size = new Size(72, 30),
                Location = new Point(472, 8),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = isDark ? _darkBtnBg : _lightBtnBg,
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnChangelog.FlatAppearance.BorderSize = 1;
            btnChangelog.FlatAppearance.BorderColor = isDark ? _darkBtnBorder : _lightBtnBorder;
            btnChangelog.Click += (s, e) => ChangelogForm.ShowChangelog(this, _isDark);

            // 强制重装按钮
            _btnForceReinstall = new Button
            {
                Text = "强制重装",
                Size = new Size(72, 30),
                Location = new Point(548, 8),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(255, 152, 0),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            _btnForceReinstall.FlatAppearance.BorderSize = 1;
            _btnForceReinstall.FlatAppearance.BorderColor = Color.FromArgb(230, 130, 0);
            _btnForceReinstall.Click += OnForceReinstallClick;

            // 关闭按钮
            _btnClose = new Button
            {
                Text = "关闭",
                Size = new Size(72, 30),
                Location = new Point(624, 8),
                Font = new Font("Microsoft YaHei UI", 10f),
                FlatStyle = FlatStyle.Flat,
                BackColor = isDark ? _darkBtnBg : _lightBtnBg,
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.OK
            };
            _btnClose.FlatAppearance.BorderSize = 1;
            _btnClose.FlatAppearance.BorderColor = isDark ? _darkBtnBorder : _lightBtnBorder;
            _btnClose.Click += (s, e) => Close();

            bottomPanel.Controls.AddRange(new Control[] {
                _lblStatus, _progressBar, _lblProgressText,
                _btnCheckUpdate, btnChangelog, _btnForceReinstall, _btnClose
            });

            Controls.AddRange(new Control[] {
                topPanel, sep, _rtbContent, bottomPanel
            });

            AcceptButton = _btnClose;
            CancelButton = _btnClose;
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
            _progressBar.Visible = true;
            _progressBar.Style = ProgressBarStyle.Marquee;

            try
            {
                _lastCheckResult = await UpdateHelper.CheckManualAsync();

                if (_lastCheckResult.HasUpdate)
                {
                    _lblStatus.Text = $"发现新版本 v{_lastCheckResult.RemoteVersion}！";

                    // 检查是否在跳过列表里
                    var skipped = UpdateHelper.GetSkippedVersions();
                    bool isSkipped = skipped.Contains(_lastCheckResult.RemoteVersion);

                    // 统一使用 UpdateNotifyForm 弹窗（点击"立即更新"后内部切换为下载进度界面）
                    var notifyForm = new UpdateNotifyForm(_lastCheckResult, _isDark);
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
                MessageBox.Show($"检查更新失败：{ex.Message}", "网络错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                _progressBar.Style = ProgressBarStyle.Continuous;
                _progressBar.Value = 0;
                _progressBar.Visible = false;
                _btnCheckUpdate.Enabled = true;
            }
        }

        private async void OnForceReinstallClick(object? sender, EventArgs e)
        {
            if (_lastCheckResult == null)
            {
                OnCheckUpdateClick(sender, e);
                return;
            }

            var result = MessageBox.Show(
                $"当前版本 v{_lastCheckResult.CurrentVersion}，确定要强制重装吗？\n" +
                "将重新下载并覆盖当前程序文件。",
                "强制重装",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes) return;

            // 用 UpdateNotifyForm 展示强制重装的下载进度
            var notifyForm = new UpdateNotifyForm(_lastCheckResult, _isDark);
            // 直接模拟点击"立即更新"触发下载
            notifyForm.ShowDialog(this);
            // UserChoice == 1 → UpdateNotifyForm 内部已处理下载和退出
        }

        /// <summary>便捷方法：显示更新窗体</summary>
        public static void ShowUpdateForm(IWin32Window owner, bool isDark = false)
        {
            using var form = new UpdateForm(isDark);
            form.ShowDialog(owner);
        }
    }
}
