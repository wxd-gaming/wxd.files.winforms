using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>
    /// 带页签的文件浏览面板容器。
    /// 内部为每个页签维护一个独立的 FileBrowserPanel 实例，完全隔离状态。
    /// 页签上限 10 个，每个页签显示关闭按钮（始终保留至少 1 个页签），鼠标悬停显示完整路径 ToolTip。
    /// </summary>
    public class TabbedFileBrowserPanel : UserControl
    {
        // ==================== 常量 ====================
        private const int MaxTabCount = 10;
        private const int TabBarHeight = 28;
        private const int TabMinWidth = 80;
        private const int TabMaxWidth = 200;
        private const int TabPaddingX = 14;
        private const int AddButtonWidth = 26;
        private const int TabCloseButtonSize = 14;
        private const int TabCloseRightMargin = 6;

        // ==================== 主题颜色 ====================
        // 浅色：页签条用比主界面浅蓝(235,248,255)更深的蓝灰，激活页签用更亮的蓝白形成对比
        private static readonly Color _colorLightTabBarBg = Color.FromArgb(218, 230, 243);
        private static readonly Color _colorLightTabBg = Color.FromArgb(218, 230, 243);
        private static readonly Color _colorLightTabActiveBg = Color.FromArgb(240, 247, 252);
        private static readonly Color _colorLightTabHoverBg = Color.FromArgb(40, 255, 255, 255);
        private static readonly Color _colorLightTabBorder = Color.FromArgb(173, 209, 237);
        private static readonly Color _colorLightTabText = Color.FromArgb(15, 15, 15);
        private static readonly Color _colorLightTabTextSecondary = Color.FromArgb(70, 90, 110);
        private static readonly Color _colorLightAddBtn = Color.FromArgb(90, 120, 140);
        private static readonly Color _colorLightAccent = Color.FromArgb(0, 120, 215);

        // 深色：非激活页签融入页签条，激活用稍亮底+底部强调线
        private static readonly Color _colorDarkTabBarBg = Color.FromArgb(37, 37, 38);
        private static readonly Color _colorDarkTabBg = Color.FromArgb(37, 37, 38);
        private static readonly Color _colorDarkTabActiveBg = Color.FromArgb(45, 45, 48);
        private static readonly Color _colorDarkTabHoverBg = Color.FromArgb(40, 255, 255, 255);
        private static readonly Color _colorDarkTabBorder = Color.FromArgb(67, 67, 70);
        private static readonly Color _colorDarkTabText = Color.FromArgb(241, 241, 241);
        private static readonly Color _colorDarkTabTextSecondary = Color.FromArgb(170, 170, 170);
        private static readonly Color _colorDarkAddBtn = Color.FromArgb(150, 150, 150);
        private static readonly Color _colorDarkAccent = Color.FromArgb(0, 151, 251);

        // ==================== 状态 ====================
        private readonly int _paneIndex;
        private readonly List<FileBrowserPanel> _tabPanels = new();
        private int _activeTabIndex = 0;
        private bool _isDark = false;
        private int _hoveredTab = -1;       // 鼠标悬停的页签索引（-1 表示无）
        private int _hoveredCloseTab = -1;  // 鼠标悬停在关闭按钮上的页签索引（-1 表示无）
        private bool _hoverAddButton = false;

        // ==================== 控件 ====================
        private readonly Panel _pnlTabBar;
        private readonly Panel _pnlContent;
        private readonly ToolTip _tabToolTip;

        // ==================== 公共代理属性 ====================
        /// <summary>当前激活页签面板的目录</summary>
        public string CurrentDirectory => GetActivePanel()?.CurrentDirectory ?? "";

        /// <summary>当前激活页签面板的过滤文本</summary>
        public string FilterText => GetActivePanel()?.FilterText ?? "";

        /// <summary>窗格索引</summary>
        public int PaneIndex => _paneIndex;

        /// <summary>当前页签数量</summary>
        public int TabCount => _tabPanels.Count;

        // ==================== 构造 ====================
        public TabbedFileBrowserPanel() : this(0) { }

        public TabbedFileBrowserPanel(int paneIndex)
        {
            _paneIndex = paneIndex;
            DoubleBuffered = true;

            // 页签条
            _pnlTabBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = TabBarHeight,
                BackColor = _colorLightTabBarBg,
            };
            _pnlTabBar.Paint += OnTabBarPaint;
            _pnlTabBar.MouseDown += OnTabBarMouseDown;
            _pnlTabBar.MouseMove += OnTabBarMouseMove;
            _pnlTabBar.MouseLeave += OnTabBarMouseLeave;

            // 内容区
            _pnlContent = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
            };

            Controls.Add(_pnlContent);
            Controls.Add(_pnlTabBar);

            // ToolTip
            _tabToolTip = new ToolTip
            {
                InitialDelay = 300,
                ReshowDelay = 100,
                ShowAlways = true,
            };
        }

        // ==================== 初始化页签（恢复配置） ====================

        /// <summary>
        /// 从 AppStateStore 恢复页签：读取保存的页签数量，为每个页签创建独立 FileBrowserPanel。
        /// 每个页签使用 (paneIndex, tabIndex) 双键从 AppStateStore 读取/写入状态。
        /// </summary>
        public void RestoreTabs(bool skipAutoNavigate)
        {
            int savedTabCount = AppStateStore.GetPaneTabCount(_paneIndex);
            int savedActiveTab = AppStateStore.GetPaneActiveTab(_paneIndex);

            for (int i = 0; i < savedTabCount; i++)
            {
                var panel = new FileBrowserPanel(_paneIndex, i, skipAutoNavigate)
                {
                    Dock = DockStyle.Fill,
                    BorderStyle = BorderStyle.None,
                };
                panel.DirectoryChanged += OnPanelDirectoryChanged;
                panel.StateChanged += OnPanelStateChanged;
                _tabPanels.Add(panel);
            }

            // 确保至少有一个页签
            if (_tabPanels.Count == 0)
            {
                var panel = new FileBrowserPanel(_paneIndex, 0, skipAutoNavigate)
                {
                    Dock = DockStyle.Fill,
                    BorderStyle = BorderStyle.None,
                };
                panel.DirectoryChanged += OnPanelDirectoryChanged;
                panel.StateChanged += OnPanelStateChanged;
                _tabPanels.Add(panel);
                savedActiveTab = 0;
            }

            _activeTabIndex = Math.Min(savedActiveTab, _tabPanels.Count - 1);
            ShowActiveTab();
            _pnlTabBar.Invalidate();
        }

        // ==================== 添加 / 切换页签 ====================

        /// <summary>
        /// 新增一个页签并导航到指定路径。
        /// 如果达到上限则不新增，返回 false。
        /// </summary>
        public bool AddTab(string path)
        {
            if (_tabPanels.Count >= MaxTabCount)
                return false;

            int tabIndex = _tabPanels.Count;

            // 创建新面板（skipAutoNavigate=true，因为我们会手动导航）
            var panel = new FileBrowserPanel(_paneIndex, tabIndex, skipAutoNavigate: true)
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
            };
            panel.DirectoryChanged += OnPanelDirectoryChanged;
            panel.StateChanged += OnPanelStateChanged;

            // 应用当前主题
            panel.ApplyTheme(_isDark);

            // 导航到指定路径
            if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
            {
                panel.NavigateTo(path, restoreFilter: false, saveState: false);
            }

            _tabPanels.Add(panel);
            _activeTabIndex = tabIndex;
            ShowActiveTab();

            // 持久化页签数量
            AppStateStore.SetPaneTabCount(_paneIndex, _tabPanels.Count);
            AppStateStore.SetPaneActiveTab(_paneIndex, _activeTabIndex);

            _pnlTabBar.Invalidate();
            return true;
        }

        /// <summary>
        /// 关闭指定页签。始终保留至少 1 个页签。
        /// 如果关闭的是激活页签，自动切换到相邻页签。
        /// </summary>
        public bool CloseTab(int tabIndex)
        {
            if (tabIndex < 0 || tabIndex >= _tabPanels.Count || _tabPanels.Count <= 1)
                return false;

            var panel = _tabPanels[tabIndex];
            panel.DirectoryChanged -= OnPanelDirectoryChanged;
            panel.StateChanged -= OnPanelStateChanged;
            _tabPanels.RemoveAt(tabIndex);

            if (!panel.IsDisposed)
                panel.Dispose();

            // 调整激活索引
            if (_activeTabIndex == tabIndex)
            {
                // 关闭的是激活页签，切换到相邻页签
                _activeTabIndex = Math.Min(tabIndex, _tabPanels.Count - 1);
            }
            else if (_activeTabIndex > tabIndex)
            {
                // 关闭的页签在激活页签之前，索引前移
                _activeTabIndex--;
            }

            ShowActiveTab();
            AppStateStore.SetPaneTabCount(_paneIndex, _tabPanels.Count);
            AppStateStore.SetPaneActiveTab(_paneIndex, _activeTabIndex);
            _pnlTabBar.Invalidate();
            return true;
        }

        /// <summary>切换到指定页签</summary>
        public void SwitchToTab(int tabIndex)
        {
            if (tabIndex < 0 || tabIndex >= _tabPanels.Count) return;
            if (tabIndex == _activeTabIndex) return;

            _activeTabIndex = tabIndex;
            ShowActiveTab();
            AppStateStore.SetPaneActiveTab(_paneIndex, _activeTabIndex);
            _pnlTabBar.Invalidate();
        }

        /// <summary>显示当前激活的页签，隐藏其余</summary>
        private void ShowActiveTab()
        {
            _pnlContent.Controls.Clear();
            if (_activeTabIndex >= 0 && _activeTabIndex < _tabPanels.Count)
            {
                var active = _tabPanels[_activeTabIndex];
                if (!active.IsDisposed)
                {
                    _pnlContent.Controls.Add(active);
                }
            }
        }

        // ==================== 公共代理方法 ====================

        /// <summary>获取当前激活的 FileBrowserPanel</summary>
        public FileBrowserPanel? GetActivePanel()
        {
            return (_activeTabIndex >= 0 && _activeTabIndex < _tabPanels.Count)
                ? _tabPanels[_activeTabIndex]
                : null;
        }

        /// <summary>异步初始化所有页签面板（启动时使用）</summary>
        public async Task InitializeDeferredAsync()
        {
            foreach (var panel in _tabPanels)
            {
                if (!panel.IsDisposed)
                    await panel.InitializeDeferredAsync();
            }
        }

        /// <summary>应用主题到所有页签面板和页签条</summary>
        public void ApplyTheme(bool isDark)
        {
            _isDark = isDark;
            _pnlTabBar.BackColor = isDark ? _colorDarkTabBarBg : _colorLightTabBarBg;
            _pnlContent.BackColor = isDark ? Color.FromArgb(30, 30, 30) : Color.White;
            foreach (var panel in _tabPanels)
            {
                if (!panel.IsDisposed)
                    panel.ApplyTheme(isDark);
            }
            _pnlTabBar.Invalidate();
        }

        /// <summary>强制保存所有页签面板的状态</summary>
        public void ForceSaveState()
        {
            for (int i = 0; i < _tabPanels.Count; i++)
            {
                var panel = _tabPanels[i];
                if (!panel.IsDisposed)
                {
                    var state = panel.GetCurrentState();
                    AppStateStore.SetTabState(_paneIndex, i, state);
                }
            }
            AppStateStore.SetPaneTabCount(_paneIndex, _tabPanels.Count);
            AppStateStore.SetPaneActiveTab(_paneIndex, _activeTabIndex);
        }

        /// <summary>刷新所有页签面板的当前目录</summary>
        public void RefreshCurrentDirectory()
        {
            foreach (var panel in _tabPanels)
            {
                if (!panel.IsDisposed)
                    panel.RefreshCurrentDirectory();
            }
        }

        // ==================== 页签条自绘 ====================

        private void OnTabBarPaint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            var rect = _pnlTabBar.ClientRectangle;

            // 底部分隔线
            using var penBorder = new Pen(_isDark ? _colorDarkTabBorder : _colorLightTabBorder, 1);
            g.DrawLine(penBorder, 0, rect.Bottom - 1, rect.Right, rect.Bottom - 1);

            // 计算每个页签宽度
            int addBtnX = rect.Right - AddButtonWidth;
            int tabAreaWidth = addBtnX;
            int tabCount = _tabPanels.Count;
            int tabWidth = tabCount > 0
                ? Math.Min(TabMaxWidth, Math.Max(TabMinWidth, tabAreaWidth / tabCount))
                : TabMinWidth;

            // 绘制各页签
            using var fontTab = new Font("Microsoft YaHei UI", 9f, FontStyle.Regular);
            using var brushText = new SolidBrush(_isDark ? _colorDarkTabText : _colorLightTabText);
            using var brushTextSecondary = new SolidBrush(_isDark ? _colorDarkTabTextSecondary : _colorLightTabTextSecondary);
            using var brushActive = new SolidBrush(_isDark ? _colorDarkTabActiveBg : _colorLightTabActiveBg);
            using var brushHover = new SolidBrush(_isDark ? _colorDarkTabHoverBg : _colorLightTabHoverBg);
            Color accentColor = _isDark ? _colorDarkAccent : _colorLightAccent;

            for (int i = 0; i < tabCount; i++)
            {
                int x = i * tabWidth;
                var tabRect = new Rectangle(x, 0, tabWidth, TabBarHeight - 1);
                bool isActive = (i == _activeTabIndex);
                bool isHover = (i == _hoveredTab && !isActive);

                // 背景：激活页签画白底（与内容区无缝衔接），悬停页签画半透明叠加，其余不画（融入页签条）
                if (isActive)
                    g.FillRectangle(brushActive, tabRect);
                else if (isHover)
                    g.FillRectangle(brushHover, tabRect);

                // 激活页签：底部 2px 强调色细线（现代精致风，替代顶部 3px 粗条）
                if (isActive)
                {
                    using var brushAccent = new SolidBrush(accentColor);
                    g.FillRectangle(brushAccent, tabRect.Left, tabRect.Bottom - 2, tabRect.Width, 2);
                }

                // 页签文字（激活态通过颜色与背景区分，不加粗以避免渲染模糊）
                string displayName = GetTabDisplayName(i);
                var textBrush = isActive ? brushText : brushTextSecondary;
                var fontToUse = fontTab;
                var sf = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter,
                    FormatFlags = StringFormatFlags.NoWrap,
                };
                int textWidth = tabRect.Width - TabPaddingX - (tabCount > 1 ? TabCloseButtonSize + TabCloseRightMargin : 0);
                var textRect = new Rectangle(tabRect.Left + TabPaddingX / 2, tabRect.Top, textWidth, tabRect.Height);
                g.DrawString(displayName, fontToUse, textBrush, textRect, sf);

                // 关闭按钮（×）：仅当页签数 > 1 时显示
                if (tabCount > 1)
                {
                    var closeRect = GetCloseButtonRect(i, tabWidth);
                    bool isCloseHover = (i == _hoveredCloseTab);

                    if (isCloseHover)
                    {
                        // 悬停时画圆形背景
                        using var closeBgBrush = new SolidBrush(_isDark ? Color.FromArgb(70, 70, 70) : Color.FromArgb(220, 220, 220));
                        g.FillEllipse(closeBgBrush, closeRect);
                    }

                    // 画 × 符号
                    using var closePen = new Pen(isCloseHover
                        ? (_isDark ? Color.FromArgb(255, 90, 90) : Color.FromArgb(200, 40, 40))
                        : (_isDark ? Color.FromArgb(170, 170, 170) : Color.FromArgb(120, 120, 120)), 1.5f);
                    int pad = 3;
                    g.DrawLine(closePen, closeRect.Left + pad, closeRect.Top + pad, closeRect.Right - pad, closeRect.Bottom - pad);
                    g.DrawLine(closePen, closeRect.Right - pad, closeRect.Top + pad, closeRect.Left + pad, closeRect.Bottom - pad);
                }
            }

            // 绘制 "+" 新增按钮
            var btnRect = new Rectangle(addBtnX + 4, 4, AddButtonWidth - 8, TabBarHeight - 9);
            bool canAdd = _tabPanels.Count < MaxTabCount;
            // 悬停时用强调色，未悬停用次级文字色；不可用时用禁用色
            var btnColor = canAdd
                ? (_hoverAddButton
                    ? accentColor
                    : (_isDark ? _colorDarkAddBtn : _colorLightAddBtn))
                : (_isDark ? Color.FromArgb(50, 50, 53) : Color.FromArgb(200, 215, 230));

            using var penBtn = new Pen(btnColor, 1.5f);
            // 横线
            int midX = btnRect.Left + btnRect.Width / 2;
            int midY = btnRect.Top + btnRect.Height / 2;
            g.DrawLine(penBtn, midX - 5, midY, midX + 5, midY);
            // 竖线
            if (canAdd)
                g.DrawLine(penBtn, midX, midY - 5, midX, midY + 5);
        }

        /// <summary>获取页签显示名（文件夹名，根目录显示盘符或路径末段）</summary>
        private string GetTabDisplayName(int tabIndex)
        {
            string path = GetTabFullPath(tabIndex);
            if (string.IsNullOrEmpty(path))
                return $"页签 {tabIndex + 1}";

            try
            {
                // 根目录（如 C:\）直接显示
                if (path.Length <= 3 && path.EndsWith(":\\"))
                    return path;

                return Path.GetFileName(path.TrimEnd('\\', '/'));
            }
            catch
            {
                return path;
            }
        }

        /// <summary>获取页签对应的完整路径</summary>
        private string GetTabFullPath(int tabIndex)
        {
            if (tabIndex < 0 || tabIndex >= _tabPanels.Count) return "";
            var panel = _tabPanels[tabIndex];
            return panel?.CurrentDirectory ?? "";
        }

        /// <summary>获取页签关闭按钮的矩形区域</summary>
        private Rectangle GetCloseButtonRect(int tabIndex, int tabWidth)
        {
            int x = tabIndex * tabWidth;
            int closeX = x + tabWidth - TabCloseRightMargin - TabCloseButtonSize;
            int closeY = (TabBarHeight - TabCloseButtonSize) / 2;
            return new Rectangle(closeX, closeY, TabCloseButtonSize, TabCloseButtonSize);
        }

        /// <summary>子面板目录变更回调：刷新页签标题（解决页签内导航后标题延迟更新）</summary>
        private void OnPanelDirectoryChanged(string newPath)
        {
            _pnlTabBar.Invalidate();
        }

        /// <summary>子面板状态变更回调：写入 AppStateStore（仅更新内存，不持久化）</summary>
        private void OnPanelStateChanged(FileBrowserPanel sender, TabPageState state)
        {
            int tabIndex = sender.TabIndex;
            if (tabIndex >= 0 && tabIndex < _tabPanels.Count)
            {
                AppStateStore.SetTabState(_paneIndex, tabIndex, state);
            }
        }

        // ==================== 页签条鼠标交互 ====================

        private void OnTabBarMouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;

            // 计算页签宽度（用于关闭按钮命中检测）
            int addBtnX = _pnlTabBar.ClientRectangle.Right - AddButtonWidth;
            int tabAreaWidth = addBtnX;
            int tabCount = _tabPanels.Count;
            int tabWidth = tabCount > 0
                ? Math.Min(TabMaxWidth, Math.Max(TabMinWidth, tabAreaWidth / tabCount))
                : TabMinWidth;

            // 检查是否点击了关闭按钮
            if (tabCount > 1)
            {
                int tabIndex = GetTabIndexAtX(e.X);
                if (tabIndex >= 0)
                {
                    var closeRect = GetCloseButtonRect(tabIndex, tabWidth);
                    if (closeRect.Contains(e.Location))
                    {
                        CloseTab(tabIndex);
                        return;
                    }
                }
            }

            // 检查是否点击了 "+" 按钮
            if (e.X >= addBtnX)
            {
                if (_tabPanels.Count < MaxTabCount)
                {
                    // 新增页签：默认打开当前激活页签的目录
                    string currentPath = GetTabFullPath(_activeTabIndex);
                    AddTab(currentPath);
                }
                return;
            }

            // 检查点击了哪个页签
            int tabIndexClick = GetTabIndexAtX(e.X);
            if (tabIndexClick >= 0)
            {
                SwitchToTab(tabIndexClick);
            }
        }

        private void OnTabBarMouseMove(object? sender, MouseEventArgs e)
        {
            // 计算页签宽度（用于关闭按钮命中检测）
            int addBtnX = _pnlTabBar.ClientRectangle.Right - AddButtonWidth;
            int tabAreaWidth = addBtnX;
            int tabCount = _tabPanels.Count;
            int tabWidth = tabCount > 0
                ? Math.Min(TabMaxWidth, Math.Max(TabMinWidth, tabAreaWidth / tabCount))
                : TabMinWidth;

            // 检查是否悬停在 "+" 按钮上
            bool overAdd = (e.X >= addBtnX && _tabPanels.Count < MaxTabCount);

            // 检查悬停在哪个页签上
            int hoveredTab = GetTabIndexAtX(e.X);

            // 检查是否悬停在关闭按钮上
            int hoveredCloseTab = -1;
            if (hoveredTab >= 0 && tabCount > 1)
            {
                var closeRect = GetCloseButtonRect(hoveredTab, tabWidth);
                if (closeRect.Contains(e.Location))
                    hoveredCloseTab = hoveredTab;
            }

            if (overAdd != _hoverAddButton || hoveredTab != _hoveredTab || hoveredCloseTab != _hoveredCloseTab)
            {
                _hoverAddButton = overAdd;
                _hoveredTab = hoveredTab;
                _hoveredCloseTab = hoveredCloseTab;
                _pnlTabBar.Invalidate();
            }

            // 更新光标
            if (hoveredCloseTab >= 0 || overAdd)
                _pnlTabBar.Cursor = Cursors.Hand;
            else
                _pnlTabBar.Cursor = Cursors.Default;

            // 更新 ToolTip
            if (hoveredCloseTab >= 0)
            {
                string tip = "关闭页签";
                string currentTip = _tabToolTip.GetToolTip(_pnlTabBar);
                if (currentTip != tip)
                    _tabToolTip.SetToolTip(_pnlTabBar, tip);
            }
            else if (hoveredTab >= 0)
            {
                string fullPath = GetTabFullPath(hoveredTab);
                string displayName = GetTabDisplayName(hoveredTab);
                string tip = string.IsNullOrEmpty(fullPath) ? displayName : fullPath;
                string currentTip = _tabToolTip.GetToolTip(_pnlTabBar);
                if (currentTip != tip)
                    _tabToolTip.SetToolTip(_pnlTabBar, tip);
            }
            else if (overAdd)
            {
                string tip = _tabPanels.Count >= MaxTabCount ? "已达到页签上限（10个）" : "新增页签";
                string currentTip = _tabToolTip.GetToolTip(_pnlTabBar);
                if (currentTip != tip)
                    _tabToolTip.SetToolTip(_pnlTabBar, tip);
            }
            else
            {
                if (!string.IsNullOrEmpty(_tabToolTip.GetToolTip(_pnlTabBar)))
                    _tabToolTip.SetToolTip(_pnlTabBar, "");
            }
        }

        private void OnTabBarMouseLeave(object? sender, EventArgs e)
        {
            if (_hoveredTab != -1 || _hoverAddButton || _hoveredCloseTab != -1)
            {
                _hoveredTab = -1;
                _hoverAddButton = false;
                _hoveredCloseTab = -1;
                _pnlTabBar.Invalidate();
            }
            _pnlTabBar.Cursor = Cursors.Default;
            _tabToolTip.SetToolTip(_pnlTabBar, "");
        }

        /// <summary>根据 X 坐标计算页签索引，返回 -1 表示不在任何页签上</summary>
        private int GetTabIndexAtX(int x)
        {
            int addBtnX = _pnlTabBar.ClientRectangle.Right - AddButtonWidth;
            int tabAreaWidth = addBtnX;
            int tabCount = _tabPanels.Count;
            if (tabCount == 0) return -1;

            int tabWidth = Math.Min(TabMaxWidth, Math.Max(TabMinWidth, tabAreaWidth / tabCount));
            int index = x / tabWidth;
            if (index < 0 || index >= tabCount) return -1;
            return index;
        }

        // ==================== Dispose ====================
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _tabToolTip?.Dispose();
                foreach (var panel in _tabPanels)
                {
                    if (!panel.IsDisposed)
                    {
                        panel.DirectoryChanged -= OnPanelDirectoryChanged;
                        panel.StateChanged -= OnPanelStateChanged;
                        panel.Dispose();
                    }
                }
                _tabPanels.Clear();
            }
            base.Dispose(disposing);
        }
    }
}
