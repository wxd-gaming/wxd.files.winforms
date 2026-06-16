@@ -1,337 +0,0 @@
# wxd.files.winforms 开发规范与技能指南

> 本文档供 AI 辅助开发使用，涵盖项目架构、编码规范、新增功能和优化功能的标准流程。

---

## 一、项目概览

### 1.1 基本信息
- **项目名称**：wxd.files.winforms
- **框架**：.NET 10.0 + Windows Forms (WinForms)
- **语言**：C# (.cs 文件统一 UTF-8 编码)
- **解决方案文件**：`wxd.files.winforms.slnx`（新格式）
- **项目文件**：`wxd.files.winforms/wxd.files.winforms.csproj`
- **入口点**：`wxd.files.winforms/Program.cs` → `Application.Run(new Form1())`
- **输出类型**：WinExe（Windows 桌面应用）
- **版本号**：1.0.0.4

### 1.2 NuGet 依赖
- `ExcelDataReader` 3.7.0
- `ExcelDataReader.DataSet` 3.7.0

**注意**：添加新 NuGet 包时，必须同时更新 `wxd.files.winforms.csproj` 中的 `<PackageReference>`。

### 1.3 命名空间
所有代码统一使用命名空间 `wxd.files.winforms`。

---

## 二、代码架构

### 2.1 文件与类职责

| 文件 | 核心类 | 职责 |
|------|--------|------|
| `Program.cs` | `Program` | 应用程序入口，`[STAThread]` |
| `Form1.cs` | `Form1 : Form` | **主窗体**：多窗格容器、工具栏、主题切换、Loading动画、布局管理 |
| `Form1.Designer.cs` | (partial) | 设计器生成代码（最小化） |
| `FileBrowserPanel.cs` | `FileBrowserPanel : UserControl` | **单窗格控件**：文件列表、导航、过滤、右键菜单、F2重命名、视图模式 |
| `CsvReaderForm.cs` | `CsvReaderForm : Form` | **CSV/Excel阅读器**：文件树+多标签页+表格显示+排序+过滤+列关联 |
| `CsvReaderStandaloneForm.cs` | `CsvReaderStandaloneForm : Form` | 独立表格查看窗体（列关联跳转时使用，无文件树） |
| `CsvSettings.cs` | `CsvSettings` (static) | **全局设置管理**：读/写 `ConfigurationManager` 和 `.wxd-work/config.json` |
| `LargeFileViewer.cs` | `LargeFileViewer : Form` | **大文件阅读器**：文件列表+多标签页 |
| `LargeFileReaderControl.cs` | `LargeFileReaderControl : UserControl` | **大文件阅读核心**：流式读取、虚拟滚动、搜索、索引、导出 |
| `WindowsShellContextMenu.cs` | `WindowsShellContextMenu` (static) | **系统右键菜单**：通过 Shell API 调用原生上下文菜单 |
| `SimpleShellContextMenu.cs` | `SimpleShellContextMenu` (static) | 简化右键菜单（备用方案） |
| `CustomCommandManagerForm.cs` | `CustomCommandManagerForm : Form` | 自定义命令管理面板（DataGridView编辑） |
| `CustomCommandItem.cs` | `CustomCommandItem` | 自定义命令数据模型 |
| `DirectoryFrozenConfig.cs` | `DirectoryFrozenConfig` | 目录级配置数据模型 |
| `ColumnLinkConfig.cs` | `ColumnLinkConfig` | 列关联配置数据模型（支持值提取） |
| `ColumnLinkChecker.cs` | `ColumnLinkChecker` (static) | 列关联颜色检查器（跨文件单元格高亮） |

### 2.2 数据流与状态管理

```
用户操作 → 控件事件 → 状态更新 → 异步加载 → UI刷新
                              ↓
                      CsvSettings 持久化
                     (ConfigurationManager
                      + .wxd-work/config.json)
```

**关键设计模式**：
- **防抖保存**：`FileBrowserPanel` 使用 `_saveDebounceTimer`（3秒延迟）避免频繁写配置
- **防抖过滤**：`_filterDebounceTimer`（300ms延迟）避免每次按键都触发过滤
- **异步加载**：文件列表加载使用 `Task.Run` + `CancellationTokenSource`
- **状态分离**：全局状态存 `dll.config`，目录级配置存 `.wxd-work/config.json`

### 2.3 配置存储层级

| 存储位置 | 内容 | 管理类 |
|----------|------|--------|
| `wxd.files.winforms.dll.config` | 窗格数量、布局、主题、各窗格目录/过滤/视图、全局冻结行、最近路径、自定义命令 | `CsvSettings` |
| `{目录}/.wxd-work/config.json` | 该目录的CSV/Excel冻结行、表头行索引、列关联、文件备注 | `CsvSettings` |

---

## 三、主题与配色规范

### 3.1 配色方案

项目支持两套主题，通过 `ThemeMode` 枚举切换：

**浅色（天蓝）配色**：
```
窗体背景:    #EBF8FF (235,248,255)
工具栏背景:  #EBF5FF (235,245,255)
边框:        #ADD1ED (173,209,237)
主文字:      #143C64 (20,60,100)
次要文字:    #5A82AA (90,130,170)
按钮主色:    #42A5F5 (66,165,245)
按钮深色:    #1E88E5 (30,136,229)
```

**深色（VS风格）配色**：
```
窗体背景:    #1E1E1E (30,30,30)
工具栏背景:  #252526 (37,37,38)
边框:        #3E3E42 (62,62,66)
主文字:      #F1F1F1 (241,241,241)
次要文字:    #C8C8C8 (200,200,200)
```

### 3.2 主题应用规则
- 所有窗体/控件必须在构造函数或 `ApplyTheme()` 方法中处理主题切换
- 使用 `_isDark` 字段判断当前主题
- 新增控件时必须同步更新 `ApplyTheme()` 方法

---

## 四、编码规范

### 4.1 命名规范
- **类名**：PascalCase（如 `FileBrowserPanel`、`CsvReaderForm`）
- **方法名**：PascalCase（如 `NavigateTo`、`ApplyTheme`）
- **私有字段**：`_camelCase` 前缀下划线（如 `_currentDir`、`_isDark`）
- **局部变量**：camelCase（如 `panelCount`、`isDark`）
- **常量**：PascalCase 或 UPPER_SNAKE_CASE（如 `GAP`、`SHGFI_ICON`）
- **控件**：`_cbo`/`_btn`/`_txt`/`_lst`/`_lbl`/`_pnl`/`_chk` 前缀 + PascalCase

### 4.2 文件编码
- 所有 `.cs` 文件使用 **UTF-8 编码**
- 创建新文件时必须确保 UTF-8 编码

### 4.3 异步操作规范
- 文件 I/O 操作使用 `async/await` + `Task.Run`
- 使用 `CancellationTokenSource` 支持取消
- 在 `catch (OperationCanceledException)` 中静默处理取消
- UI 更新必须在 UI 线程：`Invoke(() => { ... })` 或 `BeginInvoke(...)`

### 4.4 Win32 API 调用规范
- 使用 `[DllImport]` 声明，放在类顶部
- 结构体使用 `[StructLayout(LayoutKind.Sequential)]`
- COM 接口使用 `[ComImport]` + `[Guid]` + `[InterfaceType]`

### 4.5 资源释放
- 实现 `Dispose(bool disposing)` 清理托管/非托管资源
- 定时器在 Dispose 中停止和释放
- `CancellationTokenSource` 在不再需要时 Cancel 和 Dispose

---

## 五、新增功能标准流程

### 5.1 新增窗体
1. 在 `wxd.files.winforms/` 目录下创建 `{Name}.cs` 文件
2. 命名空间：`namespace wxd.files.winforms;`
3. 实现深色/浅色主题支持（参考 `CsvReaderForm.cs` 的配色字段）
4. 实现 `ApplyTheme()` 方法
5. 如需状态持久化，在 `CsvSettings` 中添加对应的 get/set 方法
6. 如需 `.resx` 资源文件，创建对应的 `{Name}.resx`

### 5.2 新增用户控件
1. 继承 `UserControl`
2. 实现 `InitializeComponent()` 和 `CreateControls()`
3. 使用 `SuspendLayout()`/`ResumeLayout()` 包裹批量控件添加
4. 设置 `Dock`、`Anchor` 确保自适应布局
5. 控件最小尺寸 `MinimumSize`

### 5.3 新增配置项
1. 在 `CsvSettings` 中添加 `const string {Key}Key = "{Key}";`
2. 添加静态属性 get/set（调用 `GetIntValue`/`SetIntValue` 或 `GetStringValue`/`SetStringValue`）
3. 如需范围校验，在 setter 中使用 `Math.Max(min, Math.Min(max, value))`

### 5.4 新增数据模型
1. 创建 `{Name}.cs` 文件
2. 类属性使用 `{ get; set; }` 自动属性
3. 如需 JSON 序列化，确保属性为 public 且有无参构造函数
4. JSON 序列化使用 `System.Text.Json.JsonSerializer`

### 5.5 在 `Form1` 中添加按钮
1. 在 `InitControlPanel()` 中创建按钮
2. 设置统一样式（`FlatStyle.Flat`、`BackColor`、`ForeColor`）
3. 添加到 `_flowPanel.Controls`
4. 注意 `Margin` 设置保证间距一致
5. 使用 `≡` 按钮的 `_showSettings` 逻辑控制可见性（可选）

---

## 六、优化功能标准流程

### 6.1 性能优化
- **文件列表加载**：使用 `Task.Run` 后台加载 + `CancellationTokenSource` 取消
- **大文件处理**：使用流式读取 + 分块 + 行偏移索引（参考 `LargeFileReaderControl`）
- **UI 防抖**：使用 `System.Windows.Forms.Timer` 实现输入防抖（300ms 过滤、3s 保存）
- **虚拟化**：大列表使用虚拟滚动（参考 `LargeFileReaderControl` 的 `HScrollListView`）

### 6.2 内存优化
- 大文件不全部加载到内存
- 使用 `ConcurrentDictionary` 缓存热数据
- 及时释放 `FileStream`、`Icon` 等资源
- `using` 语句处理 `IDisposable`

### 6.3 UI 体验优化
- 长时间操作显示 Loading 状态
- 使用 `BeginUpdate()`/`EndUpdate()` 批量更新 ListView
- `SuspendLayout()`/`ResumeLayout()` 批量添加控件
- `Application.DoEvents()` 在长循环中保持 UI 响应（谨慎使用）

### 6.4 错误处理
- 文件/目录访问使用 try-catch，显示友好提示
- `UnauthorizedAccessException` 静默跳过
- `ObjectDisposedException` 在异步回调中检查 `IsDisposed`
- 配置读写失败使用默认值，不影响主流程

---

## 七、关键代码模式

### 7.1 防抖定时器模式
```csharp
// 声明
private System.Windows.Forms.Timer? _debounceTimer;

// 初始化
_debounceTimer = new System.Windows.Forms.Timer { Interval = 300 };
_debounceTimer.Tick += (s, e) =>
{
    _debounceTimer.Stop();
    DoActualWork();
};

// 触发
_debounceTimer.Stop();
_debounceTimer.Start();
```

### 7.2 异步加载模式
```csharp
private CancellationTokenSource? _loadCts;

private async Task LoadAsync()
{
    _loadCts?.Cancel();
    _loadCts = new CancellationTokenSource();
    var token = _loadCts.Token;

    try
    {
        await Task.Run(() =>
        {
            // 后台工作，检查 token.IsCancellationRequested
        }, token);

        // UI 线程更新
        if (!token.IsCancellationRequested)
            Invoke(() => UpdateUI());
    }
    catch (OperationCanceledException) { }
}
```

### 7.3 主题切换模式
```csharp
public void ApplyTheme(bool isDark)
{
    _isDark = isDark;
    BackColor = isDark ? _colorDarkBg : _colorLightBg;
    // 更新所有子控件颜色...
    Invalidate();
}
```

### 7.4 自然排序比较器
参考 `FileBrowserPanel.NaturalSortComparer`：数字部分按数值比较，其余按字符比较。

---

## 八、禁止事项

1. **禁止修改 `wxd.files.winforms.exe`**（编译产物）
2. **禁止删除或覆盖 `.resx` 文件**中的资源
3. **禁止在 UI 线程执行同步文件 I/O**（会导致界面卡死）
4. **禁止硬编码路径**（如 `C:\xxx`），使用 `Environment.GetFolderPath` 或配置读取
5. **禁止忽略 `CancellationToken`** 在异步操作中
6. **禁止在 `Dispose` 后访问控件**，必须检查 `IsDisposed`
7. **不要创建不必要的文档文件**（.md / README 等），除非用户明确要求
8. **修改已有代码时优先使用 `replace_in_file`**，避免重写整个文件
9. **新增 NuGet 包时必须同步更新 `.csproj`**
10. **所有新建 `.cs` 文件必须是 UTF-8 编码**

---

## 九、调试与诊断

- 使用 `System.Diagnostics.Debug.WriteLine()` 输出调试信息
- 使用 `System.Diagnostics.Trace.WriteLine()` 输出追踪信息
- 关键路径（右键菜单、保存状态）已预留 Debug 输出
- 配置文件路径：`ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)`

---

## 十、版本历史

| 版本 | 日期 | 变更 |
|------|------|------|
| 1.0.0.4 | 当前 | 列关联、文件备注、目录级配置、Loading动画、自然排序、视图模式 |

---

## 十一、按功能面板的开发指南

> ⚠️ **在修改特定功能面板之前，必须先阅读对应的 skill 文档！**

### 11.1 面板 skill 文档索引

| 面板 | 文档 | 复杂度 | 核心关注点 |
|------|------|--------|-----------|
| CSV/Excel 阅读器 | [skills/CsvReader.md](skills/CsvReader.md) | ⭐⭐⭐⭐⭐ | 列显隐、列关联、双DataGridView、值提取 |
| 版本历史对比 | [skills/VersionCompare.md](skills/VersionCompare.md) | ⭐⭐⭐⭐ | 差异计算、行列索引映射、双表同步滚动、列显隐重建 |
| 文件浏览器面板 | [skills/FileBrowserPanel.md](skills/FileBrowserPanel.md) | ⭐⭐⭐ | 导航、过滤、右键菜单、防抖、自然排序 |
| 大文件阅读器 | [skills/LargeFileViewer.md](skills/LargeFileViewer.md) | ⭐⭐⭐ | 流式读取、虚拟滚动、行索引、搜索 |
| 自定义命令系统 | [skills/CustomCommand.md](skills/CustomCommand.md) | ⭐⭐ | 数据模型、按钮生成、JSON序列化 |
| 主题系统 | [skills/ThemeSystem.md](skills/ThemeSystem.md) | ⭐⭐ | 配色方案、ApplyTheme()、深色/浅色 |
| 配置持久化系统 | [skills/ConfigPersistence.md](skills/ConfigPersistence.md) | ⭐⭐ | 双层存储、JSON序列化、LRU淘汰 |

### 11.2 开发前必读流程

1. **确定要修改的功能属于哪个面板**
2. **阅读对应面板的 `skills/*.md` 文档**（特别是"修改禁区"章节）
3. **阅读全局 `skill.md` 中的编码规范**（第二节~第八节）
4. **确认修改不会违反任何"禁止事项"**（第八节）
5. **修改完成后检查常见 Bug 模式**（各面板 skill 文档中的对应章节）

### 11.3 CSV/Excel 阅读器特别提醒

CSV/Excel 阅读器是**最复杂的模块**，以下是最容易出问题的地方：

| 子系统 | 关键数据结构 | 最容易出错 |
|--------|-------------|-----------|
| 列显隐 | `_columnCheckStates` → `_visibleColumnIndices` | 状态与投影的索引映射、持久化格式兼容 |
| 列关联 | `ColumnLinkConfig` → `ColumnLinkChecker._linkLookups` | 缓存 Key 格式、值提取多值拆分顺序 |
| 冻结行 | `_dataGridHeader` + `_dataGrid` 双 DataGridView | 两个表格的投影必须同步 |
| 行过滤 | 内容过滤 + 高级过滤 | `||`/`&&` 运算符解析、过滤后 Tag 映射 |
| 版本对比 | `_leftDiffRows` + `_leftRowToOriginalIndex` + `_leftColToOriginalIndex` | 三层索引映射（原始→投影→DGV）、列显隐重建、差异编码 `(row<<16)\|col` |

**修改任何相关代码前，务必先阅读 [skills/CsvReader.md](skills/CsvReader.md)！**


## 换行符使用说明
需要使用换行符文本地方使用GlobalConst类的变量替代不要直接使用在文本内使用\r\n或者\n之类的

## 异常处理
程序捕获代码异常弹窗或者正常的弹窗提示请使用AlertForm面板进行弹窗

## 需要弹窗提示
需要使用 MessageBox.Show的地方使用AlertForm面板替代