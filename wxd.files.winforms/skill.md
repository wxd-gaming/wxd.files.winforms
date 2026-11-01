# wxd.files.winforms 开发规范与技能指南

> 本文档供 AI 辅助开发使用，涵盖项目架构、编码规范、新增功能和优化功能的标准流程。

---

## 一、项目概览

### 1.1 基本信息
- **项目名称**：wxd.files.winforms
- **框架**：.NET 10.0 + Windows Forms (WinForms)
- **语言**：C#（所有 `.cs` 文件统一 UTF-8 编码）
- **解决方案文件**：`wxd.files.winforms.slnx`（新格式）
- **项目文件**：`wxd.files.winforms/wxd.files.winforms.csproj`
- **入口点**：`Program.cs` → `Application.Run(new Form1())`
- **输出类型**：WinExe（Windows 桌面应用）
- **版本号**：1.0.0.0（在 `.csproj` 的 `<Version>` 中定义）

### 1.2 NuGet 依赖
| 包名 | 版本 | 用途 |
|------|------|------|
| `ExcelDataReader` | 3.7.0 | 读取 Excel 文件 |
| `ExcelDataReader.DataSet` | 3.7.0 | Excel DataSet 支持 |
| `MiniExcel` | 1.37.0 | 轻量级 Excel 读写 |

**注意**：新增 NuGet 包时必须同步更新 `wxd.files.winforms.csproj` 中的 `<PackageReference>`。

### 1.3 命名空间
所有代码统一使用命名空间 `wxd.files.winforms`（文件级命名空间声明：`namespace wxd.files.winforms;`）。

### 1.4 编译与输出
- `<AllowUnsafeBlocks>true</AllowUnsafeBlocks>` — 允许不安全代码（Win32 API 交互）
- `<ImplicitUsings>enable</ImplicitUsings>` — 隐式 using
- `<Nullable>enable</Nullable>` — 启用可空引用类型
- `<ApplicationIcon>app.ico</ApplicationIcon>` — 应用图标

---

## 二、代码架构

### 2.1 文件与类职责

#### 主窗体与布局
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `Program.cs` | `Program` | 应用程序入口，`[STAThread]`，全局异常捕获（UI线程/非UI线程/Task） |
| `Form1.cs` | `Form1 : Form` | 主窗体：多窗格容器、工具栏、主题切换、Loading动画、布局管理 |
| `Form1.Designer.cs` | (partial) | 设计器生成代码（最小化） |

#### 文件浏览
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `FileBrowserPanel.cs` | `FileBrowserPanel : UserControl` | 单窗格控件：文件列表、导航、过滤、右键菜单、F2重命名、视图模式 |
| `FileDataCache.cs` | `FileDataCache`（实例级，`IDisposable`） | 表格数据缓存 + 关联 lookup 缓存：每个目录阅读器窗口一份，窗口关闭即释放；解析/写盘/克隆为静态无状态方法 |
| `VcsStatusCache.cs` | `VcsStatusCache` | VCS（SVN/Git）状态缓存，异步获取文件状态 |

#### CSV/Excel 阅读器
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `CsvReaderForm.cs` | `CsvReaderForm : Form` | CSV/Excel 阅读器：文件树+多标签页+表格显示+排序+过滤+列关联 |
| `CsvReaderStandaloneForm.cs` | `CsvReaderStandaloneForm : Form` | 独立表格查看窗体（列关联跳转时使用，无文件树） |
| `TableBrowserControl.cs` | `TableBrowserControl : UserControl` | 表格浏览核心控件（双 DataGridView + 列显隐 + 过滤 + 排序 + 右键菜单） |
| `TableBrowserOptions.cs` | `TableBrowserOptions` | TableBrowserControl 统一配置对象（VO） |
| `FileTableData.cs` | `FileTableData` / `TableRow` | 通用表格数据模型（CSV/Excel 统一） |
| `DataGridLinkToolTip.cs` | `DataGridLinkToolTip` | DataGridView 关联字段自定义 tooltip |
| `CellEditDialog.cs` | `CellEditDialog : Form` | 单元格编辑弹窗（支持大文本编辑） |

#### 列关联系统
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `ColumnLinkConfig.cs` | `ColumnLinkConfig` | 列关联配置数据模型（支持值提取） |
| `ColumnLinkChecker.cs` | `ColumnLinkChecker` (static) | 列关联颜色检查器（跨文件单元格高亮） |
| `ColumnLinkChecker_SimpleCsv.cs` | (partial) | 简化 CSV 列关联检查扩展 |
| `ColumnLinkSetupDialog.cs` | `ColumnLinkSetupDialog : Form` | 列关联设置对话框 |
| `ColumnManagerDialog.cs` | `ColumnManagerDialog : Form` | 列管理对话框（列显隐 + 顺序调整） |
| `LinkTableSelectorDialog.cs` | `LinkTableSelectorDialog : Form` | 关联表选择对话框 |
| `TargetColumnSelectDialog.cs` | `TargetColumnSelectDialog : Form` | 目标列选择对话框 |
| `LineNumberJumpDialog.cs` | `LineNumberJumpDialog : Form` | 行号跳转输入框（行号区右键菜单调用） |

#### 版本对比系统
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `VersionCompareForm.cs` | `VersionCompareForm : Form` | 版本历史对比（差异计算、行列索引映射、双表同步滚动） |
| `VersionControlHelper.cs` | `VersionControlHelper` (static) | VCS（SVN/Git）命令行封装 |
| `VersionHistoryEntry.cs` | `VersionHistoryEntry` | 版本历史记录数据模型 |
| `ConflictResolveControl.cs` | `ConflictResolveControl : UserControl` | 冲突解决控件（合并编辑器） |
| `DuplicateDetailForm.cs` | `DuplicateDetailForm : Form` | 重复数据详情查看窗体 |

#### 大文件阅读器
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `LargeFileViewer.cs` | `LargeFileViewer : Form` | 大文件阅读器：文件列表+多标签页 |
| `LargeFileReaderControl.cs` | `LargeFileReaderControl : UserControl` | 大文件阅读核心：流式读取、虚拟滚动、搜索、索引、导出 |

#### 右键菜单
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `WindowsShellContextMenu.cs` | `WindowsShellContextMenu` (static) | 系统右键菜单：通过 Shell API 调用原生上下文菜单 |
| `SimpleShellContextMenu.cs` | `SimpleShellContextMenu` (static) | 简化右键菜单（备用方案） |

#### 配置管理
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `CsvSettings.cs` | `CsvSettings` (static) | 全局设置管理：读写 `ConfigurationManager` 和 `.wxd-work/config.json` |
| `DirectoryFrozenConfig.cs` | `DirectoryFrozenConfig` | 目录级配置数据模型 |
| `App.config` | — | 应用配置文件（编译后生成 `wxd.files.winforms.dll.config`） |

#### 自定义命令
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `CustomCommandManagerForm.cs` | `CustomCommandManagerForm : Form` | 自定义命令管理面板（DataGridView 编辑） |
| `CustomCommandItem.cs` | `CustomCommandItem` | 自定义命令数据模型 |

#### 更新系统
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `UpdateForm.cs` | `UpdateForm : Form` | 更新主窗体 |
| `UpdateHelper.cs` | `UpdateHelper` | 更新辅助类（检查/下载/安装） |
| `UpdateNotifyForm.cs` | `UpdateNotifyForm : Form` | 更新通知弹窗 |

#### 工具类与辅助
| 文件 | 核心类 | 职责 |
|------|--------|------|
| `AlertForm.cs` | `AlertForm : Form` | 统一弹窗（替代 MessageBox.Show，支持复制+异常输出） |
| `GlobalConst.cs` | `GlobalConst` (static) | 全局常量（换行符、分隔符），避免硬编码特殊字符 |
| `NaturalStringComparer.cs` | `NaturalStringComparer` | 自然字符串比较器（数字部分按数值比较） |
| `HelpForm.cs` | `HelpForm : Form` | 帮助信息窗体 |
| `ChangelogForm.cs` | `ChangelogForm : Form` | 更新日志窗体 |

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
| `{目录}/.wxd-work/config.json` | 该目录的 CSV/Excel 冻结行、表头行索引、列关联、文件备注 | `CsvSettings` |

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
| 类别 | 格式 | 示例 |
|------|------|------|
| 类名 | PascalCase | `FileBrowserPanel`、`CsvReaderForm` |
| 方法名 | PascalCase | `NavigateTo`、`ApplyTheme` |
| 私有字段 | `_camelCase` | `_currentDir`、`_isDark` |
| 局部变量 | camelCase | `panelCount`、`isDark` |
| 常量 | PascalCase 或 UPPER_SNAKE_CASE | `GAP`、`SHGFI_ICON` |
| 控件前缀 | `_cbo`/`_btn`/`_txt`/`_lst`/`_lbl`/`_pnl`/`_chk` + PascalCase | `_btnOk`、`_txtFilter` |

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

### 4.6 字符串与换行符
- **禁止在字符串中直接使用 `\r\n` 或 `\n` 等转义字符**
- 必须使用 `GlobalConst.NewLine`（`\r\n`）、`GlobalConst.Lf`（`\n`）、`GlobalConst.Cr`（`\r`）替代
- 使用 `GlobalConst.Tab` 替代 `\t`，`GlobalConst.Space` 替代 `" "`，`GlobalConst.Empty` 替代 `""`

### 4.7 弹窗规范
- **禁止直接使用 `MessageBox.Show()`**
- 所有弹窗必须使用 `AlertForm` 系列方法：
  - `AlertForm.Show(owner, message, title, type, showCopy, exception)` — 通用弹窗
  - `AlertForm.ShowError(owner, message, title, exception)` — 错误弹窗（自动显示复制按钮）
  - `AlertForm.ShowException(owner, ex, title)` — 异常弹窗（包含异常信息+复制按钮+输出堆栈）

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
- **全局未捕获异常**由 `Program.cs` 中的 `Application.ThreadException`、`AppDomain.CurrentDomain.UnhandledException`、`TaskScheduler.UnobservedTaskException` 统一捕获并弹窗提示

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
参考 `NaturalStringComparer`：数字部分按数值比较，其余按字符串比较（忽略大小写）。
```csharp
public class NaturalStringComparer : IComparer<string?>
{
    private readonly SortOrder _sortOrder;
    // 优先尝试 double.TryParse 数值比较，失败则 string.Compare
}
```

### 7.5 统一弹窗模式
```csharp
// 普通信息
AlertForm.Show(this, "操作成功", "提示", AlertType.Information);

// 错误（带复制按钮，自动显示堆栈）
AlertForm.ShowError(this, "文件读取失败", "错误", exception);

// 异常（自动包含堆栈信息）
AlertForm.ShowException(this, ex, "处理异常");
```

### 7.6 全局异常捕获模式
```csharp
// UI 线程异常
Application.ThreadException += (sender, e) =>
{
    AlertForm.ShowException(null, e.Exception, "UI 线程异常");
    Environment.Exit(1);
};

// 非 UI 线程异常
AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
{
    if (e.ExceptionObject is Exception ex)
        AlertForm.ShowException(null, ex, "非 UI 线程未处理异常");
    Environment.Exit(1);
};

// Task 未观察异常
TaskScheduler.UnobservedTaskException += (sender, e) =>
{
    AlertForm.ShowException(null, e.Exception, "Task 未处理异常");
    e.SetObserved();
};
```

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
11. **禁止在字符串中直接使用 `\r\n` 或 `\n`**，必须使用 `GlobalConst.NewLine` 等常量
12. **禁止直接使用 `MessageBox.Show()`**，必须使用 `AlertForm` 系列方法
13. **禁止在代码中捕获异常后不处理或只写 `Debug.WriteLine`**，必须使用 `AlertForm` 弹窗提示（全局异常由 `Program.cs` 中的事件统一处理）

---

## 九、调试与诊断

- 使用 `System.Diagnostics.Debug.WriteLine()` 输出调试信息
- 使用 `System.Diagnostics.Trace.WriteLine()` 输出追踪信息
- 关键路径（右键菜单、保存状态）已预留 Debug 输出
- 配置文件路径：`ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)`
- `AlertForm` 在传入 `exception` 参数时，会自动将完整异常信息输出到 Debug 输出窗口，并追加到弹窗消息文本中显示

---

## 十、版本历史

| 版本 | 日期 | 变更 |
|------|------|------|
| 1.0.0.0 | 当前 | 列关联、文件备注、目录级配置、Loading动画、自然排序、视图模式、版本对比、冲突解决、更新系统 |

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
| 列关联 | `ColumnLinkConfig` → `FileDataCache._linkLookups`（实例级，随窗口释放） | 缓存 Key 格式、值提取多值拆分顺序 |
| 冻结行 | `_dataGridHeader` + `_dataGrid` 双 DataGridView | 两个表格的投影必须同步 |
| 行过滤 | 内容过滤 + 高级过滤（`@line` 是虚拟列，比 `TableRow.DisplayLineNumber`） | `||`/`&&` 运算符解析、过滤后 Tag 映射、行号口径与 RowHeader 同源 |
| 版本对比 | `_leftDiffRows` + `_leftRowToOriginalIndex` + `_leftColToOriginalIndex` | 三层索引映射（原始→投影→DGV）、列显隐重建、差异编码 `(row<<16)\|col` |

**修改任何相关代码前，务必先阅读 [skills/CsvReader.md](skills/CsvReader.md)！**

---

## 十二、CSV/Excel 数据流规范（读取→写入→展示→版本对比→冲突解决）

> ⚠️ **所有涉及 CSV/Excel 数据的操作都必须遵循本规范！**
>
> 本规范描述基于 `FileDataCache`（实例级缓存+磁盘IO）和 `FileTableData/TableRow`（数据模型+差异标记）的完整数据流。

### 12.1 核心数据模型

#### 12.1.1 FileTableData — 通用表格数据模型

```csharp
// FileTableData.cs — CSV/Excel 统一的内存数据结构
public class FileTableData
{
    public string FilePath { get; set; } = "";           // 文件完整路径
    public string FileName => Path.GetFileName(FilePath); // 文件名
    public TableFileType Type { get; set; }                // Csv 或 Excel
    public List<string> ColumnNames { get; set; } = new(); // 列名（CSV=首行，Excel=A/B/C...）
    public List<TableRow> Rows { get; set; } = new();      // 数据行列表
    public char Delimiter { get; set; } = ',';              // CSV 分隔符
}
```

#### 12.1.2 TableRow — 数据行（含差异标记字段）

```csharp
// FileTableData.cs — 每行数据 + 版本对比/冲突解决标记
public class TableRow
{
    public List<string> Cells { get; set; } = new();          // 单元格值列表
    public int OriginalRowIndex { get; set; }                  // 记录序号（0-based，排序/筛选后不变）
    public int PhysicalStartLine { get; set; } = -1;           // 源文件物理起始行号（1-based，-1=来源不是文件）
    public int PhysicalEndLine { get; set; } = -1;             // 源文件物理结束行号（引号内换行会跨行）
    public bool IsDiffRow { get; set; }                        // 是否为差异行
    public HashSet<int>? DiffColumns { get; set; }             // 差异列索引集合（null=无差异）
    public DiffRowType DiffType { get; set; } = DiffRowType.None; // None/Modified/Added/Deleted
    public bool IsBlankRow { get; set; } = false;              // 空白占位行（版本对齐时插入）
    public int? DisplayLineNumber { get; }                     // 行号唯一取值口径：物理行号优先→记录序号+1→占位行 null(画 —)
}
```

#### 12.1.3 DiffRowType — 差异行类型枚举

```csharp
public enum DiffRowType
{
    None,       // 无差异
    Modified,   // 两表都有但内容不同
    Added,      // 只在右表存在（远程新增）
    Deleted     // 只在左表存在（本地删除）
}
```

#### 12.1.4 关键约定

| 约定 | 说明 |
|------|------|
| **DataTable + FileTableData 双模型** | 所有读取返回 `(DataTable, FileTableData)` 副本；DataTable 用于 DataGridView 显示，FileTableData 用于差异标记、列关联等 |
| **OriginalRowIndex 不变** | 排序/筛选/列显隐投影后，`OriginalRowIndex` 始终指向原始 DataTable 的行索引，用于回溯定位 |
| **差异编码 `(row<<16)|col`** | 单元格差异用 `int` 编码存储在 `HashSet<int>` 中，行号占高16位、列号占低16位 |
| **深拷贝副本** | `LoadFile()` 返回 `DataTable.Copy()` + `CloneFileData()`，调用方可安全修改不影响缓存 |
| **只读引用** | `GetDataTableReadOnly()` 返回缓存引用，调用方不应修改 |
| **缓存实例级归属** | `FileDataCache` 不是静态类：一个目录阅读器窗口（`CsvReaderForm`）持有一个实例，窗口关闭即 `Dispose()` 释放；A/B 目录窗口互不干扰 |

### 12.2 读取流程（FileDataCache → 视图）

#### 12.2.1 缓存读取 API

```
┌───────────────────────────────────────────────────────────────────┐
│  调用场景                    │  API（实例=某窗口的 FileDataCache） │
├──────────────────────────────┼───────────────────────────────────┤
│  CSV/Excel 阅读器正常加载     │  _cache.LoadFile(path)             │
│  内容搜索取只读表             │  _cache.GetDataTableReadOnly(path)  │
│  列关联 lookup 构建/读取      │  _cache.BuildLinkLookupsAsync()     │
│                             │  _cache.GetLookups()               │
│  预加载（不阻塞UI）           │  _cache.EnsureLoaded(path)         │
│  版本对比加载历史版本文件      │  FileDataCache.LoadFileDirect(path)│
│  冲突解决加载远端/本地文件     │  FileDataCache.LoadFileDirect(path)│
│  关联配置向导列预览           │  FileDataCache.LoadFileDirect(path)│
└──────────────────────────────┴───────────────────────────────────┘
```

**关键区别**：
- `LoadFile()` / `GetDataTableReadOnly()` / `BuildLinkLookupsAsync()` / `GetLookups()` / `EnsureLoaded()` / `SaveFile()` — **实例方法**，走调用方所属窗口的缓存；窗口 `Dispose()` 后不抛异常，读请求退化为"现场解析、不进字典"，写请求只落盘
- `LoadFileDirect()` / `SaveFileDirect()` / `CloneFileData()` / `ParseCsvFile()` / `ParseExcelFile()` — **静态无状态**，不碰任何缓存，用于版本对比临时文件、冲突解决的远端/本地版本、向导列预览
- 缓存条目**没有超时过期**：失效只靠 `LastWriteTime + FileSize` 校验（每次访问都比），别的窗口写了磁盘，本窗口下次读就自动重建

#### 12.2.2 CSV 解析流程（ParseCsvFile）

```
FileStream(FileShare.ReadWrite) → StreamReader(UTF8+BOM自动检测)
    → 全文读取 → DetectDelimiter(逗号/制表符/分号/管道符按首行频率)
    → SplitCsvRecords(按引号外的换行切记录，同时产出每条记录的物理起止行号)
    → ParseCsvLine(处理引号嵌套+转义)
    → DataTable + FileTableData（首行作为列名，重复列名自动加 _1/_2 后缀）
```

**关键细节**：
- `SplitCsvRecords()`：返回 `(Text, StartLine, EndLine)` 列表。逐字符扫描，引号内 `\n`/`\r\n`/`\r` **不分割记录**，但仍然计数物理行号，所以一条记录可以 L2–L3 这样跨行；`""` 转义为单个 `"`；空行仍产出一条空记录（保持与旧 `SplitCsvLines` 完全一致的行集合）
- `ParseCsvFile()`：把 `StartLine/EndLine` 写入 `TableRow.PhysicalStartLine/PhysicalEndLine`，这是行号显示的唯一数据来源
- `ParseExcelFile()`：ExcelDataReader 不暴露源表行号，赋 `r+1`（配置表数据区基本从 A1 起）
- `ParseCsvLine()`：引号内分隔符不分割字段，`Trim()` 每个字段值
- 列名重复处理：`while (usedNames.Contains(columnName)) { columnName = $"{header}_{suffix}"; suffix++; }`

#### 12.2.3 Excel 解析流程（ParseExcelFile）

```
FileStream(FileShare.ReadWrite) → ExcelDataReader(UseHeaderRow=false)
    → AsDataSet → 第一个Sheet → 列名=A/B/C...Z/AA/AB...
    → DataTable + FileTableData（所有单元格转为 string，null→""）
```

**关键细节**：
- `UseHeaderRow=false`：首行作为数据行而非列名，列名用 `GetExcelColumnName()` 生成（A, B, C...）
- 必须先 `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)` 注册编码提供者

#### 12.2.4 缓存自动失效机制

```csharp
// GetOrLoadEntry 内部逻辑（同一文件用 _fileLocks 里的独立锁 + double-check）
if (_disposed)                       // 窗口已关闭：现场解析一份，不进字典，不留内存
    return 临时条目(ParseFile(fullPath));

if (_entries.TryGetValue(key, out var entry))
{
    var fi = new FileInfo(fullPath);
    // ★ 文件被外部修改（LastWriteTime 或 FileSize 变化）→ 自动失效重建
    if (fi.Exists && (fi.LastWriteTimeUtc != entry.LastWriteTime || fi.Length != entry.FileSize))
    {
        _entries.TryRemove(key, out _);
        ClearLinkLookupsForFile(key);        // 同时清除关联查找缓存
        entry = LoadAndParseFile(key, fullPath); // 重新解析
    }
}
```

失效判据只有"文件变了"，**没有超时过期**：缓存跟着窗口生死，不需要 TTL。

#### 12.2.5 TableBrowserControl 加载流程

```
LoadCsvFile() / LoadExcelFile()
    → _cache.LoadFile(_filePath)                  // 走本窗口缓存，返回深拷贝
    → var (dataTable, fileData) = ...
    → _delimiter = fileData.Delimiter               // 保存分隔符（CSV写入时需要）
    → Invoke(() => OnDataLoaded(dataTable, dataTable.Clone(), infoText))
        → _dataTable = dataTable                    // 主 DataTable（后续编辑/保存的目标）
        → _fileData = BuildFileData(dataTable, _fileData) // 同步构建 FileTableData
        → 设置到 DataGridView + 冻结行 Header DataGridView
```

### 12.3 写入流程（视图 → FileDataCache → 磁盘）

#### 12.3.1 两种写入 API

```
┌─────────────────────────────────────────────────────────────────────┐
│  调用场景                    │  API                │ 缓存更新      │
├──────────────────────────────┼─────────────────────┼───────────────┤
│  TableBrowserControl 正常保存 │  SaveFile(path, dt) │ 更新缓存+清关联│
│  冲突解决保存合并结果         │  SaveFileDirect(path,dt)│ 不更新缓存   │
│  另存为新文件               │  SaveFileDirect(newPath,dt)│ 不更新缓存   │
└──────────────────────────────┴─────────────────────┴───────────────┘
```

#### 12.3.2 CSV 写入流程（WriteCsvToDisk）

```
DataTable → BuildCsvContent(dataTable, delimiter)
    → 逐行遍历，值含分隔符/引号/换行时自动加引号包裹
    → Encoding.UTF8.GetBytes(content)
    → FileStream(FileMode.Create, FileShare.ReadWrite) → Write + Flush
    → 失败时重试 3 次 × 300ms间隔（IOException=文件被占用）
```

**关键细节**：
- **不写列名行**：`BuildCsvContent()` 只写数据行，列名由 `DataTable.Columns` 管理
- 引号转义：值包含 `delimiter`、`"`、`\n`、`\r` 时 → `"value"` 包裹，内部 `"` → `""`
- `fs.Flush()` 强制刷新到磁盘

#### 12.3.3 Excel 写入流程（WriteExcelToDisk — 原子备份模式）

```
步骤1: 备份原文件 → fullPath + ".yyyyMMddHHmmssfff.bak"
步骤2: 删除原文件 → File.Delete(fullPath)
步骤3: MiniExcel.SaveAs(fullPath, dataTable, printHeader=false)
步骤4: 成功 → 删除备份 .bak 文件
步骤5: 失败 → 还原备份 .bak → 原文件（原子性保证）
    → 失败时同样重试 3 次 × 300ms间隔
```

**关键细节**：
- `printHeader: false`：不写列名行，与 CSV 行为一致
- 备份文件名包含毫秒级时间戳防止冲突
- 失败还原逻辑：`try { File.Move(backupPath, fullPath); } catch { }`

#### 12.3.4 TableBrowserControl 防抖保存机制

```
用户编辑单元格 → _hasUnsavedChanges = true → _unsavedRowIndices.Add(row)
    → ScheduleAutoSave(): _retrySaveTimer.Stop() + _retrySaveTimer.Start()
    → 3秒后 TryFlushToDisk():
        → _cache.SaveFile(_filePath, _dataTable, _delimiter)   // 落盘 + 重建本窗口缓存条目
        → 成功: _hasUnsavedChanges=false, HideDirtyIndicator(), RefreshGridLayout()
        → 失败: _retrySaveTimer.Start() 继续重试（显示⚠提示）
    → 关闭前: ForceFlushBeforeClose() → 仍失败则弹窗询问是否丢弃
```

**关键细节**：
- `SaveDebounceMs = 3000`（3秒防抖，避免频繁写磁盘）
- `SaveFile()` 更新缓存时**重新解析磁盘文件**，不从 DataTable 反推 —— 只有磁盘解析能拿到每条记录的物理起止行号，保存后重新打开该文件行号才不会退化成记录序号
- `_unsavedRowIndices` 记录未保存行（橙色背景标记）
- 失败时不弹窗打断用户，而是后台持续重试 + 状态栏提示

#### 12.3.5 冲突解决保存流程（SaveMergedResult）

```
步骤1: 备份到 %temp%/wxd-bak/（带SVN/Git版本号标识）
    → Path.GetTempPath() + "wxd-bak" + "{fileName}_backup{vcsRevision}_{timestamp}.bak"
步骤2: FileDataCache.SaveFileDirect(_filePath, _mergedTable) — 不走缓存直接写磁盘
步骤3: 清除VCS冲突标记 → VersionControlHelper.MarkConflictResolved(_filePath, _vcsType)
步骤4: AlertForm.Show 保存成功提示
步骤5: 触发 SaveCompleted 事件 → 通知父窗体刷新
```

### 12.4 展示流程（DataTable → DataGridView）

#### 12.4.1 双 DataGridView 模式

```
┌───────────────────────────────────────────────────┐
│  _dataGridHeader (冻结行)                         │
│  - 显示前 N 行（_frozenRows）                      │
│  - 列投影与 _dataGrid 主表完全同步                   │
│  - 不可编辑，始终固定在顶部                          │
├───────────────────────────────────────────────────┤
│  _dataGrid (主数据区)                              │
│  - 显示 _frozenRows 之后的所有行                    │
│  - 支持编辑、排序、过滤、列显隐                      │
│  - 行编辑模式：双击→进入编辑→Enter确认/Esc取消       │
└───────────────────────────────────────────────────┘
```

#### 12.4.2 列显隐投影机制

```
_columnCheckStates (bool[], 全列数)
    → _visibleColumnIndices (List<int>, 只包含可见列索引)
    → DataGridView 只显示可见列
    → 冻结行 Header 与主表使用相同的 _visibleColumnIndices
```

**关键约束**：两个 DataGridView 的列投影必须完全同步，否则冻结行与数据行错位！

#### 12.4.3 FileTableData 在展示中的角色

- `ColumnLinkChecker` 使用 `FileTableData.Rows[i].Cells[colIndex]` 进行关联查找
- 版本对比和冲突解决使用 `TableRow.IsDiffRow / DiffColumns / DiffType / IsBlankRow` 进行差异高亮
- `OriginalRowIndex` 用于在 DGV 行索引 → 原始行索引 的映射

### 12.5 版本对比数据流（VersionCompareForm）

#### 12.5.1 数据流概览

```
┌──────────────────────────────────────────────────────────────┐
│  VersionControlHelper.GetSvnInfo / GetGitLog                  │
│      → 获取版本历史列表 (List<VersionHistoryEntry>)            │
│                                                              │
│  选择两个版本 → VersionControlHelper.ExportFile               │
│      → 导出到 %temp%/wxd_vcs_diff_{guid}/ 临时目录              │
│      → left.csv / right.csv (或 .xlsx)                        │
│                                                              │
│  FileDataCache.LoadFileDirect(leftPath)                       │
│  FileDataCache.LoadFileDirect(rightPath)                      │
│      → (DataTable_left, FileTableData_left)                   │
│      → (DataTable_right, FileTableData_right)                 │
│                                                              │
│  ComputeDiffAsync() → 差异计算                                │
│  BuildAlignedTables() → 行对齐（双指针归并+空白行插入）          │
│      → leftAlignedFd / rightAlignedFd (FileTableData)          │
│                                                              │
│  BuildDataTableFromFileData() → 转为 DataTable 设置到视图      │
│  BuildRowIndexMap() → DGV行号→原始行号映射                     │
│  ApplyDiffHighlight() → CellFormatting 高亮绑定                │
│  SetupSyncScroll() → 三个视图滚动联动                          │
└──────────────────────────────────────────────────────────────┘
```

#### 12.5.2 差异计算（ComputeDiffAsync）

**ID对齐模式**（`_useIdAlign=true`，优先使用 `_uniqueColumnIndex` 配置列）：

```
1. 构建 remoteIdMap: id→行索引, localIdMap: id→行索引
2. 合并所有ID → 遍历每个ID:
   - 远程有+本地无 → _remoteDiffRows + _remoteDiffCells（全部列标记）
   - 本地有+远程无 → _localDiffRows + _localDiffCells（全部列标记）
   - 两边都有 → 逐列比较 → 差异列加入 _remoteDiffCells + _localDiffCells
     → 共同差异列 = _conflictCells（冲突单元格）
```

**行号对齐模式**（`_useIdAlign=false`，无ID列时回退）：

```
1. 逐行逐列比较，同位置不同值 → 标记差异行+差异单元格
2. 行数不同时，多出的行整体标记为差异
```

#### 12.5.3 行对齐算法（BuildAlignedTables — 双指针归并）

```
1. 复制固定行/表头行到 alignedFd（原样保留，不对齐）
2. 构建 remoteIdMap / localIdMap（ID→Fd行索引）
3. 合并排序所有ID（SortIds：数值优先排序，否则字典序）
4. 逐ID遍历:
   - 两表都有 → 匹配行（原样复制）
   - 远程有+本地无 → 新增行（远程数据行 + 本地空白行）
     - remoteRow.DiffType = Added, remoteRow.IsDiffRow = true
     - localRow = CreateBlankRow(DiffRowType.Added, IsBlankRow=true)
   - 本地有+远程无 → 删除行（远程空白行 + 本地数据行）
     - localRow.DiffType = Deleted, localRow.IsDiffRow = true
     - remoteRow = CreateBlankRow(DiffRowType.Deleted, IsBlankRow=true)
```

**空白行的作用**：保证远程和本地 alignedFd 行数完全一致，实现逐行同步滚动和高亮。

#### 12.5.4 三层索引映射

```
DGV显示行索引 → alignedFd行索引 → 原始DataTable行索引

_leftRowToOriginalIndex[i]  = alignedFd.Rows[i+_frozenRows].OriginalRowIndex
_leftColToOriginalIndex[j]  = 可见列j → 全列索引

差异编码: (originalRow << 16) | colIndex → _leftDiffCells / _rightDiffCells
```

#### 12.5.5 CellFormatting 高亮规则

```
isBlankRow=true && DiffType=Added  → _colorBlankRowForAdded (淡绿)
isBlankRow=true && DiffType=Deleted → _colorBlankRowForDeleted (淡橙)
DiffType=Added                      → _colorAddedRowBg (浅绿)
DiffType=Deleted                    → _colorDeletedRowBg (浅橙)
DiffType=Modified + conflictCell    → _colorConflictBg (深红)
DiffType=Modified + diffCell        → _colorDiffCellBg (浅红)
DiffType=Modified + 同行但非差异列   → _colorDiffRowBg (浅黄)
```

### 12.6 冲突解决数据流（ConflictResolveControl）

#### 12.6.1 数据流概览

```
┌──────────────────────────────────────────────────────────────┐
│  构造函数参数:                                                │
│    filePath, vcsType, remoteData, localData                   │
│    frozenRows, headerRowIndex, uniqueColumnIndex              │
│                                                              │
│  LoadComparisonAsync():                                      │
│    ComputeDiffAsync() → 差异标记                              │
│    BuildAlignedTables() → 行对齐                               │
│    SetDataTable() → 设置到远程/本地视图                         │
│    BuildRowIndexMap() → 行索引映射                             │
│    ApplyDiffHighlight() → 高亮绑定                             │
│    InitializeMergedTable() → 初始合并结果（默认=本地版本）        │
│                                                              │
│  用户操作 → UseRemoteSelected / UseLocalSelected / 编辑合并表  │
│    → 更新 _mergedTable 行内容                                  │
│                                                              │
│  SaveMergedResult():                                         │
│    备份到 %temp%/wxd-bak/                                      │
│    FileDataCache.SaveFileDirect(_filePath, _mergedTable)       │
│    MarkConflictResolved → 清除VCS冲突标记                      │
│    SaveCompleted事件 → 通知父窗体刷新                           │
└──────────────────────────────────────────────────────────────┘
```

#### 12.6.2 三区域布局

```
┌─────────────────────────────────────────────────┐
│  工具栏: ←使用远程此行 | 使用本地此行→ | 保存 |     │
│         全部使用远程 | 全部使用本地 | 全选远程 |     │
│         全选本地 | 状态信息                       │
├────────────────────────┬────────────────────────┤
│  远程版本（左）          │  本地版本（右）          │
│  _remoteView            │  _localView            │
│  CellFormatting=差异高亮 │  CellFormatting=差异高亮│
│  选中行→蓝色背景         │  选中行→绿色背景         │
├────────────────────────┴────────────────────────┤
│  合并结果（下）                                    │
│  _mergedView (可编辑)                              │
│  CellFormatting=冲突单元格高亮                      │
└─────────────────────────────────────────────────┘
```

#### 12.6.3 行选择与合并操作

```
点击远程行头 → _selectedRemoteRows.Add(originalRow) + _selectionFromRemote=true
点击本地行头 → _selectedLocalRows.Add(originalRow) + _selectionFromRemote=false
Ctrl+点击 → 多选 | Shift+点击 → 范围选

UseRemoteSelected():
  → 遍历 _selectedRemoteRows
  → 在 _mergedTable 中找到对应行 → 替换为远程版本内容
  → _selectedRemoteCell != null 时只替换单个单元格

UseLocalSelected(): 同理，替换为本地版本

UseAllRemote(): 所有差异行使用远程版本
UseAllLocal(): 所有差异行使用本地版本
```

#### 12.6.4 初始化合并表（InitializeMergedTable）

```
_mergedTable = _localData.Copy()  // ★ 默认=本地版本作为基底
→ 用户通过 UseRemoteSelected / UseLocalSelected 逐行替换差异
→ 也可直接在 _mergedView 中编辑单元格
```

#### 12.6.5 查找合并行（FindMergedRowById）

```
使用ID列对齐 → 在 _mergedTable 中查找对应ID的行索引
无ID对齐 → 回退到原始行索引直接定位
```

### 12.7 关联查找数据流（ColumnLink + LinkLookup）

#### 12.7.1 数据流概览

```
ColumnLinkConfig (配置) → BuildLinkKey(sourceRel|srcCol->targetRel:targetCol|extract:...)
    → _linkLookups[key] = LinkLookupEntry
        → Values: HashSet<string> (目标列所有唯一值)
        → DisplayNames: Dictionary<string,string> (ID→显示名映射)

BuildLinkLookupsAsync():
    → 防重入检查（_buildingLinkLookups）
    → 遍历所有 columnLinks
    → 异步加载目标文件到缓存
    → BuildLinkLookupFromMaster(FileTableData, targetColIndex, linkConfig)
    → 值提取: linkConfig.ValueExtractionEnabled → ExtractSteps处理
```

#### 12.7.2 关联查找缓存失效

```
SaveFile(updateCache=true) → ClearLinkLookupsForFile → 懒重建（下次访问时）
InvalidateFile() → 同时清除 _entries + _linkLookups
```

### 12.8 数据模型转换规则

#### 12.8.1 DataTable ↔ FileTableData 转换

```
// ParseFile → (DataTable, FileTableData)
DataTable.Columns = headers / A,B,C... (列名)
DataTable.Rows[r][c] = cellValue (string)
FileTableData.ColumnNames = 同 DataTable.Columns
FileTableData.Rows[r].Cells[c] = 同 DataTable.Rows[r][c]
FileTableData.Rows[r].OriginalRowIndex = r (0-based)

// SaveFile() 更新缓存 = 重新解析磁盘文件（保留物理行号），不再有 DataTable→FileTableData 的反推
// BuildDataTableFromFileData — 从FileTableData重建DataTable（版本对比/冲突解决用）

// CloneFileData — 深拷贝（复制所有字段包括 DiffColumns HashSet）
```

#### 12.8.2 禁止的做法

| ❌ 禁止 | ✅ 正确 |
|---------|---------|
| 直接修改 `GetDataTableReadOnly()` 返回的 DataTable | 使用 `_cache.LoadFile()` 获取副本后修改 |
| 把表格数据缓存声明回 `static`（跨窗口共享、关窗不释放） | 缓存实例由窗口持有：`CsvReaderForm._cache` → `TableBrowserOptions.Cache` 往下传，`OnFormClosed` 里 `Dispose()` |
| 在已 `Dispose()` 的缓存上假设它会抛异常 | 释放后的缓存读请求退化为"现场解析、不进字典"，写请求只落盘，照常调用即可 |
| 版本对比时忽略 `_frozenRows` 偏移 | `BuildRowIndexMap()` 跳过前 `_frozenRows` 行 |
| 冲突解决时用 `SaveFile()` 而非 `SaveFileDirect()` | 用 `SaveFileDirect()` 不走缓存 |
| 编辑后不设 `_hasUnsavedChanges=true` | 每次编辑后必须标记并调度防抖保存 |
| 忽略 `_delimiter` 保存时传入错误分隔符 | `SaveFile(path, dt, _delimiter)` 传入正确分隔符 |

### 12.9 完整数据流图

```
                    ┌─────────────────┐
                    │   磁盘文件        │
                    │  .csv / .xlsx    │
                    └────────┬────────┘
                             │
            ┌────────────────┼────────────────┐
            │  FileDataCache（每个阅读器窗口一份，窗口关闭即 Dispose）│
            │                                                            │
            │  ParseCsvFile / ParseExcelFile  （静态，不碰缓存）          │
            │  → (DataTable, FileTableData)                             │
            │                                                            │
            │  _cache.LoadFile() ──深拷贝副本──→ TableBrowserControl（阅读器）
            │  _cache.GetLookups() ──────────→ ColumnLinkChecker（列关联）
            │  LoadFileDirect() ──直接解析──→ VersionCompareForm（版本对比）
            │  LoadFileDirect() ──直接解析──→ ConflictResolveControl（冲突解决）
            │  LoadFileDirect() ──直接解析──→ ColumnLinkSetupDialog（列预览）
            │                                                            │
            │  _cache.SaveFile() ←──防抖保存──← TableBrowserControl（编辑）
            │  SaveFileDirect() ←─直接保存─← ConflictResolveControl（冲突合并）
            │                                                            │
            │  InvalidateFile() ←──外部修改──← FileSystemWatcher / 用户操作
            │  ClearLinkLookups ←──缓存失效──← SaveFile / InvalidateFile
            └────────────────┬────────────────┘
                             │
                    ┌────────┴────────┐
                    │   磁盘文件        │
                    │  .csv / .xlsx    │
                    └─────────────────┘
```

### 12.10 性能参数速查

| 参数 | 值 | 来源 |
|------|----|------|
| 缓存归属 | 一个 `CsvReaderForm`（目录阅读器窗口）一个 `FileDataCache` 实例，`OnFormClosed` 释放 | `CsvReaderForm._cache` |
| 缓存失效判据 | `LastWriteTime` + `FileSize` 不符（无超时过期） | `GetOrLoadEntry()` |
| 缓存Key格式 | `Path.GetFullPath(path).ToLowerInvariant()` | `NormalizePath()` |
| 写入重试次数 | 3次 | `WriteCsvToDisk` / `WriteExcelToDisk` |
| 写入重试间隔 | 300ms | `Thread.Sleep(300)` |
| 防抖保存延迟 | 3秒 | `TableBrowserControl.SaveDebounceMs` |
| Excel备份格式 | `{file}.{timestamp}.bak` | `WriteExcelToDisk` |
| 冲突备份目录 | `%temp%/wxd-bak/` | `SaveMergedResult` |
| 版本对比临时目录 | `%temp%/wxd_vcs_diff_{guid}/` | `VersionCompareForm` |
| 冲突解决临时目录 | `%temp%/wxd_conflict_{guid}/` | `ConflictResolveControl` |
| 差异编码 | `(row<<16)|col` | `_leftDiffCells` / `_conflictCells` |
| ID列默认值 | 第0列（`_uniqueColumnIndex=-1`时） | `DetectIdColumn()` |
| FileStream共享模式 | `FileShare.ReadWrite` | `ParseCsvFile` / `ParseExcelFile` |
| Excel列名生成 | A,B,C...Z,AA,AB... | `GetExcelColumnName()` |
| CSV分隔符检测 | 按首行频率：逗号/制表符/分号/管道符 | `DetectDelimiter()` |

---

**文档版本**: 2.0.0
**最后更新**: 2026-06-17
**适用范围**: wxd.files.winforms CSV/Excel 数据流全链路
**核心类**: FileDataCache.cs / FileTableData.cs / TableBrowserControl.cs / VersionCompareForm.cs / ConflictResolveControl.cs
