# CSV/Excel 阅读器开发指南

> ⚠️ **最重要、最复杂的模块** — 修改任何代码前必须阅读本文档

---

## 1. 架构概览

```
CsvReaderForm (主窗体)
├── 文件树 (TreeView)            — 展示目录内 CSV/Excel 文件
├── 标签页 (TabControl)          — 每个标签页嵌入 TableBrowserControl
├── 内容搜索                      — 异步遍历所有文件匹配关键词
│
TableBrowserControl (表格核心控件)
├── 双 DataGridView 架构
│   ├── _dataGridHeader (Panel1) — 冻结行显示
│   └── _dataGrid (Panel2)       — 主体数据
├── 列显隐系统                    — _columnCheckStates → _visibleColumnIndices
├── 行过滤系统                    — 内容过滤 + 高级过滤
├── 排序                          — 点击列标题切换
│
ColumnLinkConfig                 — 列关联数据模型（支持值提取）
ColumnLinkChecker (static)       — 列关联颜色检查器 + 缓存
DirectoryFrozenConfig            — 目录级配置模型
FileTableData                    — 统一表格数据模型
CsvSettings (static)             — 全局/目录级配置读写
```

---

## 2. 核心文件与职责

| 文件 | 核心类 | 行数 | 职责 |
|------|--------|------|------|
| `CsvReaderForm.cs` | `CsvReaderForm : Form` | ~2759 | 阅读器主窗体、文件树、标签页管理、内容搜索；**持有本窗口的 `FileDataCache` 实例**，关窗即 `Dispose()` |
| `TableBrowserControl.cs` | `TableBrowserControl : UserControl` | ~5842 | 表格显示核心：双DataGridView、列显隐、行过滤（含 `@line`）、排序、关联跳转高亮、单元格编辑、行号自绘与跳转 |
| `TableBrowserOptions.cs` | `TableBrowserOptions` | ~73 | 表格控件的构造参数 VO：根目录、冻结行、轻量/最小 UI、右键菜单模式、`Cache` 归属 |
| `FileDataCache.cs` | `FileDataCache : IDisposable` | ~762 | 表格数据缓存（宿主实例级）：`SplitCsvRecords()` 产出物理起止行号、per-file 锁、关联 lookup |
| `LineNumberJumpDialog.cs` | `LineNumberJumpDialog : Form` | ~67 | 跳转到行号的输入框（显示"有效行号：1–N"） |
| `CsvReaderStandaloneForm.cs` | `CsvReaderStandaloneForm : Form` | ~ | 独立表格查看窗体（列关联跳转时使用，无文件树） |
| `ColumnLinkConfig.cs` | `ColumnLinkConfig` | ~110 | 列关联数据模型：源/目标文件、列索引、值提取规则（多值分隔符、提取分隔符） |
| `ColumnLinkChecker.cs` | `ColumnLinkChecker` (static) | ~370 | 列关联颜色检查器：构建缓存、单元格着色（红/白）、有效性判断、ToolTip提示 |
| `DirectoryFrozenConfig.cs` | `DirectoryFrozenConfig` | ~40 | 目录级配置模型：冻结行、表头行、列关联、文件备注、列显隐状态 |
| `FileTableData.cs` | `FileTableData` / `TableRow` | ~86 | 统一表格数据模型：`PhysicalStartLine`/`PhysicalEndLine`、`DisplayLineNumber` 唯一口径、差异标记 |
| `CsvSettings.cs` | `CsvSettings` (static) | ~ | 全局/目录级配置读写（dll.config + .wxd-work/config.json） |

---

## 3. 数据流

```
文件加载
  → LoadCsvAsync() / LoadExcelAsync()
  → DataTable _dataTable（完整数据）
  → FileTableData _fileData（行号追踪）
  → 冻结行分离
      ├── _dataGridHeader ← _headerTable（冻结行投影到可见列）
      └── _dataGrid ← filteredTable（过滤后数据投影到可见列）
  → 列关联颜色渲染（CellFormatting 事件）
  → 排序（重建 _dataTable + _fileData）
  → 持久化（冻结行/表头行/列关联/列显隐 → .wxd-work/config.json）
```

---

## 4. 列显隐系统 ⭐

### 4.1 数据结构

| 字段 | 类型 | 说明 |
|------|------|------|
| `_columnCheckStates` | `bool[]` | 每列的选中状态（长度=原始DataTable列数） |
| `_visibleColumnIndices` | `List<int>` | 当前可见列的原始列索引列表（过滤后的列→原始列映射） |

### 4.2 GetVisibleColumnIndices()

```csharp
// 第0列（主键/唯一ID列）始终可见，其余列根据 _columnCheckStates 决定
// 行号不是数据列，画在 RowHeader 上（OnRowPostPaint → TableRow.DisplayLineNumber）
public List<int> GetVisibleColumnIndices()
{
    var visibleIndices = new List<int>();
    if (_dataTable.Columns.Count > 0) visibleIndices.Add(0); // 第0列锁定
    for (int i = 1; i < _dataTable.Columns.Count; i++)
    {
        if (i < _columnCheckStates.Length && _columnCheckStates[i])
            visibleIndices.Add(i);
    }
    return visibleIndices;
}
```

### 4.3 列选择对话框 (ColumnSelectorDialog)

- 在 `CsvReaderForm.cs` 中定义（内部类）
- 显示冻结行数据预览 + 关联列标记
- 第0列（主键/唯一ID列）锁定不可取消
- 确定后 → 更新 `_columnCheckStates` → 持久化 `CsvSettings.SetColumnVisibility()` → `RefreshGridLayout()`

### 4.4 配置存储

- 存储格式：`Dictionary<string, string>`，key=文件相对路径，value=逗号分隔的bool字符串（如 `"true,false,true,..."`）
- API：`CsvSettings.GetColumnVisibility(dir, relativePath)` / `CsvSettings.SetColumnVisibility(dir, relativePath, bool[])`
- 存储位置：`.wxd-work/config.json` 中的 `ColumnVisibility` 字段

### 4.5 修改禁区

- **❌ 不要修改 `GetVisibleColumnIndices()` 中第0列锁定的逻辑**
- **❌ 列显隐状态存储格式不能改**（旧数据是逗号分隔的bool字符串）
- **❌ `_columnCheckStates` 和 `_visibleColumnIndices` 的索引映射关系不能混淆**
  - 前者按原始列索引（DataTable列数）
  - 后者按可见列顺序排列（过滤后的索引→原始索引映射）

---

## 5. 列关联系统 ⭐⭐

### 5.1 数据模型 (ColumnLinkConfig)

| 字段 | 类型 | 说明 |
|------|------|------|
| `SourceFile` | `string` | 源文件相对路径 |
| `SourceColumnIndex` | `int` | 源文件关联列索引 |
| `TargetFile` | `string` | 目标文件相对路径 |
| `TargetColumnIndex` | `int` | 目标文件被搜索的列索引 |
| `TargetColumnName` | `string` | 目标列表头名称（显示用） |
| `TargetDisplayColumnIndex` | `int` | 目标文件显示名列索引（-1=不使用） |
| `TargetDisplayColumnName` | `string` | 显示列表头名称 |
| `ValueExtractionEnabled` | `bool` | 是否启用值提取 |
| `ExtractDelimiter` | `string` | 提取分隔符（如 `"#"`） |
| `ExtractPartIndex` | `int` | 提取的分段索引（0-based） |
| `MultiValueDelimiter` | `string` | 多值分隔符（如 `"|"`） |

### 5.2 ExtractValues() 值提取流程 ⚠️

```
原始值: "4002#1|4003#1"
  ↓ 第一步：按 MultiValueDelimiter("|") 拆分
  ["4002#1", "4003#1"]
  ↓ 第二步：每组按 ExtractDelimiter("#") 拆分，取 ExtractPartIndex(0)
  ["4002", "4003"]
  ↓ 去重返回
```

**顺序不可颠倒！** 先多值拆分，再提取分段。

### 5.3 ColumnLinkChecker 缓存机制

- 关联 lookup 存在 `FileDataCache` **实例**的 `_linkLookups: Dictionary<string, LinkLookupEntry>` 里（不是静态全局）：哪一窗口的缓存就只服务哪一窗口，窗口关闭 `Dispose()` 时一起释放
- **Key 格式**：`"{sourceRel}|{srcCol}->{targetRel}:{targetCol}"`
  - 值提取模式下追加：`"|extract:{delimiter}:{partIndex}"`
- 每个 `LinkLookupEntry` 包含：`Values` (HashSet)、`DisplayNames` (ID→显示名)、`TargetFile`、`TargetColumnName`、`DisplayColumnName`
- `ColumnLinkChecker` 是静态类，lookups 由调用方把缓存实例传进去：`CheckCellColor(_cache, ...)` / `IsLinkValueValid(_cache, ...)`（内部统一走 `_cache.GetLookups()`）
- `BuildLinkLookupsAsync()` 按源文件相对路径前缀清除旧缓存再重建（`_buildingLinkLookups` 防重入，已 Dispose 的缓存直接放弃）

### 5.4 关联颜色渲染

- `CellFormatting` 事件 → `ColumnLinkChecker.CheckCellColor()`
- 调用方负责跳过冻结行（`if (sender == _dataGridHeader) return;`）
- 不匹配 → 红色背景 `Color.FromArgb(255, 180, 180)`
- 匹配 → 恢复默认色
- 多值匹配：任一不匹配则标红

### 5.5 关联跳转

- 双击关联列单元格 → `LinkTableRequested` 事件 → `CsvReaderStandaloneForm`
- 传递高亮信息：`SetHighlight(colIndex, value, colName, linkConfig)`
- 目标表定位到对应行并高亮（黄色背景）

### 5.6 关联配置流程

1. `ColumnSelectorDialog` → "关联设置"按钮 → `ColumnLinkSetupDialog`
2. 选择目标文件 → 显示冻结行数据 → 选择目标列 → 可选显示列 → 值提取规则
3. 配置存入 `_columnLinks` → 通过 `ColumnSelectorDialog.GetColumnLinks()` 返回
4. 最终持久化到 `.wxd-work/config.json`

### 5.7 修改禁区

- **❌ 不要改变 `_linkLookups` 的 Key 格式**（任何修改都会导致缓存失效）
- **❌ `ExtractValues()` 中先多值拆分再提取分段的顺序不可颠倒**
- **❌ `CheckCellColor()` 中的颜色值不要硬编码修改**
- **❌ `BuildLinkLookups()` 中缓存随文件加载构建，需注意多文件场景**
- **❌ 跳过冻结行的逻辑由调用方在 CellFormatting 中判断 `e.RowIndex < _frozenRows`**

---

## 6. 冻结行与表头行

### 6.1 双 DataGridView 架构

- `_dataGridHeader`（Panel1）：显示冻结行，ScrollBars=None
- `_dataGrid`（Panel2）：显示主体数据
- `_pnlGridContainer`：SplitContainer 垂直分割，`FixedPanel = Panel1`

### 6.2 冻结行 = 0 时

- `Panel1Collapsed = true`（冻结行面板折叠）
- `_dataGrid.DataSource = _dataTable`（主体直接绑定完整数据）
- 不需要 `_visibleColumnIndices` 投影

### 6.3 表头行索引 (_headerRowIndex)

- 控制列标题显示内容
- `_headerRowIndex >= 0`：用 DataTable 中该行的值作为列标题
- `_headerRowIndex < 0`：用 Excel 列名（A, B, C...）
- **必须 < _frozenRows**（表头行必须在冻结行范围内）

### 6.4 RefreshGridLayout() 核心流程

```
1. GetVisibleColumnIndices()  → _visibleColumnIndices
2. 冻结行 > 0:
   a. 构建 _headerTable（冻结行投影到可见列）
   b. _dataGridHeader.DataSource = _headerTable
   c. 设置列标题（_headerRowIndex 对应的行值 或 Excel列名）
   d. Panel1Collapsed = false
   e. SplitterDistance = _frozenRows * 24 + 28
   f. 内容过滤 → GetVisibleRows() → RefreshFilteredDisplay()
3. 冻结行 = 0:
   a. Panel1Collapsed = true
   b. _dataGrid.DataSource = _dataTable
   c. 设置列标题
   d. SetRowTags()
```

### 6.5 修改禁区

- **❌ 两个 DataGridView 的投影必须同步**（冻结行和主体数据使用相同的 `_visibleColumnIndices`）
- **❌ `_headerRowIndex` 必须 < `_frozenRows`**
- **❌ 冻结行修改后需用 `FirstDisplayedScrollingRowIndex` 恢复滚动位置**

### 6.6 行号显示（RowHeader 自绘）

行号不是数据列，`__RowNum__` 隐藏列已删除，两个表格都在 `RowPostPaint` 里自绘。

| 位置 | 方法 | 取值 |
|------|------|------|
| 主体区 | `OnRowPostPaint` → `GetDisplayLineNumber(e.RowIndex)` | `_dataGrid.Rows[i].Tag as TableRow` → `DisplayLineNumber` |
| 冻结区 | `OnFrozenRowPostPaint` → `GetFrozenDisplayLineNumber(e.RowIndex)` | 冻结区没挂 Tag，按 `_fileData.Rows[i]` 取（与 `_dataTable` 前 `_frozenRows` 行一一对应） |
| 兜底 | `ResolveRowNumber(tagRow==null)` | `displayRowIndex + 1 + frozenRows`（仅刷新中间态，Tag 还没挂上时） |

- **`TableRow.DisplayLineNumber` 是唯一口径**：物理行号 `PhysicalStartLine` 优先 → 来源不是文件时用 `OriginalRowIndex + 1` → 空白占位行返回 `null`，由 `FormatRowNumber()` 画 `—`
- 新增/复制/删除行会把 `_fileData` 重建为 `PhysicalStartLine = -1`，退回记录序号；保存后重新解析文件才会恢复真实行号
- 行号区宽度由 `UpdateRowHeadersWidth()` 按 `MaxDisplayLineNumber()` 的位数自适应（`RowHeadersWidthSizeMode` 必须是 `DisableResizing`）

**行号区交互**

| 操作 | 入口 | 说明 |
|------|------|------|
| 标记行 | `OnRowHeaderMouseClick`（左键，Ctrl 多选） | 写 `_selectedRows`（显示行索引），`UpdateSelectedRowsLabel()` 在状态栏汇总"源文件 L{min}–L{max}" |
| 行号菜单 | `OnGridCellMouseDown` → `ShowRowNumberMenu()`（右键 `e.ColumnIndex < 0`） | 复制行号 / 只看该行号（写入 `@line=N`）/ 跳转到行号 / 清除标记；**不受 `ContextMenuMode` 限制**，阅读器、版本对比、冲突解决三处一致 |
| 按行号跳转 | `NavigateToPhysicalLine(long)` | 扫 `_dataGrid.Rows[i].Tag` 的 `DisplayLineNumber`；命中冻结区时滚 `_dataGridHeader`；未命中在 `_lblInfo` 区分"被过滤隐藏"和"没有这一行"。输入框为 `LineNumberJumpDialog` |

- **跳转落点靠 `CurrentCell`**：只读网格（版本对比/冲突解决）里 `row.Selected = true` 不生效，`RevealRow()` 必须设置 `CurrentCell` 才有可见反馈

---

## 7. 行过滤系统

### 7.1 内容过滤 (_txtContentFilter)

- `MatchesFilter()` 支持 `||`（或）和 `&&`（与）运算符
- 匹配所有可见列的任意单元格值
- 输入替换：`"||"` → `" | "`，`"&&"` → `" & "`

### 7.2 高级过滤 (_txtAdvancedFilter)

- 支持列表达式：`id=123`、`name like test`、`Col0>100`，可用 `and` / `or` / 括号组合
- **`@line` 虚拟列**：比较的是行号区显示的行号（`TableRow.DisplayLineNumber`），如 `@line=128`、`@line>=100 and @line<=200`。`CompileCondition()` 把它标成 `CondNode.IsLineFilter`，不占列名、不参与列名校验、不高亮单元格；占位行（`—`）永不匹配
- `CompileExpression()` → AST（`BinOpNode`/`CondNode`），只在提交时解析一次，逐行复用 `EvaluateCompiled(row, node)` 求值（需要整行 `TableRow`，因为 `@line` 要读行号）
- `TriggerFilter()` 先做列名校验（`FindUnknownFilterColumns`）：表达式里写了不存在的列名 → 状态栏红字提示并保留当前视图
- 列名解析由 `ResolveColumnIndex()` 完成，依次匹配 DataTable 列名 → 表头行单元格值 → 其它冻结行单元格值
- 没有操作符的裸关键词（如 `攻击力`）走"全列模糊匹配"，属于预期行为，不参与列名校验

### 7.3 过滤后的数据投影 (RefreshFilteredDisplay)

```
1. 创建 filteredTable（列 = _visibleColumnIndices 投影，类型统一为 string）
2. 遍历 visibleRows → 投影 → filteredTable.Rows.Add()
3. 记录 originalIndices（filteredTable行 → _dataTable行 映射）
4. _dataGrid.DataSource = filteredTable
5. 设置 Tag（使用 originalIndices 修复冻结行偏移）
6. 设置列标题
```

### 7.4 修改禁区

- **❌ `||`/`&&` 运算符的替换和解析逻辑不能随意调整**
- **❌ 过滤后 `Tag` 的设置必须使用 `originalIndices` 映射，否则冻结行偏移会导致行号错位**
- **❌ 列名找不到时不要再退回"全列模糊匹配"**：以前会静默出错行，现在必须在 `TriggerFilter()` 拦截并提示
- **❌ `@line` 的取值必须和 `RowHeader` 画的行号同源**（都读 `TableRow.DisplayLineNumber`），不要在过滤或跳转里另算 `OriginalRowIndex + 1`

---

## 8. 排序

- `_sortColumnIndex` + `_sortOrder`（None/Ascending/Descending）
- 点击列标题切换状态
- 排序后冻结行保持不变（`Take(_frozenRows)`）
- 排序后同步重建 `_fileData`（通过 `BuildFileData` 按 `OriginalRowIndex` 匹配）
- 排序后调用 `RefreshGridLayout()` + `UpdateColumnHeaderSortGlyph()`

---

## 9. 文件备注

- 右键文件树节点 → "设置备注名"
- 存储：`CsvSettings.SetFileRemark(dir, relativePath, remark)`
- 树节点显示：`remark (filename)`
- 标签页标题显示备注名
- 列关联 ToolTip 中使用备注名

---

## 10. 内容搜索

- 入口：`CsvReaderForm` 顶部搜索栏
- 异步遍历所有文件（Task.Run + CancellationTokenSource）
- 每个文件先走**本窗口的** `_cache.GetDataTableReadOnly(path)`（Parallel.ForEach 多线程共用同一实例，靠 per-file 锁 + ConcurrentDictionary 保证只解析一次），取不到再回退磁盘逐行读
- ★ 内容搜索会把整个目录的文件灌进本窗口的缓存实例：内存随窗口关闭释放，不会再滞留到整个进程
- 进度条 + 取消按钮
- 结果以过滤树形式展示（只显示有匹配的文件）

---

## 11. 常见 Bug 模式

| Bug | 原因 | 修复 |
|-----|------|------|
| 列显隐状态与 DataTable 列数不匹配 | 重新加载文件后旧状态未重置 | 加载时检查长度，不匹配则重置 |
| 冻结行修改后滚动位置丢失 | `RefreshGridLayout()` 重新绑定数据源 | 保存/恢复 `FirstDisplayedScrollingRowIndex` |
| 横向滚到右侧后拖一下列宽，整表跳回最左列并重绘 | 固定行网格 `ScrollBars=None`，改列宽时它的 `HorizontalScrollingOffset` 被布局夹回 0（发生在 `Width` setter 内部、`ColumnWidthChanged` 之前），再经 `Scroll` 事件回灌主表 | `SyncColumnWidths()` 同步期间屏蔽 Scroll 互写，结束后用 `_preferredHorizontalOffset` 复位两侧；该字段只由主表滚动和固定行键盘导航记账，固定行网格自己的偏移不是用户意图 |
| 关联 lookup 一直占内存 | 缓存曾是进程级静态，关窗不释放 | 缓存改成 `FileDataCache` 实例，由 `CsvReaderForm` 持有并在 `OnFormClosed` `Dispose()`；`ColumnLinkChecker` 靠调用方传入 `_cache` |
| 过滤后行号错位 | Tag 映射使用了错误的索引 | 使用 `originalIndices` 修复冻结行偏移 |
| 首屏行号是连续数字、不是物理行号（冻结行=0 时） | `BindDataTableInBatches()` 只绑 `DataSource` 不挂 Tag，`GetRowOf()` 在 Tag 缺失时退回记录序号，直到下一次 `RefreshGridLayout()` 才纠正 | 已在 `BindDataTableInBatches()` 两条路径（小表直绑 / 分批绑完）末尾补 `SetRowTags()`；新增绑定路径时同样要挂 Tag |
| 加载中途抓状态看到 `tagged=0` | `_dataGrid.DataSource = filteredTable` 之后才挂 Tag，`DataSourceChanged` 触发时行还没有 Tag | 诊断要看加载完成后的状态，别在 `DataSourceChanged` 里判对错（栈回溯也别写进事件里，`new StackTrace(true)` 会把 40 行的表拖成几十秒） |
| 加载完成后冻结行区一片白，像表头没渲染 | 遮罩 `Panel` 盖在整个控件上，`Controls.Remove()` 后 WinForms 不会重绘被遮住的子控件 | `HideLoadingOverlay()` 移除后 `this.Invalidate(covered, true)` + `this.Update()`，把子控件一起失效 |
| 列关联 ToolTip 显示错误 | 可见列索引→原始列索引映射出错 | 确保 `visibleColumnIndices[adjustedCol]` 正确 |

**尚未定位的现象（改动相关代码时留意）**

- 冻结行 > 0 的首屏偶发 `Rows[i].Tag == null`（约 3 次里 1 次），此时 `DataSource` 已经是 `filteredTable`（不是 `_dataTable`），行号退化成连续记录号；手动再调一次 `RefreshGridLayout()` 立刻恢复正常。怀疑与 `UpdateLoadingProgress()` 里的 `Application.DoEvents()` 重入有关（进度条刷新时把别的续套进来，交错执行到"绑完数据源但还没挂 Tag"的窗口里）。**未证实，别照着改**。
- 嵌在 `TabPage` 里时，`_dataGridHeader.RectangleToScreen(...)` 和 `SplitContainer.Panel1.RectangleToScreen(...)` 偶发返回 `{X=0,Y=0}`，而主表 `PointToScreen` 正常；此时冻结行区在屏幕截图里是空白。控件属性（`Visible`/`IsHandleCreated`/`Rows.Count`/`Bounds`）全部正常，`Invalidate(true)+Update()` 无效，且**是否出现不稳定**。自动化截图时先确认 `hdrScreen` 不是 `(0,0)` 再裁整窗。
- `DataGridLinkToolTip.CustomToolTipPanel.MeasureLayout()` 在悬停计时器里偶发抛 GDI+ `ExternalException`；同文件里表头单元格样式 `new Font(_textFont, FontStyle.Bold)` 逐个新建且不释放，疑似 GDI 句柄耗尽的来源。截图脚本会先把光标挪到标题栏规避。

---

## 12. 关键方法速查

| 方法 | 位置 | 说明 |
|------|------|------|
| `LoadCsvAsync()` | TableBrowserControl | 异步加载CSV文件 |
| `LoadExcelAsync()` | TableBrowserControl | 异步加载Excel文件 |
| `RefreshGridLayout()` | TableBrowserControl | 刷新整个表格布局（核心） |
| `BindDataTableInBatches()` | TableBrowserControl | 无冻结行/轻量模式的分批绑定；内部已负责 `SetRowTags()`，改这里别把 Tag 弄丢 |
| `SetRowTags()` | TableBrowserControl | 把 `_fileData.Rows` 挂到显示行 Tag 上（行号、跳转、`@line` 都依赖它） |
| `ShowLoadingOverlay()` / `UpdateLoadingProgress()` / `HideLoadingOverlay()` | TableBrowserControl | 加载遮罩三件套；`UpdateLoadingProgress()` 内含 `Application.DoEvents()`，是重入源头；`HideLoadingOverlay()` 移除遮罩后必须连子控件一起 `Invalidate` |
| `SyncColumnWidths()` | TableBrowserControl | 双 DataGridView 列宽双向同步 + 同步后复位横向偏移 |
| `_preferredHorizontalOffset` | TableBrowserControl | 用户横向滚动意图的唯一记账点；禁止用固定行网格的偏移反写主表 |
| `GetVisibleColumnIndices()` | TableBrowserControl | 获取可见列索引列表（第0列=主键列锁定） |
| `GetDisplayLineNumber()` / `GetFrozenDisplayLineNumber()` | TableBrowserControl | 行号绘制取值（主体区/冻结区） |
| `TableRow.DisplayLineNumber` | FileTableData | 行号唯一口径：物理行号 → 记录序号+1 → 占位行 null |
| `SplitCsvRecords()` | FileDataCache | CSV 切记录并产出每条记录的物理起止行号 |
| `TriggerFilter()` | TableBrowserControl | 过滤统一入口：先校验列名，再后台执行 |
| `FindUnknownFilterColumns()` | TableBrowserControl | 校验高级过滤里的列名是否存在 |
| `ShowRowNumberMenu()` | TableBrowserControl | 行号区右键菜单（复制/只看该行号/跳转/清除标记） |
| `NavigateToPhysicalLine()` | TableBrowserControl | 按行号定位并滚动；`LineNumberJumpDialog` 收输入 |
| `MaxDisplayLineNumber()` | TableBrowserControl | 当前表最大行号（行号区宽度、跳转输入范围共用） |
| `GetVisibleRows()` | TableBrowserControl | 获取过滤后的可见行 |
| `RefreshFilteredDisplay()` | TableBrowserControl | 刷新主体表格显示 |
| `OnSelectColumnsClick()` | TableBrowserControl | 打开列选择对话框 |
| `BuildLinkLookups()` | ColumnLinkChecker | 构建列关联缓存 |
| `CheckCellColor()` | ColumnLinkChecker | 检查并设置单元格关联颜色 |
| `ExtractValues()` | ColumnLinkConfig | 提取关联匹配值（支持多值） |
| `SetColumnVisibility()` | CsvSettings | 持久化列显隐状态 |
| `GetColumnLinksByDirectory()` | CsvSettings | 读取目录级列关联配置 |
