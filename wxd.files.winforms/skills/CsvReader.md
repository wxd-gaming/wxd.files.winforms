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
| `CsvReaderForm.cs` | `CsvReaderForm : Form` | ~3293 | 阅读器主窗体、文件树、标签页管理、`ColumnSelectorDialog`、`ColumnLinkSetupDialog`、内容搜索 |
| `TableBrowserControl.cs` | `TableBrowserControl : UserControl` | ~2500+ | 表格显示核心：双DataGridView、列显隐、行过滤、排序、关联跳转高亮、单元格编辑 |
| `CsvReaderStandaloneForm.cs` | `CsvReaderStandaloneForm : Form` | ~ | 独立表格查看窗体（列关联跳转时使用，无文件树） |
| `ColumnLinkConfig.cs` | `ColumnLinkConfig` | ~110 | 列关联数据模型：源/目标文件、列索引、值提取规则（多值分隔符、提取分隔符） |
| `ColumnLinkChecker.cs` | `ColumnLinkChecker` (static) | ~370 | 列关联颜色检查器：构建缓存、单元格着色（红/白）、有效性判断、ToolTip提示 |
| `DirectoryFrozenConfig.cs` | `DirectoryFrozenConfig` | ~40 | 目录级配置模型：冻结行、表头行、列关联、文件备注、列显隐状态 |
| `FileTableData.cs` | `FileTableData` / `TableRow` | ~60 | 统一表格数据模型：行号追踪、差异标记 |
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
// 第0列始终可见（行号列），其余列根据 _columnCheckStates 决定
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
- 第0列锁定不可取消
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

- 静态全局缓存 `_linkLookups: Dictionary<string, LinkLookupEntry>`
- **Key 格式**：`"{sourceRel}|{srcCol}->{targetRel}:{targetCol}"`
  - 值提取模式下追加：`"|extract:{delimiter}:{partIndex}"`
- 每个 `LinkLookupEntry` 包含：`Values` (HashSet)、`DisplayNames` (ID→显示名)、`TargetFile`、`TargetColumnName`、`DisplayColumnName`
- `Reset()` 清除所有缓存
- `BuildLinkLookups()` 按源文件前缀清除旧缓存再重建

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

---

## 7. 行过滤系统

### 7.1 内容过滤 (_txtContentFilter)

- `MatchesFilter()` 支持 `||`（或）和 `&&`（与）运算符
- 匹配所有可见列的任意单元格值
- 输入替换：`"||"` → `" | "`，`"&&"` → `" & "`

### 7.2 高级过滤 (_txtAdvancedFilter)

- 支持列表达式：`id=123`、`name like test`、`Col0>100`
- `EvaluateExpression()` 解析表达式

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
- 进度条 + 取消按钮
- 结果以过滤树形式展示（只显示有匹配的文件）

---

## 11. 常见 Bug 模式

| Bug | 原因 | 修复 |
|-----|------|------|
| 列显隐状态与 DataTable 列数不匹配 | 重新加载文件后旧状态未重置 | 加载时检查长度，不匹配则重置 |
| 冻结行修改后滚动位置丢失 | `RefreshGridLayout()` 重新绑定数据源 | 保存/恢复 `FirstDisplayedScrollingRowIndex` |
| 关联缓存未清理 | 切换目录时旧缓存残留 | 调用 `ColumnLinkChecker.Reset()` |
| 过滤后行号错位 | Tag 映射使用了错误的索引 | 使用 `originalIndices` 修复冻结行偏移 |
| 列关联 ToolTip 显示错误 | 可见列索引→原始列索引映射出错 | 确保 `visibleColumnIndices[adjustedCol]` 正确 |

---

## 12. 关键方法速查

| 方法 | 位置 | 说明 |
|------|------|------|
| `LoadCsvAsync()` | TableBrowserControl | 异步加载CSV文件 |
| `LoadExcelAsync()` | TableBrowserControl | 异步加载Excel文件 |
| `RefreshGridLayout()` | TableBrowserControl | 刷新整个表格布局（核心） |
| `GetVisibleColumnIndices()` | TableBrowserControl | 获取可见列索引列表 |
| `GetVisibleRows()` | TableBrowserControl | 获取过滤后的可见行 |
| `RefreshFilteredDisplay()` | TableBrowserControl | 刷新主体表格显示 |
| `OnSelectColumnsClick()` | TableBrowserControl | 打开列选择对话框 |
| `BuildLinkLookups()` | ColumnLinkChecker | 构建列关联缓存 |
| `CheckCellColor()` | ColumnLinkChecker | 检查并设置单元格关联颜色 |
| `ExtractValues()` | ColumnLinkConfig | 提取关联匹配值（支持多值） |
| `SetColumnVisibility()` | CsvSettings | 持久化列显隐状态 |
| `GetColumnLinksByDirectory()` | CsvSettings | 读取目录级列关联配置 |
