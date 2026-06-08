# 版本历史差异化对比 开发指南

> ⚠️ **本模块依赖 TableBrowserControl 对外暴露的 API，修改 TableBrowserControl 时需同步检查本模块是否受影响。**

---

## 1. 概述

- **入口**：CSV/Excel 阅读器 → 文件树右键 → "历史记录对比"
- **前提条件**：文件所在目录受 SVN 或 Git 管控（`VersionControlHelper.DetectVcs()`）
- **核心功能**：选择历史记录中的两个版本，左右并排对比，差异单元格标红、差异行非差异单元格标黄

---

## 2. 涉及文件

| 文件 | 核心类 | 行数 | 职责 |
|------|--------|------|------|
| `VersionCompareForm.cs` | `VersionCompareForm : Form` | ~1470行 | 对比主窗体：历史表格 + 双 TableBrowserControl + 差异计算 + 高亮渲染 |
| `VersionControlHelper.cs` | `VersionControlHelper` (static) / `VcsType` 枚举 | ~450行 | SVN/Git 命令行交互：检测 VCS 类型、拉取历史、导出指定版本 |
| `VersionHistoryEntry.cs` | `VersionHistoryEntry` / `VersionHistoryResult` / `ExportRevisionResult` | ~56行 | 数据模型：版本历史条目、命令执行结果 |
| `FileTableData.cs` | `FileTableData` / `TableRow` | ~60行 | `IsDiffRow`、`DiffColumns`、`OriginalRowIndex` 字段被版本对比使用 |
| `TableBrowserControl.cs` | `TableBrowserControl`（对外 API） | ~ | `SetDataTable()`、`GetDataGridView()`、`GetFileData()`、`GetColumnCheckStates()`、`GetVisibleColumnIndices()`、`SetColumnVisibility()`、`DetachGridFromParent()`、`AttachGridToParent()`、`OnExternalCellFormatting` 事件 |

---

## 3. 窗体布局架构

```
┌──────────────────────────────────────────┐
│  [ ]仅显示差异行  [ ]右侧永远最新  ℹ信息  │  ← _pnlTop (40px)
├──────────────────────────────────────────┤
│  版本号 │ 提交时间 │ 提交人 │ 提交信息     │  ← _dgvHistory (DataGridView, ~160px)
├───────────────┬──────────────────────────┤
│  旧版本 - rXX │  新版本 - rYY            │  ← _lblLeftHeader / _lblRightHeader (28px)
│               │                          │
│  TableBrowser │  TableBrowser            │  ← _leftView / _rightView (TableBrowserControl)
│  Control      │  Control                 │
│  (左表)       │  (右表)                   │
└───────────────┴──────────────────────────┘
```

- `_splitMain`：水平分割（历史表格 / 对比区域）
- `_splitCompare`：垂直分割（左表 / 右表），各占 50%

---

## 4. 完整数据流

```
[用户右键点击"历史记录对比"]
        │
        ▼
[VersionCompareForm 构造]
  ├── 接收参数: filePath, vcsType, frozenRows, headerRowIndex, columnVisibility
  ├── 创建临时目录 _tempDir
  └── Shown 事件 → LoadHistoryAsync()
        │
        ▼
[LoadHistoryAsync()]
  ├── VersionControlHelper.GetSvnHistoryAsync() / GetGitHistoryAsync()
  │     └── 解析 SVN XML 或 Git log 输出 → List<VersionHistoryEntry>
  ├── 填充 _dgvHistory.DataSource
  └── 选中第0行 → LoadComparisonAsync(0)
        │
        ▼
[LoadComparisonAsync(selectedIndex)] ★ 核心流程
  │
  ├── 阶段0: DetachGridFromParent()（移除 DGV 防止逐行绘制）
  ├── 阶段1: 并行加载左右 DataTable
  │     ├── LoadDataTableAsync(leftRevision)
  │     │     ├── 检查 _dataCache（最多10个）
  │     │     ├── VersionControlHelper.ExportRevisionAsync() → 导出临时文件
  │     │     └── ParseFileToDataTable() → DataTable
  │     └── LoadDataTableAsync(rightRevision)（同上）
  │
  ├── 阶段2: 注入列可见性 + SetDataTable()
  │     ├── _leftView.SetColumnVisibility(_savedLeftColStates ?? _columnVisibility)
  │     └── _leftView.SetDataTable(leftData, ...)
  │
  ├── 阶段3: ComputeDiffAsync() ★ 差异计算
  │     ├── 逐行逐列比较字符串值
  │     ├── 结果存入 _leftDiffRows / _rightDiffRows / _leftDiffCells / _rightDiffCells
  │     └── 差异信息回写到 _leftFullFileData / _rightFullFileData 的 TableRow.IsDiffRow
  │
  ├── 阶段3.5: 列显隐投影重建
  │     ├── 保存当前列显隐状态 → 构建投影后 FileTableData → BuildDataTableFromFileData()
  │     ├── 构建行/列索引映射 (_leftRowToOriginalIndex / _leftColToOriginalIndex)
  │     └── 重新 SetDataTable() + SetColumnVisibility()
  │
  ├── 阶段4: 差异高亮
  │     ├── ApplyDiffHighlight() → 绑定 OnDiffCellFormatting 事件
  │     └── 如勾选"仅显示差异行" → BuildProjectedFileData(diffOnly:true)
  │
  ├── SetupSyncScroll() → 左右滚动同步
  └── finally: AttachGridToParent() → 一次性渲染
```

---

## 5. 差异计算系统（关键！）

### 5.1 数据结构

```csharp
// ★ 差异行原始行号集合（0-based，对应 DataTable.Rows 的索引，包含冻结行）
private HashSet<int> _leftDiffRows = new();   // 左表差异行
private HashSet<int> _rightDiffRows = new();  // 右表差异行

// ★ 差异单元格编码集合（(row << 16) | col），原始行列索引
private HashSet<int> _leftDiffCells = new();
private HashSet<int> _rightDiffCells = new();

// ★ DGV 行索引 → 原始完整 DataTable 行索引 的映射
private List<int> _leftRowToOriginalIndex = new();
private List<int> _rightRowToOriginalIndex = new();

// ★ DGV 列索引 → 原始完整 DataTable 列索引 的映射
private List<int> _leftColToOriginalIndex = new();
private List<int> _rightColToOriginalIndex = new();

// ★ 完整列数的列隐藏状态（独立保存，不依赖 TableBrowserControl 内部状态）
private bool[]? _savedLeftColStates = null;
private bool[]? _savedRightColStates = null;

private int _diffRowCount = 0;
private bool _diffComputed = false;
```

### 5.2 ComputeDiffAsync 算法

```csharp
// 逐行逐列比较字符串值
for (int row = 0; row < maxRows; row++)
{
    for (int col = 0; col < maxCols; col++)
    {
        string? leftVal = leftData.Rows[row][col]?.ToString();
        string? rightVal = rightData.Rows[row][col]?.ToString();
        if (leftVal != rightVal)
        {
            rowHasDiff = true;
            // 三态判断：
            //   左有右无 → 只加 _leftDiffCells
            //   左无右有 → 只加 _rightDiffCells
            //   两边都有但不同 → 都加
        }
    }
    if (rowHasDiff)
    {
        _leftDiffRows.Add(row);
        _rightDiffRows.Add(row);
        // 差异单元格用 (row << 16) | col 编码
    }
}
```

### 5.3 差异回写到 TableRow

差异计算完成后，必须将差异信息回写到 `_leftFullFileData` / `_rightFullFileData` 的 `TableRow` 中：

```csharp
foreach (var row in _leftFullFileData.Rows)
    row.IsDiffRow = _leftDiffRows.Contains(row.OriginalRowIndex);
```

这一步是 `BuildProjectedFileData(diffOnly: true)` 正确筛选的前提。

---

## 6. 差异高亮渲染

### 6.1 颜色定义

```csharp
_colorDiffCellBg = Color.FromArgb(255, 200, 200);   // 红色：差异单元格
_colorDiffRowBg  = Color.FromArgb(255, 253, 220);   // 浅黄色：差异行的非差异单元格
```

### 6.2 OnDiffCellFormatting 事件

- 通过 `_leftRowToOriginalIndex[e.RowIndex]` 获取原始行索引
- 通过 `_leftColToOriginalIndex[e.ColumnIndex]` 获取原始列索引（修复列隐藏后列号偏移）
- 用 `(originalRowIndex << 16) | originalColIndex` 编码判断是否为差异单元格
- 差异单元格 → 红色，差异行非差异单元格 → 浅黄色

### 6.3 事件绑定方式

使用 `TableBrowserControl.OnExternalCellFormatting` 事件（而非直接绑定 DGV.CellFormatting）：

```csharp
view.OnExternalCellFormatting -= OnDiffCellFormatting;  // 先解绑防重复
view.OnExternalCellFormatting += OnDiffCellFormatting;
```

---

## 7. 行/列索引映射系统（最复杂！）

### 7.1 三层索引

```
原始 DataTable 行索引 (含冻结行)
        │
        ▼ BuildProjectedFileData() 投影
投影后 FileTableData 行索引 (含冻结行)
        │
        ▼ BuildRowIndexMap() 跳过冻结行
DGV 显示行索引 (不含冻结行)
```

### 7.2 映射构建

```csharp
// BuildRowIndexMap: DGV行索引 → 原始行索引
// ★ 冻结行模式下跳过冻结行
private List<int> BuildRowIndexMap(FileTableData projectedData)
{
    var map = new List<int>();
    for (int i = _frozenRows; i < projectedData.Rows.Count; i++)
        map.Add(projectedData.Rows[i].OriginalRowIndex);
    return map;
}
```

- `_leftColToOriginalIndex`：DGV 列索引 → 原始 DataTable 列索引
- 由 `_leftFullFileData` 投影时的 `leftVisibleOriginalCols` 直接赋值

### 7.3 为什么需要独立保存 `_savedLeftColStates`？

`TableBrowserControl.SetDataTable()` 会重置 `_columnCheckStates` 为投影后的列数。如果只依赖 `GetColumnCheckStates()`，列隐藏状态会丢失。因此需要 `_savedLeftColStates` / `_savedRightColStates` 独立保存完整列数的状态。

---

## 8. 列显隐与差异视图的交互

### 8.1 RefreshDiffView() 流程

1. 获取当前列显隐状态（`GetColumnCheckStates()`）
2. 从 `_leftColToOriginalIndex` 获取原始列索引
3. 根据当前显隐状态筛选可见的原始列索引
4. 调用 `BuildProjectedFileData()` 重新投影
5. `SetColumnVisibility()` + `SetDataTable()` 重新注入
6. 重建 `_leftRowToOriginalIndex` / `_leftColToOriginalIndex`
7. 重新绑定 `ApplyDiffHighlight()`

### 8.2 关键注意事项

- **不能直接用 `GetVisibleColumnIndices()`**：它返回投影后的列索引，而 `_leftFullFileData` 是完整列数的数据
- **必须用原始列索引**：通过 `_leftColToOriginalIndex` 映射回原始列
- **冻结行在投影中始终保留**：`diffOnly: true` 时不跳过冻结行

---

## 9. 版本控制系统交互

### 9.1 VCS 检测

```csharp
// VersionControlHelper.DetectVcs(filePath)
// 从文件目录向上遍历，查找 .svn 或 .git 目录
```

### 9.2 历史记录拉取

| VCS | 命令 | 解析方式 |
|-----|------|----------|
| SVN | `svn log --xml -l 100 "filePath"` | XML 解析 `logentry` 元素 |
| Git | `git log --follow --pretty=format:%H|%h|%ai|%an|%s -100 -- "relativePath"` | 按 `|` 分割 |

### 9.3 版本文件导出

| VCS | 命令 |
|-----|------|
| SVN | `svn cat -r REV "filePath"` |
| Git | `git show "REV:relativePath"` |

### 9.4 命令行可用性检查

- SVN：查找 `svn.exe`（TortoiseSVN/SlikSVN/VisualSVN 安装路径）+ 运行 `svn --version` 验证
- Git：直接运行 `git --version` 验证

---

## 10. 缓存机制

- `_dataCache: Dictionary<string, DataTable>` — 最多缓存 10 个版本的 DataTable
- Key 格式：
  - 历史版本：SVN revision 号或 Git 完整 SHA
  - 当前文件：`"__CURRENT_FILE__"`
- LRU 淘汰：超过 10 个时移除第一个（最早插入的）
- 临时文件目录 `_tempDir`：`%TEMP%/wxd_vcs_diff_{GUID}/`
- 关闭时 `CleanupTempFiles()` 清理

---

## 11. 同步滚动

```csharp
private bool _syncingScroll;  // 防递归标志

// OnLeftGridScroll:
//   垂直: rightDgv.FirstDisplayedScrollingRowIndex = leftDgv.FirstDisplayedScrollingRowIndex
//   水平: rightDgv.HorizontalScrollingOffset = leftDgv.HorizontalScrollingOffset

// OnRightGridScroll: 反向同步
```

---

## 12. 数据缓存（DataTable 解析）

### 12.1 CSV 解析
- `ReadCsvToDataTable()`：自动检测分隔符、支持引号转义、UTF-8 编码
- 列名去重：同名列追加 `_1`, `_2` 后缀

### 12.2 Excel 解析
- `ReadExcelToDataTable()`：使用 `ExcelDataReader`，列名使用 A, B, C... 格式
- `GetExcelColumnName(int columnIndex)`：0→A, 25→Z, 26→AA...

---

## 13. Loading 动画

- **独立 STA 线程**：`LoadingForm` 在独立线程中运行，即使主线程阻塞也能旋转
- 父窗体 `Enabled = false` 阻止交互
- 支持动态更新详情文字：`_loadingForm.SetDetail(text)`

---

## 14. TableBrowserControl 对外 API 依赖

版本对比依赖 `TableBrowserControl` 的以下 API：

| API | 用途 | 注意事项 |
|-----|------|----------|
| `SetDataTable(DataTable, string)` | 注入数据 | 会重置 `_columnCheckStates` |
| `GetDataGridView()` | 获取内部 DGV | 用于绑定 Scroll/CellFormatting 事件 |
| `GetFileData()` | 获取 FileTableData | 保存为 `_leftFullFileData` / `_rightFullFileData` |
| `GetColumnCheckStates()` | 获取列显隐状态 | 投影后列数，需配合 `_savedLeftColStates` |
| `GetVisibleColumnIndices()` | 获取可见列索引 | 投影后索引，**不能直接用于访问完整数据** |
| `SetColumnVisibility(bool[])` | 设置列显隐状态 | 需传入完整列数的状态 |
| `DetachGridFromParent()` | 移除 DGV | 防止设置 DataSource 时逐行绘制 |
| `AttachGridToParent()` | 加回 DGV | 一次性渲染 |
| `OnExternalCellFormatting` 事件 | 差异高亮回调 | 先解绑再绑定防重复 |

---

## 15. 修改禁区

1. **不要改变 `(row << 16) | col` 编码方式** — 差异单元格的编码格式，任何修改都会导致高亮失效
2. **`BuildRowIndexMap()` 中冻结行偏移不能动** — 冻结行模式下的行映射必须跳过 `_frozenRows`
3. **`RefreshDiffView()` 中列映射重建顺序不可改变** — 必须先获取原始列索引，再根据显隐状态筛选
4. **`_savedLeftColStates` 的保存/恢复时机不可错位** — 在 `SetDataTable()` 之前保存（完整列数），之后恢复
5. **`ComputeDiffAsync()` 中的三态判断不能合并** — 左有右无、左无右有、两边都有但不同，必须分别处理
6. **`BuildProjectedFileData()` 中冻结行不能受 `diffOnly` 影响** — 冻结行始终保留
7. **`_dataCache` 的 Key 格式不能修改** — `"__CURRENT_FILE__"` 和版本号的区分逻辑依赖此格式
8. **不要直接使用 `GetVisibleColumnIndices()` 访问完整数据** — 它返回的是投影后索引，需用 `_leftColToOriginalIndex` 映射
9. **`TableBrowserControl.OnExternalCellFormatting` 必须先解绑再绑定** — 防止多次绑定导致颜色叠加错误
10. **临时文件目录 `_tempDir` 的命名格式不能改变** — 关闭时依赖此格式清理

---

## 16. 常见 Bug 模式

| 问题 | 原因 | 修复方式 |
|------|------|----------|
| 差异高亮颜色错位 | 列隐藏后 `e.ColumnIndex` 与原始列索引不对应 | 使用 `_leftColToOriginalIndex` 映射 |
| "仅显示差异行"勾选后数据变少 | `BuildProjectedFileData` 未正确判断 `IsDiffRow` | 确保差异信息已回写到 `TableRow.IsDiffRow` |
| 切换版本后列显隐状态丢失 | `SetDataTable()` 重置了 `_columnCheckStates` | 使用 `_savedLeftColStates` 独立保存 |
| 滚动不同步 | `_syncingScroll` 未正确设置 | 确保 `finally { _syncingScroll = false; }` |
| 多次绑定事件导致颜色叠加 | `OnExternalCellFormatting` 重复绑定 | 每次绑定前先 `-=` 解绑 |
| 加载时界面卡死 | 主线程阻塞 + LoadingForm 在同一线程 | LoadingForm 在独立 STA 线程运行 |
