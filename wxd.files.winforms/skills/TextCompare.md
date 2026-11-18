# 文本比对模块（TextCompare）

> 主工具栏 **文本比对** 打开的独立窗口。比阅读器简单得多，但有几个"对齐口径"的约定，改之前先看本文。

## 1. 文件与职责

| 文件 | 内容 |
|------|------|
| `TextCompareForm.cs` | 窗体：左右文件选择、两个 DataGridView、文本/表格两种模式、差异着色、只看差异行、滚动同步 |
| `TextDiffEngine.cs` | 文本差分：前后缀裁剪 + Myers，输出"按行对齐"的 `List<TextDiffLine>` |
| `TableDiffEngine.cs` | 表格差分：按行号一一对齐、逐单元格比较，输出 `List<TableDiffRow>` |
| `SimpleTableReader.cs` | CSV / Excel 原始单元格读取；不认表头行、不做冻结行，与阅读器完全独立 |
| `TextFileReader.cs` | 按行读取文本；UTF-8 严格采样失败时退回 GB18030（代码页 936） |
| `CsvSettings.TextCompareLastDirectory` | 记住上次选文件的目录（全局设置） |

## 2. 数据流

```
选择左右文件
  → TextFileReader.ReadAllLines（后台线程）
  → TextDiffEngine.Compare（后台线程）
  → BindGrids：同一份 _display 列表生成左右两张 DataTable
  → 两个 DataGridView 共用行号：显示行 i 在左右两边的行号就是源文件行号
  → CellFormatting 按 Kind 上色
```

**核心约定：左右两个网格的行数、行序必须完全一致。** 任何"只重绑一侧"或"只对一侧排序/过滤"的改动都会让行号错位。网格列都设了 `SortMode = NotSortable`，不要打开排序。

## 2.1 两种模式（按扩展名自动切换）

| 扩展名 | 模式 | 网格列 |
|--------|------|--------|
| txt / xml / yaml / yml / json / ini / log / 其他 | 文本 | 行号 + 整行内容 |
| csv / xlsx / xls | 表格 | 行号 + A、B、C…（按两边最大列数） |

- 只有**两边都是表格**才走表格模式；一边表格一边文本会弹提示，不猜。
- 表格模式刻意"什么都不认"：第一行就是普通数据，没有表头行 / 冻结行 / 列显隐 / 列关联。
  它和阅读器是两套代码，唯一共享的只有 `TextFileReader` 的编码识别。

## 2.2 对比前按 key 排序（JSON / YAML）

- 两边都是 `.json / .yaml / .yml` 时，操作行出现「对比前按 key 排序」复选框。
- 勾选后先分别把两边文本归一，再走文本差分；**只影响内存里的比对输入，不写回文件**。
- JSON：`JsonToolEngine.Format(sortKeys: true)`，递归排序所有对象 key，并统一 2 空格缩进
  （所以顺带消除了缩进/空行差异）。
- YAML：新增 `YamlKeySorter`，只处理"块状映射"：
  - 递归按 key 排序；序列（`- `）顺序、块标量（`|` / `>`）内容、注释内容都原样保留，注释跟随所属 key
  - 遇到锚点/别名、根是序列、多文档（`---`）、无法确定的结构时**直接放弃**，按原文比对，
    状态栏提示"key 排序失败：…"
- 状态栏成功时显示「已按 key 排序」。

## 3. 差分算法（TextDiffEngine）

1. 先裁掉公共前缀、公共后缀（配置表通常只有中间一小块变化，裁完 Myers 只处理差异段）。
2. 中间段跑 Myers O((N+M)·D)，`MaxEditDistance = 1000`；超过就把中段当"整块替换"。
   - 不能无限增大：trace 每步留一份 V 快照，内存约 O(maxD²)，1000 约 8MB。
3. 输出对齐行：
   - 一对一替换 → `Modified`（左右都有行号）
   - 一侧多出来的 → `Deleted` / `Added`（另一侧行号为 null）
   - 相同行 → `Same`
4. "只保留差异行" = `OnlyDifferences()` 过滤 `Kind != Same`。

**回归测试**（`RegressionTests`）：相同/新增/删除/修改、多段差异、空侧、超限退化。
每个用例都会校验"按左右行号拼回去必须还原原文件"，改算法后这条断言是底线。

## 3.1 表格差分（TableDiffEngine）

- **按行号一一对齐**：第 N 行对第 N 行，逐单元格 `Ordinal` 比较；不按唯一列对齐。
  - 因此中间插入一行会让后面整片显示为"修改"——这是刻意的简单口径，和 Excel 里两列并排看一致。
  - 行数不齐时多出的一侧记为新增/删除，另一侧空行占位，保证左右网格行数相同。
- 单元格差异记在 `DiffColumns`（`HashSet<int>`）；`Modified` 行只染这些格子，
  新增/删除行整行染色。列数不同也算修改（缺的列按空串比）。
- 网格列头用 Excel 字母（A/B/C…AA），不是文件内容。

## 4. 编码识别（TextFileReader）

- 读文件头 64KB，按**严格 UTF-8** 试解码；尾部最多回退 3 字节，避免切在多字节字符中间造成误判。
- 解码失败 → `Encoding.GetEncoding(936)`（GB18030）。
- 随后用 `StreamReader(detectEncodingFromByteOrderMarks: true)` 按行读，BOM 交给它处理。
- 不要改成 `File.ReadAllText` + `Split('\n')`：会丢 CRLF 语义、整文件进内存。

## 5. UI 结构与易错点

- 顶部 `TableLayoutPanel`（3 行）：左文件 / 右文件 / 操作行（比对、只看差异、状态）。
- 主体 `SplitContainer`，两个 `SplitterPanel` 各放一个 Dock=Top 的标题 Label + Dock=Fill 的 DataGridView。
- **SplitContainer 初始宽度**：句柄创建前它按默认尺寸算过一次子控件宽度，句柄创建后不会自动纠正，
  表现为"左网格窄、右网格溢出面板"。`NormalizePaneLayout()` 在 Shown 里按最终面板宽度再钳一次，别删。
- `_display` 与网格行索引一一对应（不分页、不排序），`CellFormatting` 直接用 `e.RowIndex` 查 Kind。
- 差异配色走 `Theme.CellTint()`，深色皮肤下自动压暗；不要写死 Color。
- 占位侧（新增行的左侧 / 删除行的右侧）显示灰色底，不写文字，保证复制出来是空行。

## 6. 已知取舍

- 面向"比较小的文件"：大文件（>10 万行）能算，但 DataGridView 绑定行数多时滚动会变慢；日志场景建议用「日志分析」。
- 文本只做行级差分，不做行内字符级高亮；表格只做单元格级，不做行内字符级。
- 没有"忽略空白/忽略大小写"开关；需要时在比较处加归一化，并同步补回归用例。
- 表格按行号对齐，不认唯一列；两表中间插行会整片错位显示为修改（要按 ID 对齐请用阅读器的版本对比）。
- Excel 只读第一个工作表；公式取的是 Excel 里缓存的显示值（不是公式文本）。
- CSV 分隔符自动识别（逗号/制表符/分号/竖线），单元格两端空白会被 Trim。
