# 文本比对模块（TextCompare）

> 主工具栏 **文本比对** 打开的独立窗口。比阅读器简单得多，但有几个"对齐口径"的约定，改之前先看本文。

## 1. 文件与职责

| 文件 | 内容 |
|------|------|
| `TextCompareForm.cs` | 窗体：左右文件选择、两个 DataGridView、差异着色、只看差异行、滚动同步 |
| `TextDiffEngine.cs` | 纯逻辑差分：前后缀裁剪 + Myers，输出"按行对齐"的 `List<TextDiffLine>` |
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

- 大文件（>10 万行）能比对，但 DataGridView 绑定行数多时滚动会变慢；日志场景建议用「日志分析」。
- 只做行级差分，不做行内字符级高亮。
- 没有"忽略空白/忽略大小写"开关；需要时在 `TextDiffEngine` 的比较处加归一化，并同步补回归用例。
