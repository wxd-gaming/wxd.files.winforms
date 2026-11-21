# 工具箱（Toolbox）

> 主工具栏 **JSON** 打开的独立窗口，已扩成六页签工具集：JSON / YAML / 时间戳 / 编码 / 哈希 / Cron。
> 每页是一个 UserControl（`ToolPanelBase` 骨架），窗体只负责页签与主题。

## 1. 文件与职责

| 文件 | 内容 |
|------|------|
| `ToolboxForm.cs` | 窗体：自绘页签、按名称建页、统一 `ApplyTheme` |
| `ToolPanelBase.cs` | 面板骨架：顶部 46px 按钮条（FlowLayoutPanel）+ 底部 26px 状态栏 |
| `JsonToolPanel.cs` | JSON 页：源文本 + 可折叠树 + 工具栏 + 搜索（原 JsonToolForm 的主体） |
| `YamlToolPanel.cs` | YAML 页：YAML↔JSON 互转、校验 YAML |
| `TimeToolPanel.cs` | 时间戳页：秒/毫秒/微秒与日期互转，实时显示当前时间 |
| `EncodingToolPanel.cs` | 编码页：Unicode 转义/还原、UTF-8 Hex、Base64、URL、HTML 编解码 |
| `HashToolPanel.cs` | 哈希页：输入即算 MD5 / SHA256（UTF-8），另有字符数与字节数 |
| `CronToolPanel.cs` | Cron 页：解析 5/6 段表达式，中文逐字段说明 + 后续触发时间 |
| `JsonToolEngine.cs` | JSON 解析/格式化/压缩/递归 key 排序/节点计数 |
| `YamlToolEngine.cs` | YAML↔JSON 互转与校验（YamlDotNet） |
| `TimeToolEngine.cs` | 时间戳 / 日期文本换算 |
| `EncodingToolEngine.cs` | 文本与转义/编码形式互转 |
| `HashToolEngine.cs` | MD5 / SHA256（输出小写 hex） |
| `CronToolEngine.cs` | cron 解析、逐字段说明、计算后续触发时间 |
| `SafeSplitContainer.cs` | SplitContainer 的安全子类（见 §4） |
| `TextFileReader` | JSON 页打开文件时的编码识别（UTF-8 / GB18030） |

回归测试在 `wxd.files.winforms.RegressionTests/ToolboxTests.cs`（引擎级，不碰 UI）。

## 2. JSON 页

### 2.1 解析

- `System.Text.Json` 的 `JsonNode`，配置 `AllowTrailingCommas = true`、`CommentHandling = Skip`——
  JSONC/带注释的配置表能直接打开。
- 严格解析失败时会先尝试给**非字符串 key**（fastjson 常见：`{1: ...}`）补引号后再解析一次；
  仍失败才报"第 X 行，第 Y 列 + 原始消息"，状态栏红字显示，树清空。
- 顶层是字面量 `null` 也算解析成功（Root 为 null，树里显示一个 `null` 节点）。

### 2.2 格式化 / 压缩 / key 排序

- 格式化：`WriteIndented = true`、`IndentSize = 2`、`UnsafeRelaxedJsonEscaping`（中文不转成 \uXXXX）。
- 压缩：不缩进输出为一行。
- key 排序：**递归**，对象按 key 的 `StringComparer.Ordinal` 升序重排；数组元素顺序不变。
  排序会把结果重新格式化写回左侧文本（一份文档，两种视图，不维护第二份状态）。
- 数字保留原写法（走 `JsonNode`/原始 JSON 文本），不经过 double 转换。
- **JSONC 注释保护**：打开时若检测到 `//` 或 `/*` 注释，不自动格式化；点"格式化/压缩/key 排序"
  会先弹确认"操作会移除注释"，取消则不动。左侧文本直接手改并保存时注释原样保留。

### 2.3 树

- 根节点文本 `$ { } (n)` / `$ [ ] (n)`；对象子节点 `"key": 值或 { } (n)`；数组子节点 `[i]: …`。
- 字符串预览超过 80 字符截断加省略号（完整内容看左侧文本）。
- 默认只展开根节点；"全部展开/全部折叠"递归处理。

### 2.4 搜索

- 匹配的是**节点文本**（key + 值预览），大小写不敏感；命中节点 `ForeColor` 标红。
- 当前命中会展开所有祖先并选中、滚动到可见；状态栏显示"第 n/m 处"。
- 源文本一变（防抖 400ms 重新解析）就清空命中集合，避免指向已释放的 TreeNode。

### 2.5 保存

- **保存（Ctrl+S）**：写回打开时的文件；没打开过文件时自动走"另存为"。
- **原子写**：先写同目录临时文件再替换目标，写到一半崩溃不会留下半截 JSON。
- **编码沿用**：打开时按 BOM 识别 UTF-8/UTF-16/UTF-32，无 BOM 再按严格 UTF-8→GB18030；
  保存时沿用同一编码（含 BOM），GBK/UTF-16 文件不会被改成 UTF-8。
- 标题里的 `*` 表示有未保存修改（页签标题同步）；有未保存修改时打开新文件会先确认。
- 打开时的自动格式化只算视图，**不标记为未保存**；用户再编辑或点格式化才标脏。

## 3. 其余五页

### 3.1 YAML

- `YAML → JSON`：标量按内容推断为 bool / 数字 / null / 字符串，缩进 2 空格输出；
  输入若是 JSON（含注释）会提示走 `JSON → YAML`。
- `JSON → YAML`：`SerializerBuilder().WithIndentedSequences()`，序列 `- ` 缩进对齐。
- `校验 YAML`：只解析不转换，报首个语法错误的行列。
- 结果直接写回同一个编辑器（与 JSON 页同口径：一份文本，按钮改变它的形式）。

### 3.2 时间戳

- 输入 10 位秒、13 位毫秒、16 位微秒（允许小数），或 ISO 8601 / 常见日期写法。
- 输出 UTC、本地时间、秒/毫秒/微秒时间戳；底部标签每秒刷新当前 UTC 时间。
- 输入框旁有"填当前毫秒"快捷按钮；空输入和乱文本都按失败处理（不猜）。

### 3.3 编码

- 左右两栏：输入 → 输出（只读），十组按钮各自独立转换；"结果→输入"可链式操作。
- Unicode 转义默认只转非 ASCII；`Unicode 还原` 支持 `\uXXXX` / `\UXXXXXXXX` 和常见转义。
- Hex / Base64 / URL / HTML 的文本与字节统一按 UTF-8；解码失败报错，不返回替换字符。

### 3.4 哈希

- 输入即时计算（`TextChanged`），MD5 / SHA256 小写 hex；非 ASCII 按 UTF-8。
- 同时显示字符数与 UTF-8 字节数，方便排查"为什么哈希和别处对不上"。

### 3.5 Cron

- 支持 5 段（分 时 日 月 周）和 6 段（秒 分 时 日 月 周，Quartz 风格）；
  支持 `* , - /`、英文月份/星期名，`?` 当 `*`。
- 解析成功显示每个字段的中文说明、一句话概括，以及接下来若干次触发时间（本地时间）。
- 不支持 `L` / `W` / `#` 扩展，遇到会明确报错。

## 4. 易错点

- **别把 TreeNode 存到字段里跨解析复用**：每次 `ParseAndRender` 都会 `Nodes.Clear()`，
  旧节点引用会失效。搜索命中集合 `_matches` 必须和树同生命周期。
- 左侧文本的自动解析走 `_parseTimer`（400ms 防抖）；程序自己改文本时必须置
  `_suppressTextChanged`，否则格式化后会被 TextChanged 再解析一遍。
- **.NET 10 的 `SplitContainer` 会在显示会话切换时从 `RepaintSplitterRect` 抛 GDI+ `ExternalException`**，
  异常跨出窗口过程会直接杀进程（.NET 11 才修）。新面板一律用 `SafeSplitContainer`：
  在 `OnLayout` 里兜住异常，并在 `BeginInvoke` 里重复 `ResumeLayout(true)` 把泄漏的布局挂起计数清零，
  否则两栏子控件会停在旧尺寸 / 0 高。
- 面板创建初期宽度很小，固定 `PanelNMinSize = 240` 会反复抛 `InvalidOperationException`；
  三个双栏页都在 `SizeChanged` + `Load` 里按当前宽度自适应并夹紧，别改回固定值。
- 树的 `ForeColor` 是唯一能逐节点改色的属性（`BackColor` 要 OwnerDraw 才生效），
  所以搜索高亮用红字而不是底色。
- 页签是 owner-draw 的；页签标题（如 JSON 页的 `*`）要靠 `ToolboxForm` 的 `SetJsonCaption` 触发重绘。
- 用 harness 往 Win32 EDIT 里灌多行样例必须用 `\r\n`；只给 `\n` 会全部挤成一行，这不是面板 bug。

## 5. 已知取舍

- 原子写只保证"不会写一半"，没有版本备份；重要文件建议先提交 VCS。
- JSONC 注释在格式化/压缩/key 排序后会被移除（已弹确认）；目前没有保留注释的格式化实现。
- JSON 页每次文本变化都会重建整棵树；超大 JSON（几万节点）会卡，定位是小配置文件。
- JSON 搜索是线性遍历全部节点，没做索引；key 排序固定递归 + 序数排序，没有"仅当前层 / 忽略大小写"选项。
- YAML 转换走通用序列化，锚点/别名、多文档（`---`）不在支持范围；注释同样会丢。
- Cron 只给"接下来几次"的推算，不做时区/DST 特化；触发时间按本地时区展示。
