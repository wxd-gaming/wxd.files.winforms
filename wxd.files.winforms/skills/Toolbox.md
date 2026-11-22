# 工具箱（Toolbox）

> 主工具栏 **JSON** 打开的独立窗口，已扩成六页签工具集：JSON / YAML / 时间戳 / 编码 / HTTP / Cron。
> 每页是一个 UserControl（`ToolPanelBase` 骨架），窗体只负责页签与主题。

## 1. 文件与职责

| 文件 | 内容 |
|------|------|
| `ToolboxForm.cs` | 窗体：自绘页签、按名称建页、统一 `ApplyTheme` |
| `ToolPanelBase.cs` | 面板骨架：顶部 46px 按钮条（FlowLayoutPanel）+ 底部 26px 状态栏 |
| `JsonToolPanel.cs` | JSON 页：源文本 + 可折叠树 + 工具栏 + 搜索（原 JsonToolForm 的主体） |
| `YamlToolPanel.cs` | YAML 页：YAML↔JSON 互转、校验 YAML |
| `TimeToolPanel.cs` | 时间戳页：当前时间三种单位；秒/毫秒/纳秒分栏换算；年月日时分秒全下拉选择 |
| `EncodingToolPanel.cs` | 编码页：Unicode 转义/还原、UTF-8 Hex、Base64、URL、HTML 编解码，外加 MD5 / SHA256 哈希 |
| `HttpToolPanel.cs` | HTTP 页：轻量请求工具，方法 + URL 在顶部，参数/请求头键值网格，响应在下方 |
| `CronToolPanel.cs` | Cron 页：标准 / Quartz / Spring 三规范解析、语法参考、可视化生成器 + 后续触发时间 |
| `CronBuilderControl.cs` | Cron 可视化生成器：字段页签 + 每/周期/间隔/指定/特殊，实时拼表达式 |
| `JsonToolEngine.cs` | JSON 解析/格式化/压缩/递归 key 排序/节点计数 |
| `YamlToolEngine.cs` | YAML↔JSON 互转与校验（YamlDotNet） |
| `TimeToolEngine.cs` | 时间戳 / 日期文本换算 |
| `EncodingToolEngine.cs` | 文本与转义/编码形式互转 |
| `HashToolEngine.cs` | MD5 / SHA256 hex（默认小写，参数可切大写）+ 文件流式 MD5 |
| `HttpToolEngine.cs` | 键值对组装（URL 查询串 / JSON / Form）与响应 charset 解码，纯逻辑可测 |
| `CronToolEngine.cs` | cron 解析（标准 / Quartz / Spring）、逐字段说明、计算后续触发时间 |
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

- 顶部显示当前本地时间，并在同一行给出秒 / 毫秒 / 纳秒三个只读值；每次只刷新当前时间，不参与换算。
- **时间戳 → 时间** 分三行：秒、毫秒、纳秒各一个独立输入框和结果框，按对应单位显式解析，不按数值大小猜。
  回车或点行内「转换」执行；纳秒只保留 .NET tick 的 100ns 精度，结果行显示 7 位小数。
- **具体时间 → 时间戳**：年（1900–2100）、月、日、时、分、秒全部用下拉选择；日选项随年月自动调整。
  输出单位可选秒 / 毫秒 / 纳秒，结果框显示对应单位的整数时间戳。
- 修改任意时间选择或输出单位都会自动重算；点「转换」只是强制刷新一次，不再有“输入不完整”的中间状态。
  输出行的单位、转换按钮和结果按同一套列宽与上一行对齐。
- 底部详细结果列出识别类型、UTC / 本地时间（普通 + 高精度）、ISO、秒 / 毫秒 / 纳秒时间戳、
  星期和相对现在。纳秒超出 `long` 范围（约 1677-09-21 至 2262-04-11）时单独提示，秒/毫秒仍可用。
- 顶部「填入当前」把当前时间同时写进三行时间戳和年月日时间；「复制详细结果」复制底部报告。
- 当前时间计时器只在面板可见时运行；空输入和乱文本都按失败处理（不猜）。

### 3.3 编码

- 左右两栏：输入 → 输出（只读），各按钮独立转换；两栏中间的 ⇄ 按钮把两侧内容整体互换，方便链式操作。
- Unicode 转义默认只转非 ASCII；`Unicode 还原` 支持 `\uXXXX` / `\UXXXXXXXX` 和常见转义。
- Hex / Base64 / URL / HTML 的文本与字节统一按 UTF-8；解码失败报错，不返回替换字符。
- `MD5` / `SHA256` 按钮把十六进制摘要写进输出区（UTF-8 取字节），`大写输出` 勾选只影响这两个按钮；
  空文本也能算（UTF-8 0 字节），不做非空校验。原独立哈希页已并入本页。
- `文件 MD5` / `文件夹 MD5` 对本地文件流式计算（不整读进内存，GB 级文件也行），同样受 `大写输出` 影响：
  单个文件输出路径、MD5 和大小；文件夹逐文件算 MD5，按相对路径排序列出并汇总文件数/字节数，
  无权限的文件跳过并计为"读取失败"。文件夹计算在后台线程跑，期间再点会提示稍等。

### 3.4 HTTP

- 顶部工具条：方法下拉（GET/POST/PUT/DELETE）→ URL 输入框 → 次数 → 线程 → 发送 → 格式化响应 → 复制结果。
- **左右分栏**：左侧「已保存」列表 + 名称框（保存 / 删除），右侧请求详情；中间分隔条可拖动调宽。
- **已保存列表**：左侧输入名称后点「保存」把整套输入（方法 / URL / Body 类型 / 参数 / 请求头 /
  次数 / 线程）存进 `.wxd-work/http-requests.json`；同名覆盖，首次保存新建；点击列表项即自动加载，
  「删除」移除选中项。文件带临时文件原子写，损坏时先备份 `.bad` 再当空表（与 settings.json 同一套兜底）。
  旧档兼容：老档没有 BodyParams（POST 老档的 Params 就是请求体）、Content-Type 映射到 Body 单选，
  认不出的自定义值会放进请求头网格；更老的 JSON/Form 字段按老含义映射。
- **右侧一条页签栏共五页**（主题化 owner-draw）：参数 (Params) / 请求头 (Headers) / 响应 / 请求 / 响应头，
  详情三页紧跟请求头之后；发送不自动切页签。
  - **参数 (Params)**：唯一的键值网格，按方法分流——GET / DELETE 把参数拼到 URL（已有 ? 就用 & 续接）；
    POST / PUT 时网格上方出现 Body 单选行（none / application/json / x-www-form-urlencoded /
    multipart/form-data，默认 multipart），键值按单选类型组请求体，选 none 隐藏网格显示提示。
    切换方法时单选行自动显示 / 收起。
  - **请求头 (Headers)**：键值网格 + 编码下拉（UTF-8 / GBK / GB18030）。
  - **响应**：响应正文（原 `_respBody`，压测时是汇总文本）；「格式化响应」「复制结果」只作用于本页。
  - **请求**：请求行（方法 + 最终 URL，GET 已拼上参数）+ 发送的请求头网格（用户头 + Content-Type,
    multipart 会显示带 boundary 的真实值）+ 请求体（POST/PUT 的实际请求体文本，GET 为空）。
  - **响应头**：状态行（HTTP 码 / 耗时 / 字节 / 解码用编码）+ 响应头键值网格（响应头 + 内容头合并）；
    压测显示最后一次完成请求的头，全失败时提示无头可显示。
- 所有键值网格：空最后一行填完自动加新行，点行头选中整行后按 Del 删整行；只选单元格按 Del 是清空值。
  键为空的行发送时忽略。
- **请求头的键 / 值编辑时是可下拉的输入框**：点进单元格开始编辑就自动展开候选列表，输入时按前缀筛列表，
  自由输入不受限；键列候选是常用请求头（Accept / Authorization / Content-Type / User-Agent 等 17 个），
  值列候选按当前行的键给（Content-Type / User-Agent / Accept / Accept-Language / Accept-Encoding 有预设）。
  其它网格不提示。注意：可下拉网格建列必须先 `AllowUserToAddRows = false` → `ColumnCount` 建列 →
  换 `CellTemplate` → 再恢复 true；顺序反了空新行会按旧模板建成 TextBox，永远不出下拉。
- 编码下拉：请求体按所选编码出字节，form-urlencoded 按所选编码做 %XX 转义，响应没有 charset 时也按它兜底解码。
- 组装规则：POST / PUT 把参数网格的键值按单选类型组请求体，json 保持填写顺序、值统一按字符串，
  urlencoded 转义（同名键保留多份），multipart 用真实 boundary 组装（每个字段一个 text/plain part）；
  单选 none 但请求头里手写了 Content-Type 时，键值按原样 k=v&k=v 拼接不转义（游戏服手写解析的兜底通道）。
- gzip / deflate：`HttpClientHandler.AutomaticDecompression` 自动带 `Accept-Encoding: gzip, deflate`
  并透明解压响应体，无需手动设置请求头。
- 请求头原样透传（`TryAddWithoutValidation`，大小写不敏感）；`Content-Type` 特殊处理：它作用在请求体上
  （无请求体时不发送）；实际发送值以网格里手写的行优先（multipart 除外，boundary 以真实值为准）。
- **次数 > 1 进入压测**：按线程数并发发同一请求（次数按线程轮流认领，每请求独立 `HttpRequestMessage`），
  完成后响应区写汇总：总请求 / 线程 / 成功失败数、总耗时、QPS、平均 / 最快 / 最慢延迟，附首个错误；
  状态栏一行小结，全部失败才标红。次数 1–500、线程 1–64，越界和非数字静默夹紧。
- 单次请求响应按 charset 解码（GBK 等旧编码已支持；charset 没给或认不出按编码下拉的值兜底），
  二进制显示占位提示；响应头页签的状态行会标出实际解码编码名，状态码、耗时、字节数看状态栏，
  ≥400 显示红色。「格式化响应」把 JSON 缩进美化，非 JSON 会明确报错。
- 全程用自带 `HttpClient`（30 秒超时），没有引入任何第三方程序集；组装与保存序列化在 `HttpToolEngine`，可离线测试。

### 3.5 Cron

- 顶部下拉选择规范，解析规则完全不同，不能用同一套编号去套：
  - **标准 cron**：5 段（分 时 日 月 周），或沿用旧版 6 段（含秒）；
    `?` 在任意字段都当 `*`；日/周同时限定时按“或”；不支持 `L / W / #`。
  - **Quartz**：6 段（秒 分 时 日 月 周）或 7 段（末尾加年 1970-2099）；
    日/周必须有一列写 `?`；星期 `1-7 = 周日-周六`；支持 `L / L-n / LW / nW / dL / d#n`。
  - **Spring**：固定 6 段，无年字段；星期 `0/7 = 周日`；支持 `L / L-n / LW / nW / dL / d#n`
    和 `@yearly / @monthly / @weekly / @daily / @hourly` 等宏。`?` 等同 `*`，日/周同时限定时
    两个条件都要满足（“与”语义，与标准 cron 的“或”相反）。
- 5 段表达式在 Quartz/Spring 下会回退成标准 cron 并给出提示（不会直接报错）。
- 解析成功显示规范、宏展开、逐字段中文说明和接下来 10 次触发时间（本地时间）；
  Quartz 带年字段时推算上限跟着最大年份走，不会白扫 8 年。
- “语法参考”按钮把两套规范的字段表、特殊字符、常见坑和示例写进结果区。
- “示例”按钮按当前规范随机挑一个典型场景（工作日间隔、月末/第 n 个星期几、指定年份、
  每天固定时间、整点周期等），并把表达式解析回下方生成器；页签、模式、勾选项和上方输入框
  会一起同步，连续点击不会重复上一次的示例。
- 下方生成器按字段页签工作：指定（勾选网格，一个都不选等同 `*`）/ 周期（`a-b`）/ 间隔（`a/n`）；
  日、周另有“特殊”模式生成 `L / W / #`。Quartz 下改了一列日/周，生成器会自动把另一列
  退回 `?`；年份空选时省略第 7 段，生成常见的 6 段表达式。
  切到标准 cron 时，日/周的“特殊”模式会自动退回空选，避免生成标准不支持的 `L / W / #`。
  切换规范只重排生成器字段，不覆盖手工输入的表达式。

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
- 时间戳页的年月日/时间改动会自动重算；`FillDateSelectors` 和 `UpdateDayOptions` 会改选中项，
  必须先置 `_dateUpdating`，否则会触发嵌套换算。主题刷新只遍历内容区 `_content`，不要遍历整个面板，
  否则会把顶部工具条按钮和底部状态栏的字号/底色一并改掉。
- 页签是 owner-draw 的；页签标题（如 JSON 页的 `*`）要靠 `ToolboxForm` 的 `SetJsonCaption` 触发重绘。
- 用 harness 往 Win32 EDIT 里灌多行样例必须用 `\r\n`；只给 `\n` 会全部挤成一行，这不是面板 bug。

## 5. 已知取舍

- 原子写只保证"不会写一半"，没有版本备份；重要文件建议先提交 VCS。
- JSONC 注释在格式化/压缩/key 排序后会被移除（已弹确认）；目前没有保留注释的格式化实现。
- JSON 页每次文本变化都会重建整棵树；超大 JSON（几万节点）会卡，定位是小配置文件。
- JSON 搜索是线性遍历全部节点，没做索引；key 排序固定递归 + 序数排序，没有"仅当前层 / 忽略大小写"选项。
- YAML 转换走通用序列化，锚点/别名、多文档（`---`）不在支持范围；注释同样会丢。
- Cron 只给"接下来几次"的推算，不做完整时区/DST 处理：按本机时区展示，夏令时不存在的小时会被跳过，
  重复小时不去重；Quartz 带年字段时按最大年份截断。生成器不做表达式反解析，手工输入不会回填到生成器。
