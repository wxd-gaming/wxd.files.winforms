# ThemeSystem 开发指南

> 主题系统：两套皮肤（Sky / Dense）× 深色/浅色，全部取色集中在令牌层

---

## 1. 配色方案

令牌唯一来源是 `Theme.Build(UiSkin skin, bool dark)`（`Theme.cs`）。界面代码一律通过 `Theme.P.令牌名` 取色，
不要再把下面的数字抄进任何窗体——需要新色就加令牌。

| 令牌 | Sky 浅色 | Sky 深色 | Dense 浅色 | Dense 深色 |
|------|----------|----------|------------|------------|
| Window 窗体底 | #F1F7FC | #161A20 | #F2F3F5 | #1D1F22 |
| Chrome 工具条底 | #E7F2FA | #1C222A | #E9EBEF | #23262A |
| Surface 内容底 | #FFFFFF | #1F2630 | #FFFFFF | #212428 |
| Border / Divider | #C3DCEF / #DCEBF7 | #2E3947 / #263041 | #D2D6DC / #E4E7EB | #35393F / #2B2F34 |
| Text 主文字 | #143C64 | #E3EEF9 | #1B1D21 | #DCE0E5 |
| Muted 次要文字 | #5A82AA | #92A6BC | #5F666E | #8B929B |
| Accent 主操作底 | #2E8BE0 | #4E9BE8 | #3A6EA5 | #5B8DD0 |
| Selection 选中底 | #D8ECFB | #23384F | #D9E6F5 | #2C3F5A |
| RowHeight / ButtonHeight | 24 / 26 | 24 / 26 | 22 / 24 | 22 / 24 |
| Radius | 4 | 4 | 3 | 3 |

派生色一律走函数，不要另写常量：`Theme.Mix(a,b,t)` 混合、`Theme.IsDark(c)` 判明暗、
`Theme.CellTint(pastel)` 深色下压住表格底色、`Theme.Signal(c)` 深色下提亮状态色。

## 2. 主题切换机制

- 皮肤是 `UiSkin` 枚举（`Sky`=0 / `Dense`=1），深浅色是 bool，两者独立。
- 存档：`CsvSettings.ThemeSkin`（0/1）、`CsvSettings.ThemeMode`（0=深色，1=浅色）。
- 唯一入口 `Theme.Set(skin, dark)`，顺序是：重建 `Theme.P` 和字体 → `ApplyNativeMode`（原生控件明暗）
  → `ApplyCaptions`（DWM 刷已开窗口的标题栏）→ `Broadcast()` 通知订阅方。
- `Broadcast()` 逐个调用 `Changed` 的订阅方并单独兜异常：某个窗体的 `ApplyTheme` 抛错不会连累其它窗口，
  也不会把异常抛回切肤按钮。
- 字号走 `Theme.Body / List / Head / Mono / Caption / Title / TitleBold`（切肤时整体重建，旧字体不回收，
  因为可能仍被现存控件引用）；间距走 `ThemeMetrics`（Margin=10、Gap=6、Group=14、RowPad=8、RowH=ButtonHeight）。
- 订阅 `Theme.Changed` 的控件必须在 `Disposed` 退订，静态事件会攥住已关闭的窗体。

## 3. 各组件 ApplyTheme 方法

| 组件 | 方法 | 说明 |
|------|------|------|
| `Form1` | `ApplyTheme()` | 主窗体 + 顶栏 + 所有面板，皮肤选择的发起方 |
| `FileBrowserPanel` | `ApplyTheme(bool isDark)` | 多窗格文件区；isDark 参数在令牌化后已冗余，保留是因为外部调用点多 |
| `TabbedFileBrowserPanel` | `ApplyTheme(bool isDark)` | 标签式窗格容器，转调子面板 |
| `CsvReaderForm` | `ApplyTheme()` | 文件树 + 标签页 + 搜索栏 |
| `TableBrowserControl` | `ApplyTheme()` | 表格网格底色、列头、斑马线、冻结列 |
| `LineListView` | `ApplyTheme()` | 大文本列表：底色/文字/选中色 + 按皮肤调行距 |
| `LargeFileReaderControl` / `LargeFileViewer` | `ApplyTheme()` | 日志阅读器与外层窗口 |
| 各对话框 | `ApplyTheme()` | 订阅 `Theme.Changed`，运行期切肤即时跟随 |

---

## 4. 新增控件时的主题规则

1. 取色一律 `Theme.P.令牌`，判明暗一律 `Theme.Dark`；不要再往控件里加 `_isDark` 之类的第二套状态。
2. 自绘控件直接读 `Theme.P`（绘制时实时取），这样不必订阅静态事件，也避免静态事件攥住实例；
   容器/窗体才订阅 `Theme.Changed`，并在 `Disposed` 退订。
3. 尺寸不写死：行高用 `ThemeMetrics.RowH`，按钮高度用 `Theme.P.ButtonHeight`，
   间距用 `Margin/Gap/Group/RowPad`；按钮宽度用 `ThemeButton.FitWidth(text)` 按文字量算。
4. 圆角用 `Theme.P.Radius`，且必须走 `ThemeButton.RoundRect`（它会把半径钳到宽高一半，别自己拼 AddArc）。
5. 新增界面先确认它接了 `Theme.Changed`，否则运行期切肤只有标题栏跟着变、客户区留在旧配色。
6. 控件是在构造时把颜色写进 `BackColor`/`ForeColor` 的（局部变量、一次性赋值，`ApplyTheme` 里没有对应字段可刷），
   就在它的 `ApplyTheme` 末尾调一次 `ThemeRecolor.Apply(...)`：它拿 `Theme.Previous` 和新的 `Theme.P` 建旧→新映射，
   递归把正好等于旧令牌的配色平移过去，只认精确相等，写死的杂色不会被误改，重复调用也安全。

---

## 5. 修改禁区

- **❌ 不要在界面代码里写 `Color.FromArgb(数字)`**——需要新色就加令牌，并一次配齐 4 套（两皮肤 × 深浅色）。
- **❌ 不要只配深色或只配浅色**：任何按明暗分支的取色都要两条分支都写，否则另一套会露出默认色。
- **❌ 主题切换时不要遗漏任何子控件**：`ApplyTheme` 里漏一个控件，现象是切肤后那块区域保持旧配色，
  重启才正常，很容易被当成"缓存问题"。
- **❌ 不要为了切肤去重建 ListView 句柄**：原生表头颜色只有重建才刷新，但代价是丢选中/滚动位置、
  大目录要重灌条目，动之前先确认。

---

## 6. 两个「设了 ForeColor 也不生效」的坑

这两处都是深色下文字看不清的真因，改主题时容易漏：

- **`ToolStripStatusLabel.ForeColor` 没人听**：系统渲染器画状态栏标签时走自己的取色，赋 `Theme.P.Muted` 会被忽略，深色下画成黑底黑字（只剩 ClearType 边缘那点蓝像素，肉眼看着像"很暗的蓝"）。项目里唯一用 `StatusStrip` 的地方是 `NoteForm`，已换成 `Panel + Label`（`Dock=Bottom, Height=Theme.P.RowHeight`），和其他窗口一致。**新窗口直接用 Panel+Label 做状态栏，别再引 StatusStrip。**
- **RTF 把写它当时的主题色钉在文档里**：`RichTextBox.Rtf = xxx` 会覆盖控件的 `ForeColor`，因为 RTF 自带 `\colortbl`。笔记正文（`NoteStore` 存的就是 RTF 字符串）在浅色下写、切到深色打开，正文就是近黑的 30,50,80。`NoteForm.RecolorContentText()` 在载入后和 `ApplyTheme()` 里各刷一遍 `SelectionColor = Theme.P.Text`。**任何把 RTF 当存储格式的控件都要照这个做两遍**；编辑器没有用户改色功能时才能这么无条件覆盖。

> 验证只能靠真窗口像素：`tools/gridshot` 起 `dlg note`，`tools/restyle-shots/pixprobe.ps1 -In 截图 -Region x,y,w,h` 会打印该区域的底色和最出挑的文字色，直接和 `Theme.cs` 里的 token 对得上才算过。`GRIDSHOT_FLIP=1` 会让笔记窗体开 2.5 秒后运行时切一次深浅色，用来验「已经打开的窗口」这条路径。

---

## 7. 标题栏（非客户区）明暗

标题栏不是 WinForms 画的，切主题时必须单独处理，两条 API 各有分工：

- `Application.SetColorMode` 只管**新建**的窗口。当前 net10 运行库的 `SystemColorMode` 只有 `Classic / System / Dark`，**没有 `Light`**（`Theme.ApplyNativeMode` 里那句名字兜底就是为这个），所以浅色只能落 `Classic`；`System` 会跟着操作系统走，不能用。
- 已经开着的窗口翻明暗时 `SetColorMode` 不会回头改它，必须自己扫：`Theme.Set` 里 `ApplyCaptions()` 遍历 `Application.OpenForms`，对每个句柄调 `DwmSetWindowAttribute(DWMWA_USE_IMMERSIVE_DARK_MODE)`（值 20，Win10 1809~1903 退回 19）。少了这一步的现象是：深色下开过主窗后切浅色，标题栏一直黑着。

验的时候别只看客户区，`main switch` 会在同实例里真翻一次主题，截图量标题栏那条（浅色 243,243,243 / 深色 32,32,32）。`gridshot … colormode` 会把运行库实际给了哪些明暗 API 写到 `colormode.txt`，升级运行库后先跑一遍。


