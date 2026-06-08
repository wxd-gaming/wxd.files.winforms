# CustomCommand 开发指南

> 自定义命令系统，支持可配置的命令按钮

---

## 1. 涉及文件

| 文件 | 职责 |
|------|------|
| `CustomCommandItem.cs` | 命令数据模型 |
| `CustomCommandManagerForm.cs` | 命令管理面板（DataGridView 编辑） |
| `Form1.cs` | 主窗体中创建命令按钮（`_flowPanel`） |

---

## 2. 数据模型 (CustomCommandItem)

| 字段 | 类型 | 说明 |
|------|------|------|
| `Name` | `string` | 命令显示名称（按钮上的文字） |
| `Command` | `string` | 要执行的命令（exe路径 或 `cmd /c xxx`） |
| `WorkDir` | `string` | 命令执行的工作目录（留空则使用默认目录） |
| `ShowWindow` | `bool` | 是否显示 cmd 窗口（false=后台静默执行） |
| `ButtonColor` | `string` | 按钮颜色（ARGB 十六进制字符串，如 `"#7B1FA2"`） |

---

## 3. 管理面板 (CustomCommandManagerForm)

- 使用 `BindingList<CustomCommandItem>` + `DataGridView` 编辑
- 列：命令名、执行的命令、工作目录、显示窗口、按钮颜色
- 按钮：添加、删除、复制、上移、下移、确定、取消
- 工作目录列支持双击选择文件夹

---

## 4. 主窗体按钮生成

- 在 `Form1.InitControlPanel()` 中创建按钮
- 按钮样式：`FlatStyle.Flat`、`BackColor`（从 `ButtonColor` 解析）、`ForeColor`（白色）
- 点击执行：`Process.Start()` + 窗口控制（`CreateNoWindow`/`UseShellExecute`）

---

## 5. 配置持久化

- 命令列表通过 `CsvSettings` 存储到 `dll.config`
- 使用 JSON 序列化存储命令列表

---

## 6. 修改禁区

- **❌ 不要改变 `CustomCommandItem` 的字段名**（会影响 JSON 序列化/反序列化）
- **❌ 不要修改按钮颜色解析逻辑**（`ButtonColor` 为 ARGB 十六进制字符串）
- **❌ 命令执行时必须处理异常**（防止无效命令导致崩溃）
