# ConfigPersistence 开发指南

> 配置持久化系统，双层存储架构

---

## 1. 涉及文件

| 文件 | 职责 |
|------|------|
| `CsvSettings.cs` | 全局/目录级配置读写（~23KB） |
| `DirectoryFrozenConfig.cs` | 目录级配置数据模型 |
| `App.config` / `wxd.files.winforms.dll.config` | 全局配置文件 |
| `{目录}/.wxd-work/config.json` | 目录级配置文件 |

---

## 2. 双层存储架构

```
全局配置 (dll.config)
├── ThemeMode               — 主题模式
├── PaneCount               — 窗格数量
├── LayoutIndex             — 布局索引
├── Pane{N}Directory        — 各窗格目录
├── Pane{N}Filter           — 各窗格过滤条件
├── Pane{N}ViewMode         — 各窗格视图模式
├── ShowHiddenFiles         — 显示隐藏文件
├── RecentPaths             — 最近访问路径
├── CustomCommands          — 自定义命令
├── CsvReaderLastDir        — CSV阅读器上次目录
└── LargeFileViewerLastDir  — 大文件阅读器上次目录

目录级配置 (.wxd-work/config.json)
├── CsvFrozenRows           — CSV冻结行数
├── ExcelFrozenRows         — Excel冻结行数
├── CsvHeaderRowIndex       — CSV表头行索引
├── ExcelHeaderRowIndex     — Excel表头行索引
├── CsvMappingRowIndex      — CSV映射行索引
├── ExcelMappingRowIndex    — Excel映射行索引
├── CsvUniqueColumnIndex    — CSV唯一列索引
├── ExcelUniqueColumnIndex  — Excel唯一列索引
├── ColumnLinks             — 列关联配置列表
├── FileRemarks             — 文件备注映射
├── ColumnVisibility        — 列显隐配置
├── LinkTemplates           — 关联配置模板列表
└── FileSettings            — 文件级配置（备注/唯一列/隐藏列/冻结列/列关联，key=相对路径）
```

---

## 3. CsvSettings API 分类

### 3.1 全局配置（读/写 dll.config）

- 使用 `ConfigurationManager.AppSettings` 读写
- 基础方法：`GetIntValue(key, default, min, max)` / `SetIntValue(key, value)`
- 基础方法：`GetStringValue(key, default)` / `SetStringValue(key, value)`
- 保存：`CsvSettings.Save()` → `config.Save()`

### 3.2 目录级配置（读/写 .wxd-work/config.json）

- 读取：`ReadDirConfigFromFile(directoryPath)` → JSON 反序列化
- 写入：`WriteDirConfigToFile(directoryPath, cfg)` → JSON 序列化
- 自动创建 `.wxd-work` 隐藏文件夹
- LRU 淘汰：最多缓存 50 个目录配置（`MaxDirConfigs = 50`）

### 3.3 关键 API

| API | 说明 |
|-----|------|
| `GetCsvFrozenRowsByDirectory(dir)` | 获取目录CSV冻结行数 |
| `SetCsvFrozenRowsByDirectory(dir, rows)` | 设置目录CSV冻结行数 |
| `GetColumnLinksByDirectory(dir)` | 获取目录列关联配置 |
| `SetColumnLinksByDirectory(dir, links)` | 设置目录列关联配置 |
| `GetFileRemarksByDirectory(dir)` | 获取目录文件备注映射 |
| `SetFileRemark(dir, relPath, remark)` | 设置单个文件备注 |
| `GetColumnVisibility(dir, relPath)` | 获取文件列显隐状态 |
| `SetColumnVisibility(dir, relPath, states)` | 设置文件列显隐状态 |
| `GetFileFrozenColumnCount(dir, relPath)` | 获取文件冻结列数（显示列口径，0=不冻结） |
| `SetFileFrozenColumnCount(dir, relPath, count)` | 设置文件冻结列数（`<=0` 存 null，值为负时钳到 0） |

---

## 4. DirectoryFrozenConfig 数据模型

| 字段 | 类型 | 说明 |
|------|------|------|
| `DirectoryPath` | `string` | 目录路径（运行时字段，不序列化到文件） |
| `CsvFrozenRows` | `int` | CSV冻结行数（默认2） |
| `ExcelFrozenRows` | `int` | Excel冻结行数（默认2） |
| `CsvHeaderRowIndex` | `int` | CSV表头行索引（-1=默认第一行） |
| `ExcelHeaderRowIndex` | `int` | Excel表头行索引（-1=默认第一行） |
| `LastModified` | `DateTime` | 最后修改时间 |
| `LastAccessed` | `DateTime` | 最后访问时间（用于LRU淘汰） |
| `ColumnLinks` | `List<ColumnLinkConfig>` | 列关联配置列表 |
| `FileRemarks` | `Dictionary<string, string>` | 文件备注映射（已废弃，迁移到 `FileSettings.Remark`） |
| `ColumnVisibility` | `Dictionary<string, string>` | 列显隐配置（已废弃，迁移到 `FileSettings.HiddenColumns`） |
| `FileSettings` | `Dictionary<string, FileConfig>` | 文件级配置，key=文件相对路径 |

**FileConfig（FileSettings 的值）**

| 字段 | 类型 | 说明 |
|------|------|------|
| `Remark` | `string?` | 备注名（null=无备注） |
| `UniqueColumnIndex` | `int?` | 唯一列索引（null=用目录级默认） |
| `HiddenColumns` | `HashSet<int>?` | 被取消勾选的列（null/空=全部显示） |
| `FrozenColumnCount` | `int?` | 冻结列数（**显示列口径**，冻结最左侧 N 列；null/0=不冻结） |
| `ColumnLinks` | `List<ColumnLinkConfig>?` | 该文件的列关联配置 |

---

## 5. .wxd-work 目录管理

- 自动创建：首次写入配置时
- 自动隐藏：`FileAttributes.Hidden`
- `.gitignore` 建议添加 `.wxd-work/`（由用户决定）
- 序列化前清除运行时字段（`DirectoryPath`），保持文件干净

---

## 6. 修改禁区

- **❌ 不要改变 `DirectoryFrozenConfig` 的 JSON 序列化字段名**（会导致旧配置文件读取失败）
- **❌ 不要修改 `ColumnVisibility` 的存储格式**（逗号分隔的bool字符串）
- **❌ 不要修改 `.wxd-work` 目录名和 `config.json` 文件名**
- **❌ `ReadDirConfigFromFile` 中的异常必须静默处理**（配置损坏不应影响主流程）
- **❌ 序列化时必须清除 `DirectoryPath` 运行时字段**（避免绝对路径写入文件）
