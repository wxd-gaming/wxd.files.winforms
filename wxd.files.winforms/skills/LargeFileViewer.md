# LargeFileViewer 开发指南

> 大文件阅读器，支持 10GB+ 文本/日志文件秒级打开

---

## 1. 涉及文件

| 文件 | 职责 |
|------|------|
| `LargeFileViewer.cs` | 主窗体：文件列表 + 多标签页管理 |
| `LargeFileReaderControl.cs` | 核心控件：流式读取、虚拟滚动、搜索、索引、导出 |

---

## 2. 架构

```
LargeFileViewer (主窗体)
├── 左侧文件列表 (ListView)        — 展示目录内文本文件
├── 右侧标签页 (TabControl)        — 每个标签页嵌入 LargeFileReaderControl
│
LargeFileReaderControl (核心控件)
├── 工具栏                        — 打开文件、行号跳转、搜索
├── 虚拟滚动列表 (HScrollListView) — 高性能行显示
├── 行偏移索引                    — 后台异步构建
└── 缓存                          — ConcurrentDictionary<int, string>
```

---

## 3. 关键状态字段

| 字段 | 类型 | 说明 |
|------|------|------|
| `_filePath` | `string?` | 当前打开的文件路径 |
| `_fileLength` | `long` | 文件总大小 |
| `_lineOffsets` | `volatile long[]?` | 行偏移索引（每行的起始字节位置） |
| `_indexReady` | `volatile bool` | 索引是否构建完成 |
| `_readStream` | `FileStream?` | 文件读取流 |
| `_cache` | `ConcurrentDictionary<int, string>` | 行缓存 |
| `_searchResults` | `List<int>` | 搜索结果行号列表 |
| `_searchCurrentIndex` | `int` | 当前搜索导航位置 |

---

## 4. 流式读取机制

- **分块读取**：`CHUNK = 4MB`
- **最大行长度**：`MAX_LINE = 64KB`
- **不将整个文件加载到内存**，内存占用 ≤ 500MB
- 读取使用 `FileStream` + `FileShare.Read`（允许其他程序同时读取）

---

## 5. 行偏移索引

- 后台异步构建：`Task.Run` + `CancellationTokenSource`
- 进度显示：`_idxProg` / `_idxCount`
- 索引完成后才能使用搜索和行号跳转
- 索引状态：`_indexReady`

---

## 6. 虚拟滚动

- `HScrollListView`：自定义控件，支持水平滚动
- 只渲染可见区域的行
- 从缓存 `_cache` 读取行内容，缓存未命中时从 `_readStream` 读取

---

## 7. 搜索功能

- 关键词搜索、区分大小写、正则表达式
- 上下文行数可配置（`_numContextLines`）
- 搜索结果导航：第一条/上一条/下一条/最后一条
- 搜索使用 `CancellationTokenSource` 支持取消

---

## 8. 导出功能

- 将搜索结果/当前内容导出为文件
- 导出成功触发 `FileExported` 事件

---

## 9. 修改禁区

- **❌ 不要将整个文件加载到内存**（必须保持流式读取）
- **❌ 行偏移索引构建不能阻塞 UI 线程**
- **❌ `_readStream` 必须使用 `FileShare.Read`**（允许其他程序同时读取）
- **❌ 不要修改 `CHUNK` 和 `MAX_LINE` 常量**（影响性能和稳定性）
- **❌ 搜索时必须检查 `CancellationToken`**
