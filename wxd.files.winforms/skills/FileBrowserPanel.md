# FileBrowserPanel 开发指南

> 单个窗格文件夹管理器控件，是应用的核心组件

---

## 1. 涉及文件

| 文件 | 职责 |
|------|------|
| `FileBrowserPanel.cs` | 主控件（~95KB，~2500行） |
| `WindowsShellContextMenu.cs` | 系统右键菜单集成（Shell API） |
| `SimpleShellContextMenu.cs` | 备用右键菜单（简化实现） |
| `ImagePreviewToolTip.cs` | 图片预览 ToolTip |

---

## 2. 视图模式

```csharp
public enum FileListViewMode
{
    Details,    // 详细信息（默认）
    List,       // 列表
    SmallIcon,  // 小图标
    LargeIcon,  // 大图标
    Tile        // 平铺
}
```

---

## 3. 关键状态字段

| 字段 | 类型 | 说明 |
|------|------|------|
| `_currentDir` | `string` | 当前浏览目录 |
| `_filterText` | `string` | 当前过滤文本 |
| `_viewMode` | `FileListViewMode` | 视图模式 |
| `_allFiles` | `List<FileInfo>` | 当前目录文件缓存 |
| `_allDirs` | `List<DirInfo>` | 当前目录子目录缓存 |
| `_isDark` | `bool` | 当前主题（深色/浅色） |
| `_isLoading` | `bool` | 加载中标志 |
| `_hasUnsavedChanges` | `bool` | 是否有未保存的状态变更 |

---

## 4. 关键方法

| 方法 | 说明 |
|------|------|
| `NavigateTo(string path)` | 导航到目录（核心） |
| `ApplyFilter()` | 应用过滤条件 |
| `ApplyTheme(bool isDark)` | 主题切换 |
| `RefreshList()` | 刷新文件列表 |
| `SetThemeColors(bool isDark)` | 设置所有子控件颜色 |

---

## 5. 防抖机制

| 定时器 | 延迟 | 用途 |
|--------|------|------|
| `_filterDebounceTimer` | 300ms | 输入防抖，避免每次按键都触发过滤 |
| `_saveDebounceTimer` | 3s | 状态保存防抖，避免频繁写配置文件 |

---

## 6. 右键菜单

- **WindowsShellContextMenu**：通过 Shell API 调用原生上下文菜单（支持 SVN、Git 等扩展）
- **SimpleShellContextMenu**：备用方案（当 Shell API 失败时使用）
- 修改右键菜单逻辑需同时考虑两种实现

---

## 7. 自然排序

- `NaturalSortComparer`：数字部分按数值比较，其余按字符比较
- 用于文件/目录列表排序
- 修改排序逻辑时注意不要破坏数字的自然顺序

---

## 8. 路径下拉框 (_cboPath)

- 集成磁盘切换 + 最近访问路径（最多15条）
- 支持手动输入路径回车跳转
- 路径记忆通过 `CsvSettings` 持久化

---

## 9. 修改禁区

- **❌ 不要在 UI 线程同步 I/O**
- **❌ 不要绕过防抖直接刷新**（应使用 `_filterDebounceTimer`）
- **❌ 修改右键菜单逻辑需同时考虑 `WindowsShellContextMenu` 和 `SimpleShellContextMenu`**
- **❌ `NavigateTo()` 中的取消逻辑不能省略**（使用 `CancellationTokenSource`）
- **❌ 不要修改 `NaturalSortComparer` 的数字比较逻辑**

---

## 10. 配置持久化

- 当前目录：`CsvSettings.SetPaneDirectory(index, path)`
- 过滤条件：`CsvSettings.SetPaneFilter(index, filter)`
- 视图模式：`CsvSettings.SetPaneViewMode(index, mode)`
- 显示隐藏文件：`CsvSettings.ShowHiddenFiles`
