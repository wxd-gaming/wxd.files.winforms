using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wxd.files.winforms
{
    /// <summary>文件 VCS 状态枚举</summary>
    public enum FileVcsStatus
    {
        Normal,      // 正常（未修改）
        Modified,    // 已修改
        Added,       // 新增
        Deleted,     // 已删除
        Conflicted,  // 有冲突
        Untracked,   // 未跟踪
        Ignored,     // 已忽略
        Unknown      // 未知
    }

    /// <summary>单个目录的 VCS 状态快照</summary>
    public class VcsDirectoryStatus : EventArgs
    {
        /// <summary>目录路径</summary>
        public string DirectoryPath { get; set; } = "";
        /// <summary>VCS 类型</summary>
        public VcsType VcsType { get; set; } = VcsType.None;
        /// <summary>仓库根路径</summary>
        public string RepoRoot { get; set; } = "";
        /// <summary>文件名 → VCS 状态（key 是纯文件名，不含路径）</summary>
        public Dictionary<string, FileVcsStatus> FileStatuses { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>最后刷新时间</summary>
        public DateTime LastRefreshTime { get; set; }
        /// <summary>是否正在刷新</summary>
        public bool IsRefreshing { get; set; }
        /// <summary>刷新失败的错误信息（null 表示正常）</summary>
        public string? ErrorMessage { get; set; }
        /// <summary>缓存是否过时（超过 30 秒未刷新）</summary>
        public bool IsStale => !IsRefreshing && (DateTime.UtcNow - LastRefreshTime).TotalSeconds > 30;
    }

    /// <summary>
    /// 全局统一的 VCS 状态缓存管理器。
    /// 以目录路径为 key，同一目录下的 FileBrowserPanel 和 CsvReaderForm 共享缓存。
    /// 状态刷新完全异步，刷新完成后通过 StatusRefreshed 事件通知所有监听者。
    /// 使用 _refreshTasks 字典确保同一目录不会重复发起 git/svn 命令。
    ///
    /// 缓存策略（LRU）：
    /// - 最大缓存 100 个目录
    /// - 超出上限时淘汰最久未访问的缓存
    /// - Invalidate 标记过期不删除，切回时缓存命中直接返回旧数据，后台静默刷新
    /// </summary>
    public static class VcsStatusCache
    {
        static VcsStatusCache()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        // ==================== 常量 ====================
        private const int MaxCacheEntries = 100;

        // ==================== 存储 ====================
        // 目录路径 → 状态快照（不区分大小写）
        private static readonly ConcurrentDictionary<string, VcsDirectoryStatus> _cache = new(StringComparer.OrdinalIgnoreCase);

        // 访问顺序记录（用于 LRU 淘汰）
        private static readonly ConcurrentDictionary<string, DateTime> _accessOrder = new(StringComparer.OrdinalIgnoreCase);

        // 刷新任务锁：确保同一目录只有一个刷新任务在运行
        private static readonly ConcurrentDictionary<string, Task> _refreshTasks = new(StringComparer.OrdinalIgnoreCase);

        // ==================== 事件派发器 ====================
        // 使用线程安全的 List + AddListener/RemoveListener 替代 static event，
        // 避免多个实例间 -= 误删其他实例订阅的问题。
        private static readonly object _listenersLock = new();
        private static readonly List<EventHandler<VcsDirectoryStatus>> _listeners = new();

        /// <summary>添加状态刷新完成监听器</summary>
        public static void AddListener(EventHandler<VcsDirectoryStatus> listener)
        {
            lock (_listenersLock)
            {
                if (!_listeners.Contains(listener))
                    _listeners.Add(listener);
            }
        }

        /// <summary>移除状态刷新完成监听器</summary>
        public static void RemoveListener(EventHandler<VcsDirectoryStatus> listener)
        {
            lock (_listenersLock)
            {
                _listeners.Remove(listener);
            }
        }

        /// <summary>通知所有监听器状态刷新完成</summary>
        private static void NotifyStatusRefreshed(VcsDirectoryStatus status)
        {
            EventHandler<VcsDirectoryStatus>[]? snapshot;
            lock (_listenersLock)
            {
                if (_listeners.Count == 0) return;
                snapshot = _listeners.ToArray();
            }
            foreach (var listener in snapshot)
            {
                try { listener(null, status); }
                catch { /* 防止单个监听器异常影响其他监听器 */ }
            }
        }

        /// <summary>获取目录状态，缓存缺失或过期时自动在后台刷新（不会重复发起）</summary>
        public static VcsDirectoryStatus GetOrRefresh(string directoryPath)
        {
            directoryPath = directoryPath.ToLowerInvariant();
            if (_cache.TryGetValue(directoryPath, out var cached) && !cached.IsStale)
            {
                // 缓存命中且未过期，更新访问时间
                _accessOrder[directoryPath] = DateTime.UtcNow;
                return cached;
            }

            // 确保有且仅有一个后台刷新任务
            EnsureRefresh(directoryPath);

            // 记录访问时间
            _accessOrder[directoryPath] = DateTime.UtcNow;

            // 检查缓存上限，超出时淘汰最久未访问的
            TrimCache();

            return cached ?? new VcsDirectoryStatus { DirectoryPath = directoryPath, IsRefreshing = true };
        }

        /// <summary>强制异步刷新</summary>
        public static Task RefreshAsync(string directoryPath)
        {
            return EnsureRefresh(directoryPath);
        }

        /// <summary>确保指定目录的刷新任务正在执行（不会重复创建）</summary>
        private static Task EnsureRefresh(string directoryPath)
        {
            // 如果已有刷新任务在运行
            if (_refreshTasks.TryGetValue(directoryPath, out var existing))
            {
                // 如果已有 task 已完成（但 ContinueWith 还没从字典中移除），
                // 移除并重新创建，避免新订阅者错过 StatusRefreshed 事件通知
                if (existing.IsCompleted)
                {
                    _refreshTasks.TryRemove(directoryPath, out _);
                }
                else
                {
                    return existing; // 还在执行中，复用
                }
            }

            // 创建新任务
            var task = RefreshInternalAsync(directoryPath);

            // 尝试注册到字典，如果已被其他线程先注册了，丢弃当前任务
            if (_refreshTasks.TryAdd(directoryPath, task))
            {
                // 任务完成后清理字典
                _ = task.ContinueWith(_ =>
                {
                    _refreshTasks.TryRemove(directoryPath, out _);
                }, TaskContinuationOptions.ExecuteSynchronously);
            }

            return task;
        }

        private static async Task RefreshInternalAsync(string directoryPath)
        {
            directoryPath = directoryPath.ToLowerInvariant();
            var status = _cache.GetOrAdd(directoryPath, _ => new VcsDirectoryStatus { DirectoryPath = directoryPath });

            status.IsRefreshing = true;
            status.ErrorMessage = null;

            try
            {
                var vcsType = VersionControlHelper.DetectVcs(directoryPath);
                if (vcsType == VcsType.None)
                {
                    status.VcsType = VcsType.None;
                    status.RepoRoot = "";
                    status.FileStatuses = new Dictionary<string, FileVcsStatus>(StringComparer.OrdinalIgnoreCase);
                    status.LastRefreshTime = DateTime.UtcNow;
                    status.IsRefreshing = false;
                    NotifyStatusRefreshed(status);
                    return;
                }

                var repoRoot = VersionControlHelper.GetVcsRoot(directoryPath, vcsType);
                status.VcsType = vcsType;
                status.RepoRoot = repoRoot;

                // 在用户选择的目录下执行 git/svn status，只扫描该目录的直接子文件
                var fileStatuses = vcsType == VcsType.Git
                    ? await GetGitFileStatusesAsync(directoryPath)
                    : await GetSvnFileStatusesAsync(directoryPath);

                status.FileStatuses = fileStatuses;
                status.LastRefreshTime = DateTime.UtcNow;
                status.IsRefreshing = false;
                status.ErrorMessage = null;
            }
            catch (Exception ex)
            {
                status.IsRefreshing = false;
                status.ErrorMessage = $"刷新 VCS 状态失败: {ex.Message}";
                status.LastRefreshTime = DateTime.UtcNow;
            }

            NotifyStatusRefreshed(status);
        }

        /// <summary>
        /// 检查指定目录及其子目录的缓存中是否有冲突文件。
        /// 只搜索 rootDirectory 范围内的缓存，不会遍历整个仓库。
        /// </summary>
        public static bool HasAnyConflictedFiles(string rootDirectory)
        {
            var root = rootDirectory.TrimEnd('\\').ToLowerInvariant();
            foreach (var kvp in _cache)
            {
                // 只检查属于 rootDirectory 范围内的缓存
                if (!kvp.Key.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (var fs in kvp.Value.FileStatuses.Values)
                {
                    if (fs == FileVcsStatus.Conflicted)
                        return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 查询单个文件的 VCS 状态。
        /// 优先精确匹配文件所在目录的缓存；
        /// 如果未找到，向上遍历父目录查找（处理子目录中冲突文件等场景）。
        /// </summary>
        public static FileVcsStatus GetFileStatus(string filePath)
        {
            var dir = Path.GetDirectoryName(filePath) ?? "";
            var fileName = Path.GetFileName(filePath);

            // 从文件所在目录开始，逐级向上查找缓存
            var searchDir = dir;
            while (!string.IsNullOrEmpty(searchDir))
            {
                if (_cache.TryGetValue(searchDir.ToLowerInvariant(), out var cached))
                {
                    if (cached.FileStatuses.TryGetValue(fileName, out var fs))
                        return fs;
                    // 找到了缓存但没找到文件 → 停止向上遍历
                    return FileVcsStatus.Normal;
                }

                // 向上找父目录
                var parent = Path.GetDirectoryName(searchDir);
                if (string.IsNullOrEmpty(parent) ||
                    string.Equals(parent, searchDir, StringComparison.OrdinalIgnoreCase))
                    break;
                searchDir = parent;
            }

            return FileVcsStatus.Normal;
        }

        /// <summary>
        /// 标记指定目录的缓存为过期（不删除缓存）。
        /// 后续 GetOrRefresh 时发现过期会自动后台刷新，
        /// 但会立即返回旧数据，避免等待 git/svn 命令。
        /// </summary>
        public static void Invalidate(string directoryPath)
        {
            if (_cache.TryGetValue(directoryPath, out var cached))
            {
                cached.LastRefreshTime = DateTime.MinValue; // 标记为过期
            }
        }

        /// <summary>获取缓存数量（用于调试）</summary>
        public static int CacheCount => _cache.Count;

        // ==================== LRU 淘汰 ====================

        /// <summary>超出上限时淘汰最久未访问的缓存</summary>
        private static void TrimCache()
        {
            if (_cache.Count <= MaxCacheEntries) return;

            var toRemove = _accessOrder
                .OrderBy(kvp => kvp.Value)
                .Take(_cache.Count - MaxCacheEntries)
                .Select(kvp => kvp.Key)
                .ToList();

            foreach (var key in toRemove)
            {
                _cache.TryRemove(key, out _);
                _accessOrder.TryRemove(key, out _);
            }
        }

        // ==================== 显示辅助方法 ====================

        /// <summary>获取状态文本前缀（用于 TreeView 节点文本）</summary>
        public static string GetStatusPrefix(FileVcsStatus status)
        {
            return status switch
            {
                FileVcsStatus.Modified => "✏️ ",
                FileVcsStatus.Added => "➕ ",
                FileVcsStatus.Deleted => "➖ ",
                FileVcsStatus.Conflicted => "⚠️ ",
                FileVcsStatus.Untracked => "🆕 ",
                _ => ""
            };
        }

        /// <summary>获取状态对应颜色（用于 ListView 前景色）</summary>
        public static Color GetStatusColor(FileVcsStatus status)
        {
            return status switch
            {
                FileVcsStatus.Modified => Color.FromArgb(66, 133, 244),    // 蓝色
                FileVcsStatus.Added => Color.FromArgb(52, 168, 83),       // 绿色
                FileVcsStatus.Deleted => Color.FromArgb(197, 197, 197),   // 灰色
                FileVcsStatus.Conflicted => Color.FromArgb(234, 67, 53),  // 红色
                FileVcsStatus.Untracked => Color.FromArgb(154, 160, 166), // 浅灰
                _ => Color.Black
            };
        }

        // ==================== Git 状态获取 ====================

        /// <summary>
        /// 获取 Git 文件状态。
        /// 在用户目录下执行 git status，只扫描当前目录的直接子文件（不递归子文件夹）。
        /// </summary>
        private static async Task<Dictionary<string, FileVcsStatus>> GetGitFileStatusesAsync(string workingDirectory)
        {
            var result = new Dictionary<string, FileVcsStatus>(StringComparer.OrdinalIgnoreCase);

            // core.quotepath=false：否则中文文件名会被输出成 "\344\270\255..."，永远匹配不上
            var run = await VersionControlHelper.RunProcessAsync(
                "git",
                "-c core.quotepath=false status --porcelain -u .",
                workingDirectory: workingDirectory,
                timeoutMs: 30_000,
                disableInteractivePrompts: true).ConfigureAwait(false);
            if (!run.Started || run.ExitCode != 0) return result;

            var output = run.Output;

            foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                if (line.Length < 3) continue;

                var xy = line[..2];
                var path = line[3..].TrimEnd('\r');

                // 重命名/复制：porcelain 输出为 "R  old -> new"
                if ((xy[0] == 'R' || xy[0] == 'C'))
                {
                    int arrow = path.LastIndexOf(" -> ", StringComparison.Ordinal);
                    if (arrow >= 0) path = path[(arrow + 4)..];
                }

                // ★ 只处理当前目录的直接文件（不包含子目录路径）
                var fileName = Path.GetFileName(path.Replace('/', '\\'));
                if (string.IsNullOrEmpty(fileName) || fileName != path.Replace('/', '\\'))
                    continue;

                var workingStatus = xy[1];
                var stagedStatus = xy[0];

                FileVcsStatus status;
                if (xy == "??")
                    status = FileVcsStatus.Untracked;
                else if (xy == "!!")
                    status = FileVcsStatus.Ignored;
                else if (workingStatus == 'U' || stagedStatus == 'U' || xy == "DD" || xy == "AA" || xy == "UU")
                    status = FileVcsStatus.Conflicted;
                else if (workingStatus == 'M' || stagedStatus == 'M')
                    status = FileVcsStatus.Modified;
                else if (workingStatus == 'A' || stagedStatus == 'A')
                    status = FileVcsStatus.Added;
                else if (workingStatus == 'D' || stagedStatus == 'D')
                    status = FileVcsStatus.Deleted;
                else if (stagedStatus == 'R' || stagedStatus == 'C')
                    status = FileVcsStatus.Added;   // 重命名/复制：新路径视为新增
                else
                    status = FileVcsStatus.Unknown;

                result[fileName] = status;
            }

            return result;
        }

        // ==================== SVN 状态获取 ====================

        /// <summary>
        /// 获取 SVN 文件状态。
        /// 使用 --depth=immediates 只扫描当前目录的直接子文件，不递归子文件夹。
        /// </summary>
        private static async Task<Dictionary<string, FileVcsStatus>> GetSvnFileStatusesAsync(string directoryPath)
        {
            var result = new Dictionary<string, FileVcsStatus>(StringComparer.OrdinalIgnoreCase);

            var svnExe = FindSvnExe() ?? "svn";

            // --depth=immediates 只扫描当前目录的直接子文件和子文件夹，不递归
            // ★ 不设置 WorkingDirectory：绝对路径路径参数已足够，避免中文工作目录的问题
            var run = await VersionControlHelper.RunProcessAsync(
                svnExe,
                $"status --depth=immediates --non-interactive \"{directoryPath}\"",
                outputEncoding: Encoding.GetEncoding(936),
                timeoutMs: 30_000,
                disableInteractivePrompts: true).ConfigureAwait(false);
            if (!run.Started || run.ExitCode != 0) return result;

            var output = run.Output;

            foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                if (line.Length < 8) continue;

                var statusChar = line[0];
                var path = line[8..].Trim();

                // ★ 跳过子目录条目（SVN 中目录状态行以 / 或 \ 结尾）
                if (path.EndsWith("\\") || path.EndsWith("/")) continue;

                var fileName = Path.GetFileName(path);
                if (string.IsNullOrEmpty(fileName)) continue;

                // ★ 构建完整路径：支持子目录中的文件（如 In\cfg_global.xlsx）
                string fullPath;
                if (Path.IsPathRooted(path))
                {
                    fullPath = path; // 已经是绝对路径
                }
                else
                {
                    fullPath = Path.Combine(directoryPath, path); // 相对路径
                }

                // ★ 不检查 File.Exists — 信任 SVN 的输出（冲突文件可能已被删除但仍需标记）
                var fileDir = Path.GetDirectoryName(fullPath) ?? "";
                if (string.IsNullOrEmpty(fileDir)) continue;

                var status = statusChar switch
                {
                    'M' => FileVcsStatus.Modified,
                    'A' => FileVcsStatus.Added,
                    'D' => FileVcsStatus.Deleted,
                    'C' => FileVcsStatus.Conflicted,
                    '?' => FileVcsStatus.Untracked,
                    'I' => FileVcsStatus.Ignored,
                    _ => FileVcsStatus.Unknown
                };

                // ★ 如果文件在子目录中，为该子目录创建/更新缓存条目
                if (!string.Equals(fileDir, directoryPath, StringComparison.OrdinalIgnoreCase))
                {
                    var subDir = fileDir.ToLowerInvariant();
                    var subCache = _cache.GetOrAdd(subDir, _ => new VcsDirectoryStatus
                    {
                        DirectoryPath = fileDir,
                        VcsType = VcsType.Svn,
                        LastRefreshTime = DateTime.UtcNow
                    });
                    subCache.FileStatuses[fileName] = status;
                }
                else
                {
                    // 文件在当前目录下
                    result[fileName] = status;
                }
            }

            return result;
        }

        /// <summary>查找 svn.exe（从常见安装路径）</summary>
        private static string? FindSvnExe()
        {
            var paths = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "TortoiseSVN", "bin", "svn.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "TortoiseSVN", "bin", "svn.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "SlikSvn", "bin", "svn.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "SlikSvn", "bin", "svn.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "VisualSVN", "bin", "svn.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "VisualSVN", "bin", "svn.exe"),
            };

            foreach (var path in paths)
            {
                if (File.Exists(path))
                    return path;
            }

            return "svn";
        }
    }
}
