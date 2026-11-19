using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace wxd.files.winforms.RegressionTests;

/// <summary>文件操作队列 / 传输引擎 / 冲突策略的回归测试。</summary>
internal static partial class Program
{
    private static void RunFeatureTests()
    {
        Console.WriteLine("--- 文件操作：传输引擎与队列 ---");
        Run("复制文件保留内容与时间戳", Test_CopyFilePreservesMetadata);
        Run("复制目录递归", Test_CopyDirectoryRecursive);
        Run("冲突 跳过", Test_ConflictSkip);
        Run("冲突 覆盖", Test_ConflictOverwrite);
        Run("冲突 重命名", Test_ConflictRename);
        Run("冲突 保留较新者", Test_ConflictKeepNewer);
        Run("冲突 询问并全部应用", Test_ConflictAskApplyAll);
        Run("移动文件与目录", Test_MoveFileAndDirectory);
        Run("取消后清理临时文件", Test_CancelCleansTemp);
        Run("队列 顺序执行与完成状态", Test_QueueSequential);
        Run("队列 失败收集", Test_QueueFailure);
        Run("队列 失败事件", Test_QueueFailedEvent);
        Run("暂停闸门", Test_PauseToken);
        Run("唯一命名与长路径前缀", Test_UniqueAndLongPath);
        Run("统计文件数与字节", Test_Measure);

        Console.WriteLine("--- 文件搜索 ---");
        Run("搜索 文件名包含匹配(忽略大小写)", Test_SearchNameContains);
        Run("搜索 通配符与正则", Test_SearchWildcardAndRegex);
        Run("搜索 是否递归子目录", Test_SearchRecursiveToggle);
        Run("搜索 隐藏文件开关", Test_SearchHiddenToggle);
        Run("搜索 内容命中与行号", Test_SearchContentLineNumbers);
        Run("搜索 内容大小写与正则", Test_SearchContentCaseAndRegex);
        Run("搜索 大小与修改时间过滤", Test_SearchSizeAndDate);
        Run("搜索 结果上限截断", Test_SearchMaxResults);
        Run("搜索 取消与非法正则", Test_SearchCancelAndBadRegex);
        Run("搜索 二进制文件跳过", Test_SearchBinarySkip);
        Run("搜索 目录名命中", Test_SearchDirectoryHit);

        Console.WriteLine("--- 批量重命名 ---");
        Run("重命名 普通替换", Test_RenamePlainReplace);
        Run("重命名 正则捕获组", Test_RenameRegexCapture);
        Run("重命名 模板序号/扩展名", Test_RenameTemplateCounterAndExt);
        Run("重命名 仅大小写变化", Test_RenameCaseOnly);
        Run("重命名 非法名与保留名", Test_RenameInvalidAndReserved);
        Run("重命名 重名与已存在冲突", Test_RenameDuplicateAndConflict);
        Run("重命名 存在错误时整体拒绝", Test_RenameRefusesOnError);
        Run("重命名 执行与撤销", Test_RenameApplyAndUndo);
        Run("重命名 扩展名参与替换", Test_RenameIncludeExtension);

        Console.WriteLine("--- 新建文件/文件夹 ---");
        Run("新建 名称校验", Test_NewItemValidate);
        Run("新建 文件与文件夹创建", Test_NewItemCreate);
        Run("新建 重名与非法名拒绝", Test_NewItemReject);
        Run("新建 Office 模板可打开", Test_NewItemOfficeTemplates);
        Run("VCS 清理 非版本控制目录拒绝", Test_VcsCleanupNoneRejected);
        Run("VCS 清理 Git 进程可启动", Test_VcsCleanupStartsGit);
        Run("VCS 清理 SVN 进程可启动", Test_VcsCleanupStartsSvn);

        Console.WriteLine("--- 文件操作失败提示 ---");
        Run("失败提示 权限错误给出中文建议", Test_OperationErrorPermission);
        Run("失败提示 普通错误保留原始信息", Test_OperationErrorGeneric);

        Console.WriteLine("--- 系统右键菜单 ---");

        Console.WriteLine("--- UI 构造冒烟 ---");
        Run("搜索/批量重命名窗体可在 STA 线程构造", Test_UiFormsConstruct);
    }

    // ==================== 复制 ====================

    private static void Test_CopyFilePreservesMetadata()
    {
        var root = NewTempDir();
        try
        {
            var sourceDir = Path.Combine(root, "src");
            var targetDir = Path.Combine(root, "dst");
            Directory.CreateDirectory(sourceDir);
            Directory.CreateDirectory(targetDir);

            var source = Path.Combine(sourceDir, "a.txt");
            File.WriteAllText(source, "hello");
            var stamp = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc);
            File.SetLastWriteTimeUtc(source, stamp);

            long copied = 0;
            FileTransferEngine.CopyPath(source, targetDir,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite },
                null, bytes => copied += bytes, new PauseToken(), CancellationToken.None);

            var target = Path.Combine(targetDir, "a.txt");
            Assert(File.Exists(target), "目标文件应存在");
            AssertEq("hello", File.ReadAllText(target), "文件内容");
            AssertEq(stamp, File.GetLastWriteTimeUtc(target), "修改时间应保留");
            Assert(copied > 0, "应上报复制字节数");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_CopyDirectoryRecursive()
    {
        var root = NewTempDir();
        try
        {
            var source = Path.Combine(root, "src");
            var target = Path.Combine(root, "dst");
            Directory.CreateDirectory(Path.Combine(source, "sub"));
            Directory.CreateDirectory(target);
            File.WriteAllText(Path.Combine(source, "a.txt"), "A");
            File.WriteAllText(Path.Combine(source, "sub", "b.txt"), "B");

            FileTransferEngine.CopyPath(source, target,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite },
                null, null, new PauseToken(), CancellationToken.None);

            AssertEq("A", File.ReadAllText(Path.Combine(target, "src", "a.txt")), "根文件");
            AssertEq("B", File.ReadAllText(Path.Combine(target, "src", "sub", "b.txt")), "子目录文件");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    // ==================== 冲突 ====================

    private static void Test_ConflictSkip()
    {
        var root = NewTempDir();
        try
        {
            var (source, target) = CreateConflict(root, "new", "old");
            FileTransferEngine.CopyPath(source, Path.GetDirectoryName(target)!,
                new FileOperationOptions { Conflict = ConflictPolicy.Skip },
                null, null, new PauseToken(), CancellationToken.None);
            AssertEq("old", File.ReadAllText(target), "跳过时目标内容不变");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_ConflictOverwrite()
    {
        var root = NewTempDir();
        try
        {
            var (source, target) = CreateConflict(root, "new", "old");
            FileTransferEngine.CopyPath(source, Path.GetDirectoryName(target)!,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite },
                null, null, new PauseToken(), CancellationToken.None);
            AssertEq("new", File.ReadAllText(target), "覆盖后目标内容更新");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_ConflictRename()
    {
        var root = NewTempDir();
        try
        {
            var (source, target) = CreateConflict(root, "new", "old");
            FileTransferEngine.CopyPath(source, Path.GetDirectoryName(target)!,
                new FileOperationOptions { Conflict = ConflictPolicy.Rename },
                null, null, new PauseToken(), CancellationToken.None);

            AssertEq("old", File.ReadAllText(target), "原目标保持不变");
            var renamed = Path.Combine(Path.GetDirectoryName(target)!, "same (2).txt");
            Assert(File.Exists(renamed), "应生成 same (2).txt");
            AssertEq("new", File.ReadAllText(renamed), "重命名目标内容");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_ConflictKeepNewer()
    {
        var root = NewTempDir();
        try
        {
            var (source, target) = CreateConflict(root, "new", "old");
            var old = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var fresh = new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc);

            File.SetLastWriteTimeUtc(source, old);
            File.SetLastWriteTimeUtc(target, fresh);
            FileTransferEngine.CopyPath(source, Path.GetDirectoryName(target)!,
                new FileOperationOptions { Conflict = ConflictPolicy.KeepNewer },
                null, null, new PauseToken(), CancellationToken.None);
            AssertEq("old", File.ReadAllText(target), "源更旧时应跳过");

            File.SetLastWriteTimeUtc(source, fresh);
            File.SetLastWriteTimeUtc(target, old);
            FileTransferEngine.CopyPath(source, Path.GetDirectoryName(target)!,
                new FileOperationOptions { Conflict = ConflictPolicy.KeepNewer },
                null, null, new PauseToken(), CancellationToken.None);
            AssertEq("new", File.ReadAllText(target), "源较新时应覆盖");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_ConflictAskApplyAll()
    {
        var root = NewTempDir();
        try
        {
            var sourceDir = Path.Combine(root, "src");
            var targetDir = Path.Combine(root, "dst");
            Directory.CreateDirectory(sourceDir);
            Directory.CreateDirectory(targetDir);
            File.WriteAllText(Path.Combine(sourceDir, "a.txt"), "newA");
            File.WriteAllText(Path.Combine(sourceDir, "b.txt"), "newB");
            File.WriteAllText(Path.Combine(targetDir, "a.txt"), "oldA");
            File.WriteAllText(Path.Combine(targetDir, "b.txt"), "oldB");

            int asked = 0;
            using var queue = new FileOperationQueue(1);
            var job = queue.Enqueue(FileOperationKind.Copy,
                new[] { Path.Combine(sourceDir, "a.txt"), Path.Combine(sourceDir, "b.txt") },
                targetDir,
                new FileOperationOptions { Conflict = ConflictPolicy.Ask },
                _ =>
                {
                    asked++;
                    return new ConflictAnswer(ConflictDecision.Overwrite, ConflictPolicy.Overwrite);
                });

            WaitJob(job);
            AssertEq(FileOperationState.Completed, job.State, "任务应完成");
            AssertEq(1, asked, "选择全部应用后第二次冲突不应再询问");
            AssertEq("newA", File.ReadAllText(Path.Combine(targetDir, "a.txt")), "A 覆盖");
            AssertEq("newB", File.ReadAllText(Path.Combine(targetDir, "b.txt")), "B 覆盖");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    // ==================== 移动 ====================

    private static void Test_MoveFileAndDirectory()
    {
        var root = NewTempDir();
        try
        {
            var sourceDir = Path.Combine(root, "src");
            var targetDir = Path.Combine(root, "dst");
            Directory.CreateDirectory(Path.Combine(sourceDir, "folder"));
            Directory.CreateDirectory(targetDir);
            File.WriteAllText(Path.Combine(sourceDir, "a.txt"), "A");
            File.WriteAllText(Path.Combine(sourceDir, "folder", "b.txt"), "B");

            long movedBytes = 0;
            FileTransferEngine.MovePath(Path.Combine(sourceDir, "a.txt"), targetDir,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite },
                null, bytes => movedBytes += bytes, new PauseToken(), CancellationToken.None);
            AssertEq(1L, movedBytes, "同卷移动的进度字节数不应重复上报");
            Assert(!File.Exists(Path.Combine(sourceDir, "a.txt")), "移动后源文件应消失");
            AssertEq("A", File.ReadAllText(Path.Combine(targetDir, "a.txt")), "移动后目标内容");

            FileTransferEngine.MovePath(Path.Combine(sourceDir, "folder"), targetDir,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite },
                null, null, new PauseToken(), CancellationToken.None);
            Assert(!Directory.Exists(Path.Combine(sourceDir, "folder")), "移动后源目录应消失");
            AssertEq("B", File.ReadAllText(Path.Combine(targetDir, "folder", "b.txt")), "目录内文件");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    // ==================== 取消 / 队列 ====================

    private static void Test_CancelCleansTemp()
    {
        var root = NewTempDir();
        try
        {
            var sourceDir = Path.Combine(root, "src");
            var targetDir = Path.Combine(root, "dst");
            Directory.CreateDirectory(sourceDir);
            Directory.CreateDirectory(targetDir);

            var source = Path.Combine(sourceDir, "big.bin");
            var payload = new byte[2 * 1024 * 1024];
            new Random(42).NextBytes(payload);
            File.WriteAllBytes(source, payload);

            var cts = new CancellationTokenSource();
            long seen = 0;
            bool cancelled = false;
            try
            {
                FileTransferEngine.CopyPath(source, targetDir,
                    new FileOperationOptions { Conflict = ConflictPolicy.Overwrite },
                    null,
                    bytes =>
                    {
                        seen += bytes;
                        // 第一块数据已写进临时文件后再取消，确保覆盖"清理半截文件"的路径
                        if (!cancelled && seen >= 128 * 1024)
                        {
                            cancelled = true;
                            cts.Cancel();
                        }
                    },
                    new PauseToken(), cts.Token);
            }
            catch (OperationCanceledException)
            {
            }

            Assert(cancelled, "应在复制过程中发出取消");
            Assert(seen > 0, "取消前应已写入数据");
            AssertEq(0, Directory.GetFiles(targetDir, "*.wxd-tmp-*").Length, "取消后不应遗留临时文件");
            Assert(!File.Exists(Path.Combine(targetDir, "big.bin")), "取消后不应出现半截目标文件");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_QueueSequential()
    {
        var root = NewTempDir();
        try
        {
            var sourceDir = Path.Combine(root, "src");
            var dst1 = Path.Combine(root, "d1");
            var dst2 = Path.Combine(root, "d2");
            Directory.CreateDirectory(sourceDir);
            Directory.CreateDirectory(dst1);
            Directory.CreateDirectory(dst2);
            File.WriteAllText(Path.Combine(sourceDir, "a.txt"), "A");
            File.WriteAllText(Path.Combine(sourceDir, "b.txt"), "B");

            using var queue = new FileOperationQueue(1);
            var first = queue.Enqueue(FileOperationKind.Copy,
                new[] { Path.Combine(sourceDir, "a.txt") }, dst1,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite });
            var second = queue.Enqueue(FileOperationKind.Copy,
                new[] { Path.Combine(sourceDir, "b.txt") }, dst2,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite });

            WaitJob(first);
            WaitJob(second);

            AssertEq(FileOperationState.Completed, first.State, "第一个任务完成");
            AssertEq(FileOperationState.Completed, second.State, "第二个任务完成");
            Assert(first.StartedAt.HasValue && first.FinishedAt.HasValue, "第一个任务应有起止时间");
            Assert(second.StartedAt.HasValue, "第二个任务应有开始时间");
            Assert(first.RunOrder > 0 && second.RunOrder > first.RunOrder,
                $"单并发时第二个任务应在第一个之后获得执行槽（first={first.RunOrder} second={second.RunOrder}）");
            AssertEq(1, first.TotalItems, "第一个任务文件数");
            AssertEq(1, second.TotalItems, "第二个任务文件数");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_QueueFailure()
    {
        var root = NewTempDir();
        try
        {
            using var queue = new FileOperationQueue(1);
            var job = queue.Enqueue(FileOperationKind.Copy,
                new[] { Path.Combine(root, "missing.txt") }, root,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite });

            WaitJob(job);
            AssertEq(FileOperationState.Failed, job.State, "不存在的源应失败");
            Assert(job.Errors.Count > 0, "应收集失败明细");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_QueueFailedEvent()
    {
        var root = NewTempDir();
        try
        {
            using var queue = new FileOperationQueue(1);
            FileOperationJob? failed = null;
            queue.JobFailed += (_, job) => failed = job;

            var job = queue.Enqueue(FileOperationKind.Copy,
                new[] { Path.Combine(root, "missing.txt") }, root,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite });

            WaitJob(job);
            AssertEq(FileOperationState.Failed, job.State, "任务应失败");
            AssertEq(job.Id, failed?.Id, "失败事件应携带同一任务");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_PauseToken()
    {
        var token = new PauseToken();
        token.Pause();

        using var reached = new ManualResetEventSlim();
        using var released = new ManualResetEventSlim();
        var task = Task.Run(() =>
        {
            reached.Set();
            token.WaitIfPaused(CancellationToken.None);
            released.Set();
        });

        Assert(reached.Wait(2000), "等待线程应已启动");
        Assert(!released.Wait(150), "暂停期间不应继续");
        token.Resume();
        Assert(released.Wait(2000), "继续后应放行");
        task.Wait(2000);
    }

    // ==================== 辅助 ====================

    private static void Test_UniqueAndLongPath()
    {
        var root = NewTempDir();
        try
        {
            var file = Path.Combine(root, "a.txt");
            File.WriteAllText(file, "1");
            var second = FileTransferEngine.GetUniquePath(file);
            AssertEq("a (2).txt", Path.GetFileName(second), "第二个不冲突名称");
            File.WriteAllText(second, "2");
            AssertEq("a (3).txt", Path.GetFileName(FileTransferEngine.GetUniquePath(file)), "第三个不冲突名称");

            AssertEq(file, FileTransferEngine.X(file), "短路径不应加前缀");
            var longPath = Path.Combine(root, new string('x', 300) + ".txt");
            Assert(FileTransferEngine.X(longPath).StartsWith(@"\\?\"), "长路径应加 \\\\?\\ 前缀");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_Measure()
    {
        var root = NewTempDir();
        try
        {
            Directory.CreateDirectory(Path.Combine(root, "sub"));
            File.WriteAllText(Path.Combine(root, "a.txt"), "12345");
            File.WriteAllText(Path.Combine(root, "sub", "b.txt"), "123");
            var measured = FileTransferEngine.Measure(new[] { root });
            AssertEq(2, measured.Files, "文件数");
            AssertEq(8L, measured.Bytes, "总字节数");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    // ==================== 搜索 ====================

    private static void Test_SearchNameContains()
    {
        var root = NewTempDir();
        try
        {
            CreateSearchTree(root);
            var result = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, Pattern = "alpha" }, null, CancellationToken.None);

            var names = result.Hits.Select(h => h.Name).OrderBy(n => n, StringComparer.Ordinal);
            AssertSeq(new[] { "alpha.txt", "alpha2.txt" }, names, "包含匹配应忽略大小写");
            Assert(result.FilesScanned > 0, "应统计扫描文件数");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchWildcardAndRegex()
    {
        var root = NewTempDir();
        try
        {
            CreateSearchTree(root);

            var wildcard = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, Pattern = "*.txt" }, null, CancellationToken.None);
            AssertSeq(new[] { "alpha.txt", "alpha2.txt" },
                wildcard.Hits.Select(h => h.Name).OrderBy(n => n, StringComparer.Ordinal),
                "通配符 *.txt");

            var regex = FileSearchService.Search(
                new FileSearchQuery
                {
                    Roots = new[] { root },
                    Pattern = "^alpha\\d*\\.txt$",
                    UseRegex = true
                }, null, CancellationToken.None);
            AssertSeq(new[] { "alpha.txt", "alpha2.txt" },
                regex.Hits.Select(h => h.Name).OrderBy(n => n, StringComparer.Ordinal),
                "正则锚定匹配");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchRecursiveToggle()
    {
        var root = NewTempDir();
        try
        {
            CreateSearchTree(root);

            var recursive = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, Pattern = "*.txt" }, null, CancellationToken.None);
            AssertEq(2, recursive.Hits.Count, "递归时应命中根目录与子目录的 txt");

            var flat = FileSearchService.Search(
                new FileSearchQuery
                {
                    Roots = new[] { root },
                    Pattern = "*.txt",
                    IncludeSubdirectories = false
                }, null, CancellationToken.None);
            AssertSeq(new[] { "alpha.txt" }, flat.Hits.Select(h => h.Name), "不递归只返回根目录文件");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchHiddenToggle()
    {
        var root = NewTempDir();
        try
        {
            CreateSearchTree(root);

            var withoutHidden = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, Pattern = "hidden*" }, null, CancellationToken.None);
            AssertEq(0, withoutHidden.Hits.Count, "默认不返回隐藏文件");

            var withHidden = FileSearchService.Search(
                new FileSearchQuery
                {
                    Roots = new[] { root },
                    Pattern = "hidden*",
                    IncludeHidden = true
                }, null, CancellationToken.None);
            AssertEq(1, withHidden.Hits.Count, "包含隐藏文件时应命中");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchContentLineNumbers()
    {
        var root = NewTempDir();
        try
        {
            CreateSearchTree(root);
            var result = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, ContentPattern = "alpha" }, null, CancellationToken.None);

            var hits = result.Hits
                .Select(h => h.Name + ":" + h.LineNumber)
                .OrderBy(x => x, StringComparer.Ordinal)
                .ToArray();
            AssertSeq(new[] { "alpha.txt:1", "alpha2.txt:1" }, hits, "内容命中应带物理行号");
            Assert(result.Hits.All(h => !h.IsDirectory), "内容搜索不应返回目录");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchContentCaseAndRegex()
    {
        var root = NewTempDir();
        try
        {
            CreateSearchTree(root);

            var insensitive = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, ContentPattern = "beta" }, null, CancellationToken.None);
            AssertEq(1, insensitive.Hits.Count, "默认忽略大小写");
            AssertEq(2, insensitive.Hits[0].LineNumber, "Beta two 在第 2 行");

            var sensitive = FileSearchService.Search(
                new FileSearchQuery
                {
                    Roots = new[] { root },
                    ContentPattern = "beta",
                    CaseSensitive = true
                }, null, CancellationToken.None);
            AssertEq(0, sensitive.Hits.Count, "区分大小写时 beta 不应命中 Beta");

            var regex = FileSearchService.Search(
                new FileSearchQuery
                {
                    Roots = new[] { root },
                    ContentPattern = "^Beta two$",
                    ContentRegex = true
                }, null, CancellationToken.None);
            AssertEq(1, regex.Hits.Count, "内容正则锚定");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchSizeAndDate()
    {
        var root = NewTempDir();
        try
        {
            var small = Path.Combine(root, "small.txt");
            File.WriteAllText(small, "x");
            var big = Path.Combine(root, "big.txt");
            File.WriteAllText(big, new string('y', 100));

            var bySize = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, Pattern = "*.txt", MinSize = 2 },
                null, CancellationToken.None);
            AssertSeq(new[] { "big.txt" }, bySize.Hits.Select(h => h.Name), "大小过滤");

            File.SetLastWriteTimeUtc(small, new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc));
            var byDate = FileSearchService.Search(
                new FileSearchQuery
                {
                    Roots = new[] { root },
                    Pattern = "small.txt",
                    ModifiedAfter = new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Utc)
                }, null, CancellationToken.None);
            AssertEq(0, byDate.Hits.Count, "修改时间过滤");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchMaxResults()
    {
        var root = NewTempDir();
        try
        {
            for (int i = 0; i < 5; i++)
                File.WriteAllText(Path.Combine(root, "f" + i + ".txt"), "x");

            var result = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, Pattern = "*.txt", MaxResults = 2 },
                null, CancellationToken.None);
            AssertEq(2, result.Hits.Count, "结果数应被上限截断");
            Assert(result.Truncated, "应标记为截断");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchCancelAndBadRegex()
    {
        var root = NewTempDir();
        try
        {
            File.WriteAllText(Path.Combine(root, "a.txt"), "x");
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            AssertThrows<OperationCanceledException>(() => FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, Pattern = "*" }, null, cts.Token));

            AssertThrows<ArgumentException>(() => FileSearchService.BuildNameRegex("[", true, false));
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchBinarySkip()
    {
        var root = NewTempDir();
        try
        {
            var binary = Path.Combine(root, "data.bin");
            File.WriteAllBytes(binary, new byte[] { 0x61, 0x00, 0x62, 0x63 });
            var result = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, ContentPattern = "a" }, null, CancellationToken.None);
            AssertEq(0, result.Hits.Count, "含 NUL 的二进制文件不应做内容命中");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_SearchDirectoryHit()
    {
        var root = NewTempDir();
        try
        {
            CreateSearchTree(root);
            var result = FileSearchService.Search(
                new FileSearchQuery { Roots = new[] { root }, Pattern = "sub" }, null, CancellationToken.None);
            Assert(result.Hits.Any(h => h.IsDirectory && h.Name == "sub"), "目录名应参与匹配");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    // ==================== 批量重命名 ====================

    private static void Test_RenamePlainReplace()
    {
        var root = NewTempDir();
        try
        {
            var file = Path.Combine(root, "a.txt");
            File.WriteAllText(file, "x");
            var plan = BatchRenameEngine.BuildPlan(new[] { file },
                new RenameRule { Find = "A", Replace = "b" });
            AssertEq(1, plan.Items.Count, "计划项数");
            AssertEq("b.txt", plan.Items[0].TargetName, "忽略大小写替换");
            AssertEq(RenamePlanStatus.Ready, plan.Items[0].Status, "状态");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_RenameRegexCapture()
    {
        var root = NewTempDir();
        try
        {
            var file = Path.Combine(root, "item_12.txt");
            File.WriteAllText(file, "x");
            var plan = BatchRenameEngine.BuildPlan(new[] { file },
                new RenameRule { Find = "item_(\\d+)", Replace = "entry_$1", UseRegex = true });
            AssertEq("entry_12.txt", plan.Items[0].TargetName, "正则捕获组替换");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_RenameTemplateCounterAndExt()
    {
        var root = NewTempDir();
        try
        {
            var files = new[] { "a.txt", "b.txt", "c.txt" }
                .Select(n => { var p = Path.Combine(root, n); File.WriteAllText(p, "x"); return p; })
                .ToArray();

            var plan = BatchRenameEngine.BuildPlan(files, new RenameRule
            {
                Template = "{name}_{counter}",
                CounterStart = 5,
                CounterStep = 5,
                CounterDigits = 3
            });
            AssertSeq(new[] { "a_005.txt", "b_010.txt", "c_015.txt" },
                plan.Items.Select(i => i.TargetName), "模板序号与扩展名保留");

            var extPlan = BatchRenameEngine.BuildPlan(new[] { files[0] },
                new RenameRule { Template = "pre_{name}" });
            AssertEq("pre_a.txt", extPlan.Items[0].TargetName, "模板不含 {ext} 时自动保留扩展名");

            var explicitExt = BatchRenameEngine.BuildPlan(new[] { files[0] },
                new RenameRule { Template = "{name}.{ext}" });
            AssertEq("a.txt", explicitExt.Items[0].TargetName, "显式 {ext} 不重复追加");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_RenameCaseOnly()
    {
        var root = NewTempDir();
        try
        {
            var file = Path.Combine(root, "case.txt");
            File.WriteAllText(file, "x");

            var plan = BatchRenameEngine.BuildPlan(new[] { file },
                new RenameRule { Template = "{name}", CaseMode = RenameCaseMode.Upper });
            AssertEq("CASE.txt", plan.Items[0].TargetName, "仅大小写变化");
            AssertEq(RenamePlanStatus.Ready, plan.Items[0].Status, "不应与自身冲突");

            var applied = BatchRenameEngine.Apply(plan);
            Assert(applied.Success, "仅大小写变化应能执行: " + string.Join(";", applied.Errors));

            var actual = Path.GetFileName(Directory.GetFiles(root, "*.txt").Single());
            AssertEq("CASE.txt", actual, "磁盘上的实际名称");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_RenameInvalidAndReserved()
    {
        var root = NewTempDir();
        try
        {
            var file = Path.Combine(root, "a.txt");
            File.WriteAllText(file, "x");

            var invalid = BatchRenameEngine.BuildPlan(new[] { file },
                new RenameRule { Template = "{name}*" });
            AssertEq(RenamePlanStatus.Invalid, invalid.Items[0].Status, "非法字符");
            Assert(invalid.HasErrors, "非法字符应视为错误");

            var reserved = BatchRenameEngine.BuildPlan(new[] { file },
                new RenameRule { Template = "CON" });
            AssertEq(RenamePlanStatus.Invalid, reserved.Items[0].Status, "保留名");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_RenameDuplicateAndConflict()
    {
        var root = NewTempDir();
        try
        {
            var a = Path.Combine(root, "a.txt");
            var b = Path.Combine(root, "b.txt");
            var c = Path.Combine(root, "c.txt");
            File.WriteAllText(a, "a");
            File.WriteAllText(b, "b");
            File.WriteAllText(c, "c");

            var duplicate = BatchRenameEngine.BuildPlan(new[] { a, b },
                new RenameRule { Template = "same" });
            Assert(duplicate.Items.All(i => i.Status == RenamePlanStatus.Duplicate), "两个文件指向同一名称");

            var conflict = BatchRenameEngine.BuildPlan(new[] { a },
                new RenameRule { Find = "a", Replace = "c" });
            AssertEq(RenamePlanStatus.Conflict, conflict.Items[0].Status, "目标已存在应冲突");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_RenameRefusesOnError()
    {
        var root = NewTempDir();
        try
        {
            var a = Path.Combine(root, "a.txt");
            var b = Path.Combine(root, "b.txt");
            File.WriteAllText(a, "a");
            File.WriteAllText(b, "b");

            var plan = BatchRenameEngine.BuildPlan(new[] { a, b }, new RenameRule { Template = "same" });
            var result = BatchRenameEngine.Apply(plan);
            Assert(!result.Success, "存在冲突时不应执行");
            Assert(File.Exists(a) && File.Exists(b), "原文件应保持不变");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_RenameApplyAndUndo()
    {
        var root = NewTempDir();
        try
        {
            var a = Path.Combine(root, "a.txt");
            var b = Path.Combine(root, "b.txt");
            File.WriteAllText(a, "A");
            File.WriteAllText(b, "B");

            var plan = BatchRenameEngine.BuildPlan(new[] { a, b },
                new RenameRule { Template = "{name}_new" });
            var result = BatchRenameEngine.Apply(plan);
            Assert(result.Success, "执行失败: " + string.Join(";", result.Errors));
            AssertEq(2, result.Journal.Count, "日志条数");

            var aNew = Path.Combine(root, "a_new.txt");
            var bNew = Path.Combine(root, "b_new.txt");
            Assert(File.Exists(aNew) && File.Exists(bNew), "应生成新名称");
            Assert(!File.Exists(a) && !File.Exists(b), "旧名称应消失");

            Assert(BatchRenameEngine.Undo(result.Journal, out var error), "撤销失败: " + error);
            Assert(File.Exists(a) && File.Exists(b), "撤销后应恢复原名");
            Assert(!File.Exists(aNew) && !File.Exists(bNew), "撤销后新名称应消失");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_RenameIncludeExtension()
    {
        var root = NewTempDir();
        try
        {
            var file = Path.Combine(root, "a.txt");
            File.WriteAllText(file, "x");

            var plan = BatchRenameEngine.BuildPlan(new[] { file }, new RenameRule
            {
                Find = ".txt",
                Replace = ".md",
                IncludeExtensionInFind = true
            });
            AssertEq("a.md", plan.Items[0].TargetName, "扩展名参与替换后不应重复追加");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void CreateSearchTree(string root)
    {
        Directory.CreateDirectory(Path.Combine(root, "sub", "deep"));
        File.WriteAllText(Path.Combine(root, "alpha.txt"), "alpha one\nBeta two\n");
        File.WriteAllText(Path.Combine(root, "b.csv"), "x,y\n");
        File.WriteAllText(Path.Combine(root, "sub", "alpha2.txt"), "alpha in sub\n");
        File.WriteAllText(Path.Combine(root, "sub", "deep", "c.log"), "gamma\n");

        var hidden = Path.Combine(root, "hidden.txt");
        File.WriteAllText(hidden, "alpha hidden\n");
        File.SetAttributes(hidden, FileAttributes.Hidden);
    }

    private static void Test_NewItemValidate()
    {
        var root = NewTempDir();
        try
        {
            Assert(NewItemService.Validate(root, "a.txt") == null, "普通名称应可用");
            Assert(NewItemService.Validate(root, "") != null, "空名称应拒绝");
            Assert(NewItemService.Validate(root, "a*b.txt") != null, "非法字符应拒绝");
            Assert(NewItemService.Validate(root, "CON") != null, "保留名应拒绝");
            Assert(NewItemService.Validate(root, "a.") != null, "点结尾应拒绝");

            File.WriteAllText(Path.Combine(root, "exists.txt"), "x");
            Assert(NewItemService.Validate(root, "exists.txt") != null, "重名应拒绝");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_NewItemCreate()
    {
        var root = NewTempDir();
        try
        {
            Assert(NewItemService.TryCreateFile(root, "new.txt", out var file, out var fileError),
                "创建文件失败: " + fileError);
            Assert(File.Exists(file), "文件应存在");

            Assert(NewItemService.TryCreateDirectory(root, "newdir", out var directory, out var dirError),
                "创建文件夹失败: " + dirError);
            Assert(Directory.Exists(directory), "文件夹应存在");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_NewItemReject()
    {
        var root = NewTempDir();
        try
        {
            File.WriteAllText(Path.Combine(root, "dup.txt"), "x");
            Assert(!NewItemService.TryCreateFile(root, "dup.txt", out _, out var error), "重名文件应拒绝");
            Assert(error != null, "应返回原因");

            Assert(!NewItemService.TryCreateDirectory(root, "bad*name", out _, out var dirError), "非法名应拒绝");
            Assert(dirError != null, "应返回原因");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_VcsCleanupNoneRejected()
    {
        var root = NewTempDir();
        try
        {
            var result = VersionControlHelper.CleanupAsync(root, VcsType.None).GetAwaiter().GetResult();
            Assert(!result.Started, "None 类型不应启动清理进程");
            Assert(!string.IsNullOrEmpty(result.ErrorMessage), "应给出不支持的提示");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_VcsCleanupStartsGit()
    {
        var probe = VersionControlHelper.RunProcessAsync("git", "--version", timeoutMs: 10_000)
            .GetAwaiter().GetResult();
        if (!probe.Started)
            return; // 环境没有 Git，跳过

        var root = NewTempDir();
        try
        {
            Directory.CreateDirectory(Path.Combine(root, ".git"));
            var result = VersionControlHelper.CleanupAsync(root, VcsType.Git).GetAwaiter().GetResult();
            Assert(result.Started, "Git 清理应能启动进程: " + result.ErrorMessage);
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_VcsCleanupStartsSvn()
    {
        var root = NewTempDir();
        try
        {
            Directory.CreateDirectory(Path.Combine(root, ".svn"));
            var result = VersionControlHelper.CleanupAsync(root, VcsType.Svn).GetAwaiter().GetResult();
            Assert(result.Started, "SVN 清理应能启动进程: " + result.ErrorMessage);
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_NewItemOfficeTemplates()
    {
        var root = NewTempDir();
        try
        {
            Assert(NewItemService.TryCreateTypedFile(root, "a.xlsx", out var xlsx, out var xlsxError),
                "xlsx 创建失败: " + xlsxError);
            using (var zip = ZipFile.OpenRead(xlsx))
            {
                Assert(zip.GetEntry("xl/workbook.xml") != null, "xlsx 缺少 xl/workbook.xml");
                Assert(zip.GetEntry("xl/worksheets/sheet1.xml") != null, "xlsx 缺少 sheet1.xml");
            }

            Assert(NewItemService.TryCreateTypedFile(root, "b.docx", out var docx, out var docxError),
                "docx 创建失败: " + docxError);
            using (var zip = ZipFile.OpenRead(docx))
            {
                Assert(zip.GetEntry("[Content_Types].xml") != null, "docx 缺少 [Content_Types].xml");
                Assert(zip.GetEntry("word/document.xml") != null, "docx 缺少 word/document.xml");
            }

            Assert(NewItemService.TryCreateTypedFile(root, "c.txt", out var txt, out var txtError),
                "txt 创建失败: " + txtError);
            Assert(new FileInfo(txt).Length == 0, "txt 应为空文件");
        }
        finally
        {
            DeleteTempDir(root);
        }
    }

    private static void Test_OperationErrorPermission()
    {
        var job = new FileOperationJob
        {
            Kind = FileOperationKind.Copy,
            Destination = @"D:\xxx"
        };
        job.FailedItems = 1;
        job.Errors.Add(@"D:\策划配置\cfg_monster.csv: Access to the path 'D:\xxx\cfg_monster.csv.wxd-tmp-c005c002' is denied.");

        var text = OperationErrorPresenter.Build(job);
        Assert(text.Contains("没有写入权限"), "应说明是权限问题");
        Assert(text.Contains("以管理员身份"), "应给出下一步操作建议");
        Assert(text.Contains(@"D:\xxx"), "应回显目标位置");
    }

    private static void Test_OperationErrorGeneric()
    {
        var job = new FileOperationJob
        {
            Kind = FileOperationKind.Copy,
            Destination = @"D:\dst"
        };
        job.FailedItems = 1;
        job.Errors.Add(@"D:\a.txt: 源不存在");

        var text = OperationErrorPresenter.Build(job);
        Assert(text.Contains("源不存在"), "普通错误应保留原始信息");
        Assert(!text.Contains("以管理员身份"), "非权限错误不应给管理员建议");
    }

    private static void Test_UiFormsConstruct()
    {
        Exception? captured = null;
        var thread = new Thread(() =>
        {
            try
            {
                using var search = new FileSearchForm("C:\\", _ => { });
                Assert(!search.IsModifiedTimeFilterEnabled, "搜索窗口默认不应勾选修改时间过滤");
                using var newItem = new NewItemDialog("新建文件", "名称", "新建文件.txt",
                    name => NewItemService.Validate("C:\\", name));
                using var batch = new BatchRenameForm(new[]
                {
                    Path.Combine(Path.GetTempPath(), "wxd_not_exists_" + Guid.NewGuid().ToString("N") + ".txt")
                });
            }
            catch (Exception ex)
            {
                captured = ex;
            }
        })
        {
            IsBackground = true
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        Assert(thread.Join(10000), "UI 构造线程超时");
        if (captured != null) throw new Exception("窗体构造失败: " + captured);
    }

    private static void AssertThrows<T>(Action action) where T : Exception
    {
        try
        {
            action();
        }
        catch (T)
        {
            return;
        }
        catch (Exception ex)
        {
            throw new Exception("期望抛出 " + typeof(T).Name + "，实际 " + ex.GetType().Name);
        }
        throw new Exception("期望抛出 " + typeof(T).Name + "，但没有异常");
    }

    private static (string Source, string Target) CreateConflict(string root, string sourceText, string targetText)
    {
        var sourceDir = Path.Combine(root, "src");
        var targetDir = Path.Combine(root, "dst");
        Directory.CreateDirectory(sourceDir);
        Directory.CreateDirectory(targetDir);
        var source = Path.Combine(sourceDir, "same.txt");
        var target = Path.Combine(targetDir, "same.txt");
        File.WriteAllText(source, sourceText);
        File.WriteAllText(target, targetText);
        return (source, target);
    }

    private static void WaitJob(FileOperationJob job, int timeoutMs = 10000)
    {
        Assert(job.Finished.Task.Wait(timeoutMs), "任务超时，当前状态: " + job.StateText);
    }
}
