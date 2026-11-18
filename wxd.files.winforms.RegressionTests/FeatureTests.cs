using System;
using System.IO;
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
        Run("暂停闸门", Test_PauseToken);
        Run("唯一命名与长路径前缀", Test_UniqueAndLongPath);
        Run("统计文件数与字节", Test_Measure);
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

            FileTransferEngine.MovePath(Path.Combine(sourceDir, "a.txt"), targetDir,
                new FileOperationOptions { Conflict = ConflictPolicy.Overwrite },
                null, null, new PauseToken(), CancellationToken.None);
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
