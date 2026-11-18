using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;

namespace wxd.files.winforms;

/// <summary>文件操作类型。</summary>
public enum FileOperationKind
{
    Copy,
    Move,
    Delete
}

/// <summary>任务状态。</summary>
public enum FileOperationState
{
    Pending,
    Running,
    Paused,
    Completed,
    Failed,
    Cancelled
}

/// <summary>同名冲突处理策略。</summary>
public enum ConflictPolicy
{
    Ask,
    Overwrite,
    Skip,
    Rename,
    KeepNewer
}

/// <summary>单次冲突的处理结果。</summary>
public enum ConflictDecision
{
    Overwrite,
    Skip,
    Rename,
    Cancel
}

/// <summary>冲突回答；ApplyAll 非空时表示"以后同类冲突都按这个策略"。</summary>
public readonly record struct ConflictAnswer(ConflictDecision Decision, ConflictPolicy? ApplyAll = null);

/// <summary>一次同名冲突的上下文，供 UI 展示。</summary>
public sealed class ConflictInfo
{
    public string SourcePath { get; init; } = "";
    public string TargetPath { get; init; } = "";
    public bool IsDirectory { get; init; }
    public long SourceSize { get; init; }
    public long TargetSize { get; init; }
    public DateTime SourceModifiedUtc { get; init; }
    public DateTime TargetModifiedUtc { get; init; }
}

/// <summary>一次文件操作的选项。</summary>
public sealed class FileOperationOptions
{
    /// <summary>冲突策略；Ask 时会回调 UI。</summary>
    public ConflictPolicy Conflict { get; set; } = ConflictPolicy.Ask;

    /// <summary>复制完成后按 SHA256 校验。</summary>
    public bool VerifyAfterCopy { get; set; }

    /// <summary>删除时彻底删除（false 表示由调用方走回收站）。</summary>
    public bool PermanentlyDelete { get; set; }
}

/// <summary>可暂停闸门；每写一块数据前检查一次，暂停发生在块边界。</summary>
public sealed class PauseToken
{
    private readonly ManualResetEventSlim _gate = new(true);

    public bool IsPaused => !_gate.IsSet;

    public void Pause() => _gate.Reset();

    public void Resume() => _gate.Set();

    public void WaitIfPaused(CancellationToken token)
    {
        if (_gate.IsSet) return;
        _gate.Wait(token);
    }
}

/// <summary>一个文件操作任务。</summary>
public sealed class FileOperationJob
{
    private long _doneBytes;
    private int _doneItems;

    public string Id { get; } = Guid.NewGuid().ToString("N");
    public FileOperationKind Kind { get; init; }
    public IReadOnlyList<string> Sources { get; init; } = Array.Empty<string>();
    public string? Destination { get; init; }
    public FileOperationOptions Options { get; init; } = new();
    public FileOperationState State { get; internal set; } = FileOperationState.Pending;

    /// <summary>实际开始执行的序号（单并发时用于验证排队顺序）。</summary>
    public int RunOrder { get; internal set; }

    public long TotalBytes { get; internal set; }
    public long DoneBytes => Interlocked.Read(ref _doneBytes);
    public int TotalItems { get; internal set; }
    public int DoneItems => Volatile.Read(ref _doneItems);
    public int FailedItems { get; internal set; }

    public string? CurrentItem { get; internal set; }
    public string? Error { get; internal set; }
    public List<string> Errors { get; } = new();

    public DateTime CreatedAt { get; } = DateTime.Now;
    public DateTime? StartedAt { get; internal set; }
    public DateTime? FinishedAt { get; internal set; }

    public PauseToken Pause { get; } = new();
    public CancellationTokenSource Cancellation { get; } = new();

    internal Func<ConflictInfo, ConflictAnswer>? Ask { get; set; }
    internal TaskCompletionSource<bool> Finished { get; } =
        new(TaskCreationOptions.RunContinuationsAsynchronously);

    public double Percent
    {
        get
        {
            if (TotalBytes > 0) return Math.Min(100d, DoneBytes * 100d / TotalBytes);
            if (TotalItems > 0) return Math.Min(100d, DoneItems * 100d / TotalItems);
            return 0d;
        }
    }

    public string KindText => Kind switch
    {
        FileOperationKind.Copy => "复制",
        FileOperationKind.Move => "移动",
        _ => "删除"
    };

    public string StateText => State switch
    {
        FileOperationState.Pending => "等待",
        FileOperationState.Running => "进行中",
        FileOperationState.Paused => "已暂停",
        FileOperationState.Completed => "完成",
        FileOperationState.Failed => "失败",
        _ => "已取消"
    };

    internal void AddBytes(long bytes) => Interlocked.Add(ref _doneBytes, bytes);
    internal void AddItem() => Interlocked.Increment(ref _doneItems);
}

/// <summary>
/// 纯托管文件传输引擎：复制 / 移动 / 删除，支持进度、暂停、取消、冲突策略与复制后校验。
/// 不依赖 Shell，便于单元测试；删除的回收站语义仍由调用方用 Shell 处理。
/// </summary>
public static class FileTransferEngine
{
    private const int BufferSize = 128 * 1024;
    private const int LongPathThreshold = 240;

    /// <summary>统计文件数与总字节数；目录递归，忽略无权限项。</summary>
    public static (int Files, long Bytes) Measure(IEnumerable<string> paths)
    {
        int files = 0;
        long bytes = 0;
        foreach (var path in paths)
        {
            try
            {
                if (File.Exists(path))
                {
                    files++;
                    bytes += new FileInfo(path).Length;
                }
                else if (Directory.Exists(path))
                {
                    foreach (var file in Directory.EnumerateFiles(path, "*",
                        new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true }))
                    {
                        files++;
                        try { bytes += new FileInfo(file).Length; } catch { }
                    }
                }
            }
            catch
            {
                // 统计失败不影响真正执行
            }
        }
        return (files, bytes);
    }

    /// <summary>按策略决定冲突处理；Ask 时回调，可顺带把"全部应用"写回 options。</summary>
    public static ConflictDecision ResolveConflict(
        FileOperationOptions options,
        ConflictInfo info,
        Func<ConflictInfo, ConflictAnswer>? ask)
    {
        switch (options.Conflict)
        {
            case ConflictPolicy.Overwrite:
                return ConflictDecision.Overwrite;
            case ConflictPolicy.Skip:
                return ConflictDecision.Skip;
            case ConflictPolicy.Rename:
                return ConflictDecision.Rename;
            case ConflictPolicy.KeepNewer:
                return info.SourceModifiedUtc > info.TargetModifiedUtc
                    ? ConflictDecision.Overwrite
                    : ConflictDecision.Skip;
            default:
                var answer = ask?.Invoke(info) ?? new ConflictAnswer(ConflictDecision.Skip);
                if (answer.ApplyAll is ConflictPolicy policy)
                    options.Conflict = policy;
                return answer.Decision;
        }
    }

    /// <summary>目标已存在时生成 "名称 (2).ext" 形式的不冲突路径。</summary>
    public static string GetUniquePath(string path)
    {
        if (!Exists(path)) return path;

        var directory = Path.GetDirectoryName(path) ?? "";
        var name = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);
        if (string.IsNullOrEmpty(name))
        {
            name = Path.GetFileName(path);
            extension = "";
        }

        for (int i = 2; i < 100000; i++)
        {
            var candidate = Path.Combine(directory, name + " (" + i + ")" + extension);
            if (!Exists(candidate)) return candidate;
        }

        throw new IOException("无法生成不冲突的文件名: " + path);
    }

    /// <summary>按 SHA256 比较两个文件内容。</summary>
    public static bool ContentEquals(string left, string right)
    {
        using var leftStream = new FileStream(X(left), FileMode.Open, FileAccess.Read, FileShare.Read, BufferSize);
        using var rightStream = new FileStream(X(right), FileMode.Open, FileAccess.Read, FileShare.Read, BufferSize);
        using var sha = SHA256.Create();
        var leftHash = sha.ComputeHash(leftStream);
        var rightHash = sha.ComputeHash(rightStream);
        return leftHash.AsSpan().SequenceEqual(rightHash);
    }

    /// <summary>把文件或目录复制到目标目录（同名冲突按策略处理）。</summary>
    public static void CopyPath(
        string source,
        string destinationDirectory,
        FileOperationOptions options,
        Func<ConflictInfo, ConflictAnswer>? ask,
        Action<long>? onBytes,
        PauseToken pause,
        CancellationToken token)
    {
        if (!Exists(source)) throw new FileNotFoundException("源不存在", source);
        if (!Directory.Exists(destinationDirectory))
            throw new DirectoryNotFoundException("目标目录不存在: " + destinationDirectory);

        var target = Path.Combine(destinationDirectory, GetLeafName(source));
        CopyInternal(source, target, options, ask, onBytes, pause, token);
    }

    /// <summary>把文件或目录移动到目标目录（冲突按策略处理；跨卷退化为复制后删除）。</summary>
    public static void MovePath(
        string source,
        string destinationDirectory,
        FileOperationOptions options,
        Func<ConflictInfo, ConflictAnswer>? ask,
        Action<long>? onBytes,
        PauseToken pause,
        CancellationToken token)
    {
        if (!Exists(source)) throw new FileNotFoundException("源不存在", source);
        if (!Directory.Exists(destinationDirectory))
            throw new DirectoryNotFoundException("目标目录不存在: " + destinationDirectory);

        var target = Path.Combine(destinationDirectory, GetLeafName(source));
        bool targetExists = Exists(target);
        bool sourceIsDirectory = Directory.Exists(source);

        if (targetExists)
        {
            var decision = ResolveConflict(options, BuildConflictInfo(source, target, sourceIsDirectory), ask);
            switch (decision)
            {
                case ConflictDecision.Skip:
                    return;
                case ConflictDecision.Cancel:
                    throw new OperationCanceledException(token);
                case ConflictDecision.Rename:
                    target = GetUniquePath(target);
                    targetExists = false;
                    break;
            }
        }

        token.ThrowIfCancellationRequested();
        pause.WaitIfPaused(token);

        if (sourceIsDirectory)
        {
            if (!targetExists && SameVolume(source, target))
            {
                Directory.Move(X(source), X(target));
                var measured = Measure(new[] { target });
                onBytes?.Invoke(measured.Bytes);
                return;
            }

            if (File.Exists(target))
            {
                File.SetAttributes(X(target), FileAttributes.Normal);
                File.Delete(X(target));
            }
            if (Directory.Exists(target) && options.Conflict == ConflictPolicy.Overwrite)
                Directory.Delete(X(target), true);

            var copyOptions = CloneForOverwrite(options);
            CopyInternal(source, target, copyOptions, ask: null, onBytes, pause, token);
            DeletePath(source, pause, token);
            return;
        }

        if (File.Exists(target))
        {
            File.SetAttributes(X(target), FileAttributes.Normal);
            File.Delete(X(target));
        }

        var size = new FileInfo(source).Length;
        if (SameVolume(source, target))
        {
            File.Move(X(source), X(target));
        }
        else
        {
            var copyOptions = CloneForOverwrite(options);
            CopyFileInternal(source, target, copyOptions, ask: null, onBytes, pause, token);
            DeletePath(source, pause, token);
        }
        onBytes?.Invoke(size);
    }

    /// <summary>彻底删除文件或目录（递归，可取消）。</summary>
    public static void DeletePath(string path, PauseToken? pause, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        pause?.WaitIfPaused(token);

        if (File.Exists(path))
        {
            File.SetAttributes(X(path), FileAttributes.Normal);
            File.Delete(X(path));
            return;
        }

        if (!Directory.Exists(path)) return;

        foreach (var file in Directory.EnumerateFiles(path, "*",
            new EnumerationOptions { IgnoreInaccessible = true }))
        {
            token.ThrowIfCancellationRequested();
            pause?.WaitIfPaused(token);
            try
            {
                File.SetAttributes(X(file), FileAttributes.Normal);
                File.Delete(X(file));
            }
            catch
            {
                // 单个文件失败继续删其余项，最后目录删除会暴露问题
            }
        }

        foreach (var directory in Directory.EnumerateDirectories(path, "*",
            new EnumerationOptions { IgnoreInaccessible = true }))
        {
            DeletePath(directory, pause, token);
        }

        Directory.Delete(X(path), false);
    }

    private static void CopyInternal(
        string source,
        string target,
        FileOperationOptions options,
        Func<ConflictInfo, ConflictAnswer>? ask,
        Action<long>? onBytes,
        PauseToken pause,
        CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        pause.WaitIfPaused(token);

        if (!Directory.Exists(source))
        {
            CopyFileInternal(source, target, options, ask, onBytes, pause, token);
            return;
        }

        // 目录与目录同名时合并内容，不对目录本身弹冲突框
        if (File.Exists(target))
        {
            File.SetAttributes(X(target), FileAttributes.Normal);
            File.Delete(X(target));
        }
        Directory.CreateDirectory(X(target));

        foreach (var file in Directory.EnumerateFiles(source, "*",
            new EnumerationOptions { IgnoreInaccessible = true }))
        {
            token.ThrowIfCancellationRequested();
            pause.WaitIfPaused(token);
            CopyFileInternal(file, Path.Combine(target, Path.GetFileName(file)),
                options, ask, onBytes, pause, token);
        }

        foreach (var directory in Directory.EnumerateDirectories(source, "*",
            new EnumerationOptions { IgnoreInaccessible = true }))
        {
            token.ThrowIfCancellationRequested();
            pause.WaitIfPaused(token);
            CopyInternal(directory, Path.Combine(target, Path.GetFileName(directory)),
                options, ask, onBytes, pause, token);
        }

        TryCopyDirectoryMetadata(source, target);
    }

    private static void CopyFileInternal(
        string source,
        string target,
        FileOperationOptions options,
        Func<ConflictInfo, ConflictAnswer>? ask,
        Action<long>? onBytes,
        PauseToken pause,
        CancellationToken token)
    {
        if (File.Exists(target))
        {
            var decision = ResolveConflict(options, BuildConflictInfo(source, target, false), ask);
            switch (decision)
            {
                case ConflictDecision.Skip:
                    return;
                case ConflictDecision.Cancel:
                    throw new OperationCanceledException(token);
                case ConflictDecision.Rename:
                    target = GetUniquePath(target);
                    break;
                case ConflictDecision.Overwrite:
                    File.SetAttributes(X(target), FileAttributes.Normal);
                    break;
            }
        }

        var temp = target + ".wxd-tmp-" + Guid.NewGuid().ToString("N")[..8];
        try
        {
            using (var input = new FileStream(X(source), FileMode.Open, FileAccess.Read, FileShare.Read,
                       BufferSize, FileOptions.SequentialScan))
            using (var output = new FileStream(X(temp), FileMode.CreateNew, FileAccess.Write, FileShare.None,
                       BufferSize))
            {
                var buffer = new byte[BufferSize];
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    pause.WaitIfPaused(token);
                    token.ThrowIfCancellationRequested();
                    output.Write(buffer, 0, read);
                    onBytes?.Invoke(read);
                }
            }

            File.SetLastWriteTimeUtc(X(temp), File.GetLastWriteTimeUtc(X(source)));
            File.SetCreationTimeUtc(X(temp), File.GetCreationTimeUtc(X(source)));
            File.SetAttributes(X(temp), File.GetAttributes(X(source)));

            if (options.VerifyAfterCopy && !ContentEquals(source, temp))
                throw new IOException("复制后校验失败: " + target);

            File.Move(X(temp), X(target), true);
        }
        catch
        {
            TryDelete(temp);
            throw;
        }
    }

    private static ConflictInfo BuildConflictInfo(string source, string target, bool isDirectory)
    {
        long sourceSize = 0, targetSize = 0;
        DateTime sourceTime = default, targetTime = default;
        try { sourceSize = isDirectory ? 0 : new FileInfo(source).Length; } catch { }
        try { targetSize = isDirectory ? 0 : new FileInfo(target).Length; } catch { }
        try { sourceTime = File.GetLastWriteTimeUtc(X(source)); } catch { }
        try { targetTime = File.GetLastWriteTimeUtc(X(target)); } catch { }
        return new ConflictInfo
        {
            SourcePath = source,
            TargetPath = target,
            IsDirectory = isDirectory,
            SourceSize = sourceSize,
            TargetSize = targetSize,
            SourceModifiedUtc = sourceTime,
            TargetModifiedUtc = targetTime
        };
    }

    private static FileOperationOptions CloneForOverwrite(FileOperationOptions options) => new()
    {
        Conflict = ConflictPolicy.Overwrite,
        VerifyAfterCopy = options.VerifyAfterCopy,
        PermanentlyDelete = options.PermanentlyDelete
    };

    private static bool SameVolume(string left, string right)
    {
        try
        {
            return string.Equals(Path.GetPathRoot(Path.GetFullPath(left)),
                Path.GetPathRoot(Path.GetFullPath(right)), StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static string GetLeafName(string path)
    {
        var trimmed = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var name = Path.GetFileName(trimmed);
        return string.IsNullOrEmpty(name) ? trimmed : name;
    }

    private static bool Exists(string path) => File.Exists(path) || Directory.Exists(path);

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.SetAttributes(X(path), FileAttributes.Normal);
                File.Delete(X(path));
            }
        }
        catch
        {
            // 清理临时文件失败不覆盖原始异常
        }
    }

    private static void TryCopyDirectoryMetadata(string source, string target)
    {
        try { File.SetLastWriteTimeUtc(X(target), File.GetLastWriteTimeUtc(X(source))); } catch { }
        try { File.SetAttributes(X(target), File.GetAttributes(X(source))); } catch { }
    }

    /// <summary>超过阈值时加 \\?\ 前缀，避免长路径截断。</summary>
    internal static string X(string path)
    {
        if (string.IsNullOrEmpty(path)) return path;
        if (path.Length < LongPathThreshold) return path;
        if (path.StartsWith(@"\\?\", StringComparison.Ordinal)) return path;
        if (path.StartsWith(@"\\", StringComparison.Ordinal)) return @"\\?\UNC\" + path[2..];
        return @"\\?\" + Path.GetFullPath(path);
    }
}
