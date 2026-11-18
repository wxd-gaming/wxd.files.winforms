using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace wxd.files.winforms;

/// <summary>
/// 文件操作队列：任务排队执行，支持并发上限、暂停/继续/取消、失败项收集。
/// 所有进度字段由引擎在后台线程写入，UI 只读轮询，避免跨线程事件风暴。
/// </summary>
public sealed class FileOperationQueue : IDisposable
{
    /// <summary>进程级共享队列；传输中心与各窗格共用。</summary>
    public static FileOperationQueue Shared { get; } = new(1);

    private readonly object _sync = new();
    private readonly List<FileOperationJob> _jobs = new();
    private readonly SemaphoreSlim _slots;
    private int _runSequence;
    private bool _disposed;

    public FileOperationQueue(int maxConcurrentJobs = 1)
    {
        if (maxConcurrentJobs < 1) maxConcurrentJobs = 1;
        MaxConcurrentJobs = maxConcurrentJobs;
        _slots = new SemaphoreSlim(maxConcurrentJobs, maxConcurrentJobs);
    }

    public int MaxConcurrentJobs { get; }

    /// <summary>任务状态变化（入队/开始/暂停/结束）；进度变化请轮询。</summary>
    public event EventHandler<FileOperationJob>? JobChanged;

    public IReadOnlyList<FileOperationJob> Jobs
    {
        get
        {
            lock (_sync) return _jobs.ToArray();
        }
    }

    public FileOperationJob Enqueue(
        FileOperationKind kind,
        IEnumerable<string> sources,
        string? destination,
        FileOperationOptions? options = null,
        Func<ConflictInfo, ConflictAnswer>? ask = null)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(FileOperationQueue));

        var job = new FileOperationJob
        {
            Kind = kind,
            Sources = sources.ToArray(),
            Destination = destination,
            Options = options ?? new FileOperationOptions(),
            Ask = ask
        };

        lock (_sync) _jobs.Add(job);
        Raise(job);
        _ = Task.Run(() => RunAsync(job));
        return job;
    }

    public FileOperationJob? Find(string jobId)
    {
        lock (_sync) return _jobs.FirstOrDefault(j => j.Id == jobId);
    }

    public bool Pause(string jobId)
    {
        var job = Find(jobId);
        if (job == null || job.State != FileOperationState.Running) return false;
        job.Pause.Pause();
        job.State = FileOperationState.Paused;
        Raise(job);
        return true;
    }

    public bool Resume(string jobId)
    {
        var job = Find(jobId);
        if (job == null || job.State != FileOperationState.Paused) return false;
        job.State = FileOperationState.Running;
        job.Pause.Resume();
        Raise(job);
        return true;
    }

    public bool Cancel(string jobId)
    {
        var job = Find(jobId);
        if (job == null) return false;
        job.Cancellation.Cancel();
        job.Pause.Resume();
        if (job.State is FileOperationState.Pending)
        {
            Finish(job, FileOperationState.Cancelled);
        }
        return true;
    }

    public void ClearFinished()
    {
        lock (_sync)
        {
            _jobs.RemoveAll(j => j.State is FileOperationState.Completed
                or FileOperationState.Failed or FileOperationState.Cancelled);
        }
    }

    public Task WaitAsync(FileOperationJob job) => job.Finished.Task;

    /// <summary>等待所有任务结束（测试与关窗排水用）。</summary>
    public async Task WhenIdleAsync(TimeSpan? timeout = null)
    {
        var deadline = DateTime.UtcNow + (timeout ?? TimeSpan.FromMinutes(10));
        while (DateTime.UtcNow < deadline)
        {
            bool busy;
            lock (_sync)
            {
                busy = _jobs.Any(j => j.State is FileOperationState.Pending
                    or FileOperationState.Running or FileOperationState.Paused);
            }
            if (!busy) return;
            await Task.Delay(20).ConfigureAwait(false);
        }
    }

    private async Task RunAsync(FileOperationJob job)
    {
        await _slots.WaitAsync().ConfigureAwait(false);
        try
        {
            if (job.Cancellation.IsCancellationRequested)
            {
                Finish(job, FileOperationState.Cancelled);
                return;
            }

            job.State = FileOperationState.Running;
            job.StartedAt = DateTime.Now;
            job.RunOrder = Interlocked.Increment(ref _runSequence);
            Raise(job);

            var measured = FileTransferEngine.Measure(job.Sources);
            job.TotalItems = measured.Files;
            job.TotalBytes = measured.Bytes;
            Raise(job);

            foreach (var source in job.Sources)
            {
                job.Cancellation.Token.ThrowIfCancellationRequested();
                job.Pause.WaitIfPaused(job.Cancellation.Token);

                job.CurrentItem = source;
                Raise(job);

                try
                {
                    switch (job.Kind)
                    {
                        case FileOperationKind.Copy:
                            FileTransferEngine.CopyPath(source, job.Destination!,
                                job.Options, job.Ask, job.AddBytes, job.Pause, job.Cancellation.Token);
                            break;
                        case FileOperationKind.Move:
                            FileTransferEngine.MovePath(source, job.Destination!,
                                job.Options, job.Ask, job.AddBytes, job.Pause, job.Cancellation.Token);
                            break;
                        case FileOperationKind.Delete:
                            FileTransferEngine.DeletePath(source, job.Pause, job.Cancellation.Token);
                            break;
                    }
                    job.AddItem();
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    job.FailedItems++;
                    job.Errors.Add(source + ": " + ex.Message);
                }
            }

            if (job.FailedItems > 0)
            {
                job.Error = "有 " + job.FailedItems + " 项失败";
                Finish(job, FileOperationState.Failed);
            }
            else
            {
                Finish(job, FileOperationState.Completed);
            }
        }
        catch (OperationCanceledException)
        {
            Finish(job, FileOperationState.Cancelled);
        }
        catch (Exception ex)
        {
            job.Error = ex.Message;
            job.Errors.Add(ex.Message);
            Finish(job, FileOperationState.Failed);
        }
        finally
        {
            _slots.Release();
        }
    }

    private void Finish(FileOperationJob job, FileOperationState state)
    {
        job.State = state;
        job.FinishedAt = DateTime.Now;
        job.CurrentItem = null;
        Raise(job);
        job.Finished.TrySetResult(true);
    }

    private void Raise(FileOperationJob job)
    {
        var handlers = JobChanged;
        if (handlers == null) return;
        foreach (EventHandler<FileOperationJob> handler in handlers.GetInvocationList())
        {
            try { handler(this, job); } catch { }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        lock (_sync)
        {
            foreach (var job in _jobs)
            {
                try { job.Cancellation.Cancel(); } catch { }
                job.Pause.Resume();
            }
        }
        _slots.Dispose();
    }
}
