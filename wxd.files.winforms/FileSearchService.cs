using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace wxd.files.winforms;

/// <summary>一次搜索的查询条件。</summary>
public sealed class FileSearchQuery
{
    public IReadOnlyList<string> Roots { get; init; } = Array.Empty<string>();

    /// <summary>文件名条件；空表示不按名称过滤。</summary>
    public string Pattern { get; init; } = "";

    public bool UseRegex { get; init; }
    public bool CaseSensitive { get; init; }
    public bool IncludeSubdirectories { get; init; } = true;
    public bool IncludeHidden { get; init; }

    /// <summary>内容条件；非空时改为内容搜索，命中行逐条返回。</summary>
    public string ContentPattern { get; init; } = "";

    public bool ContentRegex { get; init; }

    /// <summary>0 表示不限。</summary>
    public long MinSize { get; init; }
    public long MaxSize { get; init; }

    public DateTime? ModifiedAfter { get; init; }
    public DateTime? ModifiedBefore { get; init; }

    public int MaxResults { get; init; } = 5000;
    public int MaxContentHitsPerFile { get; init; } = 3;
    public long MaxContentFileBytes { get; init; } = 32L * 1024 * 1024;
}

/// <summary>一条搜索命中；LineNumber 为 0 表示按文件名命中。</summary>
public sealed class FileSearchHit
{
    public string Path { get; init; } = "";
    public bool IsDirectory { get; init; }
    public long Size { get; init; }
    public DateTime ModifiedUtc { get; init; }
    public int LineNumber { get; init; }
    public string? Preview { get; init; }

    public string Name => System.IO.Path.GetFileName(Path);
    public string Directory => System.IO.Path.GetDirectoryName(Path) ?? "";
}

public sealed class FileSearchResult
{
    public List<FileSearchHit> Hits { get; } = new();
    public List<string> Errors { get; } = new();
    public int FilesScanned { get; set; }
    public bool Truncated { get; set; }
}

public sealed class FileSearchProgress
{
    public int Scanned { get; init; }
    public int Matched { get; init; }
    public string? Current { get; init; }
}

/// <summary>递归文件搜索：文件名通配/正则、内容行搜索、大小与修改时间过滤，可取消。</summary>
public static class FileSearchService
{
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);

    public static Task<FileSearchResult> SearchAsync(
        FileSearchQuery query,
        IProgress<FileSearchProgress>? progress,
        CancellationToken token)
    {
        return Task.Run(() => Search(query, progress, token), token);
    }

    /// <summary>同步执行搜索（测试与后台线程共用）。</summary>
    public static FileSearchResult Search(
        FileSearchQuery query,
        IProgress<FileSearchProgress>? progress,
        CancellationToken token)
    {
        var result = new FileSearchResult();
        var nameRegex = string.IsNullOrEmpty(query.Pattern)
            ? null
            : BuildNameRegex(query.Pattern, query.UseRegex, query.CaseSensitive);
        var contentRegex = string.IsNullOrEmpty(query.ContentPattern)
            ? null
            : BuildContentRegex(query.ContentPattern, query.ContentRegex, query.CaseSensitive);

        int scanned = 0;
        var lastReport = DateTime.UtcNow;

        foreach (var root in query.Roots)
        {
            token.ThrowIfCancellationRequested();

            if (File.Exists(root))
            {
                scanned++;
                ConsiderFile(root, query, nameRegex, contentRegex, result, token, progress, ref scanned, ref lastReport);
                continue;
            }

            if (!Directory.Exists(root))
            {
                result.Errors.Add("目录不存在: " + root);
                continue;
            }

            var options = new EnumerationOptions
            {
                RecurseSubdirectories = query.IncludeSubdirectories,
                IgnoreInaccessible = true,
                ReturnSpecialDirectories = false,
                AttributesToSkip = query.IncludeHidden
                    ? 0
                    : FileAttributes.Hidden | FileAttributes.System
            };

            try
            {
                // 内容搜索只返回命中行，目录没有内容可匹配
                if (contentRegex == null)
                {
                    foreach (var directory in Directory.EnumerateDirectories(root, "*", options))
                    {
                        token.ThrowIfCancellationRequested();
                        scanned++;

                        if (nameRegex != null && !nameRegex.IsMatch(Path.GetFileName(directory))) continue;
                        if (!WithinDateRange(directory, query)) continue;

                        if (!AddHit(result, query, new FileSearchHit
                            {
                                Path = directory,
                                IsDirectory = true,
                                ModifiedUtc = SafeModifiedUtc(directory)
                            }))
                        {
                            result.FilesScanned = scanned;
                            return result;
                        }
                    }
                }

                foreach (var file in Directory.EnumerateFiles(root, "*", options))
                {
                    token.ThrowIfCancellationRequested();
                    scanned++;
                    ConsiderFile(file, query, nameRegex, contentRegex, result, token, progress, ref scanned, ref lastReport);
                    if (result.Truncated)
                    {
                        result.FilesScanned = scanned;
                        return result;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                result.Errors.Add(root + ": " + ex.Message);
            }
        }

        result.FilesScanned = scanned;
        progress?.Report(new FileSearchProgress
        {
            Scanned = scanned,
            Matched = result.Hits.Count,
            Current = null
        });
        return result;
    }

    /// <summary>文件名条件编译为锚定正则；非正则模式下 * 和 ? 生效，不含通配符时按包含匹配。</summary>
    public static Regex BuildNameRegex(string pattern, bool useRegex, bool caseSensitive)
    {
        var options = RegexOptions.CultureInvariant
            | (caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);

        if (useRegex) return new Regex(pattern, options, RegexTimeout);

        var body = Regex.Escape(pattern).Replace("\\*", ".*").Replace("\\?", ".");
        if (!pattern.Contains('*') && !pattern.Contains('?'))
            body = ".*" + body + ".*";

        return new Regex("^" + body + "$", options, RegexTimeout);
    }

    public static Regex BuildContentRegex(string pattern, bool useRegex, bool caseSensitive)
    {
        var options = RegexOptions.CultureInvariant
            | (caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);
        return new Regex(useRegex ? pattern : Regex.Escape(pattern), options, RegexTimeout);
    }

    private static void ConsiderFile(
        string file,
        FileSearchQuery query,
        Regex? nameRegex,
        Regex? contentRegex,
        FileSearchResult result,
        CancellationToken token,
        IProgress<FileSearchProgress>? progress,
        ref int scanned,
        ref DateTime lastReport)
    {
        try
        {
            var attributes = File.GetAttributes(file);
            if (!query.IncludeHidden && (attributes & (FileAttributes.Hidden | FileAttributes.System)) != 0)
                return;

            var info = new FileInfo(file);
            if (query.MinSize > 0 && info.Length < query.MinSize) return;
            if (query.MaxSize > 0 && info.Length > query.MaxSize) return;
            if (!WithinDateRange(file, query)) return;

            if (nameRegex != null && !nameRegex.IsMatch(info.Name)) return;

            if (contentRegex != null)
            {
                if (info.Length > query.MaxContentFileBytes) return;
                if (LooksBinary(file)) return;
                SearchContent(file, info, contentRegex, query, result, token);
            }
            else
            {
                AddHit(result, query, new FileSearchHit
                {
                    Path = file,
                    Size = info.Length,
                    ModifiedUtc = info.LastWriteTimeUtc
                });
            }

            if (DateTime.UtcNow - lastReport > TimeSpan.FromMilliseconds(120) || scanned % 200 == 0)
            {
                lastReport = DateTime.UtcNow;
                progress?.Report(new FileSearchProgress
                {
                    Scanned = scanned,
                    Matched = result.Hits.Count,
                    Current = info.Name
                });
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            // 单个文件读不到不影响其余结果
        }
    }

    private static void SearchContent(
        string file,
        FileInfo info,
        Regex regex,
        FileSearchQuery query,
        FileSearchResult result,
        CancellationToken token)
    {
        Encoding encoding;
        try { encoding = TextFileReader.DetectEncoding(file); }
        catch { encoding = new UTF8Encoding(false); }

        using var reader = new StreamReader(file, encoding, detectEncodingFromByteOrderMarks: true,
            bufferSize: 16 * 1024);
        string? line;
        int lineNumber = 0;
        int fileHits = 0;

        while ((line = reader.ReadLine()) != null)
        {
            lineNumber++;
            if ((lineNumber & 0x3FF) == 0) token.ThrowIfCancellationRequested();

            if (!regex.IsMatch(line)) continue;

            if (!AddHit(result, query, new FileSearchHit
                {
                    Path = file,
                    Size = info.Length,
                    ModifiedUtc = info.LastWriteTimeUtc,
                    LineNumber = lineNumber,
                    Preview = TrimPreview(line)
                })) return;

            fileHits++;
            if (fileHits >= query.MaxContentHitsPerFile) return;
        }
    }

    private static bool AddHit(FileSearchResult result, FileSearchQuery query, FileSearchHit hit)
    {
        if (result.Hits.Count >= query.MaxResults)
        {
            result.Truncated = true;
            return false;
        }
        result.Hits.Add(hit);
        return true;
    }

    private static bool WithinDateRange(string path, FileSearchQuery query)
    {
        if (query.ModifiedAfter == null && query.ModifiedBefore == null) return true;
        var modified = SafeModifiedUtc(path);
        if (modified == default) return false;
        if (query.ModifiedAfter is { } after && modified < after.ToUniversalTime()) return false;
        if (query.ModifiedBefore is { } before && modified > before.ToUniversalTime()) return false;
        return true;
    }

    private static DateTime SafeModifiedUtc(string path)
    {
        try { return File.GetLastWriteTimeUtc(path); }
        catch { return default; }
    }

    private static bool LooksBinary(string file)
    {
        try
        {
            using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            Span<byte> buffer = stackalloc byte[4096];
            int read = stream.Read(buffer);
            for (int i = 0; i < read; i++)
            {
                if (buffer[i] == 0) return true;
            }
        }
        catch
        {
            return true;
        }
        return false;
    }

    private static string TrimPreview(string line)
    {
        var text = line.Trim();
        return text.Length <= 240 ? text : text[..240] + "…";
    }
}
