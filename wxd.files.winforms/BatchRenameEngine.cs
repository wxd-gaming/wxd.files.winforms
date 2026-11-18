using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace wxd.files.winforms;

/// <summary>大小写转换方式。</summary>
public enum RenameCaseMode
{
    None,
    Upper,
    Lower,
    Title
}

public enum RenamePlanStatus
{
    Ready,
    Unchanged,
    Conflict,
    Invalid,
    Duplicate
}

/// <summary>批量重命名规则。</summary>
public sealed class RenameRule
{
    public string Find { get; init; } = "";
    public string Replace { get; init; } = "";
    public bool UseRegex { get; init; }
    public bool CaseSensitive { get; init; }

    /// <summary>是否把扩展名一起参与查找替换。</summary>
    public bool IncludeExtensionInFind { get; init; }

    /// <summary>模板；支持 {name} {ext} {counter} {counter:3} {date} {parent}。</summary>
    public string Template { get; init; } = "{name}";

    public int CounterStart { get; init; } = 1;
    public int CounterStep { get; init; } = 1;
    public int CounterDigits { get; init; } = 2;
    public RenameCaseMode CaseMode { get; init; } = RenameCaseMode.None;
}

public sealed class RenamePlanItem
{
    public string SourcePath { get; init; } = "";
    public string TargetName { get; set; } = "";
    public string TargetPath { get; set; } = "";
    public RenamePlanStatus Status { get; set; } = RenamePlanStatus.Ready;
    public string? Message { get; set; }

    public string SourceName => Path.GetFileName(SourcePath);
}

public sealed class RenamePlan
{
    public List<RenamePlanItem> Items { get; } = new();

    public bool HasErrors => Items.Any(i =>
        i.Status is RenamePlanStatus.Conflict or RenamePlanStatus.Invalid or RenamePlanStatus.Duplicate);

    public int ChangeCount => Items.Count(i => i.Status == RenamePlanStatus.Ready);
}

public sealed record RenameJournalEntry(string Source, string Target);

public sealed class RenameApplyResult
{
    public bool Success { get; set; }
    public List<string> Errors { get; } = new();
    public List<RenameJournalEntry> Journal { get; } = new();
}

/// <summary>
/// 批量重命名引擎：先生成完整计划并校验冲突，再用"源→临时名→目标"两阶段执行，
/// 支持交换名、仅大小写变化，失败时回滚。
/// </summary>
public static class BatchRenameEngine
{
    private static readonly Regex CounterToken =
        new(@"\{counter(?::(\d+))?\}", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly HashSet<string> ReservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
    };

    /// <summary>生成计划；不触碰磁盘内容。</summary>
    public static RenamePlan BuildPlan(IEnumerable<string> paths, RenameRule rule)
    {
        var plan = new RenamePlan();

        int index = 0;
        foreach (var path in paths)
        {
            var fullSource = Path.GetFullPath(path);
            var item = new RenamePlanItem { SourcePath = fullSource };

            if (!File.Exists(fullSource))
            {
                item.Status = RenamePlanStatus.Invalid;
                item.Message = Directory.Exists(fullSource) ? "暂不支持批量重命名文件夹" : "文件不存在";
                plan.Items.Add(item);
                index++;
                continue;
            }

            var newName = ComputeName(fullSource, rule, index, out var error);
            if (error != null)
            {
                item.Status = RenamePlanStatus.Invalid;
                item.Message = error;
                item.TargetName = newName ?? "";
                plan.Items.Add(item);
                index++;
                continue;
            }

            item.TargetName = newName!;
            item.TargetPath = Path.Combine(Path.GetDirectoryName(fullSource) ?? "", newName!);

            if (string.Equals(item.TargetPath, fullSource, StringComparison.Ordinal))
            {
                item.Status = RenamePlanStatus.Unchanged;
            }

            plan.Items.Add(item);
            index++;
        }

        // 先标出互相撞名的项，再判断与磁盘上现有文件的冲突
        foreach (var group in plan.Items
                     .Where(i => i.Status is RenamePlanStatus.Ready or RenamePlanStatus.Unchanged)
                     .GroupBy(i => Normalize(i.TargetPath), StringComparer.OrdinalIgnoreCase)
                     .Where(g => g.Count() > 1))
        {
            foreach (var item in group)
            {
                item.Status = RenamePlanStatus.Duplicate;
                item.Message = "多个文件指向同一名称";
            }
        }

        // 只有"本轮真的会被移走"的源文件，才允许成为其他项的目标（交换名场景）
        var movingSources = new HashSet<string>(
            plan.Items.Where(i => i.Status == RenamePlanStatus.Ready)
                .Select(i => Normalize(i.SourcePath)),
            StringComparer.OrdinalIgnoreCase);

        foreach (var item in plan.Items)
        {
            if (item.Status != RenamePlanStatus.Ready) continue;
            if (File.Exists(item.TargetPath) || Directory.Exists(item.TargetPath))
            {
                if (!movingSources.Contains(Normalize(item.TargetPath)))
                {
                    item.Status = RenamePlanStatus.Conflict;
                    item.Message = "目标已存在";
                }
            }
        }

        return plan;
    }

    /// <summary>执行计划；存在冲突/无效项时整体拒绝，避免只改一半。</summary>
    public static RenameApplyResult Apply(RenamePlan plan)
    {
        var result = new RenameApplyResult { Success = false };

        if (plan.HasErrors)
        {
            result.Errors.Add("计划中存在冲突或无效项，未执行任何重命名。");
            return result;
        }

        var items = plan.Items.Where(i => i.Status == RenamePlanStatus.Ready).ToList();
        if (items.Count == 0)
        {
            result.Errors.Add("没有需要重命名的文件。");
            return result;
        }

        var temps = new List<(string Temp, RenamePlanItem Item)>();
        var finalized = new List<string>();

        try
        {
            foreach (var item in items)
            {
                var temp = item.SourcePath + ".wxd-ren-" + Guid.NewGuid().ToString("N")[..8];
                File.Move(FileTransferEngine.X(item.SourcePath), FileTransferEngine.X(temp));
                temps.Add((temp, item));
            }

            foreach (var (temp, item) in temps)
            {
                File.Move(FileTransferEngine.X(temp), FileTransferEngine.X(item.TargetPath));
                result.Journal.Add(new RenameJournalEntry(item.SourcePath, item.TargetPath));
                finalized.Add(item.SourcePath);
            }

            result.Success = true;
            return result;
        }
        catch (Exception ex)
        {
            result.Errors.Add(ex.Message);
            Rollback(items, temps, finalized, result.Errors);
            return result;
        }
    }

    /// <summary>按日志逆向恢复；同样走两阶段，可还原交换名。</summary>
    public static bool Undo(IEnumerable<RenameJournalEntry> journal, out string? error)
    {
        error = null;
        var entries = journal.ToList();
        var temps = new List<(string Temp, RenameJournalEntry Entry)>();

        try
        {
            foreach (var entry in entries)
            {
                if (!File.Exists(entry.Target)) continue;
                var temp = entry.Target + ".wxd-undo-" + Guid.NewGuid().ToString("N")[..8];
                File.Move(FileTransferEngine.X(entry.Target), FileTransferEngine.X(temp));
                temps.Add((temp, entry));
            }

            foreach (var (temp, entry) in temps)
            {
                File.Move(FileTransferEngine.X(temp), FileTransferEngine.X(entry.Source));
            }

            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            foreach (var (temp, entry) in temps)
            {
                try
                {
                    if (File.Exists(temp) && !File.Exists(entry.Source))
                        File.Move(FileTransferEngine.X(temp), FileTransferEngine.X(entry.Source));
                }
                catch
                {
                    // 尽力回滚
                }
            }
            return false;
        }
    }

    /// <summary>按规则算出一个文件的新名称；错误通过 error 返回。</summary>
    public static string? ComputeName(string sourcePath, RenameRule rule, int index, out string? error)
    {
        error = null;

        var fileName = Path.GetFileName(sourcePath);
        var extension = Path.GetExtension(fileName);
        var stem = Path.GetFileNameWithoutExtension(fileName);
        if (string.IsNullOrEmpty(stem)) stem = fileName;

        var input = rule.IncludeExtensionInFind ? fileName : stem;
        var replaced = ApplyFindReplace(input, rule);

        var template = string.IsNullOrEmpty(rule.Template) ? "{name}" : rule.Template;
        var counter = rule.CounterStart + index * rule.CounterStep;

        var templateHasExtension = template.Contains("{ext}", StringComparison.Ordinal);

        var name = template
            .Replace("{parent}", ParentName(sourcePath))
            .Replace("{date}", SafeModifiedDate(sourcePath));

        name = CounterToken.Replace(name, match =>
        {
            int digits = match.Groups[1].Success && int.TryParse(match.Groups[1].Value, out var d)
                ? Math.Clamp(d, 1, 12)
                : Math.Clamp(rule.CounterDigits, 1, 12);
            return counter.ToString("D" + digits, CultureInfo.InvariantCulture);
        });

        name = name.Replace("{ext}", extension.TrimStart('.'));
        name = name.Replace("{name}", replaced);

        if (!templateHasExtension && !rule.IncludeExtensionInFind && extension.Length > 0)
            name += extension;

        if (rule.CaseMode != RenameCaseMode.None)
            name = ApplyCase(name, rule.CaseMode);

        error = ValidateName(name);
        return name;
    }

    private static string ApplyFindReplace(string input, RenameRule rule)
    {
        if (string.IsNullOrEmpty(rule.Find)) return input;

        if (rule.UseRegex)
        {
            var options = RegexOptions.CultureInvariant
                | (rule.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);
            return Regex.Replace(input, rule.Find, rule.Replace, options, TimeSpan.FromSeconds(2));
        }

        return input.Replace(rule.Find, rule.Replace,
            rule.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
    }

    private static string ApplyCase(string name, RenameCaseMode mode)
    {
        var extension = Path.GetExtension(name);
        var stem = string.IsNullOrEmpty(extension) ? name : name[..^extension.Length];

        stem = mode switch
        {
            RenameCaseMode.Upper => stem.ToUpperInvariant(),
            RenameCaseMode.Lower => stem.ToLowerInvariant(),
            RenameCaseMode.Title => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(stem.ToLowerInvariant()),
            _ => stem
        };

        return stem + extension;
    }

    private static string? ValidateName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "名称为空";
        if (name.Length > 255) return "名称超过 255 个字符";
        if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return "名称包含非法字符";
        if (name.EndsWith(' ') || name.EndsWith('.')) return "名称不能以空格或点结尾";

        var stem = Path.GetFileNameWithoutExtension(name);
        if (ReservedNames.Contains(stem)) return "名称是 Windows 保留名";
        return null;
    }

    private static string ParentName(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (string.IsNullOrEmpty(directory)) return "";
        var trimmed = directory.TrimEnd(Path.DirectorySeparatorChar);
        return Path.GetFileName(trimmed);
    }

    private static string SafeModifiedDate(string path)
    {
        try { return File.GetLastWriteTime(path).ToString("yyyyMMdd", CultureInfo.InvariantCulture); }
        catch { return ""; }
    }

    private static string Normalize(string path)
    {
        try { return Path.GetFullPath(path); }
        catch { return path; }
    }

    private static void Rollback(
        List<RenamePlanItem> items,
        List<(string Temp, RenamePlanItem Item)> temps,
        List<string> finalized,
        List<string> errors)
    {
        var finalSet = new HashSet<string>(finalized, StringComparer.OrdinalIgnoreCase);

        foreach (var (temp, item) in temps)
        {
            try
            {
                if (File.Exists(item.TargetPath) && finalSet.Contains(item.SourcePath))
                {
                    File.Move(FileTransferEngine.X(item.TargetPath), FileTransferEngine.X(item.SourcePath));
                }
                else if (File.Exists(temp))
                {
                    File.Move(FileTransferEngine.X(temp), FileTransferEngine.X(item.SourcePath));
                }
            }
            catch (Exception ex)
            {
                errors.Add("回滚失败 " + item.SourcePath + ": " + ex.Message);
            }
        }
    }
}
