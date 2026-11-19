using System;
using System.Collections.Generic;
using System.IO;

namespace wxd.files.winforms;

/// <summary>窗格内新建文件/文件夹：名称校验与实际创建，逻辑独立可测。</summary>
internal static class NewItemService
{
    private static readonly HashSet<string> ReservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
    };

    /// <summary>返回 null 表示名称可用；否则返回中文原因。</summary>
    public static string? Validate(string directory, string? rawName)
    {
        var name = (rawName ?? "").Trim();
        if (name.Length == 0) return "名称不能为空";
        if (name.Length > 255) return "名称不能超过 255 个字符";
        if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            return "名称不能包含 \\ / : * ? \" < > | 等字符";
        if (name.EndsWith(' ') || name.EndsWith('.'))
            return "名称不能以空格或点结尾";

        var stem = Path.GetFileNameWithoutExtension(name);
        if (ReservedNames.Contains(stem)) return "该名称是 Windows 保留名";

        var full = Path.Combine(directory, name);
        if (File.Exists(full) || Directory.Exists(full)) return "已存在同名文件或文件夹";
        return null;
    }

    public static bool TryCreateFile(string directory, string? rawName, out string path, out string? error)
    {
        var name = (rawName ?? "").Trim();
        path = Path.Combine(directory, name);
        error = Validate(directory, name);
        if (error != null) return false;

        try
        {
            using (File.Create(path)) { }
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    /// <summary>按扩展名创建：.xlsx/.docx 生成最小 OOXML，其余为空文件。</summary>
    public static bool TryCreateTypedFile(string directory, string? rawName, out string path, out string? error)
    {
        var name = (rawName ?? "").Trim();
        path = Path.Combine(directory, name);
        error = Validate(directory, name);
        if (error != null) return false;

        try
        {
            var ext = Path.GetExtension(name).ToLowerInvariant();
            switch (ext)
            {
                case ".xlsx":
                    OfficeFileTemplates.CreateXlsx(path);
                    break;
                case ".docx":
                    OfficeFileTemplates.CreateDocx(path);
                    break;
                default:
                    using (File.Create(path)) { }
                    break;
            }
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            try { if (File.Exists(path)) File.Delete(path); } catch { }
            return false;
        }
    }

    public static bool TryCreateDirectory(string directory, string? rawName, out string path, out string? error)
    {
        var name = (rawName ?? "").Trim();
        path = Path.Combine(directory, name);
        error = Validate(directory, name);
        if (error != null) return false;

        try
        {
            Directory.CreateDirectory(path);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }
}
