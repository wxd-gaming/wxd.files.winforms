using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace wxd.files.winforms
{
    public enum VcsType { None, Svn, Git }

    public static class VersionControlHelper
    {
        /// <summary>从文件路径向上查找，判断版本控制类型</summary>
        public static VcsType DetectVcs(string filePath)
        {
            var dir = Path.GetDirectoryName(filePath);
            while (!string.IsNullOrEmpty(dir))
            {
                if (Directory.Exists(Path.Combine(dir, ".svn")))
                    return VcsType.Svn;
                if (Directory.Exists(Path.Combine(dir, ".git")))
                    return VcsType.Git;
                var parent = Directory.GetParent(dir);
                if (parent == null) break;
                dir = parent.FullName;
            }
            return VcsType.None;
        }

        /// <summary>获取 VCS 工作目录根路径</summary>
        public static string GetVcsRoot(string filePath, VcsType vcs)
        {
            var marker = vcs == VcsType.Svn ? ".svn" : ".git";
            var dir = Path.GetDirectoryName(filePath);
            while (!string.IsNullOrEmpty(dir))
            {
                if (Directory.Exists(Path.Combine(dir, marker)))
                    return dir;
                var parent = Directory.GetParent(dir);
                if (parent == null) break;
                dir = parent.FullName;
            }
            return Path.GetDirectoryName(filePath) ?? "";
        }

        /// <summary>获取相对路径（相对于 VCS 根目录）</summary>
        public static string GetRelativePath(string filePath, VcsType vcs)
        {
            var root = GetVcsRoot(filePath, vcs);
            var relative = filePath.Substring(root.Length)
                .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return relative.Replace('\\', '/');
        }

        /// <summary>查找 svn.exe 的完整路径</summary>
        private static string FindSvnExe()
        {
            // 常见安装路径
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

            // 最后尝试 PATH 环境变量中的 svn
            return "svn";
        }

        /// <summary>获取 SVN 历史记录</summary>
        public static async Task<List<VersionHistoryEntry>> GetSvnHistoryAsync(
            string filePath, int maxCount = 100)
        {
            var entries = new List<VersionHistoryEntry>();
            var root = GetVcsRoot(filePath, VcsType.Svn);

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = FindSvnExe(),
                    Arguments = $"log --xml -l {maxCount} \"{filePath}\"",
                    WorkingDirectory = root,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };

                using var process = Process.Start(psi);
                if (process == null) return entries;

                // 同时读取 stdout 和 stderr，避免死锁
                var outputTask = process.StandardOutput.ReadToEndAsync();
                var stderrTask = process.StandardError.ReadToEndAsync();

                var output = await outputTask;
                var stderr = await stderrTask;
                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"SVN log error (exit={process.ExitCode}): {stderr}");
                    return entries;
                }

                entries = ParseSvnLogXml(output);

                // 诊断：解析结果为空但 stdout 有内容，说明正则匹配失败
                if (entries.Count == 0 && !string.IsNullOrWhiteSpace(output))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"SVN log parse warning: got {output.Length} chars but parsed 0 entries. " +
                        $"Output preview: {output.Substring(0, Math.Min(output.Length, 500))}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SVN log error: {ex.Message}");
            }

            return entries;
        }

        private static List<VersionHistoryEntry> ParseSvnLogXml(string xml)
        {
            var entries = new List<VersionHistoryEntry>();

            try
            {
                var doc = XDocument.Parse(xml);
                int index = 0;

                foreach (var logentry in doc.Root!.Elements("logentry"))
                {
                    var revision = logentry.Attribute("revision")?.Value ?? "";
                    var author = logentry.Element("author")?.Value ?? "";
                    var dateStr = logentry.Element("date")?.Value ?? "";
                    var message = logentry.Element("msg")?.Value?.Trim() ?? "";

                    DateTime commitDate;
                    // SVN 输出 ISO 8601 格式（如 2026-05-08T03:56:51.000Z）
                    if (!DateTime.TryParseExact(dateStr,
                            new[] { "yyyy-MM-ddTHH:mm:ss.fffZ", "yyyy-MM-ddTHH:mm:ssZ" },
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                            out commitDate))
                    {
                        commitDate = DateTime.MinValue;
                    }

                    entries.Add(new VersionHistoryEntry
                    {
                        Revision = revision,
                        FullRevision = revision,
                        Author = author,
                        CommitDate = commitDate,
                        Message = message,
                        HasFileChange = true,
                        Index = index++
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SVN log XML parse error: {ex.Message}");
            }

            return entries;
        }

        /// <summary>获取 Git 历史记录（只包含当前文件有变更的）</summary>
        public static async Task<List<VersionHistoryEntry>> GetGitHistoryAsync(
            string filePath, int maxCount = 100)
        {
            var entries = new List<VersionHistoryEntry>();
            var root = GetVcsRoot(filePath, VcsType.Git);
            var relativePath = GetRelativePath(filePath, VcsType.Git);

            try
            {
                // Git log: 格式为 完整SHA|短SHA|日期|作者|提交信息
                var psi = new ProcessStartInfo
                {
                    FileName = "git",
                    Arguments = $"log --follow --pretty=format:%H|%h|%ai|%an|%s -{maxCount} -- \"{relativePath}\"",
                    WorkingDirectory = root,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };

                using var process = Process.Start(psi);
                if (process == null) return entries;

                var output = await process.StandardOutput.ReadToEndAsync();
                var stderr = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"Git log error (exit={process.ExitCode}): {stderr}");
                    return entries;
                }

                var lines = output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                int index = 0;
                foreach (var line in lines)
                {
                    var parts = line.Split('|');
                    if (parts.Length >= 5)
                    {
                        entries.Add(new VersionHistoryEntry
                        {
                            FullRevision = parts[0].Trim(),
                            Revision = parts[1].Trim(),
                            CommitDate = DateTime.Parse(parts[2].Trim()),
                            Author = parts[3].Trim(),
                            Message = parts[4].Trim(),
                            HasFileChange = true,
                            Index = index++
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Git log error: {ex.Message}");
            }

            return entries;
        }

        /// <summary>导出指定版本的文件内容到临时文件</summary>
        /// <returns>临时文件的完整路径，失败返回 null</returns>
        public static async Task<string?> ExportRevisionAsync(
            string filePath, VcsType vcs, string revision, string tempDir)
        {
            var root = GetVcsRoot(filePath, vcs);
            var relativePath = GetRelativePath(filePath, vcs);
            var fileName = Path.GetFileName(filePath);
            var tempFile = Path.Combine(tempDir, $"r{revision.Replace('/', '_').Replace('\\', '_')}_{fileName}");

            try
            {
                string arguments;
                if (vcs == VcsType.Svn)
                {
                    arguments = $"cat -r {revision} \"{filePath}\"";
                }
                else
                {
                    // Git: show <commit>:<path> — 路径使用 / 分隔，不要额外加引号
                    // 如果路径含空格，用双引号包裹整个 <commit>:<path> 部分
                    arguments = $"show \"{revision}:{relativePath}\"";
                }

                var psi = new ProcessStartInfo
                {
                    FileName = vcs == VcsType.Svn ? FindSvnExe() : "git",
                    Arguments = arguments,
                    WorkingDirectory = root,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = Process.Start(psi);
                if (process == null) return null;

                // 异步读取 stdout 和 stderr
                var stdoutTask = Task.Run(async () =>
                {
                    using var fs = new FileStream(tempFile, FileMode.Create, FileAccess.Write, FileShare.Read);
                    await process.StandardOutput.BaseStream.CopyToAsync(fs);
                });
                var stderrTask = process.StandardError.ReadToEndAsync();

                await Task.WhenAll(stdoutTask, stderrTask);
                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                {
                    var err = stderrTask.Result;
                    System.Diagnostics.Debug.WriteLine(
                        $"Export error (exit={process.ExitCode}): {err}");
                    try { File.Delete(tempFile); } catch { }
                    return null;
                }

                if (new FileInfo(tempFile).Length > 0)
                    return tempFile;

                // 空文件也算失败
                System.Diagnostics.Debug.WriteLine(
                    $"Export error: empty output for {revision}:{relativePath}");
                try { File.Delete(tempFile); } catch { }
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Export error: {ex.Message}");
                try { File.Delete(tempFile); } catch { }
                return null;
            }
        }
    }
}
