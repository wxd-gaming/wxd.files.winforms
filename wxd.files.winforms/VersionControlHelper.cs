using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

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
                    FileName = "svn",
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

                var output = await process.StandardOutput.ReadToEndAsync();
                await process.WaitForExitAsync();

                entries = ParseSvnLogXml(output);
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
            var logEntryPattern = @"<logentry\s+revision=""(\d+)"">\s*<author>([^<]*)</author>\s*<date>([^<]*)</date>\s*<msg>([^<]*)</msg>";
            var matches = Regex.Matches(xml, logEntryPattern, RegexOptions.Singleline);

            int index = 0;
            foreach (Match m in matches)
            {
                entries.Add(new VersionHistoryEntry
                {
                    Revision = m.Groups[1].Value,
                    FullRevision = m.Groups[1].Value,
                    Author = m.Groups[2].Value,
                    CommitDate = DateTime.Parse(m.Groups[3].Value),
                    Message = m.Groups[4].Value.Trim(),
                    HasFileChange = true,
                    Index = index++
                });
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
                    FileName = vcs == VcsType.Svn ? "svn" : "git",
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
