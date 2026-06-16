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
        /// <summary>
        /// 从文件/目录路径向上查找，判断版本控制类型。
        /// 如果传入的是目录路径，会先从该目录本身开始检查（不会跳过它）。
        /// </summary>
        public static VcsType DetectVcs(string path)
        {
            // 先确保从目录本身开始检查（Path.GetDirectoryName 会去掉目录名）
            var dir = Directory.Exists(path) ? path : Path.GetDirectoryName(path);
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
        public static string GetVcsRoot(string path, VcsType vcs)
        {
            var marker = vcs == VcsType.Svn ? ".svn" : ".git";
            var dir = Directory.Exists(path) ? path : Path.GetDirectoryName(path);
            while (!string.IsNullOrEmpty(dir))
            {
                if (Directory.Exists(Path.Combine(dir, marker)))
                    return dir;
                var parent = Directory.GetParent(dir);
                if (parent == null) break;
                dir = parent.FullName;
            }
            return Directory.Exists(path) ? path : (Path.GetDirectoryName(path) ?? "");
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
        private static string? FindSvnExe()
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

        /// <summary>检查 SVN 命令行是否可用（运行 svn --version 验证）</summary>
        private static string? CheckSvnAvailable()
        {
            var svnExe = FindSvnExe();
            if (svnExe == null)
                return "未找到 SVN 命令行工具（svn.exe）。请安装 TortoiseSVN（需勾选命令行工具）或 SlikSVN，并确保 svn.exe 在 PATH 中。";

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = svnExe,
                    Arguments = "--version",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var process = Process.Start(psi);
                if (process == null)
                    return $"无法启动 SVN 命令行工具（{svnExe}）。请检查安装是否正确。";
                process.WaitForExit(5000);
                if (process.ExitCode != 0)
                {
                    var err = process.StandardError.ReadToEnd();
                    return $"SVN 命令行工具异常（{svnExe}）：{err}";
                }
                return null; // 可用
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // svn 命令不在 PATH 中（"系统找不到指定的文件"）
                return "本功能需要安装 SVN 命令行工具。请安装 TortoiseSVN（安装时需勾选 \"command line client tools\"）或 SlikSVN，并确保 svn.exe 在 PATH 中。";
            }
            catch (DirectoryNotFoundException)
            {
                return "文件所在目录不存在，文件可能已被移动或删除。";
            }
            catch (Exception ex)
            {
                return $"SVN 命令行工具异常：{ex.Message}";
            }
        }

        /// <summary>检查 Git 命令行是否可用（运行 git --version 验证）</summary>
        private static string? CheckGitAvailable()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "git",
                    Arguments = "--version",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var process = Process.Start(psi);
                if (process == null)
                    return "无法启动 Git 命令行工具（git）。请确认 Git 是否已正确安装。";
                process.WaitForExit(5000);
                if (process.ExitCode != 0)
                {
                    var err = process.StandardError.ReadToEnd();
                    return $"Git 命令行工具异常：{err}";
                }
                return null; // 可用
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // git 命令不在 PATH 中（"系统找不到指定的文件"）
                return "本功能需要安装 Git 命令行工具。请安装 Git for Windows（https://git-scm.com/）并确保 git.exe 在 PATH 中。";
            }
            catch (DirectoryNotFoundException)
            {
                return "文件所在目录不存在，文件可能已被移动或删除。";
            }
            catch (Exception ex)
            {
                return $"Git 命令行工具异常：{ex.Message}";
            }
        }

        /// <summary>获取 SVN 历史记录</summary>
        public static async Task<VersionHistoryResult> GetSvnHistoryAsync(
            string filePath, int maxCount = 100)
        {
            var result = new VersionHistoryResult();

            // 先检查 SVN 命令行是否可用
            var svnError = CheckSvnAvailable();
            if (svnError != null)
            {
                result.ErrorMessage = svnError;
                return result;
            }

            var svnExe = FindSvnExe()!;
            var root = GetVcsRoot(filePath, VcsType.Svn);

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = svnExe,
                    Arguments = $"log --xml -l {maxCount} \"{filePath}\"",
                    WorkingDirectory = root,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };

                using var process = Process.Start(psi);
                if (process == null)
                {
                    result.ErrorMessage = $"无法启动 SVN 进程（{svnExe}）。";
                    return result;
                }

                // 同时读取 stdout 和 stderr，避免死锁
                var outputTask = process.StandardOutput.ReadToEndAsync();
                var stderrTask = process.StandardError.ReadToEndAsync();

                var output = await outputTask;
                var stderr = await stderrTask;
                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                {
                    result.ErrorMessage = $"SVN log 命令执行失败 (exit={process.ExitCode})：{stderr}";
                    return result;
                }

                result.Entries = ParseSvnLogXml(output);

                // 诊断：解析结果为空但 stdout 有内容，说明 XML 解析失败
                if (result.Entries.Count == 0 && !string.IsNullOrWhiteSpace(output))
                {
                    result.ErrorMessage = "SVN log 解析失败：返回的 XML 无法解析为历史记录。请检查文件路径是否正确。";
                    System.Diagnostics.Debug.WriteLine(
                        $"SVN log parse warning: got {output.Length} chars but parsed 0 entries. " +
                        $"Output preview: {output.Substring(0, Math.Min(output.Length, 500))}");
                    return result;
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                result.ErrorMessage = "本功能需要安装 SVN 命令行工具。请安装 TortoiseSVN（安装时需勾选 \"command line client tools\"）或 SlikSVN，并确保 svn.exe 在 PATH 中。";
                System.Diagnostics.Debug.WriteLine("SVN log error: svn.exe not found");
            }
            catch (DirectoryNotFoundException)
            {
                result.ErrorMessage = "文件所在目录不存在，文件可能已被移动或删除。";
                System.Diagnostics.Debug.WriteLine("SVN log error: directory not found");
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"SVN 历史记录获取失败：{ex.Message}";
                System.Diagnostics.Debug.WriteLine($"SVN log error: {ex.Message}");
            }

            return result;
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
                    // SVN 输出 ISO 8601 格式，支持多种精度：
                    //   6位微秒：2026-06-05T08:58:09.534829Z
                    //   3位毫秒：2026-05-08T03:56:51.000Z
                    //   无小数秒：2026-03-23T09:36:57Z
                    if (!DateTime.TryParseExact(dateStr,
                            new[] {
                                "yyyy-MM-ddTHH:mm:ss.ffffffZ",
                                "yyyy-MM-ddTHH:mm:ss.fffZ",
                                "yyyy-MM-ddTHH:mm:ssZ"
                            },
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                            out commitDate))
                    {
                        // 兜底：标准 DateTime 解析，覆盖其他 ISO 8601 变体
                        if (!DateTime.TryParse(dateStr, CultureInfo.InvariantCulture,
                                DateTimeStyles.AdjustToUniversal,
                                out commitDate))
                        {
                            commitDate = DateTime.MinValue;
                        }
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
        public static async Task<VersionHistoryResult> GetGitHistoryAsync(
            string filePath, int maxCount = 100)
        {
            var result = new VersionHistoryResult();

            // 先检查 Git 命令行是否可用
            var gitError = CheckGitAvailable();
            if (gitError != null)
            {
                result.ErrorMessage = gitError;
                return result;
            }

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
                if (process == null)
                {
                    result.ErrorMessage = "无法启动 Git 进程。";
                    return result;
                }

                var output = await process.StandardOutput.ReadToEndAsync();
                var stderr = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                {
                    result.ErrorMessage = $"Git log 命令执行失败 (exit={process.ExitCode})：{stderr}";
                    return result;
                }

                var lines = output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                int index = 0;
                foreach (var line in lines)
                {
                    var parts = line.Split('|');
                    if (parts.Length >= 5)
                    {
                        result.Entries.Add(new VersionHistoryEntry
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
            catch (System.ComponentModel.Win32Exception)
            {
                result.ErrorMessage = "本功能需要安装 Git 命令行工具。请安装 Git for Windows（https://git-scm.com/）并确保 git.exe 在 PATH 中。";
                System.Diagnostics.Debug.WriteLine("Git log error: git.exe not found");
            }
            catch (DirectoryNotFoundException)
            {
                result.ErrorMessage = "文件所在目录不存在，文件可能已被移动或删除。";
                System.Diagnostics.Debug.WriteLine("Git log error: directory not found");
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Git 历史记录获取失败：{ex.Message}";
                System.Diagnostics.Debug.WriteLine($"Git log error: {ex.Message}");
            }

            return result;
        }

        /// <summary>导出指定版本的文件内容到临时文件</summary>
        /// <returns>导出结果（带错误信息）</returns>
        public static async Task<ExportRevisionResult> ExportRevisionAsync(
            string filePath, VcsType vcs, string revision, string tempDir)
        {
            var result = new ExportRevisionResult();

            // 检查命令行工具是否可用
            if (vcs == VcsType.Svn)
            {
                var svnError = CheckSvnAvailable();
                if (svnError != null)
                {
                    result.ErrorMessage = svnError;
                    return result;
                }
            }
            else
            {
                var gitError = CheckGitAvailable();
                if (gitError != null)
                {
                    result.ErrorMessage = gitError;
                    return result;
                }
            }

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
                if (process == null)
                {
                    result.ErrorMessage = $"无法启动 {(vcs == VcsType.Svn ? "SVN" : "Git")} 导出进程。";
                    return result;
                }

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
                    result.ErrorMessage = $"{(vcs == VcsType.Svn ? "SVN" : "Git")} 导出失败 (exit={process.ExitCode})：{err}";
                    System.Diagnostics.Debug.WriteLine(
                        $"Export error (exit={process.ExitCode}): {err}");
                    try { File.Delete(tempFile); } catch { }
                    return result;
                }

                if (new FileInfo(tempFile).Length > 0)
                {
                    result.TempFile = tempFile;
                    return result;
                }

                // 空文件也算失败
                result.ErrorMessage = $"导出的文件为空：{revision}:{relativePath}";
                System.Diagnostics.Debug.WriteLine(
                    $"Export error: empty output for {revision}:{relativePath}");
                try { File.Delete(tempFile); } catch { }
                return result;
            }
            catch (System.ComponentModel.Win32Exception)
            {
                var tool = vcs == VcsType.Svn ? "SVN" : "Git";
                var installTip = vcs == VcsType.Svn
                    ? "请安装 TortoiseSVN（安装时需勾选 \"command line client tools\"）或 SlikSVN，并确保 svn.exe 在 PATH 中。"
                    : "请安装 Git for Windows（https://git-scm.com/）并确保 git.exe 在 PATH 中。";
                result.ErrorMessage = $"本功能需要安装 {tool} 命令行工具。{installTip}";
                System.Diagnostics.Debug.WriteLine($"Export error: {tool.ToLower()}.exe not found");
                try { File.Delete(tempFile); } catch { }
                return result;
            }
            catch (DirectoryNotFoundException)
            {
                result.ErrorMessage = "文件所在目录不存在，文件可能已被移动或删除。";
                System.Diagnostics.Debug.WriteLine("Export error: directory not found");
                try { File.Delete(tempFile); } catch { }
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"{(vcs == VcsType.Svn ? "SVN" : "Git")} 导出失败：{ex.Message}";
                System.Diagnostics.Debug.WriteLine($"Export error: {ex.Message}");
                try { File.Delete(tempFile); } catch { }
                return result;
            }
        }

        /// <summary>VCS 更新结果</summary>
        public class VcsUpdateResult
        {
            /// <summary>是否成功启动进程</summary>
            public bool Started { get; set; }

            /// <summary>进程退出码（-1 表示未启动）</summary>
            public int ExitCode { get; set; } = -1;

            /// <summary>标准输出</summary>
            public string Output { get; set; } = "";

            /// <summary>错误输出</summary>
            public string Error { get; set; } = "";

            /// <summary>错误消息（工具不可用时设置）</summary>
            public string? ErrorMessage { get; set; }

            /// <summary>检测到的冲突行列表</summary>
            public List<string> ConflictLines { get; set; } = new();

            /// <summary>是否有冲突</summary>
            public bool HasConflicts => ConflictLines.Count > 0;

            /// <summary>是否成功完成（ExitCode == 0）</summary>
            public bool Success => ExitCode == 0;
        }

        /// <summary>执行 VCS 更新操作（SVN update 或 Git pull）</summary>
        /// <param name="directory">要更新的目录路径</param>
        /// <param name="vcs">版本控制类型（Svn 或 Git）</param>
        /// <returns>更新结果</returns>
        public static async Task<VcsUpdateResult> UpdateAsync(string directory, VcsType vcs)
        {
            var result = new VcsUpdateResult();

            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
            {
                result.ErrorMessage = "指定的目录不存在。";
                return result;
            }

            try
            {
                if (vcs == VcsType.Svn)
                {
                    var svnExe = FindSvnExe();
                    if (svnExe == null)
                    {
                        result.ErrorMessage = "未找到 SVN 命令行工具（svn.exe）。请安装 TortoiseSVN（需勾选命令行工具）或 SlikSVN，并确保 svn.exe 在 PATH 中。";
                        return result;
                    }

                    var psi = new ProcessStartInfo
                    {
                        FileName = svnExe,
                        Arguments = $"update \"{directory}\" --non-interactive",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        StandardOutputEncoding = Encoding.UTF8,
                        StandardErrorEncoding = Encoding.UTF8
                    };

                    using var process = Process.Start(psi);
                    if (process == null)
                    {
                        result.ErrorMessage = $"无法启动 SVN 进程（{svnExe}）。";
                        return result;
                    }

                    var output = await process.StandardOutput.ReadToEndAsync();
                    var error = await process.StandardError.ReadToEndAsync();
                    await process.WaitForExitAsync();

                    result.Started = true;
                    result.ExitCode = process.ExitCode;
                    result.Output = output;
                    result.Error = error;
                }
                else if (vcs == VcsType.Git)
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = "git",
                        Arguments = $"pull",
                        WorkingDirectory = directory,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        StandardOutputEncoding = Encoding.UTF8,
                        StandardErrorEncoding = Encoding.UTF8
                    };

                    using var process = Process.Start(psi);
                    if (process == null)
                    {
                        result.ErrorMessage = "无法启动 Git 进程。";
                        return result;
                    }

                    var output = await process.StandardOutput.ReadToEndAsync();
                    var error = await process.StandardError.ReadToEndAsync();
                    await process.WaitForExitAsync();

                    result.Started = true;
                    result.ExitCode = process.ExitCode;
                    result.Output = output;
                    result.Error = error;
                }
                else
                {
                    result.ErrorMessage = "不支持的版本控制类型。";
                    return result;
                }

                // 检测冲突
                if (result.Started)
                {
                    var combined = result.Output + GlobalConst.Lf + result.Error;
                    result.ConflictLines = combined
                        .Split(new[] { GlobalConst.Lf, GlobalConst.Cr }, StringSplitOptions.RemoveEmptyEntries)
                        .Where(line => line.TrimStart().StartsWith("C ") ||
                                       line.Contains("Conflict discovered") ||
                                       line.Contains("conflict", StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                var toolName = vcs == VcsType.Svn ? "SVN" : "Git";
                var installTip = vcs == VcsType.Svn
                    ? "请安装 TortoiseSVN（安装时需勾选 \"command line client tools\"）或 SlikSVN，并确保 svn.exe 在 PATH 中。"
                    : "请安装 Git for Windows（https://git-scm.com/）并确保 git.exe 在 PATH 中。";
                result.ErrorMessage = $"本功能需要安装 {toolName} 命令行工具。{installTip}";
            }
            catch (DirectoryNotFoundException)
            {
                result.ErrorMessage = "目录不存在，可能已被移动或删除。";
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"更新操作异常：{ex.Message}";
            }

            return result;
        }
    }
}
