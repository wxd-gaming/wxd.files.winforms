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
        static VersionControlHelper()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

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
                // .git 在 worktree/子模块里是文件而非目录，两者都要认，否则入口被误禁用
                var gitMarker = Path.Combine(dir, ".git");
                if (Directory.Exists(gitMarker) || File.Exists(gitMarker))
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
                var markerPath = Path.Combine(dir, marker);
                // .git 在 worktree/子模块里是文件而非目录，两者都要认
                if (Directory.Exists(markerPath) || (marker == ".git" && File.Exists(markerPath)))
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
            // 用 Path.GetRelativePath 替代 Substring：对大小写不一致/junction/短路径更健壮，避免越界或错切
            var relative = Path.GetRelativePath(root, filePath);
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
        private static async Task<string?> CheckSvnAvailableAsync()
        {
            var svnExe = FindSvnExe();
            if (svnExe == null)
                return "未找到 SVN 命令行工具（svn.exe）。请安装 TortoiseSVN（需勾选命令行工具）或 SlikSVN，并确保 svn.exe 在 PATH 中。";

            try
            {
                var run = await RunProcessAsync(svnExe, "--version", timeoutMs: 10_000).ConfigureAwait(false);
                if (!run.Started)
                {
                    return run.ToolMissing
                        ? "本功能需要安装 SVN 命令行工具。请安装 TortoiseSVN（安装时需勾选 \"command line client tools\"）或 SlikSVN，并确保 svn.exe 在 PATH 中。"
                        : $"无法启动 SVN 命令行工具（{svnExe}）：{run.Error}";
                }
                if (run.TimedOut)
                    return $"SVN 命令行工具响应超时（{svnExe}）。";
                if (run.ExitCode != 0)
                    return $"SVN 命令行工具异常（{svnExe}）：{run.Error.Trim()}";
                return null; // 可用
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
        private static async Task<string?> CheckGitAvailableAsync()
        {
            try
            {
                var run = await RunProcessAsync("git", "--version", timeoutMs: 10_000).ConfigureAwait(false);
                if (!run.Started)
                {
                    return run.ToolMissing
                        ? "本功能需要安装 Git 命令行工具。请安装 Git for Windows（https://git-scm.com/）并确保 git.exe 在 PATH 中。"
                        : $"无法启动 Git 命令行工具（git）：{run.Error}";
                }
                if (run.TimedOut)
                    return "Git 命令行工具响应超时（git）。";
                if (run.ExitCode != 0)
                    return $"Git 命令行工具异常：{run.Error.Trim()}";
                return null; // 可用
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
            var svnError = await CheckSvnAvailableAsync().ConfigureAwait(false);
            if (svnError != null)
            {
                result.ErrorMessage = svnError;
                return result;
            }

            var svnExe = FindSvnExe()!;
            var root = GetVcsRoot(filePath, VcsType.Svn);

            try
            {
                var run = await RunProcessAsync(
                    svnExe,
                    $"log --xml -l {maxCount} --non-interactive \"{filePath}\"",
                    workingDirectory: root,
                    outputEncoding: Encoding.UTF8,
                    timeoutMs: 60_000,
                    disableInteractivePrompts: true).ConfigureAwait(false);

                if (!run.Started)
                {
                    result.ErrorMessage = run.ToolMissing
                        ? "本功能需要安装 SVN 命令行工具。请安装 TortoiseSVN（安装时需勾选 \"command line client tools\"）或 SlikSVN，并确保 svn.exe 在 PATH 中。"
                        : $"无法启动 SVN 进程（{svnExe}）：{run.Error}";
                    return result;
                }

                if (run.TimedOut)
                {
                    result.ErrorMessage = "SVN log 命令超时（60 秒），已终止。请检查网络或凭据配置。";
                    return result;
                }

                if (run.ExitCode != 0)
                {
                    result.ErrorMessage = $"SVN log 命令执行失败 (exit={run.ExitCode})：{run.Error}";
                    return result;
                }

                var output = run.Output;
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
            var gitError = await CheckGitAvailableAsync().ConfigureAwait(false);
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
                var run = await RunProcessAsync(
                    "git",
                    $"log --follow --pretty=format:%H|%h|%ai|%an|%s -{maxCount} -- \"{relativePath}\"",
                    workingDirectory: root,
                    outputEncoding: Encoding.UTF8,
                    timeoutMs: 60_000,
                    disableInteractivePrompts: true).ConfigureAwait(false);

                if (!run.Started)
                {
                    result.ErrorMessage = run.ToolMissing
                        ? "本功能需要安装 Git 命令行工具。请安装 Git for Windows（https://git-scm.com/）并确保 git.exe 在 PATH 中。"
                        : $"无法启动 Git 进程：{run.Error}";
                    return result;
                }

                if (run.TimedOut)
                {
                    result.ErrorMessage = "Git log 命令超时（60 秒），已终止。";
                    return result;
                }

                if (run.ExitCode != 0)
                {
                    result.ErrorMessage = $"Git log 命令执行失败 (exit={run.ExitCode})：{run.Error}";
                    return result;
                }

                result.Entries = ParseGitLogOutput(run.Output);
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

        /// <summary>
        /// 解析 git log --pretty=format:%H|%h|%ai|%an|%s 输出。
        /// 分割限 5 段以保留 subject 内部的 |；日期用 TryParse，单条坏行跳过而非让整段历史解析失败。
        /// </summary>
        internal static List<VersionHistoryEntry> ParseGitLogOutput(string output)
        {
            var entries = new List<VersionHistoryEntry>();
            var lines = output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            int index = 0;
            foreach (var line in lines)
            {
                var parts = line.Split('|', 5);
                if (parts.Length < 5) continue;

                if (!DateTime.TryParse(parts[2].Trim(), CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var commitDate))
                    continue;

                entries.Add(new VersionHistoryEntry
                {
                    FullRevision = parts[0].Trim(),
                    Revision = parts[1].Trim(),
                    CommitDate = commitDate,
                    Author = parts[3].Trim(),
                    Message = parts[4].Trim(),
                    HasFileChange = true,
                    Index = index++
                });
            }
            return entries;
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
                var svnError = await CheckSvnAvailableAsync().ConfigureAwait(false);
                if (svnError != null)
                {
                    result.ErrorMessage = svnError;
                    return result;
                }
            }
            else
            {
                var gitError = await CheckGitAvailableAsync().ConfigureAwait(false);
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
                    arguments = $"cat -r {revision} --non-interactive \"{filePath}\"";
                }
                else
                {
                    // Git: show <commit>:<path> — 路径使用 / 分隔，不要额外加引号
                    // 如果路径含空格，用双引号包裹整个 <commit>:<path> 部分
                    arguments = $"show \"{revision}:{relativePath}\"";
                }

                var run = await RunProcessToFileAsync(
                    vcs == VcsType.Svn ? FindSvnExe()! : "git",
                    arguments,
                    tempFile,
                    workingDirectory: root,
                    timeoutMs: 120_000,
                    disableInteractivePrompts: true).ConfigureAwait(false);

                if (!run.Started)
                {
                    var toolName = vcs == VcsType.Svn ? "SVN" : "Git";
                    result.ErrorMessage = run.ToolMissing
                        ? $"本功能需要安装 {toolName} 命令行工具。"
                        : $"无法启动 {toolName} 导出进程：{run.Error}";
                    try { File.Delete(tempFile); } catch { }
                    return result;
                }

                if (run.TimedOut)
                {
                    result.ErrorMessage = $"{(vcs == VcsType.Svn ? "SVN" : "Git")} 导出超时（120 秒），已终止。";
                    try { File.Delete(tempFile); } catch { }
                    return result;
                }

                if (run.ExitCode != 0)
                {
                    var err = run.Error;
                    result.ErrorMessage = $"{(vcs == VcsType.Svn ? "SVN" : "Git")} 导出失败 (exit={run.ExitCode})：{err}";
                    System.Diagnostics.Debug.WriteLine(
                        $"Export error (exit={run.ExitCode}): {err}");
                    try { File.Delete(tempFile); } catch { }
                    return result;
                }

                // 进程已成功退出(exit=0)：空文件也是合法版本内容（如新建的空CSV），不再误判为失败
                result.TempFile = tempFile;
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

        /// <summary>子进程执行结果（stdout/stderr 并发读取 + 超时）</summary>
        internal readonly struct ProcessRunResult
        {
            public int ExitCode { get; init; }
            public string Output { get; init; }
            public string Error { get; init; }
            public bool TimedOut { get; init; }
            public bool Started { get; init; }
            /// <summary>可执行文件不存在（PATH 中找不到）</summary>
            public bool ToolMissing { get; init; }
        }

        /// <summary>
        /// 运行命令行工具：并发读取 stdout/stderr（避免管道互锁），超时后结束整个进程树。
        /// </summary>
        internal static async Task<ProcessRunResult> RunProcessAsync(
            string fileName,
            string arguments,
            string? workingDirectory = null,
            Encoding? outputEncoding = null,
            int timeoutMs = 30_000,
            bool disableInteractivePrompts = false)
        {
            var encoding = outputEncoding ?? Encoding.UTF8;
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = encoding,
                StandardErrorEncoding = encoding
            };
            if (!string.IsNullOrEmpty(workingDirectory))
                psi.WorkingDirectory = workingDirectory;
            if (disableInteractivePrompts)
            {
                // 凭据/编辑器交互会让命令停在一个看不见的窗口上，直接禁用
                psi.Environment["GIT_TERMINAL_PROMPT"] = "0";
                psi.Environment["GCM_INTERACTIVE"] = "Never";
                psi.Environment["GIT_EDITOR"] = "true";
            }

            using var process = new Process { StartInfo = psi };
            try
            {
                if (!process.Start())
                    return new ProcessRunResult { ExitCode = -1, Output = "", Error = "无法启动进程。", Started = false };
            }
            catch (Exception ex)
            {
                return new ProcessRunResult
                {
                    ExitCode = -1,
                    Output = "",
                    Error = ex.Message,
                    Started = false,
                    ToolMissing = ex is System.ComponentModel.Win32Exception
                };
            }

            // 必须并发读取两个流：只读 stdout 时 stderr 写满管道缓冲会让进程卡死
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();

            bool timedOut = false;
            try
            {
                using var cts = new System.Threading.CancellationTokenSource(timeoutMs);
                await process.WaitForExitAsync(cts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                timedOut = true;
                try { process.Kill(entireProcessTree: true); } catch { }
            }

            var output = await stdoutTask.ConfigureAwait(false);
            var error = await stderrTask.ConfigureAwait(false);
            if (timedOut)
            {
                if (error.Length > 0) error += GlobalConst.Lf;
                error += $"命令超时（{timeoutMs / 1000} 秒），已终止。";
            }

            return new ProcessRunResult
            {
                ExitCode = timedOut ? -1 : process.ExitCode,
                Output = output,
                Error = error,
                TimedOut = timedOut,
                Started = true
            };
        }

        /// <summary>
        /// 运行命令行工具并把 stdout 原样写入文件（二进制安全，用于导出文件内容）。
        /// stderr 并发读取，超时后结束整个进程树。
        /// </summary>
        internal static async Task<ProcessRunResult> RunProcessToFileAsync(
            string fileName,
            string arguments,
            string outputFile,
            string? workingDirectory = null,
            int timeoutMs = 60_000,
            bool disableInteractivePrompts = false)
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                // stdout 走 BaseStream 原样落盘，不能设置 StandardOutputEncoding
                StandardErrorEncoding = Encoding.UTF8
            };
            if (!string.IsNullOrEmpty(workingDirectory))
                psi.WorkingDirectory = workingDirectory;
            if (disableInteractivePrompts)
            {
                psi.Environment["GIT_TERMINAL_PROMPT"] = "0";
                psi.Environment["GCM_INTERACTIVE"] = "Never";
                psi.Environment["GIT_EDITOR"] = "true";
            }

            using var process = new Process { StartInfo = psi };
            try
            {
                if (!process.Start())
                    return new ProcessRunResult { ExitCode = -1, Output = "", Error = "无法启动进程。", Started = false };
            }
            catch (Exception ex)
            {
                return new ProcessRunResult
                {
                    ExitCode = -1,
                    Output = "",
                    Error = ex.Message,
                    Started = false,
                    ToolMissing = ex is System.ComponentModel.Win32Exception
                };
            }

            var copyError = "";
            var copyTask = Task.Run(async () =>
            {
                using var fs = new FileStream(outputFile, FileMode.Create, FileAccess.Write, FileShare.Read);
                await process.StandardOutput.BaseStream.CopyToAsync(fs).ConfigureAwait(false);
            });
            var stderrTask = process.StandardError.ReadToEndAsync();

            bool timedOut = false;
            try
            {
                using var cts = new System.Threading.CancellationTokenSource(timeoutMs);
                await process.WaitForExitAsync(cts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                timedOut = true;
                try { process.Kill(entireProcessTree: true); } catch { }
            }

            try { await copyTask.ConfigureAwait(false); }
            catch (Exception ex) { copyError = ex.Message; }

            var error = await stderrTask.ConfigureAwait(false);
            if (copyError.Length > 0)
            {
                if (error.Length > 0) error += GlobalConst.Lf;
                error += $"写入导出文件失败：{copyError}";
            }
            if (timedOut)
            {
                if (error.Length > 0) error += GlobalConst.Lf;
                error += $"命令超时（{timeoutMs / 1000} 秒），已终止。";
            }

            return new ProcessRunResult
            {
                ExitCode = timedOut ? -1 : process.ExitCode,
                Output = "",
                Error = error,
                TimedOut = timedOut,
                Started = true
            };
        }

        /// <summary>
        /// 构造 svn update 参数：去掉目录尾部分隔符，避免 "C:\foo\" 的 \" 被 Windows argv 解析为转义引号。
        /// </summary>
        internal static string BuildSvnUpdateArguments(string directory)
            => $"update \"{Path.TrimEndingDirectorySeparator(directory)}\" --non-interactive";

        /// <summary>
        /// 从更新输出中筛选明确的冲突行：SVN 的 "C " 标记、Git 的 CONFLICT/合并失败等。
        /// 只认明确模式，避免 "no conflicts" 之类正常输出误报为冲突。
        /// </summary>
        internal static List<string> ParseConflictLines(string combined)
        {
            return combined
                .Split(new[] { GlobalConst.Lf, GlobalConst.Cr }, StringSplitOptions.RemoveEmptyEntries)
                .Where(line => line.TrimStart().StartsWith("C ", StringComparison.Ordinal) ||
                               line.StartsWith("CONFLICT", StringComparison.OrdinalIgnoreCase) ||
                               line.Contains("Conflict discovered", StringComparison.OrdinalIgnoreCase) ||
                               line.Contains("Automatic merge failed", StringComparison.OrdinalIgnoreCase) ||
                               line.StartsWith("Summary of conflicts", StringComparison.OrdinalIgnoreCase))
                .ToList();
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
                ProcessRunResult run;
                if (vcs == VcsType.Svn)
                {
                    var svnExe = FindSvnExe();
                    if (svnExe == null)
                    {
                        result.ErrorMessage = "未找到 SVN 命令行工具（svn.exe）。请安装 TortoiseSVN（需勾选命令行工具）或 SlikSVN，并确保 svn.exe 在 PATH 中。";
                        return result;
                    }

                    run = await RunProcessAsync(svnExe, BuildSvnUpdateArguments(directory),
                        outputEncoding: Encoding.GetEncoding(936),
                        timeoutMs: 120_000,
                        disableInteractivePrompts: true).ConfigureAwait(false);
                    if (!run.Started)
                    {
                        result.ErrorMessage = $"无法启动 SVN 进程（{svnExe}）：{run.Error}";
                        return result;
                    }
                }
                else if (vcs == VcsType.Git)
                {
                    run = await RunProcessAsync("git", "pull --no-edit",
                        workingDirectory: directory,
                        timeoutMs: 180_000,
                        disableInteractivePrompts: true).ConfigureAwait(false);
                    if (!run.Started)
                    {
                        result.ErrorMessage = $"无法启动 Git 进程：{run.Error}";
                        return result;
                    }
                }
                else
                {
                    result.ErrorMessage = "不支持的版本控制类型。";
                    return result;
                }

                result.Started = true;
                result.ExitCode = run.ExitCode;
                result.Output = run.Output;
                result.Error = run.Error;
                if (run.TimedOut)
                    result.ErrorMessage = vcs == VcsType.Svn
                        ? "SVN 更新超时（120 秒），已终止。请检查网络或凭据配置。"
                        : "Git 更新超时（180 秒），已终止。请检查网络或凭据配置。";

                // 检测冲突：只认明确的冲突行，避免 "no conflicts" 之类误报
                result.ConflictLines = ParseConflictLines(result.Output + GlobalConst.Lf + result.Error);
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

        /// <summary>标记冲突已解决（SVN resolved 或 Git add）</summary>
        public static void MarkConflictResolved(string filePath, VcsType vcs)
        {
            try
            {
                if (vcs == VcsType.Svn)
                {
                    var svnExe = FindSvnExe();
                    if (svnExe != null)
                    {
                        RunProcessAsync(svnExe, $"resolved \"{filePath}\" --non-interactive",
                            timeoutMs: 10_000,
                            disableInteractivePrompts: true).GetAwaiter().GetResult();
                    }
                }
                else if (vcs == VcsType.Git)
                {
                    var root = GetVcsRoot(filePath, VcsType.Git);
                    RunProcessAsync("git", $"add \"{filePath}\"",
                        workingDirectory: root,
                        timeoutMs: 10_000,
                        disableInteractivePrompts: true).GetAwaiter().GetResult();
                }
            }
            catch { }
        }
    }
}
