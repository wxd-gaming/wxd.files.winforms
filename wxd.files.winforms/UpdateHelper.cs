using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>检查更新的结果</summary>
    public class CheckResult
    {
        /// <summary>远程版本号</summary>
        public string RemoteVersion { get; set; } = "";
        /// <summary>远程文件的 MD5</summary>
        public string RemoteMd5 { get; set; } = "";
        /// <summary>是否发现新版本</summary>
        public bool HasUpdate { get; set; }
        /// <summary>当前版本号</summary>
        public string CurrentVersion { get; set; } = "";
        /// <summary>更新说明（累积历史版本的 Release body）</summary>
        public string ReleaseNotes { get; set; } = "";
    }

    /// <summary>Release 信息（从 Gitee API 解析）</summary>
    internal class ReleaseInfo
    {
        public string Version { get; set; } = "";
        public string Md5 { get; set; } = "";
        public string? DownloadUrl { get; set; }
        public string Body { get; set; } = "";
    }

    /// <summary>更新核心逻辑</summary>
    public static class UpdateHelper
    {
        // Gitee API（公开仓库无需 token）
        private const string ReleaseLatestApi =
            "https://gitee.com/api/v5/repos/wxd-gaming/wxd.files.winforms/releases/latest";
        private const string ReleasesListApi =
            "https://gitee.com/api/v5/repos/wxd-gaming/wxd.files.winforms/releases?page=1&per_page=50";

        // Config 中的键名
        private const string KeyLastCheckTime = "UpdateLastCheckTime";
        private const string KeySkippedVersions = "UpdateSkippedVersions";
        private const string KeyCurrentVersion = "UpdateCurrentVersion";

        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(8)
        };

        /// <summary>统一检查更新入口（自动/手动都调用此方法）</summary>
        /// <remarks>
        /// 核心逻辑：
        /// 1. 版本号相同 → 无更新（不弹窗，即使 MD5 不同也可能是编译差异）
        /// 2. 版本号不同 + MD5 相同 → 静默更新 Config 版本号，不弹窗
        /// 3. 版本号不同 + MD5 不同或不可比较 → 有更新（弹窗下载文件 + 更新版本号）
        /// </remarks>
        public static async Task<CheckResult?> CheckUpdateAsync()
        {
            var result = await FetchCheckResultAsync();
            if (result == null)
                return null;

            // MD5 相同但版本号不同 → 静默更新版本号，不弹窗
            if (!result.HasUpdate && !string.Equals(result.CurrentVersion, result.RemoteVersion, StringComparison.Ordinal))
            {
                SetCurrentVersion(result.RemoteVersion);
                result.CurrentVersion = result.RemoteVersion;
            }

            // 检查是否在跳过列表里
            if (result.HasUpdate)
            {
                var skipped = GetSkippedVersions();
                if (skipped.Contains(result.RemoteVersion))
                    return null; // 已跳过，不弹窗
            }

            return result;
        }

        /// <summary>启动时静默检查（异步，不抛异常，不影响启动）</summary>
        public static async Task CheckSilentAsync()
        {
            try
            {
                // 判断是否跨天（自然日，不考虑24小时间隔）
                var lastCheck = GetLastCheckTime();
                if (lastCheck != DateTime.MinValue && lastCheck.Date == DateTime.Now.Date)
                    return;

                // 更新检查时间（放在Fetch之前，由外部Timer控制调用时机，不阻塞启动）
                SetLastCheckTime(DateTime.Now);

                var result = await CheckUpdateAsync();
                if (result == null || !result.HasUpdate)
                    return;

                // 有更新 → 使用自定义弹窗（三个按钮：立即更新 / 跳过版本 / 取消）
                var notifyForm = new UpdateNotifyForm(result);
                notifyForm.ShowDialog();

                if (notifyForm.UserChoice == 2)
                {
                    // 跳过此版本
                    AddSkippedVersion(result.RemoteVersion);
                }
                // 1 = 立即更新 → UpdateNotifyForm 内部已处理下载和退出
                // 3 = 取消 → 下次启动时再次提示
            }
            catch
            {
                // 静默失败，不阻塞启动
            }
        }

        /// <summary>手动检查更新（调用方处理异常和 UI 反馈）</summary>
        public static async Task<CheckResult> CheckManualAsync()
        {
            var result = await CheckUpdateAsync();
            if (result == null)
                throw new Exception("无法获取远程版本信息，请检查网络连接。");

            return result;
        }

        /// <summary>下载并更新，带进度回调（Action&lt;int&gt; 每次读取都触发，不合并）</summary>
        public static async Task DoUpdateAsync(string remoteVersion, string expectedMd5, Action<int>? onProgress)
        {
            var currentExePath = Application.ExecutablePath;
            if (string.IsNullOrEmpty(currentExePath) || !File.Exists(currentExePath))
                throw new Exception("无法定位当前程序文件。");

            // 必须带 MD5：没有校验值就无法确认下载内容，宁可放弃更新
            if (!IsValidMd5(expectedMd5))
                throw new Exception("更新包缺少有效的 MD5 校验值，已取消更新。");

            // 从 Release API 获取附件下载地址
            var downloadUrl = await GetLatestReleaseDownloadUrlAsync().ConfigureAwait(false);
            if (string.IsNullOrEmpty(downloadUrl))
                throw new Exception("无法获取更新文件的下载地址。");

            // 临时文件必须与目标 exe 同卷，替换时才能用移动而不是复制
            var exeDir = Path.GetDirectoryName(currentExePath) ?? ".";
            var tempPath = Path.Combine(exeDir,
                $"{Path.GetFileName(currentExePath)}.{Guid.NewGuid():N}.tmp");

            try
            {
                // 下载（所有 await 使用 ConfigureAwait(false)，不回到调用线程）
                using (var response = await _httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                {
                    response.EnsureSuccessStatusCode();
                    var totalBytes = response.Content.Headers.ContentLength ?? -1;
                    using var contentStream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
                    using var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None);

                    if (totalBytes > 0 && onProgress != null)
                    {
                        var buffer = new byte[8192];
                        long readBytes = 0;
                        int bytesRead;
                        while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, bytesRead).ConfigureAwait(false);
                            readBytes += bytesRead;
                            onProgress((int)(readBytes * 100 / totalBytes));
                        }
                    }
                    else
                    {
                        await contentStream.CopyToAsync(fileStream).ConfigureAwait(false);
                        onProgress?.Invoke(100);
                    }
                }

                var downloaded = new FileInfo(tempPath);
                if (!downloaded.Exists || downloaded.Length == 0)
                    throw new Exception("下载的文件为空，已取消更新。");

                // 校验 MD5（强制）
                var downloadedMd5 = await ComputeFileMd5Async(tempPath).ConfigureAwait(false);
                if (!string.Equals(downloadedMd5, expectedMd5, StringComparison.OrdinalIgnoreCase))
                    throw new Exception("下载文件校验失败，文件可能已损坏或与发布版本不一致。");

                // 下载并校验成功，更新版本号
                if (!string.IsNullOrEmpty(remoteVersion))
                {
                    SetCurrentVersion(remoteVersion);
                }

                // 执行更新（批处理替换 + 重启）
                ExecuteUpdate(currentExePath, tempPath);
            }
            catch
            {
                TryDeleteFile(tempPath);
                throw;
            }
        }

        private static bool IsValidMd5(string? md5)
        {
            if (string.IsNullOrEmpty(md5) || md5.Length != 32)
                return false;
            foreach (var ch in md5)
            {
                if (!Uri.IsHexDigit(ch))
                    return false;
            }
            return true;
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch { }
        }

        /// <summary>批处理字面量转义：% 会被 cmd 当作变量展开</summary>
        private static string EscapeBatLiteral(string value) => value.Replace("%", "%%");

        /// <summary>执行更新：生成批处理脚本 → 退出当前程序 → 脚本替换 → 重启</summary>
        private static void ExecuteUpdate(string currentExePath, string newExePath)
        {
            if (!File.Exists(newExePath) || new FileInfo(newExePath).Length == 0)
                throw new Exception("更新包不存在或为空，已取消更新。");

            var backupPath = currentExePath + ".bak";

            // 先备份旧 exe；删除旧文件有重试上限，失败则回滚备份后重启，
            // 避免旧版本被删掉、新文件又没换上导致程序彻底消失
            var batContent = new StringBuilder();
            batContent.AppendLine("@echo off");
            batContent.AppendLine("chcp 65001 > nul");
            batContent.AppendLine("title 正在更新 wxd.files.winforms...");
            batContent.AppendLine("echo 正在更新软件，请稍候...");
            batContent.AppendLine("ping 127.0.0.1 -n 3 > nul");
            batContent.AppendLine($"set \"TARGET={EscapeBatLiteral(currentExePath)}\"");
            batContent.AppendLine($"set \"NEWFILE={EscapeBatLiteral(newExePath)}\"");
            batContent.AppendLine($"set \"BAK={EscapeBatLiteral(backupPath)}\"");
            batContent.AppendLine("copy /y \"%TARGET%\" \"%BAK%\" > nul 2>&1");
            batContent.AppendLine("set RETRY=0");
            batContent.AppendLine(":retry");
            batContent.AppendLine("del /f /q \"%TARGET%\" > nul 2>&1");
            batContent.AppendLine("if not exist \"%TARGET%\" goto replace");
            batContent.AppendLine("set /a RETRY+=1");
            batContent.AppendLine("if %RETRY% GEQ 30 goto rollback");
            batContent.AppendLine("ping 127.0.0.1 -n 2 > nul");
            batContent.AppendLine("goto retry");
            batContent.AppendLine(":replace");
            batContent.AppendLine("move /y \"%NEWFILE%\" \"%TARGET%\" > nul 2>&1");
            batContent.AppendLine("if exist \"%TARGET%\" goto restart");
            batContent.AppendLine(":rollback");
            batContent.AppendLine("if exist \"%BAK%\" copy /y \"%BAK%\" \"%TARGET%\" > nul 2>&1");
            batContent.AppendLine(":restart");
            batContent.AppendLine("if exist \"%TARGET%\" start \"\" \"%TARGET%\"");
            batContent.AppendLine("del /f /q \"%~f0\" > nul 2>&1");

            var batPath = Path.Combine(
                Path.GetTempPath(),
                $"wxd_update_{Guid.NewGuid():N}.bat");
            File.WriteAllText(batPath, batContent.ToString(), Encoding.UTF8);

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = batPath,
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                });
            }
            catch
            {
                TryDeleteFile(batPath);
                throw;
            }

            Environment.Exit(0);
        }

        /// <summary>从 Gitee Release API 获取版本检查结果</summary>
        private static async Task<CheckResult?> FetchCheckResultAsync()
        {
            // 1. 获取最新 Release 信息
            var latestRelease = await GetLatestReleaseAsync().ConfigureAwait(false);
            if (latestRelease == null)
                return null;

            var remoteVersion = latestRelease.Version;
            var remoteMd5 = latestRelease.Md5;

            // 2. 先判断版本号是否相同
            var configVersion = GetCurrentVersion();
            var asmVersion = Assembly.GetExecutingAssembly().GetName()?.Version?.ToString() ?? "0.0.0.0";
            var currentVersion = !string.IsNullOrEmpty(configVersion) ? configVersion : asmVersion;

            bool versionSame = string.Equals(currentVersion, remoteVersion, StringComparison.Ordinal);

            // 3. 获取更新日志
            string releaseNotes;
            if (versionSame)
            {
                // 版本号相同 → 取最近 3 个版本的更新日志
                releaseNotes = await FetchRecentReleaseNotesAsync(3).ConfigureAwait(false);
            }
            else
            {
                // 版本号不同 → 获取当前版本到最新版本之间的累积更新日志
                releaseNotes = await FetchReleaseNotesBetweenAsync(
                    currentVersion, remoteVersion).ConfigureAwait(false);
            }

            // 计算本地 exe 的 MD5
            var currentExePath = Application.ExecutablePath;
            string? currentMd5 = null;
            if (!string.IsNullOrEmpty(currentExePath) && File.Exists(currentExePath))
            {
                currentMd5 = await ComputeFileMd5Async(currentExePath).ConfigureAwait(false);
            }

            // 核心逻辑：版本号优先，MD5 作为辅助
            bool needUpdateFile = false;

            if (versionSame)
            {
                // 版本号相同 → 不需要更新（即使 MD5 不同，可能是编译差异或解析失败）
                needUpdateFile = false;
            }
            else if (!string.IsNullOrEmpty(currentMd5) && !string.IsNullOrEmpty(remoteMd5))
            {
                // 版本号不同 + 两边都有有效 MD5 → 以 MD5 为准
                needUpdateFile = !string.Equals(currentMd5, remoteMd5, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                // 版本号不同 + MD5 不可比较 → 保守标记为需要更新
                needUpdateFile = true;
            }

            return new CheckResult
            {
                RemoteVersion = remoteVersion,
                RemoteMd5 = remoteMd5,
                ReleaseNotes = releaseNotes,
                HasUpdate = needUpdateFile,
                CurrentVersion = currentVersion
            };
        }

        /// <summary>获取最新 Release 的 exe 下载地址</summary>
        private static async Task<string?> GetLatestReleaseDownloadUrlAsync()
        {
            var json = await _httpClient.GetStringAsync(ReleaseLatestApi).ConfigureAwait(false);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (root.TryGetProperty("assets", out var assets) && assets.GetArrayLength() > 0)
            {
                return assets[0].GetProperty("browser_download_url").GetString();
            }
            return null;
        }

        /// <summary>从 Gitee API 获取最新 Release 信息（版本号 + MD5）</summary>
        private static async Task<ReleaseInfo?> GetLatestReleaseAsync()
        {
            try
            {
                var json = await _httpClient.GetStringAsync(ReleaseLatestApi).ConfigureAwait(false);
                return ParseReleaseInfo(json);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>从 JSON 解析 Release 信息</summary>
        private static ReleaseInfo? ParseReleaseInfo(string json)
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            var tagName = root.GetProperty("tag_name").GetString() ?? "";
            var version = tagName.StartsWith("v") ? tagName[1..] : tagName;
            var body = root.GetProperty("body").GetString() ?? "";

            // 从 body 中解析 MD5: **MD5**: `xxx`
            var md5Match = Regex.Match(body, @"\*\*MD5\*\*:\s*`([a-f0-9]{32})`", RegexOptions.IgnoreCase);
            var md5 = md5Match.Success ? md5Match.Groups[1].Value : "";

            string? downloadUrl = null;
            if (root.TryGetProperty("assets", out var assets) && assets.GetArrayLength() > 0)
            {
                downloadUrl = assets[0].GetProperty("browser_download_url").GetString();
            }

            return new ReleaseInfo
            {
                Version = version,
                Md5 = md5,
                DownloadUrl = downloadUrl,
                Body = body
            };
        }

        /// <summary>从 Release body 中提取纯更新日志（去掉 MD5 等元信息行）</summary>
        private static string ExtractReleaseNotes(string body)
        {
            if (string.IsNullOrEmpty(body))
                return "";

            // 按 "---" 分割，取第一部分（更新日志）作为有效内容
            var lines = body.Split('\n');
            var sb = new StringBuilder();
            bool foundSeparator = false;

            foreach (var line in lines)
            {
                var trimmed = line.TrimEnd('\r');

                // 遇到 --- 分隔线就停止
                if (trimmed.Trim() == "---")
                {
                    foundSeparator = true;
                    continue;
                }
                if (foundSeparator)
                    continue;

                // 跳过版本标题行
                if (trimmed.Trim().StartsWith("## version", StringComparison.OrdinalIgnoreCase))
                    continue;

                sb.AppendLine(trimmed);
            }

            return sb.ToString().Trim();
        }

        /// <summary>获取两个版本之间的累积更新日志</summary>
        private static async Task<string> FetchReleaseNotesBetweenAsync(
            string currentVersion, string targetVersion)
        {
            try
            {
                var json = await _httpClient.GetStringAsync(ReleasesListApi).ConfigureAwait(false);
                using var doc = JsonDocument.Parse(json);
                var releases = doc.RootElement;

                // 解析版本号（去掉 'v' 前缀）
                static bool TryParseVersion(string? v, out Version? parsed)
                {
                    parsed = null;
                    if (string.IsNullOrEmpty(v)) return false;
                    var clean = v.StartsWith("v") ? v[1..] : v;
                    return Version.TryParse(clean, out parsed);
                }

                Version? currentVer = null;
                Version? targetVer = null;
                TryParseVersion(currentVersion, out currentVer);
                TryParseVersion(targetVersion, out targetVer);

                // 先收集符合条件的 Release 到列表
                var matchedReleases = new List<(Version Ver, string TagName, string Notes)>();

                foreach (var release in releases.EnumerateArray())
                {
                    var tagName = release.GetProperty("tag_name").GetString() ?? "";
                    var verStr = tagName.StartsWith("v") ? tagName[1..] : tagName;
                    var body = release.GetProperty("body").GetString() ?? "";

                    if (!TryParseVersion(verStr, out var relVer) || relVer == null)
                        continue;

                    // 筛选：只取 currentVersion < relVer <= targetVersion 之间的版本
                    if (targetVer != null && relVer > targetVer)
                        continue; // 超出目标版本

                    if (currentVer != null && relVer <= currentVer)
                        continue; // 不高于当前版本

                    var notes = ExtractReleaseNotes(body);
                    if (!string.IsNullOrEmpty(notes))
                    {
                        matchedReleases.Add((relVer, tagName, notes));
                    }
                }

                // 按版本号降序排列（最新在上面）
                matchedReleases.Sort((a, b) => b.Ver.CompareTo(a.Ver));

                var sb = new StringBuilder();
                foreach (var (_, tagName, notes) in matchedReleases)
                {
                    sb.AppendLine($"━━━ {tagName} ━━━");
                    sb.AppendLine(notes);
                    sb.AppendLine();
                }

                return sb.ToString().Trim();
            }
            catch
            {
                // 降级：只取最新 Release 的更新日志
                try
                {
                    var latest = await GetLatestReleaseAsync().ConfigureAwait(false);
                    if (latest != null)
                        return ExtractReleaseNotes(latest.Body);
                }
                catch { }
                return "";
            }
        }

        /// <summary>获取最近 N 个版本的更新记录（格式化纯文本），供外部调用</summary>
        public static async Task<string> GetRecentReleaseNotesAsync(int count)
        {
            return await FetchRecentReleaseNotesAsync(count).ConfigureAwait(false);
        }

        /// <summary>获取最近 N 个版本的更新日志</summary>
        private static async Task<string> FetchRecentReleaseNotesAsync(int count)
        {
            try
            {
                var json = await _httpClient.GetStringAsync(ReleasesListApi).ConfigureAwait(false);
                using var doc = JsonDocument.Parse(json);
                var releases = doc.RootElement;

                static bool TryParseVersion(string? v, out Version? parsed)
                {
                    parsed = null;
                    if (string.IsNullOrEmpty(v)) return false;
                    var clean = v.StartsWith("v") ? v[1..] : v;
                    return Version.TryParse(clean, out parsed);
                }

                var matchedReleases = new List<(Version Ver, string TagName, string Notes)>();

                foreach (var release in releases.EnumerateArray())
                {
                    var tagName = release.GetProperty("tag_name").GetString() ?? "";
                    var verStr = tagName.StartsWith("v") ? tagName[1..] : tagName;
                    var body = release.GetProperty("body").GetString() ?? "";

                    if (!TryParseVersion(verStr, out var relVer) || relVer == null)
                        continue;

                    var notes = ExtractReleaseNotes(body);
                    if (!string.IsNullOrEmpty(notes))
                    {
                        matchedReleases.Add((relVer, tagName, notes));
                    }
                }

                matchedReleases.Sort((a, b) => b.Ver.CompareTo(a.Ver));
                var topN = matchedReleases.Take(count).ToList();

                var sb = new StringBuilder();
                foreach (var (_, tagName, notes) in topN)
                {
                    sb.AppendLine($"━━━ {tagName} ━━━");
                    sb.AppendLine(notes);
                    sb.AppendLine();
                }

                return sb.ToString().Trim();
            }
            catch
            {
                return "";
            }
        }

        /// <summary>计算文件的 MD5 值</summary>
        public static async Task<string> ComputeFileMd5Async(string filePath)
        {
            using var md5 = MD5.Create();
            using var stream = File.OpenRead(filePath);
            var hash = await md5.ComputeHashAsync(stream).ConfigureAwait(false);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }

        #region Config 读写

        /// <summary>获取上次检查时间（公开版本，供 UI 显示）</summary>
        public static DateTime GetLastCheckTimePublic()
        {
            return GetLastCheckTime();
        }

        /// <summary>获取上次检查时间</summary>
        private static DateTime GetLastCheckTime()
        {
            var val = CsvSettings.GetStringValue(KeyLastCheckTime, "");
            if (string.IsNullOrEmpty(val)) return DateTime.MinValue;
            if (DateTime.TryParse(val, out var dt)) return dt;
            return DateTime.MinValue;
        }

        /// <summary>设置上次检查时间</summary>
        public static void SetLastCheckTime(DateTime time)
        {
            CsvSettings.SetStringValue(KeyLastCheckTime, time.ToString("yyyy-MM-dd HH:mm:ss"));
        }

        /// <summary>获取跳过版本的列表</summary>
        public static List<string> GetSkippedVersions()
        {
            var json = CsvSettings.GetStringValue(KeySkippedVersions, "[]");
            try
            {
                return JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
            }
            catch
            {
                return new List<string>();
            }
        }

        /// <summary>添加一个跳过的版本</summary>
        public static void AddSkippedVersion(string version)
        {
            var list = GetSkippedVersions();
            if (!list.Contains(version))
            {
                list.Add(version);
                var json = JsonSerializer.Serialize(list);
                CsvSettings.SetStringValue(KeySkippedVersions, json);
            }
        }

        /// <summary>获取 config 中记录的当前版本号</summary>
        public static string GetCurrentVersion()
        {
            return CsvSettings.GetStringValue(KeyCurrentVersion, "");
        }

        /// <summary>设置 config 中的当前版本号（更新成功后调用）</summary>
        public static void SetCurrentVersion(string version)
        {
            CsvSettings.SetStringValue(KeyCurrentVersion, version);
        }

        #endregion
    }
}
