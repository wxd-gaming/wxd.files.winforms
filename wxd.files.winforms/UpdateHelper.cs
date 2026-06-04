using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
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
    }

    /// <summary>更新核心逻辑</summary>
    public static class UpdateHelper
    {
        // Gitee raw 链接（固定地址）
        private const string Md5CheckUrl =
            "https://gitee.com/wxd-gaming/wxd.files.winforms/raw/master/md5.txt";
        private const string DownloadUrl =
            "https://gitee.com/wxd-gaming/wxd.files.winforms/raw/master/wxd.files.winforms.exe";

        // Config 中的键名
        private const string KeyLastCheckTime = "UpdateLastCheckTime";
        private const string KeySkippedVersions = "UpdateSkippedVersions";
        private const string KeyCurrentVersion = "UpdateCurrentVersion";

        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(8)
        };

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

                var result = await FetchCheckResultAsync();
                if (result == null)
                    return;

                // MD5 相同但版本号不同 → 静默更新版本号，不弹窗
                if (!result.HasUpdate && !string.Equals(result.CurrentVersion, result.RemoteVersion, StringComparison.Ordinal))
                {
                    SetCurrentVersion(result.RemoteVersion);
                    return;
                }

                if (!result.HasUpdate)
                    return;

                // 检查是否在跳过列表里
                var skipped = GetSkippedVersions();
                if (skipped.Contains(result.RemoteVersion))
                    return;

                // 有更新 → 使用自定义弹窗（三个按钮：立即更新 / 跳过版本 / 取消）
                // 注意：点击"立即更新"后，UpdateNotifyForm 内部会切换为下载进度界面并执行下载，
                // 下载完成后自动执行 Environment.Exit(0)，所以此处不需要再调用 DoUpdateAsync
                var isDark = CsvSettings.ThemeMode == 0;
                var notifyForm = new UpdateNotifyForm(result, isDark);
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
            var result = await FetchCheckResultAsync();
            if (result == null)
                throw new Exception("无法获取远程版本信息，请检查网络连接。");

            // MD5 相同但版本号不同 → 静默更新版本号，不提示更新
            if (!result.HasUpdate && !string.Equals(result.CurrentVersion, result.RemoteVersion, StringComparison.Ordinal))
            {
                SetCurrentVersion(result.RemoteVersion);
                result.CurrentVersion = result.RemoteVersion;
            }

            return result;
        }

        /// <summary>下载并更新，带进度回调</summary>
        public static async Task DoUpdateAsync(string remoteVersion, string expectedMd5, IProgress<int>? progress)
        {
            var currentExePath = Application.ExecutablePath;
            if (string.IsNullOrEmpty(currentExePath) || !File.Exists(currentExePath))
                throw new Exception("无法定位当前程序文件。");

            var tempPath = currentExePath + ".tmp";

            // 下载
            using (var response = await _httpClient.GetAsync(DownloadUrl, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();
                var totalBytes = response.Content.Headers.ContentLength ?? -1;
                using var contentStream = await response.Content.ReadAsStreamAsync();
                using var fileStream = File.Create(tempPath);

                if (totalBytes > 0 && progress != null)
                {
                    var buffer = new byte[8192];
                    long readBytes = 0;
                    int bytesRead;
                    while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                    {
                        await fileStream.WriteAsync(buffer, 0, bytesRead);
                        readBytes += bytesRead;
                        progress.Report((int)(readBytes * 100 / totalBytes));
                    }
                }
                else
                {
                    await contentStream.CopyToAsync(fileStream);
                    progress?.Report(100);
                }
            }

            // 校验 MD5
            if (!string.IsNullOrEmpty(expectedMd5))
            {
                var downloadedMd5 = await ComputeFileMd5Async(tempPath);
                if (!string.Equals(downloadedMd5, expectedMd5, StringComparison.OrdinalIgnoreCase))
                {
                    File.Delete(tempPath);
                    throw new Exception("下载文件校验失败，文件可能已损坏。");
                }
            }

            // 下载并校验成功，更新版本号
            if (!string.IsNullOrEmpty(remoteVersion))
            {
                SetCurrentVersion(remoteVersion);
            }

            // 执行更新（批处理替换 + 重启）
            ExecuteUpdate(currentExePath, tempPath);
        }

        /// <summary>执行更新：生成批处理脚本 → 退出当前程序 → 脚本替换 → 重启</summary>
        private static void ExecuteUpdate(string currentExePath, string newExePath)
        {
            var batContent = new StringBuilder();
            batContent.AppendLine("@echo off");
            batContent.AppendLine("chcp 65001 > nul");
            batContent.AppendLine("title 正在更新 wxd.files.winforms...");
            batContent.AppendLine("echo 正在更新软件，请稍候...");
            batContent.AppendLine($"ping 127.0.0.1 -n 3 > nul");
            batContent.AppendLine("");
            batContent.AppendLine(":retry");
            batContent.AppendLine($"del /f /q \"{currentExePath}\"");
            batContent.AppendLine($"if exist \"{currentExePath}\" goto retry");
            batContent.AppendLine($"move /y \"{newExePath}\" \"{currentExePath}\"");
            batContent.AppendLine($"start \"\" \"{currentExePath}\"");
            batContent.AppendLine($"del /f /q \"%~f0\"");

            var batPath = Path.Combine(
                Path.GetTempPath(),
                $"wxd_update_{Guid.NewGuid():N}.bat");
            File.WriteAllText(batPath, batContent.ToString(), Encoding.UTF8);

            Process.Start(new ProcessStartInfo
            {
                FileName = batPath,
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });

            Environment.Exit(0);
        }

        /// <summary>从远程获取版本检查结果</summary>
        private static async Task<CheckResult?> FetchCheckResultAsync()
        {
            var response = await _httpClient.GetStringAsync(Md5CheckUrl);
            var lines = response.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (lines.Length < 2) return null;

            var remoteVersion = lines[0].Trim();
            var remoteMd5 = lines[1].Trim();

            // 当前程序集版本
            var currentVersion = Assembly.GetExecutingAssembly().GetName()?.Version?.ToString() ?? "0.0.0.0";

            // 优先使用 config 中记录的版本（防止用户手动覆盖 exe 后版本号不对应）
            var configVersion = GetCurrentVersion();
            if (!string.IsNullOrEmpty(configVersion))
                currentVersion = configVersion;

            // 计算本地 exe 的 MD5
            var currentExePath = Application.ExecutablePath;
            string? currentMd5 = null;
            if (!string.IsNullOrEmpty(currentExePath) && File.Exists(currentExePath))
            {
                currentMd5 = await ComputeFileMd5Async(currentExePath);
            }

            // 核心逻辑：以 MD5 为准判断是否需要更新文件
            bool needUpdateFile = false;
            if (!string.IsNullOrEmpty(currentMd5))
            {
                if (!string.Equals(currentMd5, remoteMd5, StringComparison.OrdinalIgnoreCase))
                {
                    // MD5 不同 → 需要更新文件（无论版本号是否相同）
                    needUpdateFile = true;
                }
            }
            else
            {
                // 无法计算本地 MD5，保守起见标记为需要更新
                needUpdateFile = true;
            }

            return new CheckResult
            {
                RemoteVersion = remoteVersion,
                RemoteMd5 = remoteMd5,
                HasUpdate = needUpdateFile,
                CurrentVersion = currentVersion
            };
        }

        /// <summary>计算文件的 MD5 值</summary>
        public static async Task<string> ComputeFileMd5Async(string filePath)
        {
            using var md5 = MD5.Create();
            using var stream = File.OpenRead(filePath);
            var hash = await md5.ComputeHashAsync(stream);
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
