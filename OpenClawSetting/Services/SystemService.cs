using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenClawSetting.Models;
using System.Text.Json;
using Microsoft.Win32; // 用于操作注册表刷新环境变量using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenClawSetting.Models;
using System.Text.Json;
using Microsoft.Win32;
using NLog; // 引入 NLog

namespace OpenClawSetting.Services
{
    public class SystemService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private static readonly string UserProfilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        public static readonly string OpenClawDir = Path.Combine(UserProfilePath, ".openclaw");
        public static readonly string AgentDir = Path.Combine(OpenClawDir, "agents", "main", "agent");
        public static readonly string ConfigFilePath = Path.Combine(AgentDir, "auth-profiles.json");     
        
        // === 新增：定义目标 Node.js 版本 ===
        public const string TargetNodeVersion = "v24.14.0"; // 定义目标版本号

        // Node.js LTS 下载地址
        //public const string NodeJsDownloadUrl = "https://nodejs.org/dist/v20.14.0/node-v20.14.0-x64.msi";
        //public const string NodeJsDownloadUrl = "https://nodejs.org/dist/v20.14.0/node-v20.14.0-x64.msi";
        // 推荐保持 v22 LTS (更稳定，且满足 openclaw 要求)
        //public const string NodeJsDownloadUrl = "https://nodejs.org/dist/v22.14.0/node-v22.14.0-x64.msi";
        // 更新为 Node.js v24.14.0 (满足 OpenClaw 最新版需求)
        //public const string NodeJsDownloadUrl = "https://nodejs.org/dist/v24.14.0/node-v24.14.0-x64.msi";
        public static readonly string NodeJsDownloadUrl = $"https://nodejs.org/dist/{TargetNodeVersion}/node-{TargetNodeVersion.Substring(1)}-x64.msi";

        // Git for Windows 下载地址 (使用官方 CDN，推荐使用最新稳定版链接，或者固定版本)
        public const string GitDownloadUrl = "https://github.com/git-for-windows/git/releases/download/v2.44.0.windows.1/Git-2.44.0-64-bit.exe";

        /// <summary>
        /// 执行 CMD 命令并实时返回输出
        /// </summary>
        public async Task ExecuteCommandAsync(string command, Action<string> onOutput, CancellationToken cancellationToken)
        {
            using (Process process = new Process())
            {
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.Arguments = $"/c {command}";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.StandardOutputEncoding = Encoding.UTF8;
                process.StartInfo.StandardErrorEncoding = Encoding.UTF8;

                process.OutputDataReceived += (sender, e) => { if (e.Data != null) onOutput($"[INFO] {e.Data}"); };
                process.ErrorDataReceived += (sender, e) => { if (e.Data != null) onOutput($"[ERROR] {e.Data}"); };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                await process.WaitForExitAsync(cancellationToken);
            }
        }

        #region 下载与安装方法
        public async Task DownloadFileAsync(string url, string savePath, IProgress<(long received, long total)>? progress, CancellationToken cancellationToken)
        {
            using (HttpClient client = new HttpClient())
            {
                // ... (之前的下载代码保持不变)
                using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken))
                {
                    response.EnsureSuccessStatusCode();
                    var totalBytes = response.Content.Headers.ContentLength ?? -1L;
                    using (var stream = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(savePath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        var buffer = new byte[8192];
                        int bytesRead;
                        long totalRead = 0;
                        while ((bytesRead = await stream.ReadAsync(buffer, cancellationToken)) > 0)
                        {
                            await fileStream.WriteAsync(buffer.AsMemory(0, bytesRead), cancellationToken);
                            totalRead += bytesRead;
                            progress?.Report((totalRead, totalBytes));
                        }
                    }
                }
            }
        }

        //public async Task InstallNodeJsAsync(string msiPath, Action<string> onOutput, CancellationToken cancellationToken)
        //{
        //    string args = $"/i \"{msiPath}\" /qn /norestart";
        //    Logger.Info("正在静默安装 Node.js...");

        //    // 这里我们不需要获取详细输出，因为 MSI 静默安装没什么输出
        //    int exitCode = await RunProcessAsync("msiexec.exe", args, onOutput, null, cancellationToken);

        //    if (exitCode != 0) throw new Exception($"Node.js 安装失败，退出代码: {exitCode}");

        //    RefreshEnvironmentVariables();
        //    Logger.Info("Node.js 安装完成，环境变量已刷新。");
        //}
        public async Task InstallNodeJsAsync(string msiPath, CancellationToken cancellationToken)
        {
            // === 关键修复：安装前强制清理旧残留 ===
            Logger.Info("正在清理旧版本残留文件...");
            string roamingPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string npmFolder = Path.Combine(roamingPath, "npm");
            string npmCacheFolder = Path.Combine(roamingPath, "npm-cache");

            try
            {
                if (Directory.Exists(npmFolder))
                {
                    Logger.Info($"删除旧目录: {npmFolder}");
                    Directory.Delete(npmFolder, true);
                }
                if (Directory.Exists(npmCacheFolder))
                {
                    Logger.Info($"删除旧缓存: {npmCacheFolder}");
                    Directory.Delete(npmCacheFolder, true);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"清理残留失败(可忽略，可能文件占用): {ex.Message}");
            }
            // ======================================

            string args = $"/i \"{msiPath}\" /qn /norestart";
            Logger.Info("正在静默安装 Node.js v24.14.0 LTS...");

            int exitCode = await RunProcessAsync("msiexec.exe", args, null, null, cancellationToken);

            if (exitCode != 0) throw new Exception($"Node.js 安装失败，退出代码: {exitCode}");

            RefreshEnvironmentVariables();
            Logger.Info("Node.js 安装完成，环境变量已刷新。");
        }

        /// <summary>
        /// 新增：静默安装 Git
        /// </summary>
        public async Task InstallGitAsync(string exePath, CancellationToken cancellationToken)
        {
            Logger.Info("正在静默安装 Git...");

            // Git 安装参数：
            // /VERYSILENT - 静默安装，不显示界面
            // /NORESTART  - 安装完不重启电脑
            // /COMPONENTS="assoc,assoc_sh" - 推荐组件（可选）
            string args = "/VERYSILENT /NORESTART";

            int exitCode = await RunProcessAsync(exePath, args, null, null, cancellationToken);

            if (exitCode != 0) throw new Exception($"Git 安装失败，退出代码: {exitCode}");

            RefreshEnvironmentVariables();
            Logger.Info("Git 安装完成，环境变量已刷新。");
        }

        /// <summary>
        /// 安装 OpenClaw (集成 GitHub 加速与 SSH 修复)
        /// </summary>
        public async Task InstallOpenClawAsync(CancellationToken cancellationToken)
        {
            if (!IsAdministrator()) throw new UnauthorizedAccessException("安装 OpenClaw 需要管理员权限。");

            Logger.Info("准备安装 OpenClaw...");

            // =========================================================
            // 核心修复：配置 GitHub 镜像加速 (解决 443 连接超时问题)
            // =========================================================
            Logger.Info("正在配置 GitHub 访问加速...");

            // 1. 开启镜像加速
            // 原理：当 git 访问 https://github.com/xxx 时，自动替换为 https://mirror.ghproxy.com/https://github.com/xxx
            // mirror.ghproxy.com 是一个常用的公益加速站点
            string mirrorUrl = "https://mirror.ghproxy.com/https://github.com";

            await RunProcessAsync("cmd.exe", $"/c git config --global url.\"{mirrorUrl}\".insteadOf https://github.com", null, null, cancellationToken);

            // 针对有些包可能使用 git:// 协议，也一并转为 https 并加速
            await RunProcessAsync("cmd.exe", $"/c git config --global url.\"{mirrorUrl}\".insteadOf git://github.com", null, null, cancellationToken);

            Logger.Info("加速配置完成，将使用国内镜像下载依赖。");


            // =========================================================
            // 基础修复：SSH 协议转 HTTPS (解决 Permission denied 问题)
            // =========================================================
            Logger.Info("正在配置协议兼容性...");
            // 将 ssh://git@github.com 强转为 https 协议 (也会走上面的加速镜像)
            await RunProcessAsync("cmd.exe", "/c git config --global url.\"https://github.com/\".insteadOf git@github.com:", null, null, cancellationToken);
            await RunProcessAsync("cmd.exe", "/c git config --global url.\"https://github.com/\".insteadOf ssh://git@github.com/", null, null, cancellationToken);


            // =========================================================
            // 安装流程
            // =========================================================
            Logger.Info("正在配置 npm 镜像源...");
            await RunProcessAsync("cmd.exe", "/c npm config set registry https://registry.npmmirror.com", null, null, cancellationToken);

            Logger.Info("正在执行 npm install -g openclaw@latest ...");

            int exitCode = await RunProcessAsync("cmd.exe", "/c npm install -g openclaw@latest",
                output => Logger.Info(output),
                error => Logger.Error(error),
                cancellationToken);

            if (exitCode != 0)
            {
                // 如果加速镜像也失败了，提示用户检查网络或尝试手动下载
                throw new Exception($"OpenClaw 安装失败 (Exit Code: {exitCode})。\n建议：请检查基础网络是否正常，或尝试重新运行安装。");
            }

            RefreshEnvironmentVariables();
            Logger.Info("OpenClaw 安装成功！");
        }

        #endregion

        #region 辅助方法
        /// <summary>
        /// 核心进程执行方法，返回 ExitCode
        /// </summary>
        private async Task<int> RunProcessAsync(string fileName, string arguments, Action<string>? onOutput, Action<string>? onError, CancellationToken cancellationToken)
        {
            using (Process process = new Process())
            {
                process.StartInfo.FileName = fileName;
                process.StartInfo.Arguments = arguments;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.StandardOutputEncoding = Encoding.UTF8;
                process.StartInfo.StandardErrorEncoding = Encoding.UTF8;

                // 实时输出事件
                process.OutputDataReceived += (sender, e) => { if (!string.IsNullOrEmpty(e.Data)) onOutput?.Invoke(e.Data); };
                process.ErrorDataReceived += (sender, e) => { if (!string.IsNullOrEmpty(e.Data)) onError?.Invoke(e.Data); };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                await process.WaitForExitAsync(cancellationToken);
                return process.ExitCode;
            }
        }

        public string CheckCommandExists(string command)
        {
            // 关键修复：检查前先刷新环境变量，防止“刚装好就提示未安装”
            RefreshEnvironmentVariables();

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", $"/c {command}")
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };
                using (Process proc = Process.Start(psi))
                {
                    string output = proc.StandardOutput.ReadToEnd();
                    proc.WaitForExit();
                    return proc.ExitCode == 0 ? output.Trim() : "未安装";
                }
            }
            catch
            {
                return "检测失败";
            }
        }

        public void RefreshEnvironmentVariables()
        {
            try
            {
                var systemPath = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Session Manager\Environment")?.GetValue("Path")?.ToString();
                var userPath = Registry.CurrentUser.OpenSubKey(@"Environment")?.GetValue("Path")?.ToString();

                string newPath = "";
                if (!string.IsNullOrEmpty(systemPath)) newPath += systemPath;
                if (!string.IsNullOrEmpty(userPath)) newPath += ";" + userPath;

                Environment.SetEnvironmentVariable("Path", newPath, EnvironmentVariableTarget.Process);
            }
            catch (Exception ex)
            {
                Logger.Error($"刷新环境变量失败: {ex.Message}");
            }
        }
        public bool SaveConfigFile(string kimiKey, string? feishuAppId, string? feishuAppSecret)
        {
            try
            {
                if (!Directory.Exists(AgentDir))
                {
                    Directory.CreateDirectory(AgentDir);
                }

                long timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                string profileId = "kimi-coding:default";

                // 1. 生成 auth-profiles.json
                var authConfig = new AuthProfileConfig();
                authConfig.Profiles.Add(profileId, new AuthProfileEntry
                {
                    Key = kimiKey
                });
                authConfig.UsageStats.Add(profileId, new UsageStatsEntry
                {
                    LastUsed = timestamp
                });

                string authJson = JsonSerializer.Serialize(authConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigFilePath, authJson); // ConfigFilePath 指向 auth-profiles.json
                Logger.Info($"已生成: {ConfigFilePath}");

                // 2. 生成 models.json (同级目录)
                string modelsPath = Path.Combine(AgentDir, "models.json");

                var modelsConfig = new ModelsConfig();
                modelsConfig.Providers.Add("kimi-coding", new ProviderConfig
                {
                    ApiKey = kimiKey, // 这里也需要 Key
                    Models = new List<ModelEntry>
            {
                new ModelEntry() // 使用默认值即可，已在类中定义好
            }
                });

                string modelsJson = JsonSerializer.Serialize(modelsConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(modelsPath, modelsJson);
                Logger.Info($"已生成: {modelsPath}");

                // 3. 如果需要飞书配置，请告诉我飞书配置应该放在哪个 JSON 文件里
                // 目前根据你提供的文件，飞书配置尚未包含在内。
                if (!string.IsNullOrEmpty(feishuAppId))
                {
                    Logger.Warn("提示：飞书配置尚未写入配置文件，请确认是否需要在 models.json 中添加字段。");
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"配置文件生成失败: {ex.Message}");
                return false;
            }
        }

        public bool IsAdministrator()
        {
            using (var identity = System.Security.Principal.WindowsIdentity.GetCurrent())
            {
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
        }

        #endregion

        /// <summary>
        /// 获取 npm 包的最新版本号
        /// </summary>
        public async Task<string> GetLatestNpmVersion(string packageName)
        {
            try
            {
                // 使用 npm view 命令获取最新版本
                // 这个命令会请求 npm 仓库，可能需要几秒钟
                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", $"/c npm view {packageName} version")
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };

                using (Process proc = Process.Start(psi))
                {
                    string output = await proc.StandardOutput.ReadToEndAsync();
                    proc.WaitForExit();
                    // 输出通常是纯版本号，如 "1.0.5\n"
                    return output.Trim();
                }
            }
            catch
            {
                return "获取失败";
            }
        }
    }
}
