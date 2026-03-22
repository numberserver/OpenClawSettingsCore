using NLog; // 引入 NLog
using NLog.Windows.Forms;
using OpenClawSetting.Models;
using OpenClawSetting.Services;
using System;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing;
using System.IO;
using System.IO;
using System.Threading;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenClawSetting
{
    public partial class MainForm : Form
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly SystemService _systemService = new SystemService();
        private CancellationTokenSource? _runningCts;

        // 控件定义
        private TabControl tabControl;
        private TabPage tabEnv, tabConfig, tabRun;
        private RichTextBox logBox;
        private ProgressBar progressBar;

        // 环境页控件
        private Label lblGitStatus, lblNodeStatus, lblClawStatus;
        private Button btnInstallGit, btnInstallNode, btnInstallClaw, btnRefresh;

        // === 新增：将启动和停止按钮定义为类变量 ===
        private Button btnStart, btnStop;

        // 配置页控件
        private TextBox txtKimiKey, txtAppId, txtAppSecret;

        //// 定义全局 Logger
        //private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        //private readonly SystemService _systemService = new SystemService();
        //private CancellationTokenSource? _runningCts;

        //// 控件定义
        //private TabControl tabControl;
        //private TabPage tabEnv, tabConfig, tabRun;
        //private RichTextBox logBox;
        //private ProgressBar progressBar; // 新增进度条

        //// 环境页控件
        //private Label lblGitStatus, lblNodeStatus, lblClawStatus; // 新增 lblGitStatus
        //private Button btnInstallGit, btnInstallNode, btnInstallClaw, btnRefresh;

        //// 配置页控件
        //private TextBox txtKimiKey, txtAppId, txtAppSecret;

        public MainForm()
        {
            InitializeComponent();
            InitializeCustomLayout();
            Text = "OpenClaw 自动化部署工具 v2.1 (增强日志版)";
            Size = new Size(900, 700); // 高度稍微增加一点
            StartPosition = FormStartPosition.CenterScreen;

            // === 关键：初始化 NLog 绑定到 RichTextBox ===
            // 这里的 logBox 是你定义的 RichTextBox 控件
            // 必须在 InitializeComponent 或 InitializeCustomLayout 执行后调用
            RichTextBoxTarget.ReInitializeAllTextboxes(this);

            LoadEnvironmentStatus();
        }

        private void InitializeCustomLayout()
        {
            this.tabControl = new TabControl { Dock = DockStyle.Fill };
            this.Controls.Add(tabControl);

            // === Tab 1: 环境tab搭建 (修改部分) ===
            tabEnv = new TabPage("1. 环境搭建");
            var pnlEnv = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), AutoScroll = true };

            // 1. Git 状态
            lblGitStatus = new Label { Text = "Git 状态: 检测中...", AutoSize = true, Font = new Font("微软雅黑", 11) };
            btnInstallGit = new Button { Text = "0. 安装 Git (必需)", Width = 280, Height = 40, Margin = new Padding(0, 5, 0, 0), BackColor = Color.Lavender };
            btnInstallGit.Click += BtnInstallGit_Click;

            // 2. Node.js 状态
            lblNodeStatus = new Label { Text = "Node.js 状态: 检测中...", AutoSize = true, Font = new Font("微软雅黑", 11), Margin = new Padding(0, 20, 0, 0) };
            btnInstallNode = new Button { Text = "1. 安装 Node.js", Width = 280, Height = 40, Margin = new Padding(0, 5, 0, 0), BackColor = Color.LightBlue };
            btnInstallNode.Click += BtnInstallNode_Click;

            // 3. OpenClaw 状态
            lblClawStatus = new Label { Text = "OpenClaw 状态: 检测中...", AutoSize = true, Font = new Font("微软雅黑", 11), Margin = new Padding(0, 20, 0, 0) };
            btnInstallClaw = new Button { Text = "2. 安装 OpenClaw (需管理员)", Width = 280, Height = 40, Margin = new Padding(0, 5, 0, 0), BackColor = Color.LightGoldenrodYellow };
            btnInstallClaw.Click += BtnInstallClaw_Click;

            // 其他
            btnRefresh = new Button { Text = "🔄 刷新状态", AutoSize = true, Margin = new Padding(0, 30, 0, 0) };
            btnRefresh.Click += (s, e) => LoadEnvironmentStatus();

            progressBar = new ProgressBar { Width = 280, Height = 23, Margin = new Padding(0, 10, 0, 0) };

            pnlEnv.Controls.AddRange(new Control[] {
                lblGitStatus, btnInstallGit,
                lblNodeStatus, btnInstallNode,
                lblClawStatus, btnInstallClaw,
                btnRefresh, progressBar
            });

            tabEnv.Controls.Add(pnlEnv);
            tabControl.Controls.Add(tabEnv);

            BuildConfigTab();
            BuildRunTab();
        }

        // 抽离 Tab2 构建代码以便阅读
        private void BuildConfigTab()
        {
            tabConfig = new TabPage("2. 参数配置");
            var pnlConfig = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), AutoScroll = true };
            void AddInputRow(string label, string placeholder, out TextBox txtBox)
            {
                var lbl = new Label { Text = label, AutoSize = true, Font = new Font("微软雅黑", 10), Margin = new Padding(0, 10, 0, 0) };
                var localBox = new TextBox { Width = 400, Margin = new Padding(0, 5, 0, 20), Font = new Font("微软雅黑", 10), Text = placeholder, ForeColor = Color.Gray };
                localBox.GotFocus += (s, e) => { if (localBox.Text == placeholder) { localBox.Text = ""; localBox.ForeColor = Color.Black; } };
                pnlConfig.Controls.Add(lbl); pnlConfig.Controls.Add(localBox);
                txtBox = localBox;
            }
            AddInputRow("Kimi API Key:", "sk-xxxxxxxxxxxxxxxx", out txtKimiKey);
            AddInputRow("飞书 App ID:", "cli_xxxxxxxxxx", out txtAppId);
            AddInputRow("飞书 App Secret:", "输入密钥...", out txtAppSecret);
            var btnSave = new Button { Text = "💾 保存配置文件", Width = 200, Height = 45, Margin = new Padding(0, 30, 0, 0), BackColor = Color.LightGreen };
            btnSave.Click += BtnSaveConfig_Click;
            pnlConfig.Controls.Add(btnSave);
            tabConfig.Controls.Add(pnlConfig);
            tabControl.Controls.Add(tabConfig);

            //// 在 BuildConfigTab 方法中添加
            //var btnFeishuWizard = new Button
            //{
            //    Text = "🚀 自动创建/配置飞书应用",
            //    Width = 200,
            //    Height = 40,
            //    Margin = new Padding(0, 50, 0, 0),
            //    BackColor = Color.LightSkyBlue
            //};
            //btnFeishuWizard.Click += BtnFeishuWizard_Click;
            //pnlConfig.Controls.Add(btnFeishuWizard);
            // 在 BuildConfigTab 方法中添加
            var btnAutoCreate = new Button
            {
                Text = "🚀 一键创建飞书应用",
                Width = 150,
                Height = 30,
                Margin = new Padding(0, 10, 0, 0),
                BackColor = Color.LightSkyBlue
            };
            btnAutoCreate.Click += (s, e) =>
            {
                var wizard = new FeishuWizardFormV1();
                wizard.Show();
                //if (wizard.ShowDialog() == DialogResult.OK)
                //{
                //    // 自动回填
                //    txtAppId.Text = wizard.ResultAppId;
                //    txtAppSecret.Text = wizard.ResultAppSecret;
                //}
            };
            pnlConfig.Controls.Add(btnAutoCreate);
        }

        private void BtnFeishuWizard_Click(object? sender, EventArgs e)
        {
            Logger.Info("启动飞书配置向导...");
            var wizard = new FeishuWizardForm();

            if (wizard.ShowDialog() == DialogResult.OK)
            {
                // 用户在向导中配置成功，自动回填数据
                txtAppId.Text = wizard.ResultAppId;
                txtAppSecret.Text = wizard.ResultAppSecret;
                Logger.Info($"已自动回填 AppID: {wizard.ResultAppId}");

                // 提示用户下一步
                MessageBox.Show("凭证已获取！请在飞书后台完成版本发布，然后在此点击“保存配置文件”。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 抽离 Tab3 构建代码
        private void BuildRunTab()
        {
            tabRun = new TabPage("3. 运行与配对");
            var pnlRun = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            var lblTip = new Label { Text = "提示：点击“启动服务”后，请前往飞书私聊机器人获取配对码。", AutoSize = true, Font = new Font("微软雅黑", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 122, 204) };
            lblTip.Location = new Point(10, 10);

            // 启动按钮
            btnStart = new Button { Text = "▶ 启动服务", Location = new Point(10, 50), Width = 120, Height = 40 };
            btnStart.Click += BtnStart_Click;

            // 停止按钮 (修改了文字和颜色，警示这是强制停止)
            btnStop = new Button { Text = "⏹ 强制停止服务", Location = new Point(140, 50), Width = 120, Height = 40, Enabled = false, BackColor = Color.OrangeRed, ForeColor = Color.White };
            btnStop.Click += BtnStop_Click;

            // 配对区域 (保持不变)
            var lblPair = new Label { Text = "配对码:", Location = new Point(280, 58), AutoSize = true };
            var txtPair = new TextBox { Location = new Point(330, 55), Width = 100, Font = new Font("微软雅黑", 10) };
            var btnApprove = new Button { Text = "✅ 批准配对", Location = new Point(440, 50), Width = 120, Height = 40 };
            btnApprove.Click += (s, e) => ApprovePairing(txtPair.Text);

            logBox = new RichTextBox { Location = new Point(10, 100), Width = 840, Height = 450, ReadOnly = true, BackColor = Color.FromArgb(30, 30, 30), ForeColor = Color.LightGray, Font = new Font("Consolas", 10), BorderStyle = BorderStyle.None };

            pnlRun.Controls.AddRange(new Control[] { lblTip, btnStart, btnStop, lblPair, txtPair, btnApprove, logBox });
            tabRun.Controls.Add(pnlRun);
            tabControl.Controls.Add(tabRun);
        }

        // ================= 业务逻辑 =================
        private void LoadEnvironmentStatus()
        {   
            // === 1. 重置 UI 状态（禁用按钮防止误点）===
            // 这段代码在主线程运行，直接操作控件
            lblGitStatus.Text = "Git 状态: 检测中...";
            lblNodeStatus.Text = "Node.js 状态: 检测中...";
            lblClawStatus.Text = "OpenClaw 状态: 检测中...";

            // 【关键修改】检测期间禁用所有安装按钮
            btnInstallGit.Enabled = false;
            btnInstallNode.Enabled = false;
            btnInstallClaw.Enabled = false;

            Task.Run(async () =>
            {
                // 1. 并行获取基础环境状态
                var gitTask = Task.Run(() => _systemService.CheckCommandExists("git --version"));
                var nodeTask = Task.Run(() => _systemService.CheckCommandExists("node --version"));
                var clawTask = Task.Run(() => _systemService.CheckCommandExists("openclaw --version"));

                await Task.WhenAll(gitTask, nodeTask, clawTask);

                string gitVer = gitTask.Result;
                string nodeVer = nodeTask.Result;
                string clawVer = clawTask.Result;

                // 2. 异步获取 OpenClaw 最新版本 (联网操作)
                string latestClawVer = await _systemService.GetLatestNpmVersion("openclaw");

                // 3. 更新 UI (这里是在后台线程，必须使用 Invoke)
                // 加上 IsHandleCreated 判断，防止极端情况下窗口已关闭
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(() =>
                    {
                        // === Git 状态 ===
                        lblGitStatus.Text = $"Git 状态: {(gitVer != "未安装" ? "已安装" : "未安装")}";
                        lblGitStatus.ForeColor = gitVer != "未安装" ? Color.Green : Color.Red;
                        btnInstallGit.Enabled = gitVer == "未安装";

                        // === Node.js 状态 ===
                        bool isNodeMatch = (nodeVer == SystemService.TargetNodeVersion);
                        if (nodeVer == "未安装")
                        {
                            lblNodeStatus.Text = "Node.js 状态: 未安装";
                            lblNodeStatus.ForeColor = Color.Red;
                            btnInstallNode.Text = "1. 安装 Node.js";
                            btnInstallNode.Enabled = true;
                            btnInstallNode.BackColor = Color.LightBlue;
                        }
                        else if (isNodeMatch)
                        {
                            lblNodeStatus.Text = $"Node.js 状态: 已安装 ({nodeVer}) √";
                            lblNodeStatus.ForeColor = Color.Green;
                            btnInstallNode.Text = "1. Node.js 已就绪";
                            btnInstallNode.Enabled = false;
                            btnInstallNode.BackColor = Color.LightGray;
                        }
                        else
                        {
                            lblNodeStatus.Text = $"Node.js 状态: 已安装 ({nodeVer}) - 需更新为 {SystemService.TargetNodeVersion}";
                            lblNodeStatus.ForeColor = Color.OrangeRed;
                            btnInstallNode.Text = "1. 更新 Node.js 版本";
                            btnInstallNode.Enabled = true;
                            btnInstallNode.BackColor = Color.Orange;
                        }

                        // === OpenClaw 状态 ===
                        UpdateOpenClawUI(clawVer, latestClawVer, gitVer != "未安装" && isNodeMatch);
                    }));
                }
            });
        }

        /// <summary>
        /// 辅助方法：更新 OpenClaw 的 UI 状态
        /// </summary>
        private void UpdateOpenClawUI(string localVer, string latestVer, bool envReady)
        {
            bool isInstalled = (localVer != "未安装" && localVer != "检测失败");
            bool hasUpdate = false;

            // 简单的版本比较逻辑
            if (isInstalled && latestVer != "获取失败")
            {
                // 去掉可能存在的 'v' 前缀进行比较
                string localClean = localVer.TrimStart('v');
                string latestClean = latestVer.TrimStart('v');

                if (localClean != latestClean)
                {
                    hasUpdate = true;
                }
            }

            if (!isInstalled)
            {
                // 情况 A: 未安装
                lblClawStatus.Text = "OpenClaw 状态: 未安装";
                lblClawStatus.ForeColor = Color.Red;
                btnInstallClaw.Text = "2. 安装 OpenClaw";
                btnInstallClaw.BackColor = Color.LightGoldenrodYellow;
                btnInstallClaw.Enabled = envReady;
            }
            else if (hasUpdate)
            {
                // 情况 B: 有新版本
                lblClawStatus.Text = $"OpenClaw 状态: 当前 {localVer} (最新: {latestVer})";
                lblClawStatus.ForeColor = Color.OrangeRed; // 橙色提示更新
                btnInstallClaw.Text = "2. 更新 OpenClaw";
                btnInstallClaw.BackColor = Color.Orange; // 橙色按钮
                btnInstallClaw.Enabled = envReady;

                // 弹窗提示 (只在第一次检测到时弹出，或者直接让用户看状态即可)
                // Logger.Warn($"检测到 OpenClaw 有新版本 {latestVer}，建议更新。");
            }
            else
            {
                // 情况 C: 已是最新
                lblClawStatus.Text = $"OpenClaw 状态: 已安装 ({localVer}) √";
                lblClawStatus.ForeColor = Color.Green;
                btnInstallClaw.Text = "2. OpenClaw 已是最新";
                btnInstallClaw.BackColor = Color.LightGray;
                btnInstallClaw.Enabled = false;
            }
        }

        // 新增：Git 安装逻辑
        private async void BtnInstallGit_Click(object? sender, EventArgs e)
        {
            var btn = sender as Button;
            btn!.Enabled = false;
            Logger.Info("开始下载 Git 安装包...");

            string tempPath = Path.Combine(Path.GetTempPath(), "git-installer.exe");

            try
            {
                var progress = new Progress<(long received, long total)>(p => {
                    if (p.total > 0) progressBar.Value = (int)((p.received * 100) / p.total);
                });

                await _systemService.DownloadFileAsync(SystemService.GitDownloadUrl, tempPath, progress, CancellationToken.None);
                Logger.Info($"下载完成: {tempPath}");
                Logger.Info("正在静默安装 Git...");

                await _systemService.InstallGitAsync(tempPath, CancellationToken.None);

                Logger.Info("Git 安装流程结束。");
                MessageBox.Show("Git 安装成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error($"Git 安装失败: {ex.Message}");
                MessageBox.Show($"安装失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (File.Exists(tempPath)) File.Delete(tempPath);
                progressBar.Value = 0;
                LoadEnvironmentStatus();
            }
        }

        // 新增：Node.js 一键安装逻辑
        public async void BtnInstallNode_Click(object? sender, EventArgs e)
        {
            var btn = sender as Button;
            btn!.Enabled = false;

            // 判断是安装还是更新
            string action = btn.Text.Contains("更新") ? "更新" : "安装";
            Logger.Info($"开始下载 Node.js {SystemService.TargetNodeVersion} 安装包...");

            string tempPath = Path.Combine(Path.GetTempPath(), "node-installer.msi");

            try
            {
                // ... (下载代码保持不变) ...
                var progress = new Progress<(long received, long total)>(p => {
                    if (p.total > 0)
                    {
                        int percent = (int)((p.received * 100) / p.total);
                        progressBar.Value = percent;
                    }
                });

                await _systemService.DownloadFileAsync(SystemService.NodeJsDownloadUrl, tempPath, progress, CancellationToken.None);
                Logger.Info($"下载完成: {tempPath}");
                Logger.Info($"正在静默{action} Node.js...");

                await _systemService.InstallNodeJsAsync(tempPath, CancellationToken.None);

                Logger.Info($"Node.js {action}完成！");
                MessageBox.Show($"Node.js {action}成功！\n版本: {SystemService.TargetNodeVersion}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error($"{action}失败: {ex.Message}");
                MessageBox.Show($"{action}失败: {ex.Message}\n请尝试手动下载安装。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 清理临时文件
                if (File.Exists(tempPath)) File.Delete(tempPath);
                progressBar.Value = 0;
                LoadEnvironmentStatus(); // 刷新状态
            }
            //var btn = sender as Button;
            //btn!.Enabled = false;
            //AppendLog("开始下载 Node.js 安装包...");

            //string tempPath = Path.Combine(Path.GetTempPath(), "node-installer.msi");

            //try
            //{
            //    // 1. 下载
            //    var progress = new Progress<(long received, long total)>(p => {
            //        if (p.total > 0)
            //        {
            //            int percent = (int)((p.received * 100) / p.total);
            //            progressBar.Value = percent;
            //        }
            //    });

            //    await _systemService.DownloadFileAsync(SystemService.NodeJsDownloadUrl, tempPath, progress, CancellationToken.None);
            //    AppendLog($"下载完成: {tempPath}");
            //    AppendLog("正在静默安装 Node.js (可能需要几秒钟)...");

            //    // 2. 安装
            //    await _systemService.InstallNodeJsAsync(tempPath,  CancellationToken.None);
            //    //await _systemService.InstallNodeJsAsync(tempPath, AppendLog, CancellationToken.None);

            //    AppendLog("Node.js 安装完成！环境变量已自动刷新。");
            //    MessageBox.Show("Node.js 安装成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //catch (Exception ex)
            //{
            //    AppendLog($"安装失败: {ex.Message}");
            //    MessageBox.Show($"安装失败: {ex.Message}\n请尝试手动下载安装。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //finally
            //{
            //    // 清理临时文件
            //    if (File.Exists(tempPath)) File.Delete(tempPath);
            //    progressBar.Value = 0;
            //    LoadEnvironmentStatus();
            //}
        }

        // 新增：OpenClaw 一键安装逻辑
        private async void BtnInstallClaw_Click(object? sender, EventArgs e)
        {
            if (!_systemService.IsAdministrator())
            {
                MessageBox.Show("安装 OpenClaw 需要管理员权限。\n请关闭程序，右键点击程序图标选择“以管理员身份运行”。", "权限不足", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var btn = sender as Button;
            btn!.Enabled = false;

            Task.Run(async () =>
            {
                try
                {
                    await _systemService.InstallOpenClawAsync(CancellationToken.None);

                    // 成功提示
                    this.Invoke(new Action(() => {
                        MessageBox.Show("OpenClaw 安装流程结束，请查看日志确认是否成功。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
                catch (Exception ex)
                {
                    Logger.Error(ex.Message);
                    this.Invoke(new Action(() => {
                        MessageBox.Show($"安装出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    // 刷新界面状态
                    LoadEnvironmentStatus();
                }
            });
        }

        // 以下为之前的代码，保持不变
        private void BtnSaveConfig_Click(object? sender, EventArgs e)
        {
            string kimiKey = txtKimiKey.Text;
            string appId = txtAppId.Text;
            string appSecret = txtAppSecret.Text;

            if (string.IsNullOrWhiteSpace(kimiKey) || kimiKey == "sk-xxxxxxxxxxxxxxxx")
            {
                MessageBox.Show("请填写有效的 Kimi Key！");
                return;
            }

            try
            {
                // 调用更新后的方法
                if (_systemService.SaveConfigFile(kimiKey, appId, appSecret))
                {
                    MessageBox.Show("配置文件 已成功生成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private async void BtnStart_Click(object? sender, EventArgs e)
        {
            AppendLog("正在启动 OpenClaw Gateway...");

            // 更新按钮状态
            btnStart.Enabled = false;
            btnStop.Enabled = true;

            try
            {
                // 不需要 CancellationToken，让服务独立运行
                await _systemService.ExecuteCommandAsync("openclaw gateway --allow-unconfigured", AppendLog, CancellationToken.None);
            }
            catch (Exception ex)
            {
                AppendLog($"启动异常: {ex.Message}");
            }
            finally
            {
                // 只有在服务意外退出时才会走到这里
                // 如果是正常运行的，这个 Task 会一直等待，不会走这里
                this.Invoke(new Action(() => {
                    btnStart.Enabled = true;
                    btnStop.Enabled = false;
                }));
            }
        }

        private async void BtnStop_Click(object? sender, EventArgs e)
        {
            const string targetPort = "18789"; // 固定端口

            AppendLog($"正在查找端口 {targetPort} 占用情况...");
            btnStop.Enabled = false; // 防止重复点击

            try
            {
                // 1. 执行 netstat 命令查找 PID
                string netstatArgs = $"/c netstat -ano | findstr \"{targetPort}\"";
                var (netstatOutput, netstatCode) = await RunCmdAndGetOutput(netstatArgs);

                if (string.IsNullOrWhiteSpace(netstatOutput))
                {
                    AppendLog($"端口 {targetPort} 未被占用，服务可能已停止。");
                    btnStart.Enabled = true;
                    return;
                }

                // 2. 解析 PID
                int pid = -1;
                var lines = netstatOutput.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var line in lines)
                {
                    if (line.Contains("LISTENING"))
                    {
                        var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 0)
                        {
                            if (int.TryParse(parts[parts.Length - 1], out int foundPid))
                            {
                                pid = foundPid;
                                AppendLog($"找到监听进程 PID: {pid}");
                                break;
                            }
                        }
                    }
                }

                if (pid == -1)
                {
                    AppendLog("未能解析出 PID。");
                    btnStop.Enabled = true;
                    return;
                }

                // 3. 强制杀掉进程
                AppendLog($"正在执行强制终止进程 {pid}...");
                string killArgs = $"/c taskkill /PID {pid} /F";

                // === 关键修改：接收退出代码 ===
                var (killOutput, killExitCode) = await RunCmdAndGetOutput(killArgs);

                // 输出日志（现在应该不再是乱码了，但也无需解析它）
                AppendLog(killOutput.Trim());

                // === 通过退出代码判断成功 (0 代表成功) ===
                if (killExitCode == 0)
                {
                    AppendLog("服务已成功停止。");
                    btnStart.Enabled = true; // 恢复启动按钮
                }
                else
                {
                    AppendLog($"停止失败 (Exit Code: {killExitCode})，可能需要管理员权限。");
                    btnStop.Enabled = true; // 允许重试
                }
            }
            catch (Exception ex)
            {
                AppendLog($"停止操作出错: {ex.Message}");
                btnStop.Enabled = true;
            }
        }

        // === 辅助方法：执行 CMD 命令并返回输出字符串 ===
        private async Task<(string output, int exitCode)> RunCmdAndGetOutput(string arguments)
        {
            using (Process process = new Process())
            {
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.Arguments = arguments;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;

                // === 关键修复：使用系统默认编码 (GBK) 而不是 UTF-8，防止中文乱码 ===
                // 这里的 Default 在中文 Windows 上就是 GBK (Code Page 936)
                process.StartInfo.StandardOutputEncoding = System.Text.Encoding.Default;
                process.StartInfo.StandardErrorEncoding = System.Text.Encoding.Default;

                process.Start();

                string output = await process.StandardOutput.ReadToEndAsync();
                string error = await process.StandardError.ReadToEndAsync();

                process.WaitForExit();

                // 返回输出内容和退出代码
                return (output + error, process.ExitCode);
            }
        }

        private async void ApprovePairing(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) { MessageBox.Show("请输入配对码！"); return; }
            AppendLog($"正在批准配对码: {code} ...");
            await _systemService.ExecuteCommandAsync($"openclaw pairing approve {code}", AppendLog, CancellationToken.None);
        }

        private void AppendLog(string msg)
        {
            if (logBox.InvokeRequired) logBox.Invoke(new Action(() => AppendLog(msg)));
            else { logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\n"); logBox.ScrollToCaret(); }
        }
    }
}
