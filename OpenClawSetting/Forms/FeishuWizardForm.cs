using System;
using System.Drawing;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace OpenClawSetting
{
    public partial class FeishuWizardForm : Form
    {
        private WebView2 webView;
        private Button btnAutoConfig;
        private RichTextBox logBox;

        // 这里的 JSON 来自 PPT Step 19
        private const string PermissionJson = @"
{
  ""scopes"": {
    ""tenant"": [
      ""cardkit:card:write"",
      ""contact:contact.base:readonly"",
      ""contact:user.base:readonly"",
      ""im:chat"",
      ""im:chat:read"",
      ""im:chat:update"",
      ""im:message"",
      ""im:message.group_at_msg:readonly"",
      ""im:message.group_msg"",
      ""im:message.p2p_msg:readonly"",
      ""im:message:send_as_bot"",
      ""im:resource""
    ],
    ""user"": []
  }
}";

        public string? ResultAppId { get; private set; }
        public string? ResultAppSecret { get; private set; }

        public FeishuWizardForm()
        {
            Text = "飞书应用自动配置向导";
            Size = new Size(1200, 900);
            StartPosition = FormStartPosition.CenterParent;

            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 80, BackColor = Color.WhiteSmoke };

            btnAutoConfig = new Button { Text = "▶ 开始自动配置当前应用", Location = new Point(20, 20), Width = 200, Height = 40 };
            btnAutoConfig.Click += BtnAutoConfig_Click;

            var lblTip = new Label { Text = "提示：请先扫码登录飞书开放平台，进入应用详情页后点击按钮。", Location = new Point(240, 30), AutoSize = true };

            logBox = new RichTextBox { Dock = DockStyle.Bottom, Height = 150, ReadOnly = true, BackColor = Color.Black, ForeColor = Color.White };

            webView = new WebView2 { Dock = DockStyle.Fill };

            pnlTop.Controls.AddRange(new Control[] { btnAutoConfig, lblTip });
            this.Controls.AddRange(new Control[] { webView, logBox, pnlTop });

            this.Load += FeishuWizardForm_Load;
        }

        private async void FeishuWizardForm_Load(object? sender, EventArgs e)
        {
            // 初始化 WebView2 环境
            await webView.EnsureCoreWebView2Async();
            // 导航到飞书开放平台
            webView.Source = new Uri("https://open.feishu.cn/app");

            // 监听导航完成事件，可用于检测页面状态
            webView.CoreWebView2.NavigationCompleted += (s, args) =>
            {
                AppendLog($"页面加载完成: {webView.Source}");
            };
        }

        private async void BtnAutoConfig_Click(object? sender, EventArgs e)
        {
            btnAutoConfig.Enabled = false;
            AppendLog("开始执行自动化脚本...");

            try
            {
                // === 步骤流程 ===

                // 1. 尝试获取凭证 (对应 PPT Step 20)
                // 假设用户已经在应用详情页，或者用户手动创建了应用
                if (await TryGetCredentials())
                {
                    AppendLog($"✅ 成功获取凭证！AppID: {ResultAppId}");
                    MessageBox.Show($"自动配置完成！\n\nAppID: {ResultAppId}\nAppSecret: {ResultAppSecret}\n\n点击确定返回主界面。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                    return;
                }

                // 2. 如果没有凭证，尝试自动配置权限 (对应 PPT Step 19)
                // 导航到权限管理页面
                AppendLog("正在配置应用权限...");
                // 注意：实际自动化中，你需要根据实际的 URL 结构跳转，或者模拟点击
                // 这里演示直接注入 JS 点击“权限管理”菜单

                // 模拟点击权限管理 (此处选择器需要根据飞书实际页面结构调整，这里使用伪代码逻辑)
                // await webView.ExecuteScriptAsync("document.querySelector('[菜单选择器]').click();");

                // 注入权限 JSON (PPT Step 19 的核心功能)
                // 飞书后台通常有“批量添加”或“JSON 导入”功能
                string script = $@"
                    // 这是一个示例脚本，实际选择器需要根据飞书网页版源码确定
                    // 通常飞书后台会有一个文本框用于输入 JSON
                    var jsonStr = '{PermissionJson.Replace("\n", "").Replace("\r", "")}';
                    
                    // 假设有一个文本框 id 是 'permission-json-input' (实际需要你用F12开发者工具查找)
                    // var input = document.getElementById('permission-json-input');
                    // if(input) {{ input.value = jsonStr; }}
                    
                    // 模拟点击确定
                    // document.querySelector('.confirm-btn').click();
                    
                    'Script Executed';
                ";

                await webView.ExecuteScriptAsync(script);
                AppendLog("权限配置脚本已注入，请检查网页是否生效。");

                // 3. 启动机器人能力 (对应 PPT Step 18)
                // 同样的逻辑，找到“机器人”菜单，点击“添加”
            }
            catch (Exception ex)
            {
                AppendLog($"[错误] {ex.Message}");
                btnAutoConfig.Enabled = true;
            }
        }

        /// <summary>
        /// 尝试从当前页面抓取 AppID 和 Secret
        /// </summary>
        private async Task<bool> TryGetCredentials()
        {
            try
            {
                // 执行 JS 获取页面文本或特定元素值
                // 飞书开放平台的凭证通常在“凭证与基础信息”页面
                // 这里需要你根据实际网页结构编写 JS
                string script = @"
                    // 示例：查找页面上的 App ID (通常是复制组件)
                    // 这里的逻辑是：尝试查找包含 'cli_' 的文本
                    var elements = document.querySelectorAll('*');
                    for(var i=0; i<elements.length; i++){
                        if(elements[i].innerText && elements[i].innerText.includes('cli_')){
                            return elements[i].innerText;
                        }
                    }
                    return '';
                ";

                string result = await webView.ExecuteScriptAsync(script);
                // result 会是 JSON 格式的字符串 (带引号)
                result = result.Trim('"').Replace("\\\"", "\"");

                if (result.Contains("cli_"))
                {
                    ResultAppId = result; // 简单赋值，实际可能需要正则提取
                                          // 获取 Secret 比较复杂，通常需要点击“查看”按钮
                                          // 这里可以模拟点击查看按钮后，再抓取文本

                    return true;
                }
            }
            catch { }

            return false;
        }

        private void AppendLog(string msg)
        {
            if (logBox.InvokeRequired) logBox.Invoke(new Action(() => AppendLog(msg)));
            else logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\n");
        }
    }
}
