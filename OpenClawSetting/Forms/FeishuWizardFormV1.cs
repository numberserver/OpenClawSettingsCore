using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json; // 需要引用 Newtonsoft.Json 或 System.Text.Json

namespace OpenClawSetting
{
    public partial class FeishuWizardFormV1 : Form
    {
        private WebView2 webView;
        private TextBox txtAppName;
        private RichTextBox logBox;
        private Label lblStatus;

        // 步骤按钮
        private Button btnStep1, btnStep2, btnStep3, btnStep4, btnStep5;

        public string? ResultAppId { get; private set; }
        public string? ResultAppSecret { get; private set; }

        // PPT Step 19 权限 JSON
        private const string PermissionJson = @"{
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

        public FeishuWizardFormV1()
        {
            InitializeUI();
        }

        private void InitializeUI()
        {
            Text = "飞书应用分步配置向导";
            Size = new Size(1280, 900);
            StartPosition = FormStartPosition.CenterParent;

            // === 顶部控制面板 ===
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 130, BackColor = Color.WhiteSmoke, Padding = new Padding(10) };

            lblStatus = new Label { Text = "状态：初始化中...", Location = new Point(10, 10), AutoSize = true, Font = new Font("微软雅黑", 9, FontStyle.Bold) };

            // 第一行：应用名称
            var lblName = new Label { Text = "应用名称:", Location = new Point(10, 40), AutoSize = true };
            txtAppName = new TextBox { Location = new Point(80, 38), Width = 150, Text = "OpenClaw助手" };

            // 第二行：步骤按钮
            btnStep1 = new Button { Text = "1.打开控制台", Location = new Point(10, 70), Width = 110, Height = 35, BackColor = Color.LightGray };
            btnStep2 = new Button { Text = "2.创建应用", Location = new Point(130, 70), Width = 110, Height = 35, BackColor = Color.LightGray };
            btnStep3 = new Button { Text = "3.开启机器人", Location = new Point(250, 70), Width = 110, Height = 35, BackColor = Color.LightGray };
            btnStep4 = new Button { Text = "4.配置权限", Location = new Point(370, 70), Width = 110, Height = 35, BackColor = Color.LightGray };
            btnStep5 = new Button { Text = "5.获取凭证", Location = new Point(490, 70), Width = 110, Height = 35, BackColor = Color.LightGray };

            // 绑定事件
            btnStep1.Click += BtnStep1_Click;
            btnStep2.Click += BtnStep2_Click;
            btnStep3.Click += BtnStep3_Click;
            btnStep4.Click += BtnStep4_Click;
            btnStep5.Click += BtnStep5_Click;

            pnlTop.Controls.AddRange(new Control[] { lblStatus, lblName, txtAppName, btnStep1, btnStep2, btnStep3, btnStep4, btnStep5 });

            // === 底部日志 ===
            logBox = new RichTextBox { Dock = DockStyle.Bottom, Height = 120, ReadOnly = true, BackColor = Color.Black, ForeColor = Color.LightGray, Font = new Font("Consolas", 9) };

            // === 浏览器主体 ===
            webView = new WebView2 { Dock = DockStyle.Fill };

            Controls.AddRange(new Control[] { webView, logBox, pnlTop });

            // 加载时初始化浏览器
            this.Load += async (s, e) =>
            {
                try
                {
                    // 初始化 WebView2 环境
                    await webView.EnsureCoreWebView2Async();
                    // 导航到飞书开放平台
                    webView.Source = new Uri("https://open.feishu.cn/app");

                    // 监听导航完成事件，可用于检测页面状态
                    webView.CoreWebView2.NavigationCompleted += (s, args) =>
                    {
                        Log($"页面加载完成: {webView.Source}");
                    };
                    //await webView.EnsureCoreWebView2Async();
                    lblStatus.Text = "状态：浏览器就绪，请按步骤操作。";
                    webView.CoreWebView2.SourceChanged += (s, args) => Log($"页面跳转: {webView.Source}");
                }
                catch (Exception ex)
                {
                    Log($"初始化失败: {ex.Message}");
                    lblStatus.Text = "状态：初始化失败，请检查 WebView2 运行时。";
                }
            };
        }

        // === Step 1: 打开控制台 ===
        private void BtnStep1_Click(object? sender, EventArgs e)
        {
            Log("Step 1: 导航到飞书开放平台...");
            webView.Source = new Uri("https://open.feishu.cn/app");
        }

        // === Step 2: 创建应用 ===
        private void BtnStep2_Click(object? sender, EventArgs e)
        {
            string appName = txtAppName.Text;
            Log($"Step 2: 尝试创建应用 [{appName}]...");

            // 脚本逻辑：点击创建 -> 等待弹窗 -> 填写名称 -> 点击确定
            string script = $@"
                (function() {{
                    var createBtn = [...document.querySelectorAll('button')].find(b => b.innerText.includes('创建企业自建应用'));
                    if (createBtn) {{
                        createBtn.click();
                        
                        // 等待弹窗出现 (JS内部等待，更可靠)
                        setTimeout(function() {{
                            var inputs = document.querySelectorAll('input[type=""text""]');
                            if(inputs.length > 0) {{
                                // 设置应用名称
                                var nameInput = inputs[0];
                                nameInput.value = '{appName}';
                                nameInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                
                                // 点击确定按钮
                                setTimeout(function() {{
                                    var confirmBtn = [...document.querySelectorAll('button')].find(b => b.innerText.includes('确定') || b.innerText.includes('创建'));
                                    if(confirmBtn) confirmBtn.click();
                                }}, 300);
                            }}
                        }}, 500);
                        return '已点击创建并填写名称';
                    }}
                    return '未找到创建按钮，请确认是否在应用列表页';
                }})();
            ";
            ExecuteScriptAsync(script);
        }

        // === Step 3: 开启机器人 ===
        private void BtnStep3_Click(object? sender, EventArgs e)
        {
            Log("Step 3: 尝试开启机器人能力...");

            // 脚本逻辑：点击应用能力菜单 -> 点击机器人 -> 点击添加
            string script = $@"
                (function() {{
                    // 1. 尝试点击 '应用能力' 菜单 (左侧或顶部导航)
                    var menuItems = document.querySelectorAll('div');
                    var capMenu = null;
                    for(var i=0; i<menuItems.length; i++) {{
                        if(menuItems[i].innerText === '应用能力') {{
                            menuItems[i].click();
                            capMenu = menuItems[i];
                            break;
                        }}
                    }}

                    // 无论是否找到菜单，尝试直接点击 '机器人' (有时是卡片)
                    setTimeout(function() {{
                        var botCards = document.querySelectorAll('div');
                        for(var i=0; i<botCards.length; i++) {{
                            if(botCards[i].innerText.includes('机器人') && botCards[i].innerText.includes('添加')) {{
                                botCards[i].click();
                                
                                // 可能会有确认弹窗
                                setTimeout(function() {{
                                    var addBtn = [...document.querySelectorAll('button')].find(b => b.innerText.includes('添加') || b.innerText.includes('确定'));
                                    if(addBtn) addBtn.click();
                                }}, 300);
                                
                                return '已开启机器人';
                            }}
                        }}
                        
                        // 如果上面没找到，尝试点击独立的添加按钮
                         var addBtns = [...document.querySelectorAll('button')].filter(b => b.innerText.includes('添加能力'));
                         if(addBtns.length > 0) addBtns[0].click();
                         return '正在尝试添加能力...';

                    }}, 500);
                    
                    return '正在查找机器人配置...';
                }})();
            ";
            ExecuteScriptAsync(script);
        }

        // === Step 4: 配置权限 (优化版) ===
        private async void BtnStep4_Click(object? sender, EventArgs e)
        {
            //        Log("Step 4: 正在进入权限管理页面...");

            //        // 1. 优先查找并点击“权限管理”菜单
            //        string navScript = @"
            //    (function() {
            //        // 查找包含 '权限管理' 的元素
            //        var elements = document.querySelectorAll('div, span, a');
            //        for(var i=0; i<elements.length; i++){
            //            var text = elements[i].innerText;
            //            // 精确匹配或包含关系，排除过长文本
            //            if (text && (text === '权限管理' || (text.includes('权限管理') && text.length < 10))) {
            //                // 滚动到可见区域
            //                elements[i].scrollIntoView({block: 'center'});
            //                elements[i].click();
            //                return '已点击菜单，正在跳转...';
            //            }
            //        }

            //        // 如果没找到，尝试查找侧边栏的特定结构 (备用方案)
            //        // 这里可以根据实际F12看到的标签微调
            //        return '未找到菜单，请确认是否已在应用详情页';
            //    })();
            //";

            //        // 执行点击
            //        ExecuteScriptAsync(navScript);

            //        // 2. 延迟执行权限配置 (给页面加载留出时间)
            //        // 注意：这里我们必须等待一段时间，因为点击菜单后页面会刷新
            //        Task.Delay(2000).ContinueWith(t =>
            //        {
            //            this.Invoke(new Action(() =>
            //            {
            //                Log("页面已加载，尝试配置权限...");
            //                InjectPermissionConfig();
            //            }));
            //        });
            Log("Step 4: 准备配置权限...");

            // 1. 尝试点击左侧菜单 "权限管理"
            string navScript = @"
        (function() {
            var elements = document.querySelectorAll('div, span, a');
            for(var i=0; i<elements.length; i++){
                var text = elements[i].innerText;
                if (text && (text === '权限管理' || (text.includes('权限管理') && text.length < 10))) {
                    elements[i].scrollIntoView({block: 'center'});
                    elements[i].click();
                    return '已点击菜单';
                }
            }
            return '未找到菜单';
        })();
    ";
            await ExecuteScriptAsync(navScript);

            // 2. 等待页面加载
            await Task.Delay(1500);

            // 3. 点击“批量导入/导出权限”按钮
            string clickBtnScript = @"
        (function() {
            var batchBtn = [...document.querySelectorAll('button')].find(b => b.innerText.includes('批量导入/导出权限'));
            if (batchBtn) {
                batchBtn.click();
                return '已打开导入界面';
            }
            return '未找到按钮';
        })();
    ";
            string result = await ExecuteScriptAsync(clickBtnScript);

            // 4. 将 JSON 复制到剪贴板并提示用户
            try
            {
                // 这里的 PermissionJson 是类顶部定义的那个常量
                Clipboard.SetText(PermissionJson);

                MessageBox.Show(
                    "已自动点击“批量导入”按钮。\n\n" +
                    "权限配置 JSON 已复制到您的剪贴板！\n\n" +
                    "请直接在右侧网页输入框中按 Ctrl+V 粘贴，\n" +
                    "然后点击网页上的确定按钮。",
                    "请手动粘贴",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                Log($"剪贴板操作失败: {ex.Message}");
            }
        }

        // 独立的权限配置注入方法
        // 独立的权限配置注入方法 (优化版：适配直接录入模式)
        private async void InjectPermissionConfig()
        {
            string jsonInput = PermissionJson.Replace("\r", "").Replace("\n", "");

            // 脚本逻辑：
            // 1. 点击“批量导入/导出权限”按钮
            // 2. 短暂等待页面渲染输入区域
            // 3. 查找 textarea 或 可编辑div 并录入
            // 4. 点击确定
            string script = $@"
        (function() {{
            var jsonStr = '{jsonInput}';
            
            // 1. 查找并点击按钮
            var batchBtn = [...document.querySelectorAll('button')].find(b => b.innerText.includes('批量导入/导出权限'));
            
            if (batchBtn) {{
                batchBtn.click();
                console.log('已点击批量导入按钮');
                
                // 2. 等待输入区域出现 (给一点渲染时间)
                setTimeout(function() {{
                    // 3. 查找输入框 (兼容 textarea 和 contenteditable div)
                    var inputArea = document.querySelector('textarea');
                    if (!inputArea) {{
                        // 如果没找到 textarea，尝试找可编辑的 div
                        inputArea = document.querySelector('div[contenteditable=""true""]');
                    }}
                    
                    if (inputArea) {{
                        // 赋值
                        inputArea.value = jsonStr; // 对 textarea 有效
                        inputArea.innerText = jsonStr; // 对 div 有效
                        
                        // 触发输入事件，让飞书前端识别到变化
                        inputArea.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        inputArea.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        inputArea.dispatchEvent(new Event('blur', {{ bubbles: true }})); // 有时需要失焦触发校验
                        console.log('已填充 JSON 数据');

                        // 4. 尝试点击确定按钮 (通常在输入框下方)
                        setTimeout(function() {{
                            // 查找包含“确定”或“导入”的按钮
                            var confirmBtn = [...document.querySelectorAll('button')].find(b => 
                                (b.innerText.includes('确定') || b.innerText.includes('导入')) && 
                                !b.innerText.includes('批量') // 排除刚才点过的按钮
                            );
                            
                            if (confirmBtn) {{
                                confirmBtn.click();
                                console.log('已点击确定，配置完成');
                            }} else {{
                                console.log('未找到确定按钮，可能需要手动点击');
                            }}
                        }}, 300);
                        
                    }} else {{
                        console.log('未找到输入区域');
                        alert('未找到输入区域，请检查页面元素');
                    }}
                }}, 300); // 等待 300ms 即可，因为是页面内渲染
            }} else {{
                console.log('未找到批量导入按钮');
                alert('未找到“批量导入”按钮！');
            }}
        }})();
    ";

            await ExecuteScriptAsync(script);
        }
        //    private async void InjectPermissionConfig()
        //    {
        //        string jsonInput = PermissionJson.Replace("\r", "").Replace("\n", "");

        //        // 脚本逻辑：点击批量添加 -> 填写JSON -> 确定
        //        string script = $@"
        //    (function() {{
        //        var jsonStr = '{jsonInput}';

        //        // 查找“批量添加”按钮
        //        var batchBtn = [...document.querySelectorAll('button')].find(b => b.innerText.includes('批量导入/导出权限'));

        //        if (batchBtn) {{
        //            Log('找到批量添加按钮，正在点击...');
        //            batchBtn.click();

        //            // 等待弹窗出现
        //            setTimeout(function() {{
        //                var textarea = document.querySelector('textarea');
        //                if (textarea) {{
        //                    textarea.value = jsonStr;
        //                    textarea.dispatchEvent(new Event('input', {{ bubbles: true }}));
        //                    Log('已填充 JSON 数据');

        //                    // 点击确定
        //                    setTimeout(function() {{
        //                        var confirmBtn = [...document.querySelectorAll('button')].find(b => b.innerText.includes('确定') || b.innerText.includes('添加'));
        //                        if(confirmBtn) {{
        //                            confirmBtn.click();
        //                            Log('权限配置完成！');
        //                        }}
        //                    }}, 300);
        //                }} else {{
        //                    Log('未找到输入框');
        //                }}
        //            }}, 800);
        //        }} else {{
        //            // 如果找不到按钮，可能是页面结构变了，或者没加载出来
        //            alert('未找到“批量添加”按钮！\n请确认：\n1. 已手动进入权限管理页面。\n2. 页面已完全加载。\n\n如果没有批量添加按钮，请手动添加权限。');
        //        }}

        //        function Log(msg) {{ console.log(msg); return msg; }}
        //    }})();
        //";

        //        await ExecuteScriptAsync(script);
        //    }

        //// === Step 4: 配置权限 ===
        //private void BtnStep4_Click(object? sender, EventArgs e)
        //{
        //    Log("Step 4: 尝试配置权限...");
        //    string jsonInput = PermissionJson.Replace("\r", "").Replace("\n", "");

        //    // 脚本逻辑：点击权限管理 -> 点击批量添加 -> 填写JSON -> 确定
        //    string script = $@"
        //        (function() {{
        //            var jsonStr = '{jsonInput}';

        //            // 1. 点击权限管理菜单
        //            var menuItems = document.querySelectorAll('div');
        //            for(var i=0; i<menuItems.length; i++){{
        //                    alert(menuItems[i].innerText);
        //                if(menuItems[i].innerText.includes('权限管理') && menuItems[i].innerText.length < 10){{
        //                    menuItems[i].click();
        //                    break;
        //                }}
        //            }}

        //            // 2. 等待页面加载后，点击批量添加
        //            setTimeout(function() {{
        //                var batchBtn = [...document.querySelectorAll('button')].find(b => b.innerText.includes('批量添加'));
        //                if (batchBtn) {{
        //                    batchBtn.click();

        //                    // 3. 填写 JSON
        //                    setTimeout(function() {{
        //                        var textarea = document.querySelector('textarea');
        //                        if (textarea) {{
        //                            textarea.value = jsonStr;
        //                            textarea.dispatchEvent(new Event('input', {{ bubbles: true }}));

        //                            // 4. 点击确认
        //                            setTimeout(function() {{
        //                                var confirmBtn = [...document.querySelectorAll('button')].find(b => b.innerText.includes('确定') || b.innerText.includes('添加'));
        //                                if(confirmBtn) confirmBtn.click();
        //                            }}, 300);
        //                        }}
        //                    }}, 500);
        //                }} else {{
        //                    alert('未找到“批量添加”按钮，请检查是否已在权限管理页面');
        //                }}
        //            }}, 800);

        //            return '正在配置权限...';
        //        }})();
        //    ";
        //    ExecuteScriptAsync(script);
        //}

        // === Step 5: 获取凭证 ===
        private async void BtnStep5_Click(object? sender, EventArgs e)
        {
            Log("Step 5: 正在获取凭证...");

            // 1. 先导航到凭证页面
            string navScript = @"
                (function() {
                    var menuItems = document.querySelectorAll('div');
                    for(var i=0; i<menuItems.length; i++){

 alter(menuItems[i].innerText);
                        if(menuItems[i].innerText.includes('应用凭证') && menuItems[i].innerText.length < 15){
                            menuItems[i].click();
                            return 'Navigating...';
                        }
                    }
                    return 'Menu not found';
                })();
            ";
            await ExecuteScriptAsync(navScript);

            // 2. 给页面一点时间跳转，然后尝试抓取
            // 注意：这里用一个小延时是为了等待页面刷新，JS之间很难处理这种跨页面跳转
            await Task.Delay(1500);

            // 3. 尝试点击“查看” Secret 按钮
            string viewScript = @"
                (function() {
                    var buttons = document.querySelectorAll('span');
                    for(var i=0; i<buttons.length; i++){
                        if(buttons[i].innerText.includes('查看') || buttons[i].innerText.includes('显示')){
                            buttons[i].click();
                            return 'Clicked View Secret';
                        }
                    }
                    return 'View Button Not Found';
                })();
            ";
            await ExecuteScriptAsync(viewScript);

            // 4. 抓取内容
            await Task.Delay(1000);
            string scrapeScript = @"
                (function() {
                    var result = { appId: '', appSecret: '' };
                    var allText = document.body.innerText;
                    
                    // 提取 App ID (cli_xxx)
                    var idMatch = allText.match(/(cli_[a-zA-Z0-9]+)/);

                    if(idMatch) result.appId = idMatch[0];
 Log(idMatch); Log(allText);

                    // 提取 Secret (通常在点击查看后，很难精确抓取，这里尝试找长字符串)
                    // 如果用户手动复制了 Secret，我们也可以尝试读取剪贴板，但这里先尝试抓取页面显示
                    // 注意：这是基于当前页面结构的简单尝试

                    var inputs = document.querySelectorAll('input');
                    for(var i=0; i<inputs.length; i++){
                        var val = inputs[i].value;
                        if(val.length > 20 && !val.includes('cli_')) {
                            result.appSecret = val;
                        }
                    }
                    
                    return JSON.stringify(result);
                })();
            ";

            string resultStr = await webView.ExecuteScriptAsync(scrapeScript);
            Log($"抓取原始结果: {resultStr}");

            try
            {
                // 解析 JSON 结果
                var jsonDoc = System.Text.Json.JsonDocument.Parse(resultStr);
                var root = jsonDoc.RootElement;

                string? id = root.GetProperty("appId").GetString();

                if (!string.IsNullOrEmpty(id))
                {
                    ResultAppId = id;
                    Log($"✅ App ID: {ResultAppId}");
                    lblStatus.Text = $"已获取 App ID: {ResultAppId}";

                    // 提示用户手动处理 Secret (最稳妥)
                    MessageBox.Show($"已自动获取 App ID: {ResultAppId}\n\n" +
                                    "App Secret 可能需要手动复制。\n" +
                                    "请在右侧网页中查看 Secret，如果已显示，请手动复制到主界面。\n\n" +
                                    "点击确定关闭此窗口。", "操作完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("自动抓取失败，请手动复制 App ID 和 Secret。", "提示");
                }
            }
            catch (Exception ex)
            {
                Log($"解析结果失败: {ex.Message}");
                MessageBox.Show("自动抓取失败，请手动复制凭证。");
            }
        }

        // 辅助方法：执行脚本并记录日志
        // 辅助方法：执行脚本 (确保你的类里已经有这个方法，如果没有请加上)
        private async Task<string> ExecuteScriptAsync(string script)
        {
            try
            {
                if (webView.CoreWebView2 != null)
                {
                    string result = await webView.ExecuteScriptAsync(script);
                    // 去掉 JSON 结果的引号
                    string cleanResult = result.Trim('"').Replace("\\r\\n", "").Replace("\\n", "");
                    Log($"脚本执行: {cleanResult}");
                    return cleanResult;
                }
            }
            catch (Exception ex)
            {
                Log($"脚本错误: {ex.Message}");
            }
            return "";
        }
        //private async Task ExecuteScriptAsync(string script)
        //{
        //    try
        //    {
        //        if (webView.CoreWebView2 != null)
        //        {
        //            string result = await webView.ExecuteScriptAsync(script);
        //            Log($"执行结果: {result.Trim('"')}");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Log($"脚本执行错误: {ex.Message}");
        //    }
        //}

        private void Log(string msg)
        {
            if (logBox.InvokeRequired)
            {
                logBox.Invoke(new Action(() => Log(msg)));
            }
            else
            {
                logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\n");
            }
        }
    }
}
