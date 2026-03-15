namespace OpenClawSetting
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration. // 配置应用程序视觉样式
            ApplicationConfiguration.Initialize();
            // 运行主窗体
            Application.Run(new MainForm());
        }
    }
}