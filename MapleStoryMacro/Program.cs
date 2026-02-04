namespace MapleStoryMacro
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // 檢查管理員權限
            if (!IsAdministrator())
            {
                MessageBox.Show("此應用程序需要管理員權限才能正常運行。\n\n即將以管理員身份重新啟動...", "需要管理員權限", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RestartAsAdministrator();
                return;
            }

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }

        /// <summary>
        /// 檢查是否以管理員身份運行
        /// </summary>
        private static bool IsAdministrator()
        {
            try
            {
                var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 以管理員身份重新啟動應用程序
        /// </summary>
        private static void RestartAsAdministrator()
        {
            try
            {
                var proc = new System.Diagnostics.ProcessStartInfo();
                proc.UseShellExecute = true;
                proc.FileName = System.Reflection.Assembly.GetExecutingAssembly().Location;
                proc.Verb = "runas";
                System.Diagnostics.Process.Start(proc);
                Environment.Exit(0);
            }
            catch
            {
                MessageBox.Show("無法以管理員身份啟動應用程序。\n請手動以管理員身份運行此程序。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}