using System;
using System.Diagnostics;
using System.Security.Principal;

namespace MapleStoryMacro
{
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(System.Windows.StartupEventArgs e)
        {
            // 異常處理
            DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            // 檢查管理員權限
            if (!IsAdministrator())
            {
                System.Windows.MessageBox.Show(
                    "此應用程式需要管理員權限才能正常運行。\n\n即將以管理員身分重新啟動...",
                    "需要管理員權限", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                RestartAsAdministrator();
                Shutdown();
                return;
            }

            base.OnStartup(e);
        }

        private void App_DispatcherUnhandledException(object sender,
            System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            Debug.WriteLine($"[UI Exception] {e.Exception.Message}");
            System.Windows.MessageBox.Show($"發生未預期的錯誤：\n{e.Exception.Message}\n\n{e.Exception.StackTrace}",
                "錯誤", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            e.Handled = true;
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                System.Windows.MessageBox.Show($"發生嚴重錯誤：\n{ex.Message}",
                    "嚴重錯誤", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        private static bool IsAdministrator()
        {
            try
            {
                var identity = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        private static void RestartAsAdministrator()
        {
            try
            {
                var exePath = Environment.ProcessPath;
                if (string.IsNullOrEmpty(exePath))
                    exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

                var proc = new ProcessStartInfo
                {
                    UseShellExecute = true,
                    FileName = exePath,
                    Verb = "runas"
                };
                Process.Start(proc);
                Environment.Exit(0);
            }
            catch
            {
                System.Windows.MessageBox.Show("無法以管理員身分啟動應用程式。\n請手動以管理員身分運行此程式。",
                    "錯誤", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
    }
}
