using System;
using System.Diagnostics;
using System.Security.Principal;
using System.Threading;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // 最早設置異常處理 - 在任何其他代碼之前
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            // 檢查管理員權限
            if (!IsAdministrator())
            {
                MessageBox.Show("此應用程式需要管理員權限才能正常運行。\n\n即將以管理員身分重新啟動...", 
                    "需要管理員權限", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RestartAsAdministrator();
                return;
            }

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }

        /// <summary>
        /// UI 線程異常處理 - 抑制 CyberGroupBox 繪製錯誤
        /// </summary>
        private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            // 檢查是否為 CyberGroupBox 的繪製錯誤
            if (IsCyberGroupBoxDrawingError(e.Exception))
            {
                // 靜默忽略此錯誤，不顯示對話框
                Debug.WriteLine($"[Suppressed] CyberGroupBox drawing error: {e.Exception.Message}");
                return;
            }

            // 其他異常顯示錯誤對話框
            MessageBox.Show($"發生未預期的錯誤：\n{e.Exception.Message}\n\n{e.Exception.StackTrace}", 
                "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 非 UI 線程異常處理
        /// </summary>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                if (IsCyberGroupBoxDrawingError(ex))
                {
                    Debug.WriteLine($"[Suppressed] CyberGroupBox drawing error: {ex.Message}");
                    return;
                }

                MessageBox.Show($"發生嚴重錯誤：\n{ex.Message}", "嚴重錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 檢查是否為 CyberGroupBox 繪製相關的錯誤
        /// </summary>
        private static bool IsCyberGroupBoxDrawingError(Exception ex)
        {
            if (ex == null) return false;

            string stackTrace = ex.StackTrace ?? "";
            string message = ex.Message ?? "";

            // 檢查是否為 CyberGroupBox 或 ReaLTaiizor 繪製錯誤
            bool isReaLTaiizorError = stackTrace.Contains("CyberGroupBox") || 
                                      stackTrace.Contains("Draw_Background") ||
                                      stackTrace.Contains("ReaLTaiizor") ||
                                      stackTrace.Contains("Cyber");

            bool isDrawingError = ex is ArgumentException || 
                                  ex is InvalidOperationException ||
                                  message.Contains("Parameter is not valid") ||
                                  stackTrace.Contains("DrawPath") ||
                                  stackTrace.Contains("GraphicsPath") ||
                                  stackTrace.Contains("OnPaint") ||
                                  stackTrace.Contains("Draw");

            return isReaLTaiizorError && isDrawingError;
        }

        /// <summary>
        /// 檢查是否以管理員身分運行
        /// </summary>
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

        /// <summary>
        /// 以管理員身分重新啟動應用程式
        /// </summary>
        private static void RestartAsAdministrator()
        {
            try
            {
                var proc = new ProcessStartInfo
                {
                    UseShellExecute = true,
                    FileName = System.Reflection.Assembly.GetExecutingAssembly().Location,
                    Verb = "runas"
                };
                Process.Start(proc);
                Environment.Exit(0);
            }
            catch
            {
                MessageBox.Show("無法以管理員身分啟動應用程式。\n請手動以管理員身分運行此程式。", 
                    "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
