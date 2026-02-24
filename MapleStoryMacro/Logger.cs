using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MapleStoryMacro
{
    /// <summary>
    /// 日誌記錄器 - 提供完整的日誌功能
    /// 支援多線程、異步寫入、自動清理舊日誌
    /// </summary>
    public static class Logger
    {
        #region 配置

        private static readonly string LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MapleStoryMacro", "logs");

        private static bool _isInitialized = false;
        private static bool _enableDetailedLogging = false;
        private static int _logRetentionDays = 7;
        private static ConcurrentQueue<LogEntry> _logQueue = new ConcurrentQueue<LogEntry>();
        private static CancellationTokenSource _cts = new CancellationTokenSource();
        private static Task? _writerTask;
        private static readonly object _lockObj = new object();

        #endregion

        #region 日誌等級

        public enum LogLevel
        {
            Debug = 0,    // 詳細的調試資訊
            Info = 1,     // 一般資訊
            Warning = 2,  // 警告訊息
            Error = 3,    // 錯誤訊息
            Critical = 4  // 嚴重錯誤
        }

        #endregion

        #region 日誌條目

        private class LogEntry
        {
            public DateTime Timestamp { get; set; }
            public LogLevel Level { get; set; }
            public string Message { get; set; } = string.Empty;
            public Exception? Exception { get; set; }
            public string? StackTrace { get; set; }
            public string ThreadInfo { get; set; } = string.Empty;
        }

        #endregion

        #region 初始化和配置

        /// <summary>
        /// 初始化日誌系統
        /// </summary>
        public static void Initialize(bool enableDetailedLogging = false, int logRetentionDays = 7)
        {
            if (_isInitialized)
                return;

            lock (_lockObj)
            {
                if (_isInitialized)
                    return;

                _enableDetailedLogging = enableDetailedLogging;
                _logRetentionDays = logRetentionDays;

                try
                {
                    // 確保日誌目錄存在
                    Directory.CreateDirectory(LogDirectory);

                    // 清理舊日誌
                    CleanupOldLogs();

                    // 啟動異步寫入線程
                    _writerTask = Task.Run(() => LogWriterLoop(_cts.Token), _cts.Token);

                    _isInitialized = true;

                    Info("日誌系統已初始化");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[Logger] 初始化失敗: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 關閉日誌系統
        /// </summary>
        public static void Shutdown()
        {
            if (!_isInitialized)
                return;

            Info("日誌系統正在關閉...");

            _cts.Cancel();

            // 等待寫入完成
            _writerTask?.Wait(TimeSpan.FromSeconds(5));

            // 寫入剩餘的日誌
            FlushLogs();

            _isInitialized = false;
        }

        #endregion

        #region 日誌方法

        /// <summary>
        /// 記錄 Debug 等級日誌（只在啟用詳細日誌時記錄）
        /// </summary>
        public static void Debug(string message)
        {
            if (_enableDetailedLogging)
            {
                Log(LogLevel.Debug, message);
            }
        }

        /// <summary>
        /// 記錄 Info 等級日誌
        /// </summary>
        public static void Info(string message)
        {
            Log(LogLevel.Info, message);
        }

        /// <summary>
        /// 記錄 Warning 等級日誌
        /// </summary>
        public static void Warning(string message, Exception? ex = null)
        {
            Log(LogLevel.Warning, message, ex);
        }

        /// <summary>
        /// 記錄 Error 等級日誌
        /// </summary>
        public static void Error(string message, Exception? ex = null)
        {
            Log(LogLevel.Error, message, ex);
        }

        /// <summary>
        /// 記錄 Critical 等級日誌
        /// </summary>
        public static void Critical(string message, Exception? ex = null)
        {
            Log(LogLevel.Critical, message, ex);
        }

        /// <summary>
        /// 核心日誌方法
        /// </summary>
        private static void Log(LogLevel level, string message, Exception? ex = null, string? stackTrace = null)
        {
            if (!_isInitialized)
            {
                Initialize(); // 自動初始化
            }

            try
            {
                var entry = new LogEntry
                {
                    Timestamp = DateTime.Now,
                    Level = level,
                    Message = message,
                    Exception = ex,
                    StackTrace = stackTrace ?? (level >= LogLevel.Error ? Environment.StackTrace : null),
                    ThreadInfo = $"Thread-{Thread.CurrentThread.ManagedThreadId}"
                };

                _logQueue.Enqueue(entry);

                // 同時輸出到 Debug（開發時可見）
                System.Diagnostics.Debug.WriteLine(FormatLogEntry(entry, false));

                // 嚴重錯誤立即刷新
                if (level >= LogLevel.Critical)
                {
                    FlushLogs();
                }
            }
            catch
            {
                // 日誌系統本身的錯誤不應影響程式運行
            }
        }

        #endregion

        #region 異步寫入

        /// <summary>
        /// 異步日誌寫入循環
        /// </summary>
        private static void LogWriterLoop(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    if (_logQueue.IsEmpty)
                    {
                        Thread.Sleep(100); // 空閒等待
                        continue;
                    }

                    FlushLogs();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[Logger] 寫入錯誤: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 將隊列中的日誌寫入文件
        /// </summary>
        private static void FlushLogs()
        {
            if (_logQueue.IsEmpty)
                return;

            var batch = new System.Collections.Generic.List<LogEntry>();
            while (_logQueue.TryDequeue(out var entry))
            {
                batch.Add(entry);
                if (batch.Count >= 100) // 批量寫入
                    break;
            }

            if (batch.Count == 0)
                return;

            try
            {
                string fileName = $"log_{DateTime.Now:yyyyMMdd}.txt";
                string filePath = Path.Combine(LogDirectory, fileName);

                var sb = new StringBuilder();
                foreach (var entry in batch)
                {
                    sb.AppendLine(FormatLogEntry(entry, true));
                }

                // 追加寫入（線程安全）
                File.AppendAllText(filePath, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] 寫入文件失敗: {ex.Message}");
            }
        }

        #endregion

        #region 格式化

        /// <summary>
        /// 格式化日誌條目
        /// </summary>
        private static string FormatLogEntry(LogEntry entry, bool includeStackTrace)
        {
            var sb = new StringBuilder();

            // 時間戳記 + 等級
            sb.Append($"[{entry.Timestamp:yyyy-MM-dd HH:mm:ss.fff}] ");
            sb.Append($"[{GetLevelIcon(entry.Level)} {entry.Level,-8}] ");

            // 線程資訊（Debug 模式）
            if (_enableDetailedLogging)
            {
                sb.Append($"[{entry.ThreadInfo}] ");
            }

            // 訊息
            sb.Append(entry.Message);

            // 異常資訊
            if (entry.Exception != null)
            {
                sb.AppendLine();
                sb.AppendLine($"  異常類型: {entry.Exception.GetType().Name}");
                sb.AppendLine($"  異常訊息: {entry.Exception.Message}");

                if (includeStackTrace && !string.IsNullOrEmpty(entry.Exception.StackTrace))
                {
                    sb.AppendLine($"  堆疊追蹤:");
                    foreach (var line in entry.Exception.StackTrace.Split('\n'))
                    {
                        sb.AppendLine($"    {line.TrimEnd()}");
                    }
                }

                // 內部異常
                if (entry.Exception.InnerException != null)
                {
                    sb.AppendLine($"  內部異常: {entry.Exception.InnerException.Message}");
                }
            }

            // 堆疊追蹤（非異常情況）
            if (includeStackTrace && entry.Exception == null && !string.IsNullOrEmpty(entry.StackTrace))
            {
                sb.AppendLine($"  堆疊追蹤:");
                foreach (var line in entry.StackTrace.Split('\n'))
                {
                    if (line.Contains("at MapleStoryMacro.")) // 只顯示相關的堆疊
                    {
                        sb.AppendLine($"    {line.TrimEnd()}");
                    }
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// 取得日誌等級圖示
        /// </summary>
        private static string GetLevelIcon(LogLevel level)
        {
            return level switch
            {
                LogLevel.Debug => "🔍",
                LogLevel.Info => "ℹ️",
                LogLevel.Warning => "⚠️",
                LogLevel.Error => "❌",
                LogLevel.Critical => "🔥",
                _ => "  "
            };
        }

        #endregion

        #region 工具方法

        /// <summary>
        /// 清理舊日誌文件
        /// </summary>
        private static void CleanupOldLogs()
        {
            try
            {
                if (!Directory.Exists(LogDirectory))
                    return;

                var cutoffDate = DateTime.Now.AddDays(-_logRetentionDays);
                var files = Directory.GetFiles(LogDirectory, "log_*.txt");

                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    if (fileInfo.LastWriteTime < cutoffDate)
                    {
                        File.Delete(file);
                        System.Diagnostics.Debug.WriteLine($"[Logger] 已刪除舊日誌: {fileInfo.Name}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] 清理舊日誌失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 取得日誌目錄路徑
        /// </summary>
        public static string GetLogDirectory()
        {
            return LogDirectory;
        }

        /// <summary>
        /// 取得今日日誌文件路徑
        /// </summary>
        public static string GetTodayLogFile()
        {
            string fileName = $"log_{DateTime.Now:yyyyMMdd}.txt";
            return Path.Combine(LogDirectory, fileName);
        }

        /// <summary>
        /// 讀取今日日誌內容
        /// </summary>
        public static string ReadTodayLog()
        {
            try
            {
                string filePath = GetTodayLogFile();
                if (File.Exists(filePath))
                {
                    return File.ReadAllText(filePath, Encoding.UTF8);
                }
                return "今日尚無日誌記錄";
            }
            catch (Exception ex)
            {
                return $"讀取日誌失敗: {ex.Message}";
            }
        }

        /// <summary>
        /// 啟用/停用詳細日誌
        /// </summary>
        public static void SetDetailedLogging(bool enabled)
        {
            _enableDetailedLogging = enabled;
            Info($"詳細日誌已{(enabled ? "啟用" : "停用")}");
        }

        #endregion

        #region 性能監控

        private static ConcurrentDictionary<string, Stopwatch> _perfTimers = new ConcurrentDictionary<string, Stopwatch>();

        /// <summary>
        /// 開始性能計時
        /// </summary>
        public static void StartPerf(string operationName)
        {
            if (!_enableDetailedLogging)
                return;

            var sw = Stopwatch.StartNew();
            _perfTimers[operationName] = sw;
        }

        /// <summary>
        /// 停止性能計時並記錄
        /// </summary>
        public static void StopPerf(string operationName)
        {
            if (!_enableDetailedLogging)
                return;

            if (_perfTimers.TryRemove(operationName, out var sw))
            {
                sw.Stop();
                Debug($"⏱️ [{operationName}] 耗時: {sw.ElapsedMilliseconds}ms");
            }
        }

        /// <summary>
        /// 使用 using 語法的性能計時
        /// 用法: using (Logger.PerfScope("操作名稱")) { ... }
        /// </summary>
        public static IDisposable PerfScope(string operationName)
        {
            return new PerfTimer(operationName);
        }

        private class PerfTimer : IDisposable
        {
            private readonly string _name;
            private readonly Stopwatch _sw;

            public PerfTimer(string name)
            {
                _name = name;
                _sw = Stopwatch.StartNew();
            }

            public void Dispose()
            {
                _sw.Stop();
                if (_enableDetailedLogging)
                {
                    Debug($"⏱️ [{_name}] 耗時: {_sw.ElapsedMilliseconds}ms");
                }
            }
        }

        #endregion
    }

    #region 擴展方法

    /// <summary>
    /// 安全執行擴展方法 - 自動捕獲異常並記錄
    /// </summary>
    public static class SafeExecutionExtensions
    {
        /// <summary>
        /// 安全執行 Action（自動記錄異常）
        /// </summary>
        public static bool SafeExecute(this Action action, string operationName = "操作")
        {
            try
            {
                action();
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"{operationName}失敗", ex);
                return false;
            }
        }

        /// <summary>
        /// 安全執行 Func（自動記錄異常，返回預設值）
        /// </summary>
        public static T SafeExecute<T>(this Func<T> func, T defaultValue, string operationName = "操作")
        {
            try
            {
                return func();
            }
            catch (Exception ex)
            {
                Logger.Error($"{operationName}失敗", ex);
                return defaultValue;
            }
        }
    }

    #endregion
}
