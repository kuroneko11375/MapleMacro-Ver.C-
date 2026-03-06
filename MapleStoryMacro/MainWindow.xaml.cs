using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Threading;
using WpfBrushes = System.Windows.Media.Brushes;

namespace MapleStoryMacro
{
    public partial class MainWindow
    {
        // ═══ WPF helpers ═══
        private System.Windows.Threading.DispatcherTimer monitorTimer;

        private static System.Windows.Media.SolidColorBrush ToBrush(System.Drawing.Color c)
            => new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromArgb(c.A, c.R, c.G, c.B));

        private int LoopCountValue
        {
            get => int.TryParse(numPlayTimes.Text, out int v) ? Math.Clamp(v, 1, 9999) : 1;
            set => numPlayTimes.Text = Math.Clamp(value, 1, 9999).ToString();
        }

        private void NumPlayTimes_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !int.TryParse(e.Text, out _);
        }

        private List<MacroEvent> recordedEvents = new List<MacroEvent>();
        private bool isRecording = false;
        private volatile bool isPlaying = false;
        private double recordStartTime = 0;
        private IntPtr targetWindowHandle = IntPtr.Zero;
        private KeyboardHookDLL keyboardHook;

        // 高精度計時器（取代 Environment.TickCount，精度從 ~15ms 提升至 ~1μs）
        private readonly Stopwatch highResTimer = Stopwatch.StartNew();

        // 全局熱鍵設定
        private Keys playHotkey = Keys.F9;      // 預設播放熱鍵
        private Keys stopHotkey = Keys.F10;     // 預設停止熱鍵
        private Keys pauseHotkey = Keys.F11;    // 預設暫停熱鍵
        private Keys recordHotkey = Keys.F8;    // 預設錄製熱鍵
        private bool hotkeyEnabled = true;      // 熱鍵是否啟用
        private KeyboardHookDLL hotkeyHook;     // 全局熱鍵監聯器

        // 暫停控制
        private volatile bool isPaused = false;           // 暫停狀態
        private readonly object pauseLock = new object(); // 暫停鎖
        private volatile bool customKeysPaused = false;   // 自定義按鍵暫停狀態

        // 自定義按鍵槽位 (15個)
        private CustomKeySlot[] customKeySlots = new CustomKeySlot[15];

        // 執行統計
        private PlaybackStatistics statistics = new PlaybackStatistics();

        // 小地圖追蹤器
        private MinimapTracker? minimapTracker;

        // 位置修正器
        private PositionCorrector? positionCorrector;
        private PositionCorrectionSettings positionCorrectionSettings = new PositionCorrectionSettings();

        // 座標修正開關
        private bool positionCorrectionEnabled = false;
        private volatile bool _correctionSettingsChanged = false; // ★ 設定變更旗標
        private volatile bool isCorrecting = false; // ★ 位置修正中（阻塞腳本）
        /// <summary>
        /// ★ 技能硬直鎖：自定義按鍵施放期間為 true，位置修正器應暫停移動
        /// </summary>
        public static volatile bool IsAnimationLocked = false;

        // ★ 播放閘門：修正期間 Reset()，修正完成後 Set()，讓播放執行緒嚴格等待
        private readonly System.Threading.ManualResetEventSlim _playbackGate = new System.Threading.ManualResetEventSlim(true);

        // ★ Y 軸突變防護（雙層偵測）
        private volatile bool _hardInterruptY = false;       // 緊急掉落旗標
        private volatile int _expectedY = -1;                 // 腳本期望的 Y 座標
        private volatile bool _yAnomalyThreadRunning = false; // Y 軸偵測執行緒旗標
        private const int Y_ANOMALY_THRESHOLD = 40;          // Y 突變閾值（px）
        private const int Y_ANOMALY_INTERVAL_MS = 500;       // Y 軸偵測頻率（ms）
        private const int POST_CORRECTION_BUFFER_MS = 1200;  // 修正後緩衝延遲（ms）

        // ★ F7 邊界設定熱鍵 — 狀態機 (0=左,1=右,2=上,3=下)
        private Keys boundaryHotkey = Keys.F7;
        private int _boundarySetState = -1; // -1=未啟動, 0~3=設定中

        // 定時執行
        private List<ScheduleTask> scheduleTasks = new List<ScheduleTask>();
        private System.Windows.Threading.DispatcherTimer schedulerTimer;

        // Log System
        private readonly int MAX_LOG_LINES = 100;
        private readonly object logLock = new object();
        private string? lastLogMessage = null; // 用於合併重複日誌
        private int lastLogRepeatCount = 0; // 重複計數

        // pressedKeys 已移至 keySender.PressedKeys

        // 按鍵最短持續時間（確保遊戲能偵測到按鍵）
        private const double MIN_KEY_HOLD_SECONDS = 0.05; // ~50ms，確保遊戲至少輪詢 2-3 次（60fps ≈ 16ms/次）
        private readonly Dictionary<Keys, double> lastKeyDownTimestamp = new Dictionary<Keys, double>();

        // 當前腳本路徑
        private string? currentScriptPath = null;

        // Windows API P/Invoke
        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        // Background key sending APIs
        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        // 架構檢測 API
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool IsWow64Process(IntPtr hProcess, out bool Wow64Process);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        private const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        // Foreground key sending API
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        // SendInput API - 更現代的輸入模擬
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        // SendInput 結構定義
        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        private const uint INPUT_KEYBOARD = 1;

        // 子視窗列舉相關 API
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        // 方向鍵發送模式（ArrowKeyMode 已提取至 KeySender.cs）
        private ArrowKeyMode currentArrowKeyMode = ArrowKeyMode.SendToChild;

        // 鍵盤阻擋器（用於 Blocker 模式）
        private KeyboardBlocker? keyboardBlocker;

        // 按鍵發送引擎（已提取至 KeySender.cs）
        private KeySender keySender = new KeySender();

        // Windows message constants
        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP = 0x0101;
        private const uint WM_CHAR = 0x0102;
        private const uint WM_SYSKEYDOWN = 0x0104;
        private const uint WM_SYSKEYUP = 0x0105;
        private const uint MAPVK_VK_TO_VSC = 0;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

        // Mouse message constants
        private const uint WM_LBUTTONDOWN = 0x0201;
        private const uint WM_LBUTTONUP = 0x0202;
        private const uint MK_LBUTTON = 0x0001;

        // Mouse click P/Invoke (前景滑鼠點擊用)
        private const uint INPUT_MOUSE = 0;
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        // 子視窗搜尋 P/Invoke（用於找 MapleStoryClass）
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string? lpszClass, string? lpszWindow);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        private const uint WM_MOUSEMOVE = 0x0200;

        // 按鍵顯示名稱映射表
        private static readonly Dictionary<Keys, string> KeyDisplayNames = new Dictionary<Keys, string>
        {
            // 符號鍵 - Shift 時會變成 <> 等符號
            { Keys.OemPeriod, ". (>)" },
            { Keys.Oemcomma, ", (<)" },
            { Keys.OemQuestion, "/ (?)" },
            { Keys.OemSemicolon, "; (:)" },
            { Keys.OemQuotes, "' (\")" },
            { Keys.OemOpenBrackets, "[ ({)" },
            { Keys.OemCloseBrackets, "] (})" },
            { Keys.OemBackslash, "\\ (|)" },
            { Keys.OemMinus, "- (_)" },
            { Keys.Oemplus, "= (+)" },
            { Keys.Oemtilde, "` (~)" },
            { Keys.OemPipe, "\\ (|)" },
            // 常用功能鍵
            { Keys.Space, "空白鍵" },
            { Keys.Enter, "Enter" },
            { Keys.Escape, "Esc" },
            { Keys.Tab, "Tab" },
            { Keys.Back, "Backspace" },
            { Keys.Delete, "Delete" },
            { Keys.Insert, "Insert" },
            { Keys.Home, "Home" },
            { Keys.End, "End" },
            { Keys.PageUp, "Page Up" },
            { Keys.PageDown, "Page Down" },
            // 方向鍵
            { Keys.Left, "← 左" },
            { Keys.Right, "→ 右" },
            { Keys.Up, "↑ 上" },
            { Keys.Down, "↓ 下" },
            // 修飾鍵
            { Keys.LShiftKey, "左 Shift" },
            { Keys.RShiftKey, "右 Shift" },
            { Keys.ShiftKey, "Shift" },
            { Keys.LControlKey, "左 Ctrl" },
            { Keys.RControlKey, "右 Ctrl" },
            { Keys.ControlKey, "Ctrl" },
            { Keys.LMenu, "左 Alt" },
            { Keys.RMenu, "右 Alt" },
            { Keys.Menu, "Alt" },
            { Keys.Alt, "Alt" },
            { Keys.LWin, "左 Win" },
            { Keys.RWin, "右 Win" },
            // 數字鍵盤
            { Keys.NumPad0, "Num 0" },
            { Keys.NumPad1, "Num 1" },
            { Keys.NumPad2, "Num 2" },
            { Keys.NumPad3, "Num 3" },
            { Keys.NumPad4, "Num 4" },
            { Keys.NumPad5, "Num 5" },
            { Keys.NumPad6, "Num 6" },
            { Keys.NumPad7, "Num 7" },
            { Keys.NumPad8, "Num 8" },
            { Keys.NumPad9, "Num 9" },
            { Keys.Multiply, "Num *" },
            { Keys.Add, "Num +" },
            { Keys.Subtract, "Num -" },
            { Keys.Decimal, "Num ." },
            { Keys.Divide, "Num /" },
            { Keys.NumLock, "Num Lock" },
            // 其他
            { Keys.CapsLock, "Caps Lock" },
            { Keys.PrintScreen, "Print Screen" },
            { Keys.Scroll, "Scroll Lock" },
            { Keys.Pause, "Pause" },
        };

        /// <summary>
        /// 取得按鍵的顯示名稱（更直觀的中文名稱）
        /// </summary>
        private static string GetKeyDisplayName(Keys key)
        {
            if (KeyDisplayNames.TryGetValue(key, out var displayName))
                return displayName ?? key.ToString();
            return key.ToString();
        }

        /// <summary>
        /// 取得含修飾鍵的按鍵組合顯示名稱（例如 "Ctrl+Alt+Z"）
        /// </summary>
        private static string GetModifierKeyDisplayName(Keys key, Keys modifiers)
        {
            var parts = new List<string>();
            if ((modifiers & Keys.Control) != 0) parts.Add("Ctrl");
            if ((modifiers & Keys.Alt) != 0) parts.Add("Alt");
            if ((modifiers & Keys.Shift) != 0) parts.Add("Shift");
            parts.Add(GetKeyDisplayName(key));
            return string.Join("+", parts);
        }

        public MainWindow()
        {
            InitializeComponent();

            // 設定視窗圖示（從 XAML 移至 code-behind，避免 BAML 解析失敗）
            try
            {
                Icon = new System.Windows.Media.Imaging.BitmapImage(
                    new Uri("pack://application:,,,/MAPLE.ico"));
            }
            catch { }

            // 初始化日誌系統
            Logger.Initialize(enableDetailedLogging: false, logRetentionDays: 7);
            Logger.Info("=== MapleStory Macro 啟動 ===");
            Logger.Info($"版本: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}");

            // 初始化按鍵發送引擎
            keySender.Log = (msg) => AddLog(msg);


            // 初始化自定義按鍵槽位
            for (int i = 0; i < 15; i++)
            {
                customKeySlots[i] = new CustomKeySlot { SlotNumber = i + 1 };
            }

            // 初始化定時器
            schedulerTimer = new System.Windows.Threading.DispatcherTimer();
            schedulerTimer.Interval = TimeSpan.FromMilliseconds(1000); // 每秒檢查一次
            schedulerTimer.Tick += SchedulerTimer_Tick;

            // Bind button events not declared in XAML (lambda handlers)
            btnHotkeySettings.Click += (s, e) => OpenHotkeySettings();
            btnCustomKeys.Click += (s, e) => OpenCustomKeySettings();
            btnScheduler.Click += (s, e) => OpenSchedulerSettings();
            btnStatistics.Click += (s, e) => ShowStatistics();
            btnMemoryScanner.Click += (s, e) => OpenMemoryScannerChoice();

            // ★ 方向鍵模式下拉框事件（控件已在 Designer 中定義）
            cmbArrowMode.SelectedIndex = (int)currentArrowKeyMode;
            cmbArrowMode.SelectionChanged += (s, e) =>
            {
                currentArrowKeyMode = (ArrowKeyMode)cmbArrowMode.SelectedIndex;
                keySender.CurrentArrowKeyMode = currentArrowKeyMode;
                SaveAppSettings();
                AddLog($"方向鍵模式切換: {currentArrowKeyMode}");
            };

            // Initialize keyboard hook for recording
            keyboardHook = new KeyboardHookDLL(KeyboardHookDLL.KeyboardHookMode.LowLevel);
            keyboardHook.OnKeyEvent += KeyboardHook_OnKeyEvent;

            // Initialize global hotkey hook (延遲安裝到 MainWindow_ContentRendered)
            hotkeyHook = new KeyboardHookDLL(KeyboardHookDLL.KeyboardHookMode.LowLevel);
            hotkeyHook.OnKeyEvent += HotkeyHook_OnKeyEvent;

            // 延遲初始化 - 在視窗顯示後再載入設定和啟動定時器
            this.ContentRendered += MainWindow_ContentRendered;

            // 先設定初始狀態
            lblWindowStatus.Text = "視窗: 未鎖定";
            lblWindowStatus.Foreground = ToBrush(Color.Gray);
            UpdateUI();
            UpdatePauseButtonState();
        }

        private void MainWindow_ContentRendered(object? sender, EventArgs e)
        {
            // Initialize monitor timer
            monitorTimer = new System.Windows.Threading.DispatcherTimer();
            monitorTimer.Interval = TimeSpan.FromMilliseconds(500);
            monitorTimer.Tick += MonitorTimer_Tick;
            monitorTimer.Start();

            // 載入儲存的設定
            LoadSettings();

            // 安裝全局熱鍵鉤子
            hotkeyHook.Install();

            AddLog("應用程式已啟動");
            AddLog($"全局熱鍵：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}, 暫停={GetKeyDisplayName(pauseHotkey)}, 錄製={GetKeyDisplayName(recordHotkey)}");
        }

        /// <summary>
        /// 複製選中的日誌到剪貼簿
        /// </summary>
        private void MenuCopyLog_Click(object? sender, EventArgs e)
        {
            if (lstLog.SelectedItems.Count > 0)
            {
                var selectedLogs = lstLog.SelectedItems.Cast<string>().ToList();
                string text = string.Join(Environment.NewLine, selectedLogs);
                Clipboard.SetText(text);
                AddLog($"✅ 已複製 {selectedLogs.Count} 行日誌");
            }
        }

        /// <summary>
        /// 清除所有日誌
        /// </summary>
        private void MenuClearLog_Click(object? sender, EventArgs e)
        {
            lock (logLock)
            {
                lstLog.Items.Clear();
                lastLogMessage = null;
                lastLogRepeatCount = 0;
            }
        }

        /// <summary>
        /// 定時執行檢查
        /// </summary>
        private void SchedulerTimer_Tick(object? sender, EventArgs e)
        {
            // 檢查排程任務
            foreach (var task in scheduleTasks.ToList())
            {
                if (!task.Enabled) continue;

                // 檢查結束時間（自動停止）
                if (task.HasStarted && task.EndTime.HasValue && DateTime.Now >= task.EndTime.Value)
                {
                    task.Enabled = false;

                    // ★ 優先中斷目前的腳本任務
                    if (isPlaying)
                    {
                        AddLog($"排程結束：已到達結束時間 {task.EndTime.Value:HH:mm:ss}，強制停止腳本");
                        BtnStopPlayback_Click(this, EventArgs.Empty);
                    }

                    // 執行回程序列（在背景線程上執行，含 2000ms 冷卻延遲）
                    if (task.ReturnToTownEnabled)
                    {
                        var returnTask = task;
                        Thread returnThread = new Thread(() => ExecuteReturnToTown(returnTask))
                        {
                            IsBackground = true
                        };
                        returnThread.Start();
                    }

                    continue;
                }

                // 檢查開始時間
                if (!task.HasStarted && DateTime.Now >= task.StartTime)
                {
                    task.HasStarted = true;

                    // 載入指定腳本
                    if (!string.IsNullOrEmpty(task.ScriptPath) && File.Exists(task.ScriptPath))
                    {
                        LoadScriptFromFile(task.ScriptPath);
                        LoopCountValue = task.LoopCount;
                        AddLog($"排程觸發：載入腳本 {Path.GetFileName(task.ScriptPath)}");
                    }

                    if (!isPlaying && recordedEvents.Count > 0)
                    {
                        AddLog($"排程觸發：開始播放");
                        BtnStartPlayback_Click(this, EventArgs.Empty);
                    }

                    // 如果沒有結束時間，標記為完成
                    if (!task.EndTime.HasValue)
                    {
                        task.Enabled = false;
                    }
                }
            }

            // 移除已完成的任務
            scheduleTasks.RemoveAll(t => !t.Enabled);
            if (scheduleTasks.Count == 0)
            {
                schedulerTimer.Stop();
            }
        }

        private void AddLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");

            // 在 UI 執行緒上更新 ListBox
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => AddLogInternal(timestamp, message)));
            }
            else
            {
                AddLogInternal(timestamp, message);
            }

            System.Diagnostics.Debug.WriteLine($"[{timestamp}] {message}");
        }

        private void AddLogInternal(string timestamp, string message)
        {
            lock (logLock)
            {
                // 檢查是否與上一條日誌相同（忽略時間戳）
                if (lastLogMessage == message)
                {
                    lastLogRepeatCount++;
                    // 只在第一次重複和每 50 次時更新顯示
                    if (lastLogRepeatCount == 1 || lastLogRepeatCount % 50 == 0)
                    {
                        if (lstLog.Items.Count > 0)
                        {
                            lstLog.Items[lstLog.Items.Count - 1] = $"[{timestamp}] {message} (×{lastLogRepeatCount + 1})";
                            if (lstLog.Items.Count > 0) lstLog.ScrollIntoView(lstLog.Items[lstLog.Items.Count - 1]);
                        }
                    }
                }
                else
                {
                    // 新的日誌訊息
                    lastLogMessage = message;
                    lastLogRepeatCount = 0;

                    string logEntry = $"[{timestamp}] {message}";
                    lstLog.Items.Add(logEntry);

                    // 限制日誌數量
                    while (lstLog.Items.Count > MAX_LOG_LINES)
                    {
                        lstLog.Items.RemoveAt(0);
                    }

                    // 自動滾動到最新
                    if (lstLog.Items.Count > 0) lstLog.ScrollIntoView(lstLog.Items[lstLog.Items.Count - 1]);
                }
            }
        }


        private void MonitorTimer_Tick(object? sender, EventArgs e)
        {
            // 更新狀態列標籤
            string bgMode = (targetWindowHandle != IntPtr.Zero) ? "背景" : "前景";
            string recStatus = isRecording ? "🔴" : "⚪";
            string playStatus = isPlaying ? "▶️" : "⏹️";

            // 腳本預估時長
            string durationInfo = "";
            if (recordedEvents.Count > 0)
            {
                double maxTs = recordedEvents.Max(ev => ev.Timestamp);
                if (maxTs >= 60)
                    durationInfo = $" | 時長: {maxTs / 60:F1}分";
                else
                    durationInfo = $" | 時長: {maxTs:F1}秒";
            }

            // 排程狀態
            string schedInfo = "";
            int activeSchedules = scheduleTasks.Count(t => t.Enabled);
            if (activeSchedules > 0)
                schedInfo = $" | 排程: {activeSchedules}";

            lblStatus.Text = $"{recStatus} 錄製 | {playStatus} 播放 | 模式: {bgMode} | 事件: {recordedEvents.Count}{durationInfo}{schedInfo}";

            // 即時顯示 Hook 運作中
            if (isPlaying && keyboardBlocker != null)
            {
                lblStatus.Text += $" | Hook 運作中";
            }

            // 即時顯示座標位置（只在修正開關啟用時顯示）
            if (positionCorrectionSettings.Enabled && minimapTracker != null && minimapTracker.IsCalibrated)
            {
                var (x, y, success) = minimapTracker.ReadPosition();
                if (success)
                {
                    bool inBounds = minimapTracker.IsInBounds();
                    string boundStatus = inBounds ? "✓" : "⚠";
                    lblStatus.Text += $" | 位置: ({x}, {y}) {boundStatus}";
                }
            }
        }

        private void MainWindow_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // 當鍵盤鉤子啟用時，不重複錄製（避免雙重事件）
            if (isRecording && keyboardHook != null && !keyboardHook.IsInstalled)
            {
                var key = e.Key == System.Windows.Input.Key.System ? e.SystemKey : e.Key;
                Keys keyCode = (Keys)System.Windows.Input.KeyInterop.VirtualKeyFromKey(key);

                if (recordStartTime == 0)
                    recordStartTime = GetCurrentTime();

                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = keyCode,
                    EventType = "down",
                    Timestamp = GetCurrentTime() - recordStartTime
                });

                lblRecordingStatus.Text = $"錄製中: {recordedEvents.Count} 個事件 | 最後: {keyCode}";
                AddLog($"按鍵按下: {keyCode}");
            }
        }

        private void MainWindow_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // 當鍵盤鉤子啟用時，不重複錄製（避免雙重事件）
            if (isRecording && keyboardHook != null && !keyboardHook.IsInstalled)
            {
                var key = e.Key == System.Windows.Input.Key.System ? e.SystemKey : e.Key;
                Keys keyCode = (Keys)System.Windows.Input.KeyInterop.VirtualKeyFromKey(key);

                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = keyCode,
                    EventType = "up",
                    Timestamp = GetCurrentTime() - recordStartTime
                });

                lblRecordingStatus.Text = $"錄製中: {recordedEvents.Count} 個事件 | 最後: {keyCode}";
                AddLog($"按鍵放開: {keyCode}");
            }
        }

        private void KeyboardHook_OnKeyEvent(Keys keyCode, bool isKeyDown)
        {
            if (isRecording)
            {
                // ★ 跳過全局熱鍵，不錄製進腳本
                if (keyCode == recordHotkey || keyCode == playHotkey || keyCode == stopHotkey ||
                    keyCode == pauseHotkey || keyCode == boundaryHotkey)
                    return;

                if (recordStartTime == 0)
                    recordStartTime = GetCurrentTime();

                string eventType = isKeyDown ? "down" : "up";

                // ★ 錄製時同時記錄小地圖座標
                int rx = -1, ry = -1;
                if (minimapTracker != null && minimapTracker.IsCalibrated)
                {
                    var (px, py, pOk) = minimapTracker.ReadPosition();
                    if (pOk) { rx = px; ry = py; }
                }

                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = keyCode,
                    EventType = eventType,
                    Timestamp = GetCurrentTime() - recordStartTime,
                    RecordedX = rx,
                    RecordedY = ry
                });

                lblRecordingStatus.Text = $"錄製中: {recordedEvents.Count} 個事件 | 最後: {keyCode}";

                if (isKeyDown)
                    AddLog($"鉤子按下: {keyCode}");
                else
                    AddLog($"鉤子放開: {keyCode}");
            }
        }

        private void BtnRefreshWindow_Click(object? sender, EventArgs e)
        {
            AddLog("正在搜尋目標視窗...");
            FindGameWindow();
            UpdateWindowStatus();
        }

        private void BtnLockWindow_Click(object? sender, EventArgs e)
        {
            AddLog("正在開啟視窗選擇器...");
            SelectWindow();
        }

        private void SelectWindow()
        {
            Form windowSelector = new Form
            {
                Text = "選擇視窗 (雙擊快速選擇)",
                Width = 450,
                Height = 350,
                StartPosition = FormStartPosition.CenterScreen
            };

            // 提示標籤
            Label lblHint = new Label
            {
                Text = "★ 雙擊視窗名稱即可快速選擇",
                Dock = DockStyle.Top,
                Height = 25,
                ForeColor = Color.Blue,
                Font = new Font("microsoft yahei ui", 9F, System.Drawing.FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.LightYellow
            };

            ListBox listBox = new ListBox
            {
                Dock = DockStyle.Fill
            };

            Process[] processes = Process.GetProcesses();
            foreach (Process p in processes)
            {
                try
                {
                    if (!string.IsNullOrEmpty(p.MainWindowTitle) && IsWindow(p.MainWindowHandle))
                    {
                        bool is32 = IsProcess32Bit((uint)p.Id);
                        string arch = is32 ? "32-bit" : "64-bit";
                        listBox.Items.Add(new ProcessItem
                        {
                            Handle = p.MainWindowHandle,
                            Title = $"{p.MainWindowTitle} [{arch}] (PID: {p.Id})",
                            Is32Bit = is32
                        });
                    }
                }
                catch { }
            }

            // ★ 雙擊快速選擇
            listBox.DoubleClick += (s, args) =>
            {
                if (listBox.SelectedIndex >= 0)
                {
                    ProcessItem? selected = listBox.SelectedItem as ProcessItem;
                    if (selected != null)
                    {
                        targetWindowHandle = selected.Handle;
                        keySender.TargetWindowHandle = targetWindowHandle;
                        UpdateWindowStatus();
                        AddLog($"已鎖定視窗: {selected.Title}");
                        windowSelector.Close();
                    }
                }
            };

            Panel btnPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40
            };

            Button okBtn = new Button
            {
                Text = "確定",
                Width = 80,
                Height = 30,
                Left = 200,
                Top = 5,
                DialogResult = System.Windows.Forms.DialogResult.OK
            };

            Button cancelBtn = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 30,
                Left = 290,
                Top = 5,
                DialogResult = System.Windows.Forms.DialogResult.Cancel
            };

            btnPanel.Controls.Add(okBtn);
            btnPanel.Controls.Add(cancelBtn);

            // 控制項加入順序：先 Fill，再 Top/Bottom
            windowSelector.Controls.Add(listBox);
            windowSelector.Controls.Add(lblHint);
            windowSelector.Controls.Add(btnPanel);

            listBox.DisplayMember = "Title";

            if (windowSelector.ShowDialog() == System.Windows.Forms.DialogResult.OK && listBox.SelectedIndex >= 0)
            {
                ProcessItem? selected = listBox.SelectedItem as ProcessItem;
                if (selected != null)
                {
                    targetWindowHandle = selected.Handle;
                    keySender.TargetWindowHandle = targetWindowHandle;
                    UpdateWindowStatus();
                    AddLog($"已鎖定視窗: {selected.Title}");
                }
            }
        }

        private class ProcessItem
        {
            public IntPtr Handle { get; set; }
            public string Title { get; set; } = string.Empty;
            public bool Is32Bit { get; set; }
        }

        /// <summary>
        /// 檢測目標進程是否為 32 位元
        /// </summary>
        private static bool IsProcess32Bit(uint processId)
        {
            // 如果在 32 位元 OS 上，所有進程都是 32 位元
            if (!Environment.Is64BitOperatingSystem)
                return true;

            IntPtr hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, processId);
            if (hProcess == IntPtr.Zero)
                return false; // 無法開啟，假設未知

            try
            {
                if (IsWow64Process(hProcess, out bool isWow64))
                {
                    // WoW64 = 32 位元進程在 64 位元 OS 上執行
                    return isWow64;
                }
                return false;
            }
            finally
            {
                CloseHandle(hProcess);
            }
        }

        /// <summary>
        /// 透過視窗控制代碼檢測進程是否為 32 位元
        /// </summary>
        private static bool IsWindowProcess32Bit(IntPtr hWnd)
        {
            GetWindowThreadProcessId(hWnd, out uint processId);
            if (processId == 0) return false;
            return IsProcess32Bit(processId);
        }

        private void FindGameWindow()
        {
            string windowKeyword = txtWindowTitle?.Text?.Trim() ?? "MapleStory";
            if (string.IsNullOrEmpty(windowKeyword))
                windowKeyword = "MapleStory";

            Process[] processes = Process.GetProcesses();
            foreach (Process p in processes)
            {
                try
                {
                    if (p.MainWindowTitle.Contains(windowKeyword))
                    {
                        targetWindowHandle = p.MainWindowHandle;
                        keySender.TargetWindowHandle = targetWindowHandle;
                        break;
                    }
                }
                catch { }
            }
        }

        private void UpdateWindowStatus()
        {
            if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
            {
                bool is32 = IsWindowProcess32Bit(targetWindowHandle);
                string arch = is32 ? "32-bit" : "64-bit";
                lblWindowStatus.Text = $"視窗: 已鎖定 [{arch}] - 背景模式";
                lblWindowStatus.Foreground = ToBrush(Color.Green);
                AddLog($"已鎖定視窗: {targetWindowHandle} [{arch}]");
            }
            else
            {
                lblWindowStatus.Text = "視窗: 未找到 - 前景模式";
                lblWindowStatus.Foreground = ToBrush(Color.Red);
                targetWindowHandle = IntPtr.Zero;
                keySender.TargetWindowHandle = IntPtr.Zero;
                AddLog("未找到目標視窗");
            }
        }

        private void ReleasePressedKeys()
        {
            keySender.ReleasePressedKeys();
        }

        private void BtnStartRecording_Click(object? sender, EventArgs e)
        {
            if (isRecording) return;

            recordedEvents.Clear();
            recordStartTime = 0;
            isRecording = true;

            lblRecordingStatus.Text = "錄製中...";
            lblRecordingStatus.Foreground = ToBrush(Color.Red);
            lblStatus.Text = "按下按鍵以錄製";

            AddLog("錄製已開始");

            if (keyboardHook.Install())
            {
                AddLog("鍵盤鉤子已啟动");
            }
            else
            {
                AddLog("鍵盤鉤子啟動失敗");
                MessageBox.Show("鍵盤鉤子啟动失敗");
                isRecording = false;
                return;
            }

            UpdateUI();
        }

        private void BtnStopRecording_Click(object? sender, EventArgs e)
        {
            if (!isRecording) return;

            isRecording = false;
            lblRecordingStatus.Text = $"已停止 | 事件數: {recordedEvents.Count}";
            lblRecordingStatus.Foreground = ToBrush(Color.Green);

            keyboardHook.Uninstall();

            AddLog($"錄製已停止 - 共 {recordedEvents.Count} 個事件");
            UpdateUI();
        }

        private void BtnSaveScript_Click(object? sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                AddLog("⚠️ 沒有事件可保存");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Maple 腳本|*.mscript|舊版 JSON 腳本|*.json",
                DefaultExt = ".mscript",
                Title = "保存腳本"
            };

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    // 建立腳本資料
                    var scriptData = new ScriptData
                    {
                        Name = Path.GetFileNameWithoutExtension(sfd.FileName),
                        LoopCount = LoopCountValue,
                        ModifiedAt = DateTime.Now
                    };

                    // 轉換事件
                    bool hasAnyPath = false;
                    foreach (var evt in recordedEvents)
                    {
                        var se = new ScriptEvent
                        {
                            KeyCode = (int)evt.KeyCode,
                            EventType = evt.EventType,
                            Timestamp = evt.Timestamp,
                            RecordedX = evt.RecordedX,
                            RecordedY = evt.RecordedY,
                            SkillAnimationDelay = evt.SkillAnimationDelay // ★
                        };
                        scriptData.Events.Add(se);
                        if (evt.RecordedX >= 0) hasAnyPath = true;
                    }
                    scriptData.HasPathData = hasAnyPath;

                    // 複製自定義按鍵設定
                    for (int i = 0; i < 15; i++)
                    {
                        scriptData.CustomKeySlots[i] = new CustomKeySlotData
                        {
                            SlotNumber = customKeySlots[i].SlotNumber,
                            KeyCode = (int)customKeySlots[i].KeyCode,
                            Modifiers = (int)customKeySlots[i].Modifiers,
                            IntervalSeconds = customKeySlots[i].IntervalSeconds,
                            Enabled = customKeySlots[i].Enabled,
                            StartAtSecond = customKeySlots[i].StartAtSecond,
                            PreDelaySeconds = customKeySlots[i].PreDelaySeconds,
                            PauseScriptSeconds = customKeySlots[i].PauseScriptSeconds,
                            PauseScriptEnabled = customKeySlots[i].PauseScriptEnabled
                        };
                    }

                    string json = JsonSerializer.Serialize(scriptData, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(sfd.FileName, json);

                    // 記住最後使用的腳本路徑
                    currentScriptPath = sfd.FileName;

                    int enabledCustomKeys = customKeySlots.Count(s => s.Enabled && s.KeyCode != Keys.None);
                    string pathInfo = hasAnyPath ? ", 含座標路徑" : "";
                    AddLog($"✅ 已保存: {Path.GetFileName(sfd.FileName)} ({recordedEvents.Count} 事件{pathInfo}, {enabledCustomKeys} 自定義按鍵)");
                }
                catch (Exception ex)
                {
                    AddLog($"❌ 保存失敗: {ex.Message}");
                }
            }
        }

        private void BtnLoadScript_Click(object? sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Maple 腳本|*.mscript|舊版 JSON 腳本|*.json|所有檔案|*.*",
                Title = "載入腳本"
            };

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                LoadScriptFromFile(ofd.FileName);
            }
        }

        /// <summary>
        /// 從檔案載入腳本
        /// </summary>
        private void LoadScriptFromFile(string filePath)
        {
            try
            {
                string json = File.ReadAllText(filePath);
                string ext = Path.GetExtension(filePath).ToLowerInvariant();

                // 嘗試判斷檔案格式
                if (ext == ".mscript" || json.Contains("\"Version\"") || json.Contains("\"CustomKeySlots\""))
                {
                    // 新格式：ScriptData
                    var scriptData = JsonSerializer.Deserialize<ScriptData>(json);
                    if (scriptData == null)
                        throw new InvalidOperationException("腳本格式無效");

                    // 載入事件
                    recordedEvents.Clear();
                    foreach (var evt in scriptData.Events)
                    {
                        recordedEvents.Add(new MacroEvent
                        {
                            KeyCode = (Keys)evt.KeyCode,
                            EventType = evt.EventType,
                            Timestamp = evt.Timestamp,
                            RecordedX = evt.RecordedX,
                            RecordedY = evt.RecordedY,
                            SkillAnimationDelay = evt.SkillAnimationDelay // ★
                        });
                    }

                    // 載入循環次數
                    LoopCountValue = scriptData.LoopCount;

                    // 載入自定義按鍵設定
                    if (scriptData.CustomKeySlots != null)
                    {
                        for (int i = 0; i < Math.Min(15, scriptData.CustomKeySlots.Length); i++)
                        {
                            var data = scriptData.CustomKeySlots[i];
                            if (data != null)
                            {
                                customKeySlots[i].SlotNumber = data.SlotNumber;
                                customKeySlots[i].KeyCode = (Keys)data.KeyCode;
                                customKeySlots[i].Modifiers = (Keys)data.Modifiers;
                                customKeySlots[i].IntervalSeconds = data.IntervalSeconds;
                                customKeySlots[i].Enabled = data.Enabled;
                                customKeySlots[i].StartAtSecond = data.StartAtSecond;
                                customKeySlots[i].PreDelaySeconds = data.PreDelaySeconds;
                                customKeySlots[i].PauseScriptSeconds = data.PauseScriptSeconds;
                                customKeySlots[i].PauseScriptEnabled = data.PauseScriptEnabled;
                            }
                        }
                    }

                    int enabledCustomKeys = customKeySlots.Count(s => s.Enabled && s.KeyCode != Keys.None);
                    bool hasPath = recordedEvents.Any(e => e.RecordedX >= 0);
                    string pathStr = hasPath ? ", 含座標路徑" : "";
                    AddLog($"✅ 已載入: {Path.GetFileName(filePath)} ({recordedEvents.Count} 事件{pathStr}, {enabledCustomKeys} 自定義按鍵)");
                }
                else
                {
                    // 舊格式：純事件列表
                    var events = JsonSerializer.Deserialize<List<MacroEvent>>(json);
                    recordedEvents = events ?? new List<MacroEvent>();
                    AddLog($"✅ 已載入舊格式: {Path.GetFileName(filePath)} ({recordedEvents.Count} 事件)");
                }

                // 記住最後使用的腳本路徑
                currentScriptPath = filePath;

                // 計算腳本時長
                string durationStr = "";
                if (recordedEvents.Count > 0)
                {
                    double maxTs = recordedEvents.Max(ev => ev.Timestamp);
                    if (maxTs >= 60)
                        durationStr = $", 時長≈{maxTs / 60:F1}分";
                    else
                        durationStr = $", 時長≈{maxTs:F1}秒";
                }

                lblRecordingStatus.Text = $"已載入 | 事件數: {recordedEvents.Count}{durationStr}";

                // 更新 UI 狀態，啟用開始播放按鍵
                UpdateUI();
            }
            catch (Exception ex)
            {
                AddLog($"❌ 載入失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 格式化時間間隔為可讀字串
        /// </summary>
        private static string FormatDelta(double deltaSeconds)
        {
            if (deltaSeconds >= 1.0)
                return $"{deltaSeconds:F2} 秒";
            return $"{deltaSeconds * 1000:F0} ms";
        }

        private void BtnEditEvents_Click(object? sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                AddLog("⚠️ 沒有事件可編輯");
                return;
            }

            // 開啟摺疊式編輯器
            OpenFoldedEventEditor();
        }

        /// <summary>
        /// 摺疊事件結構 - 表示一個按鍵從按下到放開的完整動作
        /// </summary>
        private class FoldedKeyAction
        {
            public Keys KeyCode { get; set; }
            public string KeyName { get; set; } = "";
            public double PressTime { get; set; }      // 按下的時間點
            public double ReleaseTime { get; set; }    // 放開的時間點
            public double Duration => ReleaseTime - PressTime;  // 持續時間
            public int RepeatCount { get; set; } = 1;  // 重複按下次數 (連打)
            public bool IsReleased { get; set; } = true;  // 是否已放開

            public string GetDisplayText()
            {
                string durationText = Duration >= 1.0 ? $"{Duration:F2}秒" : $"{Duration * 1000:F0}ms";
                if (RepeatCount > 1)
                {
                    return $"{KeyName}  x{RepeatCount}  共{durationText}  [按下]";
                }
                return $"{KeyName}  {durationText}  [按下]";
            }

            public string GetReleaseText()
            {
                return $"{KeyName}  [放開]  @{ReleaseTime:F3}s";
            }
        }

        /// <summary>
        /// 將原始事件轉換為摺疊格式
        /// </summary>
        private List<FoldedKeyAction> ConvertToFoldedActions()
        {
            var actions = new List<FoldedKeyAction>();
            var sortedEvents = recordedEvents.OrderBy(e => e.Timestamp).ToList();
            
            // 追蹤每個按鍵的狀態
            var keyStates = new Dictionary<Keys, FoldedKeyAction>();

            foreach (var evt in sortedEvents)
            {
                if (evt.EventType == "down")
                {
                    if (keyStates.TryGetValue(evt.KeyCode, out var existing))
                    {
                        // 已經在按住狀態，增加重複計數
                        existing.RepeatCount++;
                    }
                    else
                    {
                        // 新的按下
                        keyStates[evt.KeyCode] = new FoldedKeyAction
                        {
                            KeyCode = evt.KeyCode,
                            KeyName = GetKeyDisplayName(evt.KeyCode),
                            PressTime = evt.Timestamp,
                            IsReleased = false,
                            RepeatCount = 1
                        };
                    }
                }
                else if (evt.EventType == "up")
                {
                    if (keyStates.TryGetValue(evt.KeyCode, out var action))
                    {
                        action.ReleaseTime = evt.Timestamp;
                        action.IsReleased = true;
                        actions.Add(action);
                        keyStates.Remove(evt.KeyCode);
                    }
                    else
                    {
                        // 沒有對應的按下事件，創建一個瞬間按下的動作
                        actions.Add(new FoldedKeyAction
                        {
                            KeyCode = evt.KeyCode,
                            KeyName = GetKeyDisplayName(evt.KeyCode),
                            PressTime = evt.Timestamp - 0.05,
                            ReleaseTime = evt.Timestamp,
                            IsReleased = true
                        });
                    }
                }
            }

            // 處理還沒放開的按鍵
            double lastTime = sortedEvents.Count > 0 ? sortedEvents.Last().Timestamp : 0;
            foreach (var kvp in keyStates)
            {
                kvp.Value.ReleaseTime = lastTime + 0.1;
                kvp.Value.IsReleased = false;
                actions.Add(kvp.Value);
            }

            return actions.OrderBy(a => a.PressTime).ToList();
        }

        /// <summary>
        /// 從摺疊動作重建原始事件
        /// </summary>
        private void RebuildEventsFromActions(List<FoldedKeyAction> actions)
        {
            recordedEvents.Clear();
            
            foreach (var action in actions.OrderBy(a => a.PressTime))
            {
                // 添加按下事件
                double interval = action.Duration / Math.Max(1, action.RepeatCount);
                for (int i = 0; i < action.RepeatCount; i++)
                {
                    recordedEvents.Add(new MacroEvent
                    {
                        KeyCode = action.KeyCode,
                        EventType = "down",
                        Timestamp = action.PressTime + (i * interval * 0.01)
                    });
                }

                // 添加放開事件
                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = action.KeyCode,
                    EventType = "up",
                    Timestamp = action.ReleaseTime
                });
            }
        }

        /// <summary>
        /// 開啟摺疊式事件編輯器
        /// </summary>
        private void OpenFoldedEventEditor()
        {
            OpenFoldedEventEditor_Wpf();
        }

        /// <summary>
        /// 開啟文本編輯器（原有功能）
        /// </summary>
        private void OpenTextEditor()
        {
            AddLog("正在開啟文本編輯器...");

            Form editorForm = new Form
            {
                Text = $"腳本編輯器 ({recordedEvents.Count} 個事件)",
                Width = 800,
                Height = 700,
                StartPosition = FormStartPosition.CenterScreen
            };

            // 說明標籤
            Label hintLabel = new Label
            {
                Text = "格式: 按鍵名稱 | 類型(down/up) | 時間戳(秒)    每行一個事件，可直接編輯、複製、貼上",
                Top = 10,
                Left = 10,
                Width = 760,
                ForeColor = Color.Blue,
                Font = new Font("Microsoft JhengHei", 9, System.Drawing.FontStyle.Bold)
            };

            // 文本編輯區
            RichTextBox txtEditor = new RichTextBox
            {
                Top = 35,
                Left = 10,
                Width = 760,
                Height = 520,
                Font = new Font("Consolas", 10),
                AcceptsTab = false,
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both
            };

            // 將事件轉換為文本格式
            var sb = new StringBuilder();
            sb.AppendLine("# ===== 腳本事件列表 =====");
            sb.AppendLine("# 格式: 按鍵名稱 | 類型 | 時間戳(秒)");
            sb.AppendLine("# 支援的類型: down (按下), up (放開)");
            sb.AppendLine("# 以 # 開頭的行為註解，會被忽略");
            sb.AppendLine("# ========================");
            sb.AppendLine();

            foreach (var evt in recordedEvents)
            {
                string keyName = evt.KeyCode.ToString();
                sb.AppendLine($"{keyName} | {evt.EventType} | {evt.Timestamp:F3}");
            }

            txtEditor.Text = sb.ToString();

            // 標記是否有未儲存的變更
            bool hasUnsavedChanges = false;
            txtEditor.TextChanged += (s, args) => { hasUnsavedChanges = true; };

            // 按鈕面板
            Panel btnPanel = new Panel
            {
                Top = 565,
                Left = 10,
                Width = 760,
                Height = 90,
                BorderStyle = BorderStyle.FixedSingle
            };

            Button parseBtn = new Button { Text = "📋 解析預覽", Width = 100, Height = 30, Left = 10, Top = 10 };
            Button saveBtn = new Button { Text = "💾 儲存", Width = 100, Height = 30, Left = 120, Top = 10, ForeColor = Color.Green };
            Button closeBtn = new Button { Text = "關閉", Width = 100, Height = 30, Left = 230, Top = 10 };
            Button insertDownUpBtn = new Button { Text = "插入 按下+放開", Width = 120, Height = 30, Left = 340, Top = 10 };
            Button clearBtn = new Button { Text = "清空", Width = 80, Height = 30, Left = 470, Top = 10, ForeColor = Color.Red };

            Label statusLabel = new Label
            {
                Text = $"目前事件數: {recordedEvents.Count}",
                Left = 560,
                Top = 15,
                Width = 190,
                ForeColor = Color.Gray
            };

            // 快捷按鍵提示
            Label shortcutLabel = new Label
            {
                Text = "常用按鍵名稱: A-Z, D0-D9, F1-F12, Space, Enter, Left, Right, Up, Down, LShiftKey, LControlKey, LMenu(Alt)",
                Left = 10,
                Top = 50,
                Width = 740,
                ForeColor = Color.DarkGray,
                Font = new Font("Microsoft JhengHei", 8)
            };

            // 解析文本的函數
            Func<string, List<MacroEvent>?> ParseScript = (text) =>
            {
                var events = new List<MacroEvent>();
                var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                int lineNum = 0;

                foreach (var rawLine in lines)
                {
                    lineNum++;
                    var line = rawLine.Trim();

                    // 跳過空行和註解
                    if (string.IsNullOrEmpty(line) || line.StartsWith("#"))
                        continue;

                    // 解析格式: 按鍵名稱 | 類型 | 時間戳
                    var parts = line.Split('|');
                    if (parts.Length != 3)
                    {
                        MessageBox.Show($"第 {lineNum} 行格式錯誤:\n{line}\n\n正確格式: 按鍵名稱 | 類型 | 時間戳",
                            "解析錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }

                    string keyName = parts[0].Trim();
                    string eventType = parts[1].Trim().ToLower();
                    string timestampStr = parts[2].Trim();

                    // 解析按鍵
                    if (!Enum.TryParse<Keys>(keyName, true, out Keys keyCode))
                    {
                        MessageBox.Show($"第 {lineNum} 行按鍵名稱無效: {keyName}\n\n常用名稱: A-Z, D0-D9, F1-F12, Space, Enter, Left, Right, Up, Down",
                            "解析錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }

                    // 驗證事件類型
                    if (eventType != "down" && eventType != "up")
                    {
                        MessageBox.Show($"第 {lineNum} 行事件類型無效: {eventType}\n\n有效類型: down, up",
                            "解析錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }

                    // 解析時間戳
                    if (!double.TryParse(timestampStr, out double timestamp) || timestamp < 0)
                    {
                        MessageBox.Show($"第 {lineNum} 行時間戳無效: {timestampStr}\n\n必須是非負數字",
                            "解析錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }

                    events.Add(new MacroEvent
                    {
                        KeyCode = keyCode,
                        EventType = eventType,
                        Timestamp = timestamp
                    });
                }

                return events;
            };

            // 解析預覽按鈕
            parseBtn.Click += (s, args) =>
            {
                var events = ParseScript(txtEditor.Text);
                if (events != null)
                {
                    statusLabel.Text = $"✅ 解析成功: {events.Count} 個事件";
                    statusLabel.ForeColor = Color.Green;
                    MessageBox.Show($"解析成功！\n共 {events.Count} 個事件。\n\n點擊「儲存」套用變更。",
                        "解析成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    statusLabel.Text = "❌ 解析失敗";
                    statusLabel.ForeColor = Color.Red;
                }
            };

            // 儲存按鈕
            saveBtn.Click += (s, args) =>
            {
                var events = ParseScript(txtEditor.Text);
                if (events == null) return;

                if (events.Count == 0)
                {
                    MessageBox.Show("腳本中沒有有效的事件！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 按時間戳排序
                events.Sort((a, b) => a.Timestamp.CompareTo(b.Timestamp));

                recordedEvents.Clear();
                recordedEvents.AddRange(events);

                hasUnsavedChanges = false;
                lblRecordingStatus.Text = $"已編輯 | 事件數: {recordedEvents.Count}";
                statusLabel.Text = $"✅ 已儲存: {recordedEvents.Count} 個事件";
                statusLabel.ForeColor = Color.Green;
                AddLog($"✅ 已儲存編輯 ({recordedEvents.Count} 個事件)");
                MessageBox.Show($"已儲存！共 {recordedEvents.Count} 個事件。", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            // 插入按下+放開按鈕
            insertDownUpBtn.Click += (s, args) =>
            {
                // 取得最後一個時間戳
                double lastTimestamp = 0;
                var events = ParseScript(txtEditor.Text);
                if (events != null && events.Count > 0)
                {
                    lastTimestamp = events.Max(e => e.Timestamp);
                }

                string template = $"\n# 在此輸入按鍵名稱 (例如: A, Space, F1)\nA | down | {(lastTimestamp + 0.1):F3}\nA | up | {(lastTimestamp + 0.15):F3}";
                txtEditor.AppendText(template);
                txtEditor.SelectionStart = txtEditor.Text.Length;
                txtEditor.ScrollToCaret();
            };

            // 清空按鈕
            clearBtn.Click += (s, args) =>
            {
                var result = MessageBox.Show("確定要清空所有內容嗎？", "確認清空",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    txtEditor.Text = "# ===== 腳本事件列表 =====\n# 格式: 按鍵名稱 | 類型 | 時間戳(秒)\n# ========================\n\n";
                }
            };

            // 關閉按鈕
            closeBtn.Click += (s, args) =>
            {
                if (hasUnsavedChanges)
                {
                    var result = MessageBox.Show("有未儲存的變更，是否儲存？", "未儲存的變更",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        saveBtn.PerformClick();
                        if (!hasUnsavedChanges) // 儲存成功才關閉
                            editorForm.Close();
                    }
                    else if (result == System.Windows.Forms.DialogResult.No)
                    {
                        editorForm.Close();
                    }
                }
                else
                {
                    editorForm.Close();
                }
            };

            btnPanel.Controls.Add(parseBtn);
            btnPanel.Controls.Add(saveBtn);
            btnPanel.Controls.Add(closeBtn);
            btnPanel.Controls.Add(insertDownUpBtn);
            btnPanel.Controls.Add(clearBtn);
            btnPanel.Controls.Add(statusLabel);
            btnPanel.Controls.Add(shortcutLabel);

            editorForm.Controls.Add(hintLabel);
            editorForm.Controls.Add(txtEditor);
            editorForm.Controls.Add(btnPanel);

            editorForm.ShowDialog();
            AddLog($"編輯完成 - 剩餘 {recordedEvents.Count} 個事件");
        }

        /// <summary>
        /// 整合連續重複按鍵事件的資料結構
        /// </summary>
        private class ConsolidatedKeyEvent
        {
            public Keys KeyCode { get; set; }
            public Keys OriginalKeyCode { get; set; }
            public Keys Modifiers { get; set; }
            public Keys OriginalModifiers { get; set; }
            public double StartTime { get; set; }
            public double EndTime { get; set; }
            public double Duration => EndTime - StartTime;
        }

        /// <summary>
        /// 將連續重複的按鍵事件整合為單一動作
        /// </summary>
        // IsModifierKey 和 ModifierKeyToFlag 已提取至 KeySender.cs

        private List<ConsolidatedKeyEvent> ConsolidateKeyEvents(List<MacroEvent> events)
        {
            var consolidated = new List<ConsolidatedKeyEvent>();
            if (events.Count == 0) return consolidated;

            // 追蹤每個按鍵的按下時間
            var keyDownTimes = new Dictionary<Keys, double>();

            // 追蹤目前按住的修飾鍵及其按下時間
            var activeModifiers = new Dictionary<Keys, double>();

            foreach (var evt in events.OrderBy(e => e.Timestamp))
            {
                if (evt.EventType == "down")
                {
                    if (KeySender.IsModifierKey(evt.KeyCode))
                    {
                        // 記錄修飾鍵按下
                        if (!activeModifiers.ContainsKey(evt.KeyCode))
                            activeModifiers[evt.KeyCode] = evt.Timestamp;
                    }

                    // 記錄按下時間（如果尚未追蹤）
                    if (!keyDownTimes.ContainsKey(evt.KeyCode))
                    {
                        keyDownTimes[evt.KeyCode] = evt.Timestamp;
                    }
                }
                else if (evt.EventType == "up")
                {
                    if (KeySender.IsModifierKey(evt.KeyCode))
                    {
                        activeModifiers.Remove(evt.KeyCode);
                    }

                    // 放開時計算持續時間
                    if (keyDownTimes.TryGetValue(evt.KeyCode, out double startTime))
                    {
                        // 計算此按鍵按下期間有哪些修飾鍵是按住的
                        Keys modifiers = Keys.None;
                        if (!KeySender.IsModifierKey(evt.KeyCode))
                        {
                            foreach (var mod in activeModifiers)
                            {
                                // 修飾鍵必須在主鍵按下之前或同時按下
                                if (mod.Value <= startTime)
                                {
                                    modifiers |= KeySender.ModifierKeyToFlag(mod.Key);
                                }
                            }
                        }

                        consolidated.Add(new ConsolidatedKeyEvent
                        {
                            KeyCode = evt.KeyCode,
                            OriginalKeyCode = evt.KeyCode,
                            Modifiers = modifiers,
                            OriginalModifiers = modifiers,
                            StartTime = startTime,
                            EndTime = evt.Timestamp
                        });
                        keyDownTimes.Remove(evt.KeyCode);
                    }
                }
            }

            // 處理未放開的按鍵
            foreach (var kvp in keyDownTimes)
            {
                Keys modifiers = Keys.None;
                if (!KeySender.IsModifierKey(kvp.Key))
                {
                    foreach (var mod in activeModifiers)
                    {
                        if (mod.Value <= kvp.Value)
                            modifiers |= KeySender.ModifierKeyToFlag(mod.Key);
                    }
                }

                consolidated.Add(new ConsolidatedKeyEvent
                {
                    KeyCode = kvp.Key,
                    OriginalKeyCode = kvp.Key,
                    Modifiers = modifiers,
                    OriginalModifiers = modifiers,
                    StartTime = kvp.Value,
                    EndTime = events.Max(e => e.Timestamp)
                });
            }

            // 過濾掉純修飾鍵事件（它們已經被合併到主鍵的 Modifiers 中）
            // 但保留獨立的修飾鍵（沒有搭配主鍵的）
            return consolidated.OrderBy(e => e.StartTime).ToList();
        }

        private void BtnStartPlayback_Click(object? sender, EventArgs e)
        {
            if (isPlaying || recordedEvents.Count == 0)
                return;

            isPlaying = true;

            keySender.PressedKeys.Clear();
            lock (lastKeyDownTimestamp) { lastKeyDownTimestamp.Clear(); }

            // ★ 同步 keySender 狀態
            keySender.TargetWindowHandle = targetWindowHandle;
            keySender.CurrentArrowKeyMode = currentArrowKeyMode;
            keySender.KeyboardBlocker = keyboardBlocker;

            // 重置自定義按鍵槽位的觸發狀態
            foreach (var slot in customKeySlots)
            {
                slot.Reset();
            }

            // 開始統計
            statistics.StartSession();

            // 如果使用 Blocker 模式且有目標視窗（背景模式），才初始化並啟動 KeyboardBlocker
            // 前景模式不需要 Blocker，避免安裝不必要的低層鍵盤鉤子
            if ((currentArrowKeyMode == ArrowKeyMode.ThreadAttachWithBlocker ||
                currentArrowKeyMode == ArrowKeyMode.SendInputWithBlock) &&
                targetWindowHandle != IntPtr.Zero)
            {
                if (keyboardBlocker == null)
                {
                    keyboardBlocker = new KeyboardBlocker();
                }
                keyboardBlocker.TargetWindowHandle = targetWindowHandle;
                keyboardBlocker.Install();
                keyboardBlocker.IsBlocking = true;
                keySender.KeyboardBlocker = keyboardBlocker;
                AddLog($"鍵盤阻擋器已啟用 (Blocker 模式)");
            }

            string mode = (targetWindowHandle != IntPtr.Zero) ? "背景" : "前景";
            AddLog($"播放開始 ({mode}模式, 方向鍵={currentArrowKeyMode})...");

            // 顯示啟用的自定義按鍵
            int enabledCount = customKeySlots.Count(s => s.Enabled);
            if (enabledCount > 0)
            {
                AddLog($"自定義按鍵：{enabledCount} 個已啟用");
            }

            try
            {
                int loopCount = LoopCountValue;
                // 建立事件快照，避免播放線程與 UI 線程競爭存取 recordedEvents
                var eventsSnapshot = recordedEvents.ToList();
                Thread playbackThread = new Thread(() => PlaybackThread(loopCount, eventsSnapshot))
                {
                    IsBackground = true
                };
                playbackThread.Start();
            }
            catch (Exception ex)
            {
                AddLog($"❌ 播放失敗: {ex.Message}");
                ReleasePressedKeys();
                isPlaying = false;
                statistics.EndSession();
            }

            UpdateUI();
            UpdatePauseButtonState();
        }

        private void PlaybackThread(int loopCount, List<MacroEvent> events)
        {
            bool firstKeySent = false;
            try
            {
                string mode = (targetWindowHandle != IntPtr.Zero) ? "背景" : "前景";
                Dispatcher.BeginInvoke(new Action(() => AddLog($"播放線程已啟動 ({mode}模式, {events.Count} 事件, {loopCount} 循環)")));

                // ★ 啟動 Y 軸突變偵測背景執行緒（緊急掉落防護）
                _hardInterruptY = false;
                _expectedY = -1;
                _yAnomalyThreadRunning = true;
                _playbackGate.Set(); // 確保閘門是開的
                var yAnomalyThread = new System.Threading.Thread(YAxisAnomalyDetectionThread)
                {
                    IsBackground = true,
                    Name = "Y-Anomaly-Detector"
                };
                yAnomalyThread.Start();

                for (int loop = 1; loop <= loopCount && isPlaying; loop++)
                {
                    // 檢查暫停狀態
                    CheckPauseState();

                    if (!isPlaying) break;

                    statistics.IncrementLoop();

                    // ★ 每次循環開始時檢查設定是否變更，即時套用（不重建修正器）
                    if (_correctionSettingsChanged)
                    {
                        _correctionSettingsChanged = false;
                        if (positionCorrector != null)
                        {
                            ApplyCorrectorSettings(positionCorrector);
                        }
                        Dispatcher.BeginInvoke(new Action(() => AddLog("🔄 位置修正設定已套用")));
                    }

                    // ★ 重設循環修正計數
                    if (positionCorrector != null)
                        positionCorrector.ResetLoopCorrectionCount();

                    if (!isPlaying) break;

                    int currentLoopNum = loop;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        lblPlaybackStatus.Text = $"循環: {currentLoopNum}/{loopCount}";
                        lblPlaybackStatus.Foreground = ToBrush(Color.Blue);
                        if (currentLoopNum == 1 || currentLoopNum % 10 == 0)
                            AddLog($"循環 {currentLoopNum}/{loopCount} 開始");
                    }));

                    double lastTimestamp = 0;
                    long loopStartTick = Stopwatch.GetTimestamp();
                    long lastCustomKeyCheckTick = loopStartTick;
                    long lastCorrectionCheckTick = loopStartTick; // ★ 位置修正檢測計時
                    long lastDriftTrackTick = loopStartTick;      // ★ 漂移追蹤節流計時

                    foreach (MacroEvent evt in events)
                    {
                        // 檢查暫停狀態
                        CheckPauseState();

                        if (!isPlaying) break;

                        double waitTime = evt.Timestamp - lastTimestamp;

                        // ★ 最短按鍵持續時間保護：如果是 keyup 且對應的 keydown 間隔太短，
                        // 強制等待至少 MIN_KEY_HOLD_SECONDS，確保遊戲能偵測到
                        if (evt.EventType == "up" && waitTime < MIN_KEY_HOLD_SECONDS)
                        {
                            double holdTime;
                            lock (lastKeyDownTimestamp)
                            {
                                if (lastKeyDownTimestamp.TryGetValue(evt.KeyCode, out double downTs))
                                {
                                    holdTime = evt.Timestamp - downTs;
                                }
                                else
                                {
                                    holdTime = waitTime;
                                }
                            }
                            if (holdTime < MIN_KEY_HOLD_SECONDS)
                            {
                                waitTime = Math.Max(waitTime, MIN_KEY_HOLD_SECONDS - holdTime + waitTime);
                            }
                        }

                        if (waitTime > 0)
                        {
                            // 高精度等待：混合 Sleep + Spin-wait
                            long waitStartTick = Stopwatch.GetTimestamp();
                            double waitSeconds = waitTime;

                            while (isPlaying)
                            {
                                // 檢查暫停狀態
                                CheckPauseState();

                                // ★ 緊急掉落防護：Y 軸突變 → 硬中斷，優先修正
                                if (_hardInterruptY && positionCorrectionSettings.Enabled
                                    && minimapTracker != null && minimapTracker.IsCalibrated)
                                {
                                    _hardInterruptY = false;
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        AddLog("🚨 Y軸突變！角色可能掉落，緊急修正中...");
                                        lblPlaybackStatus.Text = "緊急掉落修正...";
                                        lblPlaybackStatus.Foreground = ToBrush(Color.Red);
                                    }));
                                    _playbackGate.Reset();
                                    isCorrecting = true;
                                    customKeysPaused = true;
                                    long yStartTick = Stopwatch.GetTimestamp();

                                    double scriptTimeY = Stopwatch.GetElapsedTime(loopStartTick).TotalSeconds;
                                    PeriodicPositionCheck(events, scriptTimeY);
                                    Thread.Sleep(POST_CORRECTION_BUFFER_MS);

                                    long yElapsedTicks = Stopwatch.GetTimestamp() - yStartTick;
                                    customKeysPaused = false;
                                    isCorrecting = false;
                                    _playbackGate.Set();
                                    waitStartTick += yElapsedTicks;
                                    loopStartTick += yElapsedTicks;
                                    lastCustomKeyCheckTick = Stopwatch.GetTimestamp();
                                    lastCorrectionCheckTick = Stopwatch.GetTimestamp();

                                    int curLoopY = currentLoopNum;
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        lblPlaybackStatus.Text = $"循環: {curLoopY}/{loopCount}";
                                        lblPlaybackStatus.Foreground = ToBrush(Color.Blue);
                                    }));
                                }

                                double elapsedSeconds = Stopwatch.GetElapsedTime(waitStartTick).TotalSeconds;
                                double remaining = waitSeconds - elapsedSeconds;

                                if (remaining <= 0)
                                    break;

                                // 節流：每 ~20ms 檢查一次自定義按鍵（避免 spin 區間瘋狂呼叫）
                                if (Stopwatch.GetElapsedTime(lastCustomKeyCheckTick).TotalMilliseconds >= 20)
                                {
                                    lastCustomKeyCheckTick = Stopwatch.GetTimestamp();
                                    double currentScriptTime = Stopwatch.GetElapsedTime(loopStartTick).TotalSeconds;
                                    CheckAndTriggerCustomKeys(currentScriptTime);
                                }

                                // ★ 定期位置修正檢查（每 N 秒檢查一次是否偏離腳本位置）
                                if (positionCorrectionSettings.Enabled && minimapTracker != null && minimapTracker.IsCalibrated
                                    && positionCorrectionSettings.CorrectionCheckIntervalSec > 0)
                                {
                                    // ★ 即時套用設定變更（不中斷腳本）
                                    if (_correctionSettingsChanged)
                                    {
                                        _correctionSettingsChanged = false;
                                        if (positionCorrector != null)
                                        {
                                            ApplyCorrectorSettings(positionCorrector);
                                        }
                                        Dispatcher.BeginInvoke(new Action(() => AddLog("🔄 位置修正設定已即時套用")));
                                    }

                                    double secSinceLastCheck = Stopwatch.GetElapsedTime(lastCorrectionCheckTick).TotalSeconds;
                                    if (secSinceLastCheck >= positionCorrectionSettings.CorrectionCheckIntervalSec
                                        && !IsAnimationLocked) // ★ X 軸軟輪詢：嚴格遵守硬直鎖
                                    {
                                        // ★ 阻塞式修正：關閉閘門 → 暫停播放執行緒 spin → 修正 → 緩衝 → 開閘
                                        _playbackGate.Reset();
                                        isCorrecting = true;
                                        customKeysPaused = true;
                                        long corrStartTick = Stopwatch.GetTimestamp();

                                        double scriptTime = Stopwatch.GetElapsedTime(loopStartTick).TotalSeconds;
                                        PeriodicPositionCheck(events, scriptTime);

                                        // ★ 修正後緩衝延遲：等待角色慣性靜止
                                        Thread.Sleep(POST_CORRECTION_BUFFER_MS);

                                        long corrElapsedTicks = Stopwatch.GetTimestamp() - corrStartTick;
                                        customKeysPaused = false;
                                        isCorrecting = false;
                                        _playbackGate.Set(); // ★ 開閘，允許播放繼續

                                        // ★ 時間補償：將修正耗時從等待計時和循環計時中扣除
                                        waitStartTick += corrElapsedTicks;
                                        loopStartTick += corrElapsedTicks;
                                        lastCustomKeyCheckTick = Stopwatch.GetTimestamp();

                                        // ★ 修正完成後才重設計時器
                                        lastCorrectionCheckTick = Stopwatch.GetTimestamp();

                                        // ★ 修正後恢復播放狀態標籤
                                        int curLoop = currentLoopNum;
                                        Dispatcher.BeginInvoke(new Action(() =>
                                        {
                                            lblPlaybackStatus.Text = $"循環: {curLoop}/{loopCount}";
                                            lblPlaybackStatus.Foreground = ToBrush(Color.Blue);
                                        }));
                                    }
                                }

                                if (remaining > 0.05) // > 50ms：Sleep 較長以節省 CPU
                                {
                                    Thread.Sleep(30);
                                }
                                else if (remaining > 0.002) // 2~50ms：短 Sleep
                                {
                                    Thread.Sleep(1);
                                }
                                else // < 2ms：短 Sleep 確保 hook 線程有 CPU 時間處理按鍵
                                {
                                    Thread.Sleep(1);
                                }
                            }
                        }

                        // ★ 位置修正事件 - 不發送按鍵，而是執行位置修正
                        if (evt.EventType == "position_correct")
                        {
                            // ★ 修正未啟用時跳過腳本內的位置修正事件
                            if (!positionCorrectionSettings.Enabled)
                            {
                                Dispatcher.BeginInvoke(new Action(() =>
                                    AddLog("⏭️ 位置修正事件已跳過（修正未啟用）")));
                                lastTimestamp = evt.Timestamp;
                                continue;
                            }

                            // ★ 使用腳本錄製座標作為修正目標（不再有固定目標模式）
                            int tx = evt.RecordedX >= 0 ? evt.RecordedX : evt.CorrectTargetX;
                            int ty = evt.RecordedY >= 0 ? evt.RecordedY : evt.CorrectTargetY;

                            if (tx < 0 || ty < 0)
                            {
                                Dispatcher.BeginInvoke(new Action(() =>
                                    AddLog("⏭️ 位置修正事件已跳過（無錄製座標）")));
                                lastTimestamp = evt.Timestamp;
                                continue;
                            }

                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                lblPlaybackStatus.Text = $"腳本位置修正中... 目標({tx},{ty})";
                                lblPlaybackStatus.Foreground = ToBrush(Color.Orange);
                            }));

                            if (minimapTracker != null && minimapTracker.IsCalibrated)
                            {
                                // ★ 阻塞式修正 + 時間補償（暫停自定義按鍵）
                                isCorrecting = true;
                                customKeysPaused = true;
                                long corrStartTick = Stopwatch.GetTimestamp();

                                var corrResult = ExecutePositionCorrectionTo(tx, ty);

                                long corrElapsedTicks = Stopwatch.GetTimestamp() - corrStartTick;
                                customKeysPaused = false;
                                isCorrecting = false;

                                // ★ 時間補償
                                loopStartTick += corrElapsedTicks;
                                lastCustomKeyCheckTick = Stopwatch.GetTimestamp();
                                lastCorrectionCheckTick = Stopwatch.GetTimestamp();

                                if (corrResult != null)
                                {
                                    var r = corrResult;
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        AddLog(r.Success
                                            ? $"✅ 腳本位置修正完成: ({r.FinalX}, {r.FinalY})"
                                            : $"⚠️ 腳本位置修正: {r.Message}");
                                        // ★ 修正後恢復播放狀態
                                        lblPlaybackStatus.Text = $"播放中...";
                                        lblPlaybackStatus.Foreground = ToBrush(Color.Blue);
                                    }));
                                }
                                else
                                {
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        lblPlaybackStatus.Text = $"播放中...";
                                        lblPlaybackStatus.Foreground = ToBrush(Color.Blue);
                                    }));
                                }
                            }
                            else
                            {
                                Dispatcher.BeginInvoke(new Action(() =>
                                    AddLog("⚠️ 位置修正卡片：小地圖未校準，跳過")));
                            }

                            lastTimestamp = evt.Timestamp;
                            continue; // 跳過後面的 SendKeyEvent
                        }

                        // 記錄 keydown 時間戳，用於計算持續時間
                        if (evt.EventType == "down")
                        {
                            lock (lastKeyDownTimestamp)
                            {
                                lastKeyDownTimestamp[evt.KeyCode] = evt.Timestamp;
                            }
                        }
                        else if (evt.EventType == "up")
                        {
                            lock (lastKeyDownTimestamp)
                            {
                                lastKeyDownTimestamp.Remove(evt.KeyCode);
                            }
                        }

                        SendKeyEvent(evt);
                        lastTimestamp = evt.Timestamp;

                        // ★ 更新期望 Y 座標（供 Y 軸突變偵測執行緒使用）
                        if (evt.RecordedY >= 0)
                            _expectedY = evt.RecordedY;

                        // ★ 技能硬直延遲（MacroEvent.SkillAnimationDelay > 0 時鎖定修正器）
                        if (evt.EventType == "down" && evt.SkillAnimationDelay > 0)
                        {
                            IsAnimationLocked = true;
                            try { Thread.Sleep(evt.SkillAnimationDelay); }
                            finally { IsAnimationLocked = false; }
                        }

                        // ★ 追蹤按鍵後的實際XY座標（節流：每 2 秒最多一次，避免 ReadPosition 拖慢腳本）
                        if (positionCorrectionSettings.Enabled && minimapTracker != null && minimapTracker.IsCalibrated
                            && evt.RecordedX >= 0 && evt.RecordedY >= 0
                            && Stopwatch.GetElapsedTime(lastDriftTrackTick).TotalSeconds >= 2.0)
                        {
                            lastDriftTrackTick = Stopwatch.GetTimestamp();
                            var (ax, ay, aOk) = minimapTracker.ReadPosition();
                            if (aOk)
                            {
                                int driftX = ax - evt.RecordedX;
                                int driftY = ay - evt.RecordedY;
                                // 只在偏差明顯時記錄，避免日誌過多
                                if (Math.Abs(driftX) > positionCorrectionSettings.HorizontalTolerance ||
                                    Math.Abs(driftY) > positionCorrectionSettings.VerticalTolerance)
                                {
                                    var evtCopy = evt;
                                    Dispatcher.BeginInvoke(new Action(() =>
                                        AddLog($"📍 位置追蹤: 實際({ax},{ay}) 腳本({evtCopy.RecordedX},{evtCopy.RecordedY}) 偏差({driftX:+#;-#;0},{driftY:+#;-#;0})")
                                    ));
                                }
                            }
                        }

                        // 第一個按鍵發送後記錄
                        if (!firstKeySent)
                        {
                            firstKeySent = true;
                            var firstEvt = evt;
                            Dispatcher.BeginInvoke(new Action(() =>
                                AddLog($"✅ 第一個按鍵已發送: {GetKeyDisplayName(firstEvt.KeyCode)} ({firstEvt.EventType})")
                            ));
                        }
                    }

                    Thread.Sleep(200);
                }

                isPlaying = false;
                _yAnomalyThreadRunning = false;
                _playbackGate.Set(); // ★ 確保不卡住任何等待
                statistics.EndSession();

                // 清理 Blocker
                if (keyboardBlocker != null)
                {
                    keyboardBlocker.IsBlocking = false;
                    keyboardBlocker.Uninstall();
                }

                ReleasePressedKeys();

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    lblPlaybackStatus.Text = "播放: 已完成";
                    lblPlaybackStatus.Foreground = ToBrush(Color.Green);
                    AddLog($"播放完成 - 循環: {statistics.CurrentLoopCount}");
                    UpdateUI();
                    UpdatePauseButtonState();
                }));
            }
            catch (Exception ex)
            {
                _yAnomalyThreadRunning = false;
                _playbackGate.Set(); // ★ 確保不卡住任何等待
                statistics.EndSession();

                // 清理 Blocker
                if (keyboardBlocker != null)
                {
                    keyboardBlocker.IsBlocking = false;
                    keyboardBlocker.Uninstall();
                }

                ReleasePressedKeys();
                isPlaying = false;

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    AddLog($"❌ 播放錯誤: {ex.Message}");
                    UpdateUI();
                    UpdatePauseButtonState();
                }));
            }
        }

        #region ★ 雙層位置偵測

        /// <summary>
        /// ★ Y 軸突變防護執行緒（高頻 500ms，不受技能鎖影響）
        /// 偵測角色因 Knockback 掉落平台，設定 _hardInterruptY 旗標通知播放執行緒緊急修正。
        /// </summary>
        private void YAxisAnomalyDetectionThread()
        {
            while (_yAnomalyThreadRunning && isPlaying)
            {
                Thread.Sleep(Y_ANOMALY_INTERVAL_MS);

                if (!_yAnomalyThreadRunning || !isPlaying) break;
                if (!positionCorrectionSettings.Enabled) continue;
                if (minimapTracker == null || !minimapTracker.IsCalibrated) continue;
                if (_hardInterruptY) continue; // 旗標已設，等播放執行緒處理
                if (isCorrecting) continue;    // 修正進行中，跳過

                int expY = _expectedY;
                if (expY < 0) continue; // 尚未有期望座標

                var (cx, cy, ok) = minimapTracker.ReadPosition();
                if (!ok) continue;

                int deltaY = cy - expY; // 正值 = 掉落（Y 增大）
                if (deltaY > Y_ANOMALY_THRESHOLD)
                {
                    // ★ 緊急掉落：Y 大幅增加，代表角色掉下平台
                    _hardInterruptY = true;
                    try
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                            AddLog($"🚨 [Y突變] 現在({cx},{cy}) 期望Y={expY} 偏差={deltaY}px > {Y_ANOMALY_THRESHOLD}px，發出硬中斷")));
                    }
                    catch { }
                }
            }
        }

        // ★ 爬繩偵測可調參數（模板比對模式）
        private double _climbingMatchThreshold = 0.70;  // 模板比對匹配閾值（0~1，越高越嚴格）
        private Bitmap? _climbingTemplate = null;       // 爬繩模板圖片（記憶體中）

        private static readonly string ClimbingTemplatePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MapleMacro", "climbing_template.png");

        private static readonly string RoiDebugDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MapleMacro", "ROI_Debug");

        /// <summary>
        /// 載入爬繩模板（如果存在）
        /// </summary>
        private void LoadClimbingTemplate()
        {
            try
            {
                if (File.Exists(ClimbingTemplatePath))
                {
                    _climbingTemplate?.Dispose();
                    // 讀取後立刻複製，避免鎖定檔案
                    using (var tmp = new Bitmap(ClimbingTemplatePath))
                        _climbingTemplate = new Bitmap(tmp);
                    _templateEdgeCache = null; // 清除快取，下次偵測時重建
                    _templateHueHistogram = null; // ★ 清除顏色快取
                }
            }
            catch { _climbingTemplate = null; }
        }

        /// <summary>
        /// 爬繩偵測掃描結果（供診斷用）
        /// </summary>
        public class ClimbingScanResult
        {
            public bool IsClimbing;
            public double MatchScore;                   // 模板匹配分數 (0~1)
            public double Threshold;                    // 匹配閾值
            public int FoundX, FoundY;                  // 模板在畫面中找到的位置
            public Bitmap? CaptureBitmap;               // 匹配位置截圖（供預覽）
            public bool TemplateLoaded;                 // 模板是否已載入
            // 診斷資訊
            public int GameW, GameH;                    // 實際偵測到的視窗尺寸
            public int RoiLeft, RoiTop, RoiW, RoiH;    // 搜尋區域
            public string? Error;                       // 錯誤訊息
            public double RawNccScore;                  // 原始 NCC 分數 (未經位置加權)
            public double CenterDist;                   // 最佳匹配離搜尋中心的距離 (px)
            public double ColorSimilarity;              // 匹配區域與模板的顏色相似度 (0~1)
        }

        /// <summary>
        /// ★ 爬繩偵測：在畫面中心區域搜尋爬繩模板。
        /// 楓之谷的角色永遠在畫面中心附近（攝影機跟隨），搜尋中心區域 ± 動態 padding。
        /// 使用實際視窗尺寸計算中心位置（不硬編碼解析度）。
        /// </summary>
        private bool DetectIsClimbing() => ScanClimbingRoi().IsClimbing;

        // 模板邊緣快取（原始解析度）
        private float[]? _templateEdgeCache = null;
        private int _templateEdgeCacheW, _templateEdgeCacheH;
        // 稀疏取樣索引（預計算）
        private int[]? _templateSampleIdx = null;
        private float[]? _templateSampleVal = null;
        private int _templateSampleCount;
        private double _templateSampleMean;
        private double _templateSampleStdDev;
        private const int SAMPLE_STEP = 2; // 取樣間隔（搜尋區域小，可以用更密的取樣）
        // ★ 模板顏色直方圖快取（防止顏色不同的怪物通過邊緣匹配）
        private float[]? _templateHueHistogram = null;
        private const int HUE_BINS = 36; // 10° per bin

        /// <summary>
        /// 核心爬繩偵測：擷取畫面中心區域 → 邊緣偵測 → NCC 搜尋。
        /// </summary>
        private ClimbingScanResult ScanClimbingRoi()
        {
            var result = new ClimbingScanResult();
            result.Threshold = _climbingMatchThreshold;
            result.TemplateLoaded = _climbingTemplate != null;

            if (_climbingTemplate == null) { result.Error = "模板未載入"; return result; }
            if (minimapTracker == null) { result.Error = "追蹤器未初始化"; return result; }
            if (targetWindowHandle == IntPtr.Zero) { result.Error = "視窗未鎖定"; return result; }

            try
            {
                // 1. 確保模板邊緣快取
                EnsureTemplateEdgeCache();
                int tmpW = _templateEdgeCacheW, tmpH = _templateEdgeCacheH;
                if (_templateSampleIdx == null) { result.Error = "模板快取失敗"; return result; }

                // 2. 取得實際視窗尺寸（取代硬編碼 1600x900）
                var (gameW, gameH) = minimapTracker.GetClientAreaSize();
                if (gameW <= 0 || gameH <= 0) { result.Error = $"視窗尺寸無效: {gameW}x{gameH}"; return result; }
                result.GameW = gameW;
                result.GameH = gameH;

                // 3. 搜尋幾乎整個畫面（角色在地圖邊緣時可在螢幕任意位置）
                // 楓之谷攝影機在地圖邊緣會停止跟隨，角色可偏離中心 400+ px
                // 必須使用寬廣搜尋範圍確保角色始終在 ROI 內
                int padX = Math.Max(300, gameW * 3 / 8);  // 覆蓋螢幕寬度 ~75%
                int padY = Math.Max(200, gameH / 3);       // 覆蓋螢幕高度 ~67%
                int roiW = tmpW + padX * 2;
                int roiH = tmpH + padY * 2;
                int roiLeft = Math.Max(0, gameW / 2 - roiW / 2);
                int roiTop  = Math.Max(0, gameH / 2 - roiH / 2);
                roiW = Math.Min(roiW, gameW - roiLeft);
                roiH = Math.Min(roiH, gameH - roiTop);
                result.RoiLeft = roiLeft;
                result.RoiTop = roiTop;
                result.RoiW = roiW;
                result.RoiH = roiH;

                if (roiW < tmpW || roiH < tmpH) { result.Error = $"搜尋區域太小: {roiW}x{roiH} < 模板{tmpW}x{tmpH}"; return result; }

                // 4. 只擷取中心區域（比擷取全畫面快）
                using var roiBmp = minimapTracker.CaptureRegion(
                    new Rectangle(roiLeft, roiTop, roiW, roiH));
                if (roiBmp == null) { result.Error = "截圖失敗"; return result; }

                int srcW = roiBmp.Width, srcH = roiBmp.Height;
                if (tmpW > srcW || tmpH > srcH) return result;

                // 4. 計算邊緣圖
                var srcEdge = ComputeEdgeMap(roiBmp);

                // 5. 兩階段搜尋：粗搜（step=4）掃描整個 ROI → 精搜（step=1）精確定位
                // 搜尋範圍大幅擴展後需要兩階段避免效能問題
                int searchW = srcW - tmpW;
                int searchH = srcH - tmpH;

                double bestRawNcc = -2;
                int bestX = 0, bestY = 0;

                // Phase 1: 粗搜 — 快速掃描整個 ROI（step=4）
                const int COARSE_STEP = 4;
                for (int sy = 0; sy <= searchH; sy += COARSE_STEP)
                {
                    for (int sx = 0; sx <= searchW; sx += COARSE_STEP)
                    {
                        double ncc = ComputeSparseNCC(srcEdge, srcW, sx, sy);
                        if (ncc > bestRawNcc)
                        {
                            bestRawNcc = ncc;
                            bestX = sx; bestY = sy;
                        }
                    }
                }

                // Phase 2: 精搜 — 在粗搜最佳位置 ±4px 精確定位（step=1）
                {
                    int fxMin = Math.Max(0, bestX - COARSE_STEP);
                    int fxMax = Math.Min(searchW, bestX + COARSE_STEP);
                    int fyMin = Math.Max(0, bestY - COARSE_STEP);
                    int fyMax = Math.Min(searchH, bestY + COARSE_STEP);
                    for (int sy = fyMin; sy <= fyMax; sy++)
                    {
                        for (int sx = fxMin; sx <= fxMax; sx++)
                        {
                            double ncc = ComputeSparseNCC(srcEdge, srcW, sx, sy);
                            if (ncc > bestRawNcc)
                            {
                                bestRawNcc = ncc;
                                bestX = sx; bestY = sy;
                            }
                        }
                    }
                }

                // ★ 不使用位置先驗：角色可在地圖任意位置（邊緣偏離中心 200+ px），
                // 位置加權會導致中心的怪物/背景雜訊被誤判為爬繩
                double expectSX = searchW / 2.0;
                double expectSY = searchH / 2.0;

                // NCC 範圍 -1~1，映射到 0~1（不套用位置加權）
                result.RawNccScore = (bestRawNcc + 1.0) / 2.0;
                result.FoundX = roiLeft + bestX;
                result.FoundY = roiTop + bestY;
                result.CenterDist = Math.Sqrt((bestX - expectSX) * (bestX - expectSX)
                                            + (bestY - expectSY) * (bestY - expectSY));

                // 6. 擷取匹配位置截圖供預覽
                int prevW = Math.Min(tmpW, srcW - bestX);
                int prevH = Math.Min(tmpH, srcH - bestY);
                Bitmap? preview = null;
                if (prevW > 0 && prevH > 0)
                {
                    preview = new Bitmap(prevW, prevH, PixelFormat.Format24bppRgb);
                    using (var g = Graphics.FromImage(preview))
                        g.DrawImage(roiBmp, new Rectangle(0, 0, prevW, prevH),
                            new Rectangle(bestX, bestY, prevW, prevH), GraphicsUnit.Pixel);
                    result.CaptureBitmap = preview;
                }

                // 7. ★ 顏色直方圖驗證：比對匹配區域與模板的色彩分佈
                // 防止邊緣輪廓相似但顏色完全不同的怪物通過匹配
                double colorSim = 1.0;
                if (preview != null && _templateHueHistogram != null)
                {
                    var matchHist = ComputeHueHistogram(preview);
                    colorSim = HistogramCorrelation(_templateHueHistogram, matchHist);
                }
                result.ColorSimilarity = colorSim;

                // ★ 多閘門評分：NCC 和顏色必須同時達標才判定為爬繩
                // Gate floor：NCC×color 乘積必須 ≥ 0.40，否則直接拒絕
                // 最終分數：使用幾何平均 √(NCC×color)，兩者都要高才能通過
                const double GATE_FLOOR = 0.40;
                double gateScore = result.RawNccScore * colorSim;
                if (gateScore >= GATE_FLOOR)
                {
                    double finalScore = Math.Sqrt(result.RawNccScore * Math.Max(0.01, colorSim));
                    result.MatchScore = finalScore;
                    result.IsClimbing = finalScore >= _climbingMatchThreshold;
                }
                else
                {
                    result.MatchScore = gateScore;
                    result.IsClimbing = false;
                }
            }
            catch { /* 截圖失敗時返回 false */ }

            return result;
        }

        /// <summary>
        /// 建立模板邊緣快取 + 預計算稀疏取樣索引。
        /// </summary>
        private void EnsureTemplateEdgeCache()
        {
            if (_climbingTemplate == null) return;
            int tw = _climbingTemplate.Width, th = _climbingTemplate.Height;

            if (_templateEdgeCache != null && _templateEdgeCacheW == tw && _templateEdgeCacheH == th)
                return;

            _templateEdgeCache = ComputeEdgeMap(_climbingTemplate);
            _templateEdgeCacheW = tw;
            _templateEdgeCacheH = th;

            // 建立稀疏取樣索引（每 SAMPLE_STEP 像素取一個）
            var idxList = new System.Collections.Generic.List<int>();
            var valList = new System.Collections.Generic.List<float>();
            for (int y = 0; y < th; y += SAMPLE_STEP)
                for (int x = 0; x < tw; x += SAMPLE_STEP)
                {
                    int idx = y * tw + x;
                    idxList.Add(idx);
                    valList.Add(_templateEdgeCache[idx]);
                }

            _templateSampleIdx = idxList.ToArray();
            _templateSampleVal = valList.ToArray();
            _templateSampleCount = _templateSampleIdx.Length;

            // 預計算取樣點的均值和標準差
            double sum = 0, sum2 = 0;
            for (int i = 0; i < _templateSampleCount; i++)
            {
                float v = _templateSampleVal[i];
                sum += v;
                sum2 += (double)v * v;
            }
            _templateSampleMean = sum / _templateSampleCount;
            double var_ = sum2 / _templateSampleCount - _templateSampleMean * _templateSampleMean;
            _templateSampleStdDev = Math.Sqrt(Math.Max(0, var_));

            // ★ 同時計算模板色相直方圖（用於顏色驗證）
            if (_templateHueHistogram == null)
                _templateHueHistogram = ComputeHueHistogram(_climbingTemplate);
        }

        /// <summary>
        /// 稀疏取樣 NCC：只用預計算的取樣點計算，比完整 NCC 快 SAMPLE_STEP² 倍。
        /// </summary>
        private double ComputeSparseNCC(float[] srcEdge, int srcW, int ox, int oy)
        {
            if (_templateSampleIdx == null || _templateSampleVal == null) return 0;
            if (_templateSampleStdDev < 1e-6) return 0;

            int tmpW = _templateEdgeCacheW;
            int n = _templateSampleCount;
            double sumS = 0, sumS2 = 0, sumST = 0;

            for (int i = 0; i < n; i++)
            {
                int tIdx = _templateSampleIdx[i];
                int ty = tIdx / tmpW, tx = tIdx % tmpW;
                float s = srcEdge[(oy + ty) * srcW + (ox + tx)];
                float t = _templateSampleVal[i];
                sumS += s;
                sumS2 += (double)s * s;
                sumST += (double)s * t;
            }

            double meanS = sumS / n;
            double varS = sumS2 / n - meanS * meanS;
            if (varS < 1e-6) return 0;
            double stdS = Math.Sqrt(varS);

            double cov = sumST / n - meanS * _templateSampleMean;
            return cov / (stdS * _templateSampleStdDev);
        }

        /// <summary>
        /// 計算色相直方圖（只考慮有彩度的像素，忽略灰色/黑色/白色）
        /// </summary>
        private static float[] ComputeHueHistogram(Bitmap bmp, int bins = HUE_BINS)
        {
            float[] hist = new float[bins];
            int w = bmp.Width, h = bmp.Height;
            var data = bmp.LockBits(new Rectangle(0, 0, w, h),
                ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            int stride = data.Stride;
            byte[] pixels = new byte[stride * h];
            Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
            bmp.UnlockBits(data);

            int validCount = 0;
            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                {
                    int idx = y * stride + x * 3;
                    float b = pixels[idx] / 255f;
                    float g = pixels[idx + 1] / 255f;
                    float r = pixels[idx + 2] / 255f;
                    float max = Math.Max(r, Math.Max(g, b));
                    float min = Math.Min(r, Math.Min(g, b));
                    float delta = max - min;

                    // 跳過低飽和度像素（灰/黑/白背景）
                    float sat = max == 0 ? 0 : delta / max;
                    if (sat < 0.15f || max < 0.1f) continue;

                    float hue = 0;
                    if (delta > 0)
                    {
                        if (max == r) hue = 60 * (((g - b) / delta) % 6);
                        else if (max == g) hue = 60 * (((b - r) / delta) + 2);
                        else hue = 60 * (((r - g) / delta) + 4);
                    }
                    if (hue < 0) hue += 360;

                    int bin = Math.Min(bins - 1, (int)(hue / (360.0f / bins)));
                    hist[bin]++;
                    validCount++;
                }

            if (validCount > 0)
                for (int i = 0; i < bins; i++)
                    hist[i] /= validCount;

            return hist;
        }

        /// <summary>
        /// 直方圖餘弦相似度（0=完全不同，1=完全相同）
        /// </summary>
        private static double HistogramCorrelation(float[] h1, float[] h2)
        {
            if (h1.Length != h2.Length) return 0;
            double dot = 0, norm1 = 0, norm2 = 0;
            for (int i = 0; i < h1.Length; i++)
            {
                dot += h1[i] * h2[i];
                norm1 += (double)h1[i] * h1[i];
                norm2 += (double)h2[i] * h2[i];
            }
            double denom = Math.Sqrt(norm1) * Math.Sqrt(norm2);
            return denom < 1e-10 ? 0 : dot / denom;
        }

        /// <summary>
        /// 計算灰階 Sobel 邊緣圖（float 陣列，值 0~255）
        /// </summary>
        private static float[] ComputeEdgeMap(Bitmap bmp)
        {
            int w = bmp.Width, h = bmp.Height;
            var data = bmp.LockBits(new Rectangle(0, 0, w, h),
                ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            int stride = data.Stride;
            byte[] pixels = new byte[stride * h];
            Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
            bmp.UnlockBits(data);

            // 灰階
            float[] gray = new float[w * h];
            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                {
                    int idx = y * stride + x * 3;
                    gray[y * w + x] = pixels[idx] * 0.114f + pixels[idx + 1] * 0.587f + pixels[idx + 2] * 0.299f;
                }

            // Sobel 邊緣偵測（使用近似值避免 sqrt，加速約 3x）
            float[] edge = new float[w * h];
            for (int y = 1; y < h - 1; y++)
                for (int x = 1; x < w - 1; x++)
                {
                    float gx = -gray[(y - 1) * w + (x - 1)] + gray[(y - 1) * w + (x + 1)]
                              - 2 * gray[y * w + (x - 1)] + 2 * gray[y * w + (x + 1)]
                              - gray[(y + 1) * w + (x - 1)] + gray[(y + 1) * w + (x + 1)];
                    float gy = -gray[(y - 1) * w + (x - 1)] - 2 * gray[(y - 1) * w + x] - gray[(y - 1) * w + (x + 1)]
                              + gray[(y + 1) * w + (x - 1)] + 2 * gray[(y + 1) * w + x] + gray[(y + 1) * w + (x + 1)];
                    // |gx| + |gy| 近似 sqrt(gx²+gy²)，對 NCC 比較夠用
                    float mag = Math.Abs(gx) + Math.Abs(gy);
                    edge[y * w + x] = Math.Min(mag, 510f);
                }

            return edge;
        }


        /// <summary>
        /// ★ 診斷快照：截圖儲存到應用程式根目錄
        /// </summary>
        private (ClimbingScanResult result, string savedPath) SaveRoiDebugSnapshot(string label = "")
        {
            var result = ScanClimbingRoi();

            Directory.CreateDirectory(RoiDebugDir);

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            string stateTag  = result.IsClimbing ? "CLIMBING" : "NORMAL";
            string labelPart = string.IsNullOrEmpty(label) ? "" : $"_{label}";
            string baseName  = $"ROI_{stateTag}{labelPart}_{timestamp}";

            string savedPath = "";

            // 儲存匹配位置截圖
            if (result.CaptureBitmap != null)
            {
                string rawPath = Path.Combine(RoiDebugDir, baseName + ".png");
                result.CaptureBitmap.Save(rawPath, ImageFormat.Png);
                savedPath = rawPath;
            }

            return (result, savedPath);
        }

        #endregion

        /// <summary>
        /// 檢查並觸發自定義按鍵
        /// </summary>
        /// <returns>返回需要暫停的總時間（秒）</returns>
        private double CheckAndTriggerCustomKeys(double currentScriptTime)
        {
            // 如果自定義按鍵被暫停，直接返回
            if (customKeysPaused)
                return 0;

            double totalPauseTime = 0;

            for (int i = 0; i < customKeySlots.Length; i++)
            {
                var slot = customKeySlots[i];
                if (slot.ShouldTrigger(currentScriptTime))
                {
                    // 標記為已觸發（先標記避免重複觸發）
                    slot.MarkTriggered();
                    statistics.RecordCustomKeyTrigger(i);

                    // 計算暫停時間
                    double pauseTime = slot.PauseScriptEnabled ? slot.PauseScriptSeconds : 0;
                    double preDelay = slot.PreDelaySeconds;

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (pauseTime > 0)
                        {
                            AddLog($"⏸️ 腳本暫停 {pauseTime} 秒...");
                        }
                    }));

                    // 1. 先暫停腳本（如果啟用）
                    if (pauseTime > 0)
                    {
                        Thread.Sleep((int)(pauseTime * 1000));
                        totalPauseTime += pauseTime;
                    }

                    // ★ 技能硬直鎖：鎖定期間位置修正器暫停移動
                    IsAnimationLocked = true;
                    try
                    {
                        // 2. 發送按鍵（按下和放開）
                        SendCustomKey(slot.KeyCode, slot.Modifiers);

                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            AddLog($"⚡ 自定義按鍵 #{slot.SlotNumber}: {GetKeyDisplayName(slot.KeyCode)}");
                        }));

                        // 3. 按鍵後延遲（等待技能施放完成）
                        if (preDelay > 0)
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                AddLog($"⏳ 延遲 {preDelay} 秒後繼續...");
                            }));
                            Thread.Sleep((int)(preDelay * 1000));
                            totalPauseTime += preDelay;
                        }
                    }
                    finally
                    {
                        IsAnimationLocked = false;
                    }
                }
            }

            return totalPauseTime;
        }

        /// <summary>
        /// 發送自定義按鍵（按下後立即放開，單鍵）
        /// ★ 與主播放相同路由：方向鍵用 ATT，非方向鍵用純 PostMessage（不洩漏到前景）
        /// </summary>
        private void SendCustomKey(Keys key, Keys modifiers = Keys.None)
        {
            const int KEY_HOLD_MS = 60;         // 按鍵持續時間（毫秒）

            try
            {
                if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
                {
                    // 背景模式：根據按鍵類型選擇發送方式（避免 ATT 洩漏到前景）
                    if (IsArrowKey(key))
                    {
                        // 方向鍵：使用與主播放相同的模式（遊戲用 GetKeyState 輪詢，需要 ATT）
                        SendArrowKeyWithMode(targetWindowHandle, key, true);
                        Thread.Sleep(KEY_HOLD_MS);
                        SendArrowKeyWithMode(targetWindowHandle, key, false);
                    }
                    else if (IsAltKey(key))
                    {
                        SendAltKeyToWindow(targetWindowHandle, key, true);
                        Thread.Sleep(KEY_HOLD_MS);
                        SendAltKeyToWindow(targetWindowHandle, key, false);
                    }
                    else
                    {
                        // 非方向鍵（ZXC 技能、藥水等）：純 PostMessage，不洩漏到前景
                        SendKeyWithPostMessageOnly(targetWindowHandle, key, true);
                        Thread.Sleep(KEY_HOLD_MS);
                        SendKeyWithPostMessageOnly(targetWindowHandle, key, false);
                    }
                }
                else
                {
                    // 前景模式
                    SendKeyForeground(key, true);
                    Thread.Sleep(KEY_HOLD_MS);
                    SendKeyForeground(key, false);
                }
            }
            catch (Exception ex)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    AddLog($"自定義按鍵發送失敗: {ex.Message}");
                }));
            }
        }

        /// <summary>
        /// 使用 PostMessage 發送按鍵（非阻塞）
        /// </summary>
        private void SendKeyWithPostMessage(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            byte scanCode = GetScanCode(key);
            uint lParamValue;
            if (isKeyDown)
            {
                lParamValue = 1u | ((uint)scanCode << 16);
                if (IsExtendedKey(key)) lParamValue |= (1u << 24);
            }
            else
            {
                lParamValue = 1u | ((uint)scanCode << 16) | (1u << 30) | (1u << 31);
                if (IsExtendedKey(key)) lParamValue |= (1u << 24);
            }
            IntPtr lParam = (IntPtr)lParamValue;
            uint msg = isKeyDown ? WM_KEYDOWN : WM_KEYUP;
            PostMessage(hWnd, msg, (IntPtr)key, lParam);
        }

        /// <summary>
        /// 使用 SendMessage 發送按鍵（阻塞等待處理完成）
        /// </summary>
        private void SendKeyWithSendMessage(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            byte scanCode = GetScanCode(key);
            uint lParamValue;
            if (isKeyDown)
            {
                lParamValue = 1u | ((uint)scanCode << 16);
                if (IsExtendedKey(key)) lParamValue |= (1u << 24);
            }
            else
            {
                lParamValue = 1u | ((uint)scanCode << 16) | (1u << 30) | (1u << 31);
                if (IsExtendedKey(key)) lParamValue |= (1u << 24);
            }
            IntPtr lParam = (IntPtr)lParamValue;
            uint msg = isKeyDown ? WM_KEYDOWN : WM_KEYUP;
            SendMessage(hWnd, msg, (IntPtr)key, lParam);
        }

        private void SendKeyEvent(MacroEvent evt)
        {
            try
            {
                bool isDown = evt.EventType == "down";

                // Check if we have a valid target window for background sending
                if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
                {
                    // 對於 Alt 鍵，使用特殊的發送方式
                    if (IsAltKey(evt.KeyCode))
                    {
                        SendAltKeyToWindow(targetWindowHandle, evt.KeyCode, isDown);
                    }
                    // 方向鍵：ATT + PM 兩種訊號（必須這樣才能走路）
                    else if (IsArrowKey(evt.KeyCode))
                    {
                        SendArrowKeyWithMode(targetWindowHandle, evt.KeyCode, isDown);
                    }
                    // 英數鍵：純 PostMessage（不攔截、不洩漏）
                    else if (IsAlphaNumericKey(evt.KeyCode))
                    {
                        SendKeyWithPostMessageOnly(targetWindowHandle, evt.KeyCode, isDown);
                    }
                    // 導航鍵 (Delete, Insert, Home, End, PageUp, PageDown)：純 PM
                    else if (IsNavigationKey(evt.KeyCode))
                    {
                        SendKeyWithPostMessageOnly(targetWindowHandle, evt.KeyCode, isDown);
                    }
                    // 功能鍵 (F1-F12)：純 PM
                    else if (IsFunctionKey(evt.KeyCode))
                    {
                        SendKeyWithPostMessageOnly(targetWindowHandle, evt.KeyCode, isDown);
                    }
                    // 其他延伸鍵：根據模式決定
                    else if (IsExtendedKey(evt.KeyCode))
                    {
                        if (currentArrowKeyMode == ArrowKeyMode.SendToChild)
                        {
                            // S2C 模式下也用 PM
                            SendKeyWithPostMessageOnly(targetWindowHandle, evt.KeyCode, isDown);
                        }
                        else
                        {
                            SendKeyWithThreadAttach(targetWindowHandle, evt.KeyCode, isDown);
                        }
                    }
                    else
                    {
                        // 其他一般按鍵：純 PM
                        SendKeyWithPostMessageOnly(targetWindowHandle, evt.KeyCode, isDown);
                    }
                    Debug.WriteLine($"背景: {evt.KeyCode} ({evt.EventType})");
                }
                else
                {
                    // Foreground key sending using SendInput
                    SendKeyForeground(evt.KeyCode, isDown);
                    Debug.WriteLine($"前景: {evt.KeyCode} ({evt.EventType})");
                }

                if (evt.EventType == "down")
                {
                    keySender.PressedKeys.Add(evt.KeyCode);
                }
                else if (evt.EventType == "up")
                {
                    keySender.PressedKeys.Remove(evt.KeyCode);
                }
            }
            catch (Exception ex)
            {
                AddLog($"按鍵發送失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 檢查是否為導航鍵（Delete, Insert, Home, End, PageUp, PageDown）
        /// </summary>

        /// <summary>
        /// 供位置修正器使用的按鍵發送方法 — 與主播放完全相同的路由邏輯
        /// ★ 方向鍵：使用 SendArrowKeyWithMode（跟隨使用者選擇的模式）
        /// ★ 非方向鍵：純 PostMessage（不透過 ATT 洩漏到前景）
        /// </summary>
        private void SendKeyForCorrection(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            try
            {
                if (hWnd != IntPtr.Zero && IsWindow(hWnd))
                {
                    if (IsAltKey(key))
                    {
                        SendAltKeyToWindow(hWnd, key, isKeyDown);
                    }
                    else if (IsArrowKey(key))
                    {
                        // 方向鍵：使用與主播放相同的模式（遊戲用 GetKeyState 輪詢，需要 ATT）
                        SendArrowKeyWithMode(hWnd, key, isKeyDown);
                    }
                    else
                    {
                        // 非方向鍵：純 PostMessage（不洩漏到前景應用程式）
                        SendKeyWithPostMessageOnly(hWnd, key, isKeyDown);
                    }
                }
                else
                {
                    SendKeyForeground(key, isKeyDown);
                }
            }
            catch { }
        }

        private bool IsNavigationKey(Keys key)
        {
            return key == Keys.Delete || key == Keys.Insert ||
                   key == Keys.Home || key == Keys.End ||
                   key == Keys.PageUp || key == Keys.PageDown;
        }

        /// <summary>
        /// 檢查是否為功能鍵 (F1-F12)
        /// </summary>
        private bool IsFunctionKey(Keys key)
        {
            return key >= Keys.F1 && key <= Keys.F12;
        }

        /// <summary>
        /// 純 PostMessage 發送按鍵（不洩漏到前景）
        /// </summary>
        private void SendKeyWithPostMessageOnly(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            uint scanCode = MapVirtualKey((uint)key, MAPVK_VK_TO_VSC);
            bool isExtended = IsExtendedKey(key);

            // 建構 lParam
            uint lParamValue;
            if (isKeyDown)
            {
                lParamValue = 1u | (scanCode << 16);
                if (isExtended) lParamValue |= (1u << 24);
            }
            else
            {
                lParamValue = 1u | (scanCode << 16) | (1u << 30) | (1u << 31);
                if (isExtended) lParamValue |= (1u << 24);
            }

            uint msg = isKeyDown ? WM_KEYDOWN : WM_KEYUP;
            bool success = PostMessage(hWnd, msg, (IntPtr)key, (IntPtr)lParamValue);
            
            // 如果 PostMessage 失敗，嘗試 SendMessage
            if (!success)
            {
                SendMessage(hWnd, msg, (IntPtr)key, (IntPtr)lParamValue);
            }
        }

        /// <summary>
        /// 檢查是否為方向鍵
        /// </summary>
        private bool IsArrowKey(Keys key)
        {
            return key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down;
        }

        private bool IsAlphaNumericKey(Keys key)
        {
            return (key >= Keys.A && key <= Keys.Z) || (key >= Keys.D0 && key <= Keys.D9);
        }

        /// <summary>
        /// 根據當前模式發送方向鍵
        /// </summary>
        private void SendArrowKeyWithMode(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            switch (currentArrowKeyMode)
            {
                case ArrowKeyMode.SendToChild:
                    // SendToChild 模式：ThreadAttach + PostMessage（建立在 TA 之上）
                    SendKeyWithThreadAttach(hWnd, key, isKeyDown);
                    SendArrowKeyToChildWindow(hWnd, key, isKeyDown);
                    break;

                case ArrowKeyMode.ThreadAttachWithBlocker:
                    // ThreadAttach + Blocker 模式：嘗試攔截對前景的影響
                    keyboardBlocker?.RegisterPendingKey((uint)key);
                    // 帶 Marker 讓 Blocker 能識別
                    SendKeyWithThreadAttach(hWnd, key, isKeyDown, (UIntPtr)KeyboardBlocker.MACRO_KEY_MARKER);
                    break;

                case ArrowKeyMode.SendInputWithBlock:
                    // SendInput + Blocker 模式：嘗試攔截對前景的影響
                    keyboardBlocker?.RegisterPendingKey((uint)key);
                    SendKeyWithSendInput(key, isKeyDown);
                    // 額外發送 PostMessage 以支援對話框視窗
                    SendKeyToWindow(hWnd, key, isKeyDown);
                    break;

                default:
                    SendKeyWithThreadAttach(hWnd, key, isKeyDown);
                    SendArrowKeyToChildWindow(hWnd, key, isKeyDown);
                    break;
            }
        }

        /// <summary>
        /// 使用 SendInput 發送按鍵（純前景模式）
        /// </summary>
        private void SendKeyWithSendInput(Keys key, bool isKeyDown)
        {
            ushort vkCode = (ushort)key;
            ushort scanCode = (ushort)GetScanCode(key);
            bool isExtended = IsExtendedKey(key);

            INPUT[] inputs = new INPUT[1];
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].u.ki.wVk = vkCode;
            inputs[0].u.ki.wScan = scanCode;
            inputs[0].u.ki.dwFlags = 0;
            inputs[0].u.ki.time = 0;
            // 加入 MACRO_KEY_MARKER 讓 Blocker 能立即識別並攔截
            inputs[0].u.ki.dwExtraInfo = (IntPtr)KeyboardBlocker.MACRO_KEY_MARKER;

            if (!isKeyDown)
            {
                inputs[0].u.ki.dwFlags |= KEYEVENTF_KEYUP;
            }
            if (isExtended)
            {
                inputs[0].u.ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
            }

            uint result = SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));

            if (result == 0)
            {
                int error = Marshal.GetLastWin32Error();
                Debug.WriteLine($"[SendInput Arrow] 失敗 (錯誤碼: {error})");
                // 備援：使用 keybd_event（帶 Marker）
                uint flags = 0;
                if (!isKeyDown) flags |= KEYEVENTF_KEYUP;
                if (isExtended) flags |= KEYEVENTF_EXTENDEDKEY;
                keybd_event((byte)key, (byte)scanCode, flags, (UIntPtr)KeyboardBlocker.MACRO_KEY_MARKER);
            }
        }

        private void SendArrowKeyToChildWindow(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            List<IntPtr> childWindows = new List<IntPtr>();

            // 列舉所有子視窗
            EnumChildWindows(hWnd, (childHwnd, lParam) =>
            {
                if (IsWindowVisible(childHwnd))
                {
                    childWindows.Add(childHwnd);
                }
                return true;
            }, IntPtr.Zero);

            // 構建 lParam
            byte scanCode = GetScanCode(key);
            uint lParamValue;
            if (isKeyDown)
            {
                lParamValue = 1u | ((uint)scanCode << 16) | (1u << 24); // 延伸鍵旗標
            }
            else
            {
                lParamValue = 1u | ((uint)scanCode << 16) | (1u << 24) | (1u << 30) | (1u << 31);
            }
            IntPtr lParam = (IntPtr)lParamValue;
            uint msg = isKeyDown ? WM_KEYDOWN : WM_KEYUP;

            // 發送到所有可見的子視窗
            foreach (IntPtr childHwnd in childWindows)
            {
                PostMessage(childHwnd, msg, (IntPtr)key, lParam);
            }

            // 一定也發送到主視窗（不再 fallback 到 ThreadAttach）
            PostMessage(hWnd, msg, (IntPtr)key, lParam);

            // 額外使用 SendMessage 確保訊息被處理
            SendMessage(hWnd, msg, (IntPtr)key, lParam);
        }

        /// <summary>
        /// 檢查是否為 Alt 鍵
        /// </summary>
        private bool IsAltKey(Keys key)
        {
            return key == Keys.Alt || key == Keys.Menu || key == Keys.LMenu || key == Keys.RMenu;
        }

        /// <summary>
        /// 使用 AttachThreadInput 方法發送方向鍵到背景視窗
        /// </summary>
        private void SendKeyWithThreadAttach(IntPtr hWnd, Keys key, bool isKeyDown, UIntPtr? extraInfo = null)
        {
            uint targetThreadId = GetWindowThreadProcessId(hWnd, out uint processId);
            uint currentThreadId = GetCurrentThreadId();

            bool attached = false;
            try
            {
                if (targetThreadId != currentThreadId)
                {
                    attached = AttachThreadInput(currentThreadId, targetThreadId, true);
                }

                if (attached)
                {
                    SetFocus(hWnd);
                }

                byte vkCode = (byte)key;
                byte scanCode = GetScanCode(key);

                uint flags = 0;
                if (!isKeyDown)
                {
                    flags |= KEYEVENTF_KEYUP;
                }
                if (IsExtendedKey(key))
                {
                    flags |= KEYEVENTF_EXTENDEDKEY;
                }

                keybd_event(vkCode, scanCode, flags, extraInfo ?? UIntPtr.Zero);
            }
            finally
            {
                if (attached && targetThreadId != currentThreadId)
                {
                    AttachThreadInput(currentThreadId, targetThreadId, false);
                }
            }
        }

        /// <summary>
        /// 發送 Alt 鍵到背景視窗
        /// </summary>
        private void SendAltKeyToWindow(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            uint targetThreadId = GetWindowThreadProcessId(hWnd, out uint processId);
            uint currentThreadId = GetCurrentThreadId();

            bool attached = false;
            try
            {
                if (targetThreadId != currentThreadId)
                {
                    attached = AttachThreadInput(currentThreadId, targetThreadId, true);
                }

                if (attached)
                {
                    SetFocus(hWnd);
                }

                // Alt 鍵使用對應的虛擬鍵碼
                byte vkCode;
                if (key == Keys.LMenu)
                    vkCode = 0xA4; // VK_LMENU (左 Alt)
                else if (key == Keys.RMenu)
                    vkCode = 0xA5; // VK_RMENU (右 Alt)
                else
                    vkCode = 0x12; // VK_MENU (一般 Alt)

                byte scanCode = 0x38; // Alt 的掃描碼

                uint flags = 0;
                if (!isKeyDown)
                {
                    flags |= KEYEVENTF_KEYUP;
                }
                // 右 Alt 需要延伸鍵旗標
                if (key == Keys.RMenu)
                {
                    flags |= KEYEVENTF_EXTENDEDKEY;
                }

                keybd_event(vkCode, scanCode, flags, UIntPtr.Zero);

                // 同時也嘗試用 PostMessage 發送
                uint lParamValue;
                if (isKeyDown)
                {
                    // context code = 1 表示 Alt 鍵
                    lParamValue = 1u | ((uint)scanCode << 16) | (1u << 29);
                    if (key == Keys.RMenu) lParamValue |= (1u << 24);
                }
                else
                {
                    lParamValue = 1u | ((uint)scanCode << 16) | (1u << 29) | (1u << 30) | (1u << 31);
                    if (key == Keys.RMenu) lParamValue |= (1u << 24);
                }

                IntPtr lParam = (IntPtr)lParamValue;
                uint msg = isKeyDown ? WM_SYSKEYDOWN : WM_SYSKEYUP;
                PostMessage(hWnd, msg, (IntPtr)vkCode, lParam);

                Debug.WriteLine($"Alt 按鍵: VK=0x{vkCode:X2}, SC=0x{scanCode:X2}, 旗標=0x{flags:X}");
            }
            finally
            {
                if (attached && targetThreadId != currentThreadId)
                {
                    AttachThreadInput(currentThreadId, targetThreadId, false);
                }
            }
        }

        /// <summary>
        /// 取得按鍵的掃描碼
        /// </summary>
        private byte GetScanCode(Keys key)
        {
            return key switch
            {
                Keys.Left => 0x4B,      // 左方向鍵
                Keys.Right => 0x4D,     // 右方向鍵
                Keys.Up => 0x48,        // 上方向鍵
                Keys.Down => 0x50,      // 下方向鍵
                Keys.Insert => 0x52,    // Insert 鍵
                Keys.Delete => 0x53,    // Delete 鍵
                Keys.Home => 0x47,      // Home 鍵
                Keys.End => 0x4F,       // End 鍵
                Keys.PageUp => 0x49,    // Page Up 鍵
                Keys.PageDown => 0x51,  // Page Down 鍵
                Keys.Menu => 0x38,      // Alt 鍵
                Keys.LMenu => 0x38,     // 左 Alt 鍵
                Keys.RMenu => 0x38,     // 右 Alt 鍵
                _ => (byte)MapVirtualKey((uint)key, MAPVK_VK_TO_VSC)
            };
        }

        /// <summary>
        /// 前景模式發送按鍵（使用 SendInput API - 更可靠）
        /// </summary>
        private void SendKeyForeground(Keys key, bool isKeyDown)
        {
            // 使用 SendInput（更現代、更可靠的方式）
            ushort vkCode = (ushort)key;
            ushort scanCode = (ushort)GetScanCode(key);
            bool isExtended = IsExtendedKey(key);

            INPUT[] inputs = new INPUT[1];
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].u.ki.wVk = vkCode;
            inputs[0].u.ki.wScan = scanCode;
            inputs[0].u.ki.dwFlags = 0;
            inputs[0].u.ki.time = 0;
            inputs[0].u.ki.dwExtraInfo = IntPtr.Zero;

            if (!isKeyDown)
            {
                inputs[0].u.ki.dwFlags |= KEYEVENTF_KEYUP;
            }
            if (isExtended)
            {
                inputs[0].u.ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
            }

            uint result = SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));

            // 如果 SendInput 失敗，使用 keybd_event 作為備援
            if (result == 0)
            {
                int error = Marshal.GetLastWin32Error();
                Debug.WriteLine($"[SendInput FG] 失敗 (錯誤碼: {error})，使用 keybd_event 備援");

                // 備援：使用 keybd_event
                uint flags = 0;
                if (!isKeyDown) flags |= KEYEVENTF_KEYUP;
                if (isExtended) flags |= KEYEVENTF_EXTENDEDKEY;
                keybd_event((byte)key, (byte)scanCode, flags, UIntPtr.Zero);
            }
        }

        /// <summary>
        /// 檢查是否為延伸鍵（方向鍵、Insert、Delete 等）
        /// </summary>
        private bool IsExtendedKey(Keys key)
        {
            return key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down ||
                   key == Keys.Insert || key == Keys.Delete || key == Keys.Home || key == Keys.End ||
                   key == Keys.PageUp || key == Keys.PageDown || key == Keys.NumLock || key == Keys.PrintScreen ||
                   key == Keys.RMenu || key == Keys.RControlKey || key == Keys.RShiftKey;
        }

        /// <summary>
        /// 背景模式發送按鍵（使用 PostMessage）
        /// </summary>
        private void SendKeyToWindow(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            uint scanCode;
            if (key == Keys.Left) scanCode = 0x4B;
            else if (key == Keys.Right) scanCode = 0x4D;
            else if (key == Keys.Up) scanCode = 0x48;
            else if (key == Keys.Down) scanCode = 0x50;
            else scanCode = MapVirtualKey((uint)key, MAPVK_VK_TO_VSC);

            bool isExtendedKey = IsExtendedKey(key);
            bool isAltKey = IsAltKey(key);

            // 建構 lParam 參數
            // Bits 0-15: 重複次數 (1)
            // Bits 16-23: 掃描碼
            // Bit 24: 延伸鍵旗標
            // Bit 29: 內容代碼 (Alt 鍵為 1)
            // Bit 30: 前一個鍵狀態
            // Bit 31: 轉換狀態 (0 = 按下, 1 = 放開)
            uint lParamValue;
            if (isKeyDown)
            {
                lParamValue = 1u | (scanCode << 16);
                if (isExtendedKey) lParamValue |= (1u << 24);
                if (isAltKey) lParamValue |= (1u << 29);
            }
            else
            {
                lParamValue = 1u | (scanCode << 16) | (1u << 30) | (1u << 31);
                if (isExtendedKey) lParamValue |= (1u << 24);
                if (isAltKey) lParamValue |= (1u << 29);
            }

            IntPtr lParam = (IntPtr)lParamValue;

            uint msg;
            if (isAltKey)
            {
                msg = isKeyDown ? WM_SYSKEYDOWN : WM_SYSKEYUP;
            }
            else
            {
                msg = isKeyDown ? WM_KEYDOWN : WM_KEYUP;
            }

            bool success = PostMessage(hWnd, msg, (IntPtr)key, lParam);

            if (!success)
            {
                SendMessage(hWnd, msg, (IntPtr)key, lParam);
            }
        }

        /// <summary>
        /// 發送文字到背景視窗（使用 WM_CHAR 逐字發送）
        /// </summary>
        private void SendTextToWindow(IntPtr hWnd, string text)
        {
            foreach (char c in text)
            {
                PostMessage(hWnd, WM_CHAR, (IntPtr)c, IntPtr.Zero);
                Thread.Sleep(30);
            }
        }

        /// <summary>
        /// 發送文字到前景（使用 SendInput 逐字發送）
        /// </summary>
        private void SendTextForeground(string text)
        {
            foreach (char c in text)
            {
                INPUT[] inputs = new INPUT[2];

                // Key down (UNICODE)
                inputs[0].type = INPUT_KEYBOARD;
                inputs[0].u.ki.wVk = 0;
                inputs[0].u.ki.wScan = (ushort)c;
                inputs[0].u.ki.dwFlags = 0x0004; // KEYEVENTF_UNICODE
                inputs[0].u.ki.time = 0;
                inputs[0].u.ki.dwExtraInfo = IntPtr.Zero;

                // Key up (UNICODE)
                inputs[1].type = INPUT_KEYBOARD;
                inputs[1].u.ki.wVk = 0;
                inputs[1].u.ki.wScan = (ushort)c;
                inputs[1].u.ki.dwFlags = 0x0004 | KEYEVENTF_KEYUP;
                inputs[1].u.ki.time = 0;
                inputs[1].u.ki.dwExtraInfo = IntPtr.Zero;

                SendInput(2, inputs, Marshal.SizeOf(typeof(INPUT)));
                Thread.Sleep(30);
            }
        }

        /// <summary>
        /// 尋找遊戲渲染子視窗（MapleStoryClass 等）
        /// 如果主視窗本身就是 MapleStoryClass，則返回主視窗
        /// </summary>
        private IntPtr FindGameRenderWindow(IntPtr hWndParent)
        {
            // 1. 檢查主視窗本身
            StringBuilder className = new StringBuilder(256);
            GetClassName(hWndParent, className, 256);
            string parentClass = className.ToString();
            if (parentClass.IndexOf("MapleStory", StringComparison.OrdinalIgnoreCase) >= 0)
                return hWndParent;

            // 2. 列舉子視窗尋找 MapleStoryClass
            IntPtr found = IntPtr.Zero;
            EnumChildWindows(hWndParent, (childHwnd, lParam) =>
            {
                StringBuilder childClassName = new StringBuilder(256);
                GetClassName(childHwnd, childClassName, 256);
                if (childClassName.ToString().IndexOf("MapleStory", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    found = childHwnd;
                    return false; // 找到了，停止列舉
                }
                return true;
            }, IntPtr.Zero);

            // 3. 嘗試 FindWindowEx
            if (found == IntPtr.Zero)
                found = FindWindowEx(hWndParent, IntPtr.Zero, "MapleStoryClass", null);

            return found != IntPtr.Zero ? found : hWndParent;
        }

        /// <summary>
        /// 對話框模板圖片儲存目錄
        /// </summary>
        private static readonly string DialogTemplateDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MapleMacro", "templates");

        /// <summary>
        /// 截取遊戲視窗畫面（背景模式用 PrintWindow，前景用 null）
        /// </summary>
        private Bitmap? CaptureGameWindowForDetection()
        {
            if (targetWindowHandle == IntPtr.Zero || !IsWindow(targetWindowHandle))
                return null;

            if (minimapTracker == null)
            {
                minimapTracker = new MinimapTracker();
                minimapTracker.AttachToWindow(targetWindowHandle);
            }
            else if (minimapTracker.TargetWindow != targetWindowHandle)
            {
                minimapTracker.AttachToWindow(targetWindowHandle);
            }

            return minimapTracker.CaptureFullWindow();
        }

        /// <summary>
        /// 在大圖中搜尋模板圖片，返回匹配度最高的位置和分數
        /// 使用 Sum of Absolute Differences (SAD) 演算法
        /// </summary>
        /// <param name="source">來源大圖</param>
        /// <param name="template">要搜尋的模板小圖</param>
        /// <param name="threshold">匹配閾值 (0.0~1.0)</param>
        /// <returns>(匹配成功, 匹配度分數, 匹配位置X, 匹配位置Y)</returns>
        private static (bool found, double score, int x, int y) FindTemplateInImage(
            Bitmap source, Bitmap template, double threshold = 0.85)
        {
            if (source == null || template == null) return (false, 0, 0, 0);
            if (template.Width > source.Width || template.Height > source.Height)
                return (false, 0, 0, 0);

            int srcW = source.Width, srcH = source.Height;
            int tmpW = template.Width, tmpH = template.Height;

            // 使用 LockBits 加速像素存取
            var srcData = source.LockBits(new Rectangle(0, 0, srcW, srcH),
                ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            var tmpData = template.LockBits(new Rectangle(0, 0, tmpW, tmpH),
                ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

            int srcStride = srcData.Stride;
            int tmpStride = tmpData.Stride;
            int srcBytes = srcStride * srcH;
            int tmpBytes = tmpStride * tmpH;

            byte[] srcPixels = new byte[srcBytes];
            byte[] tmpPixels = new byte[tmpBytes];
            Marshal.Copy(srcData.Scan0, srcPixels, 0, srcBytes);
            Marshal.Copy(tmpData.Scan0, tmpPixels, 0, tmpBytes);

            source.UnlockBits(srcData);
            template.UnlockBits(tmpData);

            double bestScore = 0;
            int bestX = 0, bestY = 0;

            // 模板的像素總數 * 3（RGB 通道）
            long maxDiff = (long)tmpW * tmpH * 3 * 255;

            // 每隔 2 像素搜尋以加速（粗搜），找到候選後精確驗證
            int step = 2;
            int searchW = srcW - tmpW;
            int searchH = srcH - tmpH;

            for (int sy = 0; sy <= searchH; sy += step)
            {
                for (int sx = 0; sx <= searchW; sx += step)
                {
                    // 快速預檢：只比對四個角落 + 中心的像素
                    long quickDiff = 0;
                    int[][] checkPoints = new int[][] {
                        new[] {0, 0}, new[] {tmpW-1, 0}, new[] {0, tmpH-1}, new[] {tmpW-1, tmpH-1},
                        new[] {tmpW/2, tmpH/2}
                    };

                    bool quickFail = false;
                    foreach (var pt in checkPoints)
                    {
                        int srcIdx = (sy + pt[1]) * srcStride + (sx + pt[0]) * 3;
                        int tmpIdx = pt[1] * tmpStride + pt[0] * 3;
                        if (srcIdx + 2 >= srcBytes || tmpIdx + 2 >= tmpBytes) { quickFail = true; break; }

                        int db = Math.Abs(srcPixels[srcIdx] - tmpPixels[tmpIdx]);
                        int dg = Math.Abs(srcPixels[srcIdx + 1] - tmpPixels[tmpIdx + 1]);
                        int dr = Math.Abs(srcPixels[srcIdx + 2] - tmpPixels[tmpIdx + 2]);
                        if (db + dg + dr > 80) { quickFail = true; break; } // 單像素差距過大
                    }
                    if (quickFail) continue;

                    // 完整比對
                    long totalDiff = 0;
                    bool earlyExit = false;
                    long earlyThreshold = (long)(maxDiff * (1.0 - threshold));

                    for (int ty = 0; ty < tmpH && !earlyExit; ty++)
                    {
                        for (int tx = 0; tx < tmpW; tx++)
                        {
                            int srcIdx = (sy + ty) * srcStride + (sx + tx) * 3;
                            int tmpIdx = ty * tmpStride + tx * 3;

                            totalDiff += Math.Abs(srcPixels[srcIdx] - tmpPixels[tmpIdx]);
                            totalDiff += Math.Abs(srcPixels[srcIdx + 1] - tmpPixels[tmpIdx + 1]);
                            totalDiff += Math.Abs(srcPixels[srcIdx + 2] - tmpPixels[tmpIdx + 2]);

                            if (totalDiff > earlyThreshold) { earlyExit = true; break; }
                        }
                    }

                    if (!earlyExit)
                    {
                        double score = 1.0 - (double)totalDiff / maxDiff;
                        if (score > bestScore)
                        {
                            bestScore = score;
                            bestX = sx;
                            bestY = sy;
                        }
                    }
                }
            }

            return (bestScore >= threshold, bestScore, bestX, bestY);
        }

        /// <summary>
        /// 發送滑鼠點擊到視窗（Client Area 座標）
        /// ★ 雙重策略：PM 到 MapleStoryClass + 前景 SendInput 補強
        /// </summary>
        private void SendMouseClickToWindow(IntPtr hWnd, int clientX, int clientY)
        {
            // ★ 找到遊戲渲染子視窗（MapleStoryClass）
            IntPtr gameWnd = FindGameRenderWindow(hWnd);

            StringBuilder cn = new StringBuilder(256);
            GetClassName(gameWnd, cn, 256);
            string targetClass = cn.ToString();
            bool isGameClass = targetClass.IndexOf("MapleStory", StringComparison.OrdinalIgnoreCase) >= 0;

            Dispatcher.BeginInvoke(new Action(() =>
                AddLog($"🖱️ 目標視窗: [{targetClass}] hWnd=0x{gameWnd:X} isGame={isGameClass}")));

            // ★ 方式1: PostMessage 到 MapleStoryClass（背景點擊）
            IntPtr lParamCoord = (IntPtr)((clientY << 16) | (clientX & 0xFFFF));

            // 先送 WM_MOUSEMOVE 讓遊戲更新內部滑鼠位置狀態
            PostMessage(gameWnd, WM_MOUSEMOVE, IntPtr.Zero, lParamCoord);
            Thread.Sleep(50);

            // 送 WM_LBUTTONDOWN + WM_LBUTTONUP
            PostMessage(gameWnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, lParamCoord);
            Thread.Sleep(100);
            PostMessage(gameWnd, WM_LBUTTONUP, IntPtr.Zero, lParamCoord);

            Dispatcher.BeginInvoke(new Action(() =>
                AddLog($"🖱️ PM 點擊已送出 → [{targetClass}] ({clientX},{clientY})")));

            // ★ 方式2: 前景 SendInput 補強（確保至少一種方式有效）
            Thread.Sleep(150);

            POINT pt = new POINT { X = clientX, Y = clientY };
            ClientToScreen(gameWnd, ref pt);

            SetForegroundWindow(hWnd);
            Thread.Sleep(150);

            SetCursorPos(pt.X, pt.Y);
            Thread.Sleep(50);

            INPUT[] inputs = new INPUT[2];
            inputs[0].type = INPUT_MOUSE;
            inputs[0].u.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
            inputs[0].u.mi.time = 0;
            inputs[0].u.mi.dwExtraInfo = IntPtr.Zero;

            inputs[1].type = INPUT_MOUSE;
            inputs[1].u.mi.dwFlags = MOUSEEVENTF_LEFTUP;
            inputs[1].u.mi.time = 0;
            inputs[1].u.mi.dwExtraInfo = IntPtr.Zero;

            uint result = SendInput(2, inputs, Marshal.SizeOf(typeof(INPUT)));

            Dispatcher.BeginInvoke(new Action(() =>
                AddLog($"🖱️ 前景點擊已送出 → screen({pt.X},{pt.Y}) result={result}")));
        }

        /// <summary>
        /// 執行回程序列：停止腳本 → 冷卻 2000ms → [前置鍵] → 滑鼠點擊座標 → 偵測對話框 → [確認鍵] → 延遲 → [坐下鍵]
        /// ★ 每個按鍵只按一次，不重複確認
        /// ★ 如果啟用對話框偵測，會重複點擊直到偵測到對話框才按確認鍵
        /// </summary>
        private void ExecuteReturnToTown(ScheduleTask task)
        {
            try
            {
                Keys preKey = (Keys)task.ReturnPreKeyCode;
                Keys confirmKey = (Keys)task.ReturnConfirmKeyCode;
                Keys sitKey = (Keys)task.SitDownKeyCode;

                string preKeyName = preKey != Keys.None ? GetKeyDisplayName(preKey) : "";
                string confirmKeyName = confirmKey != Keys.None ? GetKeyDisplayName(confirmKey) : "";
                string sitKeyName = sitKey != Keys.None ? GetKeyDisplayName(sitKey) : "";

                // 組裝序列描述
                string seqDesc = $"冷卻 2000ms";
                if (preKey != Keys.None) seqDesc += $" → {preKeyName}";
                seqDesc += $" → 點擊({task.ReturnClickX},{task.ReturnClickY})";
                if (task.DialogDetectionEnabled) seqDesc += " → 偵測對話框";
                if (confirmKey != Keys.None) seqDesc += $" → {confirmKeyName}";
                if (sitKey != Keys.None) seqDesc += $" → 延遲{task.SitDownDelaySeconds}s → {sitKeyName}";

                Dispatcher.BeginInvoke(new Action(() => AddLog($"🏠 開始回程序列：{seqDesc}")));

                // 1. 冷卻延遲 2000ms（確保腳本完全停止）
                Thread.Sleep(2000);

                bool isBackground = targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle);

                if (!isBackground)
                {
                    Dispatcher.BeginInvoke(new Action(() => AddLog("⚠️ 前景模式不支援背景滑鼠點擊，跳過")));
                    return;
                }

                // 2. 前置按鍵（只按一次，例如 ESC 關閉選單）
                if (preKey != Keys.None)
                {
                    SendCustomKey(preKey);
                    Dispatcher.BeginInvoke(new Action(() => AddLog($"📨 已送出前置鍵：{preKeyName}")));
                    Thread.Sleep(500); // 等待選單關閉
                }

                // 3. 嘗試點擊並偵測對話框
                bool dialogFound = false;

                if (task.DialogDetectionEnabled && !string.IsNullOrEmpty(task.DialogTemplatePath) && File.Exists(task.DialogTemplatePath))
                {
                    // ★ 對話框偵測模式：點擊 → 截圖比對 → 迴圈直到偵測到或超時
                    Bitmap? templateBmp = null;
                    try
                    {
                        templateBmp = new Bitmap(task.DialogTemplatePath);
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.BeginInvoke(new Action(() => AddLog($"⚠️ 載入模板圖片失敗: {ex.Message}，改用直接點擊模式")));
                    }

                    if (templateBmp != null)
                    {
                        int maxRetries = Math.Max(1, task.DialogDetectionMaxRetries);
                        double matchThreshold = task.DialogDetectionThreshold;

                        for (int attempt = 1; attempt <= maxRetries && !dialogFound; attempt++)
                        {
                            int curAttempt = attempt;
                            Dispatcher.BeginInvoke(new Action(() => AddLog($"🔍 第 {curAttempt}/{maxRetries} 次嘗試：點擊 → 偵測對話框...")));

                            // 點擊 NPC
                            SendMouseClickToWindow(targetWindowHandle, task.ReturnClickX, task.ReturnClickY);
                            Thread.Sleep(800); // 等待對話框出現

                            // 截取遊戲畫面
                            using (var screenshot = CaptureGameWindowForDetection())
                            {
                                if (screenshot != null)
                                {
                                    var (found, score, mx, my) = FindTemplateInImage(screenshot, templateBmp, matchThreshold);
                                    double scoreRound = Math.Round(score * 100, 1);

                                    if (found)
                                    {
                                        dialogFound = true;
                                        Dispatcher.BeginInvoke(new Action(() => AddLog($"✅ 對話框偵測成功！匹配度={scoreRound}% 位置=({mx},{my})")));
                                    }
                                    else
                                    {
                                        Dispatcher.BeginInvoke(new Action(() => AddLog($"🔍 未偵測到對話框 (最佳匹配={scoreRound}%)，等待重試...")));
                                        Thread.Sleep(1000); // 等待一秒後重試
                                    }
                                }
                                else
                                {
                                    Dispatcher.BeginInvoke(new Action(() => AddLog($"⚠️ 截圖失敗，等待重試...")));
                                    Thread.Sleep(500);
                                }
                            }
                        }

                        templateBmp.Dispose();

                        if (!dialogFound)
                        {
                            string forceMsg = confirmKey != Keys.None
                                ? $"⚠️ {maxRetries} 次嘗試後仍未偵測到對話框，強制按 {confirmKeyName}"
                                : $"⚠️ {maxRetries} 次嘗試後仍未偵測到對話框";
                            Dispatcher.BeginInvoke(new Action(() => AddLog(forceMsg)));
                        }
                    }
                    else
                    {
                        // 模板載入失敗，退回到直接點擊
                        SendMouseClickToWindow(targetWindowHandle, task.ReturnClickX, task.ReturnClickY);
                        Dispatcher.BeginInvoke(new Action(() => AddLog($"🖱️ 已點擊座標 ({task.ReturnClickX}, {task.ReturnClickY})")));
                        Thread.Sleep(300);
                    }
                }
                else
                {
                    // ★ 直接點擊模式（無對話框偵測）
                    SendMouseClickToWindow(targetWindowHandle, task.ReturnClickX, task.ReturnClickY);
                    Dispatcher.BeginInvoke(new Action(() => AddLog($"🖱️ 已點擊座標 ({task.ReturnClickX}, {task.ReturnClickY})")));
                    Thread.Sleep(300);
                }

                // 4. 確認按鍵（只按一次）
                if (confirmKey != Keys.None)
                {
                    SendCustomKey(confirmKey);
                    Dispatcher.BeginInvoke(new Action(() => AddLog($"📨 已送出確認鍵：{confirmKeyName}")));
                }

                // 5. 等待後坐下（如設定）
                if (sitKey != Keys.None)
                {
                    int delayMs = (int)(task.SitDownDelaySeconds * 1000);
                    Dispatcher.BeginInvoke(new Action(() => AddLog($"⏳ 等待 {task.SitDownDelaySeconds} 秒後坐下...")));
                    Thread.Sleep(delayMs);

                    SendCustomKey(sitKey);
                    Dispatcher.BeginInvoke(new Action(() => AddLog($"🪑 已坐下：{sitKeyName}")));
                }

                Dispatcher.BeginInvoke(new Action(() => AddLog($"✅ 回程序列完成")));
            }
            catch (Exception ex)
            {
                Dispatcher.BeginInvoke(new Action(() => AddLog($"❌ 回程序列失敗: {ex.Message}")));
            }
        }

        /// <summary>
        /// 將文字輸入到遊戲對話框
        /// 背景模式：使用 WM_CHAR 逐字發送（ASCII 指令不受注音干擾）
        /// 前景模式：使用剪貼簿 Ctrl+V 貼上（繞過注音/IME）
        /// </summary>
        private void PasteTextToGame(IntPtr hWnd, string text, bool isBackground)
        {
            if (isBackground)
            {
                // 背景模式：WM_CHAR 逐字發送
                // @FM 等指令都是 ASCII 字元，WM_CHAR 直接送字元碼，不經過 IME 轉換
                foreach (char c in text)
                {
                    PostMessage(hWnd, WM_CHAR, (IntPtr)c, IntPtr.Zero);
                    Thread.Sleep(50);
                }
            }
            else
            {
                // 前景模式：使用剪貼簿 Ctrl+V 貼上
                string? originalClipboard = null;
                Dispatcher.Invoke(new Action(() =>
                {
                    try
                    {
                        if (Clipboard.ContainsText())
                            originalClipboard = Clipboard.GetText();
                        Clipboard.SetText(text);
                    }
                    catch { }
                }));
                Thread.Sleep(100);

                SendKeyForeground(Keys.ControlKey, true);
                Thread.Sleep(30);
                SendKeyForeground(Keys.V, true);
                Thread.Sleep(50);
                SendKeyForeground(Keys.V, false);
                Thread.Sleep(30);
                SendKeyForeground(Keys.ControlKey, false);
                Thread.Sleep(150);

                // 還原剪貼簿
                Dispatcher.Invoke(new Action(() =>
                {
                    try
                    {
                        if (originalClipboard != null)
                            Clipboard.SetText(originalClipboard);
                        else
                            Clipboard.Clear();
                    }
                    catch { }
                }));
            }
        }

        private void BtnStopPlayback_Click(object? sender, EventArgs e)
        {
            if (isPlaying)
            {
                statistics.EndSession();
            }
            // 停用 KeyboardBlocker
            if (keyboardBlocker != null)
            {
                keyboardBlocker.IsBlocking = false;
                keyboardBlocker.Uninstall();
                AddLog($"鍵盤阻擋器已停用");
            }

            ReleasePressedKeys();

            // 重置暫停狀態
            lock (pauseLock)
            {
                isPaused = false;
                System.Threading.Monitor.PulseAll(pauseLock);  // 喚醒可能暫停的線程
            }

            isPlaying = false;

            // 清除偏差修正的參考位置，下次播放時重新建立修正器
            if (positionCorrector != null)
            {
                // ★ 停止時輸出修正歷史摘要
                if (positionCorrector.HistoryCount > 0)
                {
                    AddLog($"📊 修正歷史: {positionCorrector.GetHistorySummary()}");
                }
                positionCorrector.ClearReferencePosition();
                positionCorrector.ClearHistory();
            }
            positionCorrector = null; // 重建以避免重複日誌處理器

            // ★ 清除小地圖位置記憶，避免換地圖後誤判
            minimapTracker?.ResetPositionMemory();

            lblPlaybackStatus.Text = "播放: 已停止";
            lblPlaybackStatus.Foreground = ToBrush(Color.Orange);
            AddLog("播放已停止");
            UpdatePauseButtonState();
            UpdateUI();
        }

        /// <summary>
        /// 暫停按鈕點擊事件
        /// </summary>
        private void BtnPausePlayback_Click(object? sender, EventArgs e)
        {
            if (isPlaying)
            {
                TogglePause();
            }
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (isRecording)
                BtnStopRecording_Click(this, EventArgs.Empty);
            if (isPlaying)
            {
                statistics.EndSession();
                isPlaying = false;
            }

            keyboardHook?.Uninstall();
            hotkeyHook?.Uninstall();  // 停止全局熱鍵監聽
            keyboardBlocker?.Uninstall();  // 停止鍵盤阻擋器
            keyboardBlocker?.Dispose();
            minimapTracker?.Dispose();  // 釋放小地圖追蹤器
            monitorTimer?.Stop();
            schedulerTimer?.Stop();   // 停止定時執行計時器

            // 自動儲存設定（靜默，不顯示錯誤對話框）
            try { SaveSettingsToFile(SettingsFilePath, ""); } catch { }

            Logger.Info("=== 程式關閉 ===");
            Logger.Shutdown();

            AddLog("應用程式已關閉");
        }

        /// <summary>
        /// 全局熱鍵事件處理
        /// </summary>
        private void HotkeyHook_OnKeyEvent(Keys keyCode, bool isKeyDown)
        {
            // 只處理按下事件，避免重複觸發
            if (!isKeyDown || !hotkeyEnabled)
                return;

            // ★ 錄製熱鍵：不管是否在錄製中都要處理（用來開始/停止錄製）
            if (keyCode == recordHotkey)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (isRecording)
                    {
                        AddLog($"熱鍵觸發：停止錄製 ({GetKeyDisplayName(recordHotkey)})");
                        BtnStopRecording_Click(this, EventArgs.Empty);
                    }
                    else if (!isPlaying)
                    {
                        AddLog($"熱鍵觸發：開始錄製 ({GetKeyDisplayName(recordHotkey)})");
                        BtnStartRecording_Click(this, EventArgs.Empty);
                    }
                }));
                return;
            }

            // 錄製中不處理其他熱鍵
            if (isRecording) return;

            // ★ F7 邊界設定熱鍵 — 狀態機
            if (keyCode == boundaryHotkey)
            {
                Dispatcher.BeginInvoke(new Action(() => HandleBoundaryHotkey()));
                return;
            }

            // 檢查是否為播放熱鍵
            if (keyCode == playHotkey)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (!isPlaying && recordedEvents.Count > 0)
                    {
                        AddLog($"熱鍵觸發：開始播放 ({GetKeyDisplayName(playHotkey)})");
                        BtnStartPlayback_Click(this, EventArgs.Empty);
                    }
                }));
            }
            // 檢查是否為停止熱鍵
            else if (keyCode == stopHotkey)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (isPlaying)
                    {
                        AddLog($"熱鍵觸發：停止播放 ({GetKeyDisplayName(stopHotkey)})");
                        BtnStopPlayback_Click(this, EventArgs.Empty);
                    }
                }));
            }
            // 檢查是否為暫停熱鍵
            else if (keyCode == pauseHotkey)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (isPlaying)
                    {
                        TogglePause();
                    }
                }));
            }
        }

        /// <summary>
        /// 切換暫停/繼續狀態
        /// </summary>
        private void TogglePause()
        {
            lock (pauseLock)
            {
                isPaused = !isPaused;

                if (isPaused)
                {
                    AddLog($"⏸ 播放已暫停 ({GetKeyDisplayName(pauseHotkey)})");
                    lblPlaybackStatus.Text = "狀態: 已暫停";
                    lblPlaybackStatus.Foreground = ToBrush(Color.Orange);
                }
                else
                {
                    AddLog($"▶ 播放已繼續 ({GetKeyDisplayName(pauseHotkey)})");
                    lblPlaybackStatus.Text = "狀態: 播放中";
                    lblPlaybackStatus.Foreground = ToBrush(Color.Green);

                    // 喚醒暫停中的線程
                    System.Threading.Monitor.PulseAll(pauseLock);
                }
                
                // 更新暫停按鈕狀態
                UpdatePauseButtonState();
            }
        }
        
        /// <summary>
        /// 更新暫停按鈕的外觀狀態
        /// </summary>
        private void UpdatePauseButtonState()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new Action(UpdatePauseButtonState));
                return;
            }
            
            if (isPlaying)
            {
                txtPauseHotkey.IsEnabled = true;
                if (isPaused)
                {
                    // 暫停中 - 顯示繼續
                    txtPauseHotkey.Content = "▶ 繼續";
                    txtPauseHotkey.Background = ToBrush(Color.FromArgb(0, 150, 80)); // 綠色
                }
                else
                {
                    // 播放中 - 顯示暫停
                    txtPauseHotkey.Content = "⏸ 暫停";
                    txtPauseHotkey.Background = ToBrush(Color.CornflowerBlue); // 藍色
                }
            }
            else
            {
                // 未播放 - 禁用按鈕
                txtPauseHotkey.IsEnabled = false;
                txtPauseHotkey.Content = "⏸ 暫停";
                txtPauseHotkey.Background = ToBrush(Color.FromArgb(80, 80, 85)); // 灰色
            }
            
        }

        /// <summary>
        /// 檢查暫停狀態，如果暫停則等待。
        /// ★ 同時等待 _playbackGate（修正期間阻塞）。
        /// </summary>
        private void CheckPauseState()
        {
            // ★ 嚴格阻塞：修正期間等待閘門打開（超時保護 30 秒防止死鎖）
            _playbackGate.Wait(30000);

            lock (pauseLock)
            {
                while (isPaused && isPlaying)
                {
                    // 暫停期間等待，直到被喚醒或停止
                    System.Threading.Monitor.Wait(pauseLock, 100);  // 100ms 超時，避免死鎖
                }
            }
        }





        /// <summary>
        /// 小地圖校準設定儲存路徑
        /// </summary>
        private static readonly string CalibrationFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MapleMacro",
            "minimap_calibration.json");

        /// <summary>
        /// ★ F7 邊界設定熱鍵處理 — 狀態機
        /// 第一次按：開始設定（左邊界），之後每按一次前進一步
        /// 順序：左 → 右 → 上 → 下 → 完成儲存 → 重置狀態
        /// </summary>
        private void HandleBoundaryHotkey()
        {
            if (minimapTracker == null || !minimapTracker.IsCalibrated)
            {
                AddLog("⚠️ F7 邊界設定：請先校準小地圖");
                return;
            }

            string[] stateNames = { "左邊界", "右邊界", "下邊界", "上邊界" };

            if (_boundarySetState < 0)
            {
                // 開始設定流程
                _boundarySetState = 0;
                AddLog($"🗺️ F7 邊界設定開始 — 將角色移到【{stateNames[0]}】位置，再按 F7");
                return;
            }

            // 讀取當前位置
            var (cx, cy, ok) = minimapTracker.ReadPosition();
            if (!ok)
            {
                AddLog("⚠️ F7 邊界設定：偵測失敗，請重試");
                return;
            }

            var bounds = minimapTracker.MapBounds;
            int bL = bounds.Left, bR = bounds.Right, bT = bounds.Top, bB = bounds.Bottom;

            switch (_boundarySetState)
            {
                case 0: bL = cx; AddLog($"✅ 左邊界 = {cx}"); break;
                case 1: bR = cx; AddLog($"✅ 右邊界 = {cx}"); break;
                case 2: bB = cy; AddLog($"✅ 下邊界 = {cy}"); break;
                case 3: bT = cy; AddLog($"✅ 上邊界 = {cy}"); break;
            }

            // 更新 MapBounds
            minimapTracker.MapBounds = new Rectangle(bL, bT, Math.Max(1, bR - bL), Math.Max(1, bB - bT));

            _boundarySetState++;
            if (_boundarySetState < 4)
            {
                AddLog($"🗺️ 將角色移到【{stateNames[_boundarySetState]}】位置，再按 F7");
            }
            else
            {
                // 全部設定完畢 → 儲存校準
                _boundarySetState = -1;
                minimapTracker.SaveCalibration(CalibrationFilePath);
                AddLog($"🗺️ 邊界設定完成！左={bL} 右={bR} 上={bT} 下={bB} — 已儲存");
            }
        }

        /// <summary>
        /// 開啟小地圖校準介面
        /// </summary>
        private void OpenMemoryScannerChoice()
        {
            // 直接進入校準介面（位置修正按鈕已整合在校準視窗內）
            OpenMinimapCalibration();
        }


        // HSV 轉 Color
        private static Color ColorFromHSV(double hue, double saturation, double value)
        {
            int hi = Convert.ToInt32(Math.Floor(hue / 60)) % 6;
            double f = hue / 60 - Math.Floor(hue / 60);
            value = value * 255;
            int v = Convert.ToInt32(value);
            int p = Convert.ToInt32(value * (1 - saturation));
            int q = Convert.ToInt32(value * (1 - f * saturation));
            int t = Convert.ToInt32(value * (1 - (1 - f) * saturation));
            return hi switch
            {
                0 => Color.FromArgb(255, v, t, p),
                1 => Color.FromArgb(255, q, v, p),
                2 => Color.FromArgb(255, p, v, t),
                3 => Color.FromArgb(255, p, q, v),
                4 => Color.FromArgb(255, t, p, v),
                _ => Color.FromArgb(255, v, p, q)
            };
        }

        // Color 轉 HSV
        private static void ColorToHSV(Color color, out float hue, out float saturation, out float value)
        {
            float r = color.R / 255f;
            float g = color.G / 255f;
            float b = color.B / 255f;
            float max = Math.Max(r, Math.Max(g, b));
            float min = Math.Min(r, Math.Min(g, b));
            float delta = max - min;

            hue = 0;
            if (delta > 0)
            {
                if (max == r) hue = 60 * (((g - b) / delta) % 6);
                else if (max == g) hue = 60 * (((b - r) / delta) + 2);
                else hue = 60 * (((r - g) / delta) + 4);
            }
            if (hue < 0) hue += 360;

            saturation = max == 0 ? 0 : delta / max;
            value = max;
        }

        /// <summary>
        /// 執行位置修正（偏差模式 - 比對錄製時的參考位置）
        /// </summary>
        private CorrectionResult? ExecutePositionCorrection()
        {
            if (minimapTracker == null || !minimapTracker.IsCalibrated)
                return null;

            if (positionCorrector == null)
            {
                positionCorrector = new PositionCorrector();
                positionCorrector.OnLog += (msg) =>
                {
                    try { Dispatcher.BeginInvoke(new Action(() => AddLog($"[修正] {msg}"))); } catch { }
                };
            }

            ApplyCorrectorSettings(positionCorrector);

            // ★ 偏差修正：使用參考位置
            var (refX, refY, refSet) = positionCorrector.GetReferencePosition();
            if (!refSet)
            {
                var readings = new List<(int x, int y)>();
                for (int i = 0; i < 3; i++)
                {
                    var (rx, ry, rok) = minimapTracker.ReadPosition();
                    if (rok) readings.Add((rx, ry));
                    Thread.Sleep(80);
                }

                if (readings.Count >= 2)
                {
                    readings.Sort((a, b) => a.x.CompareTo(b.x));
                    int medX = readings[readings.Count / 2].x;
                    readings.Sort((a, b) => a.y.CompareTo(b.y));
                    int medY = readings[readings.Count / 2].y;
                    positionCorrector.SetReferencePosition(medX, medY);

                    try
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                            AddLog($"📍 偏差修正：擷取參考位置 ({medX},{medY})（讀取{readings.Count}次取中位數）")));
                    } catch { }

                    return new CorrectionResult(true, $"擷取參考位置: ({medX},{medY})", medX, medY, 0, 0);
                }
                else
                {
                    try { Dispatcher.BeginInvoke(new Action(() => AddLog("⚠️ 偵測失敗，無法擷取參考位置"))); } catch { }
                    return new CorrectionResult(false, "偵測失敗，無法擷取參考位置");
                }
            }

            return positionCorrector.CorrectDeviation(minimapTracker);
        }

        /// <summary>★ 將當前設定套用到修正器</summary>
        private void ApplyCorrectorSettings(PositionCorrector corrector)
        {
            corrector.TargetWindow = targetWindowHandle;
            corrector.ExternalKeySender = SendKeyForCorrection;
            corrector.IsAnimationLockedCheck = () => IsAnimationLocked;
            corrector.IsClimbingCheck = () => DetectIsClimbing(); // ★ ROI 爬繩偵測
            corrector.Tolerance = positionCorrectionSettings.Tolerance;
            corrector.SoftToleranceMin = positionCorrectionSettings.SoftToleranceMin;
            corrector.SoftToleranceMax = positionCorrectionSettings.SoftToleranceMax;
            corrector.HorizontalTolerance = positionCorrectionSettings.HorizontalTolerance;
            corrector.VerticalTolerance = positionCorrectionSettings.VerticalTolerance;
            corrector.MaxCorrectionTimeMs = positionCorrectionSettings.MaxCorrectionTimeMs;
            corrector.MoveLeftKeys = positionCorrectionSettings.GetEffectiveLeftKeys();
            corrector.MoveRightKeys = positionCorrectionSettings.GetEffectiveRightKeys();
            corrector.MoveUpKeys = positionCorrectionSettings.GetEffectiveUpKeys();
            corrector.MoveDownKeys = positionCorrectionSettings.GetEffectiveDownKeys();
            corrector.EnableHorizontalCorrection = positionCorrectionSettings.EnableHorizontalCorrection;
            corrector.EnableVerticalCorrection = positionCorrectionSettings.EnableVerticalCorrection;
            corrector.InvertY = positionCorrectionSettings.InvertY;
            corrector.MaxCorrectionsPerLoop = positionCorrectionSettings.MaxCorrectionsPerLoop;
            corrector.MaxStepsPerCorrection = positionCorrectionSettings.MaxStepsPerCorrection;
            corrector.KeyIntervalMinMs = positionCorrectionSettings.KeyIntervalMinMs;
            corrector.KeyIntervalMaxMs = positionCorrectionSettings.KeyIntervalMaxMs;
            corrector.HorizontalKeyIntervalMinMs = positionCorrectionSettings.HorizontalKeyIntervalMinMs;
            corrector.HorizontalKeyIntervalMaxMs = positionCorrectionSettings.HorizontalKeyIntervalMaxMs;
            corrector.VerticalKeyIntervalMinMs = positionCorrectionSettings.VerticalKeyIntervalMinMs;
            corrector.VerticalKeyIntervalMaxMs = positionCorrectionSettings.VerticalKeyIntervalMaxMs;
            corrector.ConsecutiveJumpCount = positionCorrectionSettings.ConsecutiveJumpCount;
            corrector.ConsecutiveJumpIntervalMs = positionCorrectionSettings.ConsecutiveJumpIntervalMs;
            corrector.ConsecutiveJumpThreshold = positionCorrectionSettings.ConsecutiveJumpThreshold;
            corrector.HorizontalHoldWalkThreshold = positionCorrectionSettings.HorizontalHoldWalkThreshold;
            corrector.ClimbEscapeJumpKeys = positionCorrectionSettings.GetEffectiveClimbJumpKeys(); // ★ 爬繩逃脫鍵
        }

        /// <summary>
        /// ★ 定期位置檢查：比對當前位置與腳本中最近事件的錄製座標
        /// 偏差超過容差時自動修正（寬鬆修正，正負5~8px即可）
        /// </summary>
        private void PeriodicPositionCheck(List<MacroEvent> events, double currentScriptTime)
        {
            if (minimapTracker == null || !minimapTracker.IsCalibrated) return;
            if (positionCorrector != null && positionCorrector.IsLoopCorrectionLimitReached()) return;

            // ★ 找最近的有錄製座標的事件
            MacroEvent? nearest = null;
            double minDist = double.MaxValue;
            foreach (var evt in events)
            {
                if (evt.RecordedX < 0 || evt.RecordedY < 0) continue;
                double dist = Math.Abs(evt.Timestamp - currentScriptTime);
                if (dist < minDist) { minDist = dist; nearest = evt; }
            }

            if (nearest == null) return;

            var (cx, cy, ok) = minimapTracker.ReadPosition();
            if (!ok) return;

            int dx = nearest.RecordedX - cx;
            int dy = nearest.RecordedY - cy;
            // ★ 觸發閾值：分別使用水平/垂直容差，比修正停止容差寬鬆（+3 或 *1.5）
            int hTrigger = Math.Max(positionCorrectionSettings.HorizontalTolerance + 3,
                                    (int)(positionCorrectionSettings.HorizontalTolerance * 1.5));
            int vTrigger = Math.Max(positionCorrectionSettings.VerticalTolerance + 3,
                                    (int)(positionCorrectionSettings.VerticalTolerance * 1.5));

            bool xOff = positionCorrectionSettings.EnableHorizontalCorrection && Math.Abs(dx) > hTrigger;
            bool yOff = positionCorrectionSettings.EnableVerticalCorrection && Math.Abs(dy) > vTrigger;

            if (!xOff && !yOff) return; // 位置正常

            Dispatcher.BeginInvoke(new Action(() =>
            {
                AddLog($"📍 定期檢查: 目前({cx},{cy}) 腳本({nearest.RecordedX},{nearest.RecordedY}) 偏差({dx:+#;-#;0},{dy:+#;-#;0}) 觸發閾值 H±{hTrigger} V±{vTrigger} → 修正中");
                lblPlaybackStatus.Text = $"位置偏差修正中...";
                lblPlaybackStatus.Foreground = ToBrush(Color.Orange);
            }));

            if (positionCorrector == null)
            {
                positionCorrector = new PositionCorrector();
                positionCorrector.OnLog += (msg) =>
                {
                    try { Dispatcher.BeginInvoke(new Action(() => AddLog($"[修正] {msg}"))); } catch { }
                };
            }
            ApplyCorrectorSettings(positionCorrector);

            var result = positionCorrector.CorrectPosition(minimapTracker, nearest.RecordedX, nearest.RecordedY);
            if (result != null)
            {
                var r = result;
                Dispatcher.BeginInvoke(new Action(() =>
                    AddLog(r.Success
                        ? $"✅ 定期修正完成: ({r.FinalX},{r.FinalY}) 容差 H±{positionCorrector.LastHorizontalTolerance} V±{positionCorrector.LastVerticalTolerance} {r.ElapsedMs}ms"
                        : $"⚠️ 定期修正: {r.Message}")));
            }
        }

        /// <summary>
        /// 執行位置修正（可指定目標座標，-1 代表使用全域設定）
        /// </summary>
        private CorrectionResult? ExecutePositionCorrectionTo(int targetX, int targetY)
        {
            if (minimapTracker == null || !minimapTracker.IsCalibrated)
                return null;

            if (positionCorrector == null)
                positionCorrector = new PositionCorrector();

            positionCorrector.TargetWindow = targetWindowHandle;
            positionCorrector.ExternalKeySender = SendKeyForCorrection;
            positionCorrector.IsAnimationLockedCheck = () => IsAnimationLocked;
            positionCorrector.IsClimbingCheck = () => DetectIsClimbing(); // ★ ROI 爬繩偵測
            positionCorrector.Tolerance = positionCorrectionSettings.Tolerance;
            positionCorrector.SoftToleranceMin = positionCorrectionSettings.SoftToleranceMin;
            positionCorrector.SoftToleranceMax = positionCorrectionSettings.SoftToleranceMax;
            positionCorrector.HorizontalTolerance = positionCorrectionSettings.HorizontalTolerance;
            positionCorrector.VerticalTolerance = positionCorrectionSettings.VerticalTolerance;
            positionCorrector.MaxCorrectionTimeMs = positionCorrectionSettings.MaxCorrectionTimeMs;
            positionCorrector.MoveLeftKeys = positionCorrectionSettings.GetEffectiveLeftKeys();
            positionCorrector.MoveRightKeys = positionCorrectionSettings.GetEffectiveRightKeys();
            positionCorrector.MoveUpKeys = positionCorrectionSettings.GetEffectiveUpKeys();
            positionCorrector.MoveDownKeys = positionCorrectionSettings.GetEffectiveDownKeys();
            positionCorrector.ClimbEscapeJumpKeys = positionCorrectionSettings.GetEffectiveClimbJumpKeys();
            positionCorrector.EnableHorizontalCorrection = positionCorrectionSettings.EnableHorizontalCorrection;
            positionCorrector.EnableVerticalCorrection = positionCorrectionSettings.EnableVerticalCorrection;
            positionCorrector.InvertY = positionCorrectionSettings.InvertY;
            positionCorrector.KeyIntervalMinMs = positionCorrectionSettings.KeyIntervalMinMs;
            positionCorrector.KeyIntervalMaxMs = positionCorrectionSettings.KeyIntervalMaxMs;
            positionCorrector.HorizontalKeyIntervalMinMs = positionCorrectionSettings.HorizontalKeyIntervalMinMs;
            positionCorrector.HorizontalKeyIntervalMaxMs = positionCorrectionSettings.HorizontalKeyIntervalMaxMs;
            positionCorrector.VerticalKeyIntervalMinMs = positionCorrectionSettings.VerticalKeyIntervalMinMs;
            positionCorrector.VerticalKeyIntervalMaxMs = positionCorrectionSettings.VerticalKeyIntervalMaxMs;
            positionCorrector.MaxStepsPerCorrection = positionCorrectionSettings.MaxStepsPerCorrection;
            positionCorrector.ConsecutiveJumpCount = positionCorrectionSettings.ConsecutiveJumpCount;
            positionCorrector.ConsecutiveJumpIntervalMs = positionCorrectionSettings.ConsecutiveJumpIntervalMs;
            positionCorrector.ConsecutiveJumpThreshold = positionCorrectionSettings.ConsecutiveJumpThreshold;
            positionCorrector.HorizontalHoldWalkThreshold = positionCorrectionSettings.HorizontalHoldWalkThreshold;

            return positionCorrector.CorrectPosition(minimapTracker, targetX, targetY);
        }


        private void UpdateUI()
        {
            btnStartRecording.IsEnabled = !isRecording && !isPlaying;
            btnStopRecording.IsEnabled = isRecording;
            btnStartPlayback.IsEnabled = !isPlaying && recordedEvents.Count > 0 && !isRecording;
            btnStopPlayback.IsEnabled = isPlaying;
        }

        private double GetCurrentTime()
        {
            return highResTimer.Elapsed.TotalSeconds;
        }

        private void MainWindow_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            lblStatus.Text = "就緒：點擊「開始錄製」開始";
            lblRecordingStatus.Text = "錄製：尚未開始";
            lblPlaybackStatus.Text = "播放：尚未開始";
        }

        // MacroEvent 已提取至 MacroEvent.cs

        /// <summary>
        /// 儲存設定按鈕點擊事件
        /// </summary>
        private void btnSaveSettings_Click(object? sender, EventArgs e)
        {
            ExportSettings();
        }

        private void btnImportSettings_Click(object? sender, EventArgs e)
        {
            ImportSettings();
        }

        /// <summary>
        /// 設定檔案路徑
        /// </summary>
        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MapleMacro",
            "settings.json");

        /// <summary>
        /// 儲存所有設定到檔案
        /// </summary>
        private void SaveSettings()
        {
            try
            {
                SaveSettingsToFile(SettingsFilePath, "設定已儲存");
            }
            catch (Exception ex)
            {
                AddLog($"設定儲存失敗: {ex.Message}");
                MessageBox.Show($"儲存設定時發生錯誤：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportSettings()
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "JSON 設定檔|*.json",
                    DefaultExt = ".json",
                    FileName = "MapleMacroSettings.json"
                };

                if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return;

                SaveSettingsToFile(sfd.FileName, "設定已導出");
                MessageBox.Show("設定已成功導出！\n\n包含：\n• 熱鍵設定\n• 視窗標題\n• 方向鍵模式\n\n注意：自定義按鍵設定現在隨腳本 (.mscript) 一起保存",
                    "導出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AddLog($"設定導出失敗: {ex.Message}");
                MessageBox.Show($"導出設定時發生錯誤：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportSettings()
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog
                {
                    Filter = "JSON 設定檔|*.json",
                    Title = "匯入設定"
                };

                if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return;

                string json = File.ReadAllText(ofd.FileName);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);
                if (settings == null)
                    throw new InvalidOperationException("設定檔格式無效");

                ApplySettings(settings);
                SaveSettingsToFile(SettingsFilePath, "設定已匯入");

                MessageBox.Show("設定已成功匯入！\n\n包含：\n• 熱鍵設定\n• 視窗標題\n• 方向鍵模式\n\n注意：自定義按鍵設定現在隨腳本 (.mscript) 一起保存",
                    "匯入成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AddLog($"設定匯入失敗: {ex.Message}");
                MessageBox.Show($"匯入設定時發生錯誤：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveSettingsToFile(string filePath, string logMessage)
        {
            var settings = BuildSettings();

            string? directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(filePath, json);

            if (!string.IsNullOrEmpty(logMessage))
            {
                AddLog(logMessage);
            }
        }

        /// <summary>
        /// 儲存應用設定
        /// </summary>
        private void SaveAppSettings()
        {
            SaveSettingsToFile(SettingsFilePath, "");
        }

        private AppSettings BuildSettings()
        {
            var settings = new AppSettings
            {
                PlayHotkey = playHotkey,
                StopHotkey = stopHotkey,
                PauseHotkey = pauseHotkey,
                HotkeyEnabled = hotkeyEnabled,
                RecordHotkey = recordHotkey,
                WindowTitle = txtWindowTitle.Text,
                ArrowKeyMode = (int)currentArrowKeyMode,
                PositionCorrectionEnabled = positionCorrectionSettings.Enabled,
                LastScriptPath = currentScriptPath,
                ScheduleTasks = scheduleTasks.Where(t => t.Enabled && t.StartTime > DateTime.Now).ToList(),
                PositionCorrection = positionCorrectionSettings
            };

            return settings;
        }

        private void ApplySettings(AppSettings settings)
        {
            playHotkey = settings.PlayHotkey;
            stopHotkey = settings.StopHotkey;
            pauseHotkey = settings.PauseHotkey;
            hotkeyEnabled = settings.HotkeyEnabled;
            recordHotkey = settings.RecordHotkey;
            txtWindowTitle.Text = settings.WindowTitle;

            // 自動降級：無效的模式值改為 SendToChild(0)
            if (settings.ArrowKeyMode > 2)
                settings.ArrowKeyMode = 0;
            currentArrowKeyMode = (ArrowKeyMode)settings.ArrowKeyMode;
            // ★ 同步主介面方向鍵模式下拉框
            if (cmbArrowMode != null && (int)currentArrowKeyMode < cmbArrowMode.Items.Count)
                cmbArrowMode.SelectedIndex = (int)currentArrowKeyMode;

            // 位置修正設定（先載入設定，再同步 checkbox）
            if (settings.PositionCorrection != null)
            {
                positionCorrectionSettings = settings.PositionCorrection;
            }

            // 座標修正開關
            positionCorrectionEnabled = settings.PositionCorrectionEnabled;
            positionCorrectionSettings.Enabled = positionCorrectionEnabled;

            // 已停用：不再自動載入上次的腳本
            // if (!string.IsNullOrEmpty(settings.LastScriptPath) && File.Exists(settings.LastScriptPath))
            // {
            //     currentScriptPath = settings.LastScriptPath;
            //     LoadScriptFromFile(settings.LastScriptPath);
            // }

            // 載入排程任務（只恢復未來的排程）
            if (settings.ScheduleTasks != null && settings.ScheduleTasks.Count > 0)
            {
                var futureTasks = settings.ScheduleTasks.Where(t => t.StartTime > DateTime.Now && t.Enabled).ToList();
                if (futureTasks.Count > 0)
                {
                    scheduleTasks.AddRange(futureTasks);
                    schedulerTimer.Start();
                    AddLog($"已恢復 {futureTasks.Count} 個排程任務");
                }
            }

            AddLog("全域設定已載入");
            AddLog($"熱鍵：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}, 暫停={GetKeyDisplayName(pauseHotkey)}, 錄製={GetKeyDisplayName(recordHotkey)}");
            AddLog($"方向鍵模式：{currentArrowKeyMode}");

            UpdateUI();
        }

        /// <summary>
        /// 從檔案載入設定
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsFilePath))
                {
                    AddLog("未找到設定檔，使用預設值");
                    return;
                }

                string json = File.ReadAllText(SettingsFilePath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);

                if (settings != null)
                {
                    ApplySettings(settings);
                }

                // 載入爬繩模板
                LoadClimbingTemplate();
            }
            catch (Exception ex)
            {
                AddLog($"設定載入失敗: {ex.Message}");
            }
        }
    }
}
