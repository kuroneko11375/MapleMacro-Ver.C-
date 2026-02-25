using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace MapleStoryMacro
{
    public partial class Form1 : Form
    {
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

        // ★ F7 邊界設定熱鍵 — 狀態機 (0=左,1=右,2=上,3=下)
        private Keys boundaryHotkey = Keys.F7;
        private int _boundarySetState = -1; // -1=未啟動, 0~3=設定中

        // 定時執行
        private List<ScheduleTask> scheduleTasks = new List<ScheduleTask>();
        private System.Windows.Forms.Timer schedulerTimer;

        // Log System
        private readonly int MAX_LOG_LINES = 100;
        private readonly object logLock = new object();
        private string? lastLogMessage = null; // 用於合併重複日誌
        private int lastLogRepeatCount = 0; // 重複計數

        private HashSet<Keys> pressedKeys = new HashSet<Keys>();

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

        // 方向鍵發送模式
        private enum ArrowKeyMode
        {
            SendToChild,            // ThreadAttach + PostMessage（背景走路用）
            ThreadAttachWithBlocker, // ThreadAttach + Blocker（嘗試避免影響前景）
            SendInputWithBlock,     // SendInput + Blocker（嘗試避免影響前景）
        }
        private ArrowKeyMode currentArrowKeyMode = ArrowKeyMode.SendToChild;

        // 鍵盤阻擋器（用於 Blocker 模式）
        private KeyboardBlocker? keyboardBlocker;

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

        public Form1()
        {
            InitializeComponent();

            // 初始化日誌系統
            Logger.Initialize(enableDetailedLogging: false, logRetentionDays: 7);
            Logger.Info("=== MapleStory Macro 啟動 ===");
            Logger.Info($"版本: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}");


            // 初始化自定義按鍵槽位
            for (int i = 0; i < 15; i++)
            {
                customKeySlots[i] = new CustomKeySlot { SlotNumber = i + 1 };
            }

            // 初始化定時器
            schedulerTimer = new System.Windows.Forms.Timer();
            schedulerTimer.Interval = 1000; // 每秒檢查一次
            schedulerTimer.Tick += SchedulerTimer_Tick;

            // Bind all button events
            btnStartRecording.Click += BtnStartRecording_Click;
            btnStopRecording.Click += BtnStopRecording_Click;
            btnSaveScript.Click += BtnSaveScript_Click;
            btnLoadScript.Click += BtnLoadScript_Click;
            // btnClearEvents 和 btnViewEvents 功能已整合到編輯器中
            btnEditEvents.Click += BtnEditEvents_Click;
            btnStartPlayback.Click += BtnStartPlayback_Click;
            btnStopPlayback.Click += BtnStopPlayback_Click;
            txtPauseHotkey.Click += BtnPausePlayback_Click;
            btnRefreshWindow.Click += BtnRefreshWindow_Click;
            btnLockWindow.Click += BtnLockWindow_Click;
            btnHotkeySettings.Click += (s, e) => OpenHotkeySettings();
            btnCustomKeys.Click += (s, e) => OpenCustomKeySettings();
            btnScheduler.Click += (s, e) => OpenSchedulerSettings();
            btnStatistics.Click += (s, e) => ShowStatistics();
            btnMemoryScanner.Click += (s, e) => OpenMemoryScannerChoice();
            FormClosing += Form1_FormClosing;

            // ★ 方向鍵模式下拉框事件（控件已在 Designer 中定義）
            cmbArrowMode.SelectedIndex = (int)currentArrowKeyMode;
            cmbArrowMode.SelectedIndexChanged += (s, e) =>
            {
                currentArrowKeyMode = (ArrowKeyMode)cmbArrowMode.SelectedIndex;
                SaveAppSettings();
                AddLog($"方向鍵模式切換: {currentArrowKeyMode}");
            };

            // Enable KeyPreview to capture all keys
            this.KeyPreview = true;
            this.KeyDown += Form1_KeyDown;
            this.KeyUp += Form1_KeyUp;

            // Initialize keyboard hook for recording
            keyboardHook = new KeyboardHookDLL(KeyboardHookDLL.KeyboardHookMode.LowLevel);
            keyboardHook.OnKeyEvent += KeyboardHook_OnKeyEvent;

            // Initialize global hotkey hook (延遲安裝到 Form1_Shown)
            hotkeyHook = new KeyboardHookDLL(KeyboardHookDLL.KeyboardHookMode.LowLevel);
            hotkeyHook.OnKeyEvent += HotkeyHook_OnKeyEvent;

            // 延遲初始化 - 在表單顯示後再載入設定和啟動定時器
            this.Shown += Form1_Shown;

            // 先設定初始狀態
            lblWindowStatus.Text = "視窗: 未鎖定";
            lblWindowStatus.ForeColor = Color.Gray;
            UpdateUI();
            UpdatePauseButtonState();
        }

        private void Form1_Shown(object? sender, EventArgs e)
        {
            // 初始化日誌右鍵選單事件
            menuCopyLog.Click += MenuCopyLog_Click;
            menuClearLog.Click += MenuClearLog_Click;

            // Initialize log system
            monitorTimer.Interval = 500;
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
                        numPlayTimes.Value = Math.Max(1, Math.Min(9999, task.LoopCount));
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
            if (lstLog.InvokeRequired)
            {
                lstLog.BeginInvoke(new Action(() => AddLogInternal(timestamp, message)));
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
                            lstLog.BeginUpdate();
                            lstLog.Items[lstLog.Items.Count - 1] = $"[{timestamp}] {message} (×{lastLogRepeatCount + 1})";
                            lstLog.EndUpdate();
                            lstLog.TopIndex = lstLog.Items.Count - 1;
                        }
                    }
                }
                else
                {
                    // 新的日誌訊息
                    lastLogMessage = message;
                    lastLogRepeatCount = 0;

                    lstLog.BeginUpdate();
                    string logEntry = $"[{timestamp}] {message}";
                    lstLog.Items.Add(logEntry);

                    // 限制日誌數量
                    while (lstLog.Items.Count > MAX_LOG_LINES)
                    {
                        lstLog.Items.RemoveAt(0);
                    }
                    lstLog.EndUpdate();

                    // 自動滾動到最新
                    lstLog.TopIndex = lstLog.Items.Count - 1;
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

        private void Form1_KeyDown(object? sender, KeyEventArgs e)
        {
            // 當鍵盤鉤子啟用時，不重複錄製（避免雙重事件）
            if (isRecording && keyboardHook != null && !keyboardHook.IsInstalled)
            {
                if (recordStartTime == 0)
                    recordStartTime = GetCurrentTime();

                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = e.KeyCode,
                    EventType = "down",
                    Timestamp = GetCurrentTime() - recordStartTime
                });

                lblRecordingStatus.Text = $"錄製中: {recordedEvents.Count} 個事件 | 最後: {e.KeyCode}";
                AddLog($"按鍵按下: {e.KeyCode}");
            }
        }

        private void Form1_KeyUp(object? sender, KeyEventArgs e)
        {
            // 當鍵盤鉤子啟用時，不重複錄製（避免雙重事件）
            if (isRecording && keyboardHook != null && !keyboardHook.IsInstalled)
            {
                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = e.KeyCode,
                    EventType = "up",
                    Timestamp = GetCurrentTime() - recordStartTime
                });

                lblRecordingStatus.Text = $"錄製中: {recordedEvents.Count} 個事件 | 最後: {e.KeyCode}";
                AddLog($"按鍵放開: {e.KeyCode}");
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
                StartPosition = FormStartPosition.CenterParent,
                Owner = this
            };

            // 提示標籤
            Label lblHint = new Label
            {
                Text = "★ 雙擊視窗名稱即可快速選擇",
                Dock = DockStyle.Top,
                Height = 25,
                ForeColor = Color.Blue,
                Font = new Font("microsoft yahei ui", 9F, FontStyle.Bold),
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
                DialogResult = DialogResult.OK
            };

            Button cancelBtn = new Button
            {
                Text = "取消",
                Width = 80,
                Height = 30,
                Left = 290,
                Top = 5,
                DialogResult = DialogResult.Cancel
            };

            btnPanel.Controls.Add(okBtn);
            btnPanel.Controls.Add(cancelBtn);

            // 控制項加入順序：先 Fill，再 Top/Bottom
            windowSelector.Controls.Add(listBox);
            windowSelector.Controls.Add(lblHint);
            windowSelector.Controls.Add(btnPanel);

            listBox.DisplayMember = "Title";

            if (windowSelector.ShowDialog() == DialogResult.OK && listBox.SelectedIndex >= 0)
            {
                ProcessItem? selected = listBox.SelectedItem as ProcessItem;
                if (selected != null)
                {
                    targetWindowHandle = selected.Handle;
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
                lblWindowStatus.ForeColor = Color.Green;
                AddLog($"已鎖定視窗: {targetWindowHandle} [{arch}]");
            }
            else
            {
                lblWindowStatus.Text = "視窗: 未找到 - 前景模式";
                lblWindowStatus.ForeColor = Color.Red;
                targetWindowHandle = IntPtr.Zero;
                AddLog("未找到目標視窗");
            }
        }

        private void ReleasePressedKeys()
        {
            if (pressedKeys.Count == 0)
                return;

            Keys[] keysToRelease = pressedKeys.ToArray();
            pressedKeys.Clear();

            foreach (var key in keysToRelease)
            {
                SendKeyEvent(new MacroEvent
                {
                    KeyCode = key,
                    EventType = "up",
                    Timestamp = 0
                });
            }
        }

        private void BtnStartRecording_Click(object? sender, EventArgs e)
        {
            if (isRecording) return;

            recordedEvents.Clear();
            recordStartTime = 0;
            isRecording = true;

            lblRecordingStatus.Text = "錄製中...";
            lblRecordingStatus.ForeColor = Color.Red;
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
            lblRecordingStatus.ForeColor = Color.Green;

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

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // 建立腳本資料
                    var scriptData = new ScriptData
                    {
                        Name = Path.GetFileNameWithoutExtension(sfd.FileName),
                        LoopCount = (int)numPlayTimes.Value,
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
                            RecordedY = evt.RecordedY
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

            if (ofd.ShowDialog() == DialogResult.OK)
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
                            RecordedY = evt.RecordedY
                        });
                    }

                    // 載入循環次數
                    numPlayTimes.Value = Math.Max(1, Math.Min(9999, scriptData.LoopCount));

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

        private void BtnClearEvents_Click(object? sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                AddLog("⚠️ 目前沒有任何事件可清除");
                return;
            }

            // 顯示警告視窗
            DialogResult result = MessageBox.Show(
                $"確定要清除所有 {recordedEvents.Count} 個事件嗎？\n\n此操作無法復原！",
                "⚠️ 警告 - 清除所有事件",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            if (result == DialogResult.Yes)
            {
                recordedEvents.Clear();
                lblRecordingStatus.Text = "已清除 | 事件數: 0";
                AddLog("✅ 已清除所有事件");

                // 更新 UI 狀態
                UpdateUI();
            }
        }

        private void BtnViewEvents_Click(object? sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                AddLog("⚠️ 沒有事件可顯示");
                return;
            }

            Form eventViewer = new Form
            {
                Text = "檢視事件",
                Width = 700,
                Height = 500,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this
            };

            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true
            };

            dgv.Columns.Add("KeyCode", "按鍵");
            dgv.Columns.Add("EventType", "類型");
            dgv.Columns.Add("Timestamp", "時間 (秒)");

            foreach (MacroEvent evt in recordedEvents)
            {
                string keyName = GetKeyDisplayName(evt.KeyCode);
                string eventType = evt.EventType == "down" ? "按下" : "放開";
                dgv.Rows.Add(keyName, eventType, evt.Timestamp.ToString("F3"));
            }

            eventViewer.Controls.Add(dgv);
            AddLog($"正在檢視 {recordedEvents.Count} 個事件");
            eventViewer.ShowDialog();
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
            AddLog("正在開啟事件編輯器...");

            var foldedActions = ConvertToFoldedActions();
            
            Form editorForm = new Form
            {
                Text = $"腳本編輯器 ({recordedEvents.Count} 個事件 → {foldedActions.Count} 個動作)",
                Width = 750,
                Height = 650,
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Color.FromArgb(30, 30, 35),
                ForeColor = Color.White,
                Owner = this
            };

            // 說明標籤
            var lblHint = new Label
            {
                Text = "💡 雙擊修改持續時間 | Delete 或按鈕刪除 | ➕插入按鍵 | 右側可複製/匯入 JSON",
                Top = 10,
                Left = 10,
                Width = 710,
                ForeColor = Color.Cyan,
                Font = new Font("Microsoft JhengHei", 9)
            };

            // 左側：事件列表
            var lstActions = new ListView
            {
                Top = 40,
                Left = 10,
                Width = 450,
                Height = 480,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                BackColor = Color.FromArgb(25, 25, 30),
                ForeColor = Color.White,
                Font = new Font("Consolas", 10)
            };
            
            lstActions.Columns.Add("按鍵", 80);
            lstActions.Columns.Add("重複", 50);
            lstActions.Columns.Add("持續時間", 100);
            lstActions.Columns.Add("狀態", 60);
            lstActions.Columns.Add("時間點", 90);

            // 刷新列表的方法
            Action refreshList = () =>
            {
                lstActions.Items.Clear();
                foreach (var action in foldedActions)
                {
                    string duration = action.Duration >= 1.0 ? $"{action.Duration:F2}秒" : $"{action.Duration * 1000:F0}ms";
                    var item = new ListViewItem(action.KeyName);
                    item.SubItems.Add(action.RepeatCount > 1 ? $"x{action.RepeatCount}" : "-");
                    item.SubItems.Add(duration);
                    item.SubItems.Add(action.IsReleased ? "完成" : "按住中");
                    item.SubItems.Add($"{action.PressTime:F3}s");
                    item.Tag = action;
                    item.ForeColor = action.IsReleased ? Color.LightGreen : Color.Orange;
                    lstActions.Items.Add(item);
                }
                editorForm.Text = $"腳本編輯器 ({foldedActions.Count} 個動作)";
            };
            refreshList();

            // 雙擊編輯持續時間
            lstActions.DoubleClick += (s, args) =>
            {
                if (lstActions.SelectedItems.Count == 0) return;
                var action = lstActions.SelectedItems[0].Tag as FoldedKeyAction;
                if (action == null) return;

                using var editForm = new Form
                {
                    Text = $"編輯 {action.KeyName}",
                    Size = new Size(320, 200),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    BackColor = Color.FromArgb(40, 40, 45)
                };

                var lblDur = new Label { Text = "持續時間 (秒):", Location = new Point(20, 25), AutoSize = true, ForeColor = Color.White };
                var numDuration = new NumericUpDown
                {
                    Location = new Point(130, 22),
                    Size = new Size(120, 25),
                    Minimum = 0.01M,
                    Maximum = 9999,
                    DecimalPlaces = 3,
                    Value = (decimal)Math.Max(0.01, action.Duration),
                    Increment = 0.1M
                };

                var lblRepeat = new Label { Text = "重複次數:", Location = new Point(20, 60), AutoSize = true, ForeColor = Color.White };
                var numRepeat = new NumericUpDown
                {
                    Location = new Point(130, 57),
                    Size = new Size(80, 25),
                    Minimum = 1,
                    Maximum = 9999,
                    Value = action.RepeatCount
                };

                var lblInfo = new Label
                {
                    Text = $"原始: {action.Duration:F3}秒, x{action.RepeatCount}",
                    Location = new Point(20, 95),
                    AutoSize = true,
                    ForeColor = Color.Gray
                };

                var btnOk = new Button
                {
                    Text = "確定",
                    Location = new Point(70, 125),
                    Size = new Size(80, 30),
                    DialogResult = DialogResult.OK,
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(60, 140, 80),
                    ForeColor = Color.White
                };

                var btnCancel = new Button
                {
                    Text = "取消",
                    Location = new Point(160, 125),
                    Size = new Size(80, 30),
                    DialogResult = DialogResult.Cancel,
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(100, 100, 100),
                    ForeColor = Color.White
                };

                editForm.Controls.AddRange(new Control[] { lblDur, numDuration, lblRepeat, numRepeat, lblInfo, btnOk, btnCancel });
                editForm.AcceptButton = btnOk;
                editForm.CancelButton = btnCancel;

                if (editForm.ShowDialog() == DialogResult.OK)
                {
                    double newDuration = (double)numDuration.Value;
                    int newRepeat = (int)numRepeat.Value;
                    
                    action.ReleaseTime = action.PressTime + newDuration;
                    action.RepeatCount = newRepeat;
                    
                    refreshList();
                    AddLog($"✅ 已修改 {action.KeyName}: {newDuration:F3}秒 x{newRepeat}");
                }
            };

            // Delete 鍵刪除
            lstActions.KeyDown += (s, args) =>
            {
                if (args.KeyCode == Keys.Delete && lstActions.SelectedItems.Count > 0)
                {
                    var action = lstActions.SelectedItems[0].Tag as FoldedKeyAction;
                    if (action == null) return;

                    if (MessageBox.Show($"確定刪除 {action.KeyName}？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        foldedActions.Remove(action);
                        refreshList();
                        AddLog($"✅ 已刪除 {action.KeyName}");
                    }
                }
            };

            // 右側：JSON 面板
            var lblJson = new Label
            {
                Text = "📋 JSON (給 AI 用)",
                Top = 40,
                Left = 470,
                AutoSize = true,
                ForeColor = Color.White
            };

            var txtJson = new TextBox
            {
                Top = 65,
                Left = 470,
                Width = 255,
                Height = 350,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                BackColor = Color.FromArgb(20, 20, 25),
                ForeColor = Color.LightGreen,
                Font = new Font("Consolas", 9),
                WordWrap = false
            };

            // 更新 JSON 顯示
            Action updateJson = () =>
            {
                var jsonData = foldedActions.Select(a => new
                {
                    key = a.KeyName,
                    duration = Math.Round(a.Duration, 3),
                    repeat = a.RepeatCount,
                    time = Math.Round(a.PressTime, 3)
                }).ToList();
                txtJson.Text = System.Text.Json.JsonSerializer.Serialize(jsonData, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            };
            updateJson();

            // 複製 JSON 按鈕
            var btnCopyJson = new Button
            {
                Text = "📋 複製",
                Top = 425,
                Left = 470,
                Size = new Size(75, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 120, 180),
                ForeColor = Color.White
            };
            btnCopyJson.Click += (s, args) =>
            {
                updateJson();
                Clipboard.SetText(txtJson.Text);
                AddLog("✅ 已複製 JSON");
                MessageBox.Show("已複製！可貼給 AI 修改。", "複製成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            // 匯入 JSON 按鈕
            var btnImportJson = new Button
            {
                Text = "📥 匯入",
                Top = 425,
                Left = 555,
                Size = new Size(75, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 150, 100),
                ForeColor = Color.White
            };
            btnImportJson.Click += (s, args) =>
            {
                try
                {
                    var json = txtJson.Text.Trim();
                    if (string.IsNullOrEmpty(json)) return;

                    using var doc = System.Text.Json.JsonDocument.Parse(json);
                    var newActions = new List<FoldedKeyAction>();
                    double currentTime = 0;

                    foreach (var element in doc.RootElement.EnumerateArray())
                    {
                        string keyName = element.GetProperty("key").GetString() ?? "A";
                        double duration = element.TryGetProperty("duration", out var d) ? d.GetDouble() : 0.1;
                        int repeat = element.TryGetProperty("repeat", out var r) ? r.GetInt32() : 1;
                        double time = element.TryGetProperty("time", out var t) ? t.GetDouble() : currentTime;

                        if (Enum.TryParse<Keys>(keyName, true, out Keys keyCode))
                        {
                            newActions.Add(new FoldedKeyAction
                            {
                                KeyCode = keyCode,
                                KeyName = GetKeyDisplayName(keyCode),
                                PressTime = time,
                                ReleaseTime = time + duration,
                                RepeatCount = repeat,
                                IsReleased = true
                            });
                        }
                        currentTime = time + duration + 0.01;
                    }

                    if (newActions.Count > 0)
                    {
                        foldedActions.Clear();
                        foldedActions.AddRange(newActions);
                        refreshList();
                        AddLog($"✅ 已匯入 {newActions.Count} 個動作");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"JSON 格式錯誤:\n{ex.Message}", "匯入失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // 貼上按鈕
            var btnPaste = new Button
            {
                Text = "📋 貼上",
                Top = 425,
                Left = 640,
                Size = new Size(75, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(100, 100, 120),
                ForeColor = Color.White
            };
            btnPaste.Click += (s, args) =>
            {
                if (Clipboard.ContainsText())
                {
                    txtJson.Text = Clipboard.GetText();
                }
            };

            // 底部按鈕面板
            var btnClear = new Button
            {
                Text = "🗑️ 清空全部",
                Top = 530,
                Left = 10,
                Size = new Size(100, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(150, 60, 60),
                ForeColor = Color.White
            };
            btnClear.Click += (s, args) =>
            {
                if (MessageBox.Show($"確定清空全部 {foldedActions.Count} 個動作？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    foldedActions.Clear();
                    refreshList();
                    updateJson();
                    AddLog("✅ 已清空所有動作");
                }
            };

            // ★ 插入按鍵按鈕
            var btnInsert = new Button
            {
                Text = "➕ 插入按鍵",
                Top = 530,
                Left = 120,
                Size = new Size(100, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(80, 120, 180),
                ForeColor = Color.White
            };

            // ★ 刪除選取按鈕
            var btnDelete = new Button
            {
                Text = "🗑️ 刪除選取",
                Top = 530,
                Left = 228,
                Size = new Size(100, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(180, 80, 60),
                ForeColor = Color.White
            };
            btnDelete.Click += (s, args) =>
            {
                if (lstActions.SelectedItems.Count == 0) { MessageBox.Show("請先選取要刪除的動作"); return; }
                var action = lstActions.SelectedItems[0].Tag as FoldedKeyAction;
                if (action == null) return;
                if (MessageBox.Show($"確定刪除 {action.KeyName}？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    foldedActions.Remove(action);
                    refreshList();
                    updateJson();
                    AddLog($"✅ 已刪除 {action.KeyName}");
                }
            };
            btnInsert.Click += (s, args) =>
            {
                using var insertForm = new Form
                {
                    Text = "插入按鍵事件",
                    Size = new Size(350, 280),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false, MinimizeBox = false,
                    BackColor = Color.FromArgb(40, 40, 45),
                    KeyPreview = true
                };

                var lblKey = new Label { Text = "按鍵（點擊框後按鍵）:", Location = new Point(20, 20), AutoSize = true, ForeColor = Color.White };
                var txtKey = new TextBox
                {
                    Location = new Point(20, 45), Size = new Size(290, 25),
                    BackColor = Color.FromArgb(55, 55, 60), ForeColor = Color.Yellow,
                    Text = "點擊後按下按鍵...", Tag = (Keys?)null, Cursor = Cursors.Arrow
                };
                txtKey.Click += (ts, te) => { txtKey.Text = "按下按鍵..."; txtKey.ForeColor = Color.Yellow; };
                insertForm.KeyDown += (ts, te) =>
                {
                    if (txtKey.Focused || txtKey.ForeColor == Color.Yellow)
                    {
                        te.Handled = true; te.SuppressKeyPress = true;
                        txtKey.Tag = (Keys?)te.KeyCode;
                        txtKey.Text = te.KeyCode.ToString();
                        txtKey.ForeColor = Color.White;
                    }
                };

                var lblDur = new Label { Text = "持續時間(秒):", Location = new Point(20, 80), AutoSize = true, ForeColor = Color.White };
                var numDur = new NumericUpDown { Location = new Point(130, 77), Size = new Size(100, 25), Minimum = 0.01M, Maximum = 9999, DecimalPlaces = 3, Value = 0.1M, Increment = 0.05M };

                var lblRepeat = new Label { Text = "重複次數:", Location = new Point(20, 115), AutoSize = true, ForeColor = Color.White };
                var numRepeat = new NumericUpDown { Location = new Point(130, 112), Size = new Size(80, 25), Minimum = 1, Maximum = 9999, Value = 1 };

                var lblPos = new Label { Text = "插入位置:", Location = new Point(20, 150), AutoSize = true, ForeColor = Color.White };
                var rdoAfter = new RadioButton { Text = "選取項目之後", Location = new Point(130, 148), AutoSize = true, ForeColor = Color.LightGray, Checked = true };
                var rdoEnd = new RadioButton { Text = "末尾", Location = new Point(260, 148), AutoSize = true, ForeColor = Color.LightGray };

                var btnOk = new Button { Text = "插入", Location = new Point(120, 190), Size = new Size(80, 30), BackColor = Color.FromArgb(50, 150, 80), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                var btnCn = new Button { Text = "取消", Location = new Point(210, 190), Size = new Size(80, 30), BackColor = Color.FromArgb(100, 100, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

                btnOk.Click += (ts, te) =>
                {
                    if (txtKey.Tag == null) { MessageBox.Show("請先按下要插入的按鍵"); return; }
                    Keys key = (Keys)txtKey.Tag;
                    double duration = (double)numDur.Value;
                    int repeat = (int)numRepeat.Value;

                    // 計算插入位置的時間點
                    double insertTime;
                    int insertIndex;
                    if (rdoAfter.Checked && lstActions.SelectedItems.Count > 0)
                    {
                        var selAction = lstActions.SelectedItems[0].Tag as FoldedKeyAction;
                        insertIndex = foldedActions.IndexOf(selAction!) + 1;
                        insertTime = selAction!.ReleaseTime + 0.01;
                    }
                    else
                    {
                        insertIndex = foldedActions.Count;
                        insertTime = foldedActions.Count > 0 ? foldedActions.Last().ReleaseTime + 0.01 : 0;
                    }

                    var newAction = new FoldedKeyAction
                    {
                        KeyCode = key,
                        KeyName = key.ToString(),
                        PressTime = insertTime,
                        ReleaseTime = insertTime + duration,
                        RepeatCount = repeat,
                        IsReleased = true
                    };

                    foldedActions.Insert(insertIndex, newAction);
                    refreshList();
                    updateJson();
                    AddLog($"✅ 已插入 {key} (持續{duration}s x{repeat})");
                    insertForm.Close();
                };
                btnCn.Click += (ts, te) => insertForm.Close();

                insertForm.Controls.AddRange(new Control[] { lblKey, txtKey, lblDur, numDur, lblRepeat, numRepeat, lblPos, rdoAfter, rdoEnd, btnOk, btnCn });
                insertForm.ShowDialog();
            };

            var btnSave = new Button
            {
                Text = "💾 儲存並關閉",
                Top = 530,
                Left = 510,
                Size = new Size(120, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(50, 150, 80),
                ForeColor = Color.White
            };
            btnSave.Click += (s, args) =>
            {
                RebuildEventsFromActions(foldedActions);
                lblRecordingStatus.Text = $"已編輯 | 事件數: {recordedEvents.Count}";
                AddLog($"✅ 已儲存 {recordedEvents.Count} 個事件");
                editorForm.DialogResult = DialogResult.OK;
                editorForm.Close();
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Top = 530,
                Left = 640,
                Size = new Size(80, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(100, 100, 100),
                ForeColor = Color.White
            };
            btnCancel.Click += (s, args) =>
            {
                editorForm.DialogResult = DialogResult.Cancel;
                editorForm.Close();
            };

            // 狀態標籤
            var lblStatus = new Label
            {
                Text = $"動作數: {foldedActions.Count}",
                Top = 538,
                Left = 338,
                AutoSize = true,
                ForeColor = Color.Cyan
            };

            // 監聽列表選擇變化來更新 JSON
            lstActions.SelectedIndexChanged += (s, args) => updateJson();

            editorForm.Controls.AddRange(new Control[] {
                lblHint, lstActions,
                lblJson, txtJson, btnCopyJson, btnImportJson, btnPaste,
                btnClear, btnInsert, btnDelete, btnSave, btnCancel, lblStatus
            });

            editorForm.ShowDialog();
            UpdateUI();
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
                StartPosition = FormStartPosition.CenterParent,
                Owner = this
            };

            // 說明標籤
            Label hintLabel = new Label
            {
                Text = "格式: 按鍵名稱 | 類型(down/up) | 時間戳(秒)    每行一個事件，可直接編輯、複製、貼上",
                Top = 10,
                Left = 10,
                Width = 760,
                ForeColor = Color.Blue,
                Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
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
                if (result == DialogResult.Yes)
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
                    if (result == DialogResult.Yes)
                    {
                        saveBtn.PerformClick();
                        if (!hasUnsavedChanges) // 儲存成功才關閉
                            editorForm.Close();
                    }
                    else if (result == DialogResult.No)
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
        /// <summary>
        /// 檢查按鍵是否為修飾鍵（Ctrl/Alt/Shift）
        /// </summary>
        private static bool IsModifierKey(Keys key)
        {
            return key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey ||
                   key == Keys.ShiftKey || key == Keys.LShiftKey || key == Keys.RShiftKey ||
                   key == Keys.Menu || key == Keys.LMenu || key == Keys.RMenu;
        }

        /// <summary>
        /// 將具體的修飾鍵碼轉換為 Keys 修飾旗標
        /// </summary>
        private static Keys ModifierKeyToFlag(Keys key)
        {
            return key switch
            {
                Keys.ControlKey or Keys.LControlKey or Keys.RControlKey => Keys.Control,
                Keys.ShiftKey or Keys.LShiftKey or Keys.RShiftKey => Keys.Shift,
                Keys.Menu or Keys.LMenu or Keys.RMenu => Keys.Alt,
                _ => Keys.None
            };
        }

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
                    if (IsModifierKey(evt.KeyCode))
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
                    if (IsModifierKey(evt.KeyCode))
                    {
                        activeModifiers.Remove(evt.KeyCode);
                    }

                    // 放開時計算持續時間
                    if (keyDownTimes.TryGetValue(evt.KeyCode, out double startTime))
                    {
                        // 計算此按鍵按下期間有哪些修飾鍵是按住的
                        Keys modifiers = Keys.None;
                        if (!IsModifierKey(evt.KeyCode))
                        {
                            foreach (var mod in activeModifiers)
                            {
                                // 修飾鍵必須在主鍵按下之前或同時按下
                                if (mod.Value <= startTime)
                                {
                                    modifiers |= ModifierKeyToFlag(mod.Key);
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
                if (!IsModifierKey(kvp.Key))
                {
                    foreach (var mod in activeModifiers)
                    {
                        if (mod.Value <= kvp.Value)
                            modifiers |= ModifierKeyToFlag(mod.Key);
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

            pressedKeys.Clear();
            lock (lastKeyDownTimestamp) { lastKeyDownTimestamp.Clear(); }

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
                int loopCount = (int)numPlayTimes.Value;
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
                this.BeginInvoke(new Action(() => AddLog($"播放線程已啟動 ({mode}模式, {events.Count} 事件, {loopCount} 循環)")));

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
                        this.BeginInvoke(new Action(() => AddLog("🔄 位置修正設定已套用")));
                    }

                    // ★ 重設循環修正計數
                    if (positionCorrector != null)
                        positionCorrector.ResetLoopCorrectionCount();

                    if (!isPlaying) break;

                    int currentLoopNum = loop;
                    this.BeginInvoke(new Action(() =>
                    {
                        lblPlaybackStatus.Text = $"循環: {currentLoopNum}/{loopCount}";
                        lblPlaybackStatus.ForeColor = Color.Blue;
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
                                        this.BeginInvoke(new Action(() => AddLog("🔄 位置修正設定已即時套用")));
                                    }

                                    double secSinceLastCheck = Stopwatch.GetElapsedTime(lastCorrectionCheckTick).TotalSeconds;
                                    if (secSinceLastCheck >= positionCorrectionSettings.CorrectionCheckIntervalSec)
                                    {
                                        // ★ 阻塞式修正：設定 isCorrecting，暫停自定義按鍵，記錄修正耗時，修正後補償計時
                                        isCorrecting = true;
                                        customKeysPaused = true;
                                        long corrStartTick = Stopwatch.GetTimestamp();

                                        double scriptTime = Stopwatch.GetElapsedTime(loopStartTick).TotalSeconds;
                                        PeriodicPositionCheck(events, scriptTime);

                                        long corrElapsedTicks = Stopwatch.GetTimestamp() - corrStartTick;
                                        customKeysPaused = false;
                                        isCorrecting = false;

                                        // ★ 時間補償：將修正耗時從等待計時和循環計時中扣除
                                        waitStartTick += corrElapsedTicks;
                                        loopStartTick += corrElapsedTicks;
                                        lastCustomKeyCheckTick = Stopwatch.GetTimestamp();

                                        // ★ 修正完成後才重設計時器
                                        lastCorrectionCheckTick = Stopwatch.GetTimestamp();

                                        // ★ 修正後恢復播放狀態標籤
                                        int curLoop = currentLoopNum;
                                        this.BeginInvoke(new Action(() =>
                                        {
                                            lblPlaybackStatus.Text = $"循環: {curLoop}/{loopCount}";
                                            lblPlaybackStatus.ForeColor = Color.Blue;
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
                                this.BeginInvoke(new Action(() =>
                                    AddLog("⏭️ 位置修正事件已跳過（修正未啟用）")));
                                lastTimestamp = evt.Timestamp;
                                continue;
                            }

                            // ★ 使用腳本錄製座標作為修正目標（不再有固定目標模式）
                            int tx = evt.RecordedX >= 0 ? evt.RecordedX : evt.CorrectTargetX;
                            int ty = evt.RecordedY >= 0 ? evt.RecordedY : evt.CorrectTargetY;

                            if (tx < 0 || ty < 0)
                            {
                                this.BeginInvoke(new Action(() =>
                                    AddLog("⏭️ 位置修正事件已跳過（無錄製座標）")));
                                lastTimestamp = evt.Timestamp;
                                continue;
                            }

                            this.BeginInvoke(new Action(() =>
                            {
                                lblPlaybackStatus.Text = $"腳本位置修正中... 目標({tx},{ty})";
                                lblPlaybackStatus.ForeColor = Color.Orange;
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
                                    this.BeginInvoke(new Action(() =>
                                    {
                                        AddLog(r.Success
                                            ? $"✅ 腳本位置修正完成: ({r.FinalX}, {r.FinalY})"
                                            : $"⚠️ 腳本位置修正: {r.Message}");
                                        // ★ 修正後恢復播放狀態
                                        lblPlaybackStatus.Text = $"播放中...";
                                        lblPlaybackStatus.ForeColor = Color.Blue;
                                    }));
                                }
                                else
                                {
                                    this.BeginInvoke(new Action(() =>
                                    {
                                        lblPlaybackStatus.Text = $"播放中...";
                                        lblPlaybackStatus.ForeColor = Color.Blue;
                                    }));
                                }
                            }
                            else
                            {
                                this.BeginInvoke(new Action(() =>
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
                                    this.BeginInvoke(new Action(() =>
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
                            this.BeginInvoke(new Action(() =>
                                AddLog($"✅ 第一個按鍵已發送: {GetKeyDisplayName(firstEvt.KeyCode)} ({firstEvt.EventType})")
                            ));
                        }
                    }

                    Thread.Sleep(200);
                }

                isPlaying = false;
                statistics.EndSession();

                // 清理 Blocker
                if (keyboardBlocker != null)
                {
                    keyboardBlocker.IsBlocking = false;
                    keyboardBlocker.Uninstall();
                }

                ReleasePressedKeys();

                this.BeginInvoke(new Action(() =>
                {
                    lblPlaybackStatus.Text = "播放: 已完成";
                    lblPlaybackStatus.ForeColor = Color.Green;
                    AddLog($"播放完成 - 循環: {statistics.CurrentLoopCount}");
                    UpdateUI();
                    UpdatePauseButtonState();
                }));
            }
            catch (Exception ex)
            {
                statistics.EndSession();

                // 清理 Blocker
                if (keyboardBlocker != null)
                {
                    keyboardBlocker.IsBlocking = false;
                    keyboardBlocker.Uninstall();
                }

                ReleasePressedKeys();
                isPlaying = false;

                this.BeginInvoke(new Action(() =>
                {
                    AddLog($"❌ 播放錯誤: {ex.Message}");
                    UpdateUI();
                    UpdatePauseButtonState();
                }));
            }
        }

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

                    this.BeginInvoke(new Action(() =>
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

                    // 2. 發送按鍵（按下和放開）
                    SendCustomKey(slot.KeyCode, slot.Modifiers);

                    this.BeginInvoke(new Action(() =>
                    {
                        AddLog($"⚡ 自定義按鍵 #{slot.SlotNumber}: {GetKeyDisplayName(slot.KeyCode)}");
                    }));

                    // 3. 按鍵後延遲（等待技能施放完成）
                    if (preDelay > 0)
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            AddLog($"⏳ 延遲 {preDelay} 秒後繼續...");
                        }));
                        Thread.Sleep((int)(preDelay * 1000));
                        totalPauseTime += preDelay;
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
                this.BeginInvoke(new Action(() =>
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
                    pressedKeys.Add(evt.KeyCode);
                }
                else if (evt.EventType == "up")
                {
                    pressedKeys.Remove(evt.KeyCode);
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
        /// 發送滑鼠點擊到背景視窗（Client Area 座標）
        /// 使用 PostMessage 發送 WM_LBUTTONDOWN + WM_LBUTTONUP
        /// </summary>
        private void SendMouseClickToWindow(IntPtr hWnd, int clientX, int clientY)
        {
            IntPtr lParam = (IntPtr)((clientY << 16) | (clientX & 0xFFFF));
            PostMessage(hWnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, lParam);
            Thread.Sleep(50);
            PostMessage(hWnd, WM_LBUTTONUP, IntPtr.Zero, lParam);
        }

        /// <summary>
        /// 執行回程序列：停止腳本 → 冷卻 2000ms → 滑鼠點擊座標 → Enter 確認
        /// </summary>
        private void ExecuteReturnToTown(ScheduleTask task)
        {
            try
            {
                this.BeginInvoke(new Action(() => AddLog($"🏠 開始回程序列：冷卻 2000ms → 點擊({task.ReturnClickX},{task.ReturnClickY}) → Enter")));

                // 1. 冷卻延遲 2000ms（確保腳本完全停止）
                Thread.Sleep(2000);

                bool isBackground = targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle);

                // 2. 滑鼠模擬點擊目標座標
                if (isBackground)
                {
                    SendMouseClickToWindow(targetWindowHandle, task.ReturnClickX, task.ReturnClickY);
                    this.BeginInvoke(new Action(() => AddLog($"🖱️ 已點擊座標 ({task.ReturnClickX}, {task.ReturnClickY})")));
                }
                else
                {
                    this.BeginInvoke(new Action(() => AddLog("⚠️ 前景模式不支援背景滑鼠點擊，跳過")));
                }

                Thread.Sleep(300); // 等待遊戲響應點擊

                // 3. 按下 Enter（確認）
                if (isBackground)
                {
                    SendKeyToWindow(targetWindowHandle, Keys.Enter, true);
                    Thread.Sleep(50);
                    SendKeyToWindow(targetWindowHandle, Keys.Enter, false);
                }
                else
                {
                    SendKeyForeground(Keys.Enter, true);
                    Thread.Sleep(50);
                    SendKeyForeground(Keys.Enter, false);
                }

                this.BeginInvoke(new Action(() => AddLog($"📨 已送出 Enter 確認")));

                // 4. 等待後坐下（如設定）
                Keys sitKey = (Keys)task.SitDownKeyCode;
                if (sitKey != Keys.None)
                {
                    int delayMs = (int)(task.SitDownDelaySeconds * 1000);
                    this.BeginInvoke(new Action(() => AddLog($"⏳ 等待 {task.SitDownDelaySeconds} 秒後坐下...")));
                    Thread.Sleep(delayMs);

                    SendCustomKey(sitKey);
                    this.BeginInvoke(new Action(() => AddLog($"🪑 已坐下：{GetKeyDisplayName(sitKey)}")));
                }

                this.BeginInvoke(new Action(() => AddLog($"✅ 回程序列完成")));
            }
            catch (Exception ex)
            {
                this.BeginInvoke(new Action(() => AddLog($"❌ 回程序列失敗: {ex.Message}")));
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
                this.Invoke(new Action(() =>
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
                this.Invoke(new Action(() =>
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

            lblPlaybackStatus.Text = "播放: 已停止";
            lblPlaybackStatus.ForeColor = Color.Orange;
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

        private void Form1_FormClosing(object? sender, FormClosingEventArgs e)
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
                this.BeginInvoke(new Action(() =>
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
                this.BeginInvoke(new Action(() => HandleBoundaryHotkey()));
                return;
            }

            // 檢查是否為播放熱鍵
            if (keyCode == playHotkey)
            {
                this.BeginInvoke(new Action(() =>
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
                this.BeginInvoke(new Action(() =>
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
                this.BeginInvoke(new Action(() =>
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
                    lblPlaybackStatus.ForeColor = Color.Orange;
                }
                else
                {
                    AddLog($"▶ 播放已繼續 ({GetKeyDisplayName(pauseHotkey)})");
                    lblPlaybackStatus.Text = "狀態: 播放中";
                    lblPlaybackStatus.ForeColor = Color.Green;

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
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(UpdatePauseButtonState));
                return;
            }
            
            if (isPlaying)
            {
                txtPauseHotkey.Enabled = true;
                if (isPaused)
                {
                    // 暫停中 - 顯示繼續
                    txtPauseHotkey.TextButton = "▶ 繼續";
                    txtPauseHotkey.ColorBackground = Color.FromArgb(0, 150, 80); // 綠色
                }
                else
                {
                    // 播放中 - 顯示暫停
                    txtPauseHotkey.TextButton = "⏸ 暫停";
                    txtPauseHotkey.ColorBackground = Color.CornflowerBlue; // 藍色
                }
            }
            else
            {
                // 未播放 - 禁用按鈕
                txtPauseHotkey.Enabled = false;
                txtPauseHotkey.TextButton = "⏸ 暫停";
                txtPauseHotkey.ColorBackground = Color.FromArgb(80, 80, 85); // 灰色
            }
            
            txtPauseHotkey.Invalidate();
        }

        /// <summary>
        /// 檢查暫停狀態，如果暫停則等待
        /// </summary>
        private void CheckPauseState()
        {
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
        /// 開啟熱鍵設定視窗
        /// </summary>
        private void OpenHotkeySettings()
        {
            Form settingsForm = new Form
            {
                Text = "⚙ 熱鍵設定",
                Width = 450,
                Height = 350,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            // 播放熱鍵
            Label lblPlay = new Label { Text = "播放熱鍵：", Left = 20, Top = 30, Width = 100, ForeColor = Color.White };
            TextBox txtPlay = new TextBox
            {
                Left = 130,
                Top = 27,
                Width = 200,
                ReadOnly = true,
                Text = GetKeyDisplayName(playHotkey),
                Tag = playHotkey,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White
            };
            txtPlay.KeyDown += (s, e) =>
            {
                e.SuppressKeyPress = true;
                txtPlay.Text = GetKeyDisplayName(e.KeyCode);
                txtPlay.Tag = e.KeyCode;
            };


            // 停止熱鍵
            Label lblStop = new Label { Text = "停止熱鍵：", Left = 20, Top = 70, Width = 100, ForeColor = Color.White };
            TextBox txtStop = new TextBox
            {
                Left = 130,
                Top = 67,
                Width = 200,
                ReadOnly = true,
                Text = GetKeyDisplayName(stopHotkey),
                Tag = stopHotkey,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White
            };
            txtStop.KeyDown += (s, e) =>
            {
                e.SuppressKeyPress = true;
                txtStop.Text = GetKeyDisplayName(e.KeyCode);
                txtStop.Tag = e.KeyCode;
            };

            // 暫停熱鍵
            Label lblPause = new Label { Text = "暫停熱鍵：", Left = 20, Top = 110, Width = 100, ForeColor = Color.White };
            TextBox txtPause = new TextBox
            {
                Left = 130,
                Top = 107,
                Width = 200,
                ReadOnly = true,
                Text = GetKeyDisplayName(pauseHotkey),
                Tag = pauseHotkey,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White
            };
            txtPause.KeyDown += (s, e) =>
            {
                e.SuppressKeyPress = true;
                txtPause.Text = GetKeyDisplayName(e.KeyCode);
                txtPause.Tag = e.KeyCode;
            };

            // 錄製熱鍵
            Label lblRecord = new Label { Text = "錄製熱鍵：", Left = 20, Top = 150, Width = 100, ForeColor = Color.White };
            TextBox txtRecord = new TextBox
            {
                Left = 130,
                Top = 147,
                Width = 200,
                ReadOnly = true,
                Text = GetKeyDisplayName(recordHotkey),
                Tag = recordHotkey,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White
            };
            txtRecord.KeyDown += (s, e) =>
            {
                e.SuppressKeyPress = true;
                txtRecord.Text = GetKeyDisplayName(e.KeyCode);
                txtRecord.Tag = e.KeyCode;
            };

            // 啟用熱鍵
            CheckBox chkEnabled = new CheckBox
            {
                Text = "啟用全局熱鍵",
                Left = 20,
                Top = 190,
                Width = 150,
                Checked = hotkeyEnabled,
                ForeColor = Color.White
            };

            // 方向鍵模式已移至主介面「設定」群組

            // 提示文字
            Label lblHint = new Label
            {
                Text = "提示：點擊文字框後按下想要的按鍵",
                Left = 20,
                Top = 220,
                Width = 350,
                ForeColor = Color.Gray
            };

            // 按鈕
            Button btnSave = new Button
            {
                Text = "儲存",
                Left = 180,
                Top = 260,
                Width = 80,
                Height = 30,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            Button btnCancel = new Button
            {
                Text = "取消",
                Left = 270,
                Top = 260,
                Width = 80,
                Height = 30,
                BackColor = Color.FromArgb(80, 80, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnSave.Click += (s, args) =>
            {
                playHotkey = (Keys)txtPlay.Tag;
                stopHotkey = (Keys)txtStop.Tag;
                pauseHotkey = (Keys)txtPause.Tag;
                recordHotkey = (Keys)txtRecord.Tag;
                hotkeyEnabled = chkEnabled.Checked;

                AddLog($"設定已儲存：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}, 暫停={GetKeyDisplayName(pauseHotkey)}, 錄製={GetKeyDisplayName(recordHotkey)}");
                SaveAppSettings();
                settingsForm.Close();
            };

            btnCancel.Click += (s, args) => settingsForm.Close();

            settingsForm.Controls.AddRange(new Control[]
            {
                lblPlay, txtPlay, lblStop, txtStop, lblPause, txtPause,
                lblRecord, txtRecord, chkEnabled,
                lblHint, btnSave, btnCancel
            });

            settingsForm.ShowDialog();
        }

        /// <summary>
        /// 開啟自定義按鍵設定視窗
        /// </summary>
        private void OpenCustomKeySettings()
        {
            Form customForm = new Form
            {
                Text = "⚡ 自定義按鍵設定 (15 格)",
                Width = 750,
                Height = 620,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            // 說明標籤
            Label lblTitle = new Label
            {
                Text = "設定最多 15 個自定義按鍵，在腳本播放時按間隔自動施放 | 按鍵欄位點擊後按下按鍵來設定",
                Left = 10,
                Top = 10,
                Width = 720,
                ForeColor = Color.LightGray,
                Font = new Font("microsoft yahei ui", 9F)
            };

            // 建立 DataGridView
            DataGridView dgv = new DataGridView
            {
                Left = 10,
                Top = 35,
                Width = 710,
                Height = 450,
                BackgroundColor = Color.FromArgb(30, 30, 35),
                ForeColor = Color.White,
                GridColor = Color.FromArgb(60, 60, 65),
                BorderStyle = BorderStyle.FixedSingle,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(50, 50, 55),
                    ForeColor = Color.White,
                    Font = new Font("microsoft yahei ui", 9F, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(40, 40, 45),
                    ForeColor = Color.White,
                    SelectionBackColor = Color.FromArgb(0, 100, 180),
                    SelectionForeColor = Color.White,
                    Font = new Font("microsoft yahei ui", 9F)
                },
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EditMode = DataGridViewEditMode.EditOnEnter,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            // 建立欄位
            // #
            var colSlot = new DataGridViewTextBoxColumn
            {
                Name = "Slot",
                HeaderText = "#",
                Width = 30,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, ForeColor = Color.Cyan }
            };

            // 啟用
            var colEnabled = new DataGridViewCheckBoxColumn
            {
                Name = "Enabled",
                HeaderText = "啟用",
                Width = 45
            };

            // 按鍵 - ReadOnly，透過 DGV 層級的 KeyDown 捕獲按鍵（避免 EditingControlWantsInputKey 攔截導航鍵）
            var colKey = new DataGridViewTextBoxColumn
            {
                Name = "KeyCode",
                HeaderText = "按鍵 (選中後按鍵)",
                Width = 120,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(50, 60, 70), ForeColor = Color.LightGreen }
            };

            // 間隔(秒)
            var colInterval = new DataGridViewTextBoxColumn
            {
                Name = "Interval",
                HeaderText = "間隔(秒)",
                Width = 70,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
            };

            // 開始(秒)
            var colStartAt = new DataGridViewTextBoxColumn
            {
                Name = "StartAt",
                HeaderText = "開始(秒)",
                Width = 70,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
            };

            // 暫停
            var colPauseEnabled = new DataGridViewCheckBoxColumn
            {
                Name = "PauseEnabled",
                HeaderText = "暫停",
                Width = 45
            };

            // 暫停(秒)
            var colPauseSeconds = new DataGridViewTextBoxColumn
            {
                Name = "PauseSeconds",
                HeaderText = "暫停(秒)",
                Width = 70,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, ForeColor = Color.Yellow }
            };

            // 延遲(秒)
            var colDelay = new DataGridViewTextBoxColumn
            {
                Name = "Delay",
                HeaderText = "延遲(秒)",
                Width = 70,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, ForeColor = Color.Cyan }
            };

            dgv.Columns.AddRange(new DataGridViewColumn[] { colSlot, colEnabled, colKey, colInterval, colStartAt, colPauseEnabled, colPauseSeconds, colDelay });

            // 填入資料
            for (int i = 0; i < 15; i++)
            {
                var slot = customKeySlots[i];
                string keyDisplay = slot.KeyCode == Keys.None ? "(點擊設定)" : GetKeyDisplayName(slot.KeyCode);
                dgv.Rows.Add(
                    $"#{i + 1}",
                    slot.Enabled,
                    keyDisplay,
                    slot.IntervalSeconds.ToString("F0"),
                    slot.StartAtSecond.ToString("F0"),
                    slot.PauseScriptEnabled,
                    slot.PauseScriptSeconds.ToString("F1"),
                    slot.PreDelaySeconds.ToString("F1")
                );
                dgv.Rows[i].Cells["KeyCode"].Tag = slot.KeyCode;
            }

            // 攔截所有按鍵（包括 Home, End, Delete, Insert, PageUp, PageDown 等導航鍵）
            dgv.PreviewKeyDown += (s, args) =>
            {
                if (dgv.CurrentCell?.ColumnIndex == dgv.Columns["KeyCode"].Index)
                {
                    args.IsInputKey = true;
                }
            };

            // DGV 層級按鍵捕獲：選中 KeyCode 欄位後直接按鍵設定（不需進入編輯模式）
            dgv.KeyDown += (s, args) =>
            {
                if (dgv.CurrentCell?.ColumnIndex != dgv.Columns["KeyCode"].Index)
                    return;

                args.Handled = true;
                args.SuppressKeyPress = true;

                Keys newKey = args.KeyCode;
                string newKeyName = GetKeyDisplayName(newKey);

                int rowIndex = dgv.CurrentCell.RowIndex;
                dgv.CurrentCell.Value = newKeyName;
                dgv.Rows[rowIndex].Cells["KeyCode"].Tag = newKey;
            };

            // 驗證數字欄位 - 只允許數字和小數點
            dgv.CellValidating += (s, args) =>
            {
                string colName = dgv.Columns[args.ColumnIndex].Name;
                if (colName == "Interval" || colName == "StartAt" || colName == "PauseSeconds" || colName == "Delay")
                {
                    string value = args.FormattedValue?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(value) && !double.TryParse(value, out _))
                    {
                        args.Cancel = true;
                        dgv.CancelEdit();
                        MessageBox.Show("請輸入有效的數字！", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            };

            // 說明區域
            Label lblHint = new Label
            {
                Text = "【說明】間隔: 每隔幾秒觸發 | 開始: 腳本播放幾秒後開始 | 暫停: 觸發前暫停腳本 | 延遲: 按鍵後等待\n" +
                       "【執行順序】腳本暫停 → 按下按鍵 → 延遲等待 → 繼續腳本",
                Left = 10,
                Top = 490,
                Width = 710,
                Height = 35,
                ForeColor = Color.LightGray,
                Font = new Font("microsoft yahei ui", 8.5F)
            };

            // 按鈕面板
            Button btnSave = new Button
            {
                Text = "儲存",
                Left = 260,
                Top = 535,
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnCancel = new Button
            {
                Text = "取消",
                Left = 370,
                Top = 535,
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(80, 80, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnClearAll = new Button
            {
                Text = "全部清除",
                Left = 10,
                Top = 535,
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(150, 60, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnClearAll.Click += (s, args) =>
            {
                for (int i = 0; i < 15; i++)
                {
                    dgv.Rows[i].Cells["Enabled"].Value = false;
                    dgv.Rows[i].Cells["KeyCode"].Value = "(點擊設定)";
                    dgv.Rows[i].Cells["KeyCode"].Tag = Keys.None;
                    dgv.Rows[i].Cells["Interval"].Value = "30";
                    dgv.Rows[i].Cells["StartAt"].Value = "0";
                    dgv.Rows[i].Cells["PauseEnabled"].Value = false;
                    dgv.Rows[i].Cells["PauseSeconds"].Value = "0";
                    dgv.Rows[i].Cells["Delay"].Value = "0";
                }
            };

            btnSave.Click += (s, args) =>
            {
                try
                {
                    for (int i = 0; i < 15; i++)
                    {
                        var row = dgv.Rows[i];
                        customKeySlots[i].Enabled = Convert.ToBoolean(row.Cells["Enabled"].Value);
                        var tagValue = row.Cells["KeyCode"].Tag;
                        if (tagValue is Keys key)
                        {
                            customKeySlots[i].KeyCode = key;
                        }
                        else
                        {
                            customKeySlots[i].KeyCode = Keys.None;
                        }
                        customKeySlots[i].Modifiers = Keys.None;
                        customKeySlots[i].IntervalSeconds = double.TryParse(row.Cells["Interval"].Value?.ToString(), out double interval) ? interval : 30;
                        customKeySlots[i].StartAtSecond = double.TryParse(row.Cells["StartAt"].Value?.ToString(), out double startAt) ? startAt : 0;
                        customKeySlots[i].PauseScriptEnabled = Convert.ToBoolean(row.Cells["PauseEnabled"].Value);
                        customKeySlots[i].PauseScriptSeconds = double.TryParse(row.Cells["PauseSeconds"].Value?.ToString(), out double pause) ? pause : 0;
                        customKeySlots[i].PreDelaySeconds = double.TryParse(row.Cells["Delay"].Value?.ToString(), out double delay) ? delay : 0;
                    }

                    int enabledCount = customKeySlots.Count(slot => slot.Enabled && slot.KeyCode != Keys.None);
                    AddLog($"✅ 自定義按鍵設定已更新：{enabledCount} 個已啟用");
                    AddLog("💡 提示：自定義按鍵會隨腳本一起保存");
                    customForm.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"儲存失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            btnCancel.Click += (s, args) => customForm.Close();

            customForm.Controls.AddRange(new Control[] { lblTitle, dgv, lblHint, btnSave, btnCancel, btnClearAll });
            customForm.ShowDialog();
        }

        /// <summary>
        /// 自定義按鍵欄位的 PreviewKeyDown 處理（讓延伸鍵如 End, PageUp, PageDown 等被視為輸入鍵）
        /// </summary>
        private void CustomKeyCell_PreviewKeyDown(object? sender, PreviewKeyDownEventArgs e)
        {
            e.IsInputKey = true;
        }

        /// <summary>
        /// 自定義按鍵欄位的按鍵處理（單鍵設定）
        /// </summary>
        private void CustomKeyCell_KeyDown(object? sender, KeyEventArgs e)
        {
            if (sender is TextBox tb)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;

                // 設定按鍵顯示名稱
                tb.Text = GetKeyDisplayName(e.KeyCode);

                // 找到 DataGridView（需要向上尋找）
                Control? parent = tb.Parent;
                while (parent != null && !(parent is DataGridView))
                {
                    parent = parent.Parent;
                }

                if (parent is DataGridView dgv && dgv.CurrentCell != null)
                {
                    dgv.CurrentCell.Tag = e.KeyCode;
                    dgv.EndEdit();
                }
            }
        }

        /// <summary>
        /// 開啟定時執行設定視窗
        /// </summary>
        private void OpenSchedulerSettings()
        {
            Form schedForm = new Form
            {
                Text = "⏰ 排程管理",
                Width = 650,
                Height = 720,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            Label lblTitle = new Label
            {
                Text = "設定排程任務，可指定腳本、開始/結束時間",
                Left = 20,
                Top = 15,
                Width = 600,
                ForeColor = Color.LightGray,
                Font = new Font("microsoft yahei ui", 10F)
            };

            // ===== 新增排程區塊 =====
            Label lblNewTask = new Label
            {
                Text = "📋 新增排程",
                Left = 20,
                Top = 45,
                Width = 200,
                ForeColor = Color.Cyan,
                Font = new Font("microsoft yahei ui", 10F, FontStyle.Bold)
            };

            // 腳本選擇
            Label lblScript = new Label { Text = "腳本：", Left = 20, Top = 75, Width = 50, ForeColor = Color.White };
            TextBox txtScriptPath = new TextBox
            {
                Left = 75,
                Top = 72,
                Width = 420,
                ReadOnly = true,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White,
                Text = currentScriptPath ?? "(使用當前已載入的腳本)"
            };
            Button btnBrowse = new Button
            {
                Text = "...",
                Left = 500,
                Top = 71,
                Width = 35,
                Height = 25,
                BackColor = Color.FromArgb(80, 80, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            Button btnUseCurrent = new Button
            {
                Text = "當前",
                Left = 540,
                Top = 71,
                Width = 50,
                Height = 25,
                BackColor = Color.FromArgb(0, 100, 160),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnBrowse.Click += (s, args) =>
            {
                OpenFileDialog ofd = new OpenFileDialog
                {
                    Filter = "Maple 腳本|*.mscript|舊版 JSON 腳本|*.json|所有檔案|*.*",
                    Title = "選擇排程腳本"
                };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtScriptPath.Text = ofd.FileName;
                    txtScriptPath.Tag = ofd.FileName;
                }
            };

            btnUseCurrent.Click += (s, args) =>
            {
                if (recordedEvents.Count > 0)
                {
                    txtScriptPath.Text = currentScriptPath ?? "(使用當前已載入的腳本)";
                    txtScriptPath.Tag = currentScriptPath;
                }
                else
                {
                    MessageBox.Show("當前沒有已載入的腳本！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            // 開始時間
            Label lblStart = new Label { Text = "開始：", Left = 20, Top = 108, Width = 50, ForeColor = Color.White };
            DateTimePicker dtpStart = new DateTimePicker
            {
                Left = 75,
                Top = 105,
                Width = 200,
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm:ss",
                Value = DateTime.Now.AddMinutes(1)
            };

            // 結束時間
            Label lblEnd = new Label { Text = "結束：", Left = 290, Top = 108, Width = 50, ForeColor = Color.White };
            CheckBox chkEndTime = new CheckBox
            {
                Text = "",
                Left = 340,
                Top = 105,
                Width = 20,
                Checked = true,
                ForeColor = Color.White
            };
            DateTimePicker dtpEnd = new DateTimePicker
            {
                Left = 365,
                Top = 105,
                Width = 200,
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm:ss",
                Value = DateTime.Now.AddHours(1),
                Enabled = true
            };
            chkEndTime.CheckedChanged += (s, args) => dtpEnd.Enabled = chkEndTime.Checked;

            // 循環次數
            Label lblLoop = new Label { Text = "循環：", Left = 20, Top = 140, Width = 50, ForeColor = Color.White };
            NumericUpDown numLoop = new NumericUpDown
            {
                Left = 75,
                Top = 137,
                Width = 80,
                Minimum = 1,
                Maximum = 9999,
                Value = (int)numPlayTimes.Value,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White
            };

            Label lblLoopHint = new Label
            {
                Text = "勾選「結束」可設定自動停止時間",
                Left = 165,
                Top = 140,
                Width = 300,
                ForeColor = Color.Gray,
                Font = new Font("microsoft yahei ui", 8.5F)
            };

            // ===== 回程設定區塊 =====
            CheckBox chkReturnToTown = new CheckBox
            {
                Text = "🏠 回程（結束時自動點擊回城）",
                Left = 20,
                Top = 170,
                Width = 250,
                ForeColor = Color.FromArgb(100, 220, 160),
                Font = new Font("microsoft yahei ui", 9F, FontStyle.Bold),
                Checked = true
            };

            Label lblClickX = new Label { Text = "點擊 X:", Left = 280, Top = 172, Width = 50, ForeColor = Color.White };
            NumericUpDown numClickX = new NumericUpDown
            {
                Left = 330,
                Top = 169,
                Width = 65,
                Minimum = 0,
                Maximum = 3840,
                Value = 652,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.LightGreen,
                Enabled = true
            };

            Label lblClickY = new Label { Text = "Y:", Left = 400, Top = 172, Width = 20, ForeColor = Color.White };
            NumericUpDown numClickY = new NumericUpDown
            {
                Left = 420,
                Top = 169,
                Width = 65,
                Minimum = 0,
                Maximum = 2160,
                Value = 882,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.LightGreen,
                Enabled = true
            };

            Label lblSitKey = new Label { Text = "坐下：", Left = 495, Top = 172, Width = 45, ForeColor = Color.White };
            TextBox txtSitKey = new TextBox
            {
                Left = 540,
                Top = 169,
                Width = 60,
                ReadOnly = true,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.Cyan,
                Text = "(設定)",
                Enabled = true,
                Tag = Keys.None
            };

            Label lblSitDelay = new Label { Text = "延遲：", Left = 20, Top = 200, Width = 45, ForeColor = Color.White };
            NumericUpDown numSitDelay = new NumericUpDown
            {
                Left = 65,
                Top = 197,
                Width = 60,
                Minimum = 0,
                Maximum = 30,
                Value = 3,
                DecimalPlaces = 1,
                Increment = 0.5m,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White,
                Enabled = true
            };
            Label lblSitDelayUnit = new Label
            {
                Text = "秒後坐下 | 序列：停止腳本 → 冷卻2s → 點擊(X,Y) → Enter → 坐下",
                Left = 130,
                Top = 200,
                Width = 470,
                ForeColor = Color.Gray,
                Font = new Font("microsoft yahei ui", 8.5F)
            };

            // 回程勾選控制啟用/停用
            chkReturnToTown.CheckedChanged += (s, args) =>
            {
                bool enabled = chkReturnToTown.Checked;
                numClickX.Enabled = enabled;
                numClickY.Enabled = enabled;
                txtSitKey.Enabled = enabled;
                numSitDelay.Enabled = enabled;
            };

            // 坐下按鍵捕獲（單鍵，含 Ctrl/Shift/Alt/導航鍵）
            txtSitKey.PreviewKeyDown += (s, args) => { args.IsInputKey = true; };
            txtSitKey.KeyDown += (s, args) =>
            {
                args.SuppressKeyPress = true;
                args.Handled = true;

                txtSitKey.Text = GetKeyDisplayName(args.KeyCode);
                txtSitKey.Tag = args.KeyCode;
            };

            // 新增按鈕
            Button btnAddTask = new Button
            {
                Text = "➕ 新增排程",
                Left = 20,
                Top = 230,
                Width = 120,
                Height = 30,
                BackColor = Color.FromArgb(0, 150, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            // ===== 排程清單 =====
            Label lblListTitle = new Label
            {
                Text = "📅 排程清單",
                Left = 20,
                Top = 270,
                Width = 200,
                ForeColor = Color.Yellow,
                Font = new Font("microsoft yahei ui", 10F, FontStyle.Bold)
            };

            DataGridView dgv = new DataGridView
            {
                Left = 20,
                Top = 295,
                Width = 590,
                Height = 220,
                BackgroundColor = Color.FromArgb(30, 30, 35),
                ForeColor = Color.White,
                GridColor = Color.FromArgb(60, 60, 65),
                BorderStyle = BorderStyle.FixedSingle,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(50, 50, 55),
                    ForeColor = Color.White
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(40, 40, 45),
                    ForeColor = Color.White,
                    SelectionBackColor = Color.FromArgb(0, 100, 180)
                }
            };

            dgv.Columns.Add("Script", "腳本");
            dgv.Columns.Add("StartTime", "開始時間");
            dgv.Columns.Add("EndTime", "結束時間");
            dgv.Columns.Add("Loop", "循環");
            dgv.Columns.Add("Return", "回程");
            dgv.Columns.Add("Status", "狀態");

            dgv.Columns["Script"].FillWeight = 25;
            dgv.Columns["StartTime"].FillWeight = 20;
            dgv.Columns["EndTime"].FillWeight = 20;
            dgv.Columns["Loop"].FillWeight = 10;
            dgv.Columns["Return"].FillWeight = 10;
            dgv.Columns["Status"].FillWeight = 15;

            Action refreshTaskList = () =>
            {
                dgv.Rows.Clear();
                foreach (var task in scheduleTasks)
                {
                    string scriptName = string.IsNullOrEmpty(task.ScriptPath) ? "(當前腳本)" : Path.GetFileName(task.ScriptPath);
                    string endTimeStr = task.EndTime.HasValue ? task.EndTime.Value.ToString("HH:mm:ss") : "不限";
                    string status = task.HasStarted ? "已觸發" : (task.Enabled ? "等待中" : "已完成");
                    string returnStr = task.ReturnToTownEnabled ? "✔" : "";
                    dgv.Rows.Add(scriptName, task.StartTime.ToString("HH:mm:ss"), endTimeStr, task.LoopCount, returnStr, status);
                }
            };
            refreshTaskList();

            btnAddTask.Click += (s, args) =>
            {
                if (dtpStart.Value <= DateTime.Now)
                {
                    MessageBox.Show("開始時間必須為未來！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (chkEndTime.Checked && dtpEnd.Value <= dtpStart.Value)
                {
                    MessageBox.Show("結束時間必須晚於開始時間！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string? scriptPath = txtScriptPath.Tag as string;
                if (string.IsNullOrEmpty(scriptPath) && recordedEvents.Count == 0)
                {
                    MessageBox.Show("請選擇腳本或先載入腳本！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var newTask = new ScheduleTask
                {
                    ScriptPath = scriptPath ?? string.Empty,
                    StartTime = dtpStart.Value,
                    EndTime = chkEndTime.Checked ? dtpEnd.Value : null,
                    LoopCount = (int)numLoop.Value,
                    Enabled = true,
                    HasStarted = false,
                    ReturnToTownEnabled = chkReturnToTown.Checked,
                    ReturnClickX = (int)numClickX.Value,
                    ReturnClickY = (int)numClickY.Value,
                    SitDownDelaySeconds = (double)numSitDelay.Value
                };

                // 設定坐下按鍵
                if (txtSitKey.Tag is Keys sitKey)
                {
                    newTask.SitDownKeyCode = (int)sitKey;
                }

                scheduleTasks.Add(newTask);
                schedulerTimer.Start();
                refreshTaskList();

                string endInfo = chkEndTime.Checked ? $", 結束={dtpEnd.Value:HH:mm:ss}" : "";
                string returnInfo = chkReturnToTown.Checked ? $", 回程=點擊({numClickX.Value},{numClickY.Value})" : "";
                AddLog($"新增排程：{(string.IsNullOrEmpty(scriptPath) ? "當前腳本" : Path.GetFileName(scriptPath))}, 開始={dtpStart.Value:HH:mm:ss}{endInfo}, 循環={numLoop.Value}{returnInfo}");
            };

            // 刪除與清空按鈕
            Button btnRemove = new Button
            {
                Text = "🗑️ 刪除選中",
                Left = 20,
                Top = 525,
                Width = 110,
                Height = 30,
                BackColor = Color.FromArgb(150, 60, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnClearAll = new Button
            {
                Text = "清空全部",
                Left = 140,
                Top = 525,
                Width = 90,
                Height = 30,
                BackColor = Color.FromArgb(120, 80, 40),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnClose = new Button
            {
                Text = "關閉",
                Left = 530,
                Top = 525,
                Width = 80,
                Height = 30,
                BackColor = Color.FromArgb(80, 80, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnRemove.Click += (s, args) =>
            {
                if (dgv.SelectedRows.Count > 0)
                {
                    var indices = dgv.SelectedRows.Cast<DataGridViewRow>().Select(r => r.Index).OrderByDescending(i => i).ToList();
                    foreach (int idx in indices)
                    {
                        if (idx < scheduleTasks.Count)
                            scheduleTasks.RemoveAt(idx);
                    }
                    refreshTaskList();
                    AddLog($"已刪除 {indices.Count} 個排程");
                }
            };

            btnClearAll.Click += (s, args) =>
            {
                scheduleTasks.Clear();
                schedulerTimer.Stop();
                refreshTaskList();
                AddLog("已清空所有排程");
            };

            btnClose.Click += (s, args) => schedForm.Close();

            // 倒數計時
            Label lblCountdown = new Label
            {
                Left = 240,
                Top = 530,
                Width = 280,
                ForeColor = Color.Yellow,
                Font = new Font("microsoft yahei ui", 9F)
            };

            System.Windows.Forms.Timer countdownTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            countdownTimer.Tick += (s, args) =>
            {
                var nextTask = scheduleTasks.Where(t => t.Enabled && !t.HasStarted).OrderBy(t => t.StartTime).FirstOrDefault();
                if (nextTask != null)
                {
                    TimeSpan remaining = nextTask.StartTime - DateTime.Now;
                    if (remaining.TotalSeconds > 0)
                        lblCountdown.Text = $"下個排程：{remaining:hh\\:mm\\:ss} 後開始";
                    else
                        lblCountdown.Text = "正在觸發...";
                }
                else
                {
                    var activeTask = scheduleTasks.Where(t => t.HasStarted && t.Enabled && t.EndTime.HasValue).FirstOrDefault();
                    if (activeTask != null)
                    {
                        TimeSpan remaining = activeTask.EndTime!.Value - DateTime.Now;
                        if (remaining.TotalSeconds > 0)
                            lblCountdown.Text = $"自動停止：{remaining:hh\\:mm\\:ss} 後";
                        else
                            lblCountdown.Text = "正在停止...";
                    }
                    else
                    {
                        lblCountdown.Text = scheduleTasks.Count > 0 ? "所有排程已完成" : "";
                    }
                }
                refreshTaskList();
            };
            countdownTimer.Start();

            schedForm.FormClosing += (s, args) => countdownTimer.Stop();

            schedForm.Controls.AddRange(new Control[]
            {
                lblTitle, lblNewTask,
                lblScript, txtScriptPath, btnBrowse, btnUseCurrent,
                lblStart, dtpStart, lblEnd, chkEndTime, dtpEnd,
                lblLoop, numLoop, lblLoopHint,
                chkReturnToTown, lblClickX, numClickX, lblClickY, numClickY, lblSitKey, txtSitKey,
                lblSitDelay, numSitDelay, lblSitDelayUnit,
                btnAddTask,
                lblListTitle, dgv,
                btnRemove, btnClearAll, btnClose, lblCountdown
            });

            schedForm.ShowDialog();
        }

        /// <summary>
        /// 顯示即時執行統計視窗
        /// </summary>
        private void ShowStatistics()
        {
            Form statsForm = new Form
            {
                Text = "📊 即時執行統計",
                Width = 450,
                Height = 520,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(30, 30, 35)
            };

            // 狀態指示燈
            Label lblStatusIndicator = new Label
            {
                Left = 20,
                Top = 20,
                Width = 400,
                Height = 25,
                ForeColor = statistics.CurrentSessionStart.HasValue ? Color.Lime : Color.Gray,
                Font = new Font("microsoft yahei ui", 12F, FontStyle.Bold),
                Text = statistics.CurrentSessionStart.HasValue ? "● 播放中" : "○ 已停止"
            };

            // 當前會話區塊
            Label lblSessionTitle = new Label
            {
                Left = 20,
                Top = 55,
                Width = 200,
                Height = 20,
                ForeColor = Color.Cyan,
                Font = new Font("microsoft yahei ui", 10F, FontStyle.Bold),
                Text = "📌 當前會話"
            };

            Label lblSessionTime = new Label
            {
                Left = 30,
                Top = 80,
                Width = 380,
                Height = 20,
                ForeColor = Color.White,
                Font = new Font("Consolas", 10F),
                Text = "會話時長: --:--:--"
            };

            Label lblCurrentLoop = new Label
            {
                Left = 30,
                Top = 105,
                Width = 380,
                Height = 20,
                ForeColor = Color.White,
                Font = new Font("Consolas", 10F),
                Text = "當前循環: 0"
            };

            Label lblScriptInfo = new Label
            {
                Left = 30,
                Top = 130,
                Width = 380,
                Height = 20,
                ForeColor = Color.LightGray,
                Font = new Font("microsoft yahei ui", 9F),
                Text = $"腳本事件: {recordedEvents.Count} 個"
            };

            // 分隔線
            Label lblSep1 = new Label
            {
                Left = 20,
                Top = 160,
                Width = 400,
                Height = 2,
                BackColor = Color.FromArgb(60, 60, 65)
            };

            // 累計統計區塊
            Label lblTotalTitle = new Label
            {
                Left = 20,
                Top = 170,
                Width = 200,
                Height = 20,
                ForeColor = Color.Yellow,
                Font = new Font("microsoft yahei ui", 10F, FontStyle.Bold),
                Text = "📈 累計統計"
            };

            Label lblTotalPlays = new Label
            {
                Left = 30,
                Top = 195,
                Width = 380,
                Height = 20,
                ForeColor = Color.White,
                Font = new Font("Consolas", 10F),
                Text = $"播放次數: {statistics.TotalPlayCount}"
            };

            Label lblTotalTime = new Label
            {
                Left = 30,
                Top = 220,
                Width = 380,
                Height = 20,
                ForeColor = Color.White,
                Font = new Font("Consolas", 10F),
                Text = "總播放時長: 00:00:00"
            };

            Label lblLastPlay = new Label
            {
                Left = 30,
                Top = 245,
                Width = 380,
                Height = 20,
                ForeColor = Color.LightGray,
                Font = new Font("microsoft yahei ui", 9F),
                Text = $"最後播放: {(statistics.LastPlayTime.HasValue ? statistics.LastPlayTime.Value.ToString("yyyy-MM-dd HH:mm:ss") : "從未播放")}"
            };

            // 分隔線
            Label lblSep2 = new Label
            {
                Left = 20,
                Top = 275,
                Width = 400,
                Height = 2,
                BackColor = Color.FromArgb(60, 60, 65)
            };

            // 自定義按鍵統計
            Label lblCustomTitle = new Label
            {
                Left = 20,
                Top = 285,
                Width = 200,
                Height = 20,
                ForeColor = Color.FromArgb(200, 150, 255),
                Font = new Font("microsoft yahei ui", 10F, FontStyle.Bold),
                Text = "⚡ 自定義按鍵觸發"
            };

            ListBox lstCustomKeys = new ListBox
            {
                Left = 30,
                Top = 310,
                Width = 380,
                Height = 100,
                BackColor = Color.FromArgb(40, 40, 45),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 9F)
            };

            // 填入自定義按鍵統計
            Action updateCustomKeyList = () =>
            {
                lstCustomKeys.Items.Clear();
                bool hasAny = false;
                for (int i = 0; i < 15; i++)
                {
                    if (customKeySlots[i].Enabled && customKeySlots[i].KeyCode != Keys.None)
                    {
                        hasAny = true;
                        lstCustomKeys.Items.Add($"  #{i + 1} {GetKeyDisplayName(customKeySlots[i].KeyCode),-15} 觸發: {statistics.CustomKeyTriggerCounts[i]} 次");
                    }
                }
                if (!hasAny)
                {
                    lstCustomKeys.Items.Add("  (無啟用的自定義按鍵)");
                }
            };
            updateCustomKeyList();

            // 按鈕
            Button btnReset = new Button
            {
                Text = "🔄 重置統計",
                Left = 130,
                Top = 430,
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(150, 80, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnClose = new Button
            {
                Text = "關閉",
                Left = 240,
                Top = 430,
                Width = 80,
                Height = 35,
                BackColor = Color.FromArgb(80, 80, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnReset.Click += (s, args) =>
            {
                if (MessageBox.Show("確定重置所有統計資料？", "重置統計", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    statistics.Reset();
                    AddLog("統計資料已重置");
                }
            };

            btnClose.Click += (s, args) => statsForm.Close();

            // 即時更新 Timer
            System.Windows.Forms.Timer refreshTimer = new System.Windows.Forms.Timer { Interval = 500 };
            refreshTimer.Tick += (s, args) =>
            {
                bool isActive = statistics.CurrentSessionStart.HasValue;

                // 狀態指示
                lblStatusIndicator.Text = isActive ? "● 播放中" : "○ 已停止";
                lblStatusIndicator.ForeColor = isActive ? Color.Lime : Color.Gray;

                // 當前會話
                if (isActive)
                {
                    TimeSpan sessionTime = DateTime.Now - statistics.CurrentSessionStart!.Value;
                    lblSessionTime.Text = $"會話時長: {(int)sessionTime.TotalHours:D2}:{sessionTime.Minutes:D2}:{sessionTime.Seconds:D2}";
                }
                else
                {
                    lblSessionTime.Text = "會話時長: --:--:--";
                }
                lblCurrentLoop.Text = $"當前循環: {statistics.CurrentLoopCount}";
                lblScriptInfo.Text = $"腳本事件: {recordedEvents.Count} 個";

                // 累計統計（加上當前會話的時間）
                double liveTotalSeconds = statistics.TotalPlayTimeSeconds;
                if (isActive)
                {
                    liveTotalSeconds += (DateTime.Now - statistics.CurrentSessionStart!.Value).TotalSeconds;
                }
                TimeSpan totalTime = TimeSpan.FromSeconds(liveTotalSeconds);
                lblTotalPlays.Text = $"播放次數: {statistics.TotalPlayCount + (isActive ? 1 : 0)}";
                lblTotalTime.Text = $"總播放時長: {(int)totalTime.TotalHours:D2}:{totalTime.Minutes:D2}:{totalTime.Seconds:D2}";
                lblLastPlay.Text = $"最後播放: {(statistics.LastPlayTime.HasValue ? statistics.LastPlayTime.Value.ToString("yyyy-MM-dd HH:mm:ss") : (isActive ? "進行中..." : "從未播放"))}";

                // 更新自定義按鍵統計
                updateCustomKeyList();
            };
            refreshTimer.Start();

            statsForm.FormClosing += (s, args) => refreshTimer.Stop();

            statsForm.Controls.AddRange(new Control[]
            {
                lblStatusIndicator, lblSessionTitle, lblSessionTime, lblCurrentLoop, lblScriptInfo,
                lblSep1, lblTotalTitle, lblTotalPlays, lblTotalTime, lblLastPlay,
                lblSep2, lblCustomTitle, lstCustomKeys,
                btnReset, btnClose
            });

            statsForm.ShowDialog();
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

            string[] stateNames = { "左邊界", "右邊界", "上邊界", "下邊界" };

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
                case 2: bT = cy; AddLog($"✅ 上邊界 = {cy}"); break;
                case 3: bB = cy; AddLog($"✅ 下邊界 = {cy}"); break;
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

        /// <summary>
        /// 小地圖校準精靈
        /// </summary>
        /// <summary>
        /// 小地圖校準精靈
        /// </summary>
        private void OpenMinimapCalibration()
        {
            if (targetWindowHandle == IntPtr.Zero || !IsWindow(targetWindowHandle))
            {
                MessageBox.Show("請先鎖定目標視窗！\n\n步驟：\n1. 開啟遊戲\n2. 點擊「手動鎖定」選擇遊戲視窗",
                    "未鎖定視窗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 初始化追蹤器
            if (minimapTracker == null)
                minimapTracker = new MinimapTracker();

            minimapTracker.AttachToWindow(targetWindowHandle);

            // 嘗試載入已有的校準
            bool hasCalibration = minimapTracker.LoadCalibration(CalibrationFilePath);

            Form calibForm = new Form
            {
                Text = "🗺️ 小地圖校準精靈",
                Width = 750,
                Height = 580,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(35, 35, 40)
            };

            // ===== 左側：小地圖預覽區 =====
            GroupBox grpPreview = new GroupBox
            {
                Text = "小地圖預覽 (點擊角色圖標學習顏色)",
                Left = 15, Top = 15, Width = 350, Height = 300,
                ForeColor = Color.White,
                Font = new Font("microsoft yahei ui", 9F)
            };

            PictureBox picMinimap = new PictureBox
            {
                Left = 10, Top = 25, Width = 330, Height = 220,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black,
                SizeMode = PictureBoxSizeMode.Zoom,
                Cursor = Cursors.Cross
            };

            Label lblMinimapInfo = new Label
            {
                Left = 10, Top = 250, Width = 330, Height = 20,
                ForeColor = Color.LightGray,
                Font = new Font("microsoft yahei ui", 8.5F),
                Text = "區域: 未設定"
            };

            Label lblColorInfo = new Label
            {
                Left = 10, Top = 272, Width = 330, Height = 20,
                ForeColor = Color.Yellow,
                Font = new Font("microsoft yahei ui", 8.5F),
                Text = "💡 點擊預覽圖上的角色圖標來學習顏色"
            };

            grpPreview.Controls.AddRange(new Control[] { picMinimap, lblMinimapInfo, lblColorInfo });

            // ===== 右上：小地圖區域設定 =====
            GroupBox grpRegion = new GroupBox
            {
                Text = "📍 小地圖區域",
                Left = 380, Top = 15, Width = 345, Height = 100,
                ForeColor = Color.White,
                Font = new Font("microsoft yahei ui", 9F)
            };

            Label lblX = new Label { Text = "X:", Left = 15, Top = 28, Width = 25, ForeColor = Color.LightGray };
            NumericUpDown numX = new NumericUpDown { Left = 40, Top = 25, Width = 60, Minimum = 0, Maximum = 2000, Value = minimapTracker.MinimapRegion.X };
            Label lblY = new Label { Text = "Y:", Left = 110, Top = 28, Width = 25, ForeColor = Color.LightGray };
            NumericUpDown numY = new NumericUpDown { Left = 135, Top = 25, Width = 60, Minimum = 0, Maximum = 2000, Value = minimapTracker.MinimapRegion.Y };
            Label lblW = new Label { Text = "寬:", Left = 205, Top = 28, Width = 30, ForeColor = Color.LightGray };
            NumericUpDown numW = new NumericUpDown { Left = 235, Top = 25, Width = 55, Minimum = 10, Maximum = 500, Value = Math.Max(10, minimapTracker.MinimapRegion.Width) };
            Label lblH = new Label { Text = "高:", Left = 295, Top = 28, Width = 30, ForeColor = Color.LightGray };
            NumericUpDown numH = new NumericUpDown { Left = 235, Top = 55, Width = 55, Minimum = 10, Maximum = 500, Value = Math.Max(10, minimapTracker.MinimapRegion.Height) };

            Button btnSelectRegion = new Button
            {
                Text = "🖱️ 框選區域",
                Left = 15, Top = 58, Width = 100, Height = 28,
                BackColor = Color.FromArgb(0, 140, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("microsoft yahei ui", 9F, FontStyle.Bold)
            };

            Button btnCapture = new Button
            {
                Text = "📷 截取預覽",
                Left = 125, Top = 58, Width = 100, Height = 28,
                BackColor = Color.FromArgb(80, 80, 90),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            grpRegion.Controls.AddRange(new Control[] { lblX, numX, lblY, numY, lblW, numW, lblH, numH, btnSelectRegion, btnCapture });

            // ===== 右中：活動範圍邊界 =====
            GroupBox grpBounds = new GroupBox
            {
                Text = "🎯 活動範圍邊界 (移動角色後點擊設定)",
                Left = 380, Top = 120, Width = 345, Height = 130,
                ForeColor = Color.White,
                Font = new Font("microsoft yahei ui", 9F)
            };

            Label lblBoundLeft = new Label { Text = "左: ---", Left = 15, Top = 25, Width = 70, ForeColor = Color.LightGray };
            Label lblBoundRight = new Label { Text = "右: ---", Left = 95, Top = 25, Width = 70, ForeColor = Color.LightGray };
            Label lblBoundTop = new Label { Text = "上: ---", Left = 175, Top = 25, Width = 70, ForeColor = Color.LightGray };
            Label lblBoundBottom = new Label { Text = "下: ---", Left = 255, Top = 25, Width = 70, ForeColor = Color.LightGray };

            Button btnSetLeft = new Button { Text = "⬅ 左", Left = 15, Top = 50, Width = 75, Height = 28, BackColor = Color.FromArgb(60, 60, 70), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Button btnSetRight = new Button { Text = "右 ➡", Left = 95, Top = 50, Width = 75, Height = 28, BackColor = Color.FromArgb(60, 60, 70), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Button btnSetTop = new Button { Text = "⬆ 上", Left = 175, Top = 50, Width = 75, Height = 28, BackColor = Color.FromArgb(60, 60, 70), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Button btnSetBottom = new Button { Text = "下 ⬇", Left = 255, Top = 50, Width = 75, Height = 28, BackColor = Color.FromArgb(60, 60, 70), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

            Button btnResetBounds = new Button { Text = "🔄 重設", Left = 15, Top = 85, Width = 75, Height = 28, BackColor = Color.FromArgb(100, 60, 60), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

            Label lblBoundTip = new Label { Text = "F7: 依序設定左→右→上→下", Left = 100, Top = 92, Width = 220, ForeColor = Color.Cyan, Font = new Font("microsoft yahei ui", 8F) };

            grpBounds.Controls.AddRange(new Control[] { lblBoundLeft, lblBoundRight, lblBoundTop, lblBoundBottom, btnSetLeft, btnSetRight, btnSetTop, btnSetBottom, btnResetBounds, lblBoundTip });

            // ===== 右下：顏色設定 =====
            GroupBox grpColor = new GroupBox
            {
                Text = "🎨 角色圖標顏色 (點擊預覽圖學習)",
                Left = 380, Top = 255, Width = 345, Height = 60,
                ForeColor = Color.White,
                Font = new Font("microsoft yahei ui", 9F)
            };

            Label lblHue = new Label { Text = "色相:", Left = 15, Top = 25, Width = 40, ForeColor = Color.LightGray };
            NumericUpDown numHueMin = new NumericUpDown { Left = 55, Top = 22, Width = 50, Minimum = 0, Maximum = 360, Value = (decimal)minimapTracker.HueRange.Min };
            Label lblHueTo = new Label { Text = "~", Left = 108, Top = 25, Width = 15, ForeColor = Color.LightGray };
            NumericUpDown numHueMax = new NumericUpDown { Left = 123, Top = 22, Width = 50, Minimum = 0, Maximum = 360, Value = (decimal)minimapTracker.HueRange.Max };

            Label lblSat = new Label { Text = "飽和:", Left = 180, Top = 25, Width = 40, ForeColor = Color.LightGray };
            NumericUpDown numSat = new NumericUpDown { Left = 220, Top = 22, Width = 50, Minimum = 0, Maximum = 100, Value = (decimal)(minimapTracker.MinSaturation * 100), DecimalPlaces = 0 };
            Label lblSatPct = new Label { Text = "%", Left = 272, Top = 25, Width = 20, ForeColor = Color.LightGray };

            Panel pnlColorPreview = new Panel { Left = 300, Top = 20, Width = 30, Height = 30, BorderStyle = BorderStyle.FixedSingle };

            grpColor.Controls.AddRange(new Control[] { lblHue, numHueMin, lblHueTo, numHueMax, lblSat, numSat, lblSatPct, pnlColorPreview });

            // ===== 底部：偵測結果 =====
            GroupBox grpResult = new GroupBox
            {
                Text = "📊 偵測結果",
                Left = 15, Top = 320, Width = 710, Height = 100,
                ForeColor = Color.White,
                Font = new Font("microsoft yahei ui", 9F)
            };

            Label lblDetectResult = new Label
            {
                Left = 15, Top = 25, Width = 200, Height = 65,
                ForeColor = Color.Gray,
                Font = new Font("Consolas", 11F, FontStyle.Bold),
                Text = "尚未偵測"
            };

            PictureBox picDebug = new PictureBox
            {
                Left = 230, Top = 18, Width = 250, Height = 72,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black,
                SizeMode = PictureBoxSizeMode.Zoom
            };

            Button btnTest = new Button
            {
                Text = "🎯 測試偵測",
                Left = 500, Top = 25, Width = 100, Height = 30,
                BackColor = Color.FromArgb(180, 120, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnAutoTest = new Button
            {
                Text = "🔄 連續測試",
                Left = 610, Top = 25, Width = 85, Height = 30,
                BackColor = Color.FromArgb(80, 80, 90),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            grpResult.Controls.AddRange(new Control[] { lblDetectResult, picDebug, btnTest, btnAutoTest });

            // ===== 最底部按鈕 =====
            Button btnCorrection = new Button
            {
                Text = "🎯 位置修正",
                Left = 340, Top = 435, Width = 120, Height = 35,
                BackColor = Color.FromArgb(180, 120, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("microsoft yahei ui", 10F, FontStyle.Bold)
            };

            Button btnSave = new Button
            {
                Text = "💾 儲存校準",
                Left = 470, Top = 435, Width = 120, Height = 35,
                BackColor = Color.FromArgb(0, 140, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("microsoft yahei ui", 10F, FontStyle.Bold)
            };

            Button btnClose = new Button
            {
                Text = "關閉",
                Left = 605, Top = 435, Width = 80, Height = 35,
                BackColor = Color.FromArgb(80, 80, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Label lblStatus = new Label
            {
                Left = 15, Top = 445, Width = 440, Height = 25,
                ForeColor = hasCalibration ? Color.Lime : Color.Gray,
                Font = new Font("microsoft yahei ui", 9F),
                Text = hasCalibration ? "✅ 已載入先前的校準設定" : "⚠️ 尚未校準"
            };

            // ===== 變數 =====
            int boundLeft = 0, boundRight = 100, boundTop = 0, boundBottom = 100;
            bool boundsExplicitlySet = false;
            System.Windows.Forms.Timer? autoTestTimer = null;
            Bitmap? currentMinimapBitmap = null;

            if (hasCalibration && minimapTracker.MapBounds.Width > 0)
            {
                boundLeft = minimapTracker.MapBounds.Left;
                boundRight = minimapTracker.MapBounds.Right;
                boundTop = minimapTracker.MapBounds.Top;
                boundBottom = minimapTracker.MapBounds.Bottom;
                boundsExplicitlySet = true;
                lblBoundLeft.Text = $"左: {boundLeft}";
                lblBoundRight.Text = $"右: {boundRight}";
                lblBoundTop.Text = $"上: {boundTop}";
                lblBoundBottom.Text = $"下: {boundBottom}";
            }

            // ★ F7 邊界同步：HandleBoundaryHotkey 更新 MapBounds 時自動同步到對話框
            System.Windows.Forms.Timer boundarySyncTimer = new System.Windows.Forms.Timer { Interval = 300 };
            boundarySyncTimer.Tick += (s, e) =>
            {
                if (minimapTracker != null && minimapTracker.MapBounds.Width > 0)
                {
                    var mb = minimapTracker.MapBounds;
                    int mbRight = mb.Left + mb.Width;
                    int mbBottom = mb.Top + mb.Height;
                    if (mb.Left != boundLeft || mbRight != boundRight || mb.Top != boundTop || mbBottom != boundBottom)
                    {
                        boundLeft = mb.Left; boundRight = mbRight;
                        boundTop = mb.Top; boundBottom = mbBottom;
                        boundsExplicitlySet = true;
                        lblBoundLeft.Text = $"左: {boundLeft}"; lblBoundRight.Text = $"右: {boundRight}";
                        lblBoundTop.Text = $"上: {boundTop}"; lblBoundBottom.Text = $"下: {boundBottom}";
                        lblBoundLeft.ForeColor = lblBoundRight.ForeColor = lblBoundTop.ForeColor = lblBoundBottom.ForeColor = Color.Lime;
                    }
                }
                // 顯示 F7 設定進度
                if (_boundarySetState >= 0 && _boundarySetState < 4)
                {
                    string[] nextNames = { "左邊界", "右邊界", "上邊界", "下邊界" };
                    lblBoundTip.Text = $"F7: 等待設定【{nextNames[_boundarySetState]}】...";
                    lblBoundTip.ForeColor = Color.Yellow;
                }
                else
                {
                    lblBoundTip.Text = "F7: 依序設定左→右→上→下";
                    lblBoundTip.ForeColor = Color.Cyan;
                }
            };
            boundarySyncTimer.Start();

            // 更新顏色預覽
            Action updateColorPreview = () =>
            {
                float hue = ((float)numHueMin.Value + (float)numHueMax.Value) / 2f;
                pnlColorPreview.BackColor = ColorFromHSV(hue, 0.8, 0.9);
            };
            updateColorPreview();

            // ===== 輔助函數 =====
            Action updateRegion = () =>
            {
                minimapTracker.MinimapRegion = new Rectangle((int)numX.Value, (int)numY.Value, (int)numW.Value, (int)numH.Value);
                lblMinimapInfo.Text = $"區域: X={numX.Value}, Y={numY.Value}, {numW.Value}x{numH.Value}";
            };

            Action updateMapBounds = () =>
            {
                minimapTracker.MapBounds = new Rectangle(boundLeft, boundTop, boundRight - boundLeft, boundBottom - boundTop);
            };

            Action updateColorSettings = () =>
            {
                minimapTracker.HueRange = ((float)numHueMin.Value, (float)numHueMax.Value);
                minimapTracker.MinSaturation = (float)numSat.Value / 100f;
                updateColorPreview();
            };

            Func<(int x, int y, bool success)> detectCurrentPixelPos = () =>
            {
                updateRegion();
                updateColorSettings();
                return minimapTracker.ReadPosition();
            };

            // ===== 事件處理 =====
            numX.ValueChanged += (s, e) => updateRegion();
            numY.ValueChanged += (s, e) => updateRegion();
            numW.ValueChanged += (s, e) => updateRegion();
            numH.ValueChanged += (s, e) => updateRegion();
            numHueMin.ValueChanged += (s, e) => updateColorSettings();
            numHueMax.ValueChanged += (s, e) => updateColorSettings();
            numSat.ValueChanged += (s, e) => updateColorSettings();

            // 點擊預覽圖學習顏色
            picMinimap.MouseClick += (senderObj, e) =>
            {
                if (currentMinimapBitmap == null) return;

                // 計算實際點擊位置（考慮縮放）
                float scaleX = (float)currentMinimapBitmap.Width / picMinimap.Width;
                float scaleY = (float)currentMinimapBitmap.Height / picMinimap.Height;
                float scale = Math.Max(scaleX, scaleY);
                
                int offsetX = (int)((picMinimap.Width - currentMinimapBitmap.Width / scale) / 2);
                int offsetY = (int)((picMinimap.Height - currentMinimapBitmap.Height / scale) / 2);
                
                int imgX = (int)((e.X - offsetX) * scale);
                int imgY = (int)((e.Y - offsetY) * scale);

                if (imgX < 0 || imgX >= currentMinimapBitmap.Width || imgY < 0 || imgY >= currentMinimapBitmap.Height)
                    return;

                // 取得點擊位置的顏色
                Color clickedColor = currentMinimapBitmap.GetPixel(imgX, imgY);
                
                // 轉換為 HSV
                float h, s, v;
                ColorToHSV(clickedColor, out h, out s, out v);

                // 設定色相範圍 (±15度)
                numHueMin.Value = Math.Max(0, (decimal)(h - 15));
                numHueMax.Value = Math.Min(360, (decimal)(h + 15));
                numSat.Value = Math.Max(20, (decimal)(s * 100 - 20)); // 飽和度下限

                lblColorInfo.Text = $"✅ 已學習顏色: RGB({clickedColor.R},{clickedColor.G},{clickedColor.B}) H={h:F0}°";
                lblColorInfo.ForeColor = Color.Lime;
                lblStatus.Text = $"✅ 已學習顏色 (色相: {h:F0}°)";
                lblStatus.ForeColor = Color.Lime;
            };

            btnSelectRegion.Click += (s, e) =>
            {
                using (var screenshot = minimapTracker.CaptureFullWindow())
                {
                    if (screenshot == null)
                    {
                        MessageBox.Show("無法截取遊戲視窗！", "截取失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    calibForm.Hide();
                    var region = RegionSelector.SelectRegion(screenshot);
                    calibForm.Show();

                    if (region.HasValue)
                    {
                        numX.Value = Math.Min(numX.Maximum, Math.Max(numX.Minimum, region.Value.X));
                        numY.Value = Math.Min(numY.Maximum, Math.Max(numY.Minimum, region.Value.Y));
                        numW.Value = Math.Min(numW.Maximum, Math.Max(numW.Minimum, region.Value.Width));
                        numH.Value = Math.Min(numH.Maximum, Math.Max(numH.Minimum, region.Value.Height));
                        updateRegion();

                        // 自動截取預覽
                        using (var bmp = minimapTracker.CaptureMinimap())
                        {
                            if (bmp != null)
                            {
                                currentMinimapBitmap?.Dispose();
                                currentMinimapBitmap = new Bitmap(bmp);
                                picMinimap.Image?.Dispose();
                                picMinimap.Image = new Bitmap(bmp);
                            }
                        }
                        lblStatus.Text = "✅ 區域已框選！點擊預覽圖上的角色學習顏色";
                        lblStatus.ForeColor = Color.Lime;
                    }
                }
            };

            btnCapture.Click += (s, e) =>
            {
                updateRegion();
                using (var bmp = minimapTracker.CaptureMinimap())
                {
                    if (bmp != null)
                    {
                        currentMinimapBitmap?.Dispose();
                        currentMinimapBitmap = new Bitmap(bmp);
                        picMinimap.Image?.Dispose();
                        picMinimap.Image = new Bitmap(bmp);
                    }
                    else
                    {
                        MessageBox.Show("截取失敗！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            };

            // 邊界設定按鈕
            Action<string, Action<int>, Label> setBoundary = (name, setter, label) =>
            {
                var (x, y, success) = detectCurrentPixelPos();
                if (success)
                {
                    int val = name.Contains("左") || name.Contains("右") ? x : y;
                    setter(val);
                    label.Text = $"{name.Substring(0,1)}: {val}";
                    label.ForeColor = Color.Lime;
                    boundsExplicitlySet = true;
                    updateMapBounds();
                    lblStatus.Text = $"✅ {name}邊界已設定: {val}";
                    lblStatus.ForeColor = Color.Lime;
                }
                else
                {
                    MessageBox.Show("偵測失敗！請先點擊預覽圖學習角色圖標顏色", "偵測失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            btnSetLeft.Click += (s, e) => setBoundary("左", v => boundLeft = v, lblBoundLeft);
            btnSetRight.Click += (s, e) => setBoundary("右", v => boundRight = v, lblBoundRight);
            btnSetTop.Click += (s, e) => setBoundary("上", v => boundTop = v, lblBoundTop);
            btnSetBottom.Click += (s, e) => setBoundary("下", v => boundBottom = v, lblBoundBottom);

            btnResetBounds.Click += (s, e) =>
            {
                boundLeft = 0; boundRight = (int)numW.Value; boundTop = 0; boundBottom = (int)numH.Value;
                boundsExplicitlySet = false;
                lblBoundLeft.Text = "左: ---"; lblBoundRight.Text = "右: ---";
                lblBoundTop.Text = "上: ---"; lblBoundBottom.Text = "下: ---";
                lblBoundLeft.ForeColor = lblBoundRight.ForeColor = lblBoundTop.ForeColor = lblBoundBottom.ForeColor = Color.LightGray;
                updateMapBounds();
            };

            btnTest.Click += (s, e) =>
            {
                updateRegion();
                updateMapBounds();
                updateColorSettings();

                var (px, py, success) = detectCurrentPixelPos();
                if (success)
                {
                    if (boundsExplicitlySet)
                    {
                        int boundsTolerance = 5; // 邊界判定容差(px)，避免偵測抖動
                        bool inBounds = px >= (boundLeft - boundsTolerance) && px <= (boundRight + boundsTolerance) &&
                                        py >= (boundTop - boundsTolerance) && py <= (boundBottom + boundsTolerance);
                        // 計算到最近邊界的距離（負數=超出）
                        int dL = px - boundLeft, dR = boundRight - px, dT = py - boundTop, dB = boundBottom - py;
                        int minDist = Math.Min(Math.Min(dL, dR), Math.Min(dT, dB));
                        string distInfo = inBounds ? "" : $" (偏移{Math.Abs(minDist)}px)";
                        lblDetectResult.Text = $"位置: ({px}, {py})\n{(inBounds ? "✅ 在範圍內" : $"⚠️ 超出範圍{distInfo}")}\n信心度: {minimapTracker.DetectionConfidence}%";
                        lblDetectResult.ForeColor = inBounds ? Color.Lime : Color.Orange;
                    }
                    else
                    {
                        // 未設定邊界時只顯示位置和信心度，不做範圍判定
                        lblDetectResult.Text = $"位置: ({px}, {py})\n📍 (未設定活動範圍)\n信心度: {minimapTracker.DetectionConfidence}%";
                        lblDetectResult.ForeColor = Color.Lime;
                    }

                    using (var debugBmp = minimapTracker.CreateDebugImage())
                    {
                        if (debugBmp != null)
                        {
                            picDebug.Image?.Dispose();
                            picDebug.Image = new Bitmap(debugBmp);
                        }
                    }
                }
                else
                {
                    lblDetectResult.Text = "❌ 偵測失敗\n請先學習顏色";
                    lblDetectResult.ForeColor = Color.Red;
                }
            };

            // 連續測試
            btnAutoTest.Click += (s, e) =>
            {
                if (autoTestTimer == null)
                {
                    autoTestTimer = new System.Windows.Forms.Timer { Interval = 200 };
                    autoTestTimer.Tick += (ts, te) => btnTest.PerformClick();
                    autoTestTimer.Start();
                    btnAutoTest.Text = "⏹ 停止";
                    btnAutoTest.BackColor = Color.FromArgb(150, 80, 80);
                }
                else
                {
                    autoTestTimer.Stop();
                    autoTestTimer.Dispose();
                    autoTestTimer = null;
                    btnAutoTest.Text = "🔄 連續測試";
                    btnAutoTest.BackColor = Color.FromArgb(80, 80, 90);
                }
            };

            btnCorrection.Click += (s, e) => { OpenPositionCorrectionSettings(); };

            btnSave.Click += (s, e) =>
            {
                updateRegion();
                updateMapBounds();
                updateColorSettings();

                if (minimapTracker.MinimapRegion.Width <= 0)
                {
                    MessageBox.Show("請先設定小地圖區域！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                minimapTracker.SaveCalibration(CalibrationFilePath);
                lblStatus.Text = "✅ 校準設定已儲存！";
                lblStatus.ForeColor = Color.Lime;
                AddLog("💾 小地圖校準設定已儲存");
            };

            btnClose.Click += (s, e) => calibForm.Close();

            calibForm.FormClosing += (s, e) =>
            {
                autoTestTimer?.Stop();
                autoTestTimer?.Dispose();
                boundarySyncTimer?.Stop();
                boundarySyncTimer?.Dispose();
                currentMinimapBitmap?.Dispose();
                _boundarySetState = -1; // 重置 F7 狀態機
            };

            // 初始化
            updateRegion();
            updateMapBounds();

            calibForm.Controls.AddRange(new Control[] { grpPreview, grpRegion, grpBounds, grpColor, grpResult, btnCorrection, btnSave, btnClose, lblStatus });
            calibForm.ShowDialog();
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
                    try { this.BeginInvoke(new Action(() => AddLog($"[修正] {msg}"))); } catch { }
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
                        this.BeginInvoke(new Action(() =>
                            AddLog($"📍 偏差修正：擷取參考位置 ({medX},{medY})（讀取{readings.Count}次取中位數）")));
                    } catch { }

                    return new CorrectionResult(true, $"擷取參考位置: ({medX},{medY})", medX, medY, 0, 0);
                }
                else
                {
                    try { this.BeginInvoke(new Action(() => AddLog("⚠️ 偵測失敗，無法擷取參考位置"))); } catch { }
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

            this.BeginInvoke(new Action(() =>
            {
                AddLog($"📍 定期檢查: 目前({cx},{cy}) 腳本({nearest.RecordedX},{nearest.RecordedY}) 偏差({dx:+#;-#;0},{dy:+#;-#;0}) 觸發閾值 H±{hTrigger} V±{vTrigger} → 修正中");
                lblPlaybackStatus.Text = $"位置偏差修正中...";
                lblPlaybackStatus.ForeColor = Color.Orange;
            }));

            if (positionCorrector == null)
            {
                positionCorrector = new PositionCorrector();
                positionCorrector.OnLog += (msg) =>
                {
                    try { this.BeginInvoke(new Action(() => AddLog($"[修正] {msg}"))); } catch { }
                };
            }
            ApplyCorrectorSettings(positionCorrector);

            var result = positionCorrector.CorrectPosition(minimapTracker, nearest.RecordedX, nearest.RecordedY);
            if (result != null)
            {
                var r = result;
                this.BeginInvoke(new Action(() =>
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

            return positionCorrector.CorrectPosition(minimapTracker, targetX, targetY);
        }

        /// <summary>
        /// 開啟位置修正設定介面
        /// </summary>
        private void OpenPositionCorrectionSettings()
        {
            Form f = new Form
            {
                Text = "🎯 位置修正設定", Width = 620, Height = 770,
                StartPosition = FormStartPosition.CenterParent, Owner = this,
                FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.FromArgb(35, 35, 40),
                KeyPreview = true
            };

            // ===== 啟用 =====
            CheckBox chkEnabled = new CheckBox
            {
                Text = "啟用位置修正", Left = 15, Top = 10, Width = 200,
                ForeColor = Color.White, Font = new Font("microsoft yahei ui", 10F, FontStyle.Bold),
                Checked = positionCorrectionSettings.Enabled
            };

            // ===== 模式說明 =====
            Label lblModeDesc = new Label
            {
                Text = "📌 即時比對腳本錄製座標，每隔設定秒數檢查並修正偏差（水平/垂直獨立容差）",
                Left = 15, Top = 42, Width = 575, Height = 30,
                ForeColor = Color.FromArgb(120, 200, 255), Font = new Font("microsoft yahei ui", 9F)
            };

            // ===== 按鍵設定 =====
            GroupBox grpK = new GroupBox
            {
                Text = "⌨️ 移動按鍵 (點擊框後按下單鍵或組合鍵)", Left = 15, Top = 72, Width = 575, Height = 88,
                ForeColor = Color.White, Font = new Font("microsoft yahei ui", 9F)
            };
            Keys[] cL = positionCorrectionSettings.GetEffectiveLeftKeys(), cR = positionCorrectionSettings.GetEffectiveRightKeys();
            Keys[] cU = positionCorrectionSettings.GetEffectiveUpKeys(), cD = positionCorrectionSettings.GetEffectiveDownKeys();

            CheckBox chkH = new CheckBox { Text = "水平", Left = 10, Top = 22, Width = 50, ForeColor = Color.LightGray, Checked = positionCorrectionSettings.EnableHorizontalCorrection };
            Label lL = new Label { Text = "左:", Left = 65, Top = 25, Width = 22, ForeColor = Color.LightGray };
            TextBox tL = new TextBox { Left = 87, Top = 22, Width = 100, Text = PositionCorrector.KeysToDisplayString(cL), BackColor = Color.FromArgb(50, 50, 55), ForeColor = Color.White, Tag = cL, Cursor = Cursors.Arrow };
            Label lR = new Label { Text = "右:", Left = 195, Top = 25, Width = 22, ForeColor = Color.LightGray };
            TextBox tR = new TextBox { Left = 217, Top = 22, Width = 100, Text = PositionCorrector.KeysToDisplayString(cR), BackColor = Color.FromArgb(50, 50, 55), ForeColor = Color.White, Tag = cR, Cursor = Cursors.Arrow };

            CheckBox chkV = new CheckBox { Text = "垂直", Left = 10, Top = 50, Width = 50, ForeColor = Color.LightGray, Checked = positionCorrectionSettings.EnableVerticalCorrection };
            Label lU = new Label { Text = "上:", Left = 65, Top = 53, Width = 22, ForeColor = Color.LightGray };
            TextBox tU = new TextBox { Left = 87, Top = 50, Width = 100, Text = PositionCorrector.KeysToDisplayString(cU), BackColor = Color.FromArgb(50, 50, 55), ForeColor = Color.White, Tag = cU, Cursor = Cursors.Arrow };
            Label lD = new Label { Text = "下:", Left = 195, Top = 53, Width = 22, ForeColor = Color.LightGray };
            TextBox tD = new TextBox { Left = 217, Top = 50, Width = 100, Text = PositionCorrector.KeysToDisplayString(cD), BackColor = Color.FromArgb(50, 50, 55), ForeColor = Color.White, Tag = cD, Cursor = Cursors.Arrow };

            CheckBox chkInvY = new CheckBox { Text = "Y軸反轉(Y越大=越高)", Left = 330, Top = 22, Width = 200, ForeColor = Color.FromArgb(255, 200, 100), Checked = positionCorrectionSettings.InvertY };

            // ★ 按鍵輸入（支援單鍵或組合鍵，包含方向鍵、Space、Shift 等）
            TextBox? actKI = null; HashSet<Keys> capK = new HashSet<Keys>(); System.Windows.Forms.Timer? cTmr = null;

            // ★ 重設定時器：在捕獲到新按鍵後延遲 500ms 結束捕獲
            Action resetCaptureTimer = () =>
            {
                if (cTmr != null) { cTmr.Stop(); cTmr.Dispose(); }
                var currentTb = actKI;
                cTmr = new System.Windows.Forms.Timer { Interval = 500 };
                cTmr.Tick += (ts, te) =>
                {
                    cTmr!.Stop(); cTmr.Dispose(); cTmr = null;
                    if (actKI == currentTb && currentTb != null && capK.Count > 0)
                    {
                        currentTb.Tag = capK.ToArray();
                        currentTb.Text = PositionCorrector.KeysToDisplayString(capK.ToArray());
                        currentTb.ForeColor = Color.White;
                        actKI = null;
                    }
                };
                cTmr.Start();
            };

            Action<TextBox> setupCK = (tb) =>
            {
                tb.PreviewKeyDown += (s, e) => { e.IsInputKey = true; };
                tb.KeyPress += (s, e) => { e.Handled = true; };
                tb.Click += (s, e) => { actKI = tb; capK.Clear(); tb.Text = "按下..."; tb.ForeColor = Color.Yellow; };
                tb.Leave += (s, e) => { if (actKI == tb) { actKI = null; if (capK.Count > 0) { tb.Tag = capK.ToArray(); tb.Text = PositionCorrector.KeysToDisplayString(capK.ToArray()); } else tb.Text = tb.Tag is Keys[] ks ? PositionCorrector.KeysToDisplayString(ks) : "未設定"; tb.ForeColor = Color.White; } };
            };
            setupCK(tL); setupCK(tR); setupCK(tU); setupCK(tD);

            // ★ 使用 Form 層級攔截所有按鍵
            f.KeyDown += (s, e) =>
            {
                if (actKI != null)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    capK.Add(e.KeyCode);
                    if (e.Control && e.KeyCode != Keys.ControlKey) capK.Add(Keys.ControlKey);
                    if (e.Shift && e.KeyCode != Keys.ShiftKey) capK.Add(Keys.ShiftKey);
                    if (e.Alt && e.KeyCode != Keys.Menu) capK.Add(Keys.Menu);
                    actKI.Text = PositionCorrector.KeysToDisplayString(capK.ToArray());
                    resetCaptureTimer();
                }
            };

            grpK.Controls.AddRange(new Control[] { chkH, lL, tL, lR, tR, chkV, lU, tU, lD, tD, chkInvY });

            // ===== 修正參數 =====
            int paramY = 170;
            Label lChkInt = new Label { Text = "檢查間隔:", Left = 15, Top = paramY, Width = 65, ForeColor = Color.LightGray };
            NumericUpDown nChkInt = new NumericUpDown { Left = 82, Top = paramY - 3, Width = 40, Minimum = 1, Maximum = 60, Value = positionCorrectionSettings.CorrectionCheckIntervalSec };
            Label lChkInt2 = new Label { Text = "秒", Left = 125, Top = paramY, Width = 20, ForeColor = Color.LightGray };

            Label lHTol = new Label { Text = "水平容差:", Left = 155, Top = paramY, Width = 60, ForeColor = Color.LightGray };
            NumericUpDown nHTol = new NumericUpDown { Left = 218, Top = paramY - 3, Width = 38, Minimum = 1, Maximum = 50, Value = positionCorrectionSettings.HorizontalTolerance };
            Label lVTol = new Label { Text = "垂直:", Left = 262, Top = paramY, Width = 35, ForeColor = Color.LightGray };
            NumericUpDown nVTol = new NumericUpDown { Left = 297, Top = paramY - 3, Width = 38, Minimum = 1, Maximum = 50, Value = positionCorrectionSettings.VerticalTolerance };
            Label lTolPx = new Label { Text = "px", Left = 338, Top = paramY, Width = 18, ForeColor = Color.LightGray };

            Label lTO = new Label { Text = "超時:", Left = 350, Top = paramY, Width = 35, ForeColor = Color.LightGray };
            NumericUpDown nTO = new NumericUpDown { Left = 387, Top = paramY - 3, Width = 60, Minimum = 1000, Maximum = 30000, Value = positionCorrectionSettings.MaxCorrectionTimeMs, Increment = 500 };
            Label lTOms = new Label { Text = "ms", Left = 450, Top = paramY, Width = 20, ForeColor = Color.LightGray };

            int paramY2 = paramY + 32;
            Label lHKeyInt = new Label { Text = "水平間隔:", Left = 15, Top = paramY2, Width = 65, ForeColor = Color.FromArgb(255, 200, 100) };
            NumericUpDown nHKeyMin = new NumericUpDown { Left = 82, Top = paramY2 - 3, Width = 55, Minimum = 100, Maximum = 5000, Value = positionCorrectionSettings.HorizontalKeyIntervalMinMs, Increment = 100 };
            Label lHKeyTil = new Label { Text = "~", Left = 140, Top = paramY2, Width = 12, ForeColor = Color.LightGray };
            NumericUpDown nHKeyMax = new NumericUpDown { Left = 152, Top = paramY2 - 3, Width = 55, Minimum = 100, Maximum = 5000, Value = positionCorrectionSettings.HorizontalKeyIntervalMaxMs, Increment = 100 };
            Label lHKeyMs = new Label { Text = "ms", Left = 210, Top = paramY2, Width = 25, ForeColor = Color.FromArgb(255, 200, 100) };

            int paramY3 = paramY2 + 30;
            Label lVKeyInt = new Label { Text = "垂直間隔:", Left = 15, Top = paramY3, Width = 65, ForeColor = Color.FromArgb(100, 220, 255) };
            NumericUpDown nVKeyMin = new NumericUpDown { Left = 82, Top = paramY3 - 3, Width = 55, Minimum = 50, Maximum = 5000, Value = positionCorrectionSettings.VerticalKeyIntervalMinMs, Increment = 50 };
            Label lVKeyTil = new Label { Text = "~", Left = 140, Top = paramY3, Width = 12, ForeColor = Color.LightGray };
            NumericUpDown nVKeyMax = new NumericUpDown { Left = 152, Top = paramY3 - 3, Width = 55, Minimum = 50, Maximum = 5000, Value = positionCorrectionSettings.VerticalKeyIntervalMaxMs, Increment = 50 };
            Label lVKeyMs = new Label { Text = "ms (垂直較短)", Left = 210, Top = paramY3, Width = 95, ForeColor = Color.FromArgb(100, 220, 255) };

            Label lMaxCorr = new Label { Text = "步數上限:", Left = 320, Top = paramY2, Width = 60, ForeColor = Color.LightGray };
            NumericUpDown nMaxCorr = new NumericUpDown { Left = 382, Top = paramY2 - 3, Width = 50, Minimum = 0, Maximum = 999, Value = positionCorrectionSettings.MaxStepsPerCorrection };
            ToolTip tt = new ToolTip();
            tt.SetToolTip(nMaxCorr, "0=無限制，>0=每次修正最多按鍵N次（達到後停止本次修正）");
            tt.SetToolTip(nHKeyMin, "水平修正按鍵後等待的最短時間（毫秒）");
            tt.SetToolTip(nHKeyMax, "水平修正按鍵後等待的最長時間（毫秒）");
            tt.SetToolTip(nVKeyMin, "垂直修正按鍵後等待的最短時間（毫秒）— 垂直較短");
            tt.SetToolTip(nVKeyMax, "垂直修正按鍵後等待的最長時間（毫秒）");
            tt.SetToolTip(nHTol, "水平方向偏差在此範圍內即停止修正");
            tt.SetToolTip(nVTol, "垂直方向偏差在此範圍內即停止修正");

            // ===== 🔬 方向診斷 =====
            GroupBox grpDiag = new GroupBox
            {
                Text = "🔬 方向診斷 (先確認每個按鍵實際讓角色往哪移動)",
                Left = 15, Top = paramY3 + 35, Width = 575, Height = 80,
                ForeColor = Color.Yellow, Font = new Font("microsoft yahei ui", 9F)
            };
            Button bDL = new Button { Text = "測試 ←左", Left = 10, Top = 22, Width = 75, Height = 26, BackColor = Color.FromArgb(60, 60, 70), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Button bDR = new Button { Text = "測試 →右", Left = 90, Top = 22, Width = 75, Height = 26, BackColor = Color.FromArgb(60, 60, 70), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Button bDU = new Button { Text = "測試 ↑上", Left = 170, Top = 22, Width = 75, Height = 26, BackColor = Color.FromArgb(60, 60, 70), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Button bDD = new Button { Text = "測試 ↓下", Left = 250, Top = 22, Width = 75, Height = 26, BackColor = Color.FromArgb(60, 60, 70), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Label lblDiag = new Label { Left = 10, Top = 50, Width = 555, Height = 20, ForeColor = Color.LightGray, Font = new Font("Consolas", 9F), Text = "點按鈕 → 角色移動一步 → 顯示座標變化" };
            grpDiag.Controls.AddRange(new Control[] { bDL, bDR, bDU, bDD, lblDiag });

            // ===== 修正日誌 =====
            GroupBox grpLog = new GroupBox
            {
                Text = "📋 修正日誌", Left = 15, Top = paramY3 + 120, Width = 575, Height = 190,
                ForeColor = Color.White, Font = new Font("microsoft yahei ui", 9F)
            };
            TextBox txtLog = new TextBox
            {
                Left = 5, Top = 20, Width = 563, Height = 125,
                Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(25, 25, 30), ForeColor = Color.LightGreen,
                Font = new Font("Consolas", 8.5F)
            };
            Button btnTestCorr = new Button { Text = "🧪 測試修正", Left = 5, Top = 152, Width = 100, Height = 28, BackColor = Color.FromArgb(180, 120, 0), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Button btnClearLog = new Button { Text = "清除", Left = 110, Top = 152, Width = 55, Height = 28, BackColor = Color.FromArgb(60, 60, 70), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Label lblTestRes = new Label { Left = 175, Top = 157, Width = 390, Height = 20, ForeColor = Color.LightGray, Font = new Font("microsoft yahei ui", 9F) };
            grpLog.Controls.AddRange(new Control[] { txtLog, btnTestCorr, btnClearLog, lblTestRes });

            // ===== 底部 =====
            int bottomY = paramY3 + 318;
            Button btnSave = new Button { Text = "💾 儲存", Left = 400, Top = bottomY, Width = 90, Height = 32, BackColor = Color.FromArgb(0, 140, 80), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            Button btnClose = new Button { Text = "關閉", Left = 500, Top = bottomY, Width = 85, Height = 32, BackColor = Color.FromArgb(80, 80, 85), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

            // ===== Helper: 建立修正器 =====
            Func<PositionCorrector> makeCorrector = () => new PositionCorrector
            {
                TargetWindow = targetWindowHandle,
                Tolerance = (int)nHTol.Value,
                SoftToleranceMin = (int)nHTol.Value,
                SoftToleranceMax = (int)nVTol.Value,
                HorizontalTolerance = (int)nHTol.Value,
                VerticalTolerance = (int)nVTol.Value,
                MaxCorrectionTimeMs = (int)nTO.Value,
                MaxStepsPerCorrection = (int)nMaxCorr.Value,
                MoveLeftKeys = tL.Tag as Keys[] ?? new[] { Keys.Left },
                MoveRightKeys = tR.Tag as Keys[] ?? new[] { Keys.Right },
                MoveUpKeys = tU.Tag as Keys[] ?? new[] { Keys.Up },
                MoveDownKeys = tD.Tag as Keys[] ?? new[] { Keys.Down },
                EnableHorizontalCorrection = chkH.Checked,
                EnableVerticalCorrection = chkV.Checked,
                InvertY = chkInvY.Checked,
                ExternalKeySender = SendKeyForCorrection,
                KeyIntervalMinMs = (int)nHKeyMin.Value,
                KeyIntervalMaxMs = (int)nHKeyMax.Value,
                HorizontalKeyIntervalMinMs = (int)nHKeyMin.Value,
                HorizontalKeyIntervalMaxMs = (int)nHKeyMax.Value,
                VerticalKeyIntervalMinMs = (int)nVKeyMin.Value,
                VerticalKeyIntervalMaxMs = (int)nVKeyMax.Value
            };

            Action<string> appendLog = (msg) =>
            {
                if (f.IsDisposed) return;
                f.BeginInvoke(new Action(() =>
                {
                    txtLog.AppendText(msg + Environment.NewLine);
                    txtLog.SelectionStart = txtLog.TextLength;
                    txtLog.ScrollToCaret();
                }));
            };

            // ===== 事件 =====
            // 方向診斷
            Action<string, Func<Keys[]>> doDiag = (dir, getKeys) =>
            {
                if (minimapTracker == null || !minimapTracker.IsCalibrated) { lblDiag.Text = "❌ 請先校準"; return; }
                var c = makeCorrector();
                lblDiag.Text = $"測試 {dir} 中...";
                lblDiag.ForeColor = Color.Yellow;
                Task.Run(() =>
                {
                    var r = c.DiagnoseDirection(minimapTracker, dir, getKeys());
                    f.BeginInvoke(new Action(() =>
                    {
                        lblDiag.Text = r.ToString();
                        lblDiag.ForeColor = r.Error != null ? Color.Red : Color.Lime;
                        appendLog(r.ToString());
                    }));
                });
            };
            bDL.Click += (s, e) => doDiag("←左", () => tL.Tag as Keys[] ?? new[] { Keys.Left });
            bDR.Click += (s, e) => doDiag("→右", () => tR.Tag as Keys[] ?? new[] { Keys.Right });
            bDU.Click += (s, e) => doDiag("↑上", () => tU.Tag as Keys[] ?? new[] { Keys.Up });
            bDD.Click += (s, e) => doDiag("↓下", () => tD.Tag as Keys[] ?? new[] { Keys.Down });

            btnClearLog.Click += (s, e) => txtLog.Clear();

            // 測試修正
            btnTestCorr.Click += (s, e) =>
            {
                if (minimapTracker == null || !minimapTracker.IsCalibrated) { lblTestRes.Text = "❌ 未校準"; lblTestRes.ForeColor = Color.Red; return; }
                var c = makeCorrector();
                c.OnLog += appendLog;

                lblTestRes.Text = "修正中..."; lblTestRes.ForeColor = Color.Yellow;
                btnTestCorr.Enabled = false;

                Task.Run(() =>
                {
                    // ★ 即時模式：讀取當前位置作為參考點，顯示修正器能力
                    var (cx, cy, ok) = minimapTracker.ReadPosition();
                    CorrectionResult result;
                    if (ok)
                    {
                        appendLog($"即時模式 — 當前({cx},{cy}) 將作為參考起點");
                        result = new CorrectionResult(true, $"即時模式 當前({cx},{cy}) 播放時自動比對腳本位置", cx, cy);
                    }
                    else result = new CorrectionResult(false, "偵測失敗");

                    f.BeginInvoke(new Action(() =>
                    {
                        btnTestCorr.Enabled = true;
                        lblTestRes.Text = (result.Success ? "✅ " : "⚠️ ") + result.Message;
                        lblTestRes.ForeColor = result.Success ? Color.Lime : Color.Orange;
                    }));
                });
            };

            btnSave.Click += (s, e) =>
            {
                positionCorrectionSettings.Enabled = chkEnabled.Checked;
                positionCorrectionSettings.UseDeviationMode = true; // 統一即時模式
                positionCorrectionSettings.HorizontalTolerance = (int)nHTol.Value;
                positionCorrectionSettings.VerticalTolerance = (int)nVTol.Value;
                positionCorrectionSettings.SoftToleranceMin = (int)nHTol.Value; // 向後兼容
                positionCorrectionSettings.SoftToleranceMax = (int)nVTol.Value; // 向後兼容
                positionCorrectionSettings.Tolerance = (int)nHTol.Value; // 向後兼容
                positionCorrectionSettings.MaxCorrectionTimeMs = (int)nTO.Value;
                positionCorrectionSettings.InvertY = chkInvY.Checked;
                positionCorrectionSettings.MoveLeftKeys = PositionCorrector.KeysToIntArray(tL.Tag as Keys[] ?? new[] { Keys.Left });
                positionCorrectionSettings.MoveRightKeys = PositionCorrector.KeysToIntArray(tR.Tag as Keys[] ?? new[] { Keys.Right });
                positionCorrectionSettings.MoveUpKeys = PositionCorrector.KeysToIntArray(tU.Tag as Keys[] ?? new[] { Keys.Up });
                positionCorrectionSettings.MoveDownKeys = PositionCorrector.KeysToIntArray(tD.Tag as Keys[] ?? new[] { Keys.Down });
                var lk = tL.Tag as Keys[]; var rk = tR.Tag as Keys[]; var uk = tU.Tag as Keys[]; var dk = tD.Tag as Keys[];
                positionCorrectionSettings.MoveLeftKey = lk?.Length > 0 ? (int)lk[0] : (int)Keys.Left;
                positionCorrectionSettings.MoveRightKey = rk?.Length > 0 ? (int)rk[0] : (int)Keys.Right;
                positionCorrectionSettings.MoveUpKey = uk?.Length > 0 ? (int)uk[0] : (int)Keys.Up;
                positionCorrectionSettings.MoveDownKey = dk?.Length > 0 ? (int)dk[0] : (int)Keys.Down;
                positionCorrectionSettings.EnableHorizontalCorrection = chkH.Checked;
                positionCorrectionSettings.EnableVerticalCorrection = chkV.Checked;
                positionCorrectionSettings.MaxCorrectionsPerLoop = 0; // 循環觸發次數不限制
                positionCorrectionSettings.MaxStepsPerCorrection = (int)nMaxCorr.Value; // ★ 每次修正的按鍵步數上限
                positionCorrectionSettings.CorrectionCheckIntervalSec = (int)nChkInt.Value;
                positionCorrectionSettings.KeyIntervalMinMs = (int)nHKeyMin.Value;
                positionCorrectionSettings.KeyIntervalMaxMs = (int)nHKeyMax.Value;
                positionCorrectionSettings.HorizontalKeyIntervalMinMs = (int)nHKeyMin.Value;
                positionCorrectionSettings.HorizontalKeyIntervalMaxMs = (int)nHKeyMax.Value;
                positionCorrectionSettings.VerticalKeyIntervalMinMs = (int)nVKeyMin.Value;
                positionCorrectionSettings.VerticalKeyIntervalMaxMs = (int)nVKeyMax.Value;
                // ★ 設定變更旗標：下個循環自動套用
                _correctionSettingsChanged = true;
                SaveAppSettings();
                AddLog("💾 [設定完成，將於下個循環套用]");
                f.Close();
            };

            btnClose.Click += (s, e) => f.Close();

            f.Controls.AddRange(new Control[] {
                chkEnabled, lblModeDesc, grpK,
                lChkInt, nChkInt, lChkInt2,
                lHTol, nHTol, lVTol, nVTol, lTolPx,
                lTO, nTO, lTOms,
                lHKeyInt, nHKeyMin, lHKeyTil, nHKeyMax, lHKeyMs,
                lVKeyInt, nVKeyMin, lVKeyTil, nVKeyMax, lVKeyMs,
                lMaxCorr, nMaxCorr,
                grpDiag, grpLog, btnSave, btnClose
            });
            f.ShowDialog();
        }

        private void UpdateUI()
        {
            btnStartRecording.Enabled = !isRecording && !isPlaying;
            btnStopRecording.Enabled = isRecording;
            btnStartPlayback.Enabled = !isPlaying && recordedEvents.Count > 0 && !isRecording;
            btnStopPlayback.Enabled = isPlaying;
        }

        private double GetCurrentTime()
        {
            return highResTimer.Elapsed.TotalSeconds;
        }

        private void Form1_Load(object? sender, EventArgs e)
        {
            lblStatus.Text = "就緒：點擊「開始錄製」開始";
            lblRecordingStatus.Text = "錄製：尚未開始";
            lblPlaybackStatus.Text = "播放：尚未開始";
        }

        [Serializable]
        public class MacroEvent
        {
            public Keys KeyCode { get; set; }
            public string EventType { get; set; } = "down";
            public double Timestamp { get; set; }
            public int CorrectTargetX { get; set; } = -1;
            public int CorrectTargetY { get; set; } = -1;
            /// <summary>錄製時的小地圖座標（-1 = 未記錄）</summary>
            public int RecordedX { get; set; } = -1;
            public int RecordedY { get; set; } = -1;
        }

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

                if (sfd.ShowDialog() != DialogResult.OK)
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

                if (ofd.ShowDialog() != DialogResult.OK)
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
            }
            catch (Exception ex)
            {
                AddLog($"設定載入失敗: {ex.Message}");
            }
        }

        private void btnMemoryScanner_Click(object sender, EventArgs e)
        {

        }

        private void lblLoopCount_Click(object sender, EventArgs e)
        {

        }

        private void txtPauseHotkey_TextChanged(object sender, EventArgs e)
        {

        }
    }
}