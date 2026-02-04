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

namespace MapleStoryMacro
{
    public partial class Form1 : Form
    {
        private List<MacroEvent> recordedEvents = new List<MacroEvent>();
        private bool isRecording = false;
        private bool isPlaying = false;
        private double recordStartTime = 0;
        private IntPtr targetWindowHandle = IntPtr.Zero;
        private KeyboardHookDLL keyboardHook;

        // 全局熱鍵設定
        private Keys playHotkey = Keys.F9;      // 預設播放熱鍵
        private Keys stopHotkey = Keys.F10;     // 預設停止熱鍵
        private bool hotkeyEnabled = true;      // 熱鍵是否啟用
        private KeyboardHookDLL hotkeyHook;     // 全局熱鍵監聽器


        // 自定義按鍵槽位 (15個)
        private CustomKeySlot[] customKeySlots = new CustomKeySlot[15];

        // 執行統計
        private PlaybackStatistics statistics = new PlaybackStatistics();

        // 定時執行
        private DateTime? scheduledStartTime = null;
        private System.Windows.Forms.Timer schedulerTimer;

        // Log System
        private List<string> logMessages = new List<string>();
        private readonly int MAX_LOG_LINES = 100;
        private readonly object logLock = new object();
        private string lastLogMessage = "";
        private int lastLogRepeatCount = 0;

        private KeyboardBlocker arrowKeyBlocker = new KeyboardBlocker();
        private HashSet<Keys> pressedKeys = new HashSet<Keys>();
        private readonly object pressedKeysLock = new object(); // 跨線程同步鎖

        // 背景切換模式
        private enum BackgroundSwitchMode
        {
            Manual,             // 手動切換（不自動）
            Immediate,          // 立即切換到背景
            ArrowKey_05s,       // 方向鍵 0.5 秒後切換
            ArrowKey_10s,       // 方向鍵 1.0 秒後切換
            ArrowKey_15s,       // 方向鍵 1.5 秒後切換
            ArrowKey_20s,       // 方向鍵 2.0 秒後切換
            ArrowKey_30s        // 方向鍵 3.0 秒後切換
        }
        private BackgroundSwitchMode currentBackgroundSwitchMode = BackgroundSwitchMode.ArrowKey_10s;

        // 背景切換監控器
        private System.Windows.Forms.Timer backgroundSwitchTimer;
        private Keys? currentHeldArrowKey = null;      // 目前按住的方向鍵
        private DateTime arrowKeyHeldStartTime;         // 開始按住的時間
        private bool hasCompletedBackgroundSwitch = false; // 是否已完成背景切換

        /// <summary>
        /// 取得當前背景切換模式的閾值秒數
        /// </summary>
        private double GetBackgroundSwitchThreshold()
        {
            return currentBackgroundSwitchMode switch
            {
                BackgroundSwitchMode.ArrowKey_05s => 0.5,
                BackgroundSwitchMode.ArrowKey_10s => 1.0,
                BackgroundSwitchMode.ArrowKey_15s => 1.5,
                BackgroundSwitchMode.ArrowKey_20s => 2.0,
                BackgroundSwitchMode.ArrowKey_30s => 3.0,
                _ => 1.0
            };
        }

        // Windows API P/Invoke
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);


        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

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
        private static extern IntPtr GetFocus();

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        // Foreground key sending API
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        // 視窗閃爍 API（用於通知使用者）
        [DllImport("user32.dll")]
        private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        private const uint FLASHW_ALL = 3;      // 閃爍標題列和工作列按鈕
        private const uint FLASHW_TIMERNOFG = 12; // 持續閃爍直到視窗獲得焦點

        /// <summary>
        /// 閃爍視窗以提示使用者（不阻塞）
        /// </summary>
        private void FlashWindowEx(IntPtr hWnd)
        {
            FLASHWINFO fwi = new FLASHWINFO();
            fwi.cbSize = (uint)Marshal.SizeOf(typeof(FLASHWINFO));
            fwi.hwnd = hWnd;
            fwi.dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG;
            fwi.uCount = 5;  // 閃爍 5 次
            fwi.dwTimeout = 0;
            FlashWindowEx(ref fwi);
        }

        // 子視窗列舉相關 API
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool AllowSetForegroundWindow(int dwProcessId);


        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        private const int SW_SHOWNOACTIVATE = 4;
        private const int ASFW_ANY = -1;

        // 方向鍵發送模式
        private enum ArrowKeyMode
        {
            ThreadAttach,       // 線程附加模式（遊戲可用，會影響前景）
            SendToChild,        // 發送到子視窗模式（純背景，不影響前景）
            ThreadAttachWithBlocker // 線程附加 + 攔截前景
        }
        private ArrowKeyMode currentArrowKeyMode = ArrowKeyMode.ThreadAttachWithBlocker;

        // SendInput structures - 正確的 64 位元結構體定義
        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public KEYBDINPUT ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
            // 需要填充以匹配 MOUSEINPUT 的大小（在 64 位元系統上）
            public ulong padding1;
            public uint padding2;
        }

        private const uint INPUT_KEYBOARD = 1;

        // Windows message constants
        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP = 0x0101;
        private const uint WM_CHAR = 0x0102;
        private const uint WM_SYSKEYDOWN = 0x0104;
        private const uint WM_SYSKEYUP = 0x0105;
        private const uint MAPVK_VK_TO_VSC = 0;
        private const uint KEYEVENTF_KEYDOWN = 0x0000;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const uint KEYEVENTF_SCANCODE = 0x0008;

        // 按鍵顯示名稱映射表
        private static readonly Dictionary<Keys, string> KeyDisplayNames = new Dictionary<Keys, string>
        {
            // 符號鍵 - Shift 時會變成 <> 等符號
  { Keys.OemPeriod, ". (>)" },      // . 和 >
      { Keys.Oemcomma, ", (<)" },       // , 和 <
            { Keys.OemQuestion, "/ (?)" },  // / 和 ?
     { Keys.OemSemicolon, "; (:)" },   // ; 和 :
            { Keys.OemQuotes, "' (\")" },     // ' 和 "
            { Keys.OemOpenBrackets, "[ ({)" }, // [ 和 {
            { Keys.OemCloseBrackets, "] (})" }, // ] 和 }
      { Keys.OemBackslash, "\\ (|)" },  // \ 和 |
 { Keys.OemMinus, "- (_)" },    // - 和 _
{ Keys.Oemplus, "= (+)" },        // = 和 +
    { Keys.Oemtilde, "` (~)" },     // ` 和 ~
            { Keys.OemPipe, "\\ (|)" },       // \ 和 |
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
            if (KeyDisplayNames.TryGetValue(key, out string displayName))
                return displayName;
            return key.ToString();
        }

        /// <summary>
        /// 從顯示名稱還原為 Keys 列舉值
        /// </summary>
        private static Keys ParseKeyFromDisplay(string displayName)
        {
            // 先嘗試從映射表反向查找
            foreach (var kvp in KeyDisplayNames)
            {
                if (kvp.Value == displayName)
                    return kvp.Key;
            }
            // 如果找不到，嘗試直接解析
            if (Enum.TryParse<Keys>(displayName, out Keys result))
                return result;
            throw new ArgumentException($"無法識別的按鍵名稱: {displayName}");
        }

        public Form1()
        {
            InitializeComponent();

            // 初始化自定義按鍵槽位
            for (int i = 0; i < 15; i++)
            {
                customKeySlots[i] = new CustomKeySlot { SlotNumber = i + 1 };
            }

            // 初始化定時器
            schedulerTimer = new System.Windows.Forms.Timer();
            schedulerTimer.Interval = 1000; // 每秒檢查一次
            schedulerTimer.Tick += SchedulerTimer_Tick;

            // 初始化背景切換監控器
            // ★ 功能：播放時鎖定前景，方向鍵持續按住 1.5 秒後自動切到背景
            backgroundSwitchTimer = new System.Windows.Forms.Timer();
            backgroundSwitchTimer.Interval = 100; // 每 100ms 檢查一次
            backgroundSwitchTimer.Tick += BackgroundSwitchTimer_Tick;

            // Bind all button events
            btnStartRecording.Click += BtnStartRecording_Click;
            btnStopRecording.Click += BtnStopRecording_Click;
            btnSaveScript.Click += BtnSaveScript_Click;
            btnLoadScript.Click += BtnLoadScript_Click;
            btnClearEvents.Click += BtnClearEvents_Click;
            btnViewEvents.Click += BtnViewEvents_Click;
            btnEditEvents.Click += BtnEditEvents_Click;
            btnStartPlayback.Click += BtnStartPlayback_Click;
            btnStopPlayback.Click += BtnStopPlayback_Click;
            btnRefreshWindow.Click += BtnRefreshWindow_Click;
            btnLockWindow.Click += BtnLockWindow_Click;
            btnHotkeySettings.Click += (s, e) => OpenHotkeySettings();
            btnCustomKeys.Click += (s, e) => OpenCustomKeySettings();
            btnScheduler.Click += (s, e) => OpenSchedulerSettings();
            btnStatistics.Click += (s, e) => ShowStatistics();
            FormClosing += Form1_FormClosing;

            // Enable KeyPreview to capture all keys
            this.KeyPreview = true;
            this.KeyDown += Form1_KeyDown;
            this.KeyUp += Form1_KeyUp;

            // Initialize keyboard hook for recording
            keyboardHook = new KeyboardHookDLL();
            keyboardHook.OnKeyEvent += KeyboardHook_OnKeyEvent;

            // Initialize global hotkey hook
            hotkeyHook = new KeyboardHookDLL();
            hotkeyHook.OnKeyEvent += HotkeyHook_OnKeyEvent;
            hotkeyHook.Install();  // 始終啟用全局熱鍵監聽

            // Initialize log system
            monitorTimer.Interval = 100;
            monitorTimer.Tick += MonitorTimer_Tick;
            monitorTimer.Start();

            // 載入儲存的設定
            LoadSettings();

            AddLog("應用程式已啟動");
            AddLog($"全局熱鍵：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}");

            UpdateUI();
            UpdateWindowStatus();
        }

        /// <summary>
        /// 定時執行檢查
        /// </summary>
        private void SchedulerTimer_Tick(object sender, EventArgs e)
        {
            if (scheduledStartTime.HasValue && DateTime.Now >= scheduledStartTime.Value)
            {
                scheduledStartTime = null;
                schedulerTimer.Stop();
                AddLog("定時執行觸發！");

                if (!isPlaying && recordedEvents.Count > 0)
                {
                    BtnStartPlayback_Click(null, null);
                }
            }
        }


        /// <summary>
        /// 背景切換監控器 - 根據設定模式監控
        /// </summary>
        private void BackgroundSwitchTimer_Tick(object sender, EventArgs e)
        {
            // 如果播放已停止或已完成背景切換，停止監控
            if (!isPlaying)
            {
                StopBackgroundSwitch();
                return;
            }

            if (hasCompletedBackgroundSwitch)
            {
                return; // 已完成，不需要繼續監控
            }

            // 手動模式不自動切換
            if (currentBackgroundSwitchMode == BackgroundSwitchMode.Manual)
            {
                return;
            }

            // 檢查目前腳本中正在按住的方向鍵（使用 lock 保護跨線程存取）
            Keys? heldArrowKey = null;
            Keys[] currentKeys;
            lock (pressedKeysLock)
            {
                currentKeys = pressedKeys.ToArray();
            }
            
            foreach (var key in currentKeys)
            {
                if (IsArrowKey(key))
                {
                    heldArrowKey = key;
                    break;
                }
            }

            if (heldArrowKey.HasValue)
            {
                // 有方向鍵被按住
                if (currentHeldArrowKey == heldArrowKey)
                {
                    // 同一個方向鍵持續按住，檢查時間
                    double heldSeconds = (DateTime.Now - arrowKeyHeldStartTime).TotalSeconds;
                    double threshold = GetBackgroundSwitchThreshold();
                    
                    if (heldSeconds >= threshold)
                    {
                        // 超過閾值，自動跳到背景！
                        ExecuteBackgroundSwitch($"方向鍵 {GetKeyDisplayName(heldArrowKey.Value)} 持續 {heldSeconds:F1} 秒");
                    }
                }
                else
                {
                    // 換了不同的方向鍵，重新計時
                    currentHeldArrowKey = heldArrowKey;
                    arrowKeyHeldStartTime = DateTime.Now;
                }
            }
            else
            {
                // 沒有方向鍵被按住，重置狀態
                currentHeldArrowKey = null;
            }
        }

        /// <summary>
        /// 執行背景切換
        /// ★ 重點：不釋放按鍵！讓腳本繼續正常執行，跟手動切換視窗一樣
        /// </summary>
        private void ExecuteBackgroundSwitch(string reason)
        {
            hasCompletedBackgroundSwitch = true;
            
            // ★ 不要呼叫 ReleaseKeysToGameWindow()！
            // 腳本會繼續發送按鍵，讓它自己管理按鍵狀態
            // 這樣才能跟手動切換視窗一樣正常運作
            
            // 把本程式視窗帶到前景（讓遊戲變背景）
            // 使用多種方法確保切換成功
            BringWindowToTop(this.Handle);
            SetForegroundWindow(this.Handle);
            this.Activate();
            
            AddLog($"★ 背景切換完成: {reason}");
            AddLog("遊戲已切換到背景，腳本繼續執行中...");
            
            // ★ 不使用 MessageBox 避免阻塞腳本執行
            // 改用閃爍工作列圖示提醒使用者
            FlashWindowEx(this.Handle);
        }

        /// <summary>
        /// 釋放所有按鍵到遊戲視窗（專用於背景切換時）
        /// ★ 直接發送 keyup 到遊戲視窗，不經過 SendKeyEvent
        /// </summary>
        private void ReleaseKeysToGameWindow()
        {
            Keys[] keysToRelease;
            lock (pressedKeysLock)
            {
                if (pressedKeys.Count == 0)
                    return;
                keysToRelease = pressedKeys.ToArray();
                pressedKeys.Clear();
            }

            AddLog($"釋放 {keysToRelease.Length} 個按鍵到遊戲視窗");

            // 直接發送到目標視窗
            if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
            {
                foreach (var key in keysToRelease)
                {
                    try
                    {
                        if (IsArrowKey(key))
                        {
                            SendArrowKeyWithMode(targetWindowHandle, key, false);
                        }
                        else if (IsAltKey(key))
                        {
                            SendAltKeyToWindow(targetWindowHandle, key, false);
                        }
                        else if (IsExtendedKey(key))
                        {
                            SendKeyWithThreadAttach(targetWindowHandle, key, false);
                        }
                        else
                        {
                            SendKeyToWindow(targetWindowHandle, key, false);
                        }
                        AddLog($"已釋放: {GetKeyDisplayName(key)}");
                    }
                    catch (Exception ex)
                    {
                        AddLog($"釋放按鍵失敗 {key}: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// 開始背景切換監控 - 播放腳本時自動呼叫
        /// </summary>
        private void StartBackgroundSwitch()
        {
            if (targetWindowHandle == IntPtr.Zero || !IsWindow(targetWindowHandle))
                return;

            // 重置狀態
            hasCompletedBackgroundSwitch = false;
            currentHeldArrowKey = null;

            // 根據模式決定行為
            switch (currentBackgroundSwitchMode)
            {
                case BackgroundSwitchMode.Manual:
                    // 手動模式：不做任何切換，直接開始播放
                    AddLog("★ 手動模式 - 不自動切換背景");
                    break;

                case BackgroundSwitchMode.Immediate:
                    // 立即切換：不鎖定前景，直接背景播放
                    AddLog("★ 立即背景模式 - 直接在背景播放");
                    hasCompletedBackgroundSwitch = true;
                    break;

                default:
                    // 方向鍵模式：鎖定前景，等待觸發
                    SetForegroundWindow(targetWindowHandle);
                    double threshold = GetBackgroundSwitchThreshold();
                    AddLog($"★ 前景模式啟動 - 等待方向鍵持續 {threshold} 秒後自動切到背景");
                    backgroundSwitchTimer.Start();
                    break;
            }
        }

        /// <summary>
        /// 停止背景切換監控
        /// </summary>
        private void StopBackgroundSwitch()
        {
            backgroundSwitchTimer.Stop();
            currentHeldArrowKey = null;
            hasCompletedBackgroundSwitch = false;
        }

        private void AddLog(string message)
        {
            lock (logLock)
            {
                // 檢查是否與上一條訊息相同（合併重複訊息）
                if (message == lastLogMessage && logMessages.Count > 0)
                {
                    lastLogRepeatCount++;
                    // 更新最後一條訊息，加上重複次數
                    string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                    logMessages[logMessages.Count - 1] = $"[{timestamp}] {message} x{lastLogRepeatCount}";
                }
                else
                {
                    // 新訊息
                    lastLogMessage = message;
                    lastLogRepeatCount = 1;
                    string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                    string logEntry = $"[{timestamp}] {message}";
                    logMessages.Add(logEntry);

                    if (logMessages.Count > MAX_LOG_LINES)
                        logMessages.RemoveAt(0);
                }
            }

            // 更新日誌顯示（在 UI 執行緒）
            if (txtLogDisplay != null && !txtLogDisplay.IsDisposed)
            {
                try
                {
                    if (txtLogDisplay.InvokeRequired)
                    {
                        txtLogDisplay.BeginInvoke(new Action(UpdateLogDisplay));
                    }
                    else
                    {
                        UpdateLogDisplay();
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 更新日誌顯示
        /// </summary>
        private void UpdateLogDisplay()
        {
            if (txtLogDisplay == null || txtLogDisplay.IsDisposed)
                return;

            lock (logLock)
            {
                // 暫停繪製以提升效能
                txtLogDisplay.BeginUpdate();
                
                // 如果項目數量不同，重新填充
                if (txtLogDisplay.Items.Count != logMessages.Count)
                {
                    txtLogDisplay.Items.Clear();
                    foreach (var msg in logMessages)
                    {
                        txtLogDisplay.Items.Add(msg);
                    }
                }
                else if (logMessages.Count > 0)
                {
                    // 只更新最後一項（重複訊息時）
                    txtLogDisplay.Items[txtLogDisplay.Items.Count - 1] = logMessages[logMessages.Count - 1];
                }
                
                txtLogDisplay.EndUpdate();
                
                // 自動捲動到底部
                if (txtLogDisplay.Items.Count > 0)
                {
                    txtLogDisplay.TopIndex = txtLogDisplay.Items.Count - 1;
                }
            }
        }

        private void MonitorTimer_Tick(object sender, EventArgs e)
        {
            // 更新狀態列
            string bgMode = (targetWindowHandle != IntPtr.Zero) ? "BG" : "FG";
            string statusLine = $"Rec: {(isRecording ? "ON" : "OFF")} | Play: {(isPlaying ? "ON" : "OFF")} | Mode: {bgMode} | Events: {recordedEvents.Count}";
            lblStatus.Text = statusLine;
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (isRecording)
            {
                if (recordStartTime == 0)
                    recordStartTime = GetCurrentTime();

                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = e.KeyCode,
                    EventType = "down",
                    Timestamp = GetCurrentTime() - recordStartTime
                });

                lblRecordingStatus.Text = $"Recording: {recordedEvents.Count} events | Last: {e.KeyCode}";
                AddLog($"KeyDown: {e.KeyCode}");
            }
        }

        private void Form1_KeyUp(object sender, KeyEventArgs e)
        {
            if (isRecording)
            {
                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = e.KeyCode,
                    EventType = "up",
                    Timestamp = GetCurrentTime() - recordStartTime
                });

                lblRecordingStatus.Text = $"Recording: {recordedEvents.Count} events | Last: {e.KeyCode}";
                AddLog($"KeyUp: {e.KeyCode}");
            }
        }

        private void KeyboardHook_OnKeyEvent(Keys keyCode, bool isKeyDown)
        {
            if (isRecording)
            {
                if (recordStartTime == 0)
                    recordStartTime = GetCurrentTime();

                string eventType = isKeyDown ? "down" : "up";

                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = keyCode,
                    EventType = eventType,
                    Timestamp = GetCurrentTime() - recordStartTime
                });

                lblRecordingStatus.Text = $"Recording: {recordedEvents.Count} events | Last: {keyCode}";

                if (isKeyDown)
                    AddLog($"Hook Down: {keyCode}");
                else
                    AddLog($"Hook Up: {keyCode}");
            }
        }

        private void BtnRefreshWindow_Click(object sender, EventArgs e)
        {
            AddLog("Searching for target window...");
            FindGameWindow();
            UpdateWindowStatus();
        }

        private void BtnLockWindow_Click(object sender, EventArgs e)
        {
            AddLog("Opening window selector...");
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
                Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold),
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
                        listBox.Items.Add(new ProcessItem { Handle = p.MainWindowHandle, Title = $"{p.MainWindowTitle} (PID: {p.Id})" });
                    }
                }
                catch { }
            }

            // ★ 雙擊快速選擇
            listBox.DoubleClick += (s, args) =>
            {
                if (listBox.SelectedIndex >= 0)
                {
                    ProcessItem selected = listBox.SelectedItem as ProcessItem;
                    if (selected != null)
                    {
                        targetWindowHandle = selected.Handle;
                        UpdateWindowStatus();
                        AddLog($"★ 快速鎖定: {selected.Title}");
                        AddLog("Background mode ENABLED");
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
                Text = "OK",
                Width = 80,
                Height = 30,
                Left = 200,
                Top = 5,
                DialogResult = DialogResult.OK
            };

            Button cancelBtn = new Button
            {
                Text = "Cancel",
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
                ProcessItem selected = listBox.SelectedItem as ProcessItem;
                if (selected != null)
                {
                    targetWindowHandle = selected.Handle;
                    UpdateWindowStatus();
                    AddLog($"Window locked: {selected.Title}");
                    AddLog("Background mode ENABLED");
                }
            }
        }

        private class ProcessItem
        {
            public IntPtr Handle { get; set; }
            public string Title { get; set; }
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
                lblWindowStatus.Text = $"Window: Locked - BG Mode ON";
                lblWindowStatus.ForeColor = Color.Green;
                AddLog($"Window locked: {targetWindowHandle}");
            }
            else
            {
                lblWindowStatus.Text = "Window: Not found - FG Mode";
                lblWindowStatus.ForeColor = Color.Red;
                targetWindowHandle = IntPtr.Zero;
                AddLog("Target window not found");
            }

            UpdateArrowKeyBlockerState();
        }

        private void UpdateArrowKeyBlockerState()
        {
            bool shouldEnable = isPlaying
                && currentArrowKeyMode == ArrowKeyMode.ThreadAttachWithBlocker
                && targetWindowHandle != IntPtr.Zero
                && IsWindow(targetWindowHandle);

            if (shouldEnable)
            {
                arrowKeyBlocker.SetTargetWindow(targetWindowHandle);
                if (!arrowKeyBlocker.IsInstalled)
                {
                    if (arrowKeyBlocker.Install())
                    {
                        AddLog("方向鍵攔截已啟用");
                    }
                    else
                    {
                        AddLog("方向鍵攔截啟用失敗");
                    }
                }
            }
            else if (arrowKeyBlocker.IsInstalled)
            {
                arrowKeyBlocker.Uninstall();
                // 解除攔截器時釋放所有按鍵到遊戲視窗
                ReleaseKeysToGameWindow();
                AddLog("方向鍵攔截已停用，已釋放所有按鍵");
            }
        }

        private void ReleasePressedKeys()
        {
            Keys[] keysToRelease;
            lock (pressedKeysLock)
            {
                if (pressedKeys.Count == 0)
                    return;

                keysToRelease = pressedKeys.ToArray();
                pressedKeys.Clear();
            }

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

        private void BtnStartRecording_Click(object sender, EventArgs e)
        {
            if (isRecording) return;

            recordedEvents.Clear();
            recordStartTime = 0;
            isRecording = true;

            lblRecordingStatus.Text = "Recording...";
            lblRecordingStatus.ForeColor = Color.Red;
            lblStatus.Text = "Press keys to record";

            AddLog("Recording started...");

            if (keyboardHook.Install())
            {
                AddLog("Keyboard hook activated");
            }
            else
            {
                AddLog("Keyboard hook failed");
                MessageBox.Show("Failed to start keyboard hook");
                isRecording = false;
                return;
            }

            UpdateUI();
        }

        private void BtnStopRecording_Click(object sender, EventArgs e)
        {
            if (!isRecording) return;

            isRecording = false;
            lblRecordingStatus.Text = $"Stopped | Events: {recordedEvents.Count}";
            lblRecordingStatus.ForeColor = Color.Green;

            keyboardHook.Uninstall();

            AddLog($"Recording stopped - Total: {recordedEvents.Count} events");
            UpdateUI();
        }

        private void BtnSaveScript_Click(object sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                AddLog("No events to save");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "JSON Script|*.json",
                DefaultExt = ".json"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string json = JsonSerializer.Serialize(recordedEvents, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(sfd.FileName, json);
                    AddLog($"Saved: {Path.GetFileName(sfd.FileName)}");
                }
                catch (Exception ex)
                {
                    AddLog($"Save failed: {ex.Message}");
                }
            }
        }

        private void BtnLoadScript_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "JSON 腳本|*.json",
                Title = "載入腳本"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string json = File.ReadAllText(ofd.FileName);
                    recordedEvents = JsonSerializer.Deserialize<List<MacroEvent>>(json) ?? new List<MacroEvent>();
                    lblRecordingStatus.Text = $"已載入 | 事件數: {recordedEvents.Count}";
                    AddLog($"已載入: {recordedEvents.Count} 個事件");
                    UpdateUI();
                }
                catch (Exception ex)
                {
                    AddLog($"載入失敗: {ex.Message}");
                }
            }
        }

        private void BtnClearEvents_Click(object sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                MessageBox.Show("目前沒有任何事件可清除", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 顯示警告視窗
            DialogResult result = MessageBox.Show(
             $"確定要清除所有 {recordedEvents.Count} 個事件嗎？\n\n此操作無法復原！",
                 "⚠️ 警告 - 清除所有事件",
            MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
           MessageBoxDefaultButton.Button2  // 預設選擇「否」
            );

            if (result == DialogResult.Yes)
            {
                recordedEvents.Clear();
                lblRecordingStatus.Text = "已清除 | 事件數: 0";
                AddLog("已清除所有事件");
                MessageBox.Show("所有事件已清除", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 更新 UI 狀態
                UpdateUI();
            }
        }

        private void BtnViewEvents_Click(object sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                MessageBox.Show("沒有事件可顯示");
                AddLog("沒有事件可檢視");
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

        private void BtnEditEvents_Click(object sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                MessageBox.Show("沒有事件可編輯");
                AddLog("沒有事件可編輯");
                return;
            }

            AddLog("正在開啟編輯器...");

            Form editorForm = new Form
            {
                Text = "編輯腳本",
                Width = 800,
                Height = 600,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this
            };

            DataGridView dgv = new DataGridView
            {
                Top = 10,
                Left = 10,
                Width = 760,
                Height = 480,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            };

            // 建立按鍵下拉選單欄位
            DataGridViewComboBoxColumn keyColumn = new DataGridViewComboBoxColumn
            {
                Name = "KeyCode",
                HeaderText = "按鍵",
                Width = 150
            };
            // 添加所有可用的按鍵選項
            foreach (var kvp in KeyDisplayNames)
            {
                keyColumn.Items.Add(kvp.Value);
            }
            // 也添加常用的字母和數字鍵
            for (char c = 'A'; c <= 'Z'; c++)
                keyColumn.Items.Add(c.ToString());
            for (char c = '0'; c <= '9'; c++)
                keyColumn.Items.Add(c.ToString());
            // 添加 F1-F12
            for (int i = 1; i <= 12; i++)
                keyColumn.Items.Add($"F{i}");
            dgv.Columns.Add(keyColumn);

            // 建立類型下拉選單欄位
            DataGridViewComboBoxColumn typeColumn = new DataGridViewComboBoxColumn
            {
                Name = "EventType",
                HeaderText = "類型",
                Width = 80
            };
            typeColumn.Items.Add("按下");
            typeColumn.Items.Add("放開");
            dgv.Columns.Add(typeColumn);

            dgv.Columns.Add("Timestamp", "時間 (秒)");

            foreach (MacroEvent evt in recordedEvents)
            {
                string keyName = GetKeyDisplayName(evt.KeyCode);
                string eventType = evt.EventType == "down" ? "按下" : "放開";
                dgv.Rows.Add(keyName, eventType, evt.Timestamp.ToString("F3"));
            }

            Panel btnPanel = new Panel
            {
                Top = 500,
                Left = 10,
                Width = 760,
                Height = 50,
                BorderStyle = BorderStyle.FixedSingle
            };

            Button saveBtn = new Button { Text = "儲存", Width = 100, Height = 30, Left = 10, Top = 10 };
            Button cancelBtn = new Button { Text = "取消", Width = 100, Height = 30, Left = 120, Top = 10 };

            saveBtn.Click += (s, args) =>
          {
              try
              {
                  recordedEvents.Clear();
                  foreach (DataGridViewRow row in dgv.Rows)
                  {
                      if (row.Cells[0].Value != null)
                      {
                          string keyDisplayName = row.Cells[0].Value.ToString();
                          Keys keyCode = ParseKeyFromDisplay(keyDisplayName);
                          string eventTypeDisplay = row.Cells[1].Value.ToString();
                          string eventType = eventTypeDisplay == "按下" ? "down" : "up";
                          double timestamp = double.Parse(row.Cells[2].Value.ToString());

                          recordedEvents.Add(new MacroEvent
                          {
                              KeyCode = keyCode,
                              EventType = eventType,
                              Timestamp = timestamp
                          });
                      }
                  }

                  lblRecordingStatus.Text = $"已編輯 | 事件數: {recordedEvents.Count}";
                  AddLog($"已儲存編輯 - {recordedEvents.Count} 個事件");
                  editorForm.Close();
              }
              catch (Exception ex)
              {
                  AddLog($"儲存失敗: {ex.Message}");
              }
          };

            cancelBtn.Click += (s, args) => editorForm.Close();

            btnPanel.Controls.Add(saveBtn);
            btnPanel.Controls.Add(cancelBtn);

            editorForm.Controls.Add(dgv);
            editorForm.Controls.Add(btnPanel);

            editorForm.ShowDialog();
        }

        private void BtnStartPlayback_Click(object sender, EventArgs e)
        {
            if (isPlaying || recordedEvents.Count == 0)
                return;

            isPlaying = true;

            UpdateArrowKeyBlockerState();

            lock (pressedKeysLock)
            {
                pressedKeys.Clear();
            }

            // 重置自定義按鍵槽位的觸發狀態
            foreach (var slot in customKeySlots)
            {
                slot.Reset();
            }

            // 開始統計
            statistics.StartSession();

            string mode = (targetWindowHandle != IntPtr.Zero) ? "Background" : "Foreground";
            AddLog($"Playback started ({mode} mode)...");

            // 背景模式：啟動背景切換監控
            if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
            {
                StartBackgroundSwitch();
            }

            // 顯示啟用的自定義按鍵
            int enabledCount = customKeySlots.Count(s => s.Enabled);
            if (enabledCount > 0)
            {
                AddLog($"自定義按鍵：{enabledCount} 個已啟用");
            }

            try
            {
                int loopCount = (int)numPlayTimes.Value;
                Thread playbackThread = new Thread(() => PlaybackThread(loopCount))
                {
                    IsBackground = true
                };
                playbackThread.Start();
            }
            catch (Exception ex)
            {
                AddLog($"Playback failed: {ex.Message}");
                MessageBox.Show($"Playback failed: {ex.Message}");
                ReleasePressedKeys();
                isPlaying = false;
                statistics.EndSession();
                UpdateArrowKeyBlockerState();
            }

            UpdateUI();
        }

        private void PlaybackThread(int loopCount)
        {
            try
            {
                for (int loop = 1; loop <= loopCount && isPlaying; loop++)
                {
                    statistics.IncrementLoop();

                    this.Invoke(new Action(() =>
                       {
                           lblPlaybackStatus.Text = $"Loop: {loop}/{loopCount}";
                           lblPlaybackStatus.ForeColor = Color.Blue;
                       }));

                    double lastTimestamp = 0;
                    DateTime loopStartTime = DateTime.Now;

                    foreach (MacroEvent evt in recordedEvents)
                    {
                        if (!isPlaying) break;

                        double waitTime = evt.Timestamp - lastTimestamp;
                        if (waitTime > 0)
                        {
                            // 在等待期間檢查自定義按鍵
                            int waitMs = (int)(waitTime * 1000);
                            int elapsed = 0;
                            while (elapsed < waitMs && isPlaying)
                            {
                                int sleepTime = Math.Min(50, waitMs - elapsed); // 每 50ms 檢查一次
                                Thread.Sleep(sleepTime);
                                elapsed += sleepTime;

                                // 檢查並觸發自定義按鍵
                                double currentScriptTime = (DateTime.Now - loopStartTime).TotalSeconds;
                                CheckAndTriggerCustomKeys(currentScriptTime);
                            }
                        }

                        SendKeyEvent(evt);
                        lastTimestamp = evt.Timestamp;
                    }

                    Thread.Sleep(200);
                }

                isPlaying = false;
                statistics.EndSession();
                UpdateArrowKeyBlockerState();

                this.Invoke(new Action(() =>
        {
            lblPlaybackStatus.Text = "Playback: Completed";
            lblPlaybackStatus.ForeColor = Color.Green;
            AddLog($"Playback completed - 循環: {statistics.CurrentLoopCount}");
            UpdateUI();
        }));
            }
            catch (Exception ex)
            {
                statistics.EndSession();
                this.Invoke(new Action(() =>
              {
                  AddLog($"Playback error: {ex.Message}");
                  MessageBox.Show($"Playback error: {ex.Message}");
              }));
                ReleasePressedKeys();
                isPlaying = false;
                UpdateArrowKeyBlockerState();
            }
        }

        /// <summary>
        /// 檢查並觸發自定義按鍵
        /// </summary>
        /// <returns>返回需要暫停的總時間（秒）</returns>
        private double CheckAndTriggerCustomKeys(double currentScriptTime)
        {
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
                    SendCustomKey(slot.KeyCode);

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
        /// 發送自定義按鍵（按下後立即放開）
        /// </summary>
        private void SendCustomKey(Keys key)
        {
            try
            {
                if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
                {
                    // 背景模式
                    SendKeyToWindow(targetWindowHandle, key, true);  // 按下
                    Thread.Sleep(30);
                    SendKeyToWindow(targetWindowHandle, key, false); // 放開
                }
                else
                {
                    // 前景模式
                    SendKeyForeground(key, true);  // 按下
                    Thread.Sleep(30);
                    SendKeyForeground(key, false); // 放開
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

        private void SendKeyEvent(MacroEvent evt)
        {
            try
            {
                // Check if we have a valid target window for background sending
                if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
                {
                    // 對於 Alt 鍵，使用特殊的發送方式
                    if (IsAltKey(evt.KeyCode))
                    {
                        SendAltKeyToWindow(targetWindowHandle, evt.KeyCode, evt.EventType == "down");
                        AddLog($"BG(Alt): {evt.KeyCode} ({evt.EventType})");
                    }
                    // 對於方向鍵，根據設定的模式發送
                    else if (IsArrowKey(evt.KeyCode))
                    {
                        SendArrowKeyWithMode(targetWindowHandle, evt.KeyCode, evt.EventType == "down");
                        AddLog($"BG({currentArrowKeyMode}): {evt.KeyCode} ({evt.EventType})");
                    }
                    // 對於其他延伸鍵，使用線程附加模式
                    else if (IsExtendedKey(evt.KeyCode))
                    {
                        SendKeyWithThreadAttach(targetWindowHandle, evt.KeyCode, evt.EventType == "down");
                        AddLog($"BG(Attach): {evt.KeyCode} ({evt.EventType})");
                    }
                    else
                    {
                        // 一般按鍵：使用背景模式
                        SendKeyToWindow(targetWindowHandle, evt.KeyCode, evt.EventType == "down");
                        AddLog($"BG: {evt.KeyCode} ({evt.EventType})");
                    }
                }
                else
                {
                    // Foreground key sending using keybd_event
                    SendKeyForeground(evt.KeyCode, evt.EventType == "down");
                    AddLog($"FG: {evt.KeyCode} ({evt.EventType})");
                }

                if (evt.EventType == "down")
                {
                    lock (pressedKeysLock)
                    {
                        pressedKeys.Add(evt.KeyCode);
                    }
                }
                else if (evt.EventType == "up")
                {
                    lock (pressedKeysLock)
                    {
                        pressedKeys.Remove(evt.KeyCode);
                    }
                }
            }
            catch (Exception ex)
            {
                AddLog($"Send failed: {ex.Message}");
            }
        }

        /// <summary>
        /// 檢查是否為方向鍵
        /// </summary>
        private bool IsArrowKey(Keys key)
        {
            return key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down;
        }

        /// <summary>
        /// 根據當前模式發送方向鍵
        /// </summary>
        private void SendArrowKeyWithMode(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            switch (currentArrowKeyMode)
            {
                case ArrowKeyMode.ThreadAttach:
                    SendKeyWithThreadAttach(hWnd, key, isKeyDown);
                    break;

                case ArrowKeyMode.SendToChild:
                    SendArrowKeyToChildWindow(hWnd, key, isKeyDown);
                    break;

                case ArrowKeyMode.ThreadAttachWithBlocker:
                    SendArrowKeyToChildWindow(hWnd, key, isKeyDown);
                    SendKeyWithThreadAttach(hWnd, key, isKeyDown, KeyboardBlocker.MacroKeyMarker);
                    break;

                default:
                    SendKeyWithThreadAttach(hWnd, key, isKeyDown);
                    break;
            }
        }

        /// <summary>
        /// 發送方向鍵到子視窗模式 - 純 PostMessage，不影響前景
        /// </summary>
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
        /// 使用 SendInput 發送方向鍵（更可靠但需要焦點）
        /// </summary>
        private void SendArrowKeyWithSendInput(IntPtr hWnd, Keys key, bool isKeyDown)
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

                INPUT[] inputs = new INPUT[1];
                inputs[0].type = INPUT_KEYBOARD;
                inputs[0].ki.wVk = (ushort)key;
                inputs[0].ki.wScan = (ushort)GetScanCode(key);
                inputs[0].ki.dwFlags = KEYEVENTF_EXTENDEDKEY;
                if (!isKeyDown)
                {
                    inputs[0].ki.dwFlags |= KEYEVENTF_KEYUP;
                }
                inputs[0].ki.time = 0;
                inputs[0].ki.dwExtraInfo = UIntPtr.Zero;

                uint result = SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
                if (result == 0)
                {
                    AddLog($"SendInput 失敗: {Marshal.GetLastWin32Error()}");
                }
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

                AddLog($"Alt 按鍵: VK=0x{vkCode:X2}, SC=0x{scanCode:X2}, flags=0x{flags:X}");
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
        /// 前景模式發送按鍵（使用 keybd_event API）
        /// </summary>
        private void SendKeyForeground(Keys key, bool isKeyDown)
        {
            byte vkCode = (byte)key;
            byte scanCode = GetScanCode(key);
            bool isExtendedKey = IsExtendedKey(key);

            uint flags = 0;
            if (!isKeyDown)
            {
                flags |= KEYEVENTF_KEYUP;
            }
            if (isExtendedKey)
            {
                flags |= KEYEVENTF_EXTENDEDKEY;
            }

            keybd_event(vkCode, scanCode, flags, UIntPtr.Zero);
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

        private void BtnStopPlayback_Click(object sender, EventArgs e)
        {
            if (isPlaying)
            {
                statistics.EndSession();
            }
            StopBackgroundSwitch();  // 停止背景切換監控
            
            // ★ 使用 ReleaseKeysToGameWindow 確保按鍵正確釋放到遊戲
            ReleaseKeysToGameWindow();
            
            isPlaying = false;
            UpdateArrowKeyBlockerState();
            lblPlaybackStatus.Text = "Playback: Stopped";
            lblPlaybackStatus.ForeColor = Color.Orange;
            AddLog("播放已停止，已釋放所有按鍵");
            UpdateUI();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (isRecording)
                BtnStopRecording_Click(null, null);
            if (isPlaying)
            {
                statistics.EndSession();
                isPlaying = false;
            }

            keyboardHook?.Uninstall();
            hotkeyHook?.Uninstall();  // 停止全局熱鍵監聽
            arrowKeyBlocker?.Uninstall();
            monitorTimer?.Stop();
            schedulerTimer?.Stop();   // 停止定時執行計時器
            backgroundSwitchTimer?.Stop();  // 停止背景切換監控器
            AddLog("應用程式已關閉");
        }

        /// <summary>
        /// 全局熱鍵事件處理
        /// </summary>
        private void HotkeyHook_OnKeyEvent(Keys keyCode, bool isKeyDown)
        {
            // 只處理按下事件，避免重複觸發
            if (!isKeyDown || !hotkeyEnabled || isRecording)
                return;

            // 檢查是否為播放熱鍵
            if (keyCode == playHotkey)
            {
                this.BeginInvoke(new Action(() =>
                {
                    if (!isPlaying && recordedEvents.Count > 0)
                    {
                        AddLog($"熱鍵觸發：開始播放 ({GetKeyDisplayName(playHotkey)})");
                        BtnStartPlayback_Click(null, null);
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
                        BtnStopPlayback_Click(null, null);
                    }
                }));
            }
        }

        /// <summary>
        /// 設定熱鍵
        /// </summary>
        private void SetHotkeys(Keys play, Keys stop)
        {
            playHotkey = play;
            stopHotkey = stop;
            AddLog($"熱鍵已設定：播放={GetKeyDisplayName(play)}, 停止={GetKeyDisplayName(stop)}");
        }

        /// <summary>
        /// 開啟熱鍵設定視窗
        /// </summary>
        private void OpenHotkeySettings()
        {
            Form settingsForm = new Form
            {
                Text = "⚙ 熱鍵與進階設定",
                Width = 400,
                Height = 320,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            int yPos = 25;

            // 播放熱鍵
            Label lblPlay = new Label { Text = "播放熱鍵：", Left = 20, Top = yPos, Width = 100, ForeColor = Color.White };
            TextBox txtPlay = new TextBox
            {
                Left = 130,
                Top = yPos - 3,
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

            yPos += 35;

            // 停止熱鍵
            Label lblStop = new Label { Text = "停止熱鍵：", Left = 20, Top = yPos, Width = 100, ForeColor = Color.White };
            TextBox txtStop = new TextBox
            {
                Left = 130,
                Top = yPos - 3,
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

            yPos += 35;

            // 啟用熱鍵
            CheckBox chkEnabled = new CheckBox
            {
                Text = "啟用全局熱鍵",
                Left = 20,
                Top = yPos,
                Width = 150,
                Checked = hotkeyEnabled,
                ForeColor = Color.White
            };

            yPos += 40;

            // 方向鍵模式選擇（技術選項，通常不需要改）
            Label lblArrowMode = new Label
            {
                Text = "方向鍵發送：",
                Left = 20,
                Top = yPos,
                Width = 100,
                ForeColor = Color.Cyan
            };
            ComboBox cmbArrowMode = new ComboBox
            {
                Left = 130,
                Top = yPos - 3,
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White
            };
            cmbArrowMode.Items.Add("ThreadAttach+Block");
            cmbArrowMode.SelectedIndex = 0;

            yPos += 35;

            // 提示文字
            Label lblHint = new Label
            {
                Text = "提示：點擊文字框後按下想要的按鍵",
                Left = 20,
                Top = yPos,
                Width = 350,
                ForeColor = Color.Gray
            };

            yPos += 35;

            // 按鈕
            Button btnSave = new Button
            {
                Text = "儲存",
                Left = 150,
                Top = yPos,
                Width = 80,
                Height = 30,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            Button btnCancel = new Button
            {
                Text = "取消",
                Left = 240,
                Top = yPos,
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
                hotkeyEnabled = chkEnabled.Checked;
                
                UpdateArrowKeyBlockerState();
                SaveSettings();

                AddLog($"熱鍵設定已儲存：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}");
                
                MessageBox.Show(
                    $"設定已儲存！\n\n" +
                    $"播放熱鍵：{GetKeyDisplayName(playHotkey)}\n" +
                    $"停止熱鍵：{GetKeyDisplayName(stopHotkey)}\n" +
                    $"熱鍵狀態：{(hotkeyEnabled ? "已啟用" : "已停用")}",
                    "設定完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                settingsForm.Close();
            };

            btnCancel.Click += (s, args) => settingsForm.Close();

            settingsForm.Controls.AddRange(new Control[]
            {
                lblPlay, txtPlay, lblStop, txtStop, chkEnabled,
                lblArrowMode, cmbArrowMode,
                lblHint, btnSave, btnCancel
            });

            settingsForm.ShowDialog();
        }

        /// <summary>
        /// 開啟自定義按鍵設定視窗（15 個槽位，使用 DataGridView）
        /// </summary>
        private void OpenCustomKeySettings()
        {
            Form customForm = new Form
            {
                Text = "⚡ 自定義按鍵設定 (15 個槽位)",
                Width = 700,
                Height = 550,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = true,
                MinimizeBox = false,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            // 說明文字 (頂部)
            Label lblTitle = new Label
            {
                Text = "設定最多 15 個自定義按鍵，在腳本播放時按間隔自動施放",
                Dock = DockStyle.Top,
                Height = 30,
                ForeColor = Color.LightGray,
                Font = new Font("Microsoft JhengHei UI", 9F),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };

            // DataGridView
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.FromArgb(40, 40, 45),
                BorderStyle = BorderStyle.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(50, 50, 55),
                    ForeColor = Color.White,
                    SelectionBackColor = Color.FromArgb(0, 122, 204),
                    SelectionForeColor = Color.White
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(60, 60, 65),
                    ForeColor = Color.White,
                    Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold)
                },
                EnableHeadersVisualStyles = false,
                GridColor = Color.FromArgb(70, 70, 75),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EditMode = DataGridViewEditMode.EditOnEnter
            };

            // 建立欄位
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Slot", HeaderText = "#", Width = 30, ReadOnly = true });
            dgv.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "啟用", Width = 50 });
            
            // 按鍵欄位 - 使用 TextBox 並在 KeyDown 事件捕獲按鍵
            DataGridViewTextBoxColumn keyColumn = new DataGridViewTextBoxColumn
            {
                Name = "KeyCode",
                HeaderText = "按鍵 (點擊後按鍵)",
                Width = 120
            };
            dgv.Columns.Add(keyColumn);
            
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Interval", HeaderText = "間隔(秒)", Width = 70 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "StartAt", HeaderText = "開始(秒)", Width = 70 });
            dgv.Columns.Add(new DataGridViewCheckBoxColumn { Name = "PauseEnabled", HeaderText = "暫停", Width = 50 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "PauseSeconds", HeaderText = "暫停(秒)", Width = 70 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "PreDelay", HeaderText = "延遲(秒)", Width = 70 });
            
            // 清除按鈕欄位
            DataGridViewButtonColumn clearColumn = new DataGridViewButtonColumn
            {
                Name = "Clear",
                HeaderText = "清除",
                Text = "×",
                UseColumnTextForButtonValue = true,
                Width = 45
            };
            dgv.Columns.Add(clearColumn);

            // 填入資料
            for (int i = 0; i < 15; i++)
            {
                int rowIndex = dgv.Rows.Add();
                DataGridViewRow row = dgv.Rows[rowIndex];
                row.Cells["Slot"].Value = $"{i + 1:D2}";
                row.Cells["Enabled"].Value = customKeySlots[i].Enabled;
                row.Cells["KeyCode"].Value = customKeySlots[i].KeyCode == Keys.None ? "(點擊)" : GetKeyDisplayName(customKeySlots[i].KeyCode);
                row.Cells["KeyCode"].Tag = customKeySlots[i].KeyCode;
                row.Cells["Interval"].Value = customKeySlots[i].IntervalSeconds;
                row.Cells["StartAt"].Value = customKeySlots[i].StartAtSecond;
                row.Cells["PauseEnabled"].Value = customKeySlots[i].PauseScriptEnabled;
                row.Cells["PauseSeconds"].Value = customKeySlots[i].PauseScriptSeconds;
                row.Cells["PreDelay"].Value = customKeySlots[i].PreDelaySeconds;
            }

            // 按鍵欄位的 KeyDown 事件處理
            dgv.EditingControlShowing += (s, e) =>
            {
                if (dgv.CurrentCell?.ColumnIndex == dgv.Columns["KeyCode"].Index)
                {
                    TextBox tb = e.Control as TextBox;
                    if (tb != null)
                    {
                        tb.KeyDown -= CustomKeyDgv_KeyDown;
                        tb.KeyDown += CustomKeyDgv_KeyDown;
                        tb.Tag = dgv.CurrentCell;
                    }
                }
            };

            // 清除按鈕點擊事件
            dgv.CellClick += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == dgv.Columns["Clear"].Index)
                {
                    DataGridViewRow row = dgv.Rows[e.RowIndex];
                    row.Cells["Enabled"].Value = false;
                    row.Cells["KeyCode"].Value = "(點擊)";
                    row.Cells["KeyCode"].Tag = Keys.None;
                    row.Cells["Interval"].Value = 30.0;
                    row.Cells["StartAt"].Value = 0.0;
                    row.Cells["PauseEnabled"].Value = true;
                    row.Cells["PauseSeconds"].Value = 3.0;
                    row.Cells["PreDelay"].Value = 0.0;
                }
            };

            // 底部按鈕面板
            Panel btnPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.FromArgb(50, 50, 55)
            };

            Label lblHint = new Label
            {
                Text = "執行順序: 暫停腳本 → 按下按鍵 → 延遲等待 → 繼續腳本",
                Left = 10,
                Top = 5,
                Width = 350,
                ForeColor = Color.Yellow,
                Font = new Font("Microsoft JhengHei UI", 9F)
            };

            Button btnSave = new Button
            {
                Text = "✓ 儲存",
                Left = 370,
                Top = 15,
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold)
            };

            Button btnCancel = new Button
            {
                Text = "取消",
                Left = 480,
                Top = 15,
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(80, 80, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnSave.Click += (s, args) =>
            {
                SaveCustomKeySettings(dgv);
                customForm.Close();
            };

            btnCancel.Click += (s, args) => customForm.Close();

            // ★ 視窗關閉時自動儲存
            customForm.FormClosing += (s, args) =>
            {
                SaveCustomKeySettings(dgv);
            };

            btnPanel.Controls.AddRange(new Control[] { lblHint, btnSave, btnCancel });

            // 加入控制項（順序重要：先加底部，再加頂部，最後加 Fill）
            customForm.Controls.Add(dgv);
            customForm.Controls.Add(lblTitle);
            customForm.Controls.Add(btnPanel);

            customForm.ShowDialog();
        }


        /// <summary>
        /// 儲存自定義按鍵設定（從 DataGridView 讀取並儲存）
        /// </summary>
        private void SaveCustomKeySettings(DataGridView dgv)
        {
            try
            {
                for (int i = 0; i < 15; i++)
                {
                    DataGridViewRow row = dgv.Rows[i];
                    customKeySlots[i].Enabled = Convert.ToBoolean(row.Cells["Enabled"].Value);
                    customKeySlots[i].KeyCode = row.Cells["KeyCode"].Tag as Keys? ?? Keys.None;
                    customKeySlots[i].IntervalSeconds = Convert.ToDouble(row.Cells["Interval"].Value);
                    customKeySlots[i].StartAtSecond = Convert.ToDouble(row.Cells["StartAt"].Value);
                    customKeySlots[i].PauseScriptEnabled = Convert.ToBoolean(row.Cells["PauseEnabled"].Value);
                    customKeySlots[i].PauseScriptSeconds = Convert.ToDouble(row.Cells["PauseSeconds"].Value);
                    customKeySlots[i].PreDelaySeconds = Convert.ToDouble(row.Cells["PreDelay"].Value);
                }

                int enabledCount = customKeySlots.Count(slot => slot.Enabled && slot.KeyCode != Keys.None);
                AddLog($"自定義按鍵設定已儲存：{enabledCount} 個已啟用");
                SaveSettings();
            }
            catch (Exception ex)
            {
                AddLog($"自定義按鍵儲存失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 自定義按鍵 DataGridView 按鍵捕獲事件
        /// </summary>
        private void CustomKeyDgv_KeyDown(object sender, KeyEventArgs e)
        {
            TextBox tb = sender as TextBox;
            if (tb != null && tb.Tag is DataGridViewCell cell)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                tb.Text = GetKeyDisplayName(e.KeyCode);
                cell.Tag = e.KeyCode;
            }
        }

        /// <summary>
        /// 開啟定時執行設定視窗
        /// </summary>
        private void OpenSchedulerSettings()
        {
            Form schedForm = new Form
            {
                Text = "⏰ 定時執行設定",
                Width = 400,
                Height = 280,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            Label lblTitle = new Label
            {
                Text = "設定時間自動開始播放腳本",
                Left = 20,
                Top = 20,
                Width = 350,
                ForeColor = Color.LightGray,
                Font = new Font("Microsoft JhengHei UI", 10F)
            };

            // 時間選擇
            Label lblTime = new Label { Text = "執行時間：", Left = 20, Top = 60, Width = 80, ForeColor = Color.White };
            DateTimePicker dtpTime = new DateTimePicker
            {
                Left = 110,
                Top = 57,
                Width = 200,
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm:ss",
                Value = DateTime.Now.AddMinutes(5)
            };

            // 當前狀態
            Label lblStatus = new Label
            {
                Text = scheduledStartTime.HasValue
                    ? $"目前已排程：{scheduledStartTime.Value:HH:mm:ss}"
                    : "目前無排程",
                Left = 20,
                Top = 100,
                Width = 300,
                ForeColor = scheduledStartTime.HasValue ? Color.Lime : Color.Gray
            };

            // 按鈕
            Button btnSet = new Button
            {
                Text = "設定排程",
                Left = 50,
                Top = 150,
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(0, 150, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnClear = new Button
            {
                Text = "取消排程",
                Left = 160,
                Top = 150,
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(150, 80, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnClose = new Button
            {
                Text = "關閉",
                Left = 270,
                Top = 150,
                Width = 80,
                Height = 35,
                BackColor = Color.FromArgb(80, 80, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnSet.Click += (s, args) =>
            {
                if (dtpTime.Value <= DateTime.Now)
                {
                    MessageBox.Show("請選擇未來的時間！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (recordedEvents.Count == 0)
                {
                    MessageBox.Show("請先載入或錄製腳本！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                scheduledStartTime = dtpTime.Value;
                schedulerTimer.Start();
                lblStatus.Text = $"已排程：{scheduledStartTime.Value:yyyy-MM-dd HH:mm:ss}";
                lblStatus.ForeColor = Color.Lime;
                AddLog($"定時執行已設定：{scheduledStartTime.Value:HH:mm:ss}");
                MessageBox.Show($"已設定在 {scheduledStartTime.Value:HH:mm:ss} 自動開始播放！", "設定完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            btnClear.Click += (s, args) =>
            {
                scheduledStartTime = null;
                schedulerTimer.Stop();
                lblStatus.Text = "目前無排程";
                lblStatus.ForeColor = Color.Gray;
                AddLog("定時執行已取消");
            };

            btnClose.Click += (s, args) => schedForm.Close();

            // 倒數顯示
            Label lblCountdown = new Label
            {
                Left = 20,
                Top = 200,
                Width = 350,
                ForeColor = Color.Yellow
            };

            System.Windows.Forms.Timer countdownTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            countdownTimer.Tick += (s, args) =>
            {
                if (scheduledStartTime.HasValue)
                {
                    TimeSpan remaining = scheduledStartTime.Value - DateTime.Now;
                    if (remaining.TotalSeconds > 0)
                    {
                        lblCountdown.Text = $"距離執行還有：{remaining:hh\\:mm\\:ss}";
                    }
                    else
                    {
                        lblCountdown.Text = "正在執行...";
                    }
                }
                else
                {
                    lblCountdown.Text = "";
                }
            };
            countdownTimer.Start();

            schedForm.FormClosing += (s, args) => countdownTimer.Stop();
            schedForm.Controls.AddRange(new Control[] { lblTitle, lblTime, dtpTime, lblStatus, btnSet, btnClear, btnClose, lblCountdown });
            schedForm.ShowDialog();
        }

        /// <summary>
        /// 顯示執行統計
        /// </summary>
        private void ShowStatistics()
        {
            string stats = statistics.GetFormattedStats();

            // 加入自定義按鍵統計
            string customKeyStats = "\n\n自定義按鍵觸發次數：";
            for (int i = 0; i < 15; i++)
            {
                if (customKeySlots[i].Enabled && customKeySlots[i].KeyCode != Keys.None)
                {
                    customKeyStats += $"\n  #{i + 1} {GetKeyDisplayName(customKeySlots[i].KeyCode)}: {statistics.CustomKeyTriggerCounts[i]} 次";
                }
            }

            if (customKeySlots.Any(s => s.Enabled && s.KeyCode != Keys.None))
            {
                stats += customKeyStats;
            }

            DialogResult result = MessageBox.Show(
                stats + "\n\n是否重置統計？",
                "📊 執行統計",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                statistics.Reset();
                AddLog("統計資料已重置");
                MessageBox.Show("統計資料已重置！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
            return Environment.TickCount / 1000.0;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lblStatus.Text = "就緒：點擊「開始錄製」開始";
            lblRecordingStatus.Text = "錄製：尚未開始";
            lblPlaybackStatus.Text = "播放：尚未開始";
        }

        [Serializable]
        public class MacroEvent
        {
            public Keys KeyCode { get; set; }
            public string EventType { get; set; }
            public double Timestamp { get; set; }
        }

        private void btnStartRecording_Click_1(object sender, EventArgs e)
        {

        }

        private void btnCustomKeys_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 儲存設定按鈕點擊事件
        /// </summary>
        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            ExportSettings();
        }

        private void btnImportSettings_Click(object sender, EventArgs e)
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
                MessageBox.Show("設定已成功導出！\n\n包含：\n• 熱鍵設定\n• 視窗標題\n• 循環次數\n• 自定義按鍵設定\n• 方向鍵模式",
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

                MessageBox.Show("設定已成功匯入！\n\n包含：\n• 熱鍵設定\n• 視窗標題\n• 循環次數\n• 自定義按鍵設定\n• 方向鍵模式",
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

            string directory = Path.GetDirectoryName(filePath);
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

        private AppSettings BuildSettings()
        {
            var settings = new AppSettings
            {
                PlayHotkey = playHotkey,
                StopHotkey = stopHotkey,
                HotkeyEnabled = hotkeyEnabled,
                WindowTitle = txtWindowTitle.Text,
                LoopCount = (int)numPlayTimes.Value,
                ArrowKeyMode = (int)currentArrowKeyMode,
                BackgroundSwitchMode = (int)currentBackgroundSwitchMode
            };

            for (int i = 0; i < 15; i++)
            {
                settings.CustomKeySlots[i] = new CustomKeySlotData
                {
                    SlotNumber = customKeySlots[i].SlotNumber,
                    KeyCode = (int)customKeySlots[i].KeyCode,
                    IntervalSeconds = customKeySlots[i].IntervalSeconds,
                    Enabled = customKeySlots[i].Enabled,
                    StartAtSecond = customKeySlots[i].StartAtSecond,
                    PreDelaySeconds = customKeySlots[i].PreDelaySeconds,
                    PauseScriptSeconds = customKeySlots[i].PauseScriptSeconds,
                    PauseScriptEnabled = customKeySlots[i].PauseScriptEnabled
                };
            }

            return settings;
        }

        private void ApplySettings(AppSettings settings)
        {
            playHotkey = settings.PlayHotkey;
            stopHotkey = settings.StopHotkey;
            hotkeyEnabled = settings.HotkeyEnabled;
            txtWindowTitle.Text = settings.WindowTitle;
            numPlayTimes.Value = Math.Max(1, Math.Min(9999, settings.LoopCount));
            currentArrowKeyMode = (ArrowKeyMode)settings.ArrowKeyMode;
            if (currentArrowKeyMode != ArrowKeyMode.ThreadAttachWithBlocker)
            {
                currentArrowKeyMode = ArrowKeyMode.ThreadAttachWithBlocker;
            }

            // 載入背景切換模式
            currentBackgroundSwitchMode = (BackgroundSwitchMode)settings.BackgroundSwitchMode;
            if (!Enum.IsDefined(typeof(BackgroundSwitchMode), currentBackgroundSwitchMode))
            {
                currentBackgroundSwitchMode = BackgroundSwitchMode.ArrowKey_10s; // 預設 1.0 秒
            }

            if (settings.CustomKeySlots != null)
            {
                for (int i = 0; i < Math.Min(15, settings.CustomKeySlots.Length); i++)
                {
                    var data = settings.CustomKeySlots[i];
                    if (data != null)
                    {
                        customKeySlots[i].SlotNumber = data.SlotNumber;
                        customKeySlots[i].KeyCode = (Keys)data.KeyCode;
                        customKeySlots[i].IntervalSeconds = data.IntervalSeconds;
                        customKeySlots[i].Enabled = data.Enabled;
                        customKeySlots[i].StartAtSecond = data.StartAtSecond;
                        customKeySlots[i].PreDelaySeconds = data.PreDelaySeconds;
                        customKeySlots[i].PauseScriptSeconds = data.PauseScriptSeconds;
                        customKeySlots[i].PauseScriptEnabled = data.PauseScriptEnabled;
                    }
                }
            }

            AddLog("設定已載入");
            AddLog($"熱鍵：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}");
            AddLog($"背景切換模式：{currentBackgroundSwitchMode}");

            int enabledCount = customKeySlots.Count(s => s.Enabled && s.KeyCode != Keys.None);
            if (enabledCount > 0)
            {
                AddLog($"自定義按鍵：{enabledCount} 個已啟用");
            }

            UpdateUI();
            UpdateArrowKeyBlockerState();
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
    }
}
