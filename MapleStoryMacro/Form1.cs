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
        private volatile bool isPlaying = false;
        private double recordStartTime = 0;
        private IntPtr targetWindowHandle = IntPtr.Zero;
        private KeyboardHookDLL keyboardHook;

        // 高精度計時器（取代 Environment.TickCount，精度從 ~15ms 提升至 ~1μs）
        private readonly Stopwatch highResTimer = Stopwatch.StartNew();

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
            AddLog($"全局熱鍵：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}");
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
                    if (isPlaying)
                    {
                        AddLog($"排程結束：已到達結束時間 {task.EndTime.Value:HH:mm:ss}");
                        BtnStopPlayback_Click(this, EventArgs.Empty);
                    }

                    // 執行回程序列（在背景線程上執行，避免阻塞 UI）
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
                if (recordStartTime == 0)
                    recordStartTime = GetCurrentTime();

                string eventType = isKeyDown ? "down" : "up";

                recordedEvents.Add(new MacroEvent
                {
                    KeyCode = keyCode,
                    EventType = eventType,
                    Timestamp = GetCurrentTime() - recordStartTime
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
                    foreach (var evt in recordedEvents)
                    {
                        scriptData.Events.Add(new ScriptEvent
                        {
                            KeyCode = (int)evt.KeyCode,
                            EventType = evt.EventType,
                            Timestamp = evt.Timestamp
                        });
                    }

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
                    AddLog($"✅ 已保存: {Path.GetFileName(sfd.FileName)} ({recordedEvents.Count} 事件, {enabledCustomKeys} 自定義按鍵)");
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
                            Timestamp = evt.Timestamp
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
                    AddLog($"✅ 已載入: {Path.GetFileName(filePath)} ({recordedEvents.Count} 事件, {enabledCustomKeys} 自定義按鍵)");
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

            AddLog("正在開啟編輯器...");

            Form editorForm = new Form
            {
                Text = $"編輯腳本 ({recordedEvents.Count} 個事件)",
                Width = 900,
                Height = 620,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this
            };

            // 提示標籤
            Label hintLabel = new Label
            {
                Text = "★ 雙擊折疊/展開同類事件 | 選中「按鍵」欄後按鍵更改 | F2 編輯間隔",
                Top = 10,
                Left = 10,
                Width = 860,
                ForeColor = Color.Blue
            };

            // 標記是否有未儲存的變更
            bool hasUnsavedChanges = false;

            // 建立事件的編輯副本
            var editEvents = recordedEvents.Select(ev => new MacroEvent
            {
                KeyCode = ev.KeyCode,
                EventType = ev.EventType,
                Timestamp = ev.Timestamp
            }).ToList();

            // ===== 分組邏輯：連續相同按鍵+相同類型歸為一組 =====
            // groupStarts[i] = 第 i 組在 editEvents 中的起始索引
            // groupCounts[i] = 第 i 組的事件數量
            var groupStarts = new List<int>();
            var groupCounts = new List<int>();
            var expandedGroups = new HashSet<int>(); // 已展開的組別索引

            Action rebuildGroups = () =>
            {
                groupStarts.Clear();
                groupCounts.Clear();
                expandedGroups.Clear();
                int i = 0;
                while (i < editEvents.Count)
                {
                    int start = i;
                    var key = editEvents[i].KeyCode;
                    var type = editEvents[i].EventType;
                    while (i < editEvents.Count && editEvents[i].KeyCode == key && editEvents[i].EventType == type)
                        i++;
                    groupStarts.Add(start);
                    groupCounts.Add(i - start);
                }
            };
            rebuildGroups();

            DataGridView dgv = new DataGridView
            {
                Top = 35,
                Left = 10,
                Width = 860,
                Height = 470,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                EditMode = DataGridViewEditMode.EditOnF2
            };

            dgv.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn
                {
                    Name = "Index", HeaderText = "#", ReadOnly = true, FillWeight = 8
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "KeyCode", HeaderText = "按鍵 (選中後按鍵更改)", ReadOnly = true, FillWeight = 28
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "EventType", HeaderText = "類型", ReadOnly = true, FillWeight = 12
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Timestamp", HeaderText = "間隔 (秒)", FillWeight = 25,
                    DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(240, 248, 255) }
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Count", HeaderText = "數量", ReadOnly = true, FillWeight = 8,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        ForeColor = Color.DarkBlue,
                        Alignment = DataGridViewContentAlignment.MiddleCenter
                    }
                }
            });

            // Row.Tag 格式：int[] { groupIdx, eventIdx }
            //   eventIdx == -1 表示折疊的群組標題列
            //   eventIdx >= 0 表示 editEvents 中的實際索引

            // 刷新所有列（根據分組與展開狀態）
            Action refreshRows = () =>
            {
                dgv.SuspendLayout();
                dgv.Rows.Clear();

                for (int gi = 0; gi < groupStarts.Count; gi++)
                {
                    int start = groupStarts[gi];
                    int count = groupCounts[gi];
                    bool isSingle = count == 1;
                    bool isExpanded = expandedGroups.Contains(gi);

                    if (isSingle)
                    {
                        // 單一事件：直接顯示（間隔 = 與前一事件的時間差）
                        var evt = editEvents[start];
                        string keyName = GetKeyDisplayName(evt.KeyCode);
                        string eventType = evt.EventType == "down" ? "▼ 按下" : "▲ 放開";
                        double evtDelta = start == 0 ? evt.Timestamp : evt.Timestamp - editEvents[start - 1].Timestamp;

                        int ri = dgv.Rows.Add((start + 1).ToString(), keyName, eventType, evtDelta.ToString("F3"), "");
                        dgv.Rows[ri].Tag = new int[] { gi, start };
                        dgv.Rows[ri].Cells["Timestamp"].ReadOnly = false;

                        if (evt.EventType == "up")
                            dgv.Rows[ri].DefaultCellStyle.BackColor = Color.FromArgb(255, 250, 243);
                    }
                    else if (isExpanded)
                    {
                        // 展開的群組：顯示標題列 + 所有子事件
                        var firstEvt = editEvents[start];
                        string headerKey = $"▾ {GetKeyDisplayName(firstEvt.KeyCode)}";
                        string headerType = firstEvt.EventType == "down" ? "▼ 按下" : "▲ 放開";

                        // 標題列顯示所有子事件間隔的總和
                        double groupTotal = 0;
                        for (int j = 0; j < count; j++)
                        {
                            int ei2 = start + j;
                            groupTotal += ei2 == 0 ? editEvents[ei2].Timestamp : editEvents[ei2].Timestamp - editEvents[ei2 - 1].Timestamp;
                        }

                        int hri = dgv.Rows.Add("", headerKey, headerType, groupTotal.ToString("F3"), $"×{count}");
                        dgv.Rows[hri].Tag = new int[] { gi, -1 };
                        dgv.Rows[hri].Cells["Timestamp"].ReadOnly = true;
                        dgv.Rows[hri].DefaultCellStyle.BackColor = firstEvt.EventType == "down"
                            ? Color.FromArgb(218, 230, 248) : Color.FromArgb(248, 232, 218);
                        dgv.Rows[hri].DefaultCellStyle.Font = new Font(dgv.Font, FontStyle.Bold);

                        // 子事件列
                        for (int j = 0; j < count; j++)
                        {
                            int ei = start + j;
                            var evt = editEvents[ei];
                            string childKey = $"    {GetKeyDisplayName(evt.KeyCode)}";
                            string childType = evt.EventType == "down" ? "▼ 按下" : "▲ 放開";
                            double childDelta = ei == 0 ? evt.Timestamp : evt.Timestamp - editEvents[ei - 1].Timestamp;

                            int cri = dgv.Rows.Add((ei + 1).ToString(), childKey, childType, childDelta.ToString("F3"), "");
                            dgv.Rows[cri].Tag = new int[] { gi, ei };
                            dgv.Rows[cri].Cells["Timestamp"].ReadOnly = false;

                            if (evt.EventType == "up")
                                dgv.Rows[cri].DefaultCellStyle.BackColor = Color.FromArgb(255, 250, 243);
                        }
                    }
                    else
                    {
                        // 折疊的群組：只顯示標題列（間隔 = 所有子事件間隔的總和）
                        var firstEvt = editEvents[start];
                        string headerKey = $"▸ {GetKeyDisplayName(firstEvt.KeyCode)}";
                        string headerType = firstEvt.EventType == "down" ? "▼ 按下" : "▲ 放開";

                        double groupTotal = 0;
                        for (int j = 0; j < count; j++)
                        {
                            int ei = start + j;
                            groupTotal += ei == 0 ? editEvents[ei].Timestamp : editEvents[ei].Timestamp - editEvents[ei - 1].Timestamp;
                        }

                        int ri = dgv.Rows.Add("", headerKey, headerType, groupTotal.ToString("F3"), $"×{count}");
                        dgv.Rows[ri].Tag = new int[] { gi, -1 };
                        dgv.Rows[ri].Cells["Timestamp"].ReadOnly = true;
                        dgv.Rows[ri].DefaultCellStyle.BackColor = firstEvt.EventType == "down"
                            ? Color.FromArgb(225, 235, 250) : Color.FromArgb(250, 238, 225);
                        dgv.Rows[ri].DefaultCellStyle.Font = new Font(dgv.Font, FontStyle.Bold);
                    }
                }

                dgv.ResumeLayout();
            };
            refreshRows();

            // 雙擊切換折疊/展開
            dgv.CellDoubleClick += (s, args) =>
            {
                if (args.RowIndex < 0 || args.RowIndex >= dgv.Rows.Count) return;
                // 雙擊時間欄位時不切換（讓使用者編輯）
                if (args.ColumnIndex == dgv.Columns["Timestamp"]!.Index) return;

                if (dgv.Rows[args.RowIndex].Tag is int[] tag && tag.Length == 2)
                {
                    int gi = tag[0];
                    if (gi < groupStarts.Count && groupCounts[gi] > 1)
                    {
                        if (expandedGroups.Contains(gi))
                            expandedGroups.Remove(gi);
                        else
                            expandedGroups.Add(gi);
                        refreshRows();
                    }
                }
            };

            // 攔截延伸鍵（方向鍵等），讓按鍵欄能捕獲
            dgv.PreviewKeyDown += (s, args) =>
            {
                if (dgv.CurrentCell?.ColumnIndex == dgv.Columns["KeyCode"]!.Index)
                {
                    args.IsInputKey = true;
                }
            };

            // 按鍵感應：選中按鍵欄後按下按鍵即可更改
            dgv.KeyDown += (s, args) =>
            {
                if (dgv.CurrentCell?.ColumnIndex != dgv.Columns["KeyCode"]!.Index)
                    return;
                if (dgv.SelectedRows.Count == 0) return;

                args.Handled = true;
                args.SuppressKeyPress = true;

                Keys newKey = args.KeyCode;

                foreach (DataGridViewRow row in dgv.SelectedRows)
                {
                    if (row.Tag is int[] tag && tag.Length == 2)
                    {
                        int gi = tag[0], ei = tag[1];
                        if (ei == -1)
                        {
                            // 群組標題：更改整組所有事件
                            int gStart = groupStarts[gi];
                            int gCount = groupCounts[gi];
                            for (int j = 0; j < gCount; j++)
                                editEvents[gStart + j].KeyCode = newKey;
                        }
                        else
                        {
                            // 單一事件
                            editEvents[ei].KeyCode = newKey;
                        }
                        hasUnsavedChanges = true;
                    }
                }

                // 按鍵可能改變分組結構，重建
                rebuildGroups();
                refreshRows();
            };

            // 驗證時間欄位
            dgv.CellValidating += (s, args) =>
            {
                if (dgv.Columns[args.ColumnIndex].Name != "Timestamp") return;
                if (dgv.Rows[args.RowIndex].Cells["Timestamp"].ReadOnly) return;

                string value = args.FormattedValue?.ToString() ?? "";
                if (!string.IsNullOrEmpty(value) && !double.TryParse(value, out _))
                {
                    args.Cancel = true;
                    dgv.CancelEdit();
                    MessageBox.Show("請輸入有效的數字！", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (double.TryParse(value, out double ts) && ts < 0)
                {
                    args.Cancel = true;
                    dgv.CancelEdit();
                    MessageBox.Show("間隔不能為負數！", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            // 時間修改後更新
            dgv.CellValueChanged += (s, args) =>
            {
                if (args.RowIndex < 0 || args.RowIndex >= dgv.Rows.Count) return;
                if (dgv.Columns[args.ColumnIndex].Name != "Timestamp") return;

                if (dgv.Rows[args.RowIndex].Tag is int[] tag && tag.Length == 2 && tag[1] >= 0)
                {
                    int ei = tag[1];
                    string value = dgv.Rows[args.RowIndex].Cells["Timestamp"].Value?.ToString() ?? "0";
                    if (double.TryParse(value, out double newDelta) && ei < editEvents.Count)
                    {
                        // 使用者輸入的是間隔（delta），轉換回絕對時間戳
                        double prevTs = ei == 0 ? 0 : editEvents[ei - 1].Timestamp;
                        editEvents[ei].Timestamp = prevTs + newDelta;
                        hasUnsavedChanges = true;

                        // 更新後續事件的絕對時間戳（保持原始間隔不變）
                        for (int subsequent = ei + 1; subsequent < editEvents.Count; subsequent++)
                        {
                            // 後續事件不需要調整，因為它們的絕對時間戳是獨立的
                            // 只有當前事件的絕對時間改變
                            break;
                        }

                        refreshRows();
                    }
                }
            };

            Panel btnPanel = new Panel
            {
                Top = 515,
                Left = 10,
                Width = 860,
                Height = 50,
                BorderStyle = BorderStyle.FixedSingle
            };

            Button deleteBtn = new Button { Text = "刪除選中", Width = 100, Height = 30, Left = 10, Top = 10 };
            Button expandAllBtn = new Button { Text = "全部展開", Width = 85, Height = 30, Left = 340, Top = 10 };
            Button collapseAllBtn = new Button { Text = "全部折疊", Width = 85, Height = 30, Left = 430, Top = 10 };
            Button saveBtn = new Button { Text = "💾 儲存", Width = 100, Height = 30, Left = 120, Top = 10, ForeColor = Color.Green };
            Button closeBtn = new Button { Text = "關閉", Width = 100, Height = 30, Left = 230, Top = 10 };

            Label infoLabel = new Label
            {
                Text = $"事件: {editEvents.Count} | 群組: {groupStarts.Count}",
                Left = 530,
                Top = 15,
                Width = 320,
                ForeColor = Color.Gray
            };

            expandAllBtn.Click += (s, args) =>
            {
                for (int gi = 0; gi < groupStarts.Count; gi++)
                {
                    if (groupCounts[gi] > 1)
                        expandedGroups.Add(gi);
                }
                refreshRows();
            };

            collapseAllBtn.Click += (s, args) =>
            {
                expandedGroups.Clear();
                refreshRows();
            };

            deleteBtn.Click += (s, args) =>
            {
                if (dgv.SelectedRows.Count == 0) return;

                var indicesToRemove = new HashSet<int>();
                foreach (DataGridViewRow row in dgv.SelectedRows)
                {
                    if (row.Tag is int[] tag && tag.Length == 2)
                    {
                        int gi = tag[0], ei = tag[1];
                        if (ei == -1)
                        {
                            // 群組標題：刪除整組
                            int gStart = groupStarts[gi];
                            int gCount = groupCounts[gi];
                            for (int j = 0; j < gCount; j++)
                                indicesToRemove.Add(gStart + j);
                        }
                        else
                        {
                            indicesToRemove.Add(ei);
                        }
                    }
                }

                if (indicesToRemove.Count == 0) return;

                // 從後往前刪除
                foreach (int idx in indicesToRemove.OrderByDescending(x => x))
                {
                    editEvents.RemoveAt(idx);
                }

                hasUnsavedChanges = true;
                rebuildGroups();
                refreshRows();
                infoLabel.Text = $"事件: {editEvents.Count} | 群組: {groupStarts.Count}";
                AddLog($"✅ 已刪除 {indicesToRemove.Count} 個事件");
            };

            saveBtn.Click += (s, args) =>
            {
                recordedEvents.Clear();
                recordedEvents.AddRange(editEvents.Select(ev => new MacroEvent
                {
                    KeyCode = ev.KeyCode,
                    EventType = ev.EventType,
                    Timestamp = ev.Timestamp
                }));

                recordedEvents.Sort((a, b) => a.Timestamp.CompareTo(b.Timestamp));

                hasUnsavedChanges = false;
                lblRecordingStatus.Text = $"已編輯 | 事件數: {recordedEvents.Count}";
                infoLabel.Text = $"事件: {editEvents.Count} | 群組: {groupStarts.Count}";
                AddLog($"✅ 已儲存編輯 ({recordedEvents.Count} 個事件)");
                MessageBox.Show($"已儲存！共 {recordedEvents.Count} 個事件。", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            closeBtn.Click += (s, args) =>
            {
                if (hasUnsavedChanges)
                {
                    var result = MessageBox.Show("有未儲存的變更，是否儲存？", "未儲存的變更",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes)
                    {
                        saveBtn.PerformClick();
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

            btnPanel.Controls.Add(deleteBtn);
            btnPanel.Controls.Add(saveBtn);
            btnPanel.Controls.Add(closeBtn);
            btnPanel.Controls.Add(expandAllBtn);
            btnPanel.Controls.Add(collapseAllBtn);
            btnPanel.Controls.Add(infoLabel);

            editorForm.Controls.Add(hintLabel);
            editorForm.Controls.Add(dgv);
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
                    statistics.IncrementLoop();

                    int currentLoop = loop; // 避免閉包捕獲迴圈變數
                    this.BeginInvoke(new Action(() =>
                    {
                        lblPlaybackStatus.Text = $"循環: {currentLoop}/{loopCount}";
                        lblPlaybackStatus.ForeColor = Color.Blue;
                        if (currentLoop == 1 || currentLoop % 10 == 0)
                            AddLog($"循環 {currentLoop}/{loopCount} 開始");
                    }));

                    double lastTimestamp = 0;
                    long loopStartTick = Stopwatch.GetTimestamp();
                    long lastCustomKeyCheckTick = loopStartTick;

                    foreach (MacroEvent evt in events)
                    {
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
                }));
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
        /// </summary>
        private void SendCustomKey(Keys key, Keys modifiers = Keys.None)
        {
            try
            {
                if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
                {
                    // 背景模式
                    SendKeyWithThreadAttach(targetWindowHandle, key, true);
                    Thread.Sleep(30);
                    SendKeyWithThreadAttach(targetWindowHandle, key, false);
                }
                else
                {
                    // 前景模式
                    SendKeyForeground(key, true);
                    Thread.Sleep(30);
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
                    // 對於方向鍵，根據設定的模式發送
                    else if (IsArrowKey(evt.KeyCode))
                    {
                        SendArrowKeyWithMode(targetWindowHandle, evt.KeyCode, isDown);
                    }
                    // 英數鍵：使用 PostMessage（不攔截）
                    else if (IsAlphaNumericKey(evt.KeyCode))
                    {
                        SendKeyToWindow(targetWindowHandle, evt.KeyCode, isDown);
                    }
                    // 對於其他延伸鍵，使用線程附加模式
                    else if (IsExtendedKey(evt.KeyCode))
                    {
                        SendKeyWithThreadAttach(targetWindowHandle, evt.KeyCode, isDown);
                    }
                    else
                    {
                        // 一般按鍵：使用背景模式
                        SendKeyToWindow(targetWindowHandle, evt.KeyCode, isDown);
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
        /// 執行回程序列：Enter → 輸入指令 → Enter → 等待 → 坐下按鍵
        /// </summary>
        private void ExecuteReturnToTown(ScheduleTask task)
        {
            try
            {
                this.BeginInvoke(new Action(() => AddLog($"🏠 開始回程序列：{task.ReturnCommand}")));

                bool isBackground = targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle);

                // 1. 按下 Enter（開啟對話框）
                if (isBackground)
                {
                    SendKeyToWindow(targetWindowHandle, Keys.Enter, true);
                    Thread.Sleep(30);
                    SendKeyToWindow(targetWindowHandle, Keys.Enter, false);
                }
                else
                {
                    SendKeyForeground(Keys.Enter, true);
                    Thread.Sleep(30);
                    SendKeyForeground(Keys.Enter, false);
                }
                Thread.Sleep(300);

                // 2. 輸入回程指令（例如 @FM）
                if (isBackground)
                {
                    SendTextToWindow(targetWindowHandle, task.ReturnCommand);
                }
                else
                {
                    SendTextForeground(task.ReturnCommand);
                }
                Thread.Sleep(200);

                // 3. 按下 Enter（送出指令）
                if (isBackground)
                {
                    SendKeyToWindow(targetWindowHandle, Keys.Enter, true);
                    Thread.Sleep(30);
                    SendKeyToWindow(targetWindowHandle, Keys.Enter, false);
                }
                else
                {
                    SendKeyForeground(Keys.Enter, true);
                    Thread.Sleep(30);
                    SendKeyForeground(Keys.Enter, false);
                }

                this.BeginInvoke(new Action(() => AddLog($"📨 已送出指令：{task.ReturnCommand}")));

                // 4. 等待傳送完成
                Keys sitKey = (Keys)task.SitDownKeyCode;
                if (sitKey != Keys.None)
                {
                    int delayMs = (int)(task.SitDownDelaySeconds * 1000);
                    this.BeginInvoke(new Action(() => AddLog($"⏳ 等待 {task.SitDownDelaySeconds} 秒後坐下...")));
                    Thread.Sleep(delayMs);

                    // 5. 按下坐下按鍵
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
            isPlaying = false;
            lblPlaybackStatus.Text = "播放: 已停止";
            lblPlaybackStatus.ForeColor = Color.Orange;
            AddLog("播放已停止");
            UpdateUI();
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
            monitorTimer?.Stop();
            schedulerTimer?.Stop();   // 停止定時執行計時器

            // 自動儲存設定（靜默，不顯示錯誤對話框）
            try { SaveSettingsToFile(SettingsFilePath, ""); } catch { }

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
        }

        /// <summary>
        /// 開啟熱鍵設定視窗
        /// </summary>
        private void OpenHotkeySettings()
        {
            Form settingsForm = new Form
            {
                Text = "⚙ 熱鍵與進階設定",
                Width = 450,
                Height = 340,
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

            // 啟用熱鍵
            CheckBox chkEnabled = new CheckBox
            {
                Text = "啟用全局熱鍵",
                Left = 20,
                Top = 110,
                Width = 150,
                Checked = hotkeyEnabled,
                ForeColor = Color.White
            };

            // 方向鍵模式選擇
            Label lblArrowMode = new Label
            {
                Text = "方向鍵模式：",
                Left = 20,
                Top = 150,
                Width = 100,
                ForeColor = Color.Cyan
            };
            ComboBox cmbArrowMode = new ComboBox
            {
                Left = 130,
                Top = 147,
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(60, 60, 65),
                ForeColor = Color.White
            };

            // 方向鍵模式選項
            var availableModes = new (ArrowKeyMode mode, string name)[]
            {
                (ArrowKeyMode.SendToChild, "S2C (背景)"),
                (ArrowKeyMode.ThreadAttachWithBlocker, "TAB"),
                (ArrowKeyMode.SendInputWithBlock, "SWB")
            };

     foreach (var mode in availableModes)
            {
                cmbArrowMode.Items.Add(mode.name);
            }

            // 找到當前模式在列表中的索引
            int currentIndex = Array.FindIndex(availableModes, m => m.mode == currentArrowKeyMode);
            cmbArrowMode.SelectedIndex = currentIndex >= 0 ? currentIndex : 0;

            // 方向鍵模式說明
            Label lblArrowHint = new Label
            {
                Text = "S2C=背景(會洩漏) | TAB/SWB=嘗試攔截",
                Left = 20,
                Top = 180,
                Width = 400,
                ForeColor = Color.Yellow,
                Font = new Font("Microsoft JhengHei UI", 8F)
            };

            // 提示文字
            Label lblHint = new Label
            {
                Text = "提示：點擊文字框後按下想要的按鍵",
                Left = 20,
                Top = 210,
                Width = 350,
                ForeColor = Color.Gray
            };

            // 按鈕
            Button btnSave = new Button
            {
                Text = "儲存",
                Left = 180,
                Top = 250,
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
                Top = 250,
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

                // 從可用模式列表中取得實際的 enum 值
                currentArrowKeyMode = availableModes[cmbArrowMode.SelectedIndex].mode;

                AddLog($"設定已儲存：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}, 模式={currentArrowKeyMode}");
                settingsForm.Close();
            };

            btnCancel.Click += (s, args) => settingsForm.Close();

            settingsForm.Controls.AddRange(new Control[]
            {
                lblPlay, txtPlay, lblStop, txtStop, chkEnabled,
                lblArrowMode, cmbArrowMode, lblArrowHint,
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
                Font = new Font("Microsoft JhengHei UI", 9F)
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
                    Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(40, 40, 45),
                    ForeColor = Color.White,
                    SelectionBackColor = Color.FromArgb(0, 100, 180),
                    SelectionForeColor = Color.White,
                    Font = new Font("Microsoft JhengHei UI", 9F)
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
                Font = new Font("Microsoft JhengHei UI", 8.5F)
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
                Left = 20, Top = 15, Width = 600,
                ForeColor = Color.LightGray,
                Font = new Font("Microsoft JhengHei UI", 10F)
            };

            // ===== 新增排程區塊 =====
            Label lblNewTask = new Label
            {
                Text = "📋 新增排程",
                Left = 20, Top = 45, Width = 200,
                ForeColor = Color.Cyan,
                Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold)
            };

            // 腳本選擇
            Label lblScript = new Label { Text = "腳本：", Left = 20, Top = 75, Width = 50, ForeColor = Color.White };
            TextBox txtScriptPath = new TextBox
            {
                Left = 75, Top = 72, Width = 420, ReadOnly = true,
                BackColor = Color.FromArgb(60, 60, 65), ForeColor = Color.White,
                Text = currentScriptPath ?? "(使用當前已載入的腳本)"
            };
            Button btnBrowse = new Button
            {
                Text = "...", Left = 500, Top = 71, Width = 35, Height = 25,
                BackColor = Color.FromArgb(80, 80, 85), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            Button btnUseCurrent = new Button
            {
                Text = "當前", Left = 540, Top = 71, Width = 50, Height = 25,
                BackColor = Color.FromArgb(0, 100, 160), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
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
                Left = 75, Top = 105, Width = 200,
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm:ss",
                Value = DateTime.Now.AddMinutes(5)
            };

            // 結束時間
            Label lblEnd = new Label { Text = "結束：", Left = 290, Top = 108, Width = 50, ForeColor = Color.White };
            CheckBox chkEndTime = new CheckBox
            {
                Text = "", Left = 340, Top = 108, Width = 20,
                Checked = false, ForeColor = Color.White
            };
            DateTimePicker dtpEnd = new DateTimePicker
            {
                Left = 365, Top = 105, Width = 200,
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm:ss",
                Value = DateTime.Now.AddHours(1),
                Enabled = false
            };
            chkEndTime.CheckedChanged += (s, args) => dtpEnd.Enabled = chkEndTime.Checked;

            // 循環次數
            Label lblLoop = new Label { Text = "循環：", Left = 20, Top = 140, Width = 50, ForeColor = Color.White };
            NumericUpDown numLoop = new NumericUpDown
            {
                Left = 75, Top = 137, Width = 80,
                Minimum = 1, Maximum = 9999, Value = (int)numPlayTimes.Value,
                BackColor = Color.FromArgb(60, 60, 65), ForeColor = Color.White
            };

            Label lblLoopHint = new Label
            {
                Text = "勾選「結束」可設定自動停止時間",
                Left = 165, Top = 140, Width = 300,
                ForeColor = Color.Gray, Font = new Font("Microsoft JhengHei UI", 8.5F)
            };

            // ===== 回程設定區塊 =====
            CheckBox chkReturnToTown = new CheckBox
            {
                Text = "🏠 回程（結束時自動回城）",
                Left = 20, Top = 170, Width = 220,
                ForeColor = Color.FromArgb(100, 220, 160),
                Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold),
                Checked = false
            };

            Label lblReturnCmd = new Label { Text = "指令：", Left = 250, Top = 172, Width = 45, ForeColor = Color.White };
            TextBox txtReturnCmd = new TextBox
            {
                Left = 295, Top = 169, Width = 80,
                BackColor = Color.FromArgb(60, 60, 65), ForeColor = Color.LightGreen,
                Text = "@FM", Enabled = false
            };

            Label lblSitKey = new Label { Text = "坐下：", Left = 385, Top = 172, Width = 45, ForeColor = Color.White };
            TextBox txtSitKey = new TextBox
            {
                Left = 430, Top = 169, Width = 100, ReadOnly = true,
                BackColor = Color.FromArgb(60, 60, 65), ForeColor = Color.Cyan,
                Text = "(點擊設定)", Enabled = false,
                Tag = Keys.None
            };

            Label lblSitDelay = new Label { Text = "延遲：", Left = 20, Top = 200, Width = 45, ForeColor = Color.White };
            NumericUpDown numSitDelay = new NumericUpDown
            {
                Left = 65, Top = 197, Width = 60,
                Minimum = 0, Maximum = 30, Value = 3, DecimalPlaces = 1, Increment = 0.5m,
                BackColor = Color.FromArgb(60, 60, 65), ForeColor = Color.White,
                Enabled = false
            };
            Label lblSitDelayUnit = new Label
            {
                Text = "秒後坐下 | 序列：Enter → 指令 → Enter → 等待 → 坐下鍵",
                Left = 130, Top = 200, Width = 450,
                ForeColor = Color.Gray, Font = new Font("Microsoft JhengHei UI", 8.5F)
            };

            // 回程勾選控制啟用/停用
            chkReturnToTown.CheckedChanged += (s, args) =>
            {
                bool enabled = chkReturnToTown.Checked;
                txtReturnCmd.Enabled = enabled;
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
                Left = 20, Top = 230, Width = 120, Height = 30,
                BackColor = Color.FromArgb(0, 150, 80), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };

            // ===== 排程清單 =====
            Label lblListTitle = new Label
            {
                Text = "📅 排程清單",
                Left = 20, Top = 270, Width = 200,
                ForeColor = Color.Yellow,
                Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold)
            };

            DataGridView dgv = new DataGridView
            {
                Left = 20, Top = 295, Width = 590, Height = 220,
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
                    BackColor = Color.FromArgb(50, 50, 55), ForeColor = Color.White
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(40, 40, 45), ForeColor = Color.White,
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
                    ReturnCommand = txtReturnCmd.Text.Trim(),
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
                string returnInfo = chkReturnToTown.Checked ? $", 回程={txtReturnCmd.Text}" : "";
                AddLog($"新增排程：{(string.IsNullOrEmpty(scriptPath) ? "當前腳本" : Path.GetFileName(scriptPath))}, 開始={dtpStart.Value:HH:mm:ss}{endInfo}, 循環={numLoop.Value}{returnInfo}");
            };

            // 刪除與清空按鈕
            Button btnRemove = new Button
            {
                Text = "🗑️ 刪除選中",
                Left = 20, Top = 525, Width = 110, Height = 30,
                BackColor = Color.FromArgb(150, 60, 60), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };

            Button btnClearAll = new Button
            {
                Text = "清空全部",
                Left = 140, Top = 525, Width = 90, Height = 30,
                BackColor = Color.FromArgb(120, 80, 40), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };

            Button btnClose = new Button
            {
                Text = "關閉",
                Left = 530, Top = 525, Width = 80, Height = 30,
                BackColor = Color.FromArgb(80, 80, 85), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
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
                Left = 240, Top = 530, Width = 280,
                ForeColor = Color.Yellow, Font = new Font("Microsoft JhengHei UI", 9F)
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
                chkReturnToTown, lblReturnCmd, txtReturnCmd, lblSitKey, txtSitKey,
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
                Left = 20, Top = 20, Width = 400, Height = 25,
                ForeColor = statistics.CurrentSessionStart.HasValue ? Color.Lime : Color.Gray,
                Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Bold),
                Text = statistics.CurrentSessionStart.HasValue ? "● 播放中" : "○ 已停止"
            };

            // 當前會話區塊
            Label lblSessionTitle = new Label
            {
                Left = 20, Top = 55, Width = 200, Height = 20,
                ForeColor = Color.Cyan, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold),
                Text = "📌 當前會話"
            };

            Label lblSessionTime = new Label
            {
                Left = 30, Top = 80, Width = 380, Height = 20,
                ForeColor = Color.White, Font = new Font("Consolas", 10F),
                Text = "會話時長: --:--:--"
            };

            Label lblCurrentLoop = new Label
            {
                Left = 30, Top = 105, Width = 380, Height = 20,
                ForeColor = Color.White, Font = new Font("Consolas", 10F),
                Text = "當前循環: 0"
            };

            Label lblScriptInfo = new Label
            {
                Left = 30, Top = 130, Width = 380, Height = 20,
                ForeColor = Color.LightGray, Font = new Font("Microsoft JhengHei UI", 9F),
                Text = $"腳本事件: {recordedEvents.Count} 個"
            };

            // 分隔線
            Label lblSep1 = new Label
            {
                Left = 20, Top = 160, Width = 400, Height = 2,
                BackColor = Color.FromArgb(60, 60, 65)
            };

            // 累計統計區塊
            Label lblTotalTitle = new Label
            {
                Left = 20, Top = 170, Width = 200, Height = 20,
                ForeColor = Color.Yellow, Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold),
                Text = "📈 累計統計"
            };

            Label lblTotalPlays = new Label
            {
                Left = 30, Top = 195, Width = 380, Height = 20,
                ForeColor = Color.White, Font = new Font("Consolas", 10F),
                Text = $"播放次數: {statistics.TotalPlayCount}"
            };

            Label lblTotalTime = new Label
            {
                Left = 30, Top = 220, Width = 380, Height = 20,
                ForeColor = Color.White, Font = new Font("Consolas", 10F),
                Text = "總播放時長: 00:00:00"
            };

            Label lblLastPlay = new Label
            {
                Left = 30, Top = 245, Width = 380, Height = 20,
                ForeColor = Color.LightGray, Font = new Font("Microsoft JhengHei UI", 9F),
                Text = $"最後播放: {(statistics.LastPlayTime.HasValue ? statistics.LastPlayTime.Value.ToString("yyyy-MM-dd HH:mm:ss") : "從未播放")}"
            };

            // 分隔線
            Label lblSep2 = new Label
            {
                Left = 20, Top = 275, Width = 400, Height = 2,
                BackColor = Color.FromArgb(60, 60, 65)
            };

            // 自定義按鍵統計
            Label lblCustomTitle = new Label
            {
                Left = 20, Top = 285, Width = 200, Height = 20,
                ForeColor = Color.FromArgb(200, 150, 255), Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Bold),
                Text = "⚡ 自定義按鍵觸發"
            };

            ListBox lstCustomKeys = new ListBox
            {
                Left = 30, Top = 310, Width = 380, Height = 100,
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
                Left = 130, Top = 430, Width = 100, Height = 35,
                BackColor = Color.FromArgb(150, 80, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnClose = new Button
            {
                Text = "關閉",
                Left = 240, Top = 430, Width = 80, Height = 35,
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

        private AppSettings BuildSettings()
        {
            var settings = new AppSettings
            {
                PlayHotkey = playHotkey,
                StopHotkey = stopHotkey,
                HotkeyEnabled = hotkeyEnabled,
                WindowTitle = txtWindowTitle.Text,
                ArrowKeyMode = (int)currentArrowKeyMode,
                LastScriptPath = currentScriptPath,
                ScheduleTasks = scheduleTasks.Where(t => t.Enabled && t.StartTime > DateTime.Now).ToList()
            };

            return settings;
        }

        private void ApplySettings(AppSettings settings)
        {
            playHotkey = settings.PlayHotkey;
            stopHotkey = settings.StopHotkey;
            hotkeyEnabled = settings.HotkeyEnabled;
            txtWindowTitle.Text = settings.WindowTitle;

            // 自動降級：無效的模式值改為 SendToChild(0)
            if (settings.ArrowKeyMode > 2)
                settings.ArrowKeyMode = 0;
            currentArrowKeyMode = (ArrowKeyMode)settings.ArrowKeyMode;

            // 嘗試自動載入上次的腳本
            if (!string.IsNullOrEmpty(settings.LastScriptPath) && File.Exists(settings.LastScriptPath))
            {
                currentScriptPath = settings.LastScriptPath;
                LoadScriptFromFile(settings.LastScriptPath);
            }

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
            AddLog($"熱鍵：播放={GetKeyDisplayName(playHotkey)}, 停止={GetKeyDisplayName(stopHotkey)}");
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
    }
}