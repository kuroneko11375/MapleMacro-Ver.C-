using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
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
        private readonly int MAX_LOG_LINES = 100;
        private readonly object logLock = new object();
        private string? lastLogMessage = null; // 用於合併重複日誌
        private int lastLogRepeatCount = 0; // 重複計數

        private HashSet<Keys> pressedKeys = new HashSet<Keys>();

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
            SendInputWithBlock      // SendInput + Blocker（嘗試避免影響前景）
        }
        private ArrowKeyMode currentArrowKeyMode = ArrowKeyMode.SendInputWithBlock;
        
        // 鍵盤阻擋器（用於 Blocker 模式）
        private KeyboardBlocker? keyboardBlocker;

        // Windows message constants
        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP = 0x0101;
        private const uint WM_SYSKEYDOWN = 0x0104;
        private const uint WM_SYSKEYUP = 0x0105;
        private const uint MAPVK_VK_TO_VSC = 0;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

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
            keyboardHook = new KeyboardHookDLL();
            keyboardHook.OnKeyEvent += KeyboardHook_OnKeyEvent;

            // Initialize global hotkey hook (延遲安裝到 Form1_Shown)
            hotkeyHook = new KeyboardHookDLL();
            hotkeyHook.OnKeyEvent += HotkeyHook_OnKeyEvent;

            // 延遲初始化 - 在表單顯示後再載入設定和啟動定時器
            this.Shown += Form1_Shown;

            // 先設定初始狀態
            lblWindowStatus.Text = "視窗: 未鎖定";
            lblWindowStatus.ForeColor = Color.Gray;
            UpdateUI();
        }

        private void Form1_Shown(object sender, EventArgs e)
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


        private void MonitorTimer_Tick(object sender, EventArgs e)
        {
            // 更新狀態列標籤
            string bgMode = (targetWindowHandle != IntPtr.Zero) ? "背景" : "前景";
            string recStatus = isRecording ? "🔴" : "⚪";
            string playStatus = isPlaying ? "▶️" : "⏹️";

            lblStatus.Text = $"{recStatus} 錄製 | {playStatus} 播放 | 模式: {bgMode} | 事件: {recordedEvents.Count}";
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

                lblRecordingStatus.Text = $"錄製中: {recordedEvents.Count} 個事件 | 最後: {e.KeyCode}";
                AddLog($"按鍵按下: {e.KeyCode}");
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

        private void BtnRefreshWindow_Click(object sender, EventArgs e)
        {
            AddLog("正在搜尋目標視窗...");
            FindGameWindow();
            UpdateWindowStatus();
        }

        private void BtnLockWindow_Click(object sender, EventArgs e)
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
                lblWindowStatus.Text = $"視窗: 已鎖定 - 背景模式";
                lblWindowStatus.ForeColor = Color.Green;
                AddLog($"已鎖定視窗: {targetWindowHandle}");
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

        private void BtnStartRecording_Click(object sender, EventArgs e)
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

        private void BtnStopRecording_Click(object sender, EventArgs e)
        {
            if (!isRecording) return;

            isRecording = false;
            lblRecordingStatus.Text = $"已停止 | 事件數: {recordedEvents.Count}";
            lblRecordingStatus.ForeColor = Color.Green;

            keyboardHook.Uninstall();

            AddLog($"錄製已停止 - 共 {recordedEvents.Count} 個事件");
            UpdateUI();
        }

        private void BtnSaveScript_Click(object sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                AddLog("⚠️ 沒有事件可保存");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "JSON 腳本|*.json",
                DefaultExt = ".json",
                Title = "保存腳本"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string json = JsonSerializer.Serialize(recordedEvents, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(sfd.FileName, json);
                    AddLog($"✅ 已保存: {Path.GetFileName(sfd.FileName)}");
                }
                catch (Exception ex)
                {
                    AddLog($"❌ 保存失敗: {ex.Message}");
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
                    AddLog($"✅ 已載入: {Path.GetFileName(ofd.FileName)} ({recordedEvents.Count} 個事件)");

                    // 更新 UI 狀態，啟用開始播放按鍵
                    UpdateUI();
                }
                catch (Exception ex)
                {
                    AddLog($"❌ 載入失敗: {ex.Message}");
                }
            }
        }

        private void BtnClearEvents_Click(object sender, EventArgs e)
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
           MessageBoxDefaultButton.Button2  // 預設選擇「否」
            );

            if (result == DialogResult.Yes)
            {
                recordedEvents.Clear();
                lblRecordingStatus.Text = "已清除 | 事件數: 0";
                AddLog("✅ 已清除所有事件");

                // 更新 UI 狀態
                UpdateUI();
            }
        }

        private void BtnViewEvents_Click(object sender, EventArgs e)
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

        private void BtnEditEvents_Click(object sender, EventArgs e)
        {
            if (recordedEvents.Count == 0)
            {
                AddLog("⚠️ 沒有事件可編輯");
                return;
            }

            AddLog("正在開啟編輯器...");

            // 整合重複按鍵事件
            var consolidatedEvents = ConsolidateKeyEvents(recordedEvents);

            Form editorForm = new Form
            {
                Text = $"編輯腳本 (整合後: {consolidatedEvents.Count} 個動作)",
                Width = 850,
                Height = 600,
                StartPosition = FormStartPosition.CenterParent,
                Owner = this
            };

            // 提示標籤
            Label hintLabel = new Label
            {
                Text = "★ 已整合連續重複按鍵，顯示持續時間",
                Top = 10,
                Left = 10,
                Width = 400,
                ForeColor = Color.Blue
            };

            DataGridView dgv = new DataGridView
            {
                Top = 35,
                Left = 10,
                Width = 810,
                Height = 455,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            dgv.Columns.Add("KeyCode", "按鍵");
            dgv.Columns.Add("Duration", "持續時間");
            dgv.Columns.Add("StartTime", "開始時間 (秒)");
            dgv.Columns.Add("EndTime", "結束時間 (秒)");

            foreach (var evt in consolidatedEvents)
            {
                string keyName = GetKeyDisplayName(evt.KeyCode);
                string duration = evt.Duration >= 1.0
                    ? $"{evt.Duration:F2} 秒"
                    : $"{(evt.Duration * 1000):F0} ms";
                dgv.Rows.Add(keyName, duration, evt.StartTime.ToString("F3"), evt.EndTime.ToString("F3"));
            }

            Panel btnPanel = new Panel
            {
                Top = 500,
                Left = 10,
                Width = 810,
                Height = 50,
                BorderStyle = BorderStyle.FixedSingle
            };

            Button deleteBtn = new Button { Text = "刪除選中", Width = 100, Height = 30, Left = 10, Top = 10 };
            Button closeBtn = new Button { Text = "關閉", Width = 100, Height = 30, Left = 120, Top = 10 };

            Label infoLabel = new Label
            {
                Text = $"原始事件: {recordedEvents.Count} | 整合後: {consolidatedEvents.Count}",
                Left = 240,
                Top = 15,
                Width = 300,
                ForeColor = Color.Gray
            };

            deleteBtn.Click += (s, args) =>
            {
                if (dgv.SelectedRows.Count > 0)
                {
                    var selectedIndices = dgv.SelectedRows.Cast<DataGridViewRow>()
                        .Select(r => r.Index)
                        .OrderByDescending(i => i)
                        .ToList();

                    foreach (int index in selectedIndices)
                    {
                        if (index < consolidatedEvents.Count)
                        {
                            var evtToRemove = consolidatedEvents[index];
                            // 從原始事件中移除對應的事件
                            recordedEvents.RemoveAll(e =>
                                e.KeyCode == evtToRemove.KeyCode &&
                                e.Timestamp >= evtToRemove.StartTime &&
                                e.Timestamp <= evtToRemove.EndTime);
                            consolidatedEvents.RemoveAt(index);
                            dgv.Rows.RemoveAt(index);
                        }
                    }
                    lblRecordingStatus.Text = $"已編輯 | 事件數: {recordedEvents.Count}";
                    infoLabel.Text = $"原始事件: {recordedEvents.Count} | 整合後: {consolidatedEvents.Count}";
                    AddLog($"✅ 已刪除 {selectedIndices.Count} 個動作");
                }
            };

            closeBtn.Click += (s, args) => editorForm.Close();

            btnPanel.Controls.Add(deleteBtn);
            btnPanel.Controls.Add(closeBtn);
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
            public double StartTime { get; set; }
            public double EndTime { get; set; }
            public double Duration => EndTime - StartTime;
        }

        /// <summary>
        /// 將連續重複的按鍵事件整合為單一動作
        /// </summary>
        private List<ConsolidatedKeyEvent> ConsolidateKeyEvents(List<MacroEvent> events)
        {
            var consolidated = new List<ConsolidatedKeyEvent>();
            if (events.Count == 0) return consolidated;

            // 追蹤每個按鍵的按下時間
            var keyDownTimes = new Dictionary<Keys, double>();

            foreach (var evt in events.OrderBy(e => e.Timestamp))
            {
                if (evt.EventType == "down")
                {
                    // 記錄按下時間（如果尚未追蹤）
                    if (!keyDownTimes.ContainsKey(evt.KeyCode))
                    {
                        keyDownTimes[evt.KeyCode] = evt.Timestamp;
                    }
                }
                else if (evt.EventType == "up")
                {
                    // 放開時計算持續時間
                    if (keyDownTimes.TryGetValue(evt.KeyCode, out double startTime))
                    {
                        consolidated.Add(new ConsolidatedKeyEvent
                        {
                            KeyCode = evt.KeyCode,
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
                consolidated.Add(new ConsolidatedKeyEvent
                {
                    KeyCode = kvp.Key,
                    StartTime = kvp.Value,
                    EndTime = events.Max(e => e.Timestamp)
                });
            }

            return consolidated.OrderBy(e => e.StartTime).ToList();
        }

        private void BtnStartPlayback_Click(object sender, EventArgs e)
        {
            if (isPlaying || recordedEvents.Count == 0)
                return;

            isPlaying = true;

            pressedKeys.Clear();

            // 重置自定義按鍵槽位的觸發狀態
            foreach (var slot in customKeySlots)
            {
                slot.Reset();
            }

            // 開始統計
            statistics.StartSession();

            // 如果使用 Blocker 模式，初始化並啟動 KeyboardBlocker
            if (currentArrowKeyMode == ArrowKeyMode.ThreadAttachWithBlocker || 
                currentArrowKeyMode == ArrowKeyMode.SendInputWithBlock)
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
            AddLog($"播放開始 ({mode}模式)...");

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
                AddLog($"❌ 播放失敗: {ex.Message}");
                ReleasePressedKeys();
                isPlaying = false;
                statistics.EndSession();
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
                        lblPlaybackStatus.Text = $"循環: {loop}/{loopCount}";
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

                this.Invoke(new Action(() =>
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
                this.Invoke(new Action(() =>
                {
                    AddLog($"❌ 播放錯誤: {ex.Message}");
                }));
                ReleasePressedKeys();
                isPlaying = false;
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
                        AddLog($"背景(Alt): {evt.KeyCode} ({evt.EventType})");
                    }
                    // 對於方向鍵，根據設定的模式發送
                    else if (IsArrowKey(evt.KeyCode))
                    {
                        SendArrowKeyWithMode(targetWindowHandle, evt.KeyCode, evt.EventType == "down");
                        AddLog($"背景({currentArrowKeyMode}): {evt.KeyCode} ({evt.EventType})");
                    }
                    // 對於其他延伸鍵，使用線程附加模式
                    else if (IsExtendedKey(evt.KeyCode))
                    {
                        SendKeyWithThreadAttach(targetWindowHandle, evt.KeyCode, evt.EventType == "down");
                        AddLog($"背景(附加): {evt.KeyCode} ({evt.EventType})");
                    }
                    else
                    {
                        // 一般按鍵：使用背景模式
                        SendKeyToWindow(targetWindowHandle, evt.KeyCode, evt.EventType == "down");
                        AddLog($"背景: {evt.KeyCode} ({evt.EventType})");
                    }
                }
                else
                {
                    // Foreground key sending using keybd_event
                    SendKeyForeground(evt.KeyCode, evt.EventType == "down");
                    AddLog($"前景: {evt.KeyCode} ({evt.EventType})");
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
                    SendKeyWithThreadAttach(hWnd, key, isKeyDown);
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
        // 加上巨集標記，讓 Blocker 可以識別
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
            // 備援：使用 keybd_event，也要帶標記
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

                AddLog($"Alt 按鍵: VK=0x{vkCode:X2}, SC=0x{scanCode:X2}, 旗標=0x{flags:X}");
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

        private void BtnStopPlayback_Click(object sender, EventArgs e)
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
            keyboardBlocker?.Uninstall();  // 停止鍵盤阻擋器
            keyboardBlocker?.Dispose();
            monitorTimer?.Stop();
            schedulerTimer?.Stop();   // 停止定時執行計時器

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
        /// 開啟熱鍵設定視窗
        /// </summary>
        private void OpenHotkeySettings()
        {
            Form settingsForm = new Form
            {
                Text = "⚙ 熱鍵與進階設定",
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
                (ArrowKeyMode.SendInputWithBlock, "SWB (推薦)")
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
                Text = "S2C=背景 | TAB=TA+Blocker | SWB=SendInput+Blocker",
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
                Left = 130,
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
                Left = 220,
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

            // 按鍵 - 使用 TextBox，但透過事件處理按鍵輸入
            var colKey = new DataGridViewTextBoxColumn
            {
                Name = "KeyCode",
                HeaderText = "按鍵 (點擊後按鍵)",
                Width = 120,
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
                // 儲存 KeyCode 到 Tag
                dgv.Rows[i].Cells["KeyCode"].Tag = slot.KeyCode;
            }

            // 處理按鍵欄位的按鍵輸入
            dgv.EditingControlShowing += (s, args) =>
            {
                if (dgv.CurrentCell.ColumnIndex == dgv.Columns["KeyCode"].Index)
                {
                    TextBox tb = args.Control as TextBox;
                    if (tb != null)
                    {
                        // 移除舊的事件處理器
                        tb.KeyDown -= CustomKeyCell_KeyDown;
                        tb.KeyDown += CustomKeyCell_KeyDown;
                    }
                }
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

            // 按鍵欄位不可直接編輯文字
            dgv.CellBeginEdit += (s, args) =>
            {
                if (dgv.Columns[args.ColumnIndex].Name == "KeyCode")
                {
                    // 允許進入編輯模式以便捕獲按鍵
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
                        customKeySlots[i].KeyCode = (Keys)(row.Cells["KeyCode"].Tag ?? Keys.None);
                        customKeySlots[i].IntervalSeconds = double.TryParse(row.Cells["Interval"].Value?.ToString(), out double interval) ? interval : 30;
                        customKeySlots[i].StartAtSecond = double.TryParse(row.Cells["StartAt"].Value?.ToString(), out double startAt) ? startAt : 0;
                        customKeySlots[i].PauseScriptEnabled = Convert.ToBoolean(row.Cells["PauseEnabled"].Value);
                        customKeySlots[i].PauseScriptSeconds = double.TryParse(row.Cells["PauseSeconds"].Value?.ToString(), out double pause) ? pause : 0;
                        customKeySlots[i].PreDelaySeconds = double.TryParse(row.Cells["Delay"].Value?.ToString(), out double delay) ? delay : 0;
                    }

                    int enabledCount = customKeySlots.Count(slot => slot.Enabled && slot.KeyCode != Keys.None);
                    AddLog($"✅ 自定義按鍵設定已儲存：{enabledCount} 個已啟用");
                    SaveSettings();
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
        /// 自定義按鍵欄位的按鍵處理
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
                    // 儲存 KeyCode 到 Cell 的 Tag
                    dgv.CurrentCell.Tag = e.KeyCode;
                    // 結束編輯
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
                ArrowKeyMode = (int)currentArrowKeyMode
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
            AddLog($"方向鍵模式：{currentArrowKeyMode}");

            int enabledCount = customKeySlots.Count(s => s.Enabled && s.KeyCode != Keys.None);
            if (enabledCount > 0)
            {
                AddLog($"自定義按鍵：{enabledCount} 個已啟用");
            }

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