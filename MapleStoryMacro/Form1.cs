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

        // Log System
        private List<string> logMessages = new List<string>();
    private readonly int MAX_LOG_LINES = 15;
    private readonly object logLock = new object();

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
       FormClosing += Form1_FormClosing;

            // Enable KeyPreview to capture all keys
this.KeyPreview = true;
            this.KeyDown += Form1_KeyDown;
            this.KeyUp += Form1_KeyUp;

            // Initialize keyboard hook
            keyboardHook = new KeyboardHookDLL();
    keyboardHook.OnKeyEvent += KeyboardHook_OnKeyEvent;

            // Initialize log system
  monitorTimer.Interval = 100;
            monitorTimer.Tick += MonitorTimer_Tick;
  monitorTimer.Start();

            AddLog("Application started");
  AddLog("Ready for recording...");

            UpdateUI();
            UpdateWindowStatus();
        }

        private void AddLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            string logEntry = $"[{timestamp}] {message}" ;

        lock (logLock)
       {
         logMessages.Add(logEntry);

  if (logMessages.Count > MAX_LOG_LINES)
        logMessages.RemoveAt(0);
   }

          System.Diagnostics.Debug.WriteLine(logEntry);
        }

        private void MonitorTimer_Tick(object sender, EventArgs e)
        {
  // Draw logs on picPreview
            if (picPreview.Image != null)
            picPreview.Image.Dispose();

            Bitmap logBitmap = new Bitmap(picPreview.Width, picPreview.Height);
      using (Graphics g = Graphics.FromImage(logBitmap))
            {
        g.Clear(Color.Black);

                // Draw background
   using (Brush brushBg = new SolidBrush(Color.FromArgb(20, 20, 20)))
          {
       g.FillRectangle(brushBg, 0, 0, logBitmap.Width, logBitmap.Height);
 }

       // Draw log text
         using (Font font = new Font("Consolas", 9, FontStyle.Regular))
                using (Brush brushText = new SolidBrush(Color.LimeGreen))
    {
        int yPos = 10;
           List<string> logsCopy;
  lock (logLock)
   {
   logsCopy = new List<string>(logMessages);
     }
           foreach (string log in logsCopy)
           {
   g.DrawString(log, font, brushText, 10, yPos);
              yPos += 12;
        }
           }

          // Draw status bar
  using (Font fontBold = new Font("Consolas", 10, FontStyle.Bold))
          using (Brush brushStatus = new SolidBrush(Color.Yellow))
            {
        string bgMode = (targetWindowHandle != IntPtr.Zero) ? "BG" : "FG";
  string statusLine = $"Rec: {(isRecording ? "ON" : "OFF")} | Play: {(isPlaying ? "ON" : "OFF")} | Mode: {bgMode} | Events: {recordedEvents.Count}";

 g.DrawString(statusLine, fontBold, brushStatus, 10, logBitmap.Height - 30);
        }
            }

picPreview.Image = logBitmap;
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
   Text = "Select Window",
            Width = 400,
        Height = 300,
         StartPosition = FormStartPosition.CenterParent,
  Owner = this
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

     windowSelector.Controls.Add(listBox);
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
          MessageBox.Show($"Selected: {selected.Title}\n\nBackground mode is now enabled!");
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
       MessageBox.Show("No events to save");
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
    MessageBox.Show($"Saved to: {sfd.FileName}");
                }
 catch (Exception ex)
    {
       AddLog($"Save failed: {ex.Message}");
    MessageBox.Show($"Save failed: {ex.Message}");
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
MessageBox.Show($"已載入 {recordedEvents.Count} 個事件", "載入成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
       
       // 更新 UI 狀態，啟用開始播放按鍵
            UpdateUI();
  }
     catch (Exception ex)
         {
         AddLog($"載入失敗: {ex.Message}");
      MessageBox.Show($"載入失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
    MessageBox.Show("變更已儲存");
        editorForm.Close();
     }
 catch (Exception ex)
          {
      AddLog($"儲存失敗: {ex.Message}");
       MessageBox.Show($"儲存失敗: {ex.Message}");
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
       
            string mode = (targetWindowHandle != IntPtr.Zero) ? "Background" : "Foreground";
         AddLog($"Playback started ({mode} mode)...");

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
          isPlaying = false;
 }

     UpdateUI();
        }

     private void PlaybackThread(int loopCount)
        {
      try
     {
       for (int loop = 1; loop <= loopCount && isPlaying; loop++)
         {
          this.Invoke(new Action(() =>
             {
  lblPlaybackStatus.Text = $"Loop: {loop}/{loopCount}";
      lblPlaybackStatus.ForeColor = Color.Blue;
         }));

           double lastTimestamp = 0;
        foreach (MacroEvent evt in recordedEvents)
    {
         if (!isPlaying) break;

 double waitTime = evt.Timestamp - lastTimestamp;
         if (waitTime > 0)
          Thread.Sleep((int)(waitTime * 1000));

         SendKeyEvent(evt);
            lastTimestamp = evt.Timestamp;
               }

            Thread.Sleep(200);
      }

  isPlaying = false;
        this.Invoke(new Action(() =>
{
           lblPlaybackStatus.Text = "Playback: Completed";
            lblPlaybackStatus.ForeColor = Color.Green;
           AddLog("Playback completed");
         UpdateUI();
           }));
            }
    catch (Exception ex)
          {
      this.Invoke(new Action(() =>
    {
          AddLog($"Playback error: {ex.Message}");
          MessageBox.Show($"Playback error: {ex.Message}");
    }));
    isPlaying = false;
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
          // 對於方向鍵，使用 AttachThreadInput 方式
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
      }
      catch (Exception ex)
  {
 AddLog($"Send failed: {ex.Message}");
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
        private void SendKeyWithThreadAttach(IntPtr hWnd, Keys key, bool isKeyDown)
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

         keybd_event(vkCode, scanCode, flags, UIntPtr.Zero);
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
     isPlaying = false;
            lblPlaybackStatus.Text = "Playback: Stopped";
            lblPlaybackStatus.ForeColor = Color.Orange;
            AddLog("播放已停止");
UpdateUI();
        }

      private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
         if (isRecording)
        BtnStopRecording_Click(null, null);
            if (isPlaying)
          isPlaying = false;

   keyboardHook?.Uninstall();
            monitorTimer?.Stop();
            AddLog("應用程式已關閉");
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

 private void picPreview_Click(object sender, EventArgs e)
  {
      // picPreview 顯示即時日誌
        }

        [Serializable]
        public class MacroEvent
{
            public Keys KeyCode { get; set; }
            public string EventType { get; set; }
    public double Timestamp { get; set; }
        }
 }
}
