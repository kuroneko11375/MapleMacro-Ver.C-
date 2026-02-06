using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MapleStoryMacro
{
    /// <summary>
    /// 鍵盤阻擋器 - 使用 Low-Level Keyboard Hook 過濾巨集發送的按鍵
    /// 
    /// 工作原理（多重檢測 + 智能攔截）：
    /// 1. 檢測帶有 MACRO_KEY_MARKER 標記的按鍵 (dwExtraInfo)
    /// 2. 檢測 LLKHF_INJECTED 旗標（程式注入的按鍵）
    /// 3. 時間戳記追蹤（50ms 內發送的按鍵）
    /// 4. 如果前景視窗是遊戲 → 放行
    /// 5. 如果前景視窗不是遊戲 → 攔截 + PostMessage 發送給遊戲
    /// </summary>
    public class KeyboardBlocker : IDisposable
    {
        // 巨集按鍵標記 - 用於識別由巨集發送的按鍵
        public const uint MACRO_KEY_MARKER = 0x12345678;
        
        // LLKHF 旗標 - 用於檢測注入的按鍵
        private const uint LLKHF_INJECTED = 0x10;      // 程式注入的按鍵
        private const uint LLKHF_LOWER_IL_INJECTED = 0x02; // 較低完整性等級注入
        
        // 時間戳記追蹤（用於輔助識別）
        private readonly ConcurrentDictionary<uint, DateTime> _pendingKeys = new();
        private readonly TimeSpan _keyTimeout = TimeSpan.FromMilliseconds(100);

        private IntPtr _hookHandle = IntPtr.Zero;
        private LowLevelKeyboardProc _hookProc;
        private bool _isBlocking = false;
        private bool _disposed = false;

        // 目標視窗句柄 - 遊戲視窗
        public IntPtr TargetWindowHandle { get; set; } = IntPtr.Zero;

        // 統計資料
        public int BlockedKeyCount { get; private set; } = 0;
        public int PassedKeyCount { get; private set; } = 0;
        public int PostMessageSentCount { get; private set; } = 0;

        // 調試模式
        public bool DebugMode { get; set; } = false;
        
        // 攔截模式選項
        public bool UseMarkerDetection { get; set; } = true;     // 使用 dwExtraInfo 標記
        public bool UseInjectedDetection { get; set; } = true;   // 使用 LLKHF_INJECTED 旗標
        public bool UseTimestampDetection { get; set; } = true;  // 使用時間戳記追蹤
        public bool SendPostMessageOnBlock { get; set; } = true; // 攔截時發送 PostMessage

        #region Windows API

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);


        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr hWnd);

        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        #endregion

        public KeyboardBlocker()
        {
            // 保存委派參考，避免被 GC 回收
            _hookProc = HookCallback;
        }

        /// <summary>
        /// 安裝鍵盤攔截鉤子
        /// </summary>
        public bool Install()
        {
            if (_hookHandle != IntPtr.Zero)
            {
                Debug.WriteLine("[KeyboardBlocker] 鉤子已安裝");
                return true;
            }

            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule!)
            {
                _hookHandle = SetWindowsHookEx(
                    WH_KEYBOARD_LL,
                    _hookProc,
                    GetModuleHandle(curModule.ModuleName),
                    0);
            }

            if (_hookHandle == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                Debug.WriteLine($"[KeyboardBlocker] 安裝失敗，錯誤碼: {error}");
                return false;
            }

            Debug.WriteLine($"[KeyboardBlocker] 鉤子已安裝，Handle: {_hookHandle}");
            return true;
        }

        /// <summary>
        /// 卸載鍵盤攔截鉤子
        /// </summary>
        public void Uninstall()
        {
            if (_hookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookHandle);
                _hookHandle = IntPtr.Zero;
                Debug.WriteLine("[KeyboardBlocker] 鉤子已卸載");
            }
        }

        /// <summary>
        /// 啟用/停用按鍵攔截
        /// </summary>
        public bool IsBlocking
        {
            get => _isBlocking;
            set
            {
                _isBlocking = value;
                Debug.WriteLine($"[KeyboardBlocker] 攔截狀態: {(_isBlocking ? "啟用" : "停用")}");
            }
        }

        /// <summary>
        /// 重置統計資料
        /// </summary>
        public void ResetStats()
        {
            BlockedKeyCount = 0;
            PassedKeyCount = 0;
            PostMessageSentCount = 0;
            _pendingKeys.Clear();
        }
        
        /// <summary>
        /// 註冊即將發送的按鍵（用於時間戳記追蹤）
        /// </summary>
        public void RegisterPendingKey(uint vkCode)
        {
            _pendingKeys[vkCode] = DateTime.Now;
            
            // 清理過期的記錄
            var expiredKeys = new System.Collections.Generic.List<uint>();
            foreach (var kvp in _pendingKeys)
            {
                if (DateTime.Now - kvp.Value > _keyTimeout)
                {
                    expiredKeys.Add(kvp.Key);
                }
            }
            foreach (var key in expiredKeys)
            {
                _pendingKeys.TryRemove(key, out _);
            }
        }
        
        /// <summary>
        /// 檢查按鍵是否在待處理列表中（時間戳記追蹤）
        /// </summary>
        private bool IsPendingKey(uint vkCode)
        {
            if (_pendingKeys.TryGetValue(vkCode, out DateTime timestamp))
            {
                if (DateTime.Now - timestamp <= _keyTimeout)
                {
                    _pendingKeys.TryRemove(vkCode, out _);
                    return true;
                }
                _pendingKeys.TryRemove(vkCode, out _);
            }
            return false;
        }

        /// <summary>
        /// 鉤子回調函數 - 多重檢測攔截模式
        /// 使用多種方式識別巨集按鍵：dwExtraInfo 標記、LLKHF_INJECTED 旗標、時間戳記追蹤
        /// </summary>
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && _isBlocking)
            {
                int msg = (int)wParam;
                
                // 只處理按鍵訊息
                if (msg == WM_KEYDOWN || msg == WM_KEYUP || 
                    msg == WM_SYSKEYDOWN || msg == WM_SYSKEYUP)
                {
                    KBDLLHOOKSTRUCT hookStruct = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                    
                    // 多重檢測：判斷是否為巨集發送的按鍵
                    bool isMacroKey = false;
                    string detectionMethod = "";
                    
                    // 檢測方式 1：dwExtraInfo 標記
                    if (UseMarkerDetection)
                    {
                        ulong extraInfo = hookStruct.dwExtraInfo.ToUInt64();
                        if (extraInfo == MACRO_KEY_MARKER)
                        {
                            isMacroKey = true;
                            detectionMethod = "Marker";
                        }
                    }
                    
                    // 檢測方式 2：LLKHF_INJECTED 旗標（程式注入的按鍵）
                    if (!isMacroKey && UseInjectedDetection)
                    {
                        bool isInjected = (hookStruct.flags & LLKHF_INJECTED) != 0;
                        bool isLowerIL = (hookStruct.flags & LLKHF_LOWER_IL_INJECTED) != 0;
                        
                        if (isInjected || isLowerIL)
                        {
                            // 額外檢查：只攔截方向鍵和我們關注的按鍵
                            if (IsArrowOrTargetKey(hookStruct.vkCode))
                            {
                                isMacroKey = true;
                                detectionMethod = isLowerIL ? "LowerIL" : "Injected";
                            }
                        }
                    }
                    
                    // 檢測方式 3：時間戳記追蹤
                    if (!isMacroKey && UseTimestampDetection)
                    {
                        if (IsPendingKey(hookStruct.vkCode))
                        {
                            isMacroKey = true;
                            detectionMethod = "Timestamp";
                        }
                    }
                    
                    // 如果識別為巨集按鍵
                    if (isMacroKey)
                    {
                        IntPtr foregroundWindow = GetForegroundWindow();
                        bool isKeyDown = (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN);
                        
                        // 檢查目標視窗是否有效
                        if (TargetWindowHandle != IntPtr.Zero && IsWindow(TargetWindowHandle))
                        {
                            if (foregroundWindow == TargetWindowHandle)
                            {
                                // 前景視窗是遊戲 → 放行
                                PassedKeyCount++;
                                if (DebugMode)
                                {
                                    Debug.WriteLine($"[Blocker] ✓ 放行: VK=0x{hookStruct.vkCode:X2} ({detectionMethod}, 前景=目標)");
                                }
                                return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                            }
                            else
                            {
                                // 前景視窗不是遊戲 → 攔截 + 額外發送 PostMessage 確保遊戲收到
                                BlockedKeyCount++;
                                
                                if (SendPostMessageOnBlock)
                                {
                                    // 構建 lParam 並發送 PostMessage 給遊戲
                                    SendKeyMessageToTarget(hookStruct.vkCode, hookStruct.scanCode, isKeyDown);
                                    PostMessageSentCount++;
                                }
                                
                                if (DebugMode)
                                {
                                    Debug.WriteLine($"[Blocker] ★ 攔截: VK=0x{hookStruct.vkCode:X2} ({detectionMethod}, PostMsg={SendPostMessageOnBlock})");
                                }
                                return (IntPtr)1;
                            }
                        }
                        else
                        {
                            // 沒有設定目標視窗 → 預設攔截
                            BlockedKeyCount++;
                            if (DebugMode)
                            {
                                Debug.WriteLine($"[Blocker] ★ 攔截: VK=0x{hookStruct.vkCode:X2} ({detectionMethod}, 無目標)");
                            }
                            return (IntPtr)1;
                        }
                    }
                    else
                    {
                        // 非巨集按鍵 → 放行
                        PassedKeyCount++;
                    }
                }
            }

            // 繼續傳遞給下一個鉤子
            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }
        
        /// <summary>
        /// 檢查是否為方向鍵或目標按鍵
        /// </summary>
        private bool IsArrowOrTargetKey(uint vkCode)
        {
            // 方向鍵
            if (vkCode == 0x25 || vkCode == 0x26 || vkCode == 0x27 || vkCode == 0x28) // Left, Up, Right, Down
                return true;
            
            // 其他延伸鍵
            if (vkCode == 0x2D || vkCode == 0x2E) // Insert, Delete
                return true;
            if (vkCode == 0x24 || vkCode == 0x23) // Home, End
                return true;
            if (vkCode == 0x21 || vkCode == 0x22) // PageUp, PageDown
                return true;
                
            return false;
        }
        
        /// <summary>
        /// 發送按鍵到目標視窗（使用 AttachThreadInput + keybd_event，方向鍵必須用這種方式）
        /// </summary>
        private void SendKeyMessageToTarget(uint vkCode, uint scanCode, bool isKeyDown)
        {
            if (TargetWindowHandle == IntPtr.Zero || !IsWindow(TargetWindowHandle))
                return;
            
            uint targetThreadId = GetWindowThreadProcessId(TargetWindowHandle, out uint processId);
            uint currentThreadId = GetCurrentThreadId();
            
            bool attached = false;
            try
            {
                // 附加線程以便能設定焦點
                if (targetThreadId != currentThreadId)
                {
                    attached = AttachThreadInput(currentThreadId, targetThreadId, true);
                }
                
                if (attached)
                {
                    // 設定焦點到遊戲視窗（在線程附加後有效）
                    SetFocus(TargetWindowHandle);
                }
                
                // 使用 keybd_event 發送按鍵（不帶 Marker，避免被再次攔截）
                byte vk = (byte)vkCode;
                byte sc = (byte)scanCode;
                uint flags = 0;
                
                if (!isKeyDown)
                {
                    flags |= KEYEVENTF_KEYUP;
                }
                if (IsArrowOrTargetKey(vkCode))
                {
                    flags |= KEYEVENTF_EXTENDEDKEY;
                }
                
                // 不帶 Marker，直接發送
                keybd_event(vk, sc, flags, UIntPtr.Zero);
            }
            finally
            {
                // 解除線程附加
                if (attached && targetThreadId != currentThreadId)
                {
                    AttachThreadInput(currentThreadId, targetThreadId, false);
                }
            }
        }

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                Uninstall();
                _disposed = true;
            }
        }

        ~KeyboardBlocker()
        {
            Dispose(false);
        }

        #endregion
    }
}
