using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MapleStoryMacro
{
    /// <summary>
    /// 鍵盤阻擋器 - 使用 Low-Level Keyboard Hook 條件式攔截巨集按鍵
    /// 
    /// 工作原理（條件式攔截）：
    /// 1. 檢測帶有 MACRO_KEY_MARKER 標記的按鍵 (dwExtraInfo)
    /// 2. 檢測 LLKHF_INJECTED 旗標（程式注入的按鍵）
    /// 3. 時間戳記追蹤（100ms 內發送的按鍵）
    /// 4. 遊戲前景 → 攔截（return 1），防止重複處理
    /// 5. 遊戲背景 → 放行（CallNextHookEx），確保 key state 更新，
    ///    否則 GetKeyState/GetAsyncKeyState 偵測不到按鍵
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
        private readonly TimeSpan _keyTimeout = TimeSpan.FromMilliseconds(100);  // 放寬超時避免競爭條件

        private IntPtr _hookHandle = IntPtr.Zero;
        private LowLevelKeyboardProc _hookProc;
        private bool _isBlocking = false;
        private bool _disposed = false;
        
        // 目標視窗句柄 - 遊戲視窗
        public IntPtr TargetWindowHandle { get; set; } = IntPtr.Zero;

        // 統計資料
        public int BlockedKeyCount { get; private set; } = 0;
        public int PassedKeyCount { get; private set; } = 0;

        // 調試模式
        public bool DebugMode { get; set; } = false;
        
        // 攔截模式選項
        public bool UseMarkerDetection { get; set; } = true;     // 使用 dwExtraInfo 標記
        public bool UseInjectedDetection { get; set; } = true;   // 使用 LLKHF_INJECTED 旗標
        public bool UseTimestampDetection { get; set; } = true;  // 使用時間戳記追蹤

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
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr hWnd);

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
        /// 鉤子回調函數 - 條件式攔截
        /// 遊戲前景 → 攔截巨集按鍵（防止重複處理）
        /// 遊戲背景 → 放行巨集按鍵（key state 必須更新，否則 GetKeyState 失效）
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
                    
                    // ★ Marker 標記的按鍵 → 條件式處理
                    if (UseMarkerDetection)
                    {
                        ulong extraInfo = hookStruct.dwExtraInfo.ToUInt64();
                        if (extraInfo == MACRO_KEY_MARKER)
                        {
                            bool isFg = IsGameForeground();
                            if (isFg)
                            {
                                // 遊戲前景 → 攔截，防止重複處理
                                BlockedKeyCount++;
                                if (DebugMode)
                                    Debug.WriteLine($"[Blocker] 攔截(Marker): VK=0x{hookStruct.vkCode:X2}, 遊戲前景=true");
                                return (IntPtr)1;
                            }
                            else
                            {
                                // 遊戲背景 → 放行，key state 必須更新
                                PassedKeyCount++;
                                if (DebugMode)
                                    Debug.WriteLine($"[Blocker] 放行(Marker): VK=0x{hookStruct.vkCode:X2}, 遊戲前景=false");
                                return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
                            }
                        }
                    }
                    
                    // 檢測方式 2：LLKHF_INJECTED 旗標（程式注入的按鍵）
                    if (UseInjectedDetection)
                    {
                        bool isInjected = (hookStruct.flags & LLKHF_INJECTED) != 0;
                        bool isLowerIL = (hookStruct.flags & LLKHF_LOWER_IL_INJECTED) != 0;
                        
                        if (isInjected || isLowerIL)
                        {
                            // 只攔截方向鍵和目標按鍵
                            if (IsArrowOrTargetKey(hookStruct.vkCode))
                            {
                                return HandleMacroKey(hookStruct, msg, nCode, wParam, lParam, isLowerIL ? "LowerIL" : "Injected");
                            }
                        }
                    }
                    
                    // 檢測方式 3：時間戳記追蹤（備援）
                    if (UseTimestampDetection && IsPendingKey(hookStruct.vkCode))
                    {
                        return HandleMacroKey(hookStruct, msg, nCode, wParam, lParam, "Timestamp");
                    }
                }
            }

            // 非巨集按鍵或未啟用攔截 → 放行
            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }
        
        /// <summary>
        /// 檢查遊戲是否為前景視窗
        /// </summary>
        private bool IsGameForeground()
        {
            if (TargetWindowHandle == IntPtr.Zero || !IsWindow(TargetWindowHandle))
                return false;
            IntPtr foreground = GetForegroundWindow();
            return foreground == TargetWindowHandle;
        }
        
        /// <summary>
        /// 處理識別為巨集的按鍵 — 條件式攔截
        /// 遊戲前景 → 攔截（防止重複處理）
        /// 遊戲背景 → 放行（key state 必須更新）
        /// </summary>
        private IntPtr HandleMacroKey(KBDLLHOOKSTRUCT hookStruct, int msg, int origNCode, IntPtr origWParam, IntPtr origLParam, string detectionMethod)
        {
            bool isFg = IsGameForeground();
            if (isFg)
            {
                // 遊戲前景 → 攔截，防止重複處理
                BlockedKeyCount++;
                if (DebugMode)
                    Debug.WriteLine($"[Blocker] 攔截({detectionMethod}): VK=0x{hookStruct.vkCode:X2}, 遊戲前景=true");
                return (IntPtr)1;
            }
            else
            {
                // 遊戲背景 → 放行，key state 必須更新
                PassedKeyCount++;
                if (DebugMode)
                    Debug.WriteLine($"[Blocker] 放行({detectionMethod}): VK=0x{hookStruct.vkCode:X2}, 遊戲前景=false");
                return CallNextHookEx(_hookHandle, origNCode, origWParam, origLParam);
            }
        }
        
        /// <summary>
        /// 檢查是否為方向鍵或目標按鍵（用於 Injected 檢測）
        /// 注意：組合鍵時所有注入的按鍵都應該被攔截
        /// </summary>
        private bool IsArrowOrTargetKey(uint vkCode)
        {
            // 方向鍵（最常見的問題按鍵）
            if (vkCode == 0x25 || vkCode == 0x26 || vkCode == 0x27 || vkCode == 0x28) // Left, Up, Right, Down
                return true;
            
            // 其他延伸鍵
            if (vkCode == 0x2D || vkCode == 0x2E) // Insert, Delete
                return true;
            if (vkCode == 0x24 || vkCode == 0x23) // Home, End
                return true;
            if (vkCode == 0x21 || vkCode == 0x22) // PageUp, PageDown
                return true;
            
            // 一般字母鍵（組合鍵中常用，如 X+Down, Y+Right）
            if (vkCode >= 0x41 && vkCode <= 0x5A) // A-Z
                return true;
            
            // 數字鍵
            if (vkCode >= 0x30 && vkCode <= 0x39) // 0-9
                return true;
            
            // 功能鍵
            if (vkCode >= 0x70 && vkCode <= 0x7B) // F1-F12
                return true;
            
            // 空白鍵、Enter、Escape
            if (vkCode == 0x20 || vkCode == 0x0D || vkCode == 0x1B) // Space, Enter, Escape
                return true;
            
            // Shift, Ctrl, Alt（修飾鍵）
            if (vkCode == 0x10 || vkCode == 0x11 || vkCode == 0x12) // Shift, Ctrl, Alt
                return true;
            if (vkCode == 0xA0 || vkCode == 0xA1) // LShift, RShift
                return true;
            if (vkCode == 0xA2 || vkCode == 0xA3) // LCtrl, RCtrl
                return true;
            if (vkCode == 0xA4 || vkCode == 0xA5) // LAlt, RAlt
                return true;
                
            return false;
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
