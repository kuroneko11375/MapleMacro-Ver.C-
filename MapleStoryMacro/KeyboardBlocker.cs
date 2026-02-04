using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 鍵盤攔截器 - Filter 模式
    /// 工作原理：
    /// 1. 安裝 Low-Level Keyboard Hook
    /// 2. 當我們發送按鍵時，使用特殊標記
    /// 3. Hook 攔截到帶標記的按鍵時：
    ///    - 如果前景視窗是目標視窗：放行
    ///    - 如果前景視窗不是目標視窗：攔截（保護其他視窗）
    /// </summary>
    public class KeyboardBlocker
    {
        // ===== Windows API =====
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        // ===== 常數 =====
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const uint MAPVK_VK_TO_VSC = 0;

        // 使用特殊的 ExtraInfo 來標記我們自己發送的按鍵
        private static readonly UIntPtr MACRO_KEY_MARKER = new UIntPtr(0x12345678);

        public static UIntPtr MacroKeyMarker => MACRO_KEY_MARKER;

        // ===== 委託 =====
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        // ===== 結構體 =====
        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        // ===== 成員變數 =====
        private IntPtr hookHandle = IntPtr.Zero;
        private LowLevelKeyboardProc hookProc;
        private bool isInstalled = false;
        private readonly object lockObj = new object();

        // 目標視窗（只有這個視窗可以收到我們發送的按鍵）
        private IntPtr targetWindow = IntPtr.Zero;

        // 統計
        private int blockedCount = 0;
        private int passedCount = 0;

        public KeyboardBlocker()
        {
            hookProc = HookCallback;
        }

        /// <summary>
        /// 設定目標視窗（Filter：只有這個視窗可以收到按鍵）
        /// </summary>
        public void SetTargetWindow(IntPtr hWnd)
        {
            targetWindow = hWnd;
            Debug.WriteLine($"?? 目標視窗已設定: {hWnd}");
        }

        /// <summary>
        /// 安裝攔截器
        /// </summary>
        public bool Install()
        {
            if (isInstalled) return true;

            try
            {
                using (Process curProcess = Process.GetCurrentProcess())
                using (ProcessModule? curModule = curProcess.MainModule)
                {
                    if (curModule == null)
                    {
                        Debug.WriteLine("? 無法取得模組，攔截器安裝失敗");
                        return false;
                    }

                    hookHandle = SetWindowsHookEx(
                        WH_KEYBOARD_LL,
                        hookProc,
                        GetModuleHandle(curModule.ModuleName),
                        0
                    );
                }

                if (hookHandle != IntPtr.Zero)
                {
                    isInstalled = true;
                    Debug.WriteLine("? 鍵盤攔截器已安裝（Filter 模式）");
                    return true;
                }
                else
                {
                    Debug.WriteLine("? 鍵盤攔截器安裝失敗");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"? 鍵盤攔截器異常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 卸載攔截器
        /// </summary>
        public void Uninstall()
        {
            if (!isInstalled) return;

            try
            {
                if (hookHandle != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(hookHandle);
                    hookHandle = IntPtr.Zero;
                }
                isInstalled = false;
                Debug.WriteLine($"? 鍵盤攔截器已卸載 (攔截: {blockedCount}, 放行: {passedCount})");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"? 卸載攔截器異常: {ex.Message}");
            }
        }

        /// <summary>
        /// 發送按鍵（帶 Filter 標記）
        /// </summary>
        public void SendKeyBlocked(Keys key, bool isKeyDown)
        {
            byte vkCode = (byte)key;
            byte scanCode = GetScanCode(key);
            bool isExtended = IsExtendedKey(key);

            uint flags = 0;
            if (!isKeyDown) flags |= KEYEVENTF_KEYUP;
            if (isExtended) flags |= KEYEVENTF_EXTENDEDKEY;

            // 使用特殊的 dwExtraInfo 標記這是我們發送的
            keybd_event(vkCode, scanCode, flags, MACRO_KEY_MARKER);

            Debug.WriteLine($"?? 發送: {key} ({(isKeyDown ? "down" : "up")})");
        }

        /// <summary>
        /// Hook 回調函數 - Filter 模式
        /// </summary>
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();
                if (msg == WM_KEYDOWN || msg == WM_KEYUP || msg == WM_SYSKEYDOWN || msg == WM_SYSKEYUP)
                {
                    KBDLLHOOKSTRUCT hookStruct = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                    Keys key = (Keys)hookStruct.vkCode;

                    // 只處理我們發送的按鍵（帶有特殊標記）
                    if (hookStruct.dwExtraInfo == MACRO_KEY_MARKER)
                    {
                        // 取得當前前景視窗
                        IntPtr foregroundWindow = GetForegroundWindow();

                        // Filter 邏輯：
                        // 如果前景視窗是目標視窗 → 放行
                        // 如果前景視窗不是目標視窗 → 攔截（保護其他視窗）
                        if (targetWindow != IntPtr.Zero && IsWindow(targetWindow))
                        {
                            if (foregroundWindow == targetWindow)
                            {
                                // 前景是目標視窗，放行
                                passedCount++;
                                Debug.WriteLine($"? 放行: {key} (前景=目標)");
                                return CallNextHookEx(hookHandle, nCode, wParam, lParam);
                            }
                            else
                            {
                                // 前景不是目標視窗，攔截保護
                                blockedCount++;
                                Debug.WriteLine($"??? 攔截: {key} (前景≠目標，保護其他視窗)");
                                return (IntPtr)1;
                            }
                        }
                        else
                        {
                            // 沒有設定目標視窗，預設攔截
                            blockedCount++;
                            Debug.WriteLine($"??? 攔截: {key} (無目標視窗)");
                            return (IntPtr)1;
                        }
                    }
                }
            }

            return CallNextHookEx(hookHandle, nCode, wParam, lParam);
        }

        /// <summary>
        /// 取得掃描碼
        /// </summary>
        private byte GetScanCode(Keys key)
        {
            return key switch
            {
                Keys.Left => 0x4B,
                Keys.Right => 0x4D,
                Keys.Up => 0x48,
                Keys.Down => 0x50,
                Keys.Insert => 0x52,
                Keys.Delete => 0x53,
                Keys.Home => 0x47,
                Keys.End => 0x4F,
                Keys.PageUp => 0x49,
                Keys.PageDown => 0x51,
                _ => (byte)MapVirtualKey((uint)key, MAPVK_VK_TO_VSC)
            };
        }

        /// <summary>
        /// 檢查是否為延伸鍵
        /// </summary>
        private bool IsExtendedKey(Keys key)
        {
            return key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down ||
                   key == Keys.Insert || key == Keys.Delete || key == Keys.Home || key == Keys.End ||
                   key == Keys.PageUp || key == Keys.PageDown;
        }

        /// <summary>
        /// 是否已安裝
        /// </summary>
        public bool IsInstalled => isInstalled;

        /// <summary>
        /// 取得統計資訊
        /// </summary>
        public string GetStats() => $"攔截: {blockedCount}, 放行: {passedCount}";
    }
}
