using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 低階鍵盤鉤子 - 用於全局監聽鍵盤事件
    /// </summary>
    public class KeyboardHookDLL
    {
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

        private IntPtr _hookHandle = IntPtr.Zero;
        private LowLevelKeyboardProc _hookProc;

        /// <summary>
        /// 鍵盤事件回調 - 參數為 (Keys keyCode, bool isKeyDown)
        /// </summary>
        public event Action<Keys, bool>? OnKeyEvent;

        public KeyboardHookDLL()
        {
            // 保存委派參考，避免被 GC 回收
            _hookProc = HookCallback;
        }

        /// <summary>
        /// 安裝鍵盤鉤子
        /// </summary>
        /// <returns>安裝是否成功</returns>
        public bool Install()
        {
            if (_hookHandle != IntPtr.Zero)
            {
                Debug.WriteLine("[KeyboardHookDLL] 鉤子已安裝");
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
                Debug.WriteLine($"[KeyboardHookDLL] 安裝失敗，錯誤碼: {error}");
                return false;
            }

            Debug.WriteLine($"[KeyboardHookDLL] 鉤子已安裝，Handle: {_hookHandle}");
            return true;
        }

        /// <summary>
        /// 卸載鍵盤鉤子
        /// </summary>
        public void Uninstall()
        {
            if (_hookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookHandle);
                _hookHandle = IntPtr.Zero;
                Debug.WriteLine("[KeyboardHookDLL] 鉤子已卸載");
            }
        }

        /// <summary>
        /// 鉤子回調函數
        /// </summary>
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int msg = (int)wParam;

                // 只處理按鍵訊息
                if (msg == WM_KEYDOWN || msg == WM_KEYUP ||
                    msg == WM_SYSKEYDOWN || msg == WM_SYSKEYUP)
                {
                    KBDLLHOOKSTRUCT hookStruct = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                    Keys keyCode = (Keys)hookStruct.vkCode;
                    bool isKeyDown = (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN);

                    // 觸發事件
                    OnKeyEvent?.Invoke(keyCode, isKeyDown);
                }
            }

            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }
    }
}
