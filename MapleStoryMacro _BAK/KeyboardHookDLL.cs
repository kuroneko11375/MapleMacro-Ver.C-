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
        private const int WH_GETMESSAGE = 3;
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

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public POINT pt;
        }

        #endregion

        private IntPtr _hookHandle = IntPtr.Zero;
        private readonly LowLevelKeyboardProc _hookProc;
        private readonly KeyboardHookMode _mode;
        private static IntPtr _cachedModuleHandle = IntPtr.Zero; // 快取模組控制代碼，避免重複取得

        /// <summary>
        /// 鉤子是否已安裝
        /// </summary>
        public bool IsInstalled => _hookHandle != IntPtr.Zero;

        /// <summary>
        /// 鍵盤事件回調 - 參數為 (Keys keyCode, bool isKeyDown)
        /// </summary>
        public enum KeyboardHookMode
        {
            LowLevel,
            GetMessage
        }

        public event Action<Keys, bool>? OnKeyEvent;

        public KeyboardHookDLL(KeyboardHookMode mode = KeyboardHookMode.LowLevel)
        {
            // 保存委派參考，避免被 GC 回收
            _mode = mode;
            _hookProc = HookCallback;

            // 預先快取模組控制代碼（避免 Install 時阻塞 UI）
            if (_cachedModuleHandle == IntPtr.Zero)
            {
                using (Process curProcess = Process.GetCurrentProcess())
                using (ProcessModule curModule = curProcess.MainModule!)
                {
                    _cachedModuleHandle = GetModuleHandle(curModule.ModuleName);
                }
            }
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

            int hookId = _mode == KeyboardHookMode.GetMessage ? WH_GETMESSAGE : WH_KEYBOARD_LL;
            uint threadId = _mode == KeyboardHookMode.GetMessage ? GetCurrentThreadId() : 0;

            _hookHandle = SetWindowsHookEx(
                hookId,
                _hookProc,
                _cachedModuleHandle,
                threadId);

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
                if (_mode == KeyboardHookMode.GetMessage)
                {
                    MSG msg = Marshal.PtrToStructure<MSG>(lParam);
                    if (msg.message == WM_KEYDOWN || msg.message == WM_KEYUP ||
                        msg.message == WM_SYSKEYDOWN || msg.message == WM_SYSKEYUP)
                    {
                        Keys keyCode = (Keys)msg.wParam;
                        bool isKeyDown = (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN);
                        OnKeyEvent?.Invoke(keyCode, isKeyDown);
                    }
                }
                else
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
            }

            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }
    }
}
