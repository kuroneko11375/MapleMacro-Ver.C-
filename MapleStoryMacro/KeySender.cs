using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 方向鍵發送模式
    /// </summary>
    public enum ArrowKeyMode
    {
        SendToChild,             // ThreadAttach + PostMessage（背景走路用）
        ThreadAttachWithBlocker, // ThreadAttach + Blocker（嘗試避免影響前景）
        SendInputWithBlock,      // SendInput + Blocker（嘗試避免影響前景）
    }

    /// <summary>
    /// 按鍵發送引擎 — 封裝所有 Win32 按鍵/滑鼠發送邏輯
    /// </summary>
    public class KeySender
    {
        #region P/Invoke 宣告

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
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string? lpszClass, string? lpszWindow);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        #endregion

        #region 結構定義

        [StructLayout(LayoutKind.Sequential)]
        internal struct INPUT
        {
            public uint type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        internal struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct POINT
        {
            public int X;
            public int Y;
        }

        #endregion

        #region 常數

        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP = 0x0101;
        private const uint WM_CHAR = 0x0102;
        private const uint WM_SYSKEYDOWN = 0x0104;
        private const uint WM_SYSKEYUP = 0x0105;
        private const uint MAPVK_VK_TO_VSC = 0;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const uint INPUT_KEYBOARD = 1;
        private const uint INPUT_MOUSE = 0;
        private const uint WM_LBUTTONDOWN = 0x0201;
        private const uint WM_LBUTTONUP = 0x0202;
        private const uint MK_LBUTTON = 0x0001;
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint WM_MOUSEMOVE = 0x0200;

        #endregion

        #region 狀態屬性

        /// <summary>
        /// 目標視窗控制代碼（背景模式）
        /// </summary>
        public IntPtr TargetWindowHandle { get; set; }

        /// <summary>
        /// 方向鍵發送模式
        /// </summary>
        public ArrowKeyMode CurrentArrowKeyMode { get; set; } = ArrowKeyMode.SendToChild;

        /// <summary>
        /// 鍵盤阻擋器（用於 Blocker 模式）
        /// </summary>
        public KeyboardBlocker? KeyboardBlocker { get; set; }

        /// <summary>
        /// 目前按住的按鍵集合
        /// </summary>
        public HashSet<Keys> PressedKeys { get; } = new HashSet<Keys>();

        /// <summary>
        /// 日誌回呼（呼叫端負責線程安全，Form1.AddLog 已自帶 InvokeRequired 檢查）
        /// </summary>
        public Action<string>? Log { get; set; }

        #endregion

        #region 靜態輔助方法

        /// <summary>
        /// 檢查按鍵是否為修飾鍵（Ctrl/Alt/Shift）
        /// </summary>
        public static bool IsModifierKey(Keys key)
        {
            return key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey ||
                   key == Keys.ShiftKey || key == Keys.LShiftKey || key == Keys.RShiftKey ||
                   key == Keys.Menu || key == Keys.LMenu || key == Keys.RMenu;
        }

        /// <summary>
        /// 將具體的修飾鍵碼轉換為 Keys 修飾旗標
        /// </summary>
        public static Keys ModifierKeyToFlag(Keys key)
        {
            return key switch
            {
                Keys.ControlKey or Keys.LControlKey or Keys.RControlKey => Keys.Control,
                Keys.ShiftKey or Keys.LShiftKey or Keys.RShiftKey => Keys.Shift,
                Keys.Menu or Keys.LMenu or Keys.RMenu => Keys.Alt,
                _ => Keys.None
            };
        }

        #endregion

        #region 按鍵類型判斷

        /// <summary>
        /// 檢查是否為方向鍵
        /// </summary>
        public bool IsArrowKey(Keys key)
        {
            return key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down;
        }

        /// <summary>
        /// 檢查是否為英數鍵
        /// </summary>
        public bool IsAlphaNumericKey(Keys key)
        {
            return (key >= Keys.A && key <= Keys.Z) || (key >= Keys.D0 && key <= Keys.D9);
        }

        /// <summary>
        /// 檢查是否為 Alt 鍵
        /// </summary>
        public bool IsAltKey(Keys key)
        {
            return key == Keys.Alt || key == Keys.Menu || key == Keys.LMenu || key == Keys.RMenu;
        }

        /// <summary>
        /// 檢查是否為導航鍵（Delete, Insert, Home, End, PageUp, PageDown）
        /// </summary>
        public bool IsNavigationKey(Keys key)
        {
            return key == Keys.Delete || key == Keys.Insert ||
                   key == Keys.Home || key == Keys.End ||
                   key == Keys.PageUp || key == Keys.PageDown;
        }

        /// <summary>
        /// 檢查是否為功能鍵 (F1-F12)
        /// </summary>
        public bool IsFunctionKey(Keys key)
        {
            return key >= Keys.F1 && key <= Keys.F12;
        }

        /// <summary>
        /// 檢查是否為延伸鍵（方向鍵、Insert、Delete 等）
        /// </summary>
        public bool IsExtendedKey(Keys key)
        {
            return key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down ||
                   key == Keys.Insert || key == Keys.Delete || key == Keys.Home || key == Keys.End ||
                   key == Keys.PageUp || key == Keys.PageDown || key == Keys.NumLock || key == Keys.PrintScreen ||
                   key == Keys.RMenu || key == Keys.RControlKey || key == Keys.RShiftKey;
        }

        #endregion

        #region 掃描碼

        /// <summary>
        /// 取得按鍵的掃描碼
        /// </summary>
        public byte GetScanCode(Keys key)
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
                Keys.Menu => 0x38,
                Keys.LMenu => 0x38,
                Keys.RMenu => 0x38,
                _ => (byte)MapVirtualKey((uint)key, MAPVK_VK_TO_VSC)
            };
        }

        #endregion

        #region 核心按鍵發送方法

        /// <summary>
        /// 發送按鍵事件（主入口 — 根據按鍵類型自動路由）
        /// </summary>
        public void SendKeyEvent(MacroEvent evt)
        {
            try
            {
                bool isDown = evt.EventType == "down";

                if (TargetWindowHandle != IntPtr.Zero && IsWindow(TargetWindowHandle))
                {
                    if (IsAltKey(evt.KeyCode))
                    {
                        SendAltKeyToWindow(TargetWindowHandle, evt.KeyCode, isDown);
                    }
                    else if (IsArrowKey(evt.KeyCode))
                    {
                        SendArrowKeyWithMode(TargetWindowHandle, evt.KeyCode, isDown);
                    }
                    else if (IsAlphaNumericKey(evt.KeyCode))
                    {
                        SendKeyWithPostMessageOnly(TargetWindowHandle, evt.KeyCode, isDown);
                    }
                    else if (IsNavigationKey(evt.KeyCode))
                    {
                        SendKeyWithPostMessageOnly(TargetWindowHandle, evt.KeyCode, isDown);
                    }
                    else if (IsFunctionKey(evt.KeyCode))
                    {
                        SendKeyWithPostMessageOnly(TargetWindowHandle, evt.KeyCode, isDown);
                    }
                    else if (IsExtendedKey(evt.KeyCode))
                    {
                        if (CurrentArrowKeyMode == ArrowKeyMode.SendToChild)
                        {
                            SendKeyWithPostMessageOnly(TargetWindowHandle, evt.KeyCode, isDown);
                        }
                        else
                        {
                            SendKeyWithThreadAttach(TargetWindowHandle, evt.KeyCode, isDown);
                        }
                    }
                    else
                    {
                        SendKeyWithPostMessageOnly(TargetWindowHandle, evt.KeyCode, isDown);
                    }
                    Debug.WriteLine($"背景: {evt.KeyCode} ({evt.EventType})");
                }
                else
                {
                    SendKeyForeground(evt.KeyCode, isDown);
                    Debug.WriteLine($"前景: {evt.KeyCode} ({evt.EventType})");
                }

                if (evt.EventType == "down")
                {
                    PressedKeys.Add(evt.KeyCode);
                }
                else if (evt.EventType == "up")
                {
                    PressedKeys.Remove(evt.KeyCode);
                }
            }
            catch (Exception ex)
            {
                Log?.Invoke($"按鍵發送失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 發送自定義按鍵（按下後立即放開，單鍵）
        /// ★ 與主播放相同路由：方向鍵用 ATT，非方向鍵用純 PostMessage（不洩漏到前景）
        /// </summary>
        public void SendCustomKey(Keys key, Keys modifiers = Keys.None)
        {
            const int KEY_HOLD_MS = 60;

            try
            {
                if (TargetWindowHandle != IntPtr.Zero && IsWindow(TargetWindowHandle))
                {
                    if (IsArrowKey(key))
                    {
                        SendArrowKeyWithMode(TargetWindowHandle, key, true);
                        System.Threading.Thread.Sleep(KEY_HOLD_MS);
                        SendArrowKeyWithMode(TargetWindowHandle, key, false);
                    }
                    else if (IsAltKey(key))
                    {
                        SendAltKeyToWindow(TargetWindowHandle, key, true);
                        System.Threading.Thread.Sleep(KEY_HOLD_MS);
                        SendAltKeyToWindow(TargetWindowHandle, key, false);
                    }
                    else
                    {
                        SendKeyWithPostMessageOnly(TargetWindowHandle, key, true);
                        System.Threading.Thread.Sleep(KEY_HOLD_MS);
                        SendKeyWithPostMessageOnly(TargetWindowHandle, key, false);
                    }
                }
                else
                {
                    SendKeyForeground(key, true);
                    System.Threading.Thread.Sleep(KEY_HOLD_MS);
                    SendKeyForeground(key, false);
                }
            }
            catch (Exception ex)
            {
                Log?.Invoke($"自定義按鍵發送失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 供位置修正器使用的按鍵發送方法 — 與主播放完全相同的路由邏輯
        /// </summary>
        public void SendKeyForCorrection(IntPtr hWnd, Keys key, bool isKeyDown)
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
                        SendArrowKeyWithMode(hWnd, key, isKeyDown);
                    }
                    else
                    {
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

        /// <summary>
        /// 釋放所有目前按住的按鍵
        /// </summary>
        public void ReleasePressedKeys()
        {
            if (PressedKeys.Count == 0)
                return;

            Keys[] keysToRelease = PressedKeys.ToArray();
            PressedKeys.Clear();

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

        #endregion

        #region 低層按鍵發送實作

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

        /// <summary>
        /// 純 PostMessage 發送按鍵（不洩漏到前景）
        /// </summary>
        private void SendKeyWithPostMessageOnly(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            uint scanCode = MapVirtualKey((uint)key, MAPVK_VK_TO_VSC);
            bool isExtended = IsExtendedKey(key);

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

            if (!success)
            {
                SendMessage(hWnd, msg, (IntPtr)key, (IntPtr)lParamValue);
            }
        }

        /// <summary>
        /// 根據當前模式發送方向鍵
        /// </summary>
        private void SendArrowKeyWithMode(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            switch (CurrentArrowKeyMode)
            {
                case ArrowKeyMode.SendToChild:
                    SendKeyWithThreadAttach(hWnd, key, isKeyDown);
                    SendArrowKeyToChildWindow(hWnd, key, isKeyDown);
                    break;

                case ArrowKeyMode.ThreadAttachWithBlocker:
                    KeyboardBlocker?.RegisterPendingKey((uint)key);
                    SendKeyWithThreadAttach(hWnd, key, isKeyDown, (UIntPtr)KeyboardBlocker.MACRO_KEY_MARKER);
                    break;

                case ArrowKeyMode.SendInputWithBlock:
                    KeyboardBlocker?.RegisterPendingKey((uint)key);
                    SendKeyWithSendInput(key, isKeyDown);
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
                uint flags = 0;
                if (!isKeyDown) flags |= KEYEVENTF_KEYUP;
                if (isExtended) flags |= KEYEVENTF_EXTENDEDKEY;
                keybd_event((byte)key, (byte)scanCode, flags, (UIntPtr)KeyboardBlocker.MACRO_KEY_MARKER);
            }
        }

        /// <summary>
        /// 發送方向鍵到子視窗
        /// </summary>
        private void SendArrowKeyToChildWindow(IntPtr hWnd, Keys key, bool isKeyDown)
        {
            List<IntPtr> childWindows = new List<IntPtr>();

            EnumChildWindows(hWnd, (childHwnd, lParam) =>
            {
                if (IsWindowVisible(childHwnd))
                {
                    childWindows.Add(childHwnd);
                }
                return true;
            }, IntPtr.Zero);

            byte scanCode = GetScanCode(key);
            uint lParamValue;
            if (isKeyDown)
            {
                lParamValue = 1u | ((uint)scanCode << 16) | (1u << 24);
            }
            else
            {
                lParamValue = 1u | ((uint)scanCode << 16) | (1u << 24) | (1u << 30) | (1u << 31);
            }
            IntPtr lParam = (IntPtr)lParamValue;
            uint msg = isKeyDown ? WM_KEYDOWN : WM_KEYUP;

            foreach (IntPtr childHwnd in childWindows)
            {
                PostMessage(childHwnd, msg, (IntPtr)key, lParam);
            }

            PostMessage(hWnd, msg, (IntPtr)key, lParam);
            SendMessage(hWnd, msg, (IntPtr)key, lParam);
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

                byte vkCode;
                if (key == Keys.LMenu)
                    vkCode = 0xA4;
                else if (key == Keys.RMenu)
                    vkCode = 0xA5;
                else
                    vkCode = 0x12;

                byte scanCode = 0x38;

                uint flags = 0;
                if (!isKeyDown)
                {
                    flags |= KEYEVENTF_KEYUP;
                }
                if (key == Keys.RMenu)
                {
                    flags |= KEYEVENTF_EXTENDEDKEY;
                }

                keybd_event(vkCode, scanCode, flags, UIntPtr.Zero);

                uint lParamValue;
                if (isKeyDown)
                {
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
        /// 前景模式發送按鍵（使用 SendInput API）
        /// </summary>
        public void SendKeyForeground(Keys key, bool isKeyDown)
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

            if (result == 0)
            {
                int error = Marshal.GetLastWin32Error();
                Debug.WriteLine($"[SendInput FG] 失敗 (錯誤碼: {error})，使用 keybd_event 備援");

                uint flags = 0;
                if (!isKeyDown) flags |= KEYEVENTF_KEYUP;
                if (isExtended) flags |= KEYEVENTF_EXTENDEDKEY;
                keybd_event((byte)key, (byte)scanCode, flags, UIntPtr.Zero);
            }
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

        #endregion

        #region 文字發送

        /// <summary>
        /// 發送文字到背景視窗（使用 WM_CHAR 逐字發送）
        /// </summary>
        public void SendTextToWindow(IntPtr hWnd, string text)
        {
            foreach (char c in text)
            {
                PostMessage(hWnd, WM_CHAR, (IntPtr)c, IntPtr.Zero);
                System.Threading.Thread.Sleep(30);
            }
        }

        /// <summary>
        /// 發送文字到前景（使用 SendInput 逐字發送）
        /// </summary>
        public void SendTextForeground(string text)
        {
            foreach (char c in text)
            {
                INPUT[] inputs = new INPUT[2];

                inputs[0].type = INPUT_KEYBOARD;
                inputs[0].u.ki.wVk = 0;
                inputs[0].u.ki.wScan = (ushort)c;
                inputs[0].u.ki.dwFlags = 0x0004; // KEYEVENTF_UNICODE
                inputs[0].u.ki.time = 0;
                inputs[0].u.ki.dwExtraInfo = IntPtr.Zero;

                inputs[1].type = INPUT_KEYBOARD;
                inputs[1].u.ki.wVk = 0;
                inputs[1].u.ki.wScan = (ushort)c;
                inputs[1].u.ki.dwFlags = 0x0004 | KEYEVENTF_KEYUP;
                inputs[1].u.ki.time = 0;
                inputs[1].u.ki.dwExtraInfo = IntPtr.Zero;

                SendInput(2, inputs, Marshal.SizeOf(typeof(INPUT)));
                System.Threading.Thread.Sleep(30);
            }
        }

        #endregion

        #region 滑鼠 & 視窗

        /// <summary>
        /// 尋找遊戲渲染子視窗（MapleStoryClass 等）
        /// </summary>
        public IntPtr FindGameRenderWindow(IntPtr hWndParent)
        {
            StringBuilder className = new StringBuilder(256);
            GetClassName(hWndParent, className, 256);
            string parentClass = className.ToString();
            if (parentClass.IndexOf("MapleStory", StringComparison.OrdinalIgnoreCase) >= 0)
                return hWndParent;

            IntPtr found = IntPtr.Zero;
            EnumChildWindows(hWndParent, (childHwnd, lParam) =>
            {
                StringBuilder childClassName = new StringBuilder(256);
                GetClassName(childHwnd, childClassName, 256);
                if (childClassName.ToString().IndexOf("MapleStory", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    found = childHwnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);

            if (found == IntPtr.Zero)
                found = FindWindowEx(hWndParent, IntPtr.Zero, "MapleStoryClass", null);

            return found != IntPtr.Zero ? found : hWndParent;
        }

        /// <summary>
        /// 發送滑鼠點擊到視窗（Client Area 座標）
        /// ★ 雙重策略：PM 到 MapleStoryClass + 前景 SendInput 補強
        /// </summary>
        public void SendMouseClickToWindow(IntPtr hWnd, int clientX, int clientY)
        {
            IntPtr gameWnd = FindGameRenderWindow(hWnd);

            StringBuilder cn = new StringBuilder(256);
            GetClassName(gameWnd, cn, 256);
            string targetClass = cn.ToString();
            bool isGameClass = targetClass.IndexOf("MapleStory", StringComparison.OrdinalIgnoreCase) >= 0;

            Log?.Invoke($"🖱️ 目標視窗: [{targetClass}] hWnd=0x{gameWnd:X} isGame={isGameClass}");

            // 方式1: PostMessage 到 MapleStoryClass（背景點擊）
            IntPtr lParamCoord = (IntPtr)((clientY << 16) | (clientX & 0xFFFF));

            PostMessage(gameWnd, WM_MOUSEMOVE, IntPtr.Zero, lParamCoord);
            System.Threading.Thread.Sleep(50);

            PostMessage(gameWnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, lParamCoord);
            System.Threading.Thread.Sleep(100);
            PostMessage(gameWnd, WM_LBUTTONUP, IntPtr.Zero, lParamCoord);

            Log?.Invoke($"🖱️ PM 點擊已送出 → [{targetClass}] ({clientX},{clientY})");

            // 方式2: 前景 SendInput 補強
            System.Threading.Thread.Sleep(150);

            POINT pt = new POINT { X = clientX, Y = clientY };
            ClientToScreen(gameWnd, ref pt);

            SetForegroundWindow(hWnd);
            System.Threading.Thread.Sleep(150);

            SetCursorPos(pt.X, pt.Y);
            System.Threading.Thread.Sleep(50);

            INPUT[] inputs = new INPUT[2];
            inputs[0].type = INPUT_MOUSE;
            inputs[0].u.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
            inputs[0].u.mi.time = 0;
            inputs[0].u.mi.dwExtraInfo = IntPtr.Zero;

            inputs[1].type = INPUT_MOUSE;
            inputs[1].u.mi.dwFlags = MOUSEEVENTF_LEFTUP;
            inputs[1].u.mi.time = 0;
            inputs[1].u.mi.dwExtraInfo = IntPtr.Zero;

            uint result = SendInput(2, inputs, Marshal.SizeOf(typeof(INPUT)));

            Log?.Invoke($"🖱️ 前景點擊已送出 → screen({pt.X},{pt.Y}) result={result}");
        }

        #endregion
    }
}
