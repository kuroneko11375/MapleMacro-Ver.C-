using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 按鍵捕獲輔助類別 - 支援跨鍵盤布局的按鍵處理
    /// 解決不同電腦、不同鍵盤布局按鍵失效的問題
    /// </summary>
    public static class KeyCaptureHelper
    {
        #region Windows API

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKeyEx(uint uCode, uint uMapType, IntPtr dwhkl);

        [DllImport("user32.dll")]
        private static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetKeyNameTextW(int lParam, StringBuilder lpString, int nSize);

        [DllImport("user32.dll")]
        private static extern short VkKeyScanEx(char ch, IntPtr dwhkl);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        // SendInput 相關
        [DllImport("user32.dll")]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        private const uint MAPVK_VK_TO_VSC = 0;      // VirtualKey → ScanCode
        private const uint MAPVK_VSC_TO_VK = 1;      // ScanCode → VirtualKey
        private const uint MAPVK_VK_TO_CHAR = 2;     // VirtualKey → Character
        private const uint MAPVK_VSC_TO_VK_EX = 3;   // ScanCode → VirtualKey (區分左右)
        
        private const int INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;
        private const uint KEYEVENTF_SCANCODE = 0x0008;

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public int type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        #endregion

        /// <summary>
        /// 按鍵資訊 - 包含完整的跨平台資訊
        /// </summary>
        public class KeyInfo
        {
            /// <summary>
            /// 虛擬鍵碼（Virtual Key Code）
            /// </summary>
            public int VirtualKeyCode { get; set; }

            /// <summary>
            /// 硬體掃描碼（Hardware Scan Code）- 跨布局一致
            /// </summary>
            public uint ScanCode { get; set; }

            /// <summary>
            /// 按鍵名稱（用於顯示）
            /// </summary>
            public string KeyName { get; set; } = string.Empty;

            /// <summary>
            /// 是否為擴展鍵（Extended Key）
            /// 方向鍵、Home、End、Insert、Delete、PageUp、PageDown 等
            /// </summary>
            public bool IsExtendedKey { get; set; }

            /// <summary>
            /// 鍵盤布局 ID
            /// </summary>
            public int LayoutId { get; set; }

            /// <summary>
            /// 字元表示（如果有）
            /// </summary>
            public char Character { get; set; }

            public override string ToString()
            {
                return $"{KeyName} (VK:0x{VirtualKeyCode:X2}, SC:0x{ScanCode:X2}{(IsExtendedKey ? ", Ext" : "")})";
            }
        }

        /// <summary>
        /// 捕獲按鍵的完整資訊（建議在 UI 設定按鍵時調用）
        /// </summary>
        /// <param name="keyCode">Windows Forms Keys 枚舉值</param>
        /// <returns>包含跨平台資訊的 KeyInfo</returns>
        public static KeyInfo CaptureKey(Keys keyCode)
        {
            IntPtr layout = GetKeyboardLayout(0);
            uint vk = (uint)keyCode;

            // 1. 取得掃描碼
            uint scanCode = MapVirtualKeyEx(vk, MAPVK_VK_TO_VSC, layout);

            // 2. 判斷是否為擴展鍵
            bool isExtended = IsExtendedKey(keyCode);

            // 3. 取得按鍵名稱
            string keyName = GetKeyName(scanCode, isExtended);
            if (string.IsNullOrEmpty(keyName))
            {
                keyName = keyCode.ToString();
            }

            // 4. 嘗試取得字元
            char character = GetCharacterFromKey(keyCode, layout);

            return new KeyInfo
            {
                VirtualKeyCode = (int)keyCode,
                ScanCode = scanCode,
                KeyName = keyName,
                IsExtendedKey = isExtended,
                LayoutId = layout.ToInt32(),
                Character = character
            };
        }

        /// <summary>
        /// 虛擬鍵碼轉掃描碼（快速方法，不依賴布局）
        /// </summary>
        public static uint VirtualKeyToScanCode(Keys keyCode)
        {
            return MapVirtualKey((uint)keyCode, MAPVK_VK_TO_VSC);
        }

        /// <summary>
        /// 掃描碼轉虛擬鍵碼（快速方法）
        /// </summary>
        public static Keys ScanCodeToVirtualKey(uint scanCode)
        {
            return (Keys)MapVirtualKey(scanCode, MAPVK_VSC_TO_VK);
        }

        /// <summary>
        /// 判斷是否為擴展鍵
        /// </summary>
        public static bool IsExtendedKey(Keys keyCode)
        {
            // 擴展鍵列表（需要 0xE0 前綴的按鍵）
            switch (keyCode)
            {
                // 方向鍵
                case Keys.Up:
                case Keys.Down:
                case Keys.Left:
                case Keys.Right:
                // 編輯鍵
                case Keys.Insert:
                case Keys.Delete:
                case Keys.Home:
                case Keys.End:
                case Keys.PageUp:
                case Keys.PageDown:
                // 右側修飾鍵
                case Keys.RControlKey:
                case Keys.RMenu: // Right Alt
                case Keys.RWin:
                // 其他
                case Keys.NumLock:
                case Keys.Divide:  // 數字鍵盤的 /
                case Keys.Enter when ((int)keyCode == 0x0D): // 數字鍵盤的 Enter（需進一步判斷）
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 取得按鍵的本地化名稱
        /// </summary>
        private static string GetKeyName(uint scanCode, bool isExtended)
        {
            StringBuilder keyName = new StringBuilder(32);
            int lParam = (int)(scanCode << 16);
            if (isExtended)
            {
                lParam |= 0x01000000; // 設置 Extended-key 旗標
            }

            int result = GetKeyNameTextW(lParam, keyName, keyName.Capacity);
            return result > 0 ? keyName.ToString() : string.Empty;
        }

        /// <summary>
        /// 從按鍵取得對應的字元（如果有）
        /// </summary>
        private static char GetCharacterFromKey(Keys keyCode, IntPtr layout)
        {
            // 嘗試將 VK 轉換為字元
            uint charValue = MapVirtualKeyEx((uint)keyCode, MAPVK_VK_TO_CHAR, layout);
            if (charValue > 0 && charValue < 0xFFFF)
            {
                return (char)charValue;
            }
            return '\0';
        }

        #region 發送按鍵（使用 ScanCode）

        /// <summary>
        /// 使用掃描碼發送按鍵（推薦方法，跨布局一致）
        /// </summary>
        /// <param name="scanCode">硬體掃描碼</param>
        /// <param name="isExtended">是否為擴展鍵</param>
        /// <param name="isKeyDown">true=按下, false=放開</param>
        /// <param name="extraInfo">額外資訊（用於標記巨集按鍵）</param>
        public static void SendKeyByScanCode(uint scanCode, bool isExtended, bool isKeyDown, uint extraInfo = 0)
        {
            INPUT input = new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = 0, // 使用掃描碼時，VK 設為 0
                        wScan = (ushort)scanCode,
                        dwFlags = KEYEVENTF_SCANCODE |
                                 (isExtended ? KEYEVENTF_EXTENDEDKEY : 0) |
                                 (isKeyDown ? 0 : KEYEVENTF_KEYUP),
                        time = 0,
                        dwExtraInfo = (UIntPtr)extraInfo
                    }
                }
            };

            SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
        }

        /// <summary>
        /// 發送按鍵組合（使用掃描碼）
        /// </summary>
        /// <param name="modifiersScanCode">修飾鍵掃描碼（Ctrl/Alt/Shift）</param>
        /// <param name="keyScanCode">主按鍵掃描碼</param>
        /// <param name="isExtended">主按鍵是否為擴展鍵</param>
        /// <param name="extraInfo">額外資訊</param>
        public static void SendKeyCombination(uint modifiersScanCode, uint keyScanCode, bool isExtended, uint extraInfo = 0)
        {
            // 1. 按下修飾鍵
            if (modifiersScanCode != 0)
            {
                SendKeyByScanCode(modifiersScanCode, false, true, extraInfo);
                System.Threading.Thread.Sleep(10); // 短暫延遲確保按鍵順序
            }

            // 2. 按下主按鍵
            SendKeyByScanCode(keyScanCode, isExtended, true, extraInfo);
            System.Threading.Thread.Sleep(50); // 模擬人類按鍵時間

            // 3. 放開主按鍵
            SendKeyByScanCode(keyScanCode, isExtended, false, extraInfo);

            // 4. 放開修飾鍵
            if (modifiersScanCode != 0)
            {
                System.Threading.Thread.Sleep(10);
                SendKeyByScanCode(modifiersScanCode, false, false, extraInfo);
            }
        }

        /// <summary>
        /// 從 CustomKeySlotData 發送按鍵（智能選擇 ScanCode 或 VirtualKey）
        /// </summary>
        public static void SendKeyFromSlot(CustomKeySlotData slot)
        {
            uint extraInfo = KeyboardBlocker.MACRO_KEY_MARKER;

            // 優先使用 ScanCode（跨布局一致）
            if (slot.ScanCode != 0)
            {
                SendKeyCombination(slot.ModifiersScanCode, slot.ScanCode, slot.IsExtendedKey, extraInfo);
            }
            // 回退到 VirtualKey（向後兼容）
            else if (slot.KeyCode != (int)Keys.None)
            {
                SendKeyByVirtualKey((Keys)slot.KeyCode, (Keys)slot.Modifiers, extraInfo);
            }
        }

        /// <summary>
        /// 使用虛擬鍵碼發送按鍵（兼容舊版本）
        /// </summary>
        private static void SendKeyByVirtualKey(Keys keyCode, Keys modifiers, uint extraInfo)
        {
            var inputs = new System.Collections.Generic.List<INPUT>();

            // 按下修飾鍵
            if (modifiers.HasFlag(Keys.Control))
            {
                inputs.Add(CreateKeyInput(Keys.ControlKey, true, extraInfo));
            }
            if (modifiers.HasFlag(Keys.Alt))
            {
                inputs.Add(CreateKeyInput(Keys.Menu, true, extraInfo));
            }
            if (modifiers.HasFlag(Keys.Shift))
            {
                inputs.Add(CreateKeyInput(Keys.ShiftKey, true, extraInfo));
            }

            // 按下主按鍵
            inputs.Add(CreateKeyInput(keyCode, true, extraInfo));

            // 放開主按鍵
            inputs.Add(CreateKeyInput(keyCode, false, extraInfo));

            // 放開修飾鍵（逆序）
            if (modifiers.HasFlag(Keys.Shift))
            {
                inputs.Add(CreateKeyInput(Keys.ShiftKey, false, extraInfo));
            }
            if (modifiers.HasFlag(Keys.Alt))
            {
                inputs.Add(CreateKeyInput(Keys.Menu, false, extraInfo));
            }
            if (modifiers.HasFlag(Keys.Control))
            {
                inputs.Add(CreateKeyInput(Keys.ControlKey, false, extraInfo));
            }

            if (inputs.Count > 0)
            {
                SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
            }
        }

        private static INPUT CreateKeyInput(Keys keyCode, bool isKeyDown, uint extraInfo)
        {
            bool isExtended = IsExtendedKey(keyCode);
            return new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = (ushort)keyCode,
                        wScan = 0,
                        dwFlags = (isExtended ? KEYEVENTF_EXTENDEDKEY : 0) |
                                 (isKeyDown ? 0 : KEYEVENTF_KEYUP),
                        time = 0,
                        dwExtraInfo = (UIntPtr)extraInfo
                    }
                }
            };
        }

        #endregion

        #region 修飾鍵處理

        /// <summary>
        /// 取得修飾鍵的掃描碼
        /// </summary>
        public static uint GetModifierScanCode(Keys modifiers)
        {
            // 注意：如果有多個修飾鍵，只返回第一個
            // 實際使用時可能需要分別處理
            if (modifiers.HasFlag(Keys.Control))
            {
                return VirtualKeyToScanCode(Keys.ControlKey);
            }
            if (modifiers.HasFlag(Keys.Alt))
            {
                return VirtualKeyToScanCode(Keys.Menu);
            }
            if (modifiers.HasFlag(Keys.Shift))
            {
                return VirtualKeyToScanCode(Keys.ShiftKey);
            }
            return 0;
        }

        /// <summary>
        /// 取得修飾鍵名稱
        /// </summary>
        public static string GetModifierName(Keys modifiers)
        {
            var names = new System.Collections.Generic.List<string>();
            if (modifiers.HasFlag(Keys.Control))
                names.Add("Ctrl");
            if (modifiers.HasFlag(Keys.Alt))
                names.Add("Alt");
            if (modifiers.HasFlag(Keys.Shift))
                names.Add("Shift");
            return string.Join("+", names);
        }

        #endregion
    }
}
