using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MapleStoryMacro
{
    /// <summary>
    /// Rust Native DLL (macro_core.dll) 的 P/Invoke 封裝
    ///
    /// 提供三大核心功能：
    /// 1. Hardware-Level Keyboard Hook - 微秒級攔截，絕不超時
    /// 2. Flash Focus Key Sending - &lt;1ms 內完成背景按鍵發送
    /// 3. dwExtraInfo Filtering - MAGIC_TAG 精確識別
    /// </summary>
    public static class MacroCoreNative
    {
        private const string DLL_NAME = "macro_core.dll";

        /// <summary>
        /// 健康檢查 - 回傳 0xDEADBEEF 表示 DLL 正常載入
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint mc_ping();

        /// <summary>
        /// 設定目標遊戲視窗 HWND
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern void mc_set_target_hwnd(IntPtr hwnd);

        /// <summary>
        /// 取得目標遊戲視窗 HWND
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr mc_get_target_hwnd();

        /// <summary>
        /// 安裝鍵盤攔截 Hook (WH_KEYBOARD_LL)
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.U1)]
        public static extern bool mc_install_hook();

        /// <summary>
        /// 卸載鍵盤攔截 Hook
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern void mc_uninstall_hook();

        /// <summary>
        /// 啟用/停用按鍵攔截
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern void mc_set_blocking([MarshalAs(UnmanagedType.U1)] bool enabled);

        /// <summary>
        /// 查詢攔截是否啟用
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.U1)]
        public static extern bool mc_is_blocking();

        /// <summary>
        /// Flash Focus 發送按鍵到遊戲視窗
        /// 在 &lt;1ms 內完成 AttachThreadInput + SetFocus + SendInput + 還原焦點
        /// </summary>
        /// <param name="vkCode">虛擬按鍵碼</param>
        /// <param name="isKeyDown">true=按下, false=放開</param>
        /// <returns>0=成功, -1=無目標視窗, -2=Attach失敗, -3=SendInput失敗</returns>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern int mc_flash_send_key(ushort vkCode, [MarshalAs(UnmanagedType.U1)] bool isKeyDown);

        /// <summary>
        /// 前景模式 SendInput（帶 MAGIC_TAG 讓 Hook 能識別）
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern int mc_send_input_foreground(ushort vkCode, [MarshalAs(UnmanagedType.U1)] bool isKeyDown);

        /// <summary>
        /// 取得被攔截的按鍵計數
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint mc_get_blocked_count();

        /// <summary>
        /// 取得放行的按鍵計數
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint mc_get_passed_count();

        /// <summary>
        /// 重置統計計數
        /// </summary>
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        public static extern void mc_reset_stats();

        /// <summary>
        /// 檢查 Rust DLL 是否可用
        /// </summary>
        public static bool IsAvailable()
        {
            try
            {
                return mc_ping() == 0xDEADBEEF;
            }
            catch (DllNotFoundException)
            {
                return false;
            }
            catch (EntryPointNotFoundException)
            {
                return false;
            }
        }
    }

    /// <summary>
    /// Rust 鍵盤引擎 - 封裝 MacroCoreNative 提供高階 API
    /// 作為 KeyboardBlocker 的替代品使用
    /// </summary>
    public class RustKeyEngine : IDisposable
    {
        private bool _hookInstalled = false;
        private bool _disposed = false;

        /// <summary>
        /// Rust DLL 是否可用
        /// </summary>
        public bool IsRustAvailable { get; }

        /// <summary>
        /// 目標遊戲視窗
        /// </summary>
        public IntPtr TargetWindowHandle
        {
            get => MacroCoreNative.mc_get_target_hwnd();
            set => MacroCoreNative.mc_set_target_hwnd(value);
        }

        /// <summary>
        /// 攔截是否啟用
        /// </summary>
        public bool IsBlocking
        {
            get => MacroCoreNative.mc_is_blocking();
            set => MacroCoreNative.mc_set_blocking(value);
        }

        /// <summary>
        /// 攔截計數
        /// </summary>
        public uint BlockedKeyCount => MacroCoreNative.mc_get_blocked_count();

        /// <summary>
        /// 放行計數
        /// </summary>
        public uint PassedKeyCount => MacroCoreNative.mc_get_passed_count();

        public RustKeyEngine()
        {
            IsRustAvailable = MacroCoreNative.IsAvailable();
            if (IsRustAvailable)
            {
                Debug.WriteLine("[RustKeyEngine] Rust DLL 載入成功 ?");
            }
            else
            {
                Debug.WriteLine("[RustKeyEngine] Rust DLL 不可用，將回退到 C# 實作");
            }
        }

        /// <summary>
        /// 安裝 Hook
        /// </summary>
        public bool Install()
        {
            if (!IsRustAvailable) return false;
            if (_hookInstalled) return true;

            _hookInstalled = MacroCoreNative.mc_install_hook();
            Debug.WriteLine($"[RustKeyEngine] Hook 安裝: {(_hookInstalled ? "成功" : "失敗")}");
            return _hookInstalled;
        }

        /// <summary>
        /// 卸載 Hook
        /// </summary>
        public void Uninstall()
        {
            if (!IsRustAvailable) return;
            MacroCoreNative.mc_uninstall_hook();
            _hookInstalled = false;
            Debug.WriteLine("[RustKeyEngine] Hook 已卸載");
        }

        /// <summary>
        /// Flash Focus 發送按鍵（核心方法 - 背景模式）
        /// </summary>
        public int FlashSendKey(ushort vkCode, bool isKeyDown)
        {
            return MacroCoreNative.mc_flash_send_key(vkCode, isKeyDown);
        }

        /// <summary>
        /// 前景模式發送
        /// </summary>
        public int SendForeground(ushort vkCode, bool isKeyDown)
        {
            return MacroCoreNative.mc_send_input_foreground(vkCode, isKeyDown);
        }

        /// <summary>
        /// 重置統計
        /// </summary>
        public void ResetStats()
        {
            if (IsRustAvailable)
                MacroCoreNative.mc_reset_stats();
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                Uninstall();
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }

        ~RustKeyEngine()
        {
            Dispose();
        }
    }
}
