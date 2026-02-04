using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace MapleStoryMacro
{
    /// <summary>
    /// 全局鍵盤監控 - 使用 GetKeyState 輪詢
 /// 可以繞過遊戲反作弊系統
    /// </summary>
    public class KeyboardHookDLL
    {
        // ===== Windows API P/Invoke =====
        [DllImport("user32.dll", SetLastError = true)]
  private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

     [DllImport("user32.dll", SetLastError = true)]
  private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

      [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        // ===== 常數定義 =====
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
   private const int WM_KEYUP = 0x0101;

      // ===== 委託 =====
    public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        public delegate void KeyEventHandler(Keys keyCode, bool isKeyDown);
        public event KeyEventHandler? OnKeyEvent;

   // ===== 成員變量 =====
        private IntPtr hookHandle = IntPtr.Zero;
  private LowLevelKeyboardProc? hookProc = null;
        private bool isHookInstalled = false;
        private Task? pollingTask = null;
        private CancellationTokenSource? cancellationTokenSource = null;
 private Dictionary<int, bool> keyStates = new Dictionary<int, bool>();
    private readonly object lockObj = new object();

        // 監控的按鍵列表 - 精簡為常用熱鍵
        private readonly int[] MONITORED_KEYS = new int[]
        {
            // F1-F12 (熱鍵最常用)
            0x70, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x7B,
            // 常用控制鍵
            0x1B, // Escape
            0x20, // Space
            0x0D, // Enter
        };

        public KeyboardHookDLL()
        {
  }

        /// <summary>
        /// 安裝鍵盤監控
 /// 優先使用標準鉤子，失敗則使用 GetKeyState 輪詢
  /// </summary>
public bool Install()
        {
        if (isHookInstalled)
         return true;

            try
          {
           // 先嘗試標準鉤子
    hookProc = HookCallback;

       using (Process curProcess = Process.GetCurrentProcess())
       using (ProcessModule? curModule = curProcess.MainModule)
                {
     if (curModule == null)
     {
         System.Diagnostics.Debug.WriteLine("?? 無法取得模組，改用 GetKeyState 輪詢");
         StartKeyStatePolling();
         return true;
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
              isHookInstalled = true;
          System.Diagnostics.Debug.WriteLine("? 標準鍵盤鉤子已安裝");
        return true;
        }
    else
 {
     System.Diagnostics.Debug.WriteLine("?? 標準鉤子失敗，改用 GetKeyState 輪詢");
          StartKeyStatePolling();
      return true;
         }
    }
            catch (Exception ex)
            {
       System.Diagnostics.Debug.WriteLine($"? 安裝鉤子異常: {ex.Message}，改用輪詢");
       StartKeyStatePolling();
         return true;
     }
        }

     /// <summary>
        /// 啟動 GetKeyState 輪詢（更底層的方法）
   /// </summary>
        private void StartKeyStatePolling()
        {
   // 初始化鍵盤狀態
         lock (lockObj)
            {
                foreach (int key in MONITORED_KEYS)
       {
        keyStates[key] = false;
         }
  }

          cancellationTokenSource = new CancellationTokenSource();
   pollingTask = Task.Run(() => KeyStatePollingThread(cancellationTokenSource.Token));
     System.Diagnostics.Debug.WriteLine("? GetKeyState 輪詢已啟動");
        }

        /// <summary>
        /// GetKeyState 輪詢線程
        /// 使用 GetKeyState 而不是 GetAsyncKeyState（更難被反作弊阻止）
        /// </summary>
      private void KeyStatePollingThread(CancellationToken cancellationToken)
        {
            try
  {
 System.Diagnostics.Debug.WriteLine($"?? 輪詢線程啟動，監控 {MONITORED_KEYS.Length} 個按鍵");

     while (!cancellationToken.IsCancellationRequested)
             {
    try
            {
             foreach (int vkCode in MONITORED_KEYS)
    {
   // GetKeyState: 返回值高位表示當前狀態
             short state = GetKeyState(vkCode);
        bool isPressed = (state & 0x8000) != 0;

  lock (lockObj)
   {
     bool wasPressed = keyStates.ContainsKey(vkCode) ? keyStates[vkCode] : false;

      // 檢測狀態變化
    if (isPressed && !wasPressed)
                    {
                 Keys key = (Keys)vkCode;
        OnKeyEvent?.Invoke(key, true);
    System.Diagnostics.Debug.WriteLine($"?? 按下: {key}");
keyStates[vkCode] = true;
 }
       else if (!isPressed && wasPressed)
           {
        Keys key = (Keys)vkCode;
            OnKeyEvent?.Invoke(key, false);
   System.Diagnostics.Debug.WriteLine($"?? 釋放: {key}");
         keyStates[vkCode] = false;
        }
           else
       {
       keyStates[vkCode] = isPressed;
        }
          }
      }
  }
         catch (Exception ex)
       {
 System.Diagnostics.Debug.WriteLine($"? 輪詢錯誤: {ex.Message}");
        }

      // 輪詢頻率：50ms（20Hz）- 對於熱鍵監控已足夠
        Thread.Sleep(50);
              }

   System.Diagnostics.Debug.WriteLine("?? 輪詢線程已停止");
            }
            catch (Exception ex)
     {
        System.Diagnostics.Debug.WriteLine($"? 輪詢線程異常: {ex.Message}");
    }
        }

        /// <summary>
/// 卸載鍵盤監控
        /// </summary>
        public bool Uninstall()
        {
            try
        {
  // 停止輪詢
      if (pollingTask != null)
         {
              cancellationTokenSource?.Cancel();
     pollingTask?.Wait(2000);
     System.Diagnostics.Debug.WriteLine("? 輪詢已停止");
       }

                // 卸載鉤子
    if (isHookInstalled && hookHandle != IntPtr.Zero)
              {
        if (!UnhookWindowsHookEx(hookHandle))
  {
          int errorCode = Marshal.GetLastWin32Error();
    System.Diagnostics.Debug.WriteLine($"?? 卸載鉤子失敗: {errorCode}");
        }

                hookHandle = IntPtr.Zero;
             isHookInstalled = false;
     System.Diagnostics.Debug.WriteLine("? 鍵盤鉤子已卸載");
   }

                return true;
      }
      catch (Exception ex)
            {
        System.Diagnostics.Debug.WriteLine($"? 卸載異常: {ex.Message}");
         return false;
        }
   }

        /// <summary>
        /// 鉤子回調函數
        /// </summary>
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
       {
                if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_KEYUP))
      {
      int vkCode = Marshal.ReadInt32(lParam);
           Keys keyCode = (Keys)vkCode;
   bool isKeyDown = (wParam == (IntPtr)WM_KEYDOWN);

      OnKeyEvent?.Invoke(keyCode, isKeyDown);

       if (isKeyDown)
        System.Diagnostics.Debug.WriteLine($"?? Hook 按下: {keyCode}");
        else
      System.Diagnostics.Debug.WriteLine($"?? Hook 釋放: {keyCode}");
        }
    }
            catch (Exception ex)
     {
          System.Diagnostics.Debug.WriteLine($"? Hook 回調錯誤: {ex.Message}");
            }

     return CallNextHookEx(hookHandle, nCode, wParam, lParam);
        }
    }
}
