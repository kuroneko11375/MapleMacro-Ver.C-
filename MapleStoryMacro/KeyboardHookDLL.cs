using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace MapleStoryMacro
{
    /// <summary>
    /// Ы龄絃菏北 - ㄏノ GetKeyState 近高
 /// 露筁笴栏は国╰参
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

        // ===== 盽计﹚竡 =====
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
   private const int WM_KEYUP = 0x0101;

      // ===== 〆癠 =====
    public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        public delegate void KeyEventHandler(Keys keyCode, bool isKeyDown);
        public event KeyEventHandler? OnKeyEvent;

   // ===== Θ跑秖 =====
        private IntPtr hookHandle = IntPtr.Zero;
  private LowLevelKeyboardProc? hookProc = null;
        private bool isHookInstalled = false;
        private Task? pollingTask = null;
        private CancellationTokenSource? cancellationTokenSource = null;
 private Dictionary<int, bool> keyStates = new Dictionary<int, bool>();
    private readonly object lockObj = new object();

        // 菏北龄
        private readonly int[] MONITORED_KEYS = new int[]
        {
    0x20, // Space
   0x0D, // Enter
      0x1B, // Escape
     0x09, // Tab
            // ダ A-Z
       0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4A,
       0x4B, 0x4C, 0x4D, 0x4E, 0x4F, 0x50, 0x51, 0x52, 0x53, 0x54,
       0x55, 0x56, 0x57, 0x58, 0x59, 0x5A,
    // 计 0-9
         0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39,
    // よ龄
            0x25, 0x26, 0x27, 0x28,
            // 耿龄
       0x10, 0x11, 0x12, 0xA0, 0xA1, 0xA2, 0xA3,
        // F1-F12
0x70, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x7B,
            // ㄤ盽ノ龄
    0x2E, // Delete
            0x2D, // Insert
            0x21, // Page Up
0x22, // Page Down
      0x24, // Home
            0x23, // End
    };

        public KeyboardHookDLL()
        {
  }

        /// <summary>
        /// 杆龄絃菏北
 /// 纔ㄏノ夹非筥ア毖玥ㄏノ GetKeyState 近高
  /// </summary>
public bool Install()
        {
        if (isHookInstalled)
         return true;

            try
          {
           // 沽刚夹非筥
    hookProc = HookCallback;

       using (Process curProcess = Process.GetCurrentProcess())
       using (ProcessModule? curModule = curProcess.MainModule)
                {
     if (curModule == null)
     {
         System.Diagnostics.Debug.WriteLine("?? 礚猭眔家舱эノ GetKeyState 近高");
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
          System.Diagnostics.Debug.WriteLine("? 夹非龄絃筥杆");
        return true;
        }
    else
 {
     System.Diagnostics.Debug.WriteLine("?? 夹非筥ア毖эノ GetKeyState 近高");
          StartKeyStatePolling();
      return true;
         }
    }
            catch (Exception ex)
            {
       System.Diagnostics.Debug.WriteLine($"? 杆筥钵盽: {ex.Message}эノ近高");
       StartKeyStatePolling();
         return true;
     }
        }

     /// <summary>
        /// 币笆 GetKeyState 近高┏糷よ猭
   /// </summary>
        private void StartKeyStatePolling()
        {
   // ﹍て龄絃篈
         lock (lockObj)
            {
                foreach (int key in MONITORED_KEYS)
       {
        keyStates[key] = false;
         }
  }

          cancellationTokenSource = new CancellationTokenSource();
   pollingTask = Task.Run(() => KeyStatePollingThread(cancellationTokenSource.Token));
     System.Diagnostics.Debug.WriteLine("? GetKeyState 近高币笆");
        }

        /// <summary>
        /// GetKeyState 近高絬祘
        /// ㄏノ GetKeyState τぃ琌 GetAsyncKeyState螟砆は国ゎ
        /// </summary>
      private void KeyStatePollingThread(CancellationToken cancellationToken)
        {
            try
  {
 System.Diagnostics.Debug.WriteLine($"?? 近高絬祘币笆菏北 {MONITORED_KEYS.Length} 龄");

     while (!cancellationToken.IsCancellationRequested)
             {
    try
            {
             foreach (int vkCode in MONITORED_KEYS)
    {
   // GetKeyState: 蔼ボ讽玡篈
             short state = GetKeyState(vkCode);
        bool isPressed = (state & 0x8000) != 0;

  lock (lockObj)
   {
     bool wasPressed = keyStates.ContainsKey(vkCode) ? keyStates[vkCode] : false;

      // 浪代篈跑て
    if (isPressed && !wasPressed)
                    {
                 Keys key = (Keys)vkCode;
        OnKeyEvent?.Invoke(key, true);
    System.Diagnostics.Debug.WriteLine($"?? : {key}");
keyStates[vkCode] = true;
 }
       else if (!isPressed && wasPressed)
           {
        Keys key = (Keys)vkCode;
            OnKeyEvent?.Invoke(key, false);
   System.Diagnostics.Debug.WriteLine($"?? 睦: {key}");
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
 System.Diagnostics.Debug.WriteLine($"? 近高岿粇: {ex.Message}");
        }

     // 近高繵瞯10ms100Hz
        Thread.Sleep(10);
              }

   System.Diagnostics.Debug.WriteLine("?? 近高絬祘氨ゎ");
            }
            catch (Exception ex)
     {
        System.Diagnostics.Debug.WriteLine($"? 近高絬祘钵盽: {ex.Message}");
    }
        }

        /// <summary>
/// 更龄絃菏北
        /// </summary>
        public bool Uninstall()
        {
            try
        {
  // 氨ゎ近高
      if (pollingTask != null)
         {
              cancellationTokenSource?.Cancel();
     pollingTask?.Wait(2000);
     System.Diagnostics.Debug.WriteLine("? 近高氨ゎ");
       }

                // 更筥
    if (isHookInstalled && hookHandle != IntPtr.Zero)
              {
        if (!UnhookWindowsHookEx(hookHandle))
  {
          int errorCode = Marshal.GetLastWin32Error();
    System.Diagnostics.Debug.WriteLine($"?? 更筥ア毖: {errorCode}");
        }

                hookHandle = IntPtr.Zero;
             isHookInstalled = false;
     System.Diagnostics.Debug.WriteLine("? 龄絃筥更");
   }

                return true;
      }
      catch (Exception ex)
            {
        System.Diagnostics.Debug.WriteLine($"? 更钵盽: {ex.Message}");
         return false;
        }
   }

        /// <summary>
        /// 筥秸ㄧ计
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
        System.Diagnostics.Debug.WriteLine($"?? Hook : {keyCode}");
        else
      System.Diagnostics.Debug.WriteLine($"?? Hook 睦: {keyCode}");
        }
    }
            catch (Exception ex)
     {
          System.Diagnostics.Debug.WriteLine($"? Hook 秸岿粇: {ex.Message}");
            }

     return CallNextHookEx(hookHandle, nCode, wParam, lParam);
        }
    }
}
