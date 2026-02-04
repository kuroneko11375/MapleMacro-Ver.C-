using System;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 應用程式設定類別 - 用於儲存和載入所有設定
    /// </summary>
    public class AppSettings
    {
        /// <summary>
        /// 播放熱鍵
        /// </summary>
        public Keys PlayHotkey { get; set; } = Keys.F9;

        /// <summary>
        /// 停止熱鍵
        /// </summary>
        public Keys StopHotkey { get; set; } = Keys.F10;

        /// <summary>
        /// 熱鍵是否啟用
        /// </summary>
        public bool HotkeyEnabled { get; set; } = true;

        /// <summary>
        /// 目標視窗標題
        /// </summary>
        public string WindowTitle { get; set; } = "MapleStory";


        /// <summary>
        /// 循環次數
        /// </summary>
        public int LoopCount { get; set; } = 1;

        /// <summary>
        /// 方向鍵發送模式
        /// </summary>
        public int ArrowKeyMode { get; set; } = 2;

        /// <summary>
        /// 背景切換模式
        /// </summary>
        public int BackgroundSwitchMode { get; set; } = 3; // 預設 1.0 秒

        /// <summary>
        /// 自定義按鍵槽位設定 (15個)
        /// </summary>
        public CustomKeySlotData[] CustomKeySlots { get; set; } = new CustomKeySlotData[15];

        public AppSettings()
        {
            // 初始化自定義按鍵槽位
            for (int i = 0; i < 15; i++)
            {
                CustomKeySlots[i] = new CustomKeySlotData { SlotNumber = i + 1 };
            }
        }
    }

    /// <summary>
    /// 自定義按鍵槽位資料 (可序列化)
    /// </summary>
    public class CustomKeySlotData
    {
        public int SlotNumber { get; set; }
        public int KeyCode { get; set; } = (int)Keys.None;
        public double IntervalSeconds { get; set; } = 30.0;
        public bool Enabled { get; set; } = false;
        public double StartAtSecond { get; set; } = 0;
        public double PreDelaySeconds { get; set; } = 0;
        public double PauseScriptSeconds { get; set; } = 0;
        public bool PauseScriptEnabled { get; set; } = false;
    }
}
