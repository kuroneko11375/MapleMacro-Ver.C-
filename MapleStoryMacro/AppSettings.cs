using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 應用設定類別 - 用於儲存全域操作設定 (.json)
    /// 不包含腳本特定功能（事件、自定義按鍵、循環次數）
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
        /// 方向鍵發送模式
        /// 0=SendToChild, 1=ThreadAttachWithBlocker, 2=SendInputWithBlock
        /// </summary>
        public int ArrowKeyMode { get; set; } = 0;

        /// <summary>
        /// 最後載入的腳本路徑（方便下次自動載入）
        /// </summary>
        public string? LastScriptPath { get; set; }

        /// <summary>
        /// 排程任務清單
        /// </summary>
        public List<ScheduleTask> ScheduleTasks { get; set; } = new List<ScheduleTask>();
    }

    /// <summary>
    /// 排程任務資料
    /// </summary>
    public class ScheduleTask
    {
        /// <summary>
        /// 腳本檔案路徑
        /// </summary>
        public string ScriptPath { get; set; } = string.Empty;

        /// <summary>
        /// 開始時間
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 結束時間（自動停止，null 表示不限制）
        /// </summary>
        public DateTime? EndTime { get; set; }

        /// <summary>
        /// 循環次數
        /// </summary>
        public int LoopCount { get; set; } = 1;

        /// <summary>
        /// 是否啟用
        /// </summary>
        public bool Enabled { get; set; } = true;

        /// <summary>
        /// 是否已觸發開始
        /// </summary>
        public bool HasStarted { get; set; } = false;

        /// <summary>
        /// 是否啟用回程功能（Enter → @FM → Enter → 等待 → 坐下）
        /// </summary>
        public bool ReturnToTownEnabled { get; set; } = false;

        /// <summary>
        /// 回程指令（預設 @FM）
        /// </summary>
        public string ReturnCommand { get; set; } = "@FM";

        /// <summary>
        /// 坐下按鍵（回程後自動觸發）
        /// </summary>
        public int SitDownKeyCode { get; set; } = (int)Keys.None;

        /// <summary>
        /// 坐下按鍵的修飾鍵
        /// </summary>
        public int SitDownKeyModifiers { get; set; } = (int)Keys.None;

        /// <summary>
        /// 回程後等待幾秒再坐下（預設 3 秒）
        /// </summary>
        public double SitDownDelaySeconds { get; set; } = 3.0;
    }

    /// <summary>
    /// 自定義按鍵槽位資料 (可序列化)
    /// </summary>
    public class CustomKeySlotData
    {
        public int SlotNumber { get; set; }
        public int KeyCode { get; set; } = (int)Keys.None;
        public int Modifiers { get; set; } = (int)Keys.None;
        public double IntervalSeconds { get; set; } = 30.0;
        public bool Enabled { get; set; } = false;
        public double StartAtSecond { get; set; } = 0;
        public double PreDelaySeconds { get; set; } = 0;
        public double PauseScriptSeconds { get; set; } = 0;
        public bool PauseScriptEnabled { get; set; } = false;
    }
}
