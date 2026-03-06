using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 腳本資料類別 - 用於儲存和載入腳本檔案 (.mscript)
    /// 包含：錄製事件 + 自定義按鍵設定 + 循環次數
    /// </summary>
    public class ScriptData
    {
        /// <summary>
        /// 腳本版本（用於未來相容性）
        /// </summary>
        public int Version { get; set; } = 2; // ← 升版到 2，代表支持路徑座標

        /// <summary>
        /// 腳本名稱（可選）
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// 最後修改時間
        /// </summary>
        public DateTime ModifiedAt { get; set; } = DateTime.Now;

        /// <summary>
        /// 循環次數
        /// </summary>
        public int LoopCount { get; set; } = 1;

        /// <summary>
        /// 錄製的事件列表
        /// </summary>
        public List<ScriptEvent> Events { get; set; } = new List<ScriptEvent>();

        /// <summary>
        /// 自定義按鍵槽位設定 (15個)
        /// </summary>
        public CustomKeySlotData[] CustomKeySlots { get; set; } = new CustomKeySlotData[15];

        /// <summary>
        /// 是否包含路徑座標資料（錄製時有小地圖校準）
        /// </summary>
        public bool HasPathData { get; set; } = false;

        public ScriptData()
        {
            // 初始化自定義按鍵槽位
            for (int i = 0; i < 15; i++)
            {
                CustomKeySlots[i] = new CustomKeySlotData { SlotNumber = i + 1 };
            }
        }
    }

    /// <summary>
    /// 腳本事件（可序列化版本的 MacroEvent）
    /// </summary>
    public class ScriptEvent
    {
        public int KeyCode { get; set; }
        public string EventType { get; set; } = "down";
        public double Timestamp { get; set; }

        /// <summary>錄製時的小地圖 X 座標（-1 = 未記錄）</summary>
        public int RecordedX { get; set; } = -1;

        /// <summary>錄製時的小地圖 Y 座標（-1 = 未記錄）</summary>
        public int RecordedY { get; set; } = -1;

        /// <summary>
        /// ★ 技能硬直時間（毫秒）：按鍵發送後鎖定位置修正器此時長。
        /// 0 = 不鎖定。在腳本編輯器中可為長硬直技能設定此值。
        /// </summary>
        public int SkillAnimationDelay { get; set; } = 0;
    }
}
