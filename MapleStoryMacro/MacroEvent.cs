using System;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    [Serializable]
    public class MacroEvent
    {
        public Keys KeyCode { get; set; }
        public string EventType { get; set; } = "down";
        public double Timestamp { get; set; }
        public int CorrectTargetX { get; set; } = -1;
        public int CorrectTargetY { get; set; } = -1;
        /// <summary>錄製時的小地圖座標（-1 = 未記錄）</summary>
        public int RecordedX { get; set; } = -1;
        public int RecordedY { get; set; } = -1;
        /// <summary>
        /// ★ 技能硬直時間（毫秒）：按鍵發送後鎖定位置修正器此時長。
        /// 0 = 不鎖定；>0 = 發送後 IsAnimationLocked=true 持續此毫秒數。
        /// </summary>
        public int SkillAnimationDelay { get; set; } = 0;
    }
}
