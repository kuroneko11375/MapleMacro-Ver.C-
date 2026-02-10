using System;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 自定義按鍵槽位 - 可在腳本播放時按固定間隔自動施放
    /// </summary>
    public class CustomKeySlot
    {
        /// <summary>
        /// 槽位編號 (1-5)
        /// </summary>
        public int SlotNumber { get; set; }

        /// <summary>
        /// 要施放的按鍵
        /// </summary>
        public Keys KeyCode { get; set; } = Keys.None;

        /// <summary>
        /// 施放間隔（秒）
        /// </summary>
        public double IntervalSeconds { get; set; } = 30.0;

        /// <summary>
        /// 是否啟用
        /// </summary>
        public bool Enabled { get; set; } = false;

        /// <summary>
        /// 上次施放時間（用於計算下次施放）
        /// </summary>
        public DateTime LastTriggerTime { get; set; } = DateTime.MinValue;

        /// <summary>
        /// 插入時間點（腳本播放到第幾秒時開始計算，0 表示從頭開始）
        /// </summary>
        public double StartAtSecond { get; set; } = 0;

        /// <summary>
        /// 觸發前延遲（秒）- 按鍵施放前等待的時間
        /// </summary>
        public double PreDelaySeconds { get; set; } = 0;

        /// <summary>
        /// 腳本暫停時間（秒）- 觸發時暫停腳本執行的時間
        /// </summary>
        public double PauseScriptSeconds { get; set; } = 3.0;

        /// <summary>
        /// 是否啟用腳本暫停
        /// </summary>
        public bool PauseScriptEnabled { get; set; } = true;

        /// <summary>
        /// 檢查是否應該觸發
        /// </summary>
        public bool ShouldTrigger(double currentScriptTime)
        {
            if (!Enabled || KeyCode == Keys.None)
                return false;

            // 如果腳本時間還沒到開始時間點，不觸發
            if (currentScriptTime < StartAtSecond)
                return false;

            // 如果從未觸發過，且已經過了開始時間點，應該觸發
            if (LastTriggerTime == DateTime.MinValue)
                return true;

            // 檢查是否已經過了間隔時間
            double elapsed = (DateTime.Now - LastTriggerTime).TotalSeconds;
            return elapsed >= IntervalSeconds;
        }

        /// <summary>
        /// 標記為已觸發
        /// </summary>
        public void MarkTriggered()
        {
            LastTriggerTime = DateTime.Now;
        }

        /// <summary>
        /// 重置觸發狀態
        /// </summary>
        public void Reset()
        {
            LastTriggerTime = DateTime.MinValue;
        }

    }
}
