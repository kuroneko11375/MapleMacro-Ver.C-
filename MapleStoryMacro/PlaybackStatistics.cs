using System;

namespace MapleStoryMacro
{
    /// <summary>
    /// 執行統計資料
    /// </summary>
    public class PlaybackStatistics
    {
        /// <summary>
        /// 總執行次數
        /// </summary>
        public int TotalPlayCount { get; set; } = 0;

        /// <summary>
        /// 總執行時長（秒）
        /// </summary>
        public double TotalPlayTimeSeconds { get; set; } = 0;

        /// <summary>
        /// 最後執行時間
        /// </summary>
        public DateTime? LastPlayTime { get; set; } = null;

        /// <summary>
        /// 本次開始時間
        /// </summary>
        public DateTime? CurrentSessionStart { get; set; } = null;

        /// <summary>
        /// 本次執行的循環數
        /// </summary>
        public int CurrentLoopCount { get; set; } = 0;

        /// <summary>
        /// 自定義按鍵觸發次數
        /// </summary>
        public int[] CustomKeyTriggerCounts { get; set; } = new int[15];

        /// <summary>
        /// 開始新的播放階段
        /// </summary>
        public void StartSession()
        {
            CurrentSessionStart = DateTime.Now;
            CurrentLoopCount = 0;
            for (int i = 0; i < 15; i++)
                CustomKeyTriggerCounts[i] = 0;
        }

        /// <summary>
        /// 結束播放階段
        /// </summary>
        public void EndSession()
        {
            if (CurrentSessionStart.HasValue)
            {
                TotalPlayTimeSeconds += (DateTime.Now - CurrentSessionStart.Value).TotalSeconds;
                LastPlayTime = DateTime.Now;
            }
            TotalPlayCount++;
            CurrentSessionStart = null;
        }

        /// <summary>
        /// 增加循環計數
        /// </summary>
        public void IncrementLoop()
        {
            CurrentLoopCount++;
        }

        /// <summary>
        /// 記錄自定義按鍵觸發
        /// </summary>
        public void RecordCustomKeyTrigger(int slotIndex)
        {
            if (slotIndex >= 0 && slotIndex < 15)
                CustomKeyTriggerCounts[slotIndex]++;
        }

        /// <summary>
        /// 取得格式化的統計資訊
        /// </summary>
        public string GetFormattedStats()
        {
            string lastPlay = LastPlayTime.HasValue
                ? LastPlayTime.Value.ToString("yyyy-MM-dd HH:mm:ss")
                : "從未執行";

            TimeSpan totalTime = TimeSpan.FromSeconds(TotalPlayTimeSeconds);
            string totalTimeStr = $"{(int)totalTime.TotalHours:D2}:{totalTime.Minutes:D2}:{totalTime.Seconds:D2}";

            return $"總執行次數: {TotalPlayCount}\n" +
                   $"總執行時長: {totalTimeStr}\n" +
                   $"最後執行: {lastPlay}";
        }

        /// <summary>
        /// 重置所有統計
        /// </summary>
        public void Reset()
        {
            TotalPlayCount = 0;
            TotalPlayTimeSeconds = 0;
            LastPlayTime = null;
            CurrentSessionStart = null;
            CurrentLoopCount = 0;
            CustomKeyTriggerCounts = new int[15];
        }
    }
}
