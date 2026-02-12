using System;

namespace MapleStoryMacro
{
    /// <summary>
    /// 播放統計類別
    /// </summary>
    public class PlaybackStatistics
    {
        /// <summary>
        /// 總播放次數
        /// </summary>
        public int TotalPlayCount { get; set; } = 0;

        /// <summary>
        /// 總播放時長（秒）
        /// </summary>
        public double TotalPlayTimeSeconds { get; set; } = 0;

        /// <summary>
        /// 最後播放時間
        /// </summary>
        public DateTime? LastPlayTime { get; set; } = null;

        /// <summary>
        /// 本次開始時間
        /// </summary>
        public DateTime? CurrentSessionStart { get; set; } = null;

        /// <summary>
        /// 本次播放的循環次數
        /// </summary>
        public int CurrentLoopCount { get; set; } = 0;

        /// <summary>
        /// 自定義按鍵觸發次數 (15個槽位)
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
                : "從未播放";

            TimeSpan totalTime = TimeSpan.FromSeconds(TotalPlayTimeSeconds);
            string totalTimeStr = $"{(int)totalTime.TotalHours:D2}:{totalTime.Minutes:D2}:{totalTime.Seconds:D2}";

            return $"總播放次數: {TotalPlayCount}\n" +
                   $"總播放時長: {totalTimeStr}\n" +
                   $"最後播放: {lastPlay}";
        }

        /// <summary>
        /// 取得即時統計資訊（包含當前會話的即時資料）
        /// </summary>
        public string GetLiveFormattedStats()
        {
            double liveTotalSeconds = TotalPlayTimeSeconds;
            string sessionElapsed = "--:--:--";
            bool isActive = CurrentSessionStart.HasValue;

            if (isActive)
            {
                double currentSessionSeconds = (DateTime.Now - CurrentSessionStart.Value).TotalSeconds;
                liveTotalSeconds += currentSessionSeconds;
                TimeSpan sessionTime = TimeSpan.FromSeconds(currentSessionSeconds);
                sessionElapsed = $"{(int)sessionTime.TotalHours:D2}:{sessionTime.Minutes:D2}:{sessionTime.Seconds:D2}";
            }

            string lastPlay = LastPlayTime.HasValue
                ? LastPlayTime.Value.ToString("yyyy-MM-dd HH:mm:ss")
                : "從未播放";

            TimeSpan totalTime = TimeSpan.FromSeconds(liveTotalSeconds);
            string totalTimeStr = $"{(int)totalTime.TotalHours:D2}:{totalTime.Minutes:D2}:{totalTime.Seconds:D2}";

            string status = isActive ? "?? 播放中" : "?? 已停止";

            return $"狀態: {status}\n" +
                   $"當前會話時長: {sessionElapsed}\n" +
                   $"當前循環: {CurrentLoopCount}\n" +
                   $"─────────────────\n" +
                   $"累計播放次數: {TotalPlayCount + (isActive ? 1 : 0)}\n" +
                   $"累計播放時長: {totalTimeStr}\n" +
                   $"最後播放: {lastPlay}";
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
