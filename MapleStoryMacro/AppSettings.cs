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
        /// 配置檔案版本（用於未來遷移）
        /// </summary>
        public int Version { get; set; } = 2;

        /// <summary>
        /// 播放熱鍵
        /// </summary>
        public Keys PlayHotkey { get; set; } = Keys.F9;

        /// <summary>
        /// 停止熱鍵
        /// </summary>
        public Keys StopHotkey { get; set; } = Keys.F10;

        /// <summary>
        /// 暫停/繼續熱鍵
        /// </summary>
        public Keys PauseHotkey { get; set; } = Keys.F11;

        /// <summary>
        /// 錄製熱鍵
        /// </summary>
        public Keys RecordHotkey { get; set; } = Keys.F8;

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
        /// 是否啟用座標修正顯示
        /// </summary>
        public bool PositionCorrectionEnabled { get; set; } = false;

        /// <summary>
        /// 最後載入的腳本路徑（方便下次自動載入）
        /// </summary>
        public string? LastScriptPath { get; set; }

        /// <summary>
        /// 排程任務清單
        /// </summary>
        public List<ScheduleTask> ScheduleTasks { get; set; } = new List<ScheduleTask>();

        /// <summary>
        /// 座標配置（持久化指針路徑）
        /// </summary>
        public CoordinateConfig? CoordinateConfig { get; set; }

        /// <summary>
        /// 是否啟用詳細日誌記錄
        /// </summary>
        public bool EnableDetailedLogging { get; set; } = false;

        /// <summary>
        /// 日誌保留天數
        /// </summary>
        public int LogRetentionDays { get; set; } = 7;

        /// <summary>
        /// 位置修正設定
        /// </summary>
        public PositionCorrectionSettings? PositionCorrection { get; set; }
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
    /// 自定義按鍵槽位資料 (可序列化) - 改進版本
    /// 支援跨電腦、跨鍵盤布局的按鍵儲存
    /// </summary>
    public class CustomKeySlotData
    {
        public int SlotNumber { get; set; }
        
        // ===== 向後兼容：保留舊欄位 =====
        public int KeyCode { get; set; } = (int)Keys.None;
        public int Modifiers { get; set; } = (int)Keys.None;
        
        public double IntervalSeconds { get; set; } = 30.0;
        public bool Enabled { get; set; } = false;
        public double StartAtSecond { get; set; } = 0;
        public double PreDelaySeconds { get; set; } = 0;
        public double PauseScriptSeconds { get; set; } = 0;
        public bool PauseScriptEnabled { get; set; } = false;

        // ===== 新增：跨平台支援 =====
        /// <summary>
        /// 硬體掃描碼（Hardware Scan Code）- 跨鍵盤布局一致
        /// </summary>
        public uint ScanCode { get; set; } = 0;
        
        /// <summary>
        /// 修飾鍵的掃描碼
        /// </summary>
        public uint ModifiersScanCode { get; set; } = 0;
        
        /// <summary>
        /// 按鍵名稱（用於 UI 顯示和調試）
        /// </summary>
        public string KeyName { get; set; } = string.Empty;
        
        /// <summary>
        /// 修飾鍵名稱
        /// </summary>
        public string ModifiersName { get; set; } = string.Empty;
        
        /// <summary>
        /// 是否為擴展鍵（Extended Key，如方向鍵、Home、End 等）
        /// </summary>
        public bool IsExtendedKey { get; set; } = false;
        
        /// <summary>
        /// 記錄時的鍵盤布局 ID（用於調試）
        /// </summary>
        public int KeyboardLayoutId { get; set; } = 0;

        /// <summary>
        /// 檢查是否有有效的按鍵設定
        /// 優先使用 ScanCode，回退到 KeyCode
        /// </summary>
        public bool HasValidKey()
        {
            return ScanCode != 0 || KeyCode != (int)Keys.None;
        }
    }

    /// <summary>
    /// 座標配置 - 持久化指針路徑
    /// </summary>
    public class CoordinateConfig
    {
        /// <summary>
        /// X 座標指針路徑
        /// </summary>
        public PointerPath? XCoordinatePath { get; set; }
        
        /// <summary>
        /// Y 座標指針路徑
        /// </summary>
        public PointerPath? YCoordinatePath { get; set; }
        
        /// <summary>
        /// 遊戲版本（用於檢測更新）
        /// </summary>
        public string GameVersion { get; set; } = string.Empty;
        
        /// <summary>
        /// 最後驗證時間
        /// </summary>
        public DateTime LastVerified { get; set; }
        
        /// <summary>
        /// 是否自動驗證
        /// </summary>
        public bool AutoVerify { get; set; } = true;
        
        /// <summary>
        /// 驗證間隔（小時）
        /// </summary>
        public int VerifyIntervalHours { get; set; } = 24;
    }

    /// <summary>
    /// 指針鏈路徑 - 用於記憶體掃描
    /// 例如：[[MapleStory.exe+0x123456] + 0x10] + 0x4
    /// </summary>
    public class PointerPath
    {
        /// <summary>
        /// 模組名稱（如 MapleStory.exe）
        /// </summary>
        public string ModuleName { get; set; } = "MapleStory.exe";
        
        /// <summary>
        /// 基址偏移量
        /// </summary>
        public long BaseOffset { get; set; }
        
        /// <summary>
        /// 偏移量鏈（每層指針的偏移）
        /// </summary>
        public List<int> Offsets { get; set; } = new List<int>();
        
        /// <summary>
        /// 可讀的路徑描述
        /// </summary>
        public override string ToString()
        {
            string path = $"[{ModuleName}+0x{BaseOffset:X}]";
            foreach (var offset in Offsets)
            {
                path = $"[{path}+0x{offset:X}]";
            }
            return path;
        }
    }
}
