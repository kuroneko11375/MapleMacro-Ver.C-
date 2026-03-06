using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    public class PositionCorrector
    {
        #region P/Invoke
        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string? lpszClass, string? lpszWindow);
        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);
        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP = 0x0101;
        #endregion

        #region 設定
        public int Tolerance { get; set; } = 5;
        public int MaxCorrectionTimeMs { get; set; } = 5000;
        public int DetectionIntervalMs { get; set; } = 200;
        public int KeyPressDurationMs { get; set; } = 60;
        public Keys[] MoveLeftKeys { get; set; } = new[] { Keys.Left };
        public Keys[] MoveRightKeys { get; set; } = new[] { Keys.Right };
        public Keys[] MoveUpKeys { get; set; } = new[] { Keys.Up };
        public Keys[] MoveDownKeys { get; set; } = new[] { Keys.Down };
        public bool EnableHorizontalCorrection { get; set; } = true;
        public bool EnableVerticalCorrection { get; set; } = false;
        public bool InvertY { get; set; } = false;
        public IntPtr TargetWindow { get; set; } = IntPtr.Zero;
        public int MaxWrongDirectionCount { get; set; } = 3;

        /// <summary>
        /// 外部注入的按鍵發送方法 (hWnd, key, isKeyDown)
        /// </summary>
        public Action<IntPtr, Keys, bool>? ExternalKeySender { get; set; }

        /// <summary>
        /// 外部注入：檢查是否處於技能硬直（Form1.IsAnimationLocked）
        /// 硬直期間修正器暫停發送任何移動按鍵
        /// </summary>
        public Func<bool>? IsAnimationLockedCheck { get; set; }

        /// <summary>
        /// ★ 外部注入：ROI 爬繩狀態偵測。
        /// 返回 true 代表角色正在繩子上，水平修正應改用「方向鍵+跳躍」跳離繩子。
        /// </summary>
        public Func<bool>? IsClimbingCheck { get; set; }

        /// <summary>
        /// ★ 爬繩逃脫用跳躍鍵組合（預設 Keys.Alt，楓之谷預設跳躍鍵）
        /// </summary>
        public Keys[] ClimbEscapeJumpKeys { get; set; } = new[] { Keys.Alt };

        /// <summary>水平修正距離閾值：小於等於此值使用長按走路，大於此值使用設定熱鍵</summary>
        public int HorizontalHoldWalkThreshold { get; set; } = 30;

        /// <summary>每次修正呼叫內最多按鍵嘗試次數（0=無限制，僅受時間限制）</summary>
        public int MaxStepsPerCorrection { get; set; } = 0;
        /// <summary>每次循環內最多觸發修正次數（0=無限制）</summary>
        public int MaxCorrectionsPerLoop { get; set; } = 0;
        /// <summary>修正按鍵間隔下限（毫秒）— 向後兼容</summary>
        public int KeyIntervalMinMs { get; set; } = 700;
        /// <summary>修正按鍵間隔上限（毫秒）— 向後兼容</summary>
        public int KeyIntervalMaxMs { get; set; } = 1200;
        /// <summary>水平修正按鍵間隔下限（毫秒）</summary>
        public int HorizontalKeyIntervalMinMs { get; set; } = 700;
        /// <summary>水平修正按鍵間隔上限（毫秒）</summary>
        public int HorizontalKeyIntervalMaxMs { get; set; } = 1200;
        /// <summary>垂直修正按鍵間隔下限（毫秒）— 垂直通常需要較短間隔</summary>
        public int VerticalKeyIntervalMinMs { get; set; } = 300;
        /// <summary>垂直修正按鍵間隔上限（毫秒）</summary>
        public int VerticalKeyIntervalMaxMs { get; set; } = 600;
        /// <summary>柔性容差下限（修正到此範圍內即停止）</summary>
        public int SoftToleranceMin { get; set; } = 5;
        /// <summary>柔性容差上限</summary>
        public int SoftToleranceMax { get; set; } = 8;
        /// <summary>水平容差 (X軸，通常較寬鬆)</summary>
        public int HorizontalTolerance { get; set; } = 8;
        /// <summary>垂直容差 (Y軸，通常較精準)</summary>
        public int VerticalTolerance { get; set; } = 5;
        /// <summary>連續跳躍次數（向上修正時連跳幾次，1=單跳，2+=連跳）</summary>
        public int ConsecutiveJumpCount { get; set; } = 1;
        /// <summary>連續跳躍間隔（毫秒）— 按鍵到按鍵的間距（press-to-press）</summary>
        public int ConsecutiveJumpIntervalMs { get; set; } = 80;
        /// <summary>觸發連跳的Y軸偏差閾值（超過此值才連跳，0=總是使用連跳次數）</summary>
        public int ConsecutiveJumpThreshold { get; set; } = 0;
        #endregion

        // 向後兼容
        public Keys MoveLeftKey { set => MoveLeftKeys = new[] { value }; }
        public Keys MoveRightKey { set => MoveRightKeys = new[] { value }; }
        public Keys MoveUpKey { set => MoveUpKeys = new[] { value }; }
        public Keys MoveDownKey { set => MoveDownKeys = new[] { value }; }

        #region 狀態
        public bool LastCorrectionSuccess { get; private set; }
        public string LastCorrectionMessage { get; private set; } = "";
        public bool IsCorreecting { get; private set; }
        private volatile bool _cancelRequested;
        private int _referenceX = -1, _referenceY = -1;
        private bool _referenceSet = false;
        private int _correctionCountInLoop = 0;
        private readonly Random _rng = new Random();
        /// <summary>上一次修正實際使用的水平容差值</summary>
            public int LastHorizontalTolerance { get; private set; }
            /// <summary>上一次修正實際使用的垂直容差值</summary>
            public int LastVerticalTolerance { get; private set; }
            /// <summary>上一次修正實際使用的柔性容差值（向後兼容，取最大值）</summary>
            public int LastSoftTolerance => Math.Max(LastHorizontalTolerance, LastVerticalTolerance);
        #endregion

        #region 偏差歷史記錄
        private readonly List<CorrectionRecord> _history = new List<CorrectionRecord>();
        private const int MaxHistorySize = 100;

        private class CorrectionRecord
        {
            public int DeltaX, DeltaY, Steps;
            public bool Success;
        }

        public double AverageCorrectionSteps => _history.Count > 0
            ? _history.Where(r => r.Success).Select(r => r.Steps).DefaultIfEmpty(0).Average() : 0;
        public int HistoryCount => _history.Count;

        public string GetHistorySummary()
        {
            if (_history.Count == 0) return "無歷史記錄";
            int ok = _history.Count(r => r.Success);
            return $"記錄:{_history.Count}/{MaxHistorySize} 成功率:{ok * 100 / _history.Count}% 平均步數:{AverageCorrectionSteps:F1}";
        }

        private void AddHistory(int dx, int dy, int steps, bool success)
        {
            _history.Add(new CorrectionRecord { DeltaX = dx, DeltaY = dy, Steps = steps, Success = success });
            while (_history.Count > MaxHistorySize) _history.RemoveAt(0);
        }

        public void ClearHistory() => _history.Clear();

        public List<string> ExportHistoryLines() =>
            _history.Select((r, i) => $"[{i + 1}] Δ({r.DeltaX:+#;-#;0},{r.DeltaY:+#;-#;0}) {r.Steps}步 {(r.Success ? "✓" : "✗")}").ToList();
        #endregion

        public event Action<string>? OnLog;
        public event Action<string>? OnStatusUpdate;

        public void SetReferencePosition(int x, int y) { _referenceX = x; _referenceY = y; _referenceSet = true; }
        public (int x, int y, bool isSet) GetReferencePosition() => (_referenceX, _referenceY, _referenceSet);
        public void ClearReferencePosition() { _referenceX = -1; _referenceY = -1; _referenceSet = false; }
        public void ResetLoopCorrectionCount() => _correctionCountInLoop = 0;
        public bool IsLoopCorrectionLimitReached() => MaxCorrectionsPerLoop > 0 && _correctionCountInLoop >= MaxCorrectionsPerLoop;

        /// <summary>
        /// 核心修正邏輯
        /// ★ 分離容差：HorizontalTolerance (X軸) / VerticalTolerance (Y軸)
        /// ★ 按鍵間隔：每次按鍵後等待 KeyIntervalMinMs~MaxMs（避免硬直吃不到指令）
        /// ★ 步數限制：MaxStepsPerCorrection > 0 時，按鍵嘗試次數不超過此值
        /// ★ 時間限制：MaxCorrectionTimeMs 時間到即停止
        /// ★ 防過度修正：進入容差範圍後立即結束，不再嘗試逼近目標中心
        /// </summary>
        public CorrectionResult CorrectPosition(MinimapTracker tracker, int targetX, int targetY)
        {
            if (tracker == null || !tracker.IsCalibrated)
                return new CorrectionResult(false, "小地圖追蹤器未校準");
            if (TargetWindow == IntPtr.Zero)
                return new CorrectionResult(false, "未指定目標視窗");
            if (IsLoopCorrectionLimitReached())
                return new CorrectionResult(false, $"已達循環修正觸發上限({MaxCorrectionsPerLoop}次)");

            IsCorreecting = true;
            _cancelRequested = false;
            _correctionCountInLoop++;
            var sw = Stopwatch.StartNew();

            // ★ 變更 #3：修正期間鎖定偵測頻率為 50ms（結束時恢復）
            int savedDetectionInterval = DetectionIntervalMs;
            DetectionIntervalMs = 50;

            // ★ 使用分離容差：X軸與Y軸各自獨立
            int hTol = HorizontalTolerance;
            int vTol = VerticalTolerance;
            LastHorizontalTolerance = hTol;
            LastVerticalTolerance = vTol;

            // ★ 步數上限：MaxStepsPerCorrection > 0 時使用，否則僅受時間限制
            int maxSteps = MaxStepsPerCorrection > 0
                ? MaxStepsPerCorrection
                : int.MaxValue; // 無步數限制，僅受時間限制
            int keyPressCount = 0; // 實際按鍵次數（不含偵測失敗的輪次）
            int iter = 0;
            int wrongDir = 0;
            int verticalStuckCount = 0; // ★ 變更 #2：垂直向下卡住偵測計數
            int lastVerticalY = int.MinValue; // 追蹤上次垂直修正後的 Y 座標

            var (initX, initY, initOk) = tracker.ReadPosition();
            int initialDx = initOk ? targetX - initX : 0;
            int initialDy = initOk ? targetY - initY : 0;

            // ★ 先行檢查：如果已經在容差範圍內，不需要修正
            if (initOk)
            {
                bool xAlreadyOk = !EnableHorizontalCorrection || Math.Abs(initialDx) <= hTol;
                bool yAlreadyOk = !EnableVerticalCorrection || Math.Abs(initialDy) <= vTol;
                if (xAlreadyOk && yAlreadyOk)
                {
                    var skipMsg = $"已在容差範圍內 ({initX},{initY}) 偏差({initialDx},{initialDy}) 容差H±{hTol}/V±{vTol}，跳過修正";
                    Log("✅ " + skipMsg);
                    LastCorrectionSuccess = true; LastCorrectionMessage = skipMsg;
                    IsCorreecting = false;
                    return new CorrectionResult(true, skipMsg, initX, initY, 0, sw.ElapsedMilliseconds);
                }
            }

            string stepsInfo = MaxStepsPerCorrection > 0 ? $"步數上限{MaxStepsPerCorrection}" : "無步數限制";
            Log($"=== 修正開始 目標({targetX},{targetY}) 初始({initX},{initY}) 容差H±{hTol}/V±{vTol} {stepsInfo} 時限{MaxCorrectionTimeMs}ms H間隔{HorizontalKeyIntervalMinMs}~{HorizontalKeyIntervalMaxMs}ms V間隔{VerticalKeyIntervalMinMs}~{VerticalKeyIntervalMaxMs}ms ===");

            try
            {
                while (!_cancelRequested
                    && sw.ElapsedMilliseconds < MaxCorrectionTimeMs
                    && keyPressCount < maxSteps)
                {
                    iter++;

                    // ★ 變更 #4：技能硬直鎖 — 施放自定義技能期間暫停修正
                    if (IsAnimationLockedCheck?.Invoke() == true)
                    {
                        Log($"[{iter}] ⏸️ 技能硬直中，等待解鎖...");
                        if (!WaitForAnimationUnlock(sw))
                        {
                            Log($"[{iter}] ⏰ 等待硬直超時");
                            break;
                        }
                    }

                    // ★ 每次迭代開始前釋放所有修正按鍵，防止殘留 key-down 影響方向判斷
                    if (iter > 1)
                    {
                        ReleaseAllCorrectionKeys();
                        Thread.Sleep(15);
                    }

                    var (cx, cy, ok) = tracker.ReadPosition();
                    if (!ok) { Log($"[{iter}] 偵測失敗"); Thread.Sleep(200); continue; }

                    int dx = targetX - cx, dy = targetY - cy;
                    bool xOk = !EnableHorizontalCorrection || Math.Abs(dx) <= hTol;
                    bool yOk = !EnableVerticalCorrection || Math.Abs(dy) <= vTol;

                    Log($"[{iter}] 在({cx},{cy}) 差({dx:+#;-#;0},{dy:+#;-#;0}) 容差H±{hTol}/V±{vTol} 步數{keyPressCount}/{(MaxStepsPerCorrection > 0 ? MaxStepsPerCorrection.ToString() : "∞")} {(xOk ? "X✓" : "X✗")} {(yOk ? "Y✓" : "Y✗")}");

                    // ★ 進入容差範圍就立即結束，不再追求更精準
                    // ★ 變更 #2：啟用垂直修正時，需通過穩定判定
                    if (xOk && yOk)
                    {
                        bool vertStable = true;
                        if (EnableVerticalCorrection && keyPressCount > 0)
                        {
                            vertStable = VerticalStableCheck(tracker, targetY, vTol);
                            if (!vertStable)
                            {
                                Log($"[{iter}] ⚠️ 穩定確認失敗，繼續修正");
                                continue;
                            }
                        }
                        var msg = $"修正完成！({cx},{cy}) 偏差({dx:+#;-#;0},{dy:+#;-#;0}) 容差H±{hTol}/V±{vTol} 按鍵{keyPressCount}次 {sw.ElapsedMilliseconds}ms{(EnableVerticalCorrection && keyPressCount > 0 ? " (穩定確認✓)" : "")}";
                        Log("✅ " + msg);
                        LastCorrectionSuccess = true; LastCorrectionMessage = msg;
                        AddHistory(initialDx, initialDy, keyPressCount, true);
                        return new CorrectionResult(true, msg, cx, cy, keyPressCount, sw.ElapsedMilliseconds);
                    }

                    bool needVertical = EnableVerticalCorrection && !yOk;
                    bool needHorizontal = EnableHorizontalCorrection && !xOk;
                    bool xApproxOk = !EnableHorizontalCorrection || Math.Abs(dx) <= hTol * 2;

                    // ★ 爬繩逃脫：在任何方向修正前先檢查是否在繩上
                    // 繩上無法正常走路或跳躍，必須先用方向+跳躍跳離繩子
                    if ((needVertical || needHorizontal) && IsClimbingCheck?.Invoke() == true)
                    {
                        Keys escapeDir = dx >= 0 ? Keys.Right : Keys.Left;
                        var escapeKeys = new[] { escapeDir }.Concat(ClimbEscapeJumpKeys).ToArray();
                        Log($"[{iter}] 🪢 爬繩狀態！先跳離繩子 [{string.Join("+", escapeKeys.Select(k => k.ToString()))}]");
                        SendCombinedKeyPress(escapeKeys);
                        keyPressCount++;
                        Thread.Sleep(200);
                        SleepKeyInterval(false);
                        continue;
                    }

                    if (needVertical && xApproxOk)
                    {
                        bool needGoUp = InvertY ? (dy > 0) : (dy < 0);
                        var jumpKeys = needGoUp ? MoveUpKeys : MoveDownKeys;
                        string dir = needGoUp ? "↑" : "↓";

                        Keys? rawDirKey = null;
                        // ★ 往下修正時不夾帶左右方向鍵，只有往上時才搭配水平移動
                        if (EnableHorizontalCorrection && Math.Abs(dx) > 1 && needGoUp)
                        {
                            rawDirKey = dx > 0 ? Keys.Right : Keys.Left;
                            dir += dx > 0 ? "→" : "←";
                        }

                        // ★ 變更 #2：向下卡住偵測 — 連續 3 次 Y 不變則嘗試左右逃離
                        if (!needGoUp && verticalStuckCount >= 3)
                        {
                            Keys escapeKey = _rng.Next(2) == 0 ? Keys.Left : Keys.Right;
                            Log($"[{iter}] ⚠️ 向下卡住 {verticalStuckCount} 次，嘗試 {escapeKey} 逃離");
                            SendCombinedKeyPress(new[] { escapeKey });
                            keyPressCount++;
                            Thread.Sleep(200);
                            verticalStuckCount = 0;
                            lastVerticalY = int.MinValue;
                            continue; // 重新偵測
                        }

                        int jumpCount = 1;
                        if (needGoUp && ConsecutiveJumpCount > 1)
                        {
                            bool shouldConsecutive = ConsecutiveJumpThreshold <= 0 || Math.Abs(dy) > ConsecutiveJumpThreshold;
                            if (shouldConsecutive) jumpCount = ConsecutiveJumpCount;
                        }
                        Log($"[{iter}] 垂直修正 {dir} 跳躍=[{string.Join("+", jumpKeys.Select(k => k.ToString()))}]{(rawDirKey.HasValue ? $" 方向={rawDirKey}" : "")}{(jumpCount > 1 ? $" x{jumpCount}連跳 間隔{ConsecutiveJumpIntervalMs}ms" : "")}");

                        // ★ 使用專用連跳方法：方向鍵全程按住，跳躍鍵快速連拍
                        SendConsecutiveJumpSequence(jumpKeys, rawDirKey, jumpCount, ConsecutiveJumpIntervalMs);
                        keyPressCount += jumpCount;

                        // ★ 連跳後不用 SleepKeyInterval，只短等讓角色落地/移動生效
                        Thread.Sleep(jumpCount > 1 ? 80 : 50);

                        // 按鍵後立即檢查
                        var (vx, vy, vOk) = tracker.ReadPosition();
                        if (vOk)
                        {
                            // ★ 變更 #2：向下卡住偵測 — 追蹤 Y 是否有變化
                            if (!needGoUp)
                            {
                                if (lastVerticalY != int.MinValue && Math.Abs(vy - lastVerticalY) <= 1)
                                    verticalStuckCount++;
                                else
                                    verticalStuckCount = 0;
                                lastVerticalY = vy;
                            }

                            int vdx = targetX - vx, vdy = targetY - vy;
                            bool vxOk = !EnableHorizontalCorrection || Math.Abs(vdx) <= hTol;
                            bool vyOk = !EnableVerticalCorrection || Math.Abs(vdy) <= vTol;

                            // ★ 變更 #2：穩定判定 — Y 到位後等 100ms 再確認一次
                            if (vxOk && vyOk)
                            {
                                bool stable = VerticalStableCheck(tracker, targetY, vTol);
                                if (stable)
                                {
                                    // 重新讀取最終座標
                                    var (fx2, fy2, fOk2) = tracker.ReadPosition();
                                    int fdx = fOk2 ? targetX - fx2 : vdx;
                                    int fdy = fOk2 ? targetY - fy2 : vdy;
                                    int finalX = fOk2 ? fx2 : vx;
                                    int finalY = fOk2 ? fy2 : vy;
                                    var msg = $"修正完成！({finalX},{finalY}) 偏差({fdx:+#;-#;0},{fdy:+#;-#;0}) 容差H±{hTol}/V±{vTol} 按鍵{keyPressCount}次 {sw.ElapsedMilliseconds}ms (穩定確認✓)";
                                    Log("✅ " + msg);
                                    LastCorrectionSuccess = true; LastCorrectionMessage = msg;
                                    AddHistory(initialDx, initialDy, keyPressCount, true);
                                    verticalStuckCount = 0;
                                    return new CorrectionResult(true, msg, finalX, finalY, keyPressCount, sw.ElapsedMilliseconds);
                                }
                                else
                                {
                                    Log($"[{iter}] ⚠️ 穩定確認失敗（100ms 後 Y 偏移），繼續修正");
                                }
                            }
                        }
                    }
                    else if (needHorizontal)
                    {
                        int absDx = Math.Abs(dx);
                        string dir;
                        Keys[] keys;
                        int expectDx;
                        if (dx > 0) { keys = MoveRightKeys; dir = "→"; expectDx = 1; }
                        else { keys = MoveLeftKeys; dir = "←"; expectDx = -1; }

                        // ★ 變更 #1：水平修正依距離分兩種模式
                        if (absDx <= HorizontalHoldWalkThreshold)
                        {
                            // --- 短～中距離：長按原生方向鍵 + 50ms 高頻偵測，到位立即放開 ---
                            Keys rawArrow = dx > 0 ? Keys.Right : Keys.Left;
                            Log($"[{iter}] 水平修正(走路) {dir} 距離{absDx}px 長按{rawArrow}");
                            keyPressCount++;
                            bool reached = SendHeldArrowWithPolling(tracker, rawArrow, targetX, hTol, sw);

                            var (nx, ny, nOk) = tracker.ReadPosition();
                            if (nOk)
                            {
                                int ndx = targetX - nx, ndy = targetY - ny;
                                bool nxOk = !EnableHorizontalCorrection || Math.Abs(ndx) <= hTol;
                                bool nyOk = !EnableVerticalCorrection || Math.Abs(ndy) <= vTol;
                                Log($"[{iter}] 走路結果: ({cx},{cy})→({nx},{ny}) 剩餘偏差({ndx:+#;-#;0},{ndy:+#;-#;0}){(reached ? " ✓到位" : "")}");
                                if (nxOk && nyOk)
                                {
                                    var msg = $"修正完成！({nx},{ny}) 偏差({ndx:+#;-#;0},{ndy:+#;-#;0}) 容差H±{hTol}/V±{vTol} 按鍵{keyPressCount}次 {sw.ElapsedMilliseconds}ms";
                                    Log("✅ " + msg);
                                    LastCorrectionSuccess = true; LastCorrectionMessage = msg;
                                    AddHistory(initialDx, initialDy, keyPressCount, true);
                                    return new CorrectionResult(true, msg, nx, ny, keyPressCount, sw.ElapsedMilliseconds);
                                }
                            }
                        }
                        else
                        {
                            // --- 長距離 (>30px)：使用設定的熱鍵（可能是傳送/衝刺） + 冷卻 ---
                            Log($"[{iter}] 水平修正(熱鍵) {dir} 距離{absDx}px [{string.Join("+", keys.Select(k => k.ToString()))}]");
                            SendCombinedKeyPress(keys);
                            keyPressCount++;
                            SleepKeyInterval(false);

                            var (nx, ny, nOk) = tracker.ReadPosition();
                            if (nOk)
                            {
                                int mx = nx - cx;
                                bool isWrong = expectDx != 0 && mx != 0 &&
                                    ((expectDx > 0 && mx < -1) || (expectDx < 0 && mx > 1));

                                Log($"[{iter}] 結果: ({cx},{cy})→({nx},{ny}) 移動({mx:+#;-#;0},{ny - cy:+#;-#;0}){(isWrong ? " ⚠️方向反了!" : "")}");

                                if (isWrong)
                                {
                                    wrongDir++;
                                    if (wrongDir >= MaxWrongDirectionCount)
                                    {
                                        var msg = $"連續{wrongDir}次方向錯誤！按鍵{keyPressCount}次。請用「方向診斷」確認按鍵。";
                                        Log("❌ " + msg);
                                        LastCorrectionSuccess = false; LastCorrectionMessage = msg;
                                        AddHistory(initialDx, initialDy, keyPressCount, false);
                                        return new CorrectionResult(false, msg, nx, ny, keyPressCount, sw.ElapsedMilliseconds);
                                    }
                                }
                                else wrongDir = 0;

                                // 按鍵後立即檢查
                                int ndx = targetX - nx, ndy = targetY - ny;
                                bool nxOk = !EnableHorizontalCorrection || Math.Abs(ndx) <= hTol;
                                bool nyOk = !EnableVerticalCorrection || Math.Abs(ndy) <= vTol;
                                if (nxOk && nyOk)
                                {
                                    var msg = $"修正完成！({nx},{ny}) 偏差({ndx:+#;-#;0},{ndy:+#;-#;0}) 容差H±{hTol}/V±{vTol} 按鍵{keyPressCount}次 {sw.ElapsedMilliseconds}ms";
                                    Log("✅ " + msg);
                                    LastCorrectionSuccess = true; LastCorrectionMessage = msg;
                                    AddHistory(initialDx, initialDy, keyPressCount, true);
                                    return new CorrectionResult(true, msg, nx, ny, keyPressCount, sw.ElapsedMilliseconds);
                                }
                            }
                        }
                        continue;
                    }
                    else if (needVertical && !xApproxOk)
                    {
                        Keys[] keys = dx > 0 ? MoveRightKeys : MoveLeftKeys;
                        Log($"[{iter}] X偏差大 先水平修正 {(dx > 0 ? "→" : "←")}");
                        SendCombinedKeyPress(keys);
                        keyPressCount++;
                        SleepKeyInterval(false);

                        var (hx, hy, hOk) = tracker.ReadPosition();
                        if (hOk)
                        {
                            int hdx = targetX - hx, hdy = targetY - hy;
                            bool hxOk = !EnableHorizontalCorrection || Math.Abs(hdx) <= hTol;
                            bool hyOk = !EnableVerticalCorrection || Math.Abs(hdy) <= vTol;
                            if (hxOk && hyOk)
                            {
                                var msg = $"修正完成！({hx},{hy}) 偏差({hdx:+#;-#;0},{hdy:+#;-#;0}) 容差H±{hTol}/V±{vTol} 按鍵{keyPressCount}次 {sw.ElapsedMilliseconds}ms";
                                Log("✅ " + msg);
                                LastCorrectionSuccess = true; LastCorrectionMessage = msg;
                                AddHistory(initialDx, initialDy, keyPressCount, true);
                                return new CorrectionResult(true, msg, hx, hy, keyPressCount, sw.ElapsedMilliseconds);
                            }
                        }
                        continue;
                    }

                    SleepKeyInterval(false);
                }

                // ★ 判斷結束原因
                var (fX, fY, fOk) = tracker.ReadPosition();
                string reason;
                if (_cancelRequested)
                    reason = "已取消";
                else if (keyPressCount >= maxSteps)
                    reason = $"已達步數上限({maxSteps}次)";
                else
                    reason = $"已達時間上限({MaxCorrectionTimeMs}ms)";

                var tmsg = fOk
                    ? $"{reason} 最終({fX},{fY}) 目標({targetX},{targetY}) 按鍵{keyPressCount}次 {sw.ElapsedMilliseconds}ms"
                    : $"{reason} 無法偵測最終位置 按鍵{keyPressCount}次 {sw.ElapsedMilliseconds}ms";
                Log("⏰ " + tmsg);
                LastCorrectionSuccess = false; LastCorrectionMessage = tmsg;
                AddHistory(initialDx, initialDy, keyPressCount, false);
                return new CorrectionResult(false, tmsg, fX, fY, keyPressCount, sw.ElapsedMilliseconds);
            }
            finally
            {
                // ★ 安全釋放所有修正按鍵，防止卡 keydown 導致角色一直移動
                ReleaseAllCorrectionKeys();
                // ★ 變更 #3：恢復偵測頻率
                DetectionIntervalMs = savedDetectionInterval;
                IsCorreecting = false;
            }
        }

        /// <summary>
        /// ★ 隨機按鍵間隔，水平與垂直分開（垂直較短以加速跳躍修正）
        /// </summary>
        private void SleepKeyInterval(bool isVertical = false)
        {
            int min, max;
            if (isVertical)
            {
                min = Math.Max(VerticalKeyIntervalMinMs, 50);
                max = Math.Max(VerticalKeyIntervalMaxMs, min + 1);
            }
            else
            {
                min = Math.Max(HorizontalKeyIntervalMinMs, 100);
                max = Math.Max(HorizontalKeyIntervalMaxMs, min + 1);
            }
            Thread.Sleep(_rng.Next(min, max));
        }

        public CorrectionResult CorrectDeviation(MinimapTracker tracker)
        {
            if (tracker == null || !tracker.IsCalibrated)
                return new CorrectionResult(false, "未校準");
            var (cx, cy, ok) = tracker.ReadPosition();
            if (!ok) return new CorrectionResult(false, "無法偵測");
            if (!_referenceSet)
            {
                _referenceX = cx; _referenceY = cy; _referenceSet = true;
                var msg = $"擷取參考位置: ({cx},{cy})";
                Log(msg);
                return new CorrectionResult(true, msg, cx, cy, 0, 0);
            }
            return CorrectPosition(tracker, _referenceX, _referenceY);
        }

        public DiagnosticResult DiagnoseDirection(MinimapTracker tracker, string direction, Keys[] keys)
        {
            if (tracker == null || !tracker.IsCalibrated)
                return new DiagnosticResult { Error = "未校準" };
            if (TargetWindow == IntPtr.Zero)
                return new DiagnosticResult { Error = "未鎖定視窗" };

            var (bx, by, ok1) = tracker.ReadPosition();
            if (!ok1) return new DiagnosticResult { Error = "偵測失敗" };

            try
            {
                SendCombinedKeyPress(keys);
                Thread.Sleep(350);
            }
            finally
            {
                // ★ 安全釋放，防止診斷時按鍵卡住
                ReleaseAllCorrectionKeys();
            }

            var (ax, ay, ok2) = tracker.ReadPosition();
            if (!ok2) return new DiagnosticResult { Error = "按後偵測失敗" };

            return new DiagnosticResult
            {
                Direction = direction,
                KeyNames = string.Join("+", keys.Select(k => k.ToString())),
                BeforeX = bx, BeforeY = by, AfterX = ax, AfterY = ay,
                DeltaX = ax - bx, DeltaY = ay - by
            };
        }

        public (int, int, bool, bool) CheckDrift(MinimapTracker t, int ex, int ey)
        {
            if (t == null || !t.IsCalibrated) return (0, 0, false, false);
            var (cx, cy, ok) = t.ReadPosition();
            if (!ok) return (0, 0, false, false);
            int dx = ex - cx, dy = ey - cy;
            return (dx, dy, (EnableHorizontalCorrection && Math.Abs(dx) > HorizontalTolerance) || (EnableVerticalCorrection && Math.Abs(dy) > VerticalTolerance), true);
        }

        public void Cancel() => _cancelRequested = true;

        #region 按鍵
        /// <summary>
        /// ★ 等待技能硬直結束（最多等 maxWaitMs 毫秒）
        /// </summary>
        private bool WaitForAnimationUnlock(Stopwatch sw, int maxWaitMs = 3000)
        {
            if (IsAnimationLockedCheck == null) return true;
            long waitStart = sw.ElapsedMilliseconds;
            while (IsAnimationLockedCheck.Invoke() && !_cancelRequested)
            {
                if (sw.ElapsedMilliseconds - waitStart > maxWaitMs) return false;
                if (sw.ElapsedMilliseconds >= MaxCorrectionTimeMs) return false;
                Thread.Sleep(50);
            }
            return !_cancelRequested;
        }

        /// <summary>
        /// ★ 水平短距離修正：長按原生方向鍵 + 50ms 高頻偵測，到位立即放開
        /// 不使用使用者設定的快捷鍵，改用 Keys.Left/Right 純走路
        /// </summary>
        private bool SendHeldArrowWithPolling(MinimapTracker tracker, Keys rawArrowKey, int targetX, int hTol, Stopwatch sw)
        {
            IntPtr child = FindWindowEx(TargetWindow, IntPtr.Zero, null, null);
            IntPtr wnd = child != IntPtr.Zero ? child : TargetWindow;

            // 使用 ExternalKeySender 按下（保留 ATT 背景支援）
            if (ExternalKeySender != null)
                ExternalKeySender(TargetWindow, rawArrowKey, true);
            else
                PostMessage(wnd, WM_KEYDOWN, (IntPtr)rawArrowKey, GetLParam(rawArrowKey, false));

            bool reached = false;
            try
            {
                // 50ms 高頻偵測迴圈
                while (!_cancelRequested && sw.ElapsedMilliseconds < MaxCorrectionTimeMs)
                {
                    Thread.Sleep(50);

                    // 技能硬直中 → 暫時不判斷，繼續等
                    if (IsAnimationLockedCheck?.Invoke() == true) continue;

                    var (cx, cy, ok) = tracker.ReadPosition();
                    if (!ok) continue;

                    int dx = targetX - cx;
                    if (Math.Abs(dx) <= hTol)
                    {
                        reached = true;
                        break;
                    }

                    // 防走過頭：如果方向反了就立即停
                    bool goingRight = rawArrowKey == Keys.Right;
                    if ((goingRight && dx < 0) || (!goingRight && dx > 0))
                    {
                        reached = Math.Abs(dx) <= hTol;
                        break;
                    }
                }
            }
            finally
            {
                // ★ 立即放開
                if (ExternalKeySender != null)
                {
                    try { ExternalKeySender(TargetWindow, rawArrowKey, false); } catch { }
                }
                // PostMessage WM_KEYUP 補強
                PostMessage(wnd, WM_KEYUP, (IntPtr)rawArrowKey, GetLParam(rawArrowKey, true));
                Thread.Sleep(20);
            }
            return reached;
        }

        /// <summary>
        /// ★ 垂直穩定判定：達標後等 100ms 再確認一次，連續兩次在容差內才算穩定
        /// </summary>
        private bool VerticalStableCheck(MinimapTracker tracker, int targetY, int vTol)
        {
            Thread.Sleep(100);
            var (cx2, cy2, ok2) = tracker.ReadPosition();
            if (!ok2) return false;
            return Math.Abs(targetY - cy2) <= vTol;
        }

        /// <summary>
        /// ★ 連續跳躍專用：方向鍵全程按住，跳躍鍵快速連拍
        /// 
        /// 修正前的時序問題：
        ///   SendCombinedKeyPress(All) → hold 60ms → release ALL → 20ms → sleep 150ms → 再次 press ALL
        ///   = 第1跳 KeyDown 到 第2跳 KeyDown 間隔 230ms，且方向鍵中途被放開
        /// 
        /// 修正後的時序（使用 Stopwatch 高精度計時）：
        ///   Hold(方向鍵) → Tap(跳躍鍵 30ms) → Stopwatch 精確等待 50~80ms → Tap(跳躍鍵) → Release(全部)
        ///   = 間隔就是使用者設定的 ConsecutiveJumpIntervalMs，方向鍵從不中斷，
        ///     每次跳躍間的等待用 Stopwatch 取代不穩定 Thread.Sleep
        /// </summary>
        private void SendConsecutiveJumpSequence(Keys[] jumpKeys, Keys? directionKey, int jumpCount, int intervalMs)
        {
            if (jumpKeys == null || jumpKeys.Length == 0) return;

            const int JUMP_TAP_HOLD_MS = 30; // 跳躍只需短按

            IntPtr child = FindWindowEx(TargetWindow, IntPtr.Zero, null, null);
            IntPtr wnd = child != IntPtr.Zero ? child : TargetWindow;

            // 高精度計時：Stopwatch 精確到微秒，補償 OS 計時器 15ms 粒度誤差
            var sw = Stopwatch.StartNew();

            try
            {
                // 1. 方向鍵按下並全程按住
                if (directionKey.HasValue && directionKey.Value != Keys.None)
                {
                    SendSingleKey(directionKey.Value, true, wnd);
                }

                // 2. 連續拍按跳躍鍵（★ press-to-press 精確計時）
                long lastTapStartTick = 0;
                for (int i = 0; i < jumpCount; i++)
                {
                    if (_cancelRequested) break;

                    // 跳之前等間隔（第 1 跳不等）— ★ 改用 press-to-press 計時
                    // 舊版從 release 後開始計時，導致實際間隔 = hold + interval = 180ms，角色已下墜
                    // 新版從上一跳的 press 時刻起算，interval 就是真正的按鍵到按鍵間距
                    if (i > 0)
                    {
                        long intervalTicks = (long)(intervalMs * Stopwatch.Frequency / 1000.0);
                        long targetTick = lastTapStartTick + intervalTicks;
                        long remaining = targetTick - sw.ElapsedTicks;

                        if (remaining > 0)
                        {
                            // 超過 5ms 的部分用 Sleep 釋放 CPU
                            long sleepTicks = remaining - (long)(5 * Stopwatch.Frequency / 1000.0);
                            if (sleepTicks > 0)
                            {
                                int sleepMs = (int)(sleepTicks * 1000 / Stopwatch.Frequency);
                                if (sleepMs > 0) Thread.Sleep(sleepMs);
                            }

                            // 剩餘 5ms 以 Spin-wait 精確等待
                            while (sw.ElapsedTicks < targetTick)
                                System.Threading.Thread.SpinWait(10);
                        }
                    }

                    long tapStart = sw.ElapsedTicks;
                    lastTapStartTick = tapStart;

                    // Tap 跳躍鍵：按下
                    foreach (var k in jumpKeys)
                        SendSingleKey(k, true, wnd);

                    // ★ 精確按住 JUMP_TAP_HOLD_MS
                    long holdTicks = (long)(JUMP_TAP_HOLD_MS * Stopwatch.Frequency / 1000.0);
                    long holdTarget = tapStart + holdTicks;
                    long holdRemain = holdTarget - sw.ElapsedTicks;
                    long holdSleepTicks = holdRemain - (long)(2 * Stopwatch.Frequency / 1000.0);
                    if (holdSleepTicks > 0)
                    {
                        int hSleepMs = (int)(holdSleepTicks * 1000 / Stopwatch.Frequency);
                        if (hSleepMs > 0) Thread.Sleep(hSleepMs);
                    }
                    while (sw.ElapsedTicks < holdTarget)
                        System.Threading.Thread.SpinWait(10);

                    // Tap 跳躍鍵：放開
                    foreach (var k in jumpKeys)
                        SendSingleKey(k, false, wnd);
                }
            }
            finally
            {
                // 3. 全部放開（方向鍵 + 跳躍鍵都確保 up）
                if (directionKey.HasValue && directionKey.Value != Keys.None)
                {
                    SendSingleKey(directionKey.Value, false, wnd);
                }
                foreach (var k in jumpKeys)
                    SendSingleKey(k, false, wnd);

                Thread.Sleep(15); // 短等，讓遊戲處理 key-up
            }
        }

        /// <summary>
        /// 單鍵按下/放開 — 匹配 SendCombinedKeyPress 的路由邏輯：
        ///   ExternalKeySender 存在時：用它發 down/up，key-up 時 PostMessage 補強
        ///   ExternalKeySender 不存在時：純 PostMessage
        /// </summary>
        private void SendSingleKey(Keys key, bool down, IntPtr wnd)
        {
            if (ExternalKeySender != null)
            {
                try { ExternalKeySender(TargetWindow, key, down); } catch { }

                // ★ PostMessage WM_KEYUP 補強（僅 key-up，避免 keydown 雙重觸發）
                if (!down && TargetWindow != IntPtr.Zero)
                {
                    PostMessage(wnd, WM_KEYUP, (IntPtr)key, GetLParam(key, true));
                }
            }
            else
            {
                // 純 PostMessage 模式
                uint msg = down ? WM_KEYDOWN : WM_KEYUP;
                PostMessage(wnd, msg, (IntPtr)key, GetLParam(key, !down));
            }
        }
        /// <summary>
        /// 安全釋放所有修正用按鍵（防止按鍵卡住導致角色一直移動）
        /// </summary>
        private void ReleaseAllCorrectionKeys()
        {
            var allKeys = new HashSet<Keys>();
            if (MoveLeftKeys != null) foreach (var k in MoveLeftKeys) allKeys.Add(k);
            if (MoveRightKeys != null) foreach (var k in MoveRightKeys) allKeys.Add(k);
            if (MoveUpKeys != null) foreach (var k in MoveUpKeys) allKeys.Add(k);
            if (MoveDownKeys != null) foreach (var k in MoveDownKeys) allKeys.Add(k);

            // ★ 雙重釋放策略：先透過 ExternalKeySender（ATT+keybd_event），
            //   再透過 PostMessage 補強（確保遊戲訊息佇列也收到 WM_KEYUP）
            if (ExternalKeySender != null)
            {
                foreach (var k in allKeys)
                {
                    try { ExternalKeySender(TargetWindow, k, false); } catch { }
                }
            }

            // ★ PostMessage 補強釋放（即使 ATT 失敗也能透過訊息佇列釋放 key state）
            if (TargetWindow != IntPtr.Zero)
            {
                IntPtr child = FindWindowEx(TargetWindow, IntPtr.Zero, null, null);
                IntPtr wnd = child != IntPtr.Zero ? child : TargetWindow;
                foreach (var k in allKeys)
                    PostMessage(wnd, WM_KEYUP, (IntPtr)k, GetLParam(k, true));
            }
        }

        private static readonly HashSet<Keys> _arrowKeys = new HashSet<Keys>
            { Keys.Left, Keys.Right, Keys.Up, Keys.Down };

        private void SendCombinedKeyPress(Keys[] keys)
        {
            if (keys == null || keys.Length == 0) return;

            // ★ 分離方向鍵與動作鍵 — 方向鍵必須先按下，遊戲才能正確判定朝向
            var dirKeys = keys.Where(k => _arrowKeys.Contains(k)).ToArray();
            var actKeys = keys.Where(k => !_arrowKeys.Contains(k)).ToArray();

            if (ExternalKeySender != null)
            {
                try
                {
                    // 1. 先按方向鍵（確立朝向）
                    foreach (var k in dirKeys) ExternalKeySender(TargetWindow, k, true);
                    if (dirKeys.Length > 0 && actKeys.Length > 0)
                        Thread.Sleep(15); // 等遊戲註冊方向
                    // 2. 再按動作鍵（跳躍/技能）
                    foreach (var k in actKeys) ExternalKeySender(TargetWindow, k, true);
                    Thread.Sleep(KeyPressDurationMs);
                }
                finally
                {
                    // ★ 反序釋放：先放動作鍵，再放方向鍵
                    for (int i = actKeys.Length - 1; i >= 0; i--)
                        ExternalKeySender(TargetWindow, actKeys[i], false);
                    for (int i = dirKeys.Length - 1; i >= 0; i--)
                        ExternalKeySender(TargetWindow, dirKeys[i], false);

                    // ★ PostMessage WM_KEYUP 補強
                    if (TargetWindow != IntPtr.Zero)
                    {
                        IntPtr child = FindWindowEx(TargetWindow, IntPtr.Zero, null, null);
                        IntPtr wnd = child != IntPtr.Zero ? child : TargetWindow;
                        for (int i = keys.Length - 1; i >= 0; i--)
                            PostMessage(wnd, WM_KEYUP, (IntPtr)keys[i], GetLParam(keys[i], true));
                    }

                    Thread.Sleep(20); // ★ 等待遊戲輪詢到 key-up 狀態
                }
            }
            else
            {
                IntPtr child = FindWindowEx(TargetWindow, IntPtr.Zero, null, null);
                IntPtr wnd = child != IntPtr.Zero ? child : TargetWindow;
                try
                {
                    // 1. 先按方向鍵
                    foreach (var k in dirKeys) PostMessage(wnd, WM_KEYDOWN, (IntPtr)k, GetLParam(k, false));
                    if (dirKeys.Length > 0 && actKeys.Length > 0)
                        Thread.Sleep(15);
                    // 2. 再按動作鍵
                    foreach (var k in actKeys) PostMessage(wnd, WM_KEYDOWN, (IntPtr)k, GetLParam(k, false));
                    Thread.Sleep(KeyPressDurationMs);
                }
                finally
                {
                    for (int i = actKeys.Length - 1; i >= 0; i--)
                        PostMessage(wnd, WM_KEYUP, (IntPtr)actKeys[i], GetLParam(actKeys[i], true));
                    for (int i = dirKeys.Length - 1; i >= 0; i--)
                        PostMessage(wnd, WM_KEYUP, (IntPtr)dirKeys[i], GetLParam(dirKeys[i], true));
                    Thread.Sleep(20);
                }
            }
        }

        private IntPtr GetLParam(Keys key, bool up)
        {
            uint sc = MapVirtualKey((uint)key, 0);
            uint lp = 1 | (sc << 16);
            if (key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down ||
                key == Keys.Home || key == Keys.End || key == Keys.PageUp || key == Keys.PageDown ||
                key == Keys.Insert || key == Keys.Delete) lp |= (1 << 24);
            if (up) { lp |= (1u << 30); lp |= (1u << 31); }
            return (IntPtr)lp;
        }
        #endregion

        private void Log(string msg) { OnLog?.Invoke(msg); OnStatusUpdate?.Invoke(msg); }

        #region 工具
        public static string KeysToDisplayString(Keys[] k) => k == null || k.Length == 0 ? "未設定" : string.Join(" + ", k.Select(x => x.ToString()));
        public static int[] KeysToIntArray(Keys[] k) => k?.Select(x => (int)x).ToArray() ?? Array.Empty<int>();
        public static Keys[] IntArrayToKeys(int[]? a) => a == null || a.Length == 0 ? new[] { Keys.None } : a.Select(i => (Keys)i).ToArray();
        public static Keys[] SingleKeyToArray(int k) => new[] { (Keys)k };
        #endregion
    }

    public class DiagnosticResult
    {
        public string? Error { get; set; }
        public string Direction { get; set; } = "";
        public string KeyNames { get; set; } = "";
        public int BeforeX, BeforeY, AfterX, AfterY, DeltaX, DeltaY;
        public override string ToString() =>
            Error != null ? $"❌ {Error}" :
            $"{Direction}[{KeyNames}]: ({BeforeX},{BeforeY})→({AfterX},{AfterY}) ΔX={DeltaX:+#;-#;0} ΔY={DeltaY:+#;-#;0}";
    }

    public class CorrectionResult
    {
        public bool Success { get; }
        public string Message { get; }
        public int FinalX { get; }
        public int FinalY { get; }
        public int Iterations { get; }
        public long ElapsedMs { get; }
        public CorrectionResult(bool success, string message, int finalX = 0, int finalY = 0, int iterations = 0, long elapsedMs = 0)
        { Success = success; Message = message; FinalX = finalX; FinalY = finalY; Iterations = iterations; ElapsedMs = elapsedMs; }
    }

    public class PositionCorrectionSettings
    {
        public bool Enabled { get; set; } = false;
        // ★ 保留欄位向後兼容（JSON 反序列化不會報錯）
        public bool UseDeviationMode { get; set; } = true;
        public int TargetX { get; set; } = 0;
        public int TargetY { get; set; } = 0;
        public int Tolerance { get; set; } = 5;
        /// <summary>（向後兼容）舊版柔性容差下限，新版請用 HorizontalTolerance / VerticalTolerance</summary>
        public int SoftToleranceMin { get; set; } = 5;
        /// <summary>（向後兼容）舊版柔性容差上限，新版請用 HorizontalTolerance / VerticalTolerance</summary>
        public int SoftToleranceMax { get; set; } = 8;
        /// <summary>水平容差（px）— 水平偏差在此範圍內即視為到位</summary>
        public int HorizontalTolerance { get; set; } = 8;
        /// <summary>垂直容差（px）— 垂直偏差在此範圍內即視為到位</summary>
        public int VerticalTolerance { get; set; } = 5;
        public int MaxCorrectionTimeMs { get; set; } = 5000;
        public bool InvertY { get; set; } = false;
        public int[]? MoveLeftKeys { get; set; } = new[] { (int)Keys.Left };
        public int[]? MoveRightKeys { get; set; } = new[] { (int)Keys.Right };
        public int[]? MoveUpKeys { get; set; } = new[] { (int)Keys.Up };
        public int[]? MoveDownKeys { get; set; } = new[] { (int)Keys.Down };
        public int MoveLeftKey { get; set; } = (int)Keys.Left;
        public int MoveRightKey { get; set; } = (int)Keys.Right;
        public int MoveUpKey { get; set; } = (int)Keys.Up;
        public int MoveDownKey { get; set; } = (int)Keys.Down;
        public bool EnableHorizontalCorrection { get; set; } = true;
        public bool EnableVerticalCorrection { get; set; } = false;
        // ★ 保留欄位向後兼容
        public bool CorrectEveryLoop { get; set; } = true;
        public int CorrectionFrequency { get; set; } = 1;
        /// <summary>每次循環內最多觸發修正次數（0=無限制）</summary>
        public int MaxCorrectionsPerLoop { get; set; } = 0;
        /// <summary>每次修正內最多按鍵步數（0=無限制，僅受時間限制；達到後放棄本次修正）</summary>
        public int MaxStepsPerCorrection { get; set; } = 0;
        public bool UseRecordedPath { get; set; } = false;
        public int PathCheckIntervalEvents { get; set; } = 10;
        public int PathTolerance { get; set; } = 8;
        public int PathMaxCorrectionTimeMs { get; set; } = 3000;
        public bool RecordPathOnCapture { get; set; } = true;
        /// <summary>定期檢查間隔（秒），0=不檢查</summary>
        public int CorrectionCheckIntervalSec { get; set; } = 14;
        /// <summary>修正按鍵間隔下限（毫秒）— 向後兼容</summary>
        public int KeyIntervalMinMs { get; set; } = 700;
        /// <summary>修正按鍵間隔上限（毫秒）— 向後兼容</summary>
        public int KeyIntervalMaxMs { get; set; } = 1200;
        /// <summary>水平修正按鍵間隔下限（毫秒）</summary>
        public int HorizontalKeyIntervalMinMs { get; set; } = 700;
        /// <summary>水平修正按鍵間隔上限（毫秒）</summary>
        public int HorizontalKeyIntervalMaxMs { get; set; } = 1200;
        /// <summary>垂直修正按鍵間隔下限（毫秒）— 垂直通常需要較短間隔</summary>
        public int VerticalKeyIntervalMinMs { get; set; } = 300;
        /// <summary>垂直修正按鍵間隔上限（毫秒）</summary>
        public int VerticalKeyIntervalMaxMs { get; set; } = 600;
        /// <summary>連續跳躍次數（1=單跳，2+=連跳）</summary>
        public int ConsecutiveJumpCount { get; set; } = 1;
        /// <summary>連續跳躍間隔（毫秒）— 按鍵到按鍵的間距（press-to-press）</summary>
        public int ConsecutiveJumpIntervalMs { get; set; } = 80;
        /// <summary>觸發連跳的Y軸偏差閾值（0=總是連跳）</summary>
        public int ConsecutiveJumpThreshold { get; set; } = 0;
        /// <summary>水平長按走路閾值（px）— 偏差小於等於此值用長按走路，大於用熱鍵</summary>
        public int HorizontalHoldWalkThreshold { get; set; } = 30;

        /// <summary>
        /// ★ 爬繩逃脫跳躍鍵（預設 Alt，楓之谷預設跳躍鍵）
        /// 角色在繩上需要水平移動時，改用此鍵配合方向鍵跳離繩子。
        /// </summary>
        public int[]? ClimbEscapeJumpKeys { get; set; } = new[] { (int)Keys.Alt };

        public Keys[] GetEffectiveLeftKeys() => MoveLeftKeys != null && MoveLeftKeys.Length > 0 && MoveLeftKeys[0] != 0 ? PositionCorrector.IntArrayToKeys(MoveLeftKeys) : PositionCorrector.SingleKeyToArray(MoveLeftKey);
        public Keys[] GetEffectiveRightKeys() => MoveRightKeys != null && MoveRightKeys.Length > 0 && MoveRightKeys[0] != 0 ? PositionCorrector.IntArrayToKeys(MoveRightKeys) : PositionCorrector.SingleKeyToArray(MoveRightKey);
        public Keys[] GetEffectiveUpKeys() => MoveUpKeys != null && MoveUpKeys.Length > 0 && MoveUpKeys[0] != 0 ? PositionCorrector.IntArrayToKeys(MoveUpKeys) : PositionCorrector.SingleKeyToArray(MoveUpKey);
        public Keys[] GetEffectiveDownKeys() => MoveDownKeys != null && MoveDownKeys.Length > 0 && MoveDownKeys[0] != 0 ? PositionCorrector.IntArrayToKeys(MoveDownKeys) : PositionCorrector.SingleKeyToArray(MoveDownKey);
        public Keys[] GetEffectiveClimbJumpKeys() => ClimbEscapeJumpKeys != null && ClimbEscapeJumpKeys.Length > 0 ? PositionCorrector.IntArrayToKeys(ClimbEscapeJumpKeys) : new[] { Keys.Alt };
    }
}
