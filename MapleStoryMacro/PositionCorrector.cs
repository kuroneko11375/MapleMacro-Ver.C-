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
                    var (cx, cy, ok) = tracker.ReadPosition();
                    if (!ok) { Log($"[{iter}] 偵測失敗"); Thread.Sleep(200); continue; }

                    int dx = targetX - cx, dy = targetY - cy;
                    bool xOk = !EnableHorizontalCorrection || Math.Abs(dx) <= hTol;
                    bool yOk = !EnableVerticalCorrection || Math.Abs(dy) <= vTol;

                    Log($"[{iter}] 在({cx},{cy}) 差({dx:+#;-#;0},{dy:+#;-#;0}) 容差H±{hTol}/V±{vTol} 步數{keyPressCount}/{(MaxStepsPerCorrection > 0 ? MaxStepsPerCorrection.ToString() : "∞")} {(xOk ? "X✓" : "X✗")} {(yOk ? "Y✓" : "Y✗")}");

                    // ★ 進入容差範圍就立即結束，不再追求更精準
                    if (xOk && yOk)
                    {
                        var msg = $"修正完成！({cx},{cy}) 偏差({dx:+#;-#;0},{dy:+#;-#;0}) 容差H±{hTol}/V±{vTol} 按鍵{keyPressCount}次 {sw.ElapsedMilliseconds}ms";
                        Log("✅ " + msg);
                        LastCorrectionSuccess = true; LastCorrectionMessage = msg;
                        AddHistory(initialDx, initialDy, keyPressCount, true);
                        return new CorrectionResult(true, msg, cx, cy, keyPressCount, sw.ElapsedMilliseconds);
                    }

                    bool needVertical = EnableVerticalCorrection && !yOk;
                    bool needHorizontal = EnableHorizontalCorrection && !xOk;
                    bool xApproxOk = !EnableHorizontalCorrection || Math.Abs(dx) <= hTol * 2;

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

                        var keysToSend = new List<Keys>(jumpKeys);
                        if (rawDirKey.HasValue) keysToSend.Add(rawDirKey.Value);
                        var unique = keysToSend.Distinct().ToArray();

                        Log($"[{iter}] 垂直修正 {dir} [{string.Join("+", unique.Select(k => k.ToString()))}]");
                        SendCombinedKeyPress(unique);
                        keyPressCount++;
                        SleepKeyInterval(true);

                        // 按鍵後立即檢查
                        var (vx, vy, vOk) = tracker.ReadPosition();
                        if (vOk)
                        {
                            int vdx = targetX - vx, vdy = targetY - vy;
                            bool vxOk = !EnableHorizontalCorrection || Math.Abs(vdx) <= hTol;
                            bool vyOk = !EnableVerticalCorrection || Math.Abs(vdy) <= vTol;
                            if (vxOk && vyOk)
                            {
                                var msg = $"修正完成！({vx},{vy}) 偏差({vdx:+#;-#;0},{vdy:+#;-#;0}) 容差H±{hTol}/V±{vTol} 按鍵{keyPressCount}次 {sw.ElapsedMilliseconds}ms";
                                Log("✅ " + msg);
                                LastCorrectionSuccess = true; LastCorrectionMessage = msg;
                                AddHistory(initialDx, initialDy, keyPressCount, true);
                                return new CorrectionResult(true, msg, vx, vy, keyPressCount, sw.ElapsedMilliseconds);
                            }
                        }
                    }
                    else if (needHorizontal)
                    {
                        string dir;
                        Keys[] keys;
                        int expectDx;
                        if (dx > 0) { keys = MoveRightKeys; dir = "→"; expectDx = 1; }
                        else { keys = MoveLeftKeys; dir = "←"; expectDx = -1; }

                        Log($"[{iter}] 水平修正 {dir} [{string.Join("+", keys.Select(k => k.ToString()))}]");
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
            finally { IsCorreecting = false; }
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

            SendCombinedKeyPress(keys);
            Thread.Sleep(350);

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
        private void SendCombinedKeyPress(Keys[] keys)
        {
            if (keys == null || keys.Length == 0) return;
            if (ExternalKeySender != null)
            {
                foreach (var k in keys) ExternalKeySender(TargetWindow, k, true);
                Thread.Sleep(KeyPressDurationMs);
                for (int i = keys.Length - 1; i >= 0; i--) ExternalKeySender(TargetWindow, keys[i], false);
            }
            else
            {
                IntPtr child = FindWindowEx(TargetWindow, IntPtr.Zero, null, null);
                IntPtr wnd = child != IntPtr.Zero ? child : TargetWindow;
                foreach (var k in keys) PostMessage(wnd, WM_KEYDOWN, (IntPtr)k, GetLParam(k, false));
                Thread.Sleep(KeyPressDurationMs);
                for (int i = keys.Length - 1; i >= 0; i--) PostMessage(wnd, WM_KEYUP, (IntPtr)keys[i], GetLParam(keys[i], true));
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

        public Keys[] GetEffectiveLeftKeys() => MoveLeftKeys != null && MoveLeftKeys.Length > 0 && MoveLeftKeys[0] != 0 ? PositionCorrector.IntArrayToKeys(MoveLeftKeys) : PositionCorrector.SingleKeyToArray(MoveLeftKey);
        public Keys[] GetEffectiveRightKeys() => MoveRightKeys != null && MoveRightKeys.Length > 0 && MoveRightKeys[0] != 0 ? PositionCorrector.IntArrayToKeys(MoveRightKeys) : PositionCorrector.SingleKeyToArray(MoveRightKey);
        public Keys[] GetEffectiveUpKeys() => MoveUpKeys != null && MoveUpKeys.Length > 0 && MoveUpKeys[0] != 0 ? PositionCorrector.IntArrayToKeys(MoveUpKeys) : PositionCorrector.SingleKeyToArray(MoveUpKey);
        public Keys[] GetEffectiveDownKeys() => MoveDownKeys != null && MoveDownKeys.Length > 0 && MoveDownKeys[0] != 0 ? PositionCorrector.IntArrayToKeys(MoveDownKeys) : PositionCorrector.SingleKeyToArray(MoveDownKey);
    }
}
