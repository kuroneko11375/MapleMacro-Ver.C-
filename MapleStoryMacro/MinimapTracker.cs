using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace MapleStoryMacro
{
    /// <summary>
    /// 小地圖視覺追蹤器 - 透過分析小地圖上的黃色角色圖標來追蹤位置
    /// </summary>
    public class MinimapTracker : IDisposable
    {
        #region P/Invoke
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest,
            IntPtr hdcSrc, int xSrc, int ySrc, int rop);

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

        private const int SRCCOPY = 0x00CC0020;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X, Y;
        }
        #endregion

        #region 設定與狀態
        /// <summary>小地圖在遊戲視窗中的相對位置</summary>
        public Rectangle MinimapRegion { get; set; } = new Rectangle(0, 0, 0, 0);

        /// <summary>是否已校準 (小地圖區域已設定)</summary>
        public bool IsCalibrated => MinimapRegion.Width > 0 && MinimapRegion.Height > 0;

        /// <summary>地圖座標邊界 (用於像素→座標轉換)</summary>
        public Rectangle MapBounds { get; set; } = new Rectangle(0, 0, 0, 0);

        /// <summary>目標視窗句柄</summary>
        public IntPtr TargetWindow { get; private set; } = IntPtr.Zero;

        /// <summary>最後偵測到的位置 (小地圖像素座標)</summary>
        public Point LastPixelPosition { get; private set; }

        /// <summary>最後偵測到的位置 (遊戲座標)</summary>
        public Point LastGamePosition { get; private set; }

        /// <summary>最後一次偵測是否成功</summary>
        public bool LastDetectionSuccess { get; private set; }

        /// <summary>偵測信心度 (0-100)</summary>
        public int DetectionConfidence { get; private set; }
        #endregion

        #region 顏色過濾參數 (可調整)
        /// <summary>黃色色相範圍 (HSV H: 0-360)</summary>
        public (float Min, float Max) HueRange { get; set; } = (45f, 70f);

        /// <summary>最小飽和度 (0-1)</summary>
        public float MinSaturation { get; set; } = 0.55f;

        /// <summary>最小亮度 (0-1)</summary>
        public float MinValue { get; set; } = 0.70f;

        /// <summary>最小圖標像素數</summary>
        public int MinIconPixels { get; set; } = 4;

        /// <summary>最大圖標像素數</summary>
        public int MaxIconPixels { get; set; } = 100;

        /// <summary>邊界判定容差(px)</summary>
        public int BoundsTolerance { get; set; } = 5;
        #endregion

        #region 位置連續性追蹤（防止誤認 NPC/怪物）
        /// <summary>上一次偵測到的像素位置（用於連續性判斷）</summary>
        private Point? _lastKnownPixelPos = null;

        /// <summary>連續偵測失敗次數</summary>
        private int _consecutiveMissCount = 0;

        /// <summary>位置連續性：最大允許跳動距離(px)，超過此距離的聚類將被降權</summary>
        public int MaxPositionJumpPixels { get; set; } = 20;

        /// <summary>最小圖標緊湊度 (像素數 / 邊界框面積, 0~1)，過低代表形狀不規則</summary>
        public float MinCompactness { get; set; } = 0.15f;

        /// <summary>連續偵測失敗幾次後清除位置記憶（允許重新搜尋）</summary>
        public int MaxConsecutiveMissBeforeReset { get; set; } = 10;

        /// <summary>清除位置記憶（停止播放或換地圖時呼叫）</summary>
        public void ResetPositionMemory()
        {
            _lastKnownPixelPos = null;
            _consecutiveMissCount = 0;
        }

        /// <summary>最後一次偵測的聚類數量（含所有大小，供除錯顯示）</summary>
        public int LastTotalClusterCount { get; private set; }

        /// <summary>最後一次偵測通過過濾的有效聚類數量</summary>
        public int LastValidClusterCount { get; private set; }
        #endregion

        /// <summary>
        /// 附加到目標視窗
        /// </summary>
        public bool AttachToWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return false;
            TargetWindow = hWnd;
            return true;
        }

        /// <summary>
        /// 分離
        /// </summary>
        public void Detach()
        {
            TargetWindow = IntPtr.Zero;
        }

        /// <summary>
        /// 取得目標視窗的客戶區尺寸（像素）
        /// </summary>
        public (int width, int height) GetClientAreaSize()
        {
            if (TargetWindow == IntPtr.Zero) return (0, 0);
            GetClientRect(TargetWindow, out RECT r);
            return (r.Right - r.Left, r.Bottom - r.Top);
        }

        /// <summary>
        /// 使用 PrintWindow 截取客戶區（背景擷取，視窗被遮擋或失焦仍可用）
        /// </summary>
        private Bitmap? CaptureClientAreaBackground()
        {
            if (TargetWindow == IntPtr.Zero) return null;

            try
            {
                GetClientRect(TargetWindow, out RECT clientRect);
                int width = clientRect.Right - clientRect.Left;
                int height = clientRect.Bottom - clientRect.Top;
                if (width <= 0 || height <= 0) return null;

                Bitmap bmp = new Bitmap(width, height, PixelFormat.Format24bppRgb);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    IntPtr hdc = g.GetHdc();
                    // PW_CLIENTONLY (0x1) = 只截取客戶區
                    // PW_RENDERFULLCONTENT (0x2) = 含 DirectComposition 內容
                    bool ok = PrintWindow(TargetWindow, hdc, 0x1 | 0x2);
                    g.ReleaseHdc(hdc);
                    if (!ok)
                    {
                        // 備援：不帶 PW_RENDERFULLCONTENT
                        hdc = g.GetHdc();
                        ok = PrintWindow(TargetWindow, hdc, 0x1);
                        g.ReleaseHdc(hdc);
                    }
                    if (!ok)
                    {
                        bmp.Dispose();
                        return null;
                    }
                }
                return bmp;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 截取遊戲視窗的指定區域（背景模式：使用 PrintWindow）
        /// </summary>
        public Bitmap? CaptureRegion(Rectangle region)
        {
            if (TargetWindow == IntPtr.Zero) return null;

            try
            {
                int width = region.Width;
                int height = region.Height;
                if (width <= 0 || height <= 0) return null;

                // 先截取完整客戶區，再裁切指定區域
                using (var fullBmp = CaptureClientAreaBackground())
                {
                    if (fullBmp == null) return null;

                    // 確保區域不超出截圖範圍
                    int srcX = Math.Max(0, Math.Min(region.X, fullBmp.Width - 1));
                    int srcY = Math.Max(0, Math.Min(region.Y, fullBmp.Height - 1));
                    int srcW = Math.Min(width, fullBmp.Width - srcX);
                    int srcH = Math.Min(height, fullBmp.Height - srcY);
                    if (srcW <= 0 || srcH <= 0) return null;

                    Bitmap cropped = new Bitmap(srcW, srcH, PixelFormat.Format24bppRgb);
                    using (Graphics g = Graphics.FromImage(cropped))
                    {
                        g.DrawImage(fullBmp, new Rectangle(0, 0, srcW, srcH),
                            new Rectangle(srcX, srcY, srcW, srcH), GraphicsUnit.Pixel);
                    }
                    return cropped;
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 截取小地圖區域（背景模式）
        /// </summary>
        public Bitmap? CaptureMinimap()
        {
            if (MinimapRegion.Width <= 0 || MinimapRegion.Height <= 0)
                return null;
            return CaptureRegion(MinimapRegion);
        }

        /// <summary>
        /// 截取整個遊戲視窗（背景模式：使用 PrintWindow）
        /// </summary>
        public Bitmap? CaptureFullWindow()
        {
            return CaptureClientAreaBackground();
        }

        /// <summary>
        /// 從圖片中偵測角色位置
        /// ★ 使用多因素評分系統：大小、緊湊度、位置連續性、邊界範圍
        /// </summary>
        public (Point pixelPos, bool success, int confidence) DetectPositionFromBitmap(Bitmap bmp)
        {
            var yellowPixels = FindYellowPixels(bmp);

            if (yellowPixels.Count < MinIconPixels)
            {
                _consecutiveMissCount++;
                if (_consecutiveMissCount >= MaxConsecutiveMissBeforeReset)
                    _lastKnownPixelPos = null;
                return (Point.Empty, false, 0);
            }

            // 聚類分析找出獨立圖標
            var clusters = ClusterPixels(yellowPixels);
            LastTotalClusterCount = clusters.Count;

            if (clusters.Count == 0)
            {
                LastValidClusterCount = 0;
                _consecutiveMissCount++;
                if (_consecutiveMissCount >= MaxConsecutiveMissBeforeReset)
                    _lastKnownPixelPos = null;
                return (Point.Empty, false, 0);
            }

            // ★ 過濾：大小範圍 + 緊湊度
            var validClusters = new List<(List<Point> cluster, float compactness, Point center, int bboxW, int bboxH)>();
            foreach (var cluster in clusters)
            {
                if (cluster.Count < MinIconPixels || cluster.Count > MaxIconPixels)
                    continue;

                int cMinX = cluster.Min(p => p.X), cMaxX = cluster.Max(p => p.X);
                int cMinY = cluster.Min(p => p.Y), cMaxY = cluster.Max(p => p.Y);
                int bboxW = cMaxX - cMinX + 1;
                int bboxH = cMaxY - cMinY + 1;
                float bboxArea = bboxW * bboxH;
                float compactness = bboxArea > 0 ? cluster.Count / bboxArea : 0;

                // 緊湊度過濾：形狀太不規則的排除（NPC 圖標通常比較大且不規則）
                if (compactness < MinCompactness)
                    continue;

                int cx = (int)cluster.Average(p => p.X);
                int cy = (int)cluster.Average(p => p.Y);
                validClusters.Add((cluster, compactness, new Point(cx, cy), bboxW, bboxH));
            }

            if (validClusters.Count == 0)
            {
                LastValidClusterCount = 0;
                _consecutiveMissCount++;
                if (_consecutiveMissCount >= MaxConsecutiveMissBeforeReset)
                    _lastKnownPixelPos = null;
                return (Point.Empty, false, 10);
            }

            LastValidClusterCount = validClusters.Count;

            // ★ 多因素評分選擇最佳聚類
            var bestCluster = SelectBestCluster(validClusters);

            // 更新位置記憶
            _lastKnownPixelPos = bestCluster.center;
            _consecutiveMissCount = 0;

            // 計算信心度
            int confidence = CalculateConfidence(bestCluster.cluster, clusters);

            return (bestCluster.center, true, confidence);
        }

        /// <summary>
        /// ★ 多因素評分選擇最佳聚類
        /// 評分因素：位置連續性 > 邊界範圍 > 緊湊度 > 大小適中
        /// </summary>
        private (List<Point> cluster, float compactness, Point center, int bboxW, int bboxH)
            SelectBestCluster(List<(List<Point> cluster, float compactness, Point center, int bboxW, int bboxH)> candidates)
        {
            if (candidates.Count == 1)
                return candidates[0];

            double bestScore = double.MinValue;
            int bestIdx = 0;

            // 找出最大像素數用於歸一化
            int maxPixels = candidates.Max(c => c.cluster.Count);

            for (int i = 0; i < candidates.Count; i++)
            {
                var c = candidates[i];
                double score = 0;

                // 1. 位置連續性（權重最高）：離上次位置越近分數越高
                if (_lastKnownPixelPos.HasValue)
                {
                    double dist = Math.Sqrt(
                        Math.Pow(c.center.X - _lastKnownPixelPos.Value.X, 2) +
                        Math.Pow(c.center.Y - _lastKnownPixelPos.Value.Y, 2));

                    if (dist <= MaxPositionJumpPixels)
                        score += 50.0;  // 在合理跳動範圍內：大加分
                    else if (dist <= MaxPositionJumpPixels * 2)
                        score += 25.0;  // 稍遠但仍可能
                    else
                        score -= 20.0;  // 太遠：大幅降權（可能是 NPC）
                }

                // 2. 邊界範圍內加分
                if (MapBounds.Width > 0 && MapBounds.Height > 0)
                {
                    bool inBounds = c.center.X >= (MapBounds.Left - BoundsTolerance) &&
                                    c.center.X <= (MapBounds.Left + MapBounds.Width + BoundsTolerance) &&
                                    c.center.Y >= (MapBounds.Top - BoundsTolerance) &&
                                    c.center.Y <= (MapBounds.Top + MapBounds.Height + BoundsTolerance);
                    if (inBounds)
                        score += 15.0;
                    else
                        score -= 10.0;
                }

                // 3. 緊湊度加分（角色圖標通常是緊湊的小點）
                score += c.compactness * 20.0;

                // 4. 大小適中加分（角色圖標通常 5~25 像素）
                int pixelCount = c.cluster.Count;
                if (pixelCount >= 5 && pixelCount <= 25)
                    score += 10.0;
                else if (pixelCount >= MinIconPixels && pixelCount <= 40)
                    score += 5.0;

                // 5. 邊界框大小合理（角色圖標通常 ≤ 8x8）
                if (c.bboxW <= 8 && c.bboxH <= 8)
                    score += 8.0;
                else if (c.bboxW <= 12 && c.bboxH <= 12)
                    score += 3.0;

                if (score > bestScore)
                {
                    bestScore = score;
                    bestIdx = i;
                }
            }

            return candidates[bestIdx];
        }

        /// <summary>
        /// 偵測當前位置 (自動截圖並分析)
        /// 返回的是小地圖上的像素座標
        /// </summary>
        public (int x, int y, bool success) ReadPosition()
        {
            LastDetectionSuccess = false;
            DetectionConfidence = 0;

            if (TargetWindow == IntPtr.Zero || MinimapRegion.Width <= 0 || MinimapRegion.Height <= 0)
            {
                return (0, 0, false);
            }

            using (var bmp = CaptureMinimap())
            {
                if (bmp == null)
                {
                    return (0, 0, false);
                }

                var (pixelPos, success, confidence) = DetectPositionFromBitmap(bmp);

                if (!success)
                {
                    return (0, 0, false);
                }

                LastPixelPosition = pixelPos;
                LastDetectionSuccess = true;
                DetectionConfidence = confidence;
                LastGamePosition = pixelPos; // 現在直接使用像素座標

                return (pixelPos.X, pixelPos.Y, true);
            }
        }

        /// <summary>
        /// 檢查當前位置是否在設定的邊界內
        /// </summary>
        public bool IsInBounds()
        {
            if (!LastDetectionSuccess) return false;
            
            int x = LastPixelPosition.X;
            int y = LastPixelPosition.Y;
            
            return x >= (MapBounds.Left - BoundsTolerance) && x <= (MapBounds.Right + BoundsTolerance) &&
                   y >= (MapBounds.Top - BoundsTolerance) && y <= (MapBounds.Bottom + BoundsTolerance);
        }

        /// <summary>
        /// 檢查指定位置是否在設定的邊界內
        /// </summary>
        public bool IsInBounds(int x, int y)
        {
            return x >= (MapBounds.Left - BoundsTolerance) && x <= (MapBounds.Right + BoundsTolerance) &&
                   y >= (MapBounds.Top - BoundsTolerance) && y <= (MapBounds.Bottom + BoundsTolerance);
        }

        /// <summary>
        /// 將小地圖像素座標轉換為遊戲座標
        /// </summary>
        public Point PixelToGameCoord(Point pixelPos)
        {
            if (MinimapRegion.Width <= 0 || MinimapRegion.Height <= 0)
                return Point.Empty;

            // 線性插值
            double ratioX = (double)pixelPos.X / MinimapRegion.Width;
            double ratioY = (double)pixelPos.Y / MinimapRegion.Height;

            int gameX = MapBounds.Left + (int)(ratioX * MapBounds.Width);
            int gameY = MapBounds.Top + (int)(ratioY * MapBounds.Height);

            return new Point(gameX, gameY);
        }

        /// <summary>
        /// 將遊戲座標轉換為小地圖像素座標
        /// </summary>
        public Point GameCoordToPixel(Point gamePos)
        {
            if (MapBounds.Width <= 0 || MapBounds.Height <= 0)
                return Point.Empty;

            double ratioX = (double)(gamePos.X - MapBounds.Left) / MapBounds.Width;
            double ratioY = (double)(gamePos.Y - MapBounds.Top) / MapBounds.Height;

            int pixelX = (int)(ratioX * MinimapRegion.Width);
            int pixelY = (int)(ratioY * MinimapRegion.Height);

            return new Point(pixelX, pixelY);
        }

        #region 顏色分析
        /// <summary>
        /// 找出所有符合黃色條件的像素
        /// </summary>
        private List<Point> FindYellowPixels(Bitmap bmp)
        {
            var result = new List<Point>();

            BitmapData data = bmp.LockBits(
                new Rectangle(0, 0, bmp.Width, bmp.Height),
                ImageLockMode.ReadOnly,
                PixelFormat.Format24bppRgb);

            try
            {
                int stride = data.Stride;
                int width = bmp.Width;
                int height = bmp.Height;
                IntPtr scan0 = data.Scan0;

                unsafe
                {
                    byte* ptr = (byte*)scan0;

                    for (int y = 0; y < height; y++)
                    {
                        byte* row = ptr + (y * stride);
                        for (int x = 0; x < width; x++)
                        {
                            int idx = x * 3;
                            byte b = row[idx];
                            byte g = row[idx + 1];
                            byte r = row[idx + 2];

                            if (IsYellowPixel(r, g, b))
                            {
                                result.Add(new Point(x, y));
                            }
                        }
                    }
                }
            }
            finally
            {
                bmp.UnlockBits(data);
            }

            return result;
        }

        /// <summary>
        /// 判斷像素是否為黃色 (使用 HSV 色彩空間)
        /// </summary>
        private bool IsYellowPixel(byte r, byte g, byte b)
        {
            // RGB → HSV
            float rf = r / 255f;
            float gf = g / 255f;
            float bf = b / 255f;

            float max = Math.Max(rf, Math.Max(gf, bf));
            float min = Math.Min(rf, Math.Min(gf, bf));
            float delta = max - min;

            // Value (亮度)
            float v = max;
            if (v < MinValue) return false;

            // Saturation (飽和度)
            float s = (max == 0) ? 0 : delta / max;
            if (s < MinSaturation) return false;

            // Hue (色相)
            float h = 0;
            if (delta > 0)
            {
                if (max == rf)
                    h = 60f * (((gf - bf) / delta) % 6);
                else if (max == gf)
                    h = 60f * (((bf - rf) / delta) + 2);
                else
                    h = 60f * (((rf - gf) / delta) + 4);
            }
            if (h < 0) h += 360f;

            return h >= HueRange.Min && h <= HueRange.Max;
        }

        /// <summary>
        /// 將像素聚類成獨立圖標
        /// </summary>
        private List<List<Point>> ClusterPixels(List<Point> pixels)
        {
            var clusters = new List<List<Point>>();
            var used = new HashSet<Point>();
            var pixelSet = new HashSet<Point>(pixels);

            foreach (var startPixel in pixels)
            {
                if (used.Contains(startPixel)) continue;

                // BFS 找連通區域
                var cluster = new List<Point>();
                var queue = new Queue<Point>();
                queue.Enqueue(startPixel);

                while (queue.Count > 0)
                {
                    var current = queue.Dequeue();
                    if (used.Contains(current)) continue;

                    used.Add(current);
                    cluster.Add(current);

                    // 檢查 8 鄰域
                    for (int dx = -1; dx <= 1; dx++)
                    {
                        for (int dy = -1; dy <= 1; dy++)
                        {
                            if (dx == 0 && dy == 0) continue;
                            var neighbor = new Point(current.X + dx, current.Y + dy);
                            if (pixelSet.Contains(neighbor) && !used.Contains(neighbor))
                            {
                                queue.Enqueue(neighbor);
                            }
                        }
                    }
                }

                if (cluster.Count >= MinIconPixels)
                {
                    clusters.Add(cluster);
                }
            }

            return clusters;
        }

        /// <summary>
        /// 計算偵測信心度
        /// </summary>
        private int CalculateConfidence(List<Point> mainCluster, List<List<Point>> allClusters)
        {
            int confidence = 40; // 基礎分

            // 圖標大小合理 (+20)
            if (mainCluster.Count >= 5 && mainCluster.Count <= 25)
                confidence += 20;
            else if (mainCluster.Count >= MinIconPixels && mainCluster.Count <= MaxIconPixels)
                confidence += 10;

            // 只有一個黃色聚類 (+20)
            if (allClusters.Count == 1)
                confidence += 20;
            else if (allClusters.Count <= 3)
                confidence += 10;
            // 多個聚類 = 干擾較多，信心較低
            else
                confidence -= 5;

            // 聚類形狀緊湊 (+10)
            int width = mainCluster.Max(p => p.X) - mainCluster.Min(p => p.X);
            int height = mainCluster.Max(p => p.Y) - mainCluster.Min(p => p.Y);
            if (width <= 6 && height <= 6)
                confidence += 15;
            else if (width <= 8 && height <= 8)
                confidence += 10;

            // 位置連續性加分 (+10)
            if (_lastKnownPixelPos.HasValue)
            {
                int cx = (int)mainCluster.Average(p => p.X);
                int cy = (int)mainCluster.Average(p => p.Y);
                double dist = Math.Sqrt(
                    Math.Pow(cx - _lastKnownPixelPos.Value.X, 2) +
                    Math.Pow(cy - _lastKnownPixelPos.Value.Y, 2));
                if (dist <= MaxPositionJumpPixels)
                    confidence += 10;
            }

            return Math.Clamp(confidence, 0, 100);
        }
        #endregion

        #region 自動偵測小地圖位置
        /// <summary>
        /// 自動偵測小地圖在遊戲視窗中的位置
        /// 原理：小地圖通常在右上角，且有明顯的邊框
        /// </summary>
        public Rectangle? AutoDetectMinimapRegion()
        {
            if (TargetWindow == IntPtr.Zero) return null;

            // 取得視窗大小
            GetClientRect(TargetWindow, out RECT clientRect);
            int windowWidth = clientRect.Right - clientRect.Left;
            int windowHeight = clientRect.Bottom - clientRect.Top;

            // 小地圖通常在右上角，搜尋範圍
            int searchWidth = Math.Min(300, windowWidth / 3);
            int searchHeight = Math.Min(200, windowHeight / 3);
            var searchRegion = new Rectangle(windowWidth - searchWidth, 0, searchWidth, searchHeight);

            using (var screenshot = CaptureRegion(searchRegion))
            {
                if (screenshot == null) return null;

                // 尋找小地圖的特徵：
                // 1. 有黃色角色圖標
                // 2. 可能有綠色 NPC 圖標
                // 3. 通常有深色或半透明背景

                var yellowPixels = FindYellowPixels(screenshot);
                if (yellowPixels.Count < MinIconPixels) return null;

                // 找到黃色像素的邊界
                int minX = yellowPixels.Min(p => p.X);
                int maxX = yellowPixels.Max(p => p.X);
                int minY = yellowPixels.Min(p => p.Y);
                int maxY = yellowPixels.Max(p => p.Y);

                // 擴展邊界來包含整個小地圖 (估計)
                int padding = 50;
                int mapX = Math.Max(0, minX - padding);
                int mapY = Math.Max(0, minY - padding);
                int mapWidth = Math.Min(screenshot.Width - mapX, (maxX - minX) + padding * 2);
                int mapHeight = Math.Min(screenshot.Height - mapY, (maxY - minY) + padding * 2);

                // 轉換回視窗座標
                return new Rectangle(
                    searchRegion.X + mapX,
                    searchRegion.Y + mapY,
                    mapWidth,
                    mapHeight
                );
            }
        }
        #endregion

        #region 儲存/載入設定
        /// <summary>
        /// 儲存校準設定
        /// </summary>
        public void SaveCalibration(string filePath)
        {
            var data = new CalibrationData
            {
                MinimapX = MinimapRegion.X,
                MinimapY = MinimapRegion.Y,
                MinimapWidth = MinimapRegion.Width,
                MinimapHeight = MinimapRegion.Height,
                MapLeft = MapBounds.Left,
                MapTop = MapBounds.Top,
                MapWidth = MapBounds.Width,
                MapHeight = MapBounds.Height,
                HueMin = HueRange.Min,
                HueMax = HueRange.Max,
                MinSaturation = MinSaturation,
                MinValue = MinValue,
                MaxPositionJumpPixels = MaxPositionJumpPixels,
                MinCompactness = MinCompactness,
                SavedAt = DateTime.Now
            };

            string json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
            string? dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            File.WriteAllText(filePath, json);
        }

        /// <summary>
        /// 載入校準設定
        /// </summary>
        public bool LoadCalibration(string filePath)
        {
            if (!File.Exists(filePath)) return false;

            try
            {
                string json = File.ReadAllText(filePath);
                var data = JsonSerializer.Deserialize<CalibrationData>(json);

                if (data != null && data.MinimapWidth > 0)
                {
                    MinimapRegion = new Rectangle(data.MinimapX, data.MinimapY, data.MinimapWidth, data.MinimapHeight);
                    MapBounds = new Rectangle(data.MapLeft, data.MapTop, data.MapWidth, data.MapHeight);
                    HueRange = (data.HueMin, data.HueMax);
                    MinSaturation = data.MinSaturation;
                    MinValue = data.MinValue;
                    // ★ 新參數（向後相容：舊檔案預設為 0，用預設值代替）
                    if (data.MaxPositionJumpPixels > 0)
                        MaxPositionJumpPixels = data.MaxPositionJumpPixels;
                    if (data.MinCompactness > 0)
                        MinCompactness = data.MinCompactness;
                    return true;
                }
            }
            catch { }

            return false;
        }

        public class CalibrationData
        {
            public int MinimapX { get; set; }
            public int MinimapY { get; set; }
            public int MinimapWidth { get; set; }
            public int MinimapHeight { get; set; }
            public int MapLeft { get; set; }
            public int MapTop { get; set; }
            public int MapWidth { get; set; }
            public int MapHeight { get; set; }
            public float HueMin { get; set; }
            public float HueMax { get; set; }
            public float MinSaturation { get; set; }
            public float MinValue { get; set; }
            public int MaxPositionJumpPixels { get; set; }
            public float MinCompactness { get; set; }
            public DateTime SavedAt { get; set; }
        }
        #endregion

        #region 除錯工具
        /// <summary>
        /// 產生除錯圖片，標記偵測到的黃色像素和聚類
        /// ★ 不同聚類用不同顏色標記，被選中的用紅色，被淘汰的用灰色
        /// </summary>
        public Bitmap? CreateDebugImage()
        {
            using (var bmp = CaptureMinimap())
            {
                if (bmp == null) return null;

                var debugBmp = new Bitmap(bmp);
                var yellowPixels = FindYellowPixels(bmp);
                var clusters = ClusterPixels(yellowPixels);

                // 聚類顏色列表（被淘汰的聚類用不同顏色區分）
                Color[] clusterColors = { Color.Orange, Color.Magenta, Color.Yellow, Color.Lime, Color.Aqua };

                // 標記所有聚類
                for (int ci = 0; ci < clusters.Count; ci++)
                {
                    var cluster = clusters[ci];
                    Color clrMark = clusterColors[ci % clusterColors.Length];
                    // 過小或過大的聚類用深灰色表示被淘汰
                    if (cluster.Count < MinIconPixels || cluster.Count > MaxIconPixels)
                        clrMark = Color.FromArgb(80, 80, 80);

                    foreach (var p in cluster)
                    {
                        if (p.X >= 0 && p.X < debugBmp.Width && p.Y >= 0 && p.Y < debugBmp.Height)
                            debugBmp.SetPixel(p.X, p.Y, clrMark);
                    }
                }

                // 標記偵測到的中心點
                if (LastDetectionSuccess)
                {
                    int cx = LastPixelPosition.X;
                    int cy = LastPixelPosition.Y;
                    for (int dx = -3; dx <= 3; dx++)
                    {
                        for (int dy = -3; dy <= 3; dy++)
                        {
                            int px = cx + dx;
                            int py = cy + dy;
                            if (px >= 0 && px < debugBmp.Width && py >= 0 && py < debugBmp.Height)
                            {
                                debugBmp.SetPixel(px, py, Color.Cyan);
                            }
                        }
                    }
                }

                return debugBmp;
            }
        }
        #endregion

        public void Dispose()
        {
            Detach();
        }
    }
}
