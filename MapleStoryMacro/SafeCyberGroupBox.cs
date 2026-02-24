using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using ReaLTaiizor.Controls;

namespace MapleStoryMacro
{
    /// <summary>
    /// 安全的 CyberGroupBox 替代品
    /// 完全自定義繪製邏輯，不調用 ReaLTaiizor 的任何繪製方法
    /// 保持與 CyberGroupBox 相同的屬性接口以確保 Designer 兼容性
    /// </summary>
    public class SafeCyberGroupBox : CyberGroupBox
    {
        public SafeCyberGroupBox() : base()
        {
            this.SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw |
                ControlStyles.SupportsTransparentBackColor,
                true);
            
            this.DoubleBuffered = true;
            this.MinimumSize = new Size(40, 40);
        }

        /// <summary>
        /// 完全覆寫 OnPaint，不調用 base.OnPaint() 以避免 ReaLTaiizor 的繪製錯誤
        /// </summary>
        protected override void OnPaint(PaintEventArgs e)
        {
            // 不調用 base.OnPaint(e) - 這是關鍵！
            // 完全使用自定義繪製邏輯

            if (this.Width <= 0 || this.Height <= 0)
                return;

            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = this.TextRenderingHint;

            try
            {
                DrawBackground(g);
                DrawBorder(g);
            }
            catch
            {
                // 備用方案：簡單填充
                try
                {
                    using (var brush = new SolidBrush(this.ColorBackground))
                    {
                        g.FillRectangle(brush, this.ClientRectangle);
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 繪製背景
        /// </summary>
        private void DrawBackground(Graphics g)
        {
            Rectangle rect = this.ClientRectangle;
            
            // 確保有有效的繪製區域
            if (rect.Width <= 0 || rect.Height <= 0)
                return;

            int radius = GetSafeRadius(rect);

            if (this.LinearGradient_Background)
            {
                // 漸層背景
                using (var brush = new LinearGradientBrush(
                    rect,
                    this.ColorBackground_1,
                    this.ColorBackground_2,
                    LinearGradientMode.Vertical))
                {
                    if (this.Rounding && radius > 0)
                    {
                        using (var path = CreateRoundedRectPath(rect, radius))
                        {
                            g.FillPath(brush, path);
                        }
                    }
                    else
                    {
                        g.FillRectangle(brush, rect);
                    }
                }
            }
            else
            {
                // 純色背景
                using (var brush = new SolidBrush(this.ColorBackground))
                {
                    if (this.Rounding && radius > 0)
                    {
                        using (var path = CreateRoundedRectPath(rect, radius))
                        {
                            g.FillPath(brush, path);
                        }
                    }
                    else
                    {
                        g.FillRectangle(brush, rect);
                    }
                }
            }
        }

        /// <summary>
        /// 繪製邊框
        /// </summary>
        private void DrawBorder(Graphics g)
        {
            if (!this.BackgroundPen)
                return;

            float penWidth = Math.Max(1, this.Background_WidthPen);
            int offset = (int)Math.Ceiling(penWidth / 2) + 1;
            
            Rectangle borderRect = new Rectangle(
                offset,
                offset,
                this.Width - offset * 2,
                this.Height - offset * 2);

            if (borderRect.Width <= 0 || borderRect.Height <= 0)
                return;

            int radius = GetSafeRadius(borderRect);

            using (var pen = new Pen(this.ColorBackground_Pen, penWidth))
            {
                if (this.LinearGradientPen)
                {
                    using (var brush = new LinearGradientBrush(
                        borderRect,
                        this.ColorPen_1,
                        this.ColorPen_2,
                        LinearGradientMode.Vertical))
                    {
                        pen.Brush = brush;
                        DrawBorderPath(g, pen, borderRect, radius);
                    }
                }
                else
                {
                    DrawBorderPath(g, pen, borderRect, radius);
                }
            }
        }

        /// <summary>
        /// 繪製邊框路徑
        /// </summary>
        private void DrawBorderPath(Graphics g, Pen pen, Rectangle rect, int radius)
        {
            if (this.Rounding && radius > 0)
            {
                using (var path = CreateRoundedRectPath(rect, radius))
                {
                    g.DrawPath(pen, path);
                }
            }
            else
            {
                g.DrawRectangle(pen, rect);
            }
        }

        /// <summary>
        /// 獲取安全的圓角半徑
        /// </summary>
        private int GetSafeRadius(Rectangle rect)
        {
            if (!this.Rounding)
                return 0;

            int maxRadius = Math.Min(rect.Width, rect.Height) / 2 - 1;
            maxRadius = Math.Max(0, maxRadius);
            
            return Math.Min(this.RoundingInt, maxRadius);
        }

        /// <summary>
        /// 創建圓角矩形路徑
        /// </summary>
        private GraphicsPath CreateRoundedRectPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();

            if (radius <= 0 || rect.Width <= 0 || rect.Height <= 0)
            {
                path.AddRectangle(rect.Width > 0 && rect.Height > 0 ? rect : new Rectangle(0, 0, 1, 1));
                return path;
            }

            // 確保直徑不超過矩形尺寸
            int diameter = Math.Min(radius * 2, Math.Min(rect.Width, rect.Height));
            radius = diameter / 2;

            if (diameter < 2)
            {
                path.AddRectangle(rect);
                return path;
            }

            // 左上角
            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            // 右上角
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            // 右下角
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            // 左下角
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);

            path.CloseFigure();
            return path;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            this.Invalidate();
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            this.Invalidate();
        }

        /// <summary>
        /// 攔截 WndProc 以防止某些繪製消息觸發基類錯誤
        /// </summary>
        protected override void WndProc(ref Message m)
        {
            const int WM_PAINT = 0x000F;
            const int WM_ERASEBKGND = 0x0014;
            const int WM_PRINTCLIENT = 0x0318;

            try
            {
                base.WndProc(ref m);
            }
            catch (ArgumentException)
            {
                // 忽略繪製相關的異常
                if (m.Msg == WM_PAINT || m.Msg == WM_ERASEBKGND || m.Msg == WM_PRINTCLIENT)
                {
                    return;
                }
                throw;
            }
            catch (InvalidOperationException)
            {
                if (m.Msg == WM_PAINT || m.Msg == WM_ERASEBKGND || m.Msg == WM_PRINTCLIENT)
                {
                    return;
                }
                throw;
            }
        }
    }
}
