using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// 區域框選工具 - 讓使用者在截圖上拖拽框選區域
    /// </summary>
    public class RegionSelector : Form
    {
        private Bitmap screenshot;
        private Point startPoint;
        private Point endPoint;
        private bool isSelecting = false;
        private Rectangle selectedRegion = Rectangle.Empty;

        /// <summary>
        /// 使用者選擇的區域 (相對於截圖)
        /// </summary>
        public Rectangle SelectedRegion => selectedRegion;

        /// <summary>
        /// 是否已選擇區域
        /// </summary>
        public bool HasSelection => selectedRegion.Width > 0 && selectedRegion.Height > 0;

        public RegionSelector(Bitmap screenshot)
        {
            this.screenshot = new Bitmap(screenshot);

            // 設定表單屬性
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.DoubleBuffered = true;
            this.BackColor = Color.Black;
            this.Cursor = Cursors.Cross;
            this.TopMost = true;
            this.KeyPreview = true;

            // 計算適當的顯示大小 (最大不超過螢幕的 90%)
            var screen = Screen.PrimaryScreen.WorkingArea;
            float scaleX = (float)(screen.Width * 0.9) / screenshot.Width;
            float scaleY = (float)(screen.Height * 0.9) / screenshot.Height;
            float scale = Math.Min(scaleX, scaleY);
            scale = Math.Min(scale, 1.0f); // 不放大

            this.Width = (int)(screenshot.Width * scale);
            this.Height = (int)(screenshot.Height * scale) + 40; // 額外空間給提示文字

            // 縮放截圖
            if (scale < 1.0f)
            {
                var scaled = new Bitmap(this.Width, (int)(screenshot.Height * scale));
                using (var g = Graphics.FromImage(scaled))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBilinear;
                    g.DrawImage(screenshot, 0, 0, scaled.Width, scaled.Height);
                }
                this.screenshot.Dispose();
                this.screenshot = scaled;
            }

            // 事件
            this.MouseDown += OnMouseDown;
            this.MouseMove += OnMouseMove;
            this.MouseUp += OnMouseUp;
            this.Paint += OnPaint;
            this.KeyDown += OnKeyDown;
        }

        private void OnMouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && e.Y < this.Height - 40)
            {
                startPoint = e.Location;
                endPoint = e.Location;
                isSelecting = true;
                Invalidate();
            }
        }

        private void OnMouseMove(object? sender, MouseEventArgs e)
        {
            if (isSelecting)
            {
                endPoint = e.Location;
                // 限制在截圖範圍內
                endPoint.X = Math.Max(0, Math.Min(endPoint.X, screenshot.Width));
                endPoint.Y = Math.Max(0, Math.Min(endPoint.Y, screenshot.Height));
                Invalidate();
            }
        }

        private void OnMouseUp(object? sender, MouseEventArgs e)
        {
            if (isSelecting && e.Button == MouseButtons.Left)
            {
                isSelecting = false;
                endPoint = e.Location;
                endPoint.X = Math.Max(0, Math.Min(endPoint.X, screenshot.Width));
                endPoint.Y = Math.Max(0, Math.Min(endPoint.Y, screenshot.Height));

                // 計算選擇區域
                int x = Math.Min(startPoint.X, endPoint.X);
                int y = Math.Min(startPoint.Y, endPoint.Y);
                int w = Math.Abs(endPoint.X - startPoint.X);
                int h = Math.Abs(endPoint.Y - startPoint.Y);

                if (w > 5 && h > 5)
                {
                    selectedRegion = new Rectangle(x, y, w, h);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    Invalidate();
                }
            }
        }

        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void OnPaint(object? sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.InterpolationMode = InterpolationMode.HighQualityBilinear;

            // 繪製截圖
            g.DrawImage(screenshot, 0, 0);

            // 繪製半透明遮罩
            using (var overlay = new SolidBrush(Color.FromArgb(100, 0, 0, 0)))
            {
                g.FillRectangle(overlay, 0, 0, screenshot.Width, screenshot.Height);
            }

            // 如果正在選擇，繪製選擇區域
            if (isSelecting || selectedRegion != Rectangle.Empty)
            {
                Rectangle rect;
                if (isSelecting)
                {
                    int x = Math.Min(startPoint.X, endPoint.X);
                    int y = Math.Min(startPoint.Y, endPoint.Y);
                    int w = Math.Abs(endPoint.X - startPoint.X);
                    int h = Math.Abs(endPoint.Y - startPoint.Y);
                    rect = new Rectangle(x, y, w, h);
                }
                else
                {
                    rect = selectedRegion;
                }

                // 在選擇區域內顯示原始圖片（清除遮罩）
                if (rect.Width > 0 && rect.Height > 0)
                {
                    g.SetClip(rect);
                    g.DrawImage(screenshot, 0, 0);
                    g.ResetClip();

                    // 繪製選擇框
                    using (var pen = new Pen(Color.Lime, 2))
                    {
                        pen.DashStyle = DashStyle.Dash;
                        g.DrawRectangle(pen, rect);
                    }

                    // 顯示尺寸
                    string sizeText = $"{rect.Width} x {rect.Height}";
                    using (var font = new Font("Microsoft JhengHei UI", 10, FontStyle.Bold))
                    using (var brush = new SolidBrush(Color.Yellow))
                    using (var bgBrush = new SolidBrush(Color.FromArgb(180, 0, 0, 0)))
                    {
                        var textSize = g.MeasureString(sizeText, font);
                        var textRect = new RectangleF(rect.X, rect.Y - textSize.Height - 5, textSize.Width + 10, textSize.Height + 4);
                        if (textRect.Y < 0) textRect.Y = rect.Bottom + 5;
                        g.FillRectangle(bgBrush, textRect);
                        g.DrawString(sizeText, font, brush, textRect.X + 5, textRect.Y + 2);
                    }
                }
            }

            // 繪製底部提示
            using (var bgBrush = new SolidBrush(Color.FromArgb(220, 30, 30, 35)))
            {
                g.FillRectangle(bgBrush, 0, this.Height - 40, this.Width, 40);
            }

            string hint = "🖱️ 拖拽框選小地圖區域  |  按 ESC 取消";
            using (var font = new Font("Microsoft JhengHei UI", 11, FontStyle.Bold))
            using (var brush = new SolidBrush(Color.White))
            {
                var textSize = g.MeasureString(hint, font);
                float textX = (this.Width - textSize.Width) / 2;
                g.DrawString(hint, font, brush, textX, this.Height - 32);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                screenshot?.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// 顯示區域選擇器並返回選擇的區域
        /// </summary>
        /// <param name="screenshot">要在其上選擇的截圖</param>
        /// <param name="scale">截圖相對於原始視窗的縮放比例 (用於還原座標)</param>
        /// <returns>選擇的區域 (原始座標)，如果取消則返回 null</returns>
        public static Rectangle? SelectRegion(Bitmap screenshot, float scale = 1.0f)
        {
            using (var selector = new RegionSelector(screenshot))
            {
                if (selector.ShowDialog() == DialogResult.OK && selector.HasSelection)
                {
                    // 如果有縮放，需要還原到原始座標
                    var region = selector.SelectedRegion;
                    if (Math.Abs(scale - 1.0f) > 0.01f)
                    {
                        return new Rectangle(
                            (int)(region.X / scale),
                            (int)(region.Y / scale),
                            (int)(region.Width / scale),
                            (int)(region.Height / scale)
                        );
                    }
                    return region;
                }
            }
            return null;
        }
    }
}
