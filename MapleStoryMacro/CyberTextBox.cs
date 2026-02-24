using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace MapleStoryMacro
{
    /// <summary>
    /// Cyber 風格的深色顯示框（只讀文字顯示）
    /// 配合 ReaLTaiizor Cyber 主題
    /// </summary>
    public class CyberTextBox : UserControl
    {
        private string _text = "";
        private Color _borderColor = Color.FromArgb(29, 200, 238);
        private Color _backgroundColor = Color.FromArgb(35, 35, 40);
        private int _borderRadius = 12;
        private int _borderWidth = 2;

        public new string Text
        {
            get => _text;
            set { _text = value; Invalidate(); TextChanged?.Invoke(this, EventArgs.Empty); }
        }

        public Color BorderColor
        {
            get => _borderColor;
            set { _borderColor = value; Invalidate(); }
        }

        public Color BackgroundColor
        {
            get => _backgroundColor;
            set { _backgroundColor = value; Invalidate(); }
        }

        public int BorderRadius
        {
            get => _borderRadius;
            set { _borderRadius = value; Invalidate(); }
        }

        public int BorderWidth
        {
            get => _borderWidth;
            set { _borderWidth = value; Invalidate(); }
        }

        // 保持 TextBox 兼容性的屬性
        public bool ReadOnly { get; set; } = true;
        public HorizontalAlignment TextAlign { get; set; } = HorizontalAlignment.Center;

        public new event EventHandler? TextChanged;

        public CyberTextBox()
        {
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw |
                ControlStyles.SupportsTransparentBackColor,
                true);

            this.DoubleBuffered = true;
            this.Size = new Size(65, 35);
            this.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Bold);
            this.ForeColor = Color.White;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            Rectangle rect = new Rectangle(0, 0, Width - 1, Height - 1);

            // 繪製背景
            using (var bgBrush = new SolidBrush(_backgroundColor))
            {
                using (var path = CreateRoundedRectangle(rect, _borderRadius))
                {
                    g.FillPath(bgBrush, path);
                }
            }

            // 繪製邊框
            Rectangle borderRect = new Rectangle(
                _borderWidth / 2, 
                _borderWidth / 2,
                Width - _borderWidth - 1, 
                Height - _borderWidth - 1);

            using (var borderPen = new Pen(_borderColor, _borderWidth))
            {
                using (var path = CreateRoundedRectangle(borderRect, _borderRadius - _borderWidth))
                {
                    g.DrawPath(borderPen, path);
                }
            }

            // 繪製文字
            if (!string.IsNullOrEmpty(_text))
            {
                StringFormat sf = new StringFormat();
                sf.Alignment = TextAlign == HorizontalAlignment.Center ? StringAlignment.Center :
                               TextAlign == HorizontalAlignment.Right ? StringAlignment.Far : StringAlignment.Near;
                sf.LineAlignment = StringAlignment.Center;

                using (var textBrush = new SolidBrush(ForeColor))
                {
                    g.DrawString(_text, Font, textBrush, rect, sf);
                }
            }
        }

        private GraphicsPath CreateRoundedRectangle(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();

            if (rect.Width <= 0 || rect.Height <= 0)
            {
                path.AddRectangle(new Rectangle(0, 0, 1, 1));
                return path;
            }

            radius = Math.Max(1, Math.Min(radius, Math.Min(rect.Width, rect.Height) / 2));
            int diameter = radius * 2;

            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();

            return path;
        }
    }
}
