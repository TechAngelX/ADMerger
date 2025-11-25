// UI/Controls/ModernButton.cs

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ADMerger.UI.Controls
{
    public class ModernButton : Button
    {
        private bool _isSecondary = false;
        private bool _isRounded = false;

        public ModernButton()
        {
            this.FlatStyle = FlatStyle.Flat;
            this.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.Cursor = Cursors.Hand;
            UpdateStyle();
        }

        public void SetSecondary()
        {
            _isSecondary = true;
            UpdateStyle();
        }

        public void SetRounded()
        {
            _isRounded = true;
            this.Paint += ModernButton_Paint;
        }

        private void ModernButton_Paint(object sender, PaintEventArgs e)
        {
            if (_isRounded)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                
                var rect = new Rectangle(0, 0, this.Width - 1, this.Height - 1);
                var path = GetRoundedRectanglePath(rect, 10);
                
                this.Region = new Region(path);
                
                using (var brush = new SolidBrush(this.BackColor))
                {
                    e.Graphics.FillPath(brush, path);
                }
                
                using (var pen = new Pen(this.BackColor, 2))
                {
                    e.Graphics.DrawPath(pen, path);
                }
                
                TextRenderer.DrawText(e.Graphics, this.Text, this.Font, rect, this.ForeColor, 
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
        }

        private GraphicsPath GetRoundedRectanglePath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius, radius, 180, 90);
            path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 90);
            path.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90);
            path.AddArc(rect.X, rect.Bottom - radius, radius, radius, 90, 90);
            path.CloseFigure();
            return path;
        }

        private void UpdateStyle()
        {
            if (_isSecondary)
            {
                this.BackColor = Color.White;
                this.ForeColor = ColorTranslator.FromHtml("#3B82F6");
                this.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#3B82F6");
                this.FlatAppearance.BorderSize = 2;
                this.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#EFF6FF");
            }
            else
            {
                this.BackColor = ColorTranslator.FromHtml("#3B82F6");
                this.ForeColor = Color.White;
                this.FlatAppearance.BorderSize = 0;
                this.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#2563EB");
            }
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            if (!Enabled && !_isSecondary)
            {
                this.BackColor = ColorTranslator.FromHtml("#CBD5E1");
                this.ForeColor = ColorTranslator.FromHtml("#94A3B8");
            }
            else
            {
                UpdateStyle();
            }
        }
    }
}