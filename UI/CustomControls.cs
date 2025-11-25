using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ADMerger.UI
{
    public class ModernFilePanel : Panel
    {
        private Label label;
        private bool fileLoaded = false;
        private Action<string> onFileDropped;

        public ModernFilePanel(string text, int xPos, int yPos)
        {
            this.Location = new Point(xPos, yPos);
            this.Size = new Size(830, 90);
            this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
            this.BorderStyle = BorderStyle.None;
            this.Cursor = Cursors.Hand;
            this.AllowDrop = true;
            
            this.Paint += (s, e) =>
            {
                using (Pen pen = new Pen(ColorTranslator.FromHtml("#CBD5E1"), 2))
                {
                    pen.DashStyle = DashStyle.Dash;
                    e.Graphics.DrawRectangle(pen, 1, 1, this.Width - 3, this.Height - 3);
                }
            };

            label = new Label
            {
                Text = text,
                Font = new Font("Segoe UI", 11F),
                ForeColor = ColorTranslator.FromHtml("#64748B"),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Cursor = Cursors.Hand,
                BackColor = Color.Transparent
            };
            
            label.Click += (s, e) => this.OnClick(e);
            label.MouseDown += (s, e) => this.OnMouseDown(e);
            label.MouseMove += (s, e) => this.OnMouseMove(e);
            label.MouseUp += (s, e) => this.OnMouseUp(e);
            
            this.Controls.Add(label);

            this.DragEnter += ModernFilePanel_DragEnter;
            this.DragDrop += ModernFilePanel_DragDrop;
            this.DragLeave += ModernFilePanel_DragLeave;
            this.DragOver += ModernFilePanel_DragOver;

            this.MouseEnter += (s, e) =>
            {
                if (this.Enabled && !fileLoaded)
                {
                    this.BackColor = ColorTranslator.FromHtml("#DBEAFE");
                    this.Invalidate();
                }
            };
            
            this.MouseLeave += (s, e) =>
            {
                if (!fileLoaded && this.Enabled)
                {
                    this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                    this.Invalidate();
                }
            };
        }

        private void ModernFilePanel_DragOver(object sender, DragEventArgs e)
        {
            if (!this.Enabled) return;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                e.Effect = (files.Length == 1 && files[0].EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    ? DragDropEffects.Copy : DragDropEffects.None;
            }
        }

        private void ModernFilePanel_DragEnter(object sender, DragEventArgs e)
        {
            if (!this.Enabled) return;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1 && files[0].EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    e.Effect = DragDropEffects.Copy;
                    this.BackColor = ColorTranslator.FromHtml("#BFDBFE");
                    this.Invalidate();
                }
            }
        }

        private void ModernFilePanel_DragLeave(object sender, EventArgs e)
        {
            if (!fileLoaded && this.Enabled)
            {
                this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                this.Invalidate();
            }
        }

        private void ModernFilePanel_DragDrop(object sender, DragEventArgs e)
        {
            if (!this.Enabled) return;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1 && files[0].EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    onFileDropped?.Invoke(files[0]);
            }
            
            if (!fileLoaded && this.Enabled)
            {
                this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                this.Invalidate();
            }
        }

        public void SetDropHandler(Action<string> handler) => onFileDropped = handler;

        public void UpdateText(string text) => label.Text = text;

        public void SetFileLoaded(string filename, int recordCount)
        {
            fileLoaded = true;
            this.BackColor = ColorTranslator.FromHtml("#DCFCE7");
            
            string displayName = filename.Length > 50 ? filename.Substring(0, 47) + "..." : filename;
            label.Text = $"{displayName}\n{recordCount} records loaded";
            label.ForeColor = ColorTranslator.FromHtml("#16A34A");
            label.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            
            this.Invalidate();
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            if (!Enabled)
            {
                this.BackColor = ColorTranslator.FromHtml("#F1F5F9");
                label.ForeColor = ColorTranslator.FromHtml("#94A3B8");
                this.Cursor = Cursors.Default;
                label.Cursor = Cursors.Default;
            }
            else
            {
                this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                label.ForeColor = ColorTranslator.FromHtml("#64748B");
                this.Cursor = Cursors.Hand;
                label.Cursor = Cursors.Hand;
            }
        }
    }

    public class ModernButton : Button
    {
        private bool isSecondary = false;
        private bool isRounded = false;

        public ModernButton()
        {
            this.FlatStyle = FlatStyle.Flat;
            this.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.Cursor = Cursors.Hand;
            UpdateStyle();
        }

        public void SetSecondary()
        {
            isSecondary = true;
            UpdateStyle();
        }

        public void SetRounded()
        {
            isRounded = true;
            this.Paint += ModernButton_Paint;
        }

        private void ModernButton_Paint(object sender, PaintEventArgs e)
        {
            if (isRounded)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                
                var rect = new Rectangle(0, 0, this.Width - 1, this.Height - 1);
                var path = GetRoundedRectanglePath(rect, 10);
                
                this.Region = new Region(path);
                
                using (var brush = new SolidBrush(this.BackColor))
                    e.Graphics.FillPath(brush, path);
                
                using (var pen = new Pen(this.BackColor, 2))
                    e.Graphics.DrawPath(pen, path);
                
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
            if (isSecondary)
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
            if (!Enabled && !isSecondary)
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
