// UI/Controls/ModernFilePanel.cs

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ADMerger.UI.Controls
{
    public class ModernFilePanel : Panel
    {
        private Label _label;
        private string _originalText;
        private bool _fileLoaded = false;
        private Action<string> _onFileDropped;

        public ModernFilePanel(string text, int xPos, int yPos)
        {
            _originalText = text;
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

            _label = new Label();
            _label.Text = text;
            _label.Font = new Font("Segoe UI", 11F);
            _label.ForeColor = ColorTranslator.FromHtml("#64748B");
            _label.TextAlign = ContentAlignment.MiddleCenter;
            _label.Dock = DockStyle.Fill;
            _label.Cursor = Cursors.Hand;
            _label.BackColor = Color.Transparent;
            _label.Click += (s, e) => this.OnClick(e);
            
            _label.MouseDown += (s, e) => this.OnMouseDown(e);
            _label.MouseMove += (s, e) => this.OnMouseMove(e);
            _label.MouseUp += (s, e) => this.OnMouseUp(e);
            
            this.Controls.Add(_label);

            this.DragEnter += ModernFilePanel_DragEnter;
            this.DragDrop += ModernFilePanel_DragDrop;
            this.DragLeave += ModernFilePanel_DragLeave;
            this.DragOver += ModernFilePanel_DragOver;

            this.MouseEnter += (s, e) =>
            {
                if (this.Enabled && !_fileLoaded)
                {
                    this.BackColor = ColorTranslator.FromHtml("#DBEAFE");
                    this.Invalidate();
                }
            };
            this.MouseLeave += (s, e) =>
            {
                if (!_fileLoaded && this.Enabled)
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
                if (files.Length == 1 && files[0].EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
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
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
        }

        private void ModernFilePanel_DragLeave(object sender, EventArgs e)
        {
            if (!_fileLoaded && this.Enabled)
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
                {
                    _onFileDropped?.Invoke(files[0]);
                }
            }
            
            if (!_fileLoaded && this.Enabled)
            {
                this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                this.Invalidate();
            }
        }

        public void SetDropHandler(Action<string> handler)
        {
            _onFileDropped = handler;
        }

        public void UpdateText(string text)
        {
            _label.Text = text;
        }

        public void SetFileLoaded(string filename, int recordCount)
        {
            _fileLoaded = true;
            this.BackColor = ColorTranslator.FromHtml("#DCFCE7");
            
            string displayName = filename.Length > 50 ? filename.Substring(0, 47) + "..." : filename;
            _label.Text = $"{displayName}\n{recordCount} records loaded";
            _label.ForeColor = ColorTranslator.FromHtml("#16A34A");
            _label.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            
            this.Invalidate();
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            if (!Enabled)
            {
                this.BackColor = ColorTranslator.FromHtml("#F1F5F9");
                _label.ForeColor = ColorTranslator.FromHtml("#94A3B8");
                this.Cursor = Cursors.Default;
                _label.Cursor = Cursors.Default;
            }
            else
            {
                this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                _label.ForeColor = ColorTranslator.FromHtml("#64748B");
                this.Cursor = Cursors.Hand;
                _label.Cursor = Cursors.Hand;
            }
        }
    }
}