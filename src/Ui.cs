using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace MailAktarici
{
    // Görünüm: renkler, yazı tipleri ve küçük özel denetimler.
    static class Theme
    {
        public static readonly Color Accent = Color.FromArgb(28, 94, 168);
        public static readonly Color AccentDark = Color.FromArgb(21, 72, 130);
        public static readonly Color AccentSoft = Color.FromArgb(237, 243, 251);
        public static readonly Color AccentDisabled = Color.FromArgb(170, 191, 216);
        public static readonly Color Bg = Color.FromArgb(242, 244, 247);
        public static readonly Color Surface = Color.White;
        public static readonly Color Border = Color.FromArgb(220, 226, 233);
        public static readonly Color Text = Color.FromArgb(30, 38, 48);
        public static readonly Color Muted = Color.FromArgb(98, 109, 123);
        public static readonly Color Warn = Color.FromArgb(178, 84, 0);
        public static readonly Color LogBg = Color.FromArgb(248, 250, 252);

        // Ekran ölçeği (%125, %150 ...): elle çizilen denetimlerde koordinatlar bununla büyütülür.
        public static float Scale = 1f;
        public static int S(int px) { return (int)Math.Round(px * Scale); }

        public static readonly Font Base = new Font("Segoe UI", 9.5f);
        public static readonly Font Small = new Font("Segoe UI", 8.75f);
        public static readonly Font Bold = new Font("Segoe UI Semibold", 9.5f);
        public static readonly Font Title = new Font("Segoe UI Semibold", 11.5f);
        public static readonly Font Mono = new Font("Consolas", 8.75f);

        static Icon appIcon;

        public static Icon AppIcon(int size)
        {
            try
            {
                using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("MailAktarici.app.ico"))
                    if (s != null) return new Icon(s, size, size);
            }
            catch { }
            if (appIcon == null)
            {
                try { appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
                catch { appIcon = SystemIcons.Application; }
            }
            return appIcon;
        }

        public static Button Primary(string text, EventHandler click)
        {
            var b = BaseButton(text, click);
            b.BackColor = Accent;
            b.ForeColor = Color.White;
            b.Font = new Font("Segoe UI Semibold", 10.5f);
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = AccentDark;
            b.FlatAppearance.MouseDownBackColor = AccentDark;
            b.Padding = new Padding(24, 4, 24, 4);
            b.EnabledChanged += (s, e) => b.BackColor = b.Enabled ? Accent : AccentDisabled;
            return b;
        }

        public static Button Secondary(string text, EventHandler click)
        {
            var b = BaseButton(text, click);
            b.BackColor = Surface;
            b.ForeColor = Text;
            b.FlatAppearance.BorderColor = Color.FromArgb(198, 207, 219);
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.MouseOverBackColor = AccentSoft;
            b.FlatAppearance.MouseDownBackColor = Color.FromArgb(222, 233, 247);
            return b;
        }

        static Button BaseButton(string text, EventHandler click)
        {
            var b = new Button
            {
                Text = text, FlatStyle = FlatStyle.Flat, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8, 0, 8, 0), Margin = new Padding(0, 0, 8, 0), Cursor = Cursors.Hand, UseVisualStyleBackColor = false
            };
            b.Click += click;
            return b;
        }

        public static Label Hint(string text)
        {
            return new Label { Text = text, AutoSize = true, ForeColor = Muted, Font = Small, Margin = new Padding(0, 2, 0, 6) };
        }

        public static Label Label(string text)
        {
            return new Label { Text = text, AutoSize = true, ForeColor = Text, Margin = new Padding(0, 7, 8, 0) };
        }

        // İnce kenarlıklı çerçeve (eski 3B kenarlık yerine)
        public static Panel Framed(Control inner, int height)
        {
            var p = new Panel { BackColor = Border, Padding = new Padding(1), Height = height, Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 4) };
            inner.Dock = DockStyle.Fill;
            p.Controls.Add(inner);
            return p;
        }

        public static Control Separator()
        {
            return new Panel { Height = 1, Dock = DockStyle.Fill, BackColor = Color.FromArgb(232, 236, 241), Margin = new Padding(0, 10, 0, 8) };
        }

        public static GraphicsPath Rounded(Rectangle r, int radius)
        {
            var p = new GraphicsPath();
            int d = radius * 2;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        public static void OpenUrl(string url)
        {
            try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
            catch { }
        }
    }

    // Köşeleri yuvarlak beyaz kart
    sealed class Card : Panel
    {
        public Card()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
            BackColor = Theme.Bg;
            Padding = new Padding(18, 14, 18, 14);
            Margin = new Padding(0, 0, 0, 12);
            Dock = DockStyle.Fill;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            if (Width < 4 || Height < 4) return;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = Theme.Rounded(r, Theme.S(10)))
            using (var fill = new SolidBrush(Theme.Surface))
            using (var pen = new Pen(Theme.Border))
            {
                e.Graphics.FillPath(fill, path);
                e.Graphics.DrawPath(pen, path);
            }
        }
    }

    // "1  Nereden alınacak?" gibi numaralı bölüm başlığı
    sealed class StepHeader : Control
    {
        readonly int number;
        readonly string title, subtitle;

        public StepHeader(int number, string title, string subtitle)
        {
            this.number = number;
            this.title = title;
            this.subtitle = subtitle;
            DoubleBuffered = true;
            ResizeRedraw = true;
            BackColor = Theme.Surface;
            Height = string.IsNullOrEmpty(subtitle) ? 30 : 42;
            Dock = DockStyle.Fill;
            Margin = new Padding(0, 0, 0, 6);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            int x = 0;
            if (number > 0)
            {
                int d = Theme.S(26);
                var circle = new Rectangle(0, Theme.S(2), d, d);
                using (var b = new SolidBrush(Theme.Accent)) g.FillEllipse(b, circle);
                TextRenderer.DrawText(g, number.ToString(), Theme.Bold, circle, Color.White,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                x = d + Theme.S(10);
            }
            TextRenderer.DrawText(g, title, Theme.Title, new Point(x, 0), Theme.Text);
            if (!string.IsNullOrEmpty(subtitle))
                TextRenderer.DrawText(g, subtitle, Theme.Small, new Rectangle(x + Theme.S(1), Theme.S(22), Math.Max(10, Width - x), Theme.S(18)), Theme.Muted,
                    TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }
    }

    // Kalın başlıklı onay kutusu + altında açıklama
    sealed class OptionRow : TableLayoutPanel
    {
        public readonly CheckBox Box;
        public Label Description;

        public OptionRow(string title, string description, bool on)
        {
            ColumnCount = 1;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Dock = DockStyle.Fill;
            BackColor = Theme.Surface;
            Margin = new Padding(0, 2, 0, 4);
            Box = new CheckBox { Text = title, Checked = on, AutoSize = true, Font = Theme.Bold, ForeColor = Theme.Text, Margin = new Padding(0, 4, 0, 0), Cursor = Cursors.Hand };
            Controls.Add(Box);
            if (!string.IsNullOrEmpty(description))
            {
                var d = Theme.Hint(description);
                d.Margin = new Padding(18, 0, 0, 1);
                d.Cursor = Cursors.Hand;
                d.Click += (s, e) => { if (Box.Enabled) Box.Checked = !Box.Checked; };
                Controls.Add(d);
                Description = d;
            }
        }
    }

    // Mavi başlık bandı
    sealed class HeaderBand : Control
    {
        readonly Icon icon;

        public HeaderBand()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
            BackColor = Theme.Accent;
            Height = 60;
            Dock = DockStyle.Fill;
            Margin = Padding.Empty;
            icon = Theme.AppIcon(48);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            if (Width <= 0 || Height <= 0) return;
            using (var br = new LinearGradientBrush(ClientRectangle, Theme.Accent, Color.FromArgb(19, 70, 132), LinearGradientMode.Horizontal))
                g.FillRectangle(br, ClientRectangle);
            int x = Theme.S(18), size = Theme.S(38);
            if (icon != null)
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawIcon(icon, new Rectangle(x, (Height - size) / 2, size, size));
                x += size + Theme.S(12);
            }
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            using (var big = new Font("Segoe UI Semibold", 14f))
                TextRenderer.DrawText(g, "Mail Aktarıcı", big, new Point(x, Theme.S(6)), Color.White);
            TextRenderer.DrawText(g, "Outlook maillerinin tamamını tek klasöre kopyalar · hiçbir mail silinmez veya değiştirilmez",
                Theme.Small, new Point(x + Theme.S(2), Theme.S(34)), Color.FromArgb(214, 228, 245));
        }
    }
}
