// ============================================================================
// UI/ModernControls.cs   —   v5.1 UI toolkit
//
// Lightweight custom WinForms controls for the 2026 WorshipHelper UI refresh.
//
// Goals:
//   • Keep the green / white / gold palette of the existing product.
//   • Round all interactive surfaces (buttons, textboxes, cards) to 8 px.
//   • Smooth hover / press transitions without pulling in a heavy UI toolkit.
//   • Drop-in replacements for System.Windows.Forms.Button / TextBox / Panel
//     so existing Designer code can opt in one control at a time.
//
// All controls are GDI+-based and antialiased. Nothing here uses P/Invoke
// except a tiny ApplyRoundRegion helper (already needed for form-level
// rounding in SpeechConfirmForm / AutoScriptureMode).
//
// Thread-safety: these are UI controls — only touch them from the UI thread.
// ============================================================================

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WorshipHelperVSTO.UI
{
    // --------------------------------------------------------------------
    // Shared palette — single source of truth for the product colours.
    // --------------------------------------------------------------------
    public static class Palette
    {
        public static readonly Color AccentGreen   = Color.FromArgb(46, 125, 50);    // #2E7D32
        public static readonly Color AccentGreenLt = Color.FromArgb(76, 175, 80);    // #4CAF50
        public static readonly Color DarkGreen     = Color.FromArgb(27, 94, 32);     // #1B5E20
        public static readonly Color LightGreen    = Color.FromArgb(232, 245, 233);  // #E8F5E9
        public static readonly Color HoverGreen    = Color.FromArgb(200, 230, 201);  // #C8E6C9
        public static readonly Color Gold          = Color.FromArgb(184, 134, 11);   // #B8860B
        public static readonly Color TextDark      = Color.FromArgb(33, 33, 33);     // #212121
        public static readonly Color TextMuted     = Color.FromArgb(117, 117, 117);  // #757575
        public static readonly Color BorderLight   = Color.FromArgb(224, 228, 224);  // soft green-grey
        public static readonly Color SurfaceTop    = Color.White;
        public static readonly Color SurfaceBottom = Color.FromArgb(248, 250, 248);
        public static readonly Color ErrorRed      = Color.FromArgb(198, 40, 40);

        /// <summary>Applies a rounded-rectangle Region to a form/control.</summary>
        public static void ApplyRoundRegion(Control c, int radius = 12)
        {
            if (c == null || c.Width <= 0 || c.Height <= 0) return;
            try
            {
                IntPtr rgn = CreateRoundRectRgn(0, 0, c.Width + 1, c.Height + 1, radius, radius);
                c.Region = Region.FromHrgn(rgn);
                DeleteObject(rgn);
            }
            catch { /* non-fatal — just leaves square corners */ }
        }

        /// <summary>Builds a GraphicsPath for a rounded rectangle. Caller disposes.</summary>
        public static GraphicsPath RoundedPath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            if (radius <= 0 || rect.Width <= 0 || rect.Height <= 0) { path.AddRectangle(rect); return path; }
            int d = radius * 2;
            if (d > rect.Width)  d = rect.Width;
            if (d > rect.Height) d = rect.Height;
            path.AddArc(rect.X,                  rect.Y,                   d, d, 180, 90);
            path.AddArc(rect.Right - d,          rect.Y,                   d, d, 270, 90);
            path.AddArc(rect.Right - d,          rect.Bottom - d,          d, d,   0, 90);
            path.AddArc(rect.X,                  rect.Bottom - d,          d, d,  90, 90);
            path.CloseFigure();
            return path;
        }

        [DllImport("Gdi32.dll")]
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);

        [DllImport("Gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);
    }

    // --------------------------------------------------------------------
    // ModernButton — rounded, antialiased, hover/press tinted.
    //
    // Usage:
    //   var b = new ModernButton { Text = "Insert", Primary = true };
    //   container.Controls.Add(b); b.Click += ...
    //
    // "Primary" = filled accent background (the accept/confirm action).
    // "Primary=false" = outline style with subtle hover fill (cancel etc.).
    // --------------------------------------------------------------------
    public sealed class ModernButton : Button
    {
        private bool _hover;
        private bool _down;

        /// <summary>Corner radius in px.</summary>
        public int CornerRadius { get; set; } = 10;

        /// <summary>Filled accent colour. If false, outline style.</summary>
        public bool Primary { get; set; } = true;

        /// <summary>Overrides the default accent background. Null = use Palette.AccentGreen.</summary>
        public Color? AccentOverride { get; set; }

        public ModernButton()
        {
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Font = new Font("Segoe UI Semibold", 9.5f, FontStyle.Bold);
            Cursor = Cursors.Hand;
            BackColor = Color.Transparent;
            ForeColor = Color.White;
            Size = new Size(120, 34);
            SetStyle(ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.ResizeRedraw |
                     ControlStyles.SupportsTransparentBackColor, true);
            UseVisualStyleBackColor = false;
        }

        protected override void OnMouseEnter(EventArgs e)   { _hover = true;  Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e)   { _hover = false; _down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e){ _down = true;  Invalidate(); base.OnMouseDown(e);  }
        protected override void OnMouseUp(MouseEventArgs e) { _down = false; Invalidate(); base.OnMouseUp(e);    }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode     = SmoothingMode.AntiAlias;
            g.PixelOffsetMode   = PixelOffsetMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = Palette.RoundedPath(r, CornerRadius))
            {
                Color accent = AccentOverride ?? Palette.AccentGreen;
                if (!Enabled) accent = Color.FromArgb(180, 180, 180);

                if (Primary)
                {
                    Color top = accent;
                    Color bot = Shift(accent, -15);
                    if (_down)       { top = Shift(accent, -20); bot = Shift(accent, -30); }
                    else if (_hover) { top = Shift(accent,  +8); bot = accent;              }
                    using (var br = new LinearGradientBrush(r, top, bot, LinearGradientMode.Vertical))
                        g.FillPath(br, path);
                    ForeColor = Color.White;
                }
                else
                {
                    Color fill = _down   ? Palette.HoverGreen
                               : (_hover ? Palette.LightGreen : Color.White);
                    using (var br = new SolidBrush(fill))
                        g.FillPath(br, path);
                    using (var pen = new Pen(_hover ? accent : Palette.BorderLight, 1.2f))
                        g.DrawPath(pen, path);
                    ForeColor = accent;
                }

                // Focus ring
                if (Focused && Enabled)
                {
                    using (var pen = new Pen(Color.FromArgb(80, accent), 1.6f))
                        g.DrawPath(pen, path);
                }

                // Text
                TextRenderer.DrawText(g, Text, Font, r, ForeColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                    TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
            }
        }

        private static Color Shift(Color c, int delta) =>
            Color.FromArgb(c.A,
                Clamp(c.R + delta, 0, 255),
                Clamp(c.G + delta, 0, 255),
                Clamp(c.B + delta, 0, 255));

        private static int Clamp(int v, int lo, int hi) => v < lo ? lo : (v > hi ? hi : v);
    }

    // --------------------------------------------------------------------
    // ModernTextBox — a System.Windows.Forms.TextBox wrapped in a custom
    // Panel that paints a rounded border around it. The text box itself
    // is kept borderless; the panel does the drawing so we get perfect
    // rounding without reimplementing text editing.
    // --------------------------------------------------------------------
    public sealed class ModernTextBox : Panel
    {
        private readonly TextBox _inner;
        private bool _focused;

        /// <summary>Corner radius in px.</summary>
        public int CornerRadius { get; set; } = 8;

        public TextBox InnerTextBox => _inner;

        public override string Text
        {
            get => _inner?.Text ?? string.Empty;
            set { if (_inner != null) _inner.Text = value; }
        }

        public new Font Font
        {
            get => _inner?.Font ?? base.Font;
            set { if (_inner != null) _inner.Font = value; base.Font = value; Relayout(); }
        }

        public HorizontalAlignment TextAlign
        {
            get => _inner?.TextAlign ?? HorizontalAlignment.Left;
            set { if (_inner != null) _inner.TextAlign = value; }
        }

        public bool Multiline
        {
            get => _inner?.Multiline ?? false;
            set { if (_inner != null) _inner.Multiline = value; Relayout(); }
        }

        public bool ReadOnly
        {
            get => _inner?.ReadOnly ?? false;
            set { if (_inner != null) _inner.ReadOnly = value; }
        }

        public event EventHandler TextValueChanged;

        public ModernTextBox()
        {
            BackColor = Color.White;
            Padding   = new Padding(10, 6, 10, 6);
            Size      = new Size(240, 30);
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.ResizeRedraw, true);

            _inner = new TextBox
            {
                BorderStyle = BorderStyle.None,
                Font        = new Font("Segoe UI", 10f),
                BackColor   = Color.White,
                ForeColor   = Palette.TextDark,
            };
            _inner.GotFocus  += (s, e) => { _focused = true;  Invalidate(); };
            _inner.LostFocus += (s, e) => { _focused = false; Invalidate(); };
            _inner.TextChanged += (s, e) => TextValueChanged?.Invoke(this, EventArgs.Empty);
            Controls.Add(_inner);
            Relayout();
            Resize += (s, e) => Relayout();
        }

        private void Relayout()
        {
            if (_inner == null) return;
            int padX = Padding.Left;
            int padY = Padding.Top;
            if (_inner.Multiline)
            {
                _inner.Location = new Point(padX, padY);
                _inner.Size     = new Size(Math.Max(1, Width - padX - Padding.Right),
                                           Math.Max(1, Height - padY - Padding.Bottom));
            }
            else
            {
                // Vertically centre a single-line textbox.
                int h = _inner.PreferredHeight;
                _inner.Location = new Point(padX, Math.Max(padY, (Height - h) / 2));
                _inner.Size     = new Size(Math.Max(1, Width - padX - Padding.Right), h);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode   = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = Palette.RoundedPath(r, CornerRadius))
            {
                using (var br = new SolidBrush(BackColor))
                    g.FillPath(br, path);

                Color border = _focused ? Palette.AccentGreen : Palette.BorderLight;
                float w      = _focused ? 1.6f               : 1f;
                using (var pen = new Pen(border, w))
                    g.DrawPath(pen, path);
            }
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            _inner?.Focus();
        }
    }

    // --------------------------------------------------------------------
    // CardPanel — white rounded container with an optional left accent
    // stripe. Use for the main content area of forms.
    // --------------------------------------------------------------------
    public sealed class CardPanel : Panel
    {
        /// <summary>Corner radius in px.</summary>
        public int CornerRadius { get; set; } = 12;

        /// <summary>Accent stripe colour. Null = no stripe.</summary>
        public Color? AccentStripe { get; set; } = Palette.AccentGreen;

        /// <summary>Stripe width in px.</summary>
        public int StripeWidth { get; set; } = 4;

        /// <summary>Drop shadow?</summary>
        public bool HasShadow { get; set; } = true;

        public CardPanel()
        {
            BackColor = Palette.SurfaceTop;
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.ResizeRedraw |
                     ControlStyles.SupportsTransparentBackColor, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            Rectangle r = HasShadow
                ? new Rectangle(0, 0, Width - 4, Height - 4)
                : new Rectangle(0, 0, Width - 1, Height - 1);

            if (HasShadow)
            {
                // Soft two-layer shadow
                for (int i = 0; i < 4; i++)
                {
                    var sr = new Rectangle(r.X + i, r.Y + i, r.Width, r.Height);
                    using (var path = Palette.RoundedPath(sr, CornerRadius))
                    using (var br   = new SolidBrush(Color.FromArgb(8, 0, 0, 0)))
                        g.FillPath(br, path);
                }
            }

            using (var path = Palette.RoundedPath(r, CornerRadius))
            {
                using (var br = new LinearGradientBrush(r, Palette.SurfaceTop, Palette.SurfaceBottom,
                                                        LinearGradientMode.Vertical))
                    g.FillPath(br, path);

                if (AccentStripe.HasValue && StripeWidth > 0)
                {
                    // Clip to rounded shape so the stripe doesn't overflow
                    var oldClip = g.Clip;
                    g.SetClip(path);
                    using (var br = new LinearGradientBrush(
                        new Rectangle(r.X, r.Y, StripeWidth, r.Height),
                        Palette.AccentGreenLt, AccentStripe.Value,
                        LinearGradientMode.Vertical))
                        g.FillRectangle(br, r.X, r.Y, StripeWidth, r.Height);
                    g.Clip = oldClip;
                }

                using (var pen = new Pen(Palette.BorderLight, 1))
                    g.DrawPath(pen, path);
            }
        }
    }

    // --------------------------------------------------------------------
    // SectionHeader — a tiny all-caps label for "MIC LEVEL" / "CONFIDENCE"
    // style section separators on the debug panel.
    // --------------------------------------------------------------------
    public sealed class SectionHeader : Label
    {
        public SectionHeader()
        {
            AutoSize  = true;
            Font      = new Font("Segoe UI Semibold", 7.8f, FontStyle.Bold);
            ForeColor = Palette.TextMuted;
            BackColor = Color.Transparent;
        }

        public override string Text
        {
            get => base.Text;
            set => base.Text = (value ?? string.Empty).ToUpperInvariant();
        }
    }
}
