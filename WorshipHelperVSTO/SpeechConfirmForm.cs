// ============================================================================
// SpeechConfirmForm.cs  —  v5.1 (modernised)
//
// A premium-looking toast notification that appears when speech recognition
// detects a Bible reference. Users can edit, accept (Enter / Insert button),
// or dismiss (Esc / Dismiss button). Auto-dismisses after a timeout.
//
// What's new in v5.1:
//   • Uses the new ModernButton + Palette from UI/ModernControls.cs so the
//     Insert / Dismiss buttons are genuinely rounded (not just square with
//     a flat border), with smooth hover + press states.
//   • Soft 4-layer shadow behind the card for depth.
//   • Countdown bar is a thin pill at the bottom, fully inside the card's
//     rounded region — previously it touched the card border which looked
//     amateur.
//   • Nothing else in the public API has changed — TestRibbonItem keeps
//     working exactly as before.
// ============================================================================

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WorshipHelperVSTO.UI;

namespace WorshipHelperVSTO
{
    public class SpeechConfirmForm : Form
    {
        // Local aliases for readability.
        private static readonly Color SurfaceTop    = Palette.SurfaceTop;
        private static readonly Color SurfaceBottom = Palette.SurfaceBottom;
        private static readonly Color AccentGreen   = Palette.AccentGreen;
        private static readonly Color AccentGreenLt = Palette.AccentGreenLt;
        private static readonly Color DarkGreen     = Palette.DarkGreen;
        private static readonly Color LightGreen    = Palette.LightGreen;
        private static readonly Color TextPrimary   = Palette.TextDark;
        private static readonly Color TextMuted     = Palette.TextMuted;
        private static readonly Color BorderSoft    = Palette.BorderLight;

        // ── Controls ────────────────────────────────────────────────────────
        private Label        _lblIcon;
        private Label        _lblHeading;
        private Label        _lblSubheading;
        private TextBox      _txtReference;
        private Panel        _txtReferenceWrap;
        private Label        _lblHint;
        private ModernButton _btnInsert;
        private ModernButton _btnDismiss;
        private CountdownBar _countdown;

        // ── Timers / state ──────────────────────────────────────────────────
        private Timer _autoCloseTimer;
        private Timer _countdownTicker;
        private Timer _fadeTimer;

        private int      _autoCloseMs  = 12_000;      // 12s default
        private DateTime _closeAtUtc;
        private bool     _userHasEdited;
        private bool     _closing;

        // Slide-in animation state
        private Point _targetLocation;
        private Point _offscreenLocation;
        private Timer _slideInTimer;
        private int   _slideInStep;
        private const int SlideInSteps = 10;

        // Shadow margin used when painting the card.
        private const int ShadowMargin = 8;

        // ── Public API ──────────────────────────────────────────────────────
        public string Reference => _txtReference?.Text?.Trim() ?? string.Empty;

        /// <summary>Kept for backwards compatibility with old ShowDialog() callers.</summary>
        public bool Confirmed { get; private set; }

        /// <summary>Fired when the user accepts the reference (Enter / Insert click).</summary>
        public event EventHandler<string> ReferenceConfirmed;

        /// <summary>Fired when the user dismisses (Esc / Dismiss / timeout).</summary>
        public event EventHandler Dismissed;

        public SpeechConfirmForm(string detectedReference)
        {
            BuildUI(detectedReference ?? "");
            StartAutoClose(_autoCloseMs);
        }

        // ─────────────────────────────────────────────────────────────────────
        // UI construction
        // ─────────────────────────────────────────────────────────────────────
        private void BuildUI(string reference)
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition   = FormStartPosition.Manual;
            TopMost         = true;
            ShowInTaskbar   = false;
            BackColor       = Color.Magenta;    // transparency key
            TransparencyKey = Color.Magenta;    // lets the soft shadow look shadowy
            DoubleBuffered  = true;
            Size            = new Size(420, 196);
            Padding         = new Padding(0);

            // Anchor bottom-right of primary screen
            var screen = Screen.PrimaryScreen.WorkingArea;
            _targetLocation    = new Point(screen.Right - Width - 24, screen.Bottom - Height - 24);
            _offscreenLocation = new Point(screen.Right + 8,          _targetLocation.Y);
            Location           = _offscreenLocation;

            // Custom paint: shadow + rounded card + gradient + left accent stripe
            Paint += OnPaintCard;

            // Inner card bounds (everything is positioned relative to these).
            Rectangle card = CardBounds();

            // ── Icon badge ───────────────────────────────────────────────
            _lblIcon = new Label
            {
                Text      = "\uD83D\uDCD6",             // 📖
                Font      = new Font("Segoe UI Emoji", 18f, FontStyle.Regular),
                ForeColor = AccentGreen,
                BackColor = Color.Transparent,
                AutoSize  = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Size      = new Size(44, 44),
                Location  = new Point(card.X + 14, card.Y + 14),
            };
            Controls.Add(_lblIcon);

            // ── Heading ──────────────────────────────────────────────────
            _lblHeading = new Label
            {
                Text      = "Heard a reference",
                Font      = new Font("Segoe UI Semibold", 10.5f, FontStyle.Bold),
                ForeColor = DarkGreen,
                BackColor = Color.Transparent,
                AutoSize  = true,
                Location  = new Point(card.X + 64, card.Y + 14),
            };
            Controls.Add(_lblHeading);

            // ── Subheading / instructions ────────────────────────────────
            _lblSubheading = new Label
            {
                Text      = "Press Enter to insert  •  Esc to dismiss",
                Font      = new Font("Segoe UI", 8.25f, FontStyle.Regular),
                ForeColor = TextMuted,
                BackColor = Color.Transparent,
                AutoSize  = true,
                Location  = new Point(card.X + 64, card.Y + 32),
            };
            Controls.Add(_lblSubheading);

            // ── Reference textbox wrapped in a rounded panel ─────────────
            _txtReferenceWrap = new Panel
            {
                Location   = new Point(card.X + 16, card.Y + 62),
                Size       = new Size(card.Width - 32, 38),
                BackColor  = LightGreen,
                Padding    = new Padding(12, 6, 12, 6),
            };
            _txtReferenceWrap.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var rr = new Rectangle(0, 0, _txtReferenceWrap.Width - 1, _txtReferenceWrap.Height - 1);
                using (var path = Palette.RoundedPath(rr, 10))
                using (var br   = new SolidBrush(LightGreen))
                    g.FillPath(br, path);
            };
            Controls.Add(_txtReferenceWrap);

            _txtReference = new TextBox
            {
                Text        = reference,
                Font        = new Font("Segoe UI", 15f, FontStyle.Bold),
                ForeColor   = TextPrimary,
                BackColor   = LightGreen,
                BorderStyle = BorderStyle.None,
                Dock        = DockStyle.Fill,
                TextAlign   = HorizontalAlignment.Left,
            };
            _txtReference.KeyDown     += OnReferenceKeyDown;
            _txtReference.TextChanged += OnTextChangedFlagUserEdit;
            _txtReferenceWrap.Controls.Add(_txtReference);

            // Caret to end on focus
            _txtReference.GotFocus += (s, e) =>
            {
                try { _txtReference.SelectionStart = _txtReference.TextLength; }
                catch { /* ignore */ }
            };

            // ── Countdown bar ────────────────────────────────────────────
            _countdown = new CountdownBar
            {
                Location = new Point(card.X + 16, card.Y + 108),
                Size     = new Size(card.Width - 32, 4),
            };
            Controls.Add(_countdown);

            // ── Hint ─────────────────────────────────────────────────────
            _lblHint = new Label
            {
                Text      = "Tip: click the box to edit",
                Font      = new Font("Segoe UI", 8f, FontStyle.Italic),
                ForeColor = TextMuted,
                BackColor = Color.Transparent,
                AutoSize  = true,
                Location  = new Point(card.X + 16, card.Y + 118),
            };
            Controls.Add(_lblHint);

            // ── Modern rounded buttons ───────────────────────────────────
            _btnDismiss = new ModernButton
            {
                Text    = "Dismiss",
                Primary = false,
                Size    = new Size(88, 32),
                TabStop = false,
            };
            _btnDismiss.Click += (s, e) => Dismiss();
            Controls.Add(_btnDismiss);

            _btnInsert = new ModernButton
            {
                Text    = "Insert  \u2192",
                Primary = true,
                Size    = new Size(116, 32),
                TabStop = false,
            };
            _btnInsert.Click += (s, e) => Confirm();
            Controls.Add(_btnInsert);

            LayoutButtons();
        }

        private Rectangle CardBounds()
        {
            // Leave room for the shadow on all sides.
            return new Rectangle(ShadowMargin, ShadowMargin,
                                 Width - ShadowMargin * 2,
                                 Height - ShadowMargin * 2);
        }

        private void LayoutButtons()
        {
            if (_btnDismiss == null || _btnInsert == null) return;
            Rectangle card = CardBounds();
            int y = card.Bottom - _btnInsert.Height - 14;
            _btnInsert.Location  = new Point(card.Right - _btnInsert.Width - 14, y);
            _btnDismiss.Location = new Point(_btnInsert.Left - _btnDismiss.Width - 8, y);
        }

        private void OnPaintCard(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode   = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            // Clear the form with the transparency key so we actually get
            // rounded-with-shadow instead of showing square magenta edges.
            g.Clear(TransparencyKey);

            Rectangle card = CardBounds();

            // Soft drop shadow (4 expanding layers of low-alpha fills)
            for (int i = 1; i <= 4; i++)
            {
                var sr = new Rectangle(card.X - i, card.Y + i, card.Width + i * 2, card.Height + i * 2);
                using (var path = Palette.RoundedPath(sr, 16))
                using (var br   = new SolidBrush(Color.FromArgb(14 - i * 2, 0, 0, 0)))
                    g.FillPath(br, path);
            }

            // Card body with gradient
            using (var path = Palette.RoundedPath(card, 14))
            {
                using (var br = new LinearGradientBrush(card, SurfaceTop, SurfaceBottom,
                                                        LinearGradientMode.Vertical))
                    g.FillPath(br, path);

                // Left accent stripe, clipped to rounded path
                var oldClip = g.Clip;
                g.SetClip(path);
                using (var accent = new LinearGradientBrush(
                    new Rectangle(card.X, card.Y, 4, card.Height),
                    AccentGreenLt, AccentGreen, LinearGradientMode.Vertical))
                    g.FillRectangle(accent, card.X, card.Y, 4, card.Height);
                g.Clip = oldClip;

                // Soft 1px border
                using (var pen = new Pen(BorderSoft, 1))
                    g.DrawPath(pen, path);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Modeless / activation
        // ─────────────────────────────────────────────────────────────────────
        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_EX_NOACTIVATE = 0x08000000;
                const int WS_EX_TOOLWINDOW  = 0x00000080;
                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            StartSlideIn();
        }

        // ─────────────────────────────────────────────────────────────────────
        // Animation helpers
        // ─────────────────────────────────────────────────────────────────────
        private void StartSlideIn()
        {
            _slideInStep = 0;
            _slideInTimer?.Stop();
            _slideInTimer = new Timer { Interval = 14 };
            _slideInTimer.Tick += (s, e) =>
            {
                _slideInStep++;
                if (_slideInStep >= SlideInSteps)
                {
                    Location = _targetLocation;
                    _slideInTimer.Stop();
                    _slideInTimer.Dispose();
                    _slideInTimer = null;
                    return;
                }
                // easeOutCubic
                float t    = (float)_slideInStep / SlideInSteps;
                float ease = 1f - (float)Math.Pow(1 - t, 3);
                int   dx   = _targetLocation.X - _offscreenLocation.X;
                Location = new Point(_offscreenLocation.X + (int)(dx * ease), _targetLocation.Y);
            };
            _slideInTimer.Start();
        }

        private void StartFadeOut(Action afterFade)
        {
            _fadeTimer?.Stop();
            _fadeTimer = new Timer { Interval = 18 };
            _fadeTimer.Tick += (s, e) =>
            {
                Opacity -= 0.14;
                if (Opacity <= 0.02)
                {
                    Opacity = 0;
                    _fadeTimer.Stop();
                    _fadeTimer.Dispose();
                    _fadeTimer = null;
                    afterFade?.Invoke();
                }
            };
            _fadeTimer.Start();
        }

        // ─────────────────────────────────────────────────────────────────────
        // Public: edit the reference shown in the toast in place
        // ─────────────────────────────────────────────────────────────────────
        public void UpdateReference(string newReference)
        {
            if (IsDisposed || _closing) return;
            if (string.IsNullOrWhiteSpace(newReference)) return;

            if (InvokeRequired)
            {
                try { BeginInvoke((Action)(() => UpdateReference(newReference))); }
                catch { /* tearing down */ }
                return;
            }

            try
            {
                if (!_userHasEdited)
                {
                    _txtReference.TextChanged -= OnTextChangedFlagUserEdit;
                    try
                    {
                        _txtReference.Text = newReference;
                        _txtReference.SelectionStart  = _txtReference.Text.Length;
                        _txtReference.SelectionLength = 0;
                    }
                    finally
                    {
                        _txtReference.TextChanged += OnTextChangedFlagUserEdit;
                    }

                    // Nudge the heading so the user notices the change
                    _lblHeading.Text = "Reference updated";
                }

                RestartAutoClose();
            }
            catch
            {
                // never break the pipeline for a cosmetic update
            }
        }

        private void OnTextChangedFlagUserEdit(object s, EventArgs e) => _userHasEdited = true;

        // ─────────────────────────────────────────────────────────────────────
        // Input handling
        // ─────────────────────────────────────────────────────────────────────
        private void OnReferenceKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                e.Handled          = true;
                Confirm();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                e.Handled          = true;
                Dismiss();
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape) { Dismiss(); return true; }
            if (keyData == Keys.Enter)  { Confirm(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ─────────────────────────────────────────────────────────────────────
        // Confirm / dismiss
        // ─────────────────────────────────────────────────────────────────────
        private void Confirm()
        {
            if (IsDisposed || _closing) return;
            _closing  = true;
            Confirmed = true;
            StopAllTimers();

            string final = Reference;
            try { ReferenceConfirmed?.Invoke(this, final); }
            catch { /* subscriber errors must not block close */ }

            StartFadeOut(() => { try { Close(); } catch { } });
        }

        private void Dismiss()
        {
            if (IsDisposed || _closing) return;
            _closing  = true;
            Confirmed = false;
            StopAllTimers();

            try { Dismissed?.Invoke(this, EventArgs.Empty); }
            catch { /* subscriber errors must not block close */ }

            StartFadeOut(() => { try { Close(); } catch { } });
        }

        // ─────────────────────────────────────────────────────────────────────
        // Auto-close + countdown bar
        // ─────────────────────────────────────────────────────────────────────
        private void StartAutoClose(int milliseconds)
        {
            _autoCloseMs = milliseconds;
            _closeAtUtc  = DateTime.UtcNow.AddMilliseconds(milliseconds);

            _autoCloseTimer = new Timer { Interval = milliseconds };
            _autoCloseTimer.Tick += (s, e) => Dismiss();
            _autoCloseTimer.Start();

            _countdownTicker = new Timer { Interval = 60 };
            _countdownTicker.Tick += (s, e) =>
            {
                try
                {
                    double remaining = (_closeAtUtc - DateTime.UtcNow).TotalMilliseconds;
                    float  progress  = (float)Math.Max(0, Math.Min(1, remaining / _autoCloseMs));
                    _countdown.Progress = progress;
                }
                catch { /* ignore */ }
            };
            _countdownTicker.Start();
        }

        private void RestartAutoClose()
        {
            if (_autoCloseTimer == null) return;
            _autoCloseTimer.Stop();
            _autoCloseTimer.Interval = _autoCloseMs;
            _autoCloseTimer.Start();
            _closeAtUtc = DateTime.UtcNow.AddMilliseconds(_autoCloseMs);
        }

        private void StopAllTimers()
        {
            try { _autoCloseTimer?.Stop();    _autoCloseTimer?.Dispose();    } catch { }
            try { _countdownTicker?.Stop();   _countdownTicker?.Dispose();   } catch { }
            _autoCloseTimer  = null;
            _countdownTicker = null;
        }

        protected override void Dispose(bool disposing)
        {
            StopAllTimers();
            try { _slideInTimer?.Dispose(); } catch { }
            try { _fadeTimer?.Dispose();    } catch { }
            base.Dispose(disposing);
        }

        // ─────────────────────────────────────────────────────────────────────
        // CountdownBar — thin pill under the reference
        // ─────────────────────────────────────────────────────────────────────
        private sealed class CountdownBar : Control
        {
            private float _progress = 1f;
            public float Progress
            {
                get => _progress;
                set
                {
                    float v = value < 0 ? 0 : (value > 1 ? 1 : value);
                    if (Math.Abs(v - _progress) < 0.001f) return;
                    _progress = v;
                    Invalidate();
                }
            }

            public CountdownBar()
            {
                SetStyle(ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.UserPaint |
                         ControlStyles.ResizeRedraw |
                         ControlStyles.SupportsTransparentBackColor, true);
                BackColor = Color.Transparent;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode   = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                var r = new Rectangle(0, 0, Width, Height);
                int radius = Math.Min(r.Height, 4);

                // Track
                using (var path = Palette.RoundedPath(r, radius))
                using (var br   = new SolidBrush(Color.FromArgb(224, 232, 224)))
                    g.FillPath(br, path);

                int fillW = (int)(r.Width * _progress);
                if (fillW <= 0) return;

                // Fill — transitions green → amber → red as time runs out
                Color c = _progress > 0.5f
                    ? AccentGreen
                    : (_progress > 0.25f ? Color.FromArgb(230, 180, 50)
                                         : Color.FromArgb(220, 80, 60));

                var fillRect = new Rectangle(0, 0, fillW, Height);
                using (var path = Palette.RoundedPath(fillRect, radius))
                using (var br   = new LinearGradientBrush(
                    fillRect, Color.FromArgb(210, c), c, LinearGradientMode.Horizontal))
                    g.FillPath(br, path);
            }
        }

        // (DllImports previously used for rounded region now live in Palette.)
    }
}
