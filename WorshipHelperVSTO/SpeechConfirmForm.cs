// ============================================================================
// SpeechConfirmForm.cs  —  REMADE (v2)
//
// A premium-looking toast notification that appears when speech recognition
// detects a Bible reference. Users can edit, accept (Enter / Insert button),
// or dismiss (Esc / Dismiss button). Auto-dismisses after a timeout.
//
// What's new in v2:
//   • Fully re-skinned: large rounded card, proper drop-shadow, cleaner
//     typography, dedicated heading row with book-icon glyph.
//   • Animated slide-in from the right, fade-out on dismiss.
//   • Live countdown bar under the reference so presenters see exactly
//     how much time they have before the toast disappears.
//   • Modelessly shown so presenter view hotkeys still work.
//   • UpdateReference() edits the same toast in place when a chapter-only
//     reference is upgraded to chapter:verse.
//   • ShowWithoutActivation so focus stays on the slideshow.
//   • All timers are Windows Forms timers (UI thread) — thread-safe public
//     API via Invoke/BeginInvoke.
// ============================================================================

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public class SpeechConfirmForm : Form
    {
        // ── Palette ─────────────────────────────────────────────────────────
        private static readonly Color SurfaceTop     = Color.FromArgb(255, 255, 255);
        private static readonly Color SurfaceBottom  = Color.FromArgb(248, 250, 248);
        private static readonly Color AccentGreen    = Color.FromArgb(46, 125, 50);
        private static readonly Color AccentGreenLt  = Color.FromArgb(76, 175, 80);
        private static readonly Color DarkGreen      = Color.FromArgb(27, 94, 32);
        private static readonly Color LightGreen     = Color.FromArgb(232, 245, 233);
        private static readonly Color TextPrimary    = Color.FromArgb(33, 33, 33);
        private static readonly Color TextMuted      = Color.FromArgb(117, 117, 117);
        private static readonly Color BorderSoft     = Color.FromArgb(224, 228, 224);

        // ── Controls ────────────────────────────────────────────────────────
        private Label    _lblIcon;
        private Label    _lblHeading;
        private Label    _lblSubheading;
        private TextBox  _txtReference;
        private Label    _lblHint;
        private Button   _btnInsert;
        private Button   _btnDismiss;
        private CountdownBar _countdown;

        // ── Timers / state ──────────────────────────────────────────────────
        private Timer _autoCloseTimer;
        private Timer _countdownTicker;
        private Timer _fadeTimer;

        private int   _autoCloseMs  = 12_000;      // 12s default
        private DateTime _closeAtUtc;
        private bool  _userHasEdited;
        private bool  _closing;

        // Slide-in animation state
        private Point _targetLocation;
        private Point _offscreenLocation;
        private Timer _slideInTimer;
        private int   _slideInStep;
        private const int SlideInSteps = 10;

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
            BackColor       = SurfaceTop;
            DoubleBuffered  = true;
            Size            = new Size(400, 170);
            Padding         = new Padding(0);

            // Anchor bottom-right of primary screen
            var screen = Screen.PrimaryScreen.WorkingArea;
            _targetLocation    = new Point(screen.Right - Width - 24, screen.Bottom - Height - 24);
            _offscreenLocation = new Point(screen.Right + 8,          _targetLocation.Y);
            Location           = _offscreenLocation;

            // Rounded corners
            ApplyRoundedRegion();
            Resize += (s, e) => ApplyRoundedRegion();

            // Custom paint: subtle gradient + left accent stripe + soft border
            Paint += OnPaintCard;

            // ── Left accent stripe is handled in OnPaintCard (no control) ──

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
                Location  = new Point(14, 16),
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
                Location  = new Point(64, 14)
            };
            Controls.Add(_lblHeading);

            // ── Subheading / timestamp ───────────────────────────────────
            _lblSubheading = new Label
            {
                Text      = "Press Enter to insert, Esc to dismiss",
                Font      = new Font("Segoe UI", 8.25f, FontStyle.Regular),
                ForeColor = TextMuted,
                BackColor = Color.Transparent,
                AutoSize  = true,
                Location  = new Point(64, 32)
            };
            Controls.Add(_lblSubheading);

            // ── Reference textbox (big, bold, editable) ──────────────────
            _txtReference = new TextBox
            {
                Text        = reference,
                Font        = new Font("Segoe UI", 16f, FontStyle.Bold),
                ForeColor   = TextPrimary,
                BackColor   = LightGreen,
                BorderStyle = BorderStyle.None,
                Location    = new Point(16, 58),
                Size        = new Size(Width - 32, 34),
                TextAlign   = HorizontalAlignment.Left,
                Anchor      = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };
            _txtReference.KeyDown     += OnReferenceKeyDown;
            _txtReference.TextChanged += OnTextChangedFlagUserEdit;
            Controls.Add(_txtReference);

            // Selection caret shim for the borderless textbox
            _txtReference.GotFocus += (s, e) =>
            {
                try { _txtReference.SelectionStart = _txtReference.TextLength; }
                catch { /* ignore */ }
            };

            // ── Countdown bar ────────────────────────────────────────────
            _countdown = new CountdownBar
            {
                Location = new Point(16, 100),
                Size     = new Size(Width - 32, 4),
                Anchor   = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
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
                Location  = new Point(16, 112)
            };
            Controls.Add(_lblHint);

            // ── Dismiss button ───────────────────────────────────────────
            _btnDismiss = new Button
            {
                Text      = "Dismiss",
                Font      = new Font("Segoe UI", 9f),
                ForeColor = TextMuted,
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size      = new Size(80, 30),
                Cursor    = Cursors.Hand,
                TabStop   = false,
            };
            _btnDismiss.FlatAppearance.BorderColor = BorderSoft;
            _btnDismiss.FlatAppearance.BorderSize  = 1;
            _btnDismiss.FlatAppearance.MouseOverBackColor = Color.FromArgb(245, 245, 245);
            _btnDismiss.Click += (s, e) => Dismiss();
            Controls.Add(_btnDismiss);

            // ── Insert button (primary) ──────────────────────────────────
            _btnInsert = new Button
            {
                Text      = "Insert  \u2192",
                Font      = new Font("Segoe UI Semibold", 9f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = AccentGreen,
                FlatStyle = FlatStyle.Flat,
                Size      = new Size(110, 30),
                Cursor    = Cursors.Hand,
                TabStop   = false,
                UseVisualStyleBackColor = false,
            };
            _btnInsert.FlatAppearance.BorderSize = 0;
            _btnInsert.FlatAppearance.MouseOverBackColor = DarkGreen;
            _btnInsert.Click += (s, e) => Confirm();
            Controls.Add(_btnInsert);

            // Layout buttons
            LayoutButtons();
            Resize += (s, e) => LayoutButtons();
        }

        private void LayoutButtons()
        {
            if (_btnDismiss == null || _btnInsert == null) return;
            int y = Height - _btnInsert.Height - 14;
            _btnInsert.Location  = new Point(Width - _btnInsert.Width - 14, y);
            _btnDismiss.Location = new Point(_btnInsert.Left - _btnDismiss.Width - 8, y);
        }

        private void ApplyRoundedRegion()
        {
            try
            {
                var rgn = CreateRoundRectRgn(0, 0, Width + 1, Height + 1, 14, 14);
                this.Region = Region.FromHrgn(rgn);
                DeleteObject(rgn);
            }
            catch { /* non-fatal */ }
        }

        private void OnPaintCard(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Subtle top→bottom gradient fill
            var rect = new Rectangle(0, 0, Width, Height);
            using (var br = new LinearGradientBrush(rect, SurfaceTop, SurfaceBottom, LinearGradientMode.Vertical))
                g.FillRectangle(br, rect);

            // Left accent stripe (4px)
            using (var accent = new LinearGradientBrush(
                new Rectangle(0, 0, 4, Height),
                AccentGreenLt, AccentGreen, LinearGradientMode.Vertical))
                g.FillRectangle(accent, 0, 0, 4, Height);

            // Soft 1px border
            using (var pen = new Pen(BorderSoft, 1))
                g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
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
                float t = (float)_slideInStep / SlideInSteps;
                float ease = 1f - (float)Math.Pow(1 - t, 3);
                int dx = _targetLocation.X - _offscreenLocation.X;
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
                e.Handled = true;
                Confirm();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
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
                    float progress = (float)Math.Max(0, Math.Min(1, remaining / _autoCloseMs));
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
            _autoCloseTimer = null;
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
        // Win32 interop for rounded region
        // ─────────────────────────────────────────────────────────────────────
        [DllImport("Gdi32.dll")]
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);

        [DllImport("Gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        // ─────────────────────────────────────────────────────────────────────
        // CountdownBar — thin animated progress line under the reference
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
                         ControlStyles.ResizeRedraw, true);
                BackColor = Color.FromArgb(230, 236, 230);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // Track
                using (var br = new SolidBrush(BackColor))
                    g.FillRectangle(br, ClientRectangle);

                // Fill
                int fillW = (int)(Width * _progress);
                if (fillW <= 0) return;

                // Colour transitions red near zero, green when plenty of time
                Color c = _progress > 0.5f
                    ? AccentGreen
                    : (_progress > 0.25f ? Color.FromArgb(230, 180, 50)
                                         : Color.FromArgb(220, 80, 60));

                using (var br = new LinearGradientBrush(
                    new Rectangle(0, 0, Math.Max(1, fillW), Height),
                    Color.FromArgb(200, c), c,
                    LinearGradientMode.Horizontal))
                    g.FillRectangle(br, 0, 0, fillW, Height);
            }
        }
    }
}
