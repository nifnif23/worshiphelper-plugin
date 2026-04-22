// ============================================================================
// AutoScriptureMode.cs  —  v2
// Manages the "Auto Scripture" listening mode.
//
// When active, any scripture reference detected from speech is inserted
// immediately into the presentation — no button press, no form, no
// confirmation dialog. The presenter just says the reference aloud and
// the slide appears.
//
// v2 changes:
//   • Upgraded the inline toast to match the new SpeechConfirmForm look:
//     rounded card, left accent stripe, icon badge, slide-in animation.
//   • Confined toast lifetime is bullet-proof — a new toast cancels any
//     previous one cleanly (no overlap / memory leak).
//   • Shows a short-lived "Inserted X" toast after auto-inserts.
// ============================================================================

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using log4net;

namespace WorshipHelperVSTO
{
    public sealed class AutoScriptureMode
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(AutoScriptureMode));
        private static readonly AutoScriptureMode _instance = new AutoScriptureMode();
        public static AutoScriptureMode Instance => _instance;

        private AutoScriptureMode() { }

        public bool IsEnabled { get; private set; }

        public event EventHandler<bool> StateChanged;

        public void Enable(SpeechToScriptureService service)
        {
            if (IsEnabled) return;
            IsEnabled = true;
            log.Info("AutoScriptureMode: Enabled.");

            if (service != null && !service.IsListening)
            {
                try { service.Start(); }
                catch (Exception ex)
                {
                    log.Error("AutoScriptureMode: Failed to start speech service.", ex);
                    IsEnabled = false;
                    RaiseStateChanged(false);
                    return;
                }
            }

            ShowToast("Auto Scripture", "Listening for spoken references", ToastKind.Info, 2500);
            RaiseStateChanged(true);
        }

        public void Disable()
        {
            if (!IsEnabled) return;
            IsEnabled = false;
            log.Info("AutoScriptureMode: Disabled.");
            ShowToast("Auto Scripture", "Turned off", ToastKind.Muted, 1500);
            RaiseStateChanged(false);
        }

        private void RaiseStateChanged(bool newState)
        {
            try { StateChanged?.Invoke(this, newState); }
            catch (Exception ex) { log.Debug($"AutoScriptureMode: StateChanged subscriber threw: {ex.Message}"); }
        }

        public bool Toggle(SpeechToScriptureService service)
        {
            if (IsEnabled) { Disable(); return false; }
            else { Enable(service); return true; }
        }

        public void HandleDetectedReference(string normalisedReference,
                                            string spokenText,
                                            Action<string> insertAction)
        {
            if (!IsEnabled) return;

            log.Info($"AutoScriptureMode: Auto-inserting \"{normalisedReference}\" (spoken: \"{spokenText}\")");

            try
            {
                insertAction(normalisedReference);
                ShowToast("Inserted", normalisedReference, ToastKind.Success, 2200);
            }
            catch (Exception ex)
            {
                log.Error($"AutoScriptureMode: Insert failed for \"{normalisedReference}\"", ex);
                ShowToast("Insert failed", normalisedReference, ToastKind.Error, 2500);
            }
        }

        // ──────────────────────────────────────────────────────────────────
        // Toast notification (re-skinned to match the new SpeechConfirmForm)
        // ──────────────────────────────────────────────────────────────────
        private Form _toastForm;

        private enum ToastKind { Info, Success, Error, Muted }

        private void ShowToast(string heading, string detail, ToastKind kind, int durationMs = 2000)
        {
            try
            {
                // Replace any existing toast
                try { _toastForm?.Close(); } catch { }
                _toastForm = null;

                var toast = new MiniToast(heading, detail, kind);
                var timer = new Timer { Interval = durationMs };
                timer.Tick += (s, e) =>
                {
                    timer.Stop(); timer.Dispose();
                    try { toast.BeginFadeOutAndClose(); } catch { }
                };
                toast.Shown += (s, e) => timer.Start();
                toast.FormClosed += (s, e) =>
                {
                    if (ReferenceEquals(_toastForm, s)) _toastForm = null;
                };

                toast.Show();
                _toastForm = toast;
            }
            catch (Exception ex)
            {
                log.Debug($"AutoScriptureMode: Toast failed (non-fatal): {ex.Message}");
            }
        }

        // ──────────────────────────────────────────────────────────────────
        // MiniToast — compact, self-contained, styled to match SpeechConfirmForm
        // ──────────────────────────────────────────────────────────────────
        private sealed class MiniToast : Form
        {
            private static readonly Color SurfaceTop     = Color.FromArgb(255, 255, 255);
            private static readonly Color SurfaceBottom  = Color.FromArgb(248, 250, 248);
            private static readonly Color BorderSoft     = Color.FromArgb(220, 224, 220);
            private static readonly Color TextPrimary    = Color.FromArgb(33, 33, 33);
            private static readonly Color TextMuted      = Color.FromArgb(117, 117, 117);

            private static readonly Color AccentInfo     = Color.FromArgb(46, 125, 50);
            private static readonly Color AccentSuccess  = Color.FromArgb(56, 142, 60);
            private static readonly Color AccentError    = Color.FromArgb(198, 40, 40);
            private static readonly Color AccentMuted    = Color.FromArgb(117, 117, 117);

            private readonly Color _accent;
            private readonly string _icon;
            private readonly string _heading;
            private readonly string _detail;

            // Slide-in animation state
            private Point _target;
            private Point _start;
            private Timer _slide;
            private int   _slideStep;
            private const int SlideSteps = 10;

            private Timer _fade;

            public MiniToast(string heading, string detail, ToastKind kind)
            {
                _heading = heading ?? "";
                _detail  = detail ?? "";
                switch (kind)
                {
                    case ToastKind.Success: _accent = AccentSuccess; _icon = "\u2714"; break;
                    case ToastKind.Error:   _accent = AccentError;   _icon = "\u26A0"; break;
                    case ToastKind.Muted:   _accent = AccentMuted;   _icon = "\u25CF"; break;
                    default:                _accent = AccentInfo;    _icon = "\uD83C\uDFA4"; break;
                }

                FormBorderStyle = FormBorderStyle.None;
                StartPosition   = FormStartPosition.Manual;
                TopMost         = true;
                ShowInTaskbar   = false;
                DoubleBuffered  = true;
                BackColor       = SurfaceTop;
                Size            = new Size(360, 64);

                var sc = Screen.PrimaryScreen.WorkingArea;
                _target = new Point(sc.Right - Width - 24, sc.Bottom - Height - 24);
                _start  = new Point(sc.Right + 8,          _target.Y);
                Location = _start;

                ApplyRoundedRegion();
                Resize += (s, e) => ApplyRoundedRegion();
                Paint  += OnPaintCard;
            }

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
                _slideStep = 0;
                _slide = new Timer { Interval = 14 };
                _slide.Tick += (s, ev) =>
                {
                    _slideStep++;
                    if (_slideStep >= SlideSteps)
                    {
                        Location = _target;
                        _slide.Stop(); _slide.Dispose(); _slide = null;
                        return;
                    }
                    float t = (float)_slideStep / SlideSteps;
                    float ease = 1f - (float)Math.Pow(1 - t, 3);
                    int dx = _target.X - _start.X;
                    Location = new Point(_start.X + (int)(dx * ease), _target.Y);
                };
                _slide.Start();
            }

            public void BeginFadeOutAndClose()
            {
                if (IsDisposed) return;
                _fade?.Stop();
                _fade = new Timer { Interval = 20 };
                _fade.Tick += (s, e) =>
                {
                    Opacity -= 0.14;
                    if (Opacity <= 0.02)
                    {
                        Opacity = 0;
                        _fade.Stop(); _fade.Dispose(); _fade = null;
                        try { Close(); } catch { }
                    }
                };
                _fade.Start();
            }

            private void ApplyRoundedRegion()
            {
                try
                {
                    var rgn = CreateRoundRectRgn(0, 0, Width + 1, Height + 1, 12, 12);
                    Region = Region.FromHrgn(rgn);
                    DeleteObject(rgn);
                }
                catch { /* non-fatal */ }
            }

            private void OnPaintCard(object sender, PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                var rect = new Rectangle(0, 0, Width, Height);
                using (var br = new LinearGradientBrush(rect, SurfaceTop, SurfaceBottom, LinearGradientMode.Vertical))
                    g.FillRectangle(br, rect);

                // Left accent stripe
                using (var br = new SolidBrush(_accent))
                    g.FillRectangle(br, 0, 0, 4, Height);

                // Icon
                using (var fnt = new Font("Segoe UI Emoji", 14f, FontStyle.Regular))
                using (var br  = new SolidBrush(_accent))
                    g.DrawString(_icon, fnt, br, 14, 18);

                // Heading
                using (var fnt = new Font("Segoe UI Semibold", 10f, FontStyle.Bold))
                using (var br  = new SolidBrush(TextPrimary))
                    g.DrawString(_heading, fnt, br, 50, 10);

                // Detail
                using (var fnt = new Font("Segoe UI", 9f, FontStyle.Regular))
                using (var br  = new SolidBrush(TextMuted))
                    g.DrawString(_detail, fnt, br, 50, 30);

                using (var pen = new Pen(BorderSoft, 1))
                    g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            }

            [DllImport("Gdi32.dll")]
            private static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect,
                int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);

            [DllImport("Gdi32.dll")]
            private static extern bool DeleteObject(IntPtr hObject);

            protected override void Dispose(bool disposing)
            {
                try { _slide?.Dispose(); } catch { }
                try { _fade?.Dispose();  } catch { }
                base.Dispose(disposing);
            }
        }
    }
}
