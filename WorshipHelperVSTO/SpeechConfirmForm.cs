// ============================================================================
// SpeechConfirmForm.cs
//
// A small toast-style popup that appears when speech recognition detects a
// Bible reference. The user can edit the reference and press Enter to insert,
// or Escape to dismiss. Auto-dismisses after a timeout if no interaction.
//
// UPDATED: The toast is now designed to be shown MODELESSLY (Show(), not
// ShowDialog()) so it can live alongside continued speech processing:
//
//   - A preliminary chapter-only reference (e.g. "Genesis 1") can be shown
//     immediately the moment it's heard.
//   - If a richer reference ("Genesis 1:5") is detected a moment later, the
//     caller invokes UpdateReference(...) to edit the same toast in place
//     rather than stacking a second popup on top.
//   - The caller subscribes to ReferenceConfirmed / Dismissed events rather
//     than checking DialogResult.
// ============================================================================

using System;
using System.Drawing;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public class SpeechConfirmForm : Form
    {
        private Label _lblHeading;
        private TextBox _txtReference;
        private Label _lblHint;
        private Button _btnInsert;
        private Button _btnDismiss;
        private Timer _autoCloseTimer;
        private int _autoCloseMs = 12000;

        /// <summary>
        /// True if the user has already interacted with the reference text
        /// (typed in it). When true, automatic UpdateReference calls stop
        /// overwriting the text \u2014 we don't want to clobber the user's edits.
        /// </summary>
        private bool _userHasEdited;

        public string Reference => _txtReference.Text.Trim();

        /// <summary>
        /// Kept for backwards compatibility with the old ShowDialog() flow.
        /// New callers should subscribe to <see cref="ReferenceConfirmed"/>
        /// or <see cref="Dismissed"/> instead.
        /// </summary>
        public bool Confirmed { get; private set; } = false;

        /// <summary>
        /// Raised when the user accepts the shown reference (Enter key or Insert button).
        /// Carries the final (possibly edited) reference string.
        /// </summary>
        public event EventHandler<string> ReferenceConfirmed;

        /// <summary>
        /// Raised when the user dismisses the toast without accepting
        /// (Esc, Dismiss button, or auto-close timeout).
        /// </summary>
        public event EventHandler Dismissed;

        private static readonly Color AccentGreen  = Color.FromArgb(46, 125, 50);
        private static readonly Color DarkGreen    = Color.FromArgb(27, 94, 32);
        private static readonly Color LightGreen   = Color.FromArgb(232, 245, 233);
        private static readonly Color TextMuted    = Color.FromArgb(117, 117, 117);

        public SpeechConfirmForm(string detectedReference)
        {
            BuildUI(detectedReference);
            StartAutoClose(_autoCloseMs); // 12s then dismiss
        }

        private void BuildUI(string reference)
        {
            // \u2500\u2500 Form shell \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
            FormBorderStyle = FormBorderStyle.None;
            StartPosition   = FormStartPosition.Manual;
            TopMost         = true;
            BackColor       = Color.White;
            Size            = new Size(340, 130);
            Padding         = new Padding(0);
            ShowInTaskbar   = false;

            // Position bottom-right of the primary screen
            var screen = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(screen.Right - Width - 16, screen.Bottom - Height - 16);

            // Subtle drop-shadow border
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 8, 8));

            // \u2500\u2500 Green top strip \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
            var strip = new Panel
            {
                BackColor = AccentGreen,
                Dock      = DockStyle.Top,
                Height    = 6
            };
            Controls.Add(strip);

            // \u2500\u2500 Mic icon label \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
            _lblHeading = new Label
            {
                Text      = "\ud83c\udfa4  Heard a reference",
                Font      = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = DarkGreen,
                Location  = new Point(14, 16),
                AutoSize  = true
            };
            Controls.Add(_lblHeading);

            // \u2500\u2500 Editable reference box \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
            _txtReference = new TextBox
            {
                Text      = reference,
                Font      = new Font("Segoe UI", 13f, FontStyle.Bold),
                ForeColor = DarkGreen,
                BackColor = LightGreen,
                BorderStyle = BorderStyle.None,
                Location  = new Point(14, 40),
                Size      = new Size(312, 28),
                TextAlign = HorizontalAlignment.Left
            };
            _txtReference.KeyDown += TxtReference_KeyDown;
            _txtReference.TextChanged += TxtReference_TextChanged;
            Controls.Add(_txtReference);

            // \u2500\u2500 Hint \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
            _lblHint = new Label
            {
                Text      = "Enter to insert  \u00b7  Esc to dismiss",
                Font      = new Font("Segoe UI", 8f, FontStyle.Italic),
                ForeColor = TextMuted,
                Location  = new Point(14, 72),
                AutoSize  = true
            };
            Controls.Add(_lblHint);

            // \u2500\u2500 Insert button \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
            _btnInsert = new Button
            {
                Text      = "Insert",
                Font      = new Font("Segoe UI Semibold", 8.5f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = AccentGreen,
                FlatStyle = FlatStyle.Flat,
                Size      = new Size(72, 26),
                Location  = new Point(174, 68),
                Cursor    = Cursors.Hand
            };
            _btnInsert.FlatAppearance.BorderSize = 0;
            _btnInsert.Click += (s, e) => Confirm();
            Controls.Add(_btnInsert);

            // \u2500\u2500 Dismiss button \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
            _btnDismiss = new Button
            {
                Text      = "Dismiss",
                Font      = new Font("Segoe UI", 8.5f),
                ForeColor = TextMuted,
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size      = new Size(62, 26),
                Location  = new Point(252, 68),
                Cursor    = Cursors.Hand
            };
            _btnDismiss.FlatAppearance.BorderColor = Color.FromArgb(220, 220, 220);
            _btnDismiss.Click += (s, e) => Dismiss();
            Controls.Add(_btnDismiss);

            // Thin border around whole form
            Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(180, 180, 180), 1))
                    e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            };
        }

        /// <summary>
        /// Show the form modelessly without stealing focus from PowerPoint's
        /// presenter view / slideshow. Without this override, calling Show()
        /// on a TopMost form will pull focus away from the active slideshow.
        /// </summary>
        protected override bool ShowWithoutActivation => true;

        /// <summary>
        /// Updates the reference shown in the toast. Called when the speech
        /// pipeline detects a richer reference (e.g. the user initially said
        /// "Genesis 1" and then followed up with "verse 5" \u2192 we edit the
        /// existing toast to read "Genesis 1:5" rather than opening a
        /// second popup).
        ///
        /// Safe to call from any thread \u2014 marshals to the UI thread.
        ///
        /// If the user has already manually edited the textbox we leave their
        /// text alone (they took control) and just reset the auto-close timer.
        /// </summary>
        public void UpdateReference(string newReference)
        {
            if (IsDisposed) return;
            if (string.IsNullOrWhiteSpace(newReference)) return;

            if (InvokeRequired)
            {
                try { BeginInvoke((Action)(() => UpdateReference(newReference))); }
                catch { /* form tearing down \u2014 ignore */ }
                return;
            }

            try
            {
                if (!_userHasEdited)
                {
                    // Temporarily detach TextChanged so our programmatic update
                    // isn't mistaken for the user editing the box.
                    _txtReference.TextChanged -= TxtReference_TextChanged;
                    try
                    {
                        _txtReference.Text = newReference;
                        // Keep the caret at the end so if the user now wants to
                        // append they continue naturally.
                        _txtReference.SelectionStart  = _txtReference.Text.Length;
                        _txtReference.SelectionLength = 0;
                    }
                    finally
                    {
                        _txtReference.TextChanged += TxtReference_TextChanged;
                    }

                    // Nudge the heading so the user notices the update.
                    _lblHeading.Text = "\ud83c\udfa4  Updated reference";
                }

                // Always reset the auto-close window on update \u2014 the user
                // needs a fresh chance to read the new text.
                RestartAutoClose();
            }
            catch
            {
                // Never let a UI hiccup break the speech pipeline.
            }
        }

        private void TxtReference_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                Confirm();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                Dismiss();
            }
        }

        private void TxtReference_TextChanged(object sender, EventArgs e)
        {
            // User is typing \u2014 stop letting the speech pipeline overwrite them.
            _userHasEdited = true;
        }

        private void Confirm()
        {
            if (IsDisposed) return;
            Confirmed = true;
            _autoCloseTimer?.Stop();
            string final = Reference;
            try { ReferenceConfirmed?.Invoke(this, final); }
            catch { /* subscriber error must not prevent close */ }
            Close();
        }

        private void Dismiss()
        {
            if (IsDisposed) return;
            Confirmed = false;
            _autoCloseTimer?.Stop();
            try { Dismissed?.Invoke(this, EventArgs.Empty); }
            catch { /* subscriber error must not prevent close */ }
            Close();
        }

        private void StartAutoClose(int milliseconds)
        {
            _autoCloseMs = milliseconds;
            _autoCloseTimer = new Timer { Interval = milliseconds };
            _autoCloseTimer.Tick += (s, e) => Dismiss();
            _autoCloseTimer.Start();
        }

        /// <summary>
        /// Resets the auto-close countdown. Called whenever the reference is
        /// updated so the user always has a fair chance to read the latest
        /// text before the toast vanishes.
        /// </summary>
        private void RestartAutoClose()
        {
            if (_autoCloseTimer == null) return;
            _autoCloseTimer.Stop();
            _autoCloseTimer.Interval = _autoCloseMs;
            _autoCloseTimer.Start();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            // Because the form is shown without activation, don't grab focus
            // away from the slideshow. The user can still click into the
            // textbox to edit if they want to.
            // (Previous behaviour was SelectAll() + Focus(); that stole focus
            // from presenter view and caused exactly the kind of hotkey-loss
            // this refactor is fixing.)
        }

        protected override void Dispose(bool disposing)
        {
            _autoCloseTimer?.Dispose();
            base.Dispose(disposing);
        }

        // Win32 rounded rect region
        [System.Runtime.InteropServices.DllImport("Gdi32.dll")]
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);
    }
}
