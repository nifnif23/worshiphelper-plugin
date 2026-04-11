// ============================================================================
// SpeechConfirmForm.cs
//
// A small toast-style popup that appears when speech recognition detects a
// Bible reference. The user can edit the reference and press Enter to insert,
// or Escape to dismiss. Auto-dismisses after a timeout if no interaction.
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

        public string Reference => _txtReference.Text.Trim();
        public bool Confirmed { get; private set; } = false;

        private static readonly Color AccentGreen  = Color.FromArgb(46, 125, 50);
        private static readonly Color DarkGreen    = Color.FromArgb(27, 94, 32);
        private static readonly Color LightGreen   = Color.FromArgb(232, 245, 233);
        private static readonly Color TextMuted    = Color.FromArgb(117, 117, 117);

        public SpeechConfirmForm(string detectedReference)
        {
            BuildUI(detectedReference);
            StartAutoClose(12000); // 12s then dismiss
        }

        private void BuildUI(string reference)
        {
            // ── Form shell ──────────────────────────────────────────────
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

            // ── Green top strip ──────────────────────────────────────────
            var strip = new Panel
            {
                BackColor = AccentGreen,
                Dock      = DockStyle.Top,
                Height    = 6
            };
            Controls.Add(strip);

            // ── Mic icon label ───────────────────────────────────────────
            _lblHeading = new Label
            {
                Text      = "🎤  Heard a reference",
                Font      = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = DarkGreen,
                Location  = new Point(14, 16),
                AutoSize  = true
            };
            Controls.Add(_lblHeading);

            // ── Editable reference box ───────────────────────────────────
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
            Controls.Add(_txtReference);

            // ── Hint ─────────────────────────────────────────────────────
            _lblHint = new Label
            {
                Text      = "Enter to insert  ·  Esc to dismiss",
                Font      = new Font("Segoe UI", 8f, FontStyle.Italic),
                ForeColor = TextMuted,
                Location  = new Point(14, 72),
                AutoSize  = true
            };
            Controls.Add(_lblHint);

            // ── Insert button ─────────────────────────────────────────────
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

            // ── Dismiss button ────────────────────────────────────────────
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

        private void Confirm()
        {
            Confirmed = true;
            _autoCloseTimer?.Stop();
            Close();
        }

        private void Dismiss()
        {
            Confirmed = false;
            _autoCloseTimer?.Stop();
            Close();
        }

        private void StartAutoClose(int milliseconds)
        {
            _autoCloseTimer = new Timer { Interval = milliseconds };
            _autoCloseTimer.Tick += (s, e) => Dismiss();
            _autoCloseTimer.Start();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _txtReference.SelectAll();
            _txtReference.Focus();
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
