// ============================================================================
// SpeechDebugPanel.cs
//
// A floating diagnostic panel that shows live speech recognition activity.
// Displays: mic status, raw heard text, detection result, errors.
// Stays on top so you can watch it while speaking.
// ============================================================================

using System;
using System.Drawing;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public class SpeechDebugPanel : Form
    {
        private Label  _lblStatus;
        private Label  _lblHeard;
        private Label  _lblHeardVal;
        private Label  _lblDetected;
        private Label  _lblDetectedVal;
        private Label  _lblLog;
        private RichTextBox _rtbLog;
        private Button _btnClear;
        private Panel  _stripTop;

        private static readonly Color BgColor      = Color.FromArgb(30, 30, 30);
        private static readonly Color AccentGreen  = Color.FromArgb(46, 125, 50);
        private static readonly Color AccentRed    = Color.FromArgb(183, 28, 28);
        private static readonly Color TextMain     = Color.FromArgb(240, 240, 240);
        private static readonly Color TextMuted    = Color.FromArgb(150, 150, 150);
        private static readonly Color TextGreen    = Color.FromArgb(129, 199, 132);
        private static readonly Color TextYellow   = Color.FromArgb(255, 213, 79);
        private static readonly Color TextRed      = Color.FromArgb(239, 154, 154);

        public SpeechDebugPanel()
        {
            BuildUI();
        }

        private void BuildUI()
        {
            Text            = "Speech Monitor — WorshipHelper";
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            StartPosition   = FormStartPosition.Manual;
            TopMost         = true;
            BackColor       = BgColor;
            Size            = new Size(420, 460);
            MinimumSize     = new Size(320, 300);
            ShowInTaskbar   = false;

            // Position top-right
            var screen = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(screen.Right - Width - 16, screen.Top + 16);

            // ── Top status strip ─────────────────────────────────────────
            _stripTop = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 36,
                BackColor = AccentRed  // red = not listening
            };
            _lblStatus = new Label
            {
                Text      = "⏹  Not listening",
                Font      = new Font("Segoe UI Semibold", 10f, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize  = false,
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            _stripTop.Controls.Add(_lblStatus);
            Controls.Add(_stripTop);

            // ── Last heard ───────────────────────────────────────────────
            int y = 46;
            _lblHeard = MakeLabel("LAST HEARD", TextMuted, new Point(12, y), bold: false, size: 7.5f);
            Controls.Add(_lblHeard);

            y += 18;
            _lblHeardVal = MakeLabel("—", TextMain, new Point(12, y), bold: true, size: 11f);
            _lblHeardVal.Size = new Size(396, 22);
            Controls.Add(_lblHeardVal);

            // ── Detected reference ───────────────────────────────────────
            y += 30;
            _lblDetected = MakeLabel("DETECTED REFERENCE", TextMuted, new Point(12, y), bold: false, size: 7.5f);
            Controls.Add(_lblDetected);

            y += 18;
            _lblDetectedVal = MakeLabel("—", TextGreen, new Point(12, y), bold: true, size: 14f);
            _lblDetectedVal.Size = new Size(396, 28);
            Controls.Add(_lblDetectedVal);

            // ── Log ──────────────────────────────────────────────────────
            y += 36;
            _lblLog = MakeLabel("LOG", TextMuted, new Point(12, y), bold: false, size: 7.5f);
            Controls.Add(_lblLog);

            y += 18;
            _rtbLog = new RichTextBox
            {
                Location    = new Point(12, y),
                Size        = new Size(396, 310),
                Anchor      = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor   = Color.FromArgb(20, 20, 20),
                ForeColor   = TextMain,
                Font        = new Font("Consolas", 8.5f),
                ReadOnly    = true,
                BorderStyle = BorderStyle.None,
                ScrollBars  = RichTextBoxScrollBars.Vertical,
                WordWrap    = true
            };
            Controls.Add(_rtbLog);

            // ── Clear button ─────────────────────────────────────────────
            _btnClear = new Button
            {
                Text      = "Clear",
                Font      = new Font("Segoe UI", 8f),
                ForeColor = TextMuted,
                BackColor = Color.FromArgb(50, 50, 50),
                FlatStyle = FlatStyle.Flat,
                Size      = new Size(56, 22),
                Anchor    = AnchorStyles.Bottom | AnchorStyles.Right,
                Cursor    = Cursors.Hand
            };
            _btnClear.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
            _btnClear.Click += (s, e) => _rtbLog.Clear();
            Controls.Add(_btnClear);

            // Reposition clear button on resize
            Resize += (s, e) => PositionClearButton();
            PositionClearButton();
        }

        private void PositionClearButton()
        {
            _btnClear.Location = new Point(
                ClientSize.Width - _btnClear.Width - 12,
                ClientSize.Height - _btnClear.Height - 8);
        }

        // ── Public update methods (call from UI thread) ──────────────────

        public void SetListening(bool isListening)
        {
            _stripTop.BackColor = isListening ? AccentGreen : AccentRed;
            _lblStatus.Text     = isListening ? "🎤  Listening..." : "⏹  Not listening";
        }

        public void SetHeard(string text)
        {
            _lblHeardVal.Text = string.IsNullOrWhiteSpace(text) ? "—" : text;
            AppendLog($"Heard: {text}", TextMain);
        }

        public void SetDetected(string reference)
        {
            if (string.IsNullOrWhiteSpace(reference))
            {
                _lblDetectedVal.Text      = "no match";
                _lblDetectedVal.ForeColor = TextMuted;
                AppendLog("No reference detected", TextMuted);
            }
            else
            {
                _lblDetectedVal.Text      = reference;
                _lblDetectedVal.ForeColor = TextGreen;
                AppendLog($"Detected: {reference}", TextGreen);
            }
        }

        public void SetError(string message)
        {
            AppendLog($"ERROR: {message}", TextRed);
        }

        public void SetStatus(string message, bool isError = false)
        {
            AppendLog(message, isError ? TextRed : TextYellow);
        }

        private void AppendLog(string message, Color color)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { BeginInvoke(new Action(() => AppendLog(message, color))); return; }

            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            _rtbLog.SelectionStart  = _rtbLog.TextLength;
            _rtbLog.SelectionLength = 0;
            _rtbLog.SelectionColor  = TextMuted;
            _rtbLog.AppendText($"[{timestamp}] ");
            _rtbLog.SelectionColor  = color;
            _rtbLog.AppendText(message + "\n");
            _rtbLog.ScrollToCaret();
        }

        private Label MakeLabel(string text, Color color, Point location, bool bold, float size)
        {
            return new Label
            {
                Text      = text,
                Font      = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = color,
                Location  = location,
                AutoSize  = true,
                BackColor = Color.Transparent
            };
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Don't actually close, just hide — so the caller can reopen it
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
            base.OnFormClosing(e);
        }
    }
}
