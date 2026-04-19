// ============================================================================
// SpeechDebugPanel.cs
//
// A floating diagnostic panel that shows live speech recognition activity.
//
// Features:
//   * Live mic-level bar — a custom-drawn control driven by its own
//     WaveInEvent capture. Green → yellow → red as volume rises.
//   * Confidence meter — each recognition result pushes a colour-coded bar
//     (red <50%, yellow <75%, green ≥75%) with the percentage overlaid.
//   * Phase badge — SCAN / FOCUS: <book>, fed by SpeechListener.PhaseChanged.
//   * Copy button — copies the last detected reference to the clipboard.
//   * Log auto-trim — the RichTextBox is capped at 500 lines so it never
//     grows without bound.
//
// Stays on top so you can watch it while speaking.
// ============================================================================

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using NAudio.Wave;

namespace WorshipHelperVSTO
{
    // -----------------------------------------------------------------------
    // MicLevelBar — custom-drawn horizontal level meter
    // -----------------------------------------------------------------------
    internal class MicLevelBar : Control
    {
        private float _level; // 0.0 – 1.0
        private WaveInEvent _waveIn;
        private readonly object _lock = new object();

        public MicLevelBar()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.ResizeRedraw, true);
            BackColor = Color.FromArgb(20, 20, 20);
        }

        public void StartMonitoring()
        {
            lock (_lock)
            {
                if (_waveIn != null) return;
                try
                {
                    _waveIn = new WaveInEvent
                    {
                        WaveFormat = new WaveFormat(16000, 1),
                        BufferMilliseconds = 50,
                    };
                    _waveIn.DataAvailable += OnData;
                    _waveIn.StartRecording();
                }
                catch
                {
                    // Silent — the panel still works without mic monitoring.
                    _waveIn = null;
                }
            }
        }

        public void StopMonitoring()
        {
            lock (_lock)
            {
                try
                {
                    if (_waveIn != null)
                    {
                        _waveIn.DataAvailable -= OnData;
                        _waveIn.StopRecording();
                        _waveIn.Dispose();
                    }
                }
                catch { /* ignore */ }
                _waveIn = null;
                _level = 0f;
            }
            SafeInvalidate();
        }

        private void OnData(object sender, WaveInEventArgs e)
        {
            // Compute RMS of this 50ms buffer (16-bit PCM mono).
            long sum = 0;
            int samples = e.BytesRecorded / 2;
            for (int i = 0; i < e.BytesRecorded; i += 2)
            {
                short s = (short)(e.Buffer[i] | (e.Buffer[i + 1] << 8));
                sum += s * s;
            }
            double rms = samples > 0 ? Math.Sqrt((double)sum / samples) : 0;
            // Map RMS (0..32768) to 0..1 with a gentle curve so quiet speech
            // still reads on the bar.
            float raw = (float)Math.Min(1.0, rms / 8000.0);
            float eased = (float)Math.Pow(raw, 0.6);

            // Smooth with simple attack/decay for a natural-looking meter.
            float newLevel;
            lock (_lock)
            {
                newLevel = eased > _level ? eased : _level * 0.85f + eased * 0.15f;
                _level = newLevel;
            }
            SafeInvalidate();
        }

        private void SafeInvalidate()
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { BeginInvoke(new Action(SafeInvalidate)); return; }
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var r = ClientRectangle;
            using (var bg = new SolidBrush(BackColor))
                g.FillRectangle(bg, r);

            float level;
            lock (_lock) level = _level;

            int fillW = (int)(r.Width * level);
            if (fillW < 1 && level > 0) fillW = 1;

            // Draw segment-style bar with 3-zone colouring.
            int segCount = 28;
            int segGap = 2;
            int segW = Math.Max(2, (r.Width - (segCount - 1) * segGap) / segCount);
            int totalW = segCount * segW + (segCount - 1) * segGap;
            int startX = (r.Width - totalW) / 2;

            for (int i = 0; i < segCount; i++)
            {
                int x = startX + i * (segW + segGap);
                float segLevel = (i + 0.5f) / segCount;
                bool lit = level >= segLevel;

                Color c;
                if (segLevel < 0.55f) c = Color.FromArgb(70, 180, 90);       // green
                else if (segLevel < 0.85f) c = Color.FromArgb(235, 200, 60); // yellow
                else c = Color.FromArgb(230, 80, 70);                        // red

                if (!lit) c = Color.FromArgb(40, c);
                using (var b = new SolidBrush(c))
                    g.FillRectangle(b, x, r.Y + 2, segW, r.Height - 4);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) StopMonitoring();
            base.Dispose(disposing);
        }
    }

    // -----------------------------------------------------------------------
    // ConfidenceBar — custom-drawn horizontal bar with overlaid percentage
    // -----------------------------------------------------------------------
    internal class ConfidenceBar : Control
    {
        private float _value; // 0.0 – 1.0

        public ConfidenceBar()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.ResizeRedraw, true);
            BackColor = Color.FromArgb(20, 20, 20);
            ForeColor = Color.White;
            Font = new Font("Segoe UI Semibold", 8.5f, FontStyle.Bold);
        }

        public float Value
        {
            get => _value;
            set
            {
                float v = value;
                if (v < 0) v = 0; else if (v > 1) v = 1;
                if (Math.Abs(_value - v) < 0.001f) return;
                _value = v;
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var r = ClientRectangle;
            using (var bg = new SolidBrush(BackColor))
                g.FillRectangle(bg, r);

            int fillW = (int)(r.Width * _value);

            Color c;
            if (_value < 0.50f) c = Color.FromArgb(220, 80, 70);       // red
            else if (_value < 0.75f) c = Color.FromArgb(230, 190, 60); // yellow
            else c = Color.FromArgb(80, 180, 90);                      // green

            using (var b = new SolidBrush(c))
                g.FillRectangle(b, r.X, r.Y, fillW, r.Height);

            // Border
            using (var p = new Pen(Color.FromArgb(60, 60, 60)))
                g.DrawRectangle(p, r.X, r.Y, r.Width - 1, r.Height - 1);

            // Percentage text
            string text = (_value * 100f).ToString("0") + "%";
            var ts = g.MeasureString(text, Font);
            var tp = new PointF((r.Width - ts.Width) / 2f, (r.Height - ts.Height) / 2f);
            using (var shadow = new SolidBrush(Color.FromArgb(160, 0, 0, 0)))
                g.DrawString(text, Font, shadow, tp.X + 1, tp.Y + 1);
            using (var fg = new SolidBrush(ForeColor))
                g.DrawString(text, Font, fg, tp);
        }
    }

    // -----------------------------------------------------------------------
    // SpeechDebugPanel — the main form
    // -----------------------------------------------------------------------
    public class SpeechDebugPanel : Form
    {
        private Label  _lblStatus;
        private Label  _lblPhase;
        private Label  _lblMic;
        private MicLevelBar _micBar;
        private Label  _lblHeard;
        private Label  _lblHeardVal;
        private Label  _lblConf;
        private ConfidenceBar _confBar;
        private Label  _lblDetected;
        private Label  _lblDetectedVal;
        private Label  _lblLog;
        private RichTextBox _rtbLog;
        private Button _btnClear;
        private Button _btnCopy;
        private Button _btnTestMic;
        private Panel  _stripTop;

        private string _lastDetected = "";

        private const int MaxLogLines = 500;

        private static readonly Color BgColor      = Color.FromArgb(30, 30, 30);
        private static readonly Color AccentGreen  = Color.FromArgb(46, 125, 50);
        private static readonly Color AccentRed    = Color.FromArgb(183, 28, 28);
        private static readonly Color TextMain     = Color.FromArgb(240, 240, 240);
        private static readonly Color TextMuted    = Color.FromArgb(150, 150, 150);
        private static readonly Color TextGreen    = Color.FromArgb(129, 199, 132);
        private static readonly Color TextYellow   = Color.FromArgb(255, 213, 79);
        private static readonly Color TextRed      = Color.FromArgb(239, 154, 154);
        private static readonly Color BadgeScan    = Color.FromArgb(70, 90, 140);
        private static readonly Color BadgeFocus   = Color.FromArgb(140, 90, 40);

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
            Size            = new Size(460, 560);
            MinimumSize     = new Size(380, 420);
            ShowInTaskbar   = false;

            // Position top-right
            var screen = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(screen.Right - Width - 16, screen.Top + 16);

            // ── Top status strip with phase badge ────────────────────────
            _stripTop = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 38,
                BackColor = AccentRed
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

            _lblPhase = new Label
            {
                Text      = "SCAN",
                Font      = new Font("Segoe UI Semibold", 8.5f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = BadgeScan,
                AutoSize  = false,
                Size      = new Size(120, 22),
                TextAlign = ContentAlignment.MiddleCenter,
                Anchor    = AnchorStyles.Top | AnchorStyles.Right
            };
            _lblPhase.Location = new Point(_stripTop.Width - _lblPhase.Width - 8, 8);
            _stripTop.Resize += (s, e) =>
                _lblPhase.Location = new Point(_stripTop.Width - _lblPhase.Width - 8, 8);
            _stripTop.Controls.Add(_lblPhase);
            _lblPhase.BringToFront();

            Controls.Add(_stripTop);

            // ── Mic level ────────────────────────────────────────────────
            int y = 46;
            _lblMic = MakeLabel("MIC LEVEL", TextMuted, new Point(12, y), bold: false, size: 7.5f);
            Controls.Add(_lblMic);

            y += 18;
            _micBar = new MicLevelBar
            {
                Location = new Point(12, y),
                Size = new Size(ClientSize.Width - 24, 18),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            Controls.Add(_micBar);

            // ── Last heard ───────────────────────────────────────────────
            y += 28;
            _lblHeard = MakeLabel("LAST HEARD", TextMuted, new Point(12, y), bold: false, size: 7.5f);
            Controls.Add(_lblHeard);

            y += 18;
            _lblHeardVal = MakeLabel("—", TextMain, new Point(12, y), bold: true, size: 11f);
            _lblHeardVal.Size = new Size(ClientSize.Width - 24, 22);
            _lblHeardVal.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_lblHeardVal);

            // ── Confidence bar ───────────────────────────────────────────
            y += 26;
            _lblConf = MakeLabel("CONFIDENCE", TextMuted, new Point(12, y), bold: false, size: 7.5f);
            Controls.Add(_lblConf);

            y += 18;
            _confBar = new ConfidenceBar
            {
                Location = new Point(12, y),
                Size = new Size(ClientSize.Width - 24, 18),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            Controls.Add(_confBar);

            // ── Detected reference ───────────────────────────────────────
            y += 28;
            _lblDetected = MakeLabel("DETECTED REFERENCE", TextMuted, new Point(12, y), bold: false, size: 7.5f);
            Controls.Add(_lblDetected);

            y += 18;
            _lblDetectedVal = MakeLabel("—", TextGreen, new Point(12, y), bold: true, size: 14f);
            _lblDetectedVal.Size = new Size(ClientSize.Width - 24, 28);
            _lblDetectedVal.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_lblDetectedVal);

            // ── Log ──────────────────────────────────────────────────────
            y += 36;
            _lblLog = MakeLabel("LOG", TextMuted, new Point(12, y), bold: false, size: 7.5f);
            Controls.Add(_lblLog);

            y += 18;
            _rtbLog = new RichTextBox
            {
                Location    = new Point(12, y),
                Size        = new Size(ClientSize.Width - 24, ClientSize.Height - y - 40),
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

            // ── Bottom buttons ───────────────────────────────────────────
            _btnTestMic = MakeButton("Test Mic");
            _btnTestMic.Click += (s, e) => TestMic();
            Controls.Add(_btnTestMic);

            _btnCopy = MakeButton("📋 Copy");
            _btnCopy.Enabled = false;
            _btnCopy.Click += (s, e) => CopyLastDetected();
            Controls.Add(_btnCopy);

            _btnClear = MakeButton("Clear");
            _btnClear.Click += (s, e) => _rtbLog.Clear();
            Controls.Add(_btnClear);

            Resize += (s, e) => LayoutBottomButtons();
            LayoutBottomButtons();
        }

        private Button MakeButton(string text)
        {
            var btn = new Button
            {
                Text      = text,
                Font      = new Font("Segoe UI", 8f),
                ForeColor = TextMain,
                BackColor = Color.FromArgb(50, 50, 50),
                FlatStyle = FlatStyle.Flat,
                Size      = new Size(72, 24),
                Anchor    = AnchorStyles.Bottom | AnchorStyles.Right,
                Cursor    = Cursors.Hand,
                UseVisualStyleBackColor = false,
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
            return btn;
        }

        private void LayoutBottomButtons()
        {
            int y = ClientSize.Height - _btnClear.Height - 8;
            int x = ClientSize.Width - 12;

            _btnClear.Location   = new Point(x - _btnClear.Width, y);
            x -= _btnClear.Width + 6;

            _btnCopy.Location    = new Point(x - _btnCopy.Width, y);
            x -= _btnCopy.Width + 6;

            _btnTestMic.Location = new Point(x - _btnTestMic.Width, y);
        }

        // ── Public update methods (UI-thread safe) ───────────────────────

        public void SetListening(bool isListening)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { BeginInvoke(new Action(() => SetListening(isListening))); return; }

            _stripTop.BackColor = isListening ? AccentGreen : AccentRed;
            _lblStatus.Text     = isListening ? "🎤  Listening..." : "⏹  Not listening";

            if (isListening) _micBar.StartMonitoring();
            else             _micBar.StopMonitoring();
        }

        /// <summary>
        /// Display the last recognised utterance along with its confidence.
        /// Confidence is shown both as a percentage and on the coloured bar.
        /// </summary>
        public void SetHeard(string text, float confidence)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { BeginInvoke(new Action(() => SetHeard(text, confidence))); return; }

            _lblHeardVal.Text = string.IsNullOrWhiteSpace(text) ? "—" : text;
            _confBar.Value    = confidence;
            AppendLog($"Heard: {text} (conf {confidence:0.00})", TextMain);
        }

        /// <summary>
        /// Backwards-compatible overload — if no confidence is available
        /// we just leave the bar where it is.
        /// </summary>
        public void SetHeard(string text) => SetHeard(text, _confBar?.Value ?? 0f);

        public void SetDetected(string reference)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { BeginInvoke(new Action(() => SetDetected(reference))); return; }

            if (string.IsNullOrWhiteSpace(reference))
            {
                _lblDetectedVal.Text      = "no match";
                _lblDetectedVal.ForeColor = TextMuted;
                _btnCopy.Enabled = false;
                AppendLog("No reference detected", TextMuted);
            }
            else
            {
                _lblDetectedVal.Text      = reference;
                _lblDetectedVal.ForeColor = TextGreen;
                _lastDetected = reference;
                _btnCopy.Enabled = true;
                AppendLog($"Detected: {reference}", TextGreen);
            }
        }

        public void SetPhase(string phase, string bookName)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { BeginInvoke(new Action(() => SetPhase(phase, bookName))); return; }

            bool isFocus = string.Equals(phase, "Focus", StringComparison.OrdinalIgnoreCase);
            if (isFocus)
            {
                _lblPhase.Text      = string.IsNullOrEmpty(bookName)
                    ? "FOCUS"
                    : $"FOCUS: {bookName}";
                _lblPhase.BackColor = BadgeFocus;
            }
            else
            {
                _lblPhase.Text      = "SCAN";
                _lblPhase.BackColor = BadgeScan;
            }
            // Keep it pinned to the right — in case text length changed
            _lblPhase.Width = Math.Max(120, TextRenderer.MeasureText(_lblPhase.Text, _lblPhase.Font).Width + 16);
            _lblPhase.Location = new Point(_stripTop.Width - _lblPhase.Width - 8, 8);
        }

        public void SetError(string message)
        {
            AppendLog($"ERROR: {message}", TextRed);
        }

        public void SetStatus(string message, bool isError = false)
        {
            AppendLog(message, isError ? TextRed : TextYellow);
        }

        private void CopyLastDetected()
        {
            if (string.IsNullOrEmpty(_lastDetected)) return;
            try
            {
                Clipboard.SetText(_lastDetected);
                AppendLog($"Copied: {_lastDetected}", TextYellow);
            }
            catch (Exception ex)
            {
                AppendLog($"Copy failed: {ex.Message}", TextRed);
            }
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

            TrimLog();
            _rtbLog.ScrollToCaret();
        }

        private void TrimLog()
        {
            int lines = _rtbLog.Lines.Length;
            if (lines <= MaxLogLines) return;

            // Remove oldest (lines - MaxLogLines) lines
            int toRemove = lines - MaxLogLines;
            int firstKeep = _rtbLog.GetFirstCharIndexFromLine(toRemove);
            if (firstKeep > 0)
            {
                _rtbLog.Select(0, firstKeep);
                _rtbLog.SelectedText = string.Empty;
            }
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

        private void TestMic()
        {
            AppendLog("Testing microphone access...", TextYellow);
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    int deviceCount = NAudio.Wave.WaveInEvent.DeviceCount;
                    if (deviceCount == 0)
                    {
                        AppendLog("✗ No recording devices found", TextRed);
                        return;
                    }
                    var caps = NAudio.Wave.WaveInEvent.GetCapabilities(0);
                    AppendLog($"✓ Mic accessible — {deviceCount} device(s), default: \"{caps.ProductName}\"", TextGreen);
                }
                catch (Exception ex)
                {
                    AppendLog($"✗ Mic error: {ex.Message}", TextRed);
                }
            });
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
                _micBar?.StopMonitoring();
            }
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { _micBar?.StopMonitoring(); } catch { /* ignore */ }
            }
            base.Dispose(disposing);
        }
    }
}
