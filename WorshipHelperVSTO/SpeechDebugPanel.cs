// ============================================================================
// SpeechDebugPanel.cs   —   v5.1 (modernised)
//
// Floating diagnostic panel for live speech recognition activity.
//
// v5.1 UI refresh:
//   * Card-based layout with rounded section backgrounds (was a flat dark
//     rectangle).
//   * Phase badge (SCAN / FOCUS: <book>) is now in a dedicated right-docked
//     slot — previously it was painted on TOP of a DockStyle.Fill status
//     label, which relied on z-order and could be covered at runtime.
//   * Top status strip uses the product green/red accent colours and
//     rounded corners at the top of the window.
//   * Bottom buttons are ModernButton instances with proper hover/press
//     feedback.
//   * Layout is driven by a TableLayoutPanel so nothing overlaps when the
//     user resizes the window.
//   * Log panel still caps at 500 lines.
// ============================================================================

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using NAudio.Wave;
using WorshipHelperVSTO.UI;

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
            BackColor = Color.FromArgb(24, 28, 24);
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
                        WaveFormat         = new WaveFormat(16000, 1),
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
                _level  = 0f;
            }
            SafeInvalidate();
        }

        private void OnData(object sender, WaveInEventArgs e)
        {
            // Compute RMS of this 50ms buffer (16-bit PCM mono).
            long sum     = 0;
            int  samples = e.BytesRecorded / 2;
            for (int i = 0; i < e.BytesRecorded; i += 2)
            {
                short s = (short)(e.Buffer[i] | (e.Buffer[i + 1] << 8));
                sum += s * s;
            }
            double rms = samples > 0 ? Math.Sqrt((double)sum / samples) : 0;
            float raw   = (float)Math.Min(1.0, rms / 8000.0);
            float eased = (float)Math.Pow(raw, 0.6);

            lock (_lock)
                _level = eased > _level ? eased : _level * 0.85f + eased * 0.15f;

            SafeInvalidate();
        }

        private void SafeInvalidate()
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(SafeInvalidate)); } catch { } return; }
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var r = ClientRectangle;
            using (var path = Palette.RoundedPath(new Rectangle(r.X, r.Y, r.Width - 1, r.Height - 1), 6))
            using (var bg   = new SolidBrush(BackColor))
                g.FillPath(bg, path);

            float level;
            lock (_lock) level = _level;

            // Segment-style bar with 3-zone colouring.
            int segCount = 28;
            int segGap   = 2;
            int segW     = Math.Max(2, (r.Width - 8 - (segCount - 1) * segGap) / segCount);
            int totalW   = segCount * segW + (segCount - 1) * segGap;
            int startX   = (r.Width - totalW) / 2;

            for (int i = 0; i < segCount; i++)
            {
                int x = startX + i * (segW + segGap);
                float segLevel = (i + 0.5f) / segCount;
                bool  lit      = level >= segLevel;

                Color c;
                if      (segLevel < 0.55f) c = Color.FromArgb(70, 180, 90);
                else if (segLevel < 0.85f) c = Color.FromArgb(235, 200, 60);
                else                       c = Color.FromArgb(230, 80, 70);

                if (!lit) c = Color.FromArgb(48, c);
                using (var b = new SolidBrush(c))
                    g.FillRectangle(b, x, r.Y + 3, segW, r.Height - 6);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) StopMonitoring();
            base.Dispose(disposing);
        }
    }

    // -----------------------------------------------------------------------
    // ConfidenceBar — rounded, with percentage overlay
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
            BackColor = Color.FromArgb(24, 28, 24);
            ForeColor = Color.White;
            Font      = new Font("Segoe UI Semibold", 8.5f, FontStyle.Bold);
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

            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = Palette.RoundedPath(r, 6))
            using (var bg   = new SolidBrush(BackColor))
                g.FillPath(bg, path);

            int fillW = (int)(r.Width * _value);
            if (fillW > 0)
            {
                Color c;
                if      (_value < 0.50f) c = Color.FromArgb(220, 80, 70);
                else if (_value < 0.75f) c = Color.FromArgb(230, 190, 60);
                else                     c = Color.FromArgb(80, 180, 90);

                var fillR = new Rectangle(0, 0, fillW, r.Height);
                using (var fillPath = Palette.RoundedPath(fillR, 6))
                using (var br       = new LinearGradientBrush(
                    fillR, Color.FromArgb(220, c), c, LinearGradientMode.Horizontal))
                    g.FillPath(br, fillPath);
            }

            // Percentage text
            string text = (_value * 100f).ToString("0") + "%";
            var ts = g.MeasureString(text, Font);
            var tp = new PointF((r.Width - ts.Width) / 2f, (r.Height - ts.Height) / 2f);
            using (var shadow = new SolidBrush(Color.FromArgb(180, 0, 0, 0)))
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
        private Panel         _topBar;
        private Label         _lblStatus;
        private Label         _lblPhase;
        private MicLevelBar   _micBar;
        private Label         _lblHeardVal;
        private ConfidenceBar _confBar;
        private Label         _lblDetectedVal;
        private RichTextBox   _rtbLog;
        private ModernButton  _btnClear;
        private ModernButton  _btnCopy;
        private ModernButton  _btnTestMic;

        private string _lastDetected = "";

        private const int MaxLogLines = 500;

        private static readonly Color BgColor    = Color.FromArgb(32, 34, 32);
        private static readonly Color CardBg     = Color.FromArgb(44, 48, 44);
        private static readonly Color AccentRed  = Color.FromArgb(183, 28, 28);
        private static readonly Color TextMain   = Color.FromArgb(238, 238, 232);
        private static readonly Color TextMuted  = Color.FromArgb(150, 158, 150);
        private static readonly Color TextGreen  = Color.FromArgb(129, 199, 132);
        private static readonly Color TextYellow = Color.FromArgb(255, 213, 79);
        private static readonly Color TextRed    = Color.FromArgb(239, 154, 154);
        private static readonly Color BadgeScan  = Color.FromArgb(70, 90, 140);
        private static readonly Color BadgeFocus = Color.FromArgb(140, 90, 40);

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
            Size            = new Size(480, 600);
            MinimumSize     = new Size(400, 460);
            ShowInTaskbar   = false;
            DoubleBuffered  = true;

            var screen = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(screen.Right - Width - 16, screen.Top + 16);

            // Root table: row 0 = top bar, row 1 = content, row 2 = buttons
            var root = new TableLayoutPanel
            {
                Dock        = DockStyle.Fill,
                ColumnCount = 1,
                RowCount    = 3,
                BackColor   = BgColor,
                Padding     = new Padding(0),
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52f));
            Controls.Add(root);

            // ---- Top bar: status (fill) + phase badge (right) ----
            _topBar = new Panel
            {
                Dock      = DockStyle.Fill,
                BackColor = AccentRed,
                Padding   = new Padding(0),
            };
            _topBar.Paint += OnTopBarPaint;

            // Badge docked to the right — added FIRST so Fill label goes
            // into the remaining space, not over it.
            _lblPhase = new Label
            {
                Text      = "SCAN",
                Font      = new Font("Segoe UI Semibold", 8.5f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = BadgeScan,
                AutoSize  = false,
                Width     = 130,
                Dock      = DockStyle.Right,
                TextAlign = ContentAlignment.MiddleCenter,
                Margin    = new Padding(0, 6, 8, 6),
                Padding   = new Padding(6, 4, 6, 4),
            };
            _topBar.Controls.Add(_lblPhase);

            _lblStatus = new Label
            {
                Text      = "⏹  Not listening",
                Font      = new Font("Segoe UI Semibold", 10.5f, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize  = false,
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.Transparent,
            };
            _topBar.Controls.Add(_lblStatus);
            _lblStatus.BringToFront(); // ensure text is in front of top-bar paint

            root.Controls.Add(_topBar, 0, 0);

            // ---- Content area with card backgrounds ----
            var content = new TableLayoutPanel
            {
                Dock        = DockStyle.Fill,
                ColumnCount = 1,
                RowCount    = 5,
                BackColor   = BgColor,
                Padding     = new Padding(12, 12, 12, 6),
            };
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 52f));   // Mic
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 64f));   // Last heard
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 52f));   // Confidence
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 68f));   // Detected reference
            content.RowStyles.Add(new RowStyle(SizeType.Percent,  100f));  // Log
            root.Controls.Add(content, 0, 1);

            _micBar = new MicLevelBar();
            content.Controls.Add(WrapCard("MIC LEVEL", _micBar, 20, 22), 0, 0);

            _lblHeardVal = MakeLabel("—", TextMain, bold: true, size: 11.5f);
            _lblHeardVal.AutoEllipsis = true;
            _lblHeardVal.TextAlign    = ContentAlignment.MiddleLeft;
            content.Controls.Add(WrapCard("LAST HEARD", _lblHeardVal, 22, 28), 0, 1);

            _confBar = new ConfidenceBar();
            content.Controls.Add(WrapCard("CONFIDENCE", _confBar, 20, 22), 0, 2);

            _lblDetectedVal = MakeLabel("—", TextGreen, bold: true, size: 15f);
            _lblDetectedVal.AutoEllipsis = true;
            _lblDetectedVal.TextAlign    = ContentAlignment.MiddleLeft;
            content.Controls.Add(WrapCard("DETECTED REFERENCE", _lblDetectedVal, 22, 34), 0, 3);

            // Log
            var logCard = new Panel { Dock = DockStyle.Fill, BackColor = BgColor, Padding = new Padding(0, 4, 0, 0) };
            var lblLog = new Label
            {
                Text      = "LOG",
                Font      = new Font("Segoe UI Semibold", 7.8f, FontStyle.Bold),
                ForeColor = TextMuted,
                AutoSize  = true,
                Location  = new Point(4, 0),
            };
            logCard.Controls.Add(lblLog);

            _rtbLog = new RichTextBox
            {
                Location    = new Point(0, 18),
                Anchor      = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor   = Color.FromArgb(24, 28, 24),
                ForeColor   = TextMain,
                Font        = new Font("Consolas", 8.75f),
                ReadOnly    = true,
                BorderStyle = BorderStyle.None,
                ScrollBars  = RichTextBoxScrollBars.Vertical,
                WordWrap    = true,
            };
            logCard.Resize += (s, e) =>
            {
                _rtbLog.Size = new Size(logCard.Width, Math.Max(0, logCard.Height - 18));
            };
            _rtbLog.Size = new Size(logCard.Width, Math.Max(0, logCard.Height - 18));
            logCard.Controls.Add(_rtbLog);
            content.Controls.Add(logCard, 0, 4);

            // ---- Bottom buttons ----
            var btnRow = new Panel { Dock = DockStyle.Fill, BackColor = BgColor, Padding = new Padding(12, 8, 12, 12) };
            _btnTestMic = new ModernButton
            {
                Text = "Test Mic", Primary = false, Size = new Size(90, 30),
                AccentOverride = Palette.AccentGreen,
            };
            _btnTestMic.Click += (s, e) => TestMic();

            _btnCopy = new ModernButton
            {
                Text = "📋  Copy", Primary = false, Size = new Size(90, 30),
                AccentOverride = Palette.AccentGreen,
                Enabled = false,
            };
            _btnCopy.Click += (s, e) => CopyLastDetected();

            _btnClear = new ModernButton
            {
                Text = "Clear", Primary = false, Size = new Size(80, 30),
                AccentOverride = Palette.AccentGreen,
            };
            _btnClear.Click += (s, e) => _rtbLog.Clear();

            btnRow.Controls.Add(_btnTestMic);
            btnRow.Controls.Add(_btnCopy);
            btnRow.Controls.Add(_btnClear);
            btnRow.Resize += (s, e) => LayoutBottomButtons(btnRow);
            LayoutBottomButtons(btnRow);

            root.Controls.Add(btnRow, 0, 2);
        }

        private void OnTopBarPaint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = _topBar.ClientRectangle;
            using (var br = new LinearGradientBrush(r,
                    _topBar.BackColor,
                    Color.FromArgb(Math.Max(0, _topBar.BackColor.R - 20),
                                   Math.Max(0, _topBar.BackColor.G - 20),
                                   Math.Max(0, _topBar.BackColor.B - 20)),
                    LinearGradientMode.Vertical))
                g.FillRectangle(br, r);
        }

        /// <summary>Wraps a content control in a rounded dark "card" with a caption.</summary>
        private Panel WrapCard(string caption, Control body, int bodyY, int bodyH)
        {
            var card = new Panel
            {
                Dock    = DockStyle.Fill,
                Margin  = new Padding(0, 0, 0, 6),
                BackColor = Color.Transparent,
                Padding = new Padding(10, 6, 10, 6),
            };
            card.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var r = new Rectangle(0, 0, card.Width - 1, card.Height - 1);
                using (var path = Palette.RoundedPath(r, 8))
                using (var br   = new SolidBrush(CardBg))
                    g.FillPath(br, path);
            };

            var lbl = new Label
            {
                Text      = caption,
                Font      = new Font("Segoe UI Semibold", 7.5f, FontStyle.Bold),
                ForeColor = TextMuted,
                AutoSize  = true,
                Location  = new Point(10, 4),
                BackColor = Color.Transparent,
            };
            card.Controls.Add(lbl);

            body.Location = new Point(10, bodyY);
            body.Size     = new Size(card.Width - 20, bodyH);
            body.Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            body.BackColor = body.BackColor == Color.Transparent ? Color.Transparent : body.BackColor;
            card.Resize += (s, e) =>
            {
                body.Size = new Size(card.Width - 20, bodyH);
            };
            card.Controls.Add(body);

            return card;
        }

        private void LayoutBottomButtons(Panel row)
        {
            int y = (row.Height - _btnClear.Height) / 2;
            int x = row.ClientSize.Width - row.Padding.Right;

            _btnClear.Location   = new Point(x - _btnClear.Width,   y);
            x -= _btnClear.Width + 6;
            _btnCopy.Location    = new Point(x - _btnCopy.Width,    y);
            x -= _btnCopy.Width + 6;
            _btnTestMic.Location = new Point(x - _btnTestMic.Width, y);
        }

        // ── Public update methods (UI-thread safe) ───────────────────────

        public void SetListening(bool isListening)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetListening(isListening))); } catch { } return; }

            _topBar.BackColor = isListening ? Palette.AccentGreen : AccentRed;
            _lblStatus.Text   = isListening ? "🎤  Listening…"     : "⏹  Not listening";
            _topBar.Invalidate();

            if (isListening) _micBar.StartMonitoring();
            else             _micBar.StopMonitoring();
        }

        public void SetHeard(string text, float confidence)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetHeard(text, confidence))); } catch { } return; }

            _lblHeardVal.Text = string.IsNullOrWhiteSpace(text) ? "—" : text;
            _confBar.Value    = confidence;
            AppendLog($"Heard: {text} (conf {confidence:0.00})", TextMain);
        }

        public void SetHeard(string text) => SetHeard(text, _confBar?.Value ?? 0f);

        public void SetDetected(string reference)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetDetected(reference))); } catch { } return; }

            if (string.IsNullOrWhiteSpace(reference))
            {
                _lblDetectedVal.Text      = "no match";
                _lblDetectedVal.ForeColor = TextMuted;
                _btnCopy.Enabled          = false;
                AppendLog("No reference detected", TextMuted);
            }
            else
            {
                _lblDetectedVal.Text      = reference;
                _lblDetectedVal.ForeColor = TextGreen;
                _lastDetected             = reference;
                _btnCopy.Enabled          = true;
                AppendLog($"Detected: {reference}", TextGreen);
            }
        }

        public void SetPhase(string phase, string bookName)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetPhase(phase, bookName))); } catch { } return; }

            bool isFocus = string.Equals(phase, "Focus", StringComparison.OrdinalIgnoreCase);
            if (isFocus)
            {
                _lblPhase.Text      = string.IsNullOrEmpty(bookName) ? "FOCUS" : $"FOCUS: {bookName}";
                _lblPhase.BackColor = BadgeFocus;
            }
            else
            {
                _lblPhase.Text      = "SCAN";
                _lblPhase.BackColor = BadgeScan;
            }
            // Resize badge to fit text (but clamp so it never eats the status).
            int want = TextRenderer.MeasureText(_lblPhase.Text, _lblPhase.Font).Width + 24;
            _lblPhase.Width = Math.Max(110, Math.Min(200, want));
        }

        public void SetError(string message)   => AppendLog($"ERROR: {message}", TextRed);

        public void SetStatus(string message, bool isError = false) =>
            AppendLog(message, isError ? TextRed : TextYellow);

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
            if (InvokeRequired) { try { BeginInvoke(new Action(() => AppendLog(message, color))); } catch { } return; }

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

            int toRemove  = lines - MaxLogLines;
            int firstKeep = _rtbLog.GetFirstCharIndexFromLine(toRemove);
            if (firstKeep > 0)
            {
                _rtbLog.Select(0, firstKeep);
                _rtbLog.SelectedText = string.Empty;
            }
        }

        private Label MakeLabel(string text, Color color, bool bold, float size)
        {
            return new Label
            {
                Text      = text,
                Font      = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = color,
                AutoSize  = false,
                BackColor = Color.Transparent,
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
                try { _micBar?.StopMonitoring(); } catch { }
            }
            base.Dispose(disposing);
        }
    }
}
