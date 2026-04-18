// ============================================================================
// AutoScriptureMode.cs
// Manages the "Auto Scripture" listening mode.
//
// When active, any scripture reference detected from speech is inserted
// immediately into the presentation — no button press, no form, no
// confirmation dialog. The presenter just says the reference aloud and
// the slide appears.
//
// How it works:
//   1. AutoScriptureMode.Enable() is called (ribbon button or Shift hotkey).
//   2. SpeechToScriptureService starts listening (if not already).
//   3. OnReferenceDetected fires → InsertScriptureFromSpeech() is called
//      directly (same path as the manual flow).
//   4. A brief on-screen toast shows what was inserted so the presenter
//      knows it worked.
//   5. AutoScriptureMode.Disable() stops the auto-insert behaviour.
//      The speech service stays active but stops auto-inserting.
//
// The mode is indicated by:
//   - A ribbon toggle button (btnAutoScripture)
//   - A small overlay label on the presenter view (optional, gracefully skipped)
//
// Thread safety:
//   IsEnabled is read/written on the main STA thread via Invoke.
//   The speech callback fires on a background thread — it uses BeginInvoke
//   to marshal insertion to the main thread.
// ============================================================================

using System;
using System.Windows.Forms;
using log4net;
using Microsoft.Office.Interop.PowerPoint;

namespace WorshipHelperVSTO
{
    /// <summary>
    /// Singleton that manages the Auto Scripture listening mode.
    /// Access via AutoScriptureMode.Instance.
    /// </summary>
    public sealed class AutoScriptureMode
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(AutoScriptureMode));
        private static readonly AutoScriptureMode _instance = new AutoScriptureMode();
        public static AutoScriptureMode Instance => _instance;

        private AutoScriptureMode() { }

        // -----------------------------------------------------------------------
        // State
        // -----------------------------------------------------------------------

        /// <summary>
        /// True when Auto Scripture Mode is active and inserting references
        /// automatically from speech.
        /// </summary>
        public bool IsEnabled { get; private set; }

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Enables Auto Scripture Mode.
        /// Starts the speech service if it isn't already running.
        /// </summary>
        public void Enable(SpeechToScriptureService service)
        {
            if (IsEnabled) return;
            IsEnabled = true;

            log.Info("AutoScriptureMode: Enabled.");

            // Start listening if not already
            if (service != null && !service.IsListening)
            {
                try { service.Start(); }
                catch (Exception ex)
                {
                    log.Error("AutoScriptureMode: Failed to start speech service.", ex);
                    IsEnabled = false;
                    return;
                }
            }

            ShowToast("Auto Scripture ON — listening for references", durationMs: 2500);
        }

        /// <summary>
        /// Disables Auto Scripture Mode.
        /// The speech service keeps running but stops auto-inserting.
        /// </summary>
        public void Disable()
        {
            if (!IsEnabled) return;
            IsEnabled = false;
            log.Info("AutoScriptureMode: Disabled.");
            ShowToast("Auto Scripture OFF", durationMs: 1500);
        }

        /// <summary>
        /// Toggles Auto Scripture Mode. Returns new state.
        /// </summary>
        public bool Toggle(SpeechToScriptureService service)
        {
            if (IsEnabled) { Disable(); return false; }
            else { Enable(service); return true; }
        }

        /// <summary>
        /// Called by the speech pipeline when a reference is detected and
        /// Auto Scripture Mode is active. Performs the actual insertion and
        /// shows feedback.
        ///
        /// Must be called from the main STA thread (use BeginInvoke if on
        /// a background thread).
        /// </summary>
        public void HandleDetectedReference(string normalisedReference,
                                            string spokenText,
                                            Action<string> insertAction)
        {
            if (!IsEnabled) return;

            log.Info($"AutoScriptureMode: Auto-inserting \"{normalisedReference}\" " +
                     $"(spoken: \"{spokenText}\")");

            try
            {
                insertAction(normalisedReference);
                ShowToast($"Inserted: {normalisedReference}", durationMs: 2000);
            }
            catch (Exception ex)
            {
                log.Error($"AutoScriptureMode: Insert failed for \"{normalisedReference}\"", ex);
                ShowToast($"Insert failed: {normalisedReference}", durationMs: 2000);
            }
        }

        // -----------------------------------------------------------------------
        // Toast notification
        // -----------------------------------------------------------------------

        private Form _toastForm;

        /// <summary>
        /// Shows a brief non-modal overlay message to the presenter.
        /// Appears bottom-right of the screen, auto-dismisses after durationMs.
        /// Fails silently if UI is unavailable (no crash during slideshow).
        /// </summary>
        private void ShowToast(string message, int durationMs = 2000)
        {
            try
            {
                // Dismiss any existing toast
                _toastForm?.Close();
                _toastForm = null;

                var toast = new Form
                {
                    FormBorderStyle = FormBorderStyle.None,
                    StartPosition   = FormStartPosition.Manual,
                    BackColor       = System.Drawing.Color.FromArgb(30, 30, 30),
                    Opacity         = 0.88,
                    ShowInTaskbar   = false,
                    TopMost         = true,
                    Size            = new System.Drawing.Size(380, 52),
                };

                // Position bottom-right of primary screen
                var screen = Screen.PrimaryScreen.WorkingArea;
                toast.Location = new System.Drawing.Point(
                    screen.Right - toast.Width - 20,
                    screen.Bottom - toast.Height - 20);

                var label = new Label
                {
                    Text      = message,
                    ForeColor = System.Drawing.Color.White,
                    BackColor = System.Drawing.Color.Transparent,
                    Font      = new System.Drawing.Font("Segoe UI", 11f, System.Drawing.FontStyle.Regular),
                    Dock      = DockStyle.Fill,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                    Padding   = new Padding(8, 0, 8, 0),
                };
                toast.Controls.Add(label);

                // Auto-dismiss timer
                var timer = new Timer { Interval = durationMs };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    try { toast.Close(); } catch { }
                };

                toast.Shown += (s, e) => timer.Start();
                toast.Show();
                _toastForm = toast;
            }
            catch (Exception ex)
            {
                // Toast is cosmetic only — never crash on failure
                log.Debug($"AutoScriptureMode: Toast failed (non-fatal): {ex.Message}");
            }
        }
    }
}
