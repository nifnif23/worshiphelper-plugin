// ============================================================================
// SpeechToScriptureService.cs
// Orchestrator that wires together:
//   SpeechListener → BibleReferenceDetector → OnReferenceDetected event
//
// This is the single entry-point you interact with from your add-in.
//
// Drop into:  WorshipHelperVSTO/SpeechToScriptureService.cs
// Namespace:  WorshipHelperVSTO
// ============================================================================

using System;
using System.Collections.Generic;
using System.Timers;
using log4net;

namespace WorshipHelperVSTO
{
    // -----------------------------------------------------------------------
    // Event args for the reference-detected event
    // -----------------------------------------------------------------------

    /// <summary>
    /// Carries a normalised Bible reference string to subscribers.
    /// </summary>
    public class ReferenceDetectedEventArgs : EventArgs
    {
        /// <summary>
        /// The fully normalised reference, e.g. "John 3:16", "1 Corinthians 13:4-7".
        /// Ready to pass into InsertScripture / FullReferenceParser.ParseFullReference.
        /// </summary>
        public string NormalisedReference { get; set; }

        /// <summary>
        /// The canonical book name, e.g. "John", "1 Corinthians".
        /// </summary>
        public string BookName { get; set; }

        /// <summary>
        /// The numeric reference fragment, e.g. "3:16", "13:4-7".
        /// </summary>
        public string ReferenceFragment { get; set; }

        /// <summary>
        /// The original spoken text that was recognised.
        /// </summary>
        public string SpokenText { get; set; }

        /// <summary>
        /// Overall confidence (combination of speech engine + Bible detection).
        /// </summary>
        public double Confidence { get; set; }
    }

    /// <summary>
    /// Carries status updates from the service to the UI.
    /// </summary>
    public class ServiceStatusEventArgs : EventArgs
    {
        public string Message { get; set; }
        public bool IsError { get; set; }
        public bool IsListening { get; set; }
    }

    // -----------------------------------------------------------------------
    // SpeechToScriptureService
    // -----------------------------------------------------------------------

    /// <summary>
    /// Complete speech-to-scripture pipeline.
    ///
    /// Architecture:
    ///   Microphone → SpeechListener (System.Speech, offline)
    ///       → raw text
    ///       → BibleReferenceDetector (book name + spoken number parsing)
    ///       → normalised reference string
    ///       → OnReferenceDetected event
    ///       → your existing InsertScripture logic
    ///
    /// Features:
    ///   - Start / Stop / Toggle listening.
    ///   - Duplicate suppression: same reference won't fire twice within a configurable cooldown.
    ///   - Conservative detection: only fires when confidence exceeds threshold.
    ///   - Fully event-driven: subscribe to OnReferenceDetected and OnStatusChanged.
    ///   - Thread-safe.
    ///
    /// Usage:
    ///   var service = new SpeechToScriptureService();
    ///   service.OnReferenceDetected += (s, e) => InsertScripture(e.NormalisedReference);
    ///   service.Start();
    /// </summary>
    public class SpeechToScriptureService : IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(SpeechToScriptureService));

        private readonly SpeechListener _listener;

        /// <summary>
        /// Public access to the underlying speech listener so the UI can
        /// subscribe to low-level events like <see cref="SpeechListener.PhaseChanged"/>.
        /// </summary>
        public SpeechListener Listener => _listener;
        private bool _disposed;
        private Bible _validationBible; // cached for reference validation

        // Duplicate suppression
        private readonly Dictionary<string, DateTime> _recentReferences = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
        private readonly object _recentLock = new object();
        private Timer _cleanupTimer;

        // -----------------------------------------------------------------------
        // Configuration
        // -----------------------------------------------------------------------

        /// <summary>
        /// Cooldown period: how long (in seconds) before the same reference can
        /// be detected again.  Prevents rapid duplicate insertions if the speaker
        /// repeats themselves or the engine produces the same result twice.
        /// Default: 30 seconds.
        /// </summary>
        public int DuplicateCooldownSeconds { get; set; } = 30;

        /// <summary>
        /// Minimum combined confidence required for a detection to fire the event.
        /// This combines the speech engine confidence and the Bible detection confidence.
        /// Range: 0.0 – 1.0.  Default: 0.4
        /// </summary>
        public double MinCombinedConfidence { get; set; } = 0.4;

        /// <summary>
        /// Minimum speech engine confidence.  Phrases below this are discarded
        /// before reaching the Bible detector.  Default: 0.3
        /// </summary>
        public float MinSpeechConfidence
        {
            get => _listener.MinEngineConfidence;
            set => _listener.MinEngineConfidence = value;
        }

        /// <summary>
        /// Minimum Bible reference detector confidence.
        /// Default: 0.5
        /// </summary>
        public double MinDetectorConfidence
        {
            get => BibleReferenceDetector.MinConfidence;
            set => BibleReferenceDetector.MinConfidence = value;
        }

        // -----------------------------------------------------------------------
        // Events
        // -----------------------------------------------------------------------

        /// <summary>
        /// Fired when a Bible reference is detected in speech.
        /// Subscribe to this event and wire it into your InsertScripture logic.
        ///
        /// NOTE: This event fires on a background thread (the speech engine's thread).
        /// If you need to update PowerPoint objects or UI, marshal to the main thread
        /// using Control.Invoke or similar.
        /// </summary>
        public event EventHandler<ReferenceDetectedEventArgs> OnReferenceDetected;

        /// <summary>
        /// Fired for every raw phrase the speech engine hears, before reference detection.
        /// Useful for diagnostics — lets you see what the mic is actually picking up.
        /// </summary>
        public event EventHandler<SpeechRecognisedEventArgs> OnRawSpeech;

        /// <summary>
        /// Fired when the service status changes (started, stopped, errors).
        /// Useful for updating a status indicator in the ribbon / UI.
        /// </summary>
        public event EventHandler<ServiceStatusEventArgs> OnStatusChanged;

        // -----------------------------------------------------------------------
        // Constructor
        // -----------------------------------------------------------------------

        public SpeechToScriptureService()
        {
            _listener = new SpeechListener();
            _listener.SpeechRecognised += OnSpeechRecognised;
            _listener.StatusChanged += OnListenerStatusChanged;

            // Periodic cleanup of the duplicate suppression cache
            _cleanupTimer = new Timer(60_000); // Every 60 seconds
            _cleanupTimer.Elapsed += (s, e) => CleanupRecentReferences();
            _cleanupTimer.AutoReset = true;
            _cleanupTimer.Start();
        }

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Returns true if the service is currently listening for speech.
        /// </summary>
        public bool IsListening => _listener.IsListening;

        /// <summary>
        /// Starts listening to the microphone and processing speech for Bible references.
        /// </summary>
        public void Start()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(SpeechToScriptureService));

            log.Info("SpeechToScriptureService: Starting…");
            _listener.Start();

            RaiseStatus("Listening for Bible references…", isListening: true);
        }

        /// <summary>
        /// Stops listening.  Can be restarted later.
        /// </summary>
        public void Stop()
        {
            log.Info("SpeechToScriptureService: Stopping…");
            _listener.Stop();

            RaiseStatus("Speech listening stopped.", isListening: false);
        }

        /// <summary>
        /// Toggles between listening and stopped.
        /// Returns the new state (true = listening).
        /// </summary>
        public bool Toggle()
        {
            bool newState = _listener.Toggle();
            RaiseStatus(
                newState ? "Listening for Bible references…" : "Speech listening stopped.",
                isListening: newState);
            return newState;
        }

        /// <summary>
        /// Clears the duplicate suppression cache.
        /// Call this when starting a new section of the presentation
        /// where references may legitimately be repeated.
        /// </summary>
        public void ClearDuplicateCache()
        {
            lock (_recentLock)
            {
                _recentReferences.Clear();
            }
            log.Debug("SpeechToScriptureService: Duplicate cache cleared.");
        }

        /// <summary>
        /// Manually process a text string through the Bible reference detection pipeline.
        /// Useful for testing without a microphone.
        /// </summary>
        public void ProcessText(string text)
        {
            OnSpeechRecognised(this, new SpeechRecognisedEventArgs
            {
                Text = text,
                Confidence = 1.0f  // Manual input = full confidence
            });
        }

        // -----------------------------------------------------------------------
        // Internal pipeline
        // -----------------------------------------------------------------------

        private void OnSpeechRecognised(object sender, SpeechRecognisedEventArgs e)
        {
            try
            {
                log.Debug($"Pipeline: Received speech \"{e.Text}\" (conf={e.Confidence:F2})");

                // Fire raw speech event for diagnostics
                OnRawSpeech?.Invoke(this, e);

                // Step 1: Run Bible reference detection
                var detected = BibleReferenceDetector.DetectBest(e.Text);
                if (detected == null)
                {
                    log.Debug("Pipeline: No Bible reference detected.");
                    return;
                }

                log.Info($"Pipeline: Detected \"{detected.NormalisedReference}\" " +
                         $"(det_conf={detected.Confidence:F2}, speech_conf={e.Confidence:F2})");

                // Step 1b: Validate and repair against real Bible data.
                // Catches impossible references and applies heuristics:
                //   "Zech 5:49"  → "Zech 5:4-9"   (verse as range)
                //   "Zech 49"    → "Zech 4:9"      (collapsed chapter:verse)
                //   "Zech 40:1"  → "Zech 14:1"     (forty/fourteen swap)
                //   "Ps 1:19"    → stays "Ps 1:19" (verse 19 exists in Ps 1)
                //   "Ps 119:1"   → stays "Ps 119:1" (Ps 119 is a real chapter)
                Bible validationBible = null;
                try
                {
                    // Load the default bible for validation.
                    // We use ESV as the validation source — it has all books/chapters/verses.
                    // This is a fast cached operation after first load.
                    validationBible = _validationBible ?? (_validationBible = TryLoadValidationBible());
                }
                catch { /* non-fatal — skip validation if bible unavailable */ }

                if (validationBible != null)
                {
                    var validated = ReferenceValidator.Validate(
                        validationBible, detected.BookName, detected.ReferenceFragment);

                    if (validated == null)
                    {
                        log.Warn($"Pipeline: \"{detected.NormalisedReference}\" failed validation and could not be repaired. Dropping.");
                        return;
                    }

                    if (validated.Outcome != ValidationOutcome.Valid)
                    {
                        log.Info($"Pipeline: Repaired \"{detected.NormalisedReference}\" → " +
                                 $"\"{validated.NormalisedReference}\" ({validated.Outcome})");

                        // Rebuild detected reference from the repaired values
                        detected = new DetectedReference
                        {
                            NormalisedReference = validated.NormalisedReference,
                            BookName            = validated.BookName,
                            ReferenceFragment   = validated.ReferenceFragment,
                            MatchedRawText      = detected.MatchedRawText,
                            Confidence          = detected.Confidence,
                        };
                    }
                }

                // Step 2: Compute combined confidence
                double combined = (e.Confidence * 0.4) + (detected.Confidence * 0.6);
                if (combined < MinCombinedConfidence)
                {
                    log.Debug($"Pipeline: Combined confidence {combined:F2} below threshold {MinCombinedConfidence:F2}, ignoring.");
                    return;
                }

                // Step 3: Duplicate suppression
                string refKey = detected.NormalisedReference;
                lock (_recentLock)
                {
                    if (_recentReferences.TryGetValue(refKey, out DateTime lastTime))
                    {
                        if ((DateTime.UtcNow - lastTime).TotalSeconds < DuplicateCooldownSeconds)
                        {
                            log.Debug($"Pipeline: Duplicate suppressed for \"{refKey}\" (cooldown={DuplicateCooldownSeconds}s).");
                            return;
                        }
                    }
                    _recentReferences[refKey] = DateTime.UtcNow;
                }

                // Step 4: Fire the event!
                log.Info($"Pipeline: Firing OnReferenceDetected → \"{detected.NormalisedReference}\"");

                OnReferenceDetected?.Invoke(this, new ReferenceDetectedEventArgs
                {
                    NormalisedReference = detected.NormalisedReference,
                    BookName = detected.BookName,
                    ReferenceFragment = detected.ReferenceFragment,
                    SpokenText = e.Text,
                    Confidence = combined,
                });
            }
            catch (Exception ex)
            {
                log.Error("Pipeline: Unhandled exception in speech processing.", ex);
            }
        }

        private void OnListenerStatusChanged(object sender, SpeechListenerStatusEventArgs e)
        {
            RaiseStatus(e.Message, e.IsError, _listener.IsListening);
        }

        // -----------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------

        private void CleanupRecentReferences()
        {
            lock (_recentLock)
            {
                var cutoff = DateTime.UtcNow.AddSeconds(-DuplicateCooldownSeconds * 2);
                var expired = new List<string>();
                foreach (var kvp in _recentReferences)
                {
                    if (kvp.Value < cutoff)
                        expired.Add(kvp.Key);
                }
                foreach (var key in expired)
                    _recentReferences.Remove(key);
            }
        }

        private void RaiseStatus(string message, bool isError = false, bool isListening = false)
        {
            OnStatusChanged?.Invoke(this, new ServiceStatusEventArgs
            {
                Message = message,
                IsError = isError,
                IsListening = isListening,
            });
        }

        /// <summary>
        /// Loads the ESV bible for reference validation. Cached after first load.
        /// Returns null if the bible file is unavailable (e.g. first run before install).
        /// </summary>
        private static Bible TryLoadValidationBible()
        {
            try
            {
                return OpenSongBibleReader.LoadTranslation("ESV");
            }
            catch
            {
                try { return OpenSongBibleReader.LoadTranslation("NASB"); }
                catch { return null; }
            }
        }

        // -----------------------------------------------------------------------
        // IDisposable
        // -----------------------------------------------------------------------

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            _cleanupTimer?.Stop();
            _cleanupTimer?.Dispose();
            _listener?.Dispose();

            log.Info("SpeechToScriptureService: Disposed.");
        }
    }
}
