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
        // Chapter-only debounce state
        //
        // When a chapter-only reference is detected (e.g. "Psalms 23"), we
        // hold it for ChapterOnlyGraceMs in case the verse is about to be
        // announced. The pending entry is guarded by _pendingLock — never
        // read or write any _pending* field outside that lock.
        // -----------------------------------------------------------------------
        private readonly object _pendingLock = new object();
        private DetectedReference _pendingDetection; // chapter-only, awaiting possible verse
        private float _pendingSpeechConfidence;
        private string _pendingSpokenText;
        private int _pendingChapterNumber;           // parsed chapter of the pending reference
        private System.Threading.Timer _pendingTimer;

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
        /// How long (in milliseconds) to hold a chapter-only reference back
        /// in case a verse number arrives in a following utterance.
        ///
        /// Ministers often speak scripture references slowly, with a noticeable
        /// pause between the chapter and the verse:
        ///     "Let's turn to Psalm twenty-three ... verse four."
        ///
        /// Without this grace window, the first utterance ("Psalm 23") would
        /// fire immediately and the presenter would get the whole chapter on
        /// screen just as the verse was being announced.
        ///
        /// If a chapter:verse reference is detected during the grace period
        /// for the same book+chapter, the pending chapter-only is cancelled
        /// and the fuller reference fires instead. If nothing arrives before
        /// the timer elapses, the chapter-only reference is released as-is.
        ///
        /// Default: 10000ms (10 seconds). Set to 0 to disable debouncing and
        /// fire chapter-only references immediately (legacy behaviour).
        /// </summary>
        public int ChapterOnlyGraceMs { get; set; } = 10_000;

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
        /// Fired as soon as a chapter-only reference (e.g. "Genesis 1") is
        /// heard, BEFORE the chapter-only grace window expires. Subscribers
        /// can use this to show an immediate notification so the presenter
        /// gets instant visual feedback that the mic heard them.
        ///
        /// If the chapter-only detection is later upgraded to a richer
        /// reference (e.g. "Genesis 1:5" after the minister says "verse
        /// five"), <see cref="OnReferenceDetected"/> will fire with the
        /// richer reference. Subscribers are expected to EDIT their existing
        /// notification rather than opening a second one.
        ///
        /// Also fires on a background thread — marshal as needed.
        /// </summary>
        public event EventHandler<ReferenceDetectedEventArgs> OnReferencePreliminary;

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
            CancelPendingReference();
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

                // Step 1a: Chapter-only debounce — if we have a pending
                // chapter-only reference from a previous utterance, try to
                // upgrade it by splicing this new text onto the pending
                // book/chapter context (e.g. the minister's follow-up
                // "verse four" after saying "Psalm twenty-three").
                if (detected == null || !detected.ReferenceFragment.Contains(":"))
                {
                    if (TryUpgradePendingWithFollowUp(e))
                    {
                        // The upgrade path already fired the richer reference.
                        return;
                    }
                }

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

                // Step 3: Chapter-only debounce.
                //
                // Ministers read slowly — they often say the chapter, pause,
                // then announce the verse. Firing the chapter-only reference
                // on the first utterance would show the whole chapter on
                // screen just as the verse number is being spoken.
                //
                // So: if this detection is chapter-only, stash it and wait
                // ChapterOnlyGraceMs for a verse to arrive. If a chapter:verse
                // detection lands first, cancel the pending and fire the
                // richer one. If the timer elapses, release the chapter-only.
                bool hasVerse = detected.ReferenceFragment != null
                                && detected.ReferenceFragment.Contains(":");

                if (!hasVerse && ChapterOnlyGraceMs > 0)
                {
                    StashPendingChapterOnly(detected, e.Text, e.Confidence);
                    // Let the UI show an immediate "Heard Genesis 1" notification.
                    // If a verse arrives before grace expires, OnReferenceDetected
                    // will fire with the richer "Genesis 1:5" and the UI should
                    // edit the same notification in place.
                    RaisePreliminary(detected, e.Text, e.Confidence);
                    return;
                }

                // We have a chapter:verse reference. If a pending chapter-only
                // is waiting for the same book+chapter, it's now superseded
                // — drop it without firing.
                if (hasVerse)
                {
                    DiscardPendingIfMatches(detected);
                }

                FireDetection(detected, e.Text, e.Confidence);
            }
            catch (Exception ex)
            {
                log.Error("Pipeline: Unhandled exception in speech processing.", ex);
            }
        }

        /// <summary>
        /// Fires the preliminary "heard something" event for a chapter-only
        /// detection that has been stashed for the grace window. Applies the
        /// same combined-confidence threshold as the main fire path so the
        /// toast never pops for a borderline detection that we'd ultimately
        /// have thrown away.
        /// </summary>
        private void RaisePreliminary(DetectedReference detected, string spokenText, float speechConfidence)
        {
            if (detected == null) return;
            try
            {
                double combined = (speechConfidence * 0.4) + (detected.Confidence * 0.6);
                if (combined < MinCombinedConfidence) return;

                OnReferencePreliminary?.Invoke(this, new ReferenceDetectedEventArgs
                {
                    NormalisedReference = detected.NormalisedReference,
                    BookName            = detected.BookName,
                    ReferenceFragment   = detected.ReferenceFragment,
                    SpokenText          = spokenText,
                    Confidence          = combined,
                });
            }
            catch (Exception ex)
            {
                log.Debug($"Pipeline: OnReferencePreliminary subscriber threw: {ex.Message}");
            }
        }

        /// <summary>
        /// Fires the OnReferenceDetected event after applying duplicate
        /// suppression and confidence threshold checks. Shared by the
        /// immediate-fire path and the pending-release path.
        /// </summary>
        private void FireDetection(DetectedReference detected, string spokenText, float speechConfidence)
        {
            if (detected == null) return;

            double combined = (speechConfidence * 0.4) + (detected.Confidence * 0.6);
            if (combined < MinCombinedConfidence)
            {
                log.Debug($"Pipeline: Combined confidence {combined:F2} below threshold {MinCombinedConfidence:F2}, ignoring.");
                return;
            }

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

            log.Info($"Pipeline: Firing OnReferenceDetected → \"{detected.NormalisedReference}\"");

            OnReferenceDetected?.Invoke(this, new ReferenceDetectedEventArgs
            {
                NormalisedReference = detected.NormalisedReference,
                BookName            = detected.BookName,
                ReferenceFragment   = detected.ReferenceFragment,
                SpokenText          = spokenText,
                Confidence          = combined,
            });
        }

        // -----------------------------------------------------------------------
        // Chapter-only debounce helpers
        // -----------------------------------------------------------------------

        /// <summary>
        /// Parses the leading chapter number out of a reference fragment
        /// like "23" or "23:4-7". Returns 0 if it can't be parsed.
        /// </summary>
        private static int ParseChapterNumber(string refFragment)
        {
            if (string.IsNullOrEmpty(refFragment)) return 0;
            int split = refFragment.IndexOfAny(new[] { ':', '-' });
            string chapterPart = split < 0 ? refFragment : refFragment.Substring(0, split);
            return int.TryParse(chapterPart.Trim(), out int n) ? n : 0;
        }

        /// <summary>
        /// Stores a chapter-only detection as "pending" and (re)starts the
        /// grace-period timer. If a previous pending exists for a DIFFERENT
        /// book/chapter, that one is released immediately — the minister
        /// has clearly moved on to a new reference.
        /// </summary>
        private void StashPendingChapterOnly(DetectedReference detected, string spokenText, float speechConfidence)
        {
            DetectedReference toReleaseNow = null;
            string            releaseSpoken = null;
            float             releaseConf   = 0f;

            int incomingChapter = ParseChapterNumber(detected.ReferenceFragment);

            lock (_pendingLock)
            {
                if (_pendingDetection != null)
                {
                    bool sameTarget = string.Equals(
                            _pendingDetection.BookName,
                            detected.BookName,
                            StringComparison.OrdinalIgnoreCase)
                        && _pendingChapterNumber == incomingChapter;

                    if (!sameTarget)
                    {
                        // Minister has moved on — release the previous pending
                        // so we don't silently drop it. Stash the new one.
                        toReleaseNow = _pendingDetection;
                        releaseSpoken = _pendingSpokenText;
                        releaseConf   = _pendingSpeechConfidence;
                    }
                    // Else: same book+chapter restated; just reset the timer below.
                }

                _pendingDetection        = detected;
                _pendingSpokenText       = spokenText;
                _pendingSpeechConfidence = speechConfidence;
                _pendingChapterNumber    = incomingChapter;

                _pendingTimer?.Dispose();
                _pendingTimer = new System.Threading.Timer(
                    OnPendingGraceExpired,
                    state: null,
                    dueTime: ChapterOnlyGraceMs,
                    period: System.Threading.Timeout.Infinite);
            }

            log.Debug($"Pipeline: Holding chapter-only \"{detected.NormalisedReference}\" " +
                      $"for up to {ChapterOnlyGraceMs}ms in case a verse follows.");

            if (toReleaseNow != null)
            {
                log.Info($"Pipeline: Releasing previous pending \"{toReleaseNow.NormalisedReference}\" " +
                          "(superseded by a different reference).");
                FireDetection(toReleaseNow, releaseSpoken, releaseConf);
            }
        }

        /// <summary>
        /// If the pending chapter-only reference matches the incoming
        /// chapter:verse reference (same book + same chapter), discard the
        /// pending — it's about to be superseded by something richer.
        /// </summary>
        private void DiscardPendingIfMatches(DetectedReference richer)
        {
            int richerChapter = ParseChapterNumber(richer.ReferenceFragment);

            lock (_pendingLock)
            {
                if (_pendingDetection == null) return;

                bool sameTarget = string.Equals(
                        _pendingDetection.BookName,
                        richer.BookName,
                        StringComparison.OrdinalIgnoreCase)
                    && _pendingChapterNumber == richerChapter;

                if (!sameTarget) return;

                log.Debug($"Pipeline: Pending \"{_pendingDetection.NormalisedReference}\" " +
                          $"upgraded to \"{richer.NormalisedReference}\".");

                _pendingDetection = null;
                _pendingSpokenText = null;
                _pendingSpeechConfidence = 0f;
                _pendingChapterNumber = 0;
                _pendingTimer?.Dispose();
                _pendingTimer = null;
            }
        }

        /// <summary>
        /// Called on a background thread when ChapterOnlyGraceMs elapses
        /// without any follow-up verse. Releases the pending chapter-only
        /// reference as the minister's intended insertion.
        /// </summary>
        private void OnPendingGraceExpired(object state)
        {
            DetectedReference detected;
            string spokenText;
            float  speechConfidence;

            lock (_pendingLock)
            {
                if (_pendingDetection == null) return;

                detected         = _pendingDetection;
                spokenText       = _pendingSpokenText;
                speechConfidence = _pendingSpeechConfidence;

                _pendingDetection = null;
                _pendingSpokenText = null;
                _pendingSpeechConfidence = 0f;
                _pendingChapterNumber = 0;
                _pendingTimer?.Dispose();
                _pendingTimer = null;
            }

            log.Info($"Pipeline: Grace period elapsed — releasing chapter-only " +
                     $"\"{detected.NormalisedReference}\".");
            FireDetection(detected, spokenText, speechConfidence);
        }

        /// <summary>
        /// Tries to promote the pending chapter-only reference by splicing
        /// this follow-up utterance onto the pending book + chapter context.
        ///
        /// If the minister said "Psalm twenty-three" and then "verse four",
        /// the second utterance has no book name of its own — the detector
        /// would return nothing. By prepending "Psalms chapter 23" we give
        /// the detector the context it needs to parse "verse four" as "23:4".
        ///
        /// Returns true if an upgrade was detected, validated and fired
        /// (in which case the caller should stop processing).
        /// </summary>
        private bool TryUpgradePendingWithFollowUp(SpeechRecognisedEventArgs followUp)
        {
            DetectedReference pending;
            string            pendingSpoken;
            float             pendingConfidence;
            int               pendingChapter;

            lock (_pendingLock)
            {
                if (_pendingDetection == null) return false;
                pending           = _pendingDetection;
                pendingSpoken     = _pendingSpokenText;
                pendingConfidence = _pendingSpeechConfidence;
                pendingChapter    = _pendingChapterNumber;
            }

            if (string.IsNullOrWhiteSpace(followUp?.Text)) return false;

            string spliced = $"{pending.BookName} chapter {pendingChapter} {followUp.Text}";
            var upgraded = BibleReferenceDetector.DetectBest(spliced);

            if (upgraded == null) return false;
            if (upgraded.ReferenceFragment == null || !upgraded.ReferenceFragment.Contains(":"))
                return false;
            if (!string.Equals(upgraded.BookName, pending.BookName, StringComparison.OrdinalIgnoreCase))
                return false;
            if (ParseChapterNumber(upgraded.ReferenceFragment) != pendingChapter)
                return false;

            // Clear the pending before firing so a subsequent detection on the
            // same text can't double-fire.
            lock (_pendingLock)
            {
                _pendingDetection = null;
                _pendingSpokenText = null;
                _pendingSpeechConfidence = 0f;
                _pendingChapterNumber = 0;
                _pendingTimer?.Dispose();
                _pendingTimer = null;
            }

            // Pass the pending speech through validation via the normal path.
            // Run the validator on the upgraded reference too.
            Bible validationBible = null;
            try { validationBible = _validationBible ?? (_validationBible = TryLoadValidationBible()); }
            catch { /* non-fatal */ }

            if (validationBible != null)
            {
                var validated = ReferenceValidator.Validate(
                    validationBible, upgraded.BookName, upgraded.ReferenceFragment);
                if (validated == null)
                {
                    log.Warn($"Pipeline: Upgraded \"{upgraded.NormalisedReference}\" failed validation. " +
                             $"Falling back to pending chapter-only release.");
                    FireDetection(pending, pendingSpoken, pendingConfidence);
                    return true;
                }
                if (validated.Outcome != ValidationOutcome.Valid)
                {
                    upgraded = new DetectedReference
                    {
                        NormalisedReference = validated.NormalisedReference,
                        BookName            = validated.BookName,
                        ReferenceFragment   = validated.ReferenceFragment,
                        MatchedRawText      = upgraded.MatchedRawText,
                        Confidence          = upgraded.Confidence,
                    };
                }
            }

            log.Info($"Pipeline: Upgraded pending \"{pending.NormalisedReference}\" → " +
                      $"\"{upgraded.NormalisedReference}\" via follow-up \"{followUp.Text}\".");

            // Use the stronger of the two speech confidences.
            float fusedSpeechConf = Math.Max(pendingConfidence, followUp.Confidence);
            string fusedSpoken = $"{pendingSpoken} {followUp.Text}".Trim();

            FireDetection(upgraded, fusedSpoken, fusedSpeechConf);
            return true;
        }

        /// <summary>
        /// Drops any pending chapter-only reference without firing. Useful
        /// when the user manually stops listening or disables auto mode.
        /// </summary>
        public void CancelPendingReference()
        {
            lock (_pendingLock)
            {
                if (_pendingDetection == null) return;
                log.Debug($"Pipeline: Discarding pending \"{_pendingDetection.NormalisedReference}\" (cancelled).");
                _pendingDetection = null;
                _pendingSpokenText = null;
                _pendingSpeechConfidence = 0f;
                _pendingChapterNumber = 0;
                _pendingTimer?.Dispose();
                _pendingTimer = null;
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

            lock (_pendingLock)
            {
                _pendingTimer?.Dispose();
                _pendingTimer     = null;
                _pendingDetection = null;
            }

            _listener?.Dispose();

            log.Info("SpeechToScriptureService: Disposed.");
        }
    }
}
